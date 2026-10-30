#!/usr/bin/env python3
"""Palms Cartridge Server: holds the Game Library cartridges, lets the host approve new ones,
and checks they haven't been tampered with. (In the code a cartridge is a "pak".)"""

import base64
import configparser
import difflib
import getpass
import hashlib
import hmac
import html
import io
import os
import re
import shutil
import struct
import subprocess
import sys
import threading
import time
import urllib.parse
from collections import deque
from http.server import BaseHTTPRequestHandler, HTTPServer
from socketserver import ThreadingMixIn


class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True

VERSION = "2.5.1"
CONFIG = "/etc/palms-paks.ini"
ADMINS = "/etc/palms-paks/admins"
REALM = "Palms Cartridge Server"
OLD_REALM = "Palms pak server"     # the login realm of older versions: those accounts need their password set again
SCAN_SECONDS = 5        # how often to look at the folders
QUIET_SECONDS = 30      # a pak without archive_bytes is ready once nothing has changed for this long
SETTLE_SECONDS = 15     # a finished pak sits in the inbox this long, untouched, before it's locked for review
PAK_FORMAT = 2          # the newest pak.ini format this program understands (2: packs of games)
AREAS = ("inbox", "waiting", "library", "trash")
LOCKOUT_TRIES = 5       # wrong passwords from one PC...
LOCKOUT_SECONDS = 600   # ...within this long locks it out for this long

try:
    from PIL import Image
except ImportError:
    Image = None
try:
    import pefile
except ImportError:
    pefile = None


def log(*parts):
    print(time.strftime("%H:%M:%S"), *parts, flush=True)


def stamp(t=None):
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(t))


# ------------------------------------------------------------------ settings

def load_settings(path=CONFIG):
    cp = configparser.ConfigParser(interpolation=None)
    cp.read(path, encoding="utf-8")
    s = cp["paks"] if cp.has_section("paks") else {}
    url = s.get("url", "").strip()
    if url and not url.endswith("/"):
        url += "/"
    return {
        "folder": s.get("folder", "/srv/paks"),
        "state": s.get("state", "/var/lib/palms-paks"),
        "url": url,
        "admin_port": to_int(s.get("admin_port", "8081"), 8081),
        "admins": s.get("admins", ADMINS),
        "trash_days": to_int(s.get("trash_days", "7"), 7),
        "check_hour": to_int(s.get("check_hour", "4"), 4),
    }


# ------------------------------------------------------------------ reading a pak

def read_text(path):
    """pak.ini as the Pak Maker writes it (Windows ANSI), or as Notepad might save it."""
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8", "replace")
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16", "replace")
    return raw.decode("cp1252", "replace")

def files_in(path):
    """{lower-case name: real name} - Windows doesn't care about case, Linux does."""
    try:
        return {e.name.lower(): e.name for e in os.scandir(path) if e.is_file()}
    except OSError:
        return {}

def is_id(s):
    return 2 <= len(s) <= 24 and all(c.isdigit() or ("a" <= c <= "z") or c == "-" for c in s)

def to_int(s, default=-1):
    try:
        return int(str(s).strip())
    except ValueError:
        return default


class Pak:
    """One pak folder, and what we made of it."""

    def __init__(self, area, root, rel):
        self.area = area        # inbox, waiting or library
        self.root = root        # that area's folder
        self.rel = rel          # the pak's folder in it, e.g. halflife-v1.1.1.0 (or folder/pak in the inbox)
        self.info = {}
        self.status = ""        # inbox: ready, copying, problem, waiting, settling
        self.message = ""       # waiting/library: checking, review, live, tampered, unapproved, hidden
        self.pieces = []        # real file names of the archive pieces
        self.got_bytes = 0
        self.sig = ""           # changes whenever the pak does
        self.icon = ""          # icons/<name>.bmp once made
        self.png = ""
        self.shot = ""          # icons/<name>-shot.jpg, if the cartridge has a screenshot
        self.games = []         # a pack's games: dicts with id, name, players, exe, args...
        self.meta = {}          # from its fingerprints: when, who approved it...

    @property
    def path(self):
        return os.path.join(self.root, self.rel)

    @property
    def name(self):
        return self.rel.split("/")[-1]


def look_at(pak, memory, now):
    """Work out whether a pak folder is complete. memory keeps sizes between looks."""
    path = pak.path
    names = files_in(path)
    if "pak.ini" not in names:
        pak.status = "waiting"
        pak.message = "pak.ini not found"
        return
    ini_path = os.path.join(path, names["pak.ini"])
    cp = configparser.ConfigParser(interpolation=None, strict=False)
    try:
        cp.read_string(read_text(ini_path))
    except (configparser.Error, OSError) as e:
        pak.status = "problem"
        pak.message = "pak.ini parse error: %s" % str(e).splitlines()[0]
        return
    if not cp.has_section("pak"):
        pak.status = "problem"
        pak.message = "pak.ini: no [pak] section"
        return
    info = {k: v.strip() for k, v in cp["pak"].items()}
    pak.info = info

    fmt = to_int(info.get("format", "1"))
    if fmt > PAK_FORMAT:
        pak.status = "problem"
        pak.message = "Unsupported cartridge format %d (max %d)" % (fmt, PAK_FORMAT)
        return
    said = info.get("id", "")
    pid = said.lower()
    info["id"] = pid
    if not is_id(pid):
        pak.status = "problem"
        pak.message = "Invalid id '%s' (2-24 chars, a-z 0-9 -)" % said
        return
    if not info.get("name"):
        pak.status = "problem"
        pak.message = "pak.ini: name missing"
        return
    pak.games = read_games(cp, info)
    if fmt >= 2 or pak.games:
        if not pak.games:
            pak.status = "problem"
            pak.message = "pak.ini: no complete [gameN] sections"
            return
    elif not info.get("exe"):
        pak.status = "problem"
        pak.message = "pak.ini: exe missing"
        return
    count = to_int(info.get("pieces", ""))
    if count < 1:
        pak.status = "problem"
        pak.message = "pak.ini: pieces missing"
        return

    archive = info.get("archive") or "game.7z"
    pieces, missing, total, sizes = [], [], 0, []
    for n in range(1, count + 1):
        want = ("%s.%03d" % (archive, n)).lower()
        if want in names:
            real = names[want]
            pieces.append(real)
            try:
                st = os.stat(os.path.join(path, real))
                total += st.st_size
                sizes.append((st.st_size, int(st.st_mtime)))
            except OSError:
                missing.append(want)
        else:
            missing.append(want)
    pak.pieces = pieces
    pak.got_bytes = total
    try:
        ini_stamp = int(os.stat(ini_path).st_mtime)
    except OSError:
        ini_stamp = 0

    if missing:
        pak.status = "copying"
        pak.message = "Waiting for %s (%d/%d pieces)" % (missing[0], len(pieces), count)
        return
    # the small files pak.ini mentions (copying a pak in Explorer can bring pak.ini before setup.reg)
    small = [info.get(k, "") for k in ("reg", "notes", "icon", "screenshot")] + [g.get("icon", "") for g in pak.games]
    for f in small:
        if f and f.lower() not in names:
            pak.status = "copying"
            pak.message = "Waiting for %s" % f
            return

    expect = to_int(info.get("archive_bytes", ""))
    if expect >= 0:
        if total < expect:
            pak.status = "copying"
            pak.message = "%s / %s received" % (size_text(total), size_text(expect))
            return
        if total > expect:
            pak.status = "problem"
            pak.message = ("Size mismatch: pieces %s bytes, archive_bytes %s"
                           % ("{:,}".format(total), "{:,}".format(expect)))
            return
    else:
        # no archive_bytes in pak.ini: ready once the pieces stop changing
        key = tuple(sizes)
        seen = memory.get(pak.rel)
        if not seen or seen[0] != key:
            memory[pak.rel] = (key, now)
            pak.status = "copying"
            pak.message = "No archive_bytes: waiting for pieces to stop changing"
            return
        if now - seen[1] < QUIET_SECONDS:
            pak.status = "copying"
            pak.message = "No archive_bytes: waiting for pieces to stop changing"
            return

    pak.status = "ready"
    pak.message = ""
    pak.sig = hashlib.sha1(("%s|%d|%d|%s" % (pak.name, ini_stamp, total, sizes)).encode()).hexdigest()[:10]


def read_games(cp, info):
    """A pack's games: [game1], [game2]... in pak.ini, in that order. Each game's own
    args come first, then the pack's, which go after every game (like -nomaster -steam)."""
    sections = []
    for name in cp.sections():
        low = name.lower()
        if low.startswith("game") and low[4:].isdigit():
            sections.append((int(low[4:]), name))
    games, seen = [], set()
    for _, name in sorted(sections):
        g = {k: v.strip() for k, v in cp[name].items()}
        gid = g.get("id", "").lower()
        if not is_id(gid) or gid in seen or not g.get("name") or not g.get("exe"):
            continue
        seen.add(gid)
        g["id"] = gid
        g["args"] = " ".join(x for x in (g.get("args", ""), info.get("args", "")) if x)
        g["icon_url"], g["png"] = "", ""
        games.append(g)
    return games


def find_paks(area, root):
    """Every pak folder in an area. In the inbox, one level further down too, for a folder of paks."""
    found = []
    try:
        tops = sorted((e.name for e in os.scandir(root) if e.is_dir() and not e.name.startswith(".")),
                      key=str.lower)
    except OSError as e:
        log("Unable to read", root, "-", e)
        return found
    for top in tops:
        path = os.path.join(root, top)
        if area != "inbox" or "pak.ini" in files_in(path):
            found.append(Pak(area, root, top))
            continue
        try:
            subs = sorted((e.name for e in os.scandir(path) if e.is_dir() and not e.name.startswith(".")),
                          key=str.lower)
        except OSError:
            subs = []
        inner = [s for s in subs if "pak.ini" in files_in(os.path.join(path, s))]
        if inner:
            found.extend(Pak(area, root, top + "/" + s) for s in inner)
        else:
            found.append(Pak(area, root, top))
    return found


def folder_stamp(path):
    """Every file in a folder with its size and time: changes whenever anything in it does."""
    out = []
    for dirpath, dirs, files in os.walk(path):
        for f in files:
            try:
                st = os.stat(os.path.join(dirpath, f))
                out.append((os.path.relpath(os.path.join(dirpath, f), path), st.st_size, st.st_mtime_ns))
            except OSError:
                pass
    return tuple(sorted(out))


# ------------------------------------------------------------------ icons

def ico_from_program(data):
    """The program's own icon (the one Explorer shows), as the bytes of an .ico file."""
    if pefile is None:
        return None
    try:
        pe = pefile.PE(data=data, fast_load=True)
        pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_RESOURCE"]])
    except Exception:
        return None
    if not hasattr(pe, "DIRECTORY_ENTRY_RESOURCE"):
        return None
    icons, groups = {}, []
    for kind in pe.DIRECTORY_ENTRY_RESOURCE.entries:
        if kind.id not in (3, 14) or not hasattr(kind, "directory"):
            continue
        for entry in kind.directory.entries:
            if not hasattr(entry, "directory") or not entry.directory.entries:
                continue
            res = entry.directory.entries[0].data.struct
            blob = pe.get_data(res.OffsetToData, res.Size)
            if kind.id == 3:
                icons[entry.id] = blob
            else:
                groups.append(blob)
    if not groups:
        return None
    group = groups[0]
    try:
        count = struct.unpack_from("<HHH", group, 0)[2]
        parts = []
        for i in range(count):
            w, h, colours, res, planes, bits, size, nid = struct.unpack_from("<BBBBHHIH", group, 6 + i * 14)
            if nid in icons:
                parts.append((w, h, colours, res, planes, bits, icons[nid]))
    except struct.error:
        return None
    if not parts:
        return None
    out = io.BytesIO()
    out.write(struct.pack("<HHH", 0, 1, len(parts)))
    offset = 6 + 16 * len(parts)
    for w, h, colours, res, planes, bits, blob in parts:
        out.write(struct.pack("<BBBBHHII", w, h, colours, res, planes, bits, len(blob), offset))
        offset += len(blob)
    for part in parts:
        out.write(part[6])
    return out.getvalue()

def open_picture(data):
    """Two pictures from an icon or image file: (about 32 px, 16 px or None)."""
    im = Image.open(io.BytesIO(data))
    if im.format == "ICO":
        sizes = sorted(im.ico.sizes())
        def nearest(target):
            exact = [s for s in sizes if s[0] == target]
            if exact:
                return exact[0]
            bigger = [s for s in sizes if s[0] > target]
            return bigger[0] if bigger else sizes[-1]
        big = im.ico.getimage(nearest(32))
        small = im.ico.getimage((16, 16)) if any(s[0] == 16 for s in sizes) else None
        return big, small
    im.load()
    return im, None


def from_archive(pak, inside):
    """One file out of the pak's archive, without unpacking the rest."""
    first = os.path.join(pak.path, pak.pieces[0])
    name = inside.replace("\\", "/")
    try:
        r = subprocess.run(["7z", "e", "-so", "-bd", first, name],
                           stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, timeout=120)
    except (OSError, subprocess.TimeoutExpired) as e:
        log("  Unable to extract", inside, "from", pak.rel, "-", e)
        return None
    return r.stdout if r.returncode == 0 and r.stdout else None


def flatten(im, size):
    """A square picture on white, plus its mask (black = solid)."""
    im = im.convert("RGBA")
    if im.size != (size, size):
        im = im.resize((size, size), Image.LANCZOS)
    pic = Image.new("RGB", (size, size), (0, 0, 0))
    mask = Image.new("RGB", (size, size), (255, 255, 255))
    flat = Image.new("RGB", (size, size), (255, 255, 255))
    src = im.load()
    for y in range(size):
        for x in range(size):
            r, g, b, a = src[x, y]
            # soft edges blended onto white, the colour of the Game Library list
            r = (r * a + 255 * (255 - a)) // 255
            g = (g * a + 255 * (255 - a)) // 255
            b = (b * a + 255 * (255 - a)) // 255
            if a >= 64:
                pic.putpixel((x, y), (r, g, b))
                mask.putpixel((x, y), (0, 0, 0))
                flat.putpixel((x, y), (r, g, b))
    return pic, mask, flat


def make_icon(state, pak, icon, icon_from, name):
    """icons/<name>.bmp for the LAN Center, icons/<name>.png for the web page and
    icons/<name>.ico for the Start menu shortcut.

    The .bmp is 96 x 32: the 32 px picture, its mask, then the 16 px picture and
    its mask in the top half. A plain bitmap and a mask is what old PCs draw
    fastest, so all the work of reading icons happens here, once."""
    if Image is None:
        return "", ""
    data, said = None, ""
    if icon:
        names = files_in(pak.path)
        real = names.get(icon.lower())
        if real:
            with open(os.path.join(pak.path, real), "rb") as f:
                data = f.read()
            said = icon
    elif icon_from:
        data = from_archive(pak, icon_from)
        said = icon_from
    if not data:
        return "", ""
    ext = os.path.splitext(said)[1].lower()
    try:
        if ext in (".exe", ".dll"):
            data = ico_from_program(data)
            if not data:
                log("  %s contains no icon" % said)
                return "", ""
        big, small = open_picture(data)
        pic32, mask32, flat32 = flatten(big, 32)
        pic16, mask16, flat16 = flatten(small or big, 16)
    except Exception as e:
        log("  Unable to create an icon from %s: %s" % (said, e))
        return "", ""
    sheet = Image.new("RGB", (96, 32), (255, 255, 255))
    sheet.paste(pic32, (0, 0))
    sheet.paste(mask32, (32, 0))
    sheet.paste(pic16, (64, 0))
    sheet.paste(mask16, (80, 0))
    folder = os.path.join(state, "icons")
    os.makedirs(folder, exist_ok=True)
    sheet.save(os.path.join(folder, name + ".bmp"), "BMP")
    flat32.save(os.path.join(folder, name + ".png"), "PNG")
    try:
        gif_icon(big, os.path.join(folder, name + ".gif"))
    except Exception as e:
        log("  Unable to create %s.gif: %s" % (name, e))
    flat16.save(os.path.join(folder, name + "-16.png"), "PNG")
    try:
        write_ico(os.path.join(folder, name + ".ico"), [(pic32, mask32), (pic16, mask16)])
    except Exception as e:
        log("  Unable to create %s.ico: %s" % (name, e))
    return "icons/%s.bmp" % name, "icons/%s.png" % name


def write_ico(path, images):
    """An .ico the old way (a 24-bit picture and a 1-bit mask for each size), which
    Windows 98 and XP both show. Newer icon files that hold PNGs don't work on XP."""
    parts = []
    for pic, mask in images:
        w, h = pic.size
        row = (w * 3 + 3) // 4 * 4
        mrow = (w + 31) // 32 * 4
        xor, andm = bytearray(), bytearray()
        px, mk = pic.load(), mask.load()
        for y in range(h - 1, -1, -1):          # bitmaps go from the bottom up
            line = bytearray()
            for x in range(w):
                r, g, b = px[x, y][:3]
                line += bytes((b, g, r))
            xor += line + bytes(row - len(line))
            bits = bytearray(mrow)
            for x in range(w):
                if mk[x, y][0] > 127:           # white in the mask: see-through
                    bits[x // 8] |= 0x80 >> (x % 8)
            andm += bits
        head = struct.pack("<IiiHHIIiiII", 40, w, h * 2, 1, 24, 0, len(xor) + len(andm), 0, 0, 0, 0)
        parts.append((w, h, head + bytes(xor) + bytes(andm)))
    out = io.BytesIO()
    out.write(struct.pack("<HHH", 0, 1, len(parts)))
    offset = 6 + 16 * len(parts)
    for w, h, blob in parts:
        out.write(struct.pack("<BBBBHHII", w % 256, h % 256, 0, 0, 1, 24, len(blob), offset))
        offset += len(blob)
    for _, _, blob in parts:
        out.write(blob)
    with open(path, "wb") as f:
        f.write(out.getvalue())

def gif_icon(im, path):
    """A 32 px GIF with a see-through background and no white edges, for the projector (dark, IE5)."""
    im = im.convert("RGBA")
    if im.size != (32, 32):
        im = im.resize((32, 32), Image.LANCZOS)
    rgb = Image.new("RGB", im.size, (0, 0, 0))
    rgb.paste(im, mask=im.split()[3])
    pal = rgb.quantize(colors=255)
    alpha = im.split()[3].load()
    px = pal.load()
    for y in range(32):
        for x in range(32):
            if alpha[x, y] < 128:
                px[x, y] = 255
    pal.save(path, "GIF", transparency=255)


def make_shot(state, pak, shot, name):
    """icons/<name>-shot.jpg: the cartridge's screenshot, at most 640 x 480, for the projector."""
    if Image is None or not shot:
        return ""
    real = files_in(pak.path).get(shot.lower())
    if not real:
        return ""
    try:
        im = Image.open(os.path.join(pak.path, real))
        im.load()
        im = im.convert("RGB")
        im.thumbnail((640, 480), Image.LANCZOS)
        folder = os.path.join(state, "icons")
        os.makedirs(folder, exist_ok=True)
        im.save(os.path.join(folder, name + "-shot.jpg"), "JPEG", quality=85)
    except Exception as e:
        log("  Unable to read the screenshot %s: %s" % (shot, e))
        return ""
    return "icons/%s-shot.jpg" % name


def tidy_icons(state, keep):
    folder = os.path.join(state, "icons")
    try:
        for f in os.scandir(folder):
            stem = os.path.splitext(f.name)[0]
            if stem.endswith("-16"):
                stem = stem[:-3]
            elif stem.endswith("-shot"):
                stem = stem[:-5]
            if f.is_file() and stem not in keep:
                os.remove(f.path)
    except OSError:
        pass


# ------------------------------------------------------------------ the list

def size_text(n):
    if n >= 1024 ** 3:
        return "%.1f GB" % (n / 1024 ** 3)
    if n >= 1024 ** 2:
        return "%d MB" % round(n / 1024 ** 2)
    return "%d KB" % max(1, n // 1024)

def ver_text(info):
    """A cartridge's version for showing: 1.09 (Rev. 2). Either part may be missing."""
    ver, rev = info.get("version", ""), info.get("revision", "")
    if not rev:
        return ver
    return "%s (Rev. %s)" % (ver, rev) if ver else "Rev. %s" % rev


def clean(s):
    return str(s).replace("|", "/").replace("\r", " ").replace("\n", " ").strip()

CATALOG_FIELDS = ("id", "name", "version", "players", "folder", "exe", "args", "config_exe",
                  "config_args", "reg", "notes")

GAME_FIELDS = ("id", "name", "players", "exe", "args", "start_in", "config_exe", "config_args")

def catalog_lines(ready, url):
    """The list, the same shape the LAN Center server passes on:
    PAKS|format|how many paks|when|address of this server
    PAK|id|name|version|players|folder|exe|args|config exe|config args|reg|notes|
        pak folder|archive|pieces|archive bytes|installed size|files|icon|made|revision|screenshot
    and after a pack's PAK line, one line for each game in it (its exe is blank):
    GAME|pak id|game id|name|players|exe|args|start in|config exe|config args|icon"""
    lines = ["PAKS|2|%d|%s|%s" % (len(ready), time.strftime("%Y-%m-%d %H:%M:%S"), url)]
    for p in ready:
        i = dict(p.info)
        if p.games:
            for k in ("exe", "args", "config_exe", "config_args"):
                i[k] = ""
        archive = p.pieces[0][:-4] if p.pieces else i.get("archive", "game.7z")
        folder_url = urllib.parse.quote("paks/" + p.rel + "/")
        lines.append("|".join(["PAK"] + [clean(i.get(k, "")) for k in CATALOG_FIELDS] +
                              [folder_url, clean(archive), str(len(p.pieces)), str(p.got_bytes),
                               clean(i.get("size", "")), clean(i.get("files", "")), p.icon,
                               clean(i.get("made", "")), clean(i.get("revision", "")), p.shot]))
        for g in p.games:
            lines.append("|".join(["GAME", i["id"]] + [clean(g.get(k, "")) for k in GAME_FIELDS] + [g["icon_url"]]))
    return lines

def write_if_changed(path, text, encoding):
    data = text.encode(encoding, "replace")
    try:
        with open(path, "rb") as f:
            if f.read() == data:
                return False
    except OSError:
        pass
    tmp = path + ".new"
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, path)
    return True


# ------------------------------------------------------------------ fingerprints
# A pak's fingerprints: every file's SHA-256, size, time and inode. Kept in
# <state>/manifests/<area>/<pak>.txt, out of reach of the share.

def fingerprint(path, progress=None):
    """[(file, size, mtime_ns, inode, sha256)] for every file in a pak folder."""
    files = []
    for dirpath, dirs, names in os.walk(path):
        dirs.sort()
        for n in sorted(names):
            files.append(os.path.relpath(os.path.join(dirpath, n), path))
    total = 0
    for f in files:
        try:
            total += os.stat(os.path.join(path, f)).st_size
        except OSError:
            pass
    done, out = 0, []
    for f in files:
        full = os.path.join(path, f)
        h = hashlib.sha256()
        with open(full, "rb") as fh:
            while True:
                block = fh.read(4 * 1024 * 1024)
                if not block:
                    break
                h.update(block)
                done += len(block)
                if progress:
                    progress(done, total)
        st = os.stat(full)
        out.append((f.replace(os.sep, "/"), st.st_size, st.st_mtime_ns, st.st_ino, h.hexdigest()))
    return out


def entries_digest(entries):
    """One fingerprint for the whole pak."""
    h = hashlib.sha256()
    for f, size, _, _, sha in entries:
        h.update(("%s|%d|%s\n" % (f.lower(), size, sha)).encode("utf-8"))
    return h.hexdigest()


def quick_check(path, entries):
    """What's changed since the fingerprints, going by sizes, times and inodes. [] if nothing."""
    issues = []
    want = {f.lower(): (f, size, mtime, ino) for f, size, mtime, ino, _ in entries}
    seen = set()
    for dirpath, dirs, names in os.walk(path):
        for n in names:
            rel = os.path.relpath(os.path.join(dirpath, n), path).replace(os.sep, "/")
            seen.add(rel.lower())
            if rel.lower() not in want:
                issues.append("%s: added" % rel)
    if not os.path.isdir(path):
        return ["folder missing"]
    for low, (f, size, mtime, ino) in want.items():
        if low not in seen:
            issues.append("%s: missing" % f)
            continue
        try:
            st = os.stat(os.path.join(path, f))
        except OSError:
            issues.append("%s: unreadable" % f)
            continue
        if st.st_size != size:
            a, b = "{:,}".format(st.st_size), "{:,}".format(size)
            issues.append("%s: size %s, expected %s" % (f, a, b))
        elif st.st_mtime_ns != mtime or st.st_ino != ino:
            issues.append("%s: mtime/inode changed" % f)
    return issues


def full_check(path, entries):
    """quick_check, then every file read again and its SHA-256 compared."""
    issues = quick_check(path, entries)
    if issues:
        return issues
    for f, size, _, _, sha in entries:
        h = hashlib.sha256()
        try:
            with open(os.path.join(path, f), "rb") as fh:
                while True:
                    block = fh.read(4 * 1024 * 1024)
                    if not block:
                        break
                    h.update(block)
        except OSError:
            issues.append("%s: unreadable" % f)
            continue
        if h.hexdigest() != sha:
            issues.append("%s: SHA-256 mismatch" % f)
    return issues


def save_manifest(mdir, area, name, meta, entries):
    folder = os.path.join(mdir, area)
    os.makedirs(folder, exist_ok=True)
    lines = ["# Palms pak fingerprints: %s" % name]
    for k in sorted(meta):
        lines.append("%s=%s" % (k, clean(meta[k])))
    for f, size, mtime, ino, sha in entries:
        lines.append("F|%s|%d|%d|%d|%s" % (f, size, mtime, ino, sha))
    path = os.path.join(folder, name + ".txt")
    tmp = path + ".new"
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    os.replace(tmp, path)


def load_manifest(mdir, area, name):
    path = os.path.join(mdir, area, name + ".txt")
    try:
        with open(path, encoding="utf-8") as fh:
            text = fh.read()
    except OSError:
        return None
    meta, entries = {}, []
    for line in text.splitlines():
        if line.startswith("F|"):
            p = line.split("|")
            if len(p) >= 6:
                entries.append((p[1], int(p[2]), int(p[3]), int(p[4]), p[5]))
        elif "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            meta[k] = v
    return meta, entries


def drop_manifest(mdir, area, name):
    try:
        os.remove(os.path.join(mdir, area, name + ".txt"))
    except OSError:
        pass


# ------------------------------------------------------------------ what a pak does to a PC
# A pak's setup.reg goes into every PC's registry, and its programs run there, so the
# admin page points out anything a game has no business doing.

RISKY_KEYS = [
    ("\\currentversion\\run", "Autorun (Run)"),
    ("\\currentversion\\runonce", "Autorun (RunOnce)"),
    ("\\currentversion\\runservices", "Autorun (RunServices)"),
    ("\\currentversion\\policies", "Policies key"),
    ("\\winlogon", "Winlogon key"),
    ("\\image file execution options", "IFEO (debugger hijack)"),
    ("\\browser helper objects", "IE Browser Helper Object"),
    ("\\shellserviceobjectdelayload", "ShellServiceObjectDelayLoad"),
    ("\\shellexecutehooks", "ShellExecuteHooks"),
    ("\\shell extensions", "Shell extension"),
    ("\\active setup\\installed components", "Active Setup (per-user run)"),
    ("\\windows nt\\currentversion\\windows", "Windows key (AppInit_DLLs)"),
    ("\\session manager", "Session Manager"),
    ("\\currentcontrolset\\services", "Services key (service/driver)"),
    ("\\controlset001\\services", "Services key (service/driver)"),
    ("\\sharedaccess", "Firewall (SharedAccess)"),
    ("\\internet settings", "Internet Settings"),
    ("\\control\\lsa", "LSA"),
    ("\\environment", "Environment (PATH etc.)"),
    ("hkey_classes_root", "File associations (Classes)"),
    ("\\software\\classes", "File associations (Classes)"),
]
PROGRAM_EXTS = (".exe", ".bat", ".cmd", ".com", ".pif", ".scr", ".vbs", ".vbe", ".js", ".jse", ".wsf", ".hta", ".dll", ".cpl")


def check_reg(text):
    """[(line number, line, what's worrying about it or "")] for a registry file."""
    out, key = [], ""
    for n, line in enumerate(text.splitlines(), 1):
        t = line.strip()
        why = ""
        if t.startswith("["):
            key = t.strip("[]").lower()
            if key.startswith("-"):
                why = "Key deletion"
                key = key[1:]
            root_ok = key.startswith("hkey_local_machine\\software\\") or key.startswith("hkey_current_user\\software\\")
            for frag, what in RISKY_KEYS:
                if frag in key:
                    why = what
                    break
            if not why and not root_ok:
                why = "Non-standard root (expected HKLM/HKCU\\SOFTWARE)"
        elif "=" in t and not t.startswith(";"):
            low = t.lower()
            data = low.split("=", 1)[1]
            if "=-" in low.replace(" ", ""):
                why = "Value deletion"
            elif any(x in data for x in PROGRAM_EXTS) and "%installdir" not in data:
                why = "Executable path outside %INSTALLDIR%"
            elif data.startswith("hex"):
                pass
            for frag, what in RISKY_KEYS:
                if not why and frag in key:
                    why = what
        out.append((n, line, why))
    return out


def program_worries(exe, args):
    w = []
    ext = os.path.splitext(exe.lower())[1]
    if ext != ".exe":
        w.append("Not an .exe (%s)" % (ext or "no extension"))
    a = args.lower()
    if "http://" in a or "https://" in a or "\\\\" in a:
        w.append("Network path in arguments")
    if ".." in exe or ":" in exe or exe.startswith(("\\", "/")):
        w.append("Executable outside game folder")
    return w


# ------------------------------------------------------------------ the audit log

class Audit:
    """audit.log: when|event|pak|who|from|what. Newest at the bottom."""

    def __init__(self, path):
        self.path = path
        self.lock = threading.Lock()

    def add(self, event, pak="", who="", ip="", what=""):
        line = "|".join([stamp(), event, clean(pak), clean(who), clean(ip), clean(what)])
        with self.lock:
            try:
                with open(self.path, "a", encoding="utf-8") as fh:
                    fh.write(line + "\n")
            except OSError as e:
                log("Unable to write to the audit log:", e)
        log("AUDIT", line)

    def recent(self, n=200, pak=None):
        try:
            with open(self.path, "rb") as fh:
                fh.seek(0, 2)
                size = fh.tell()
                fh.seek(max(0, size - 400000))
                text = fh.read().decode("utf-8", "replace")
        except OSError:
            return []
        rows = [l.split("|", 5) for l in text.splitlines() if l.count("|") >= 5]
        if pak:
            rows = [r for r in rows if r[2] == pak]
        return rows[-n:][::-1]


# ------------------------------------------------------------------ admins

def read_admins(path):
    """{name: HA1} - HA1 is MD5(name:realm:password), what Digest logins check against."""
    out = {}
    try:
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line and not line.startswith("#") and ":" in line:
                    name, ha1 = line.split(":", 1)
                    out[name.strip()] = ha1.strip()
    except OSError:
        pass
    return out


def write_admins(path, admins):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".new"
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write("# Palms Cartridge Server host accounts: name:MD5(name:%s:password). Change with palms_paks.py --add-admin\n" % REALM)
        for name in sorted(admins):
            fh.write("%s:%s\n" % (name, admins[name]))
    os.chmod(tmp, 0o640)
    try:
        import grp
        os.chown(tmp, 0, grp.getgrnam("palms").gr_gid)
    except (KeyError, PermissionError, ImportError, OSError):
        pass
    os.replace(tmp, path)


def ha1_of(name, password):
    return hashlib.md5(("%s:%s:%s" % (name, REALM, password)).encode("utf-8")).hexdigest()


# ------------------------------------------------------------------ the pak server

class PakServer:
    def __init__(self, settings):
        self.base = settings["folder"]
        self.state = settings["state"]
        self.url = settings["url"]
        self.settings = settings
        self.dirs = {a: os.path.join(self.base, a) for a in AREAS}
        self.mdir = os.path.join(self.state, "manifests")
        self.audit = Audit(os.path.join(self.state, "audit.log"))
        self.lock = threading.RLock()
        self.wake = threading.Event()
        self.memory = {}        # inbox: sizes of paks still copying
        self.settle = {}        # inbox: pak -> (everything in it, since when)
        self.icons = {}         # pak signature -> (bmp, png)
        self.shots = {}         # pak signature -> icons/<name>-shot.jpg
        self.last = {}          # pak -> status, for the log
        self.view = {"inbox": [], "waiting": [], "library": []}
        self.ready = []
        self.jobs = deque()     # ("fingerprint", area, name, extra) or ("verify", ...)
        self.job_now = None     # what the checker is doing: {"what", "pak", "done", "total"}
        self.jobs_wake = threading.Event()
        self.writers = {}       # inbox folder -> {ip: count}, from Samba's audit
        self.share_events = {}  # (ip, folder, what) -> [count, first, last]
        self.share_audit = "starting"
        self.share_seen = 0
        self.last_full = ""     # when every pak was last read right through
        self.last_full_day = ""
        self.last_purge_day = ""
        self.secret = os.urandom(32)
        self.flash = {}         # admin -> message for their next page
        self.nonces = {}        # Digest login nonces -> [made, used ncs]
        self.fails = {}         # ip -> [times of wrong passwords]
        self.page_key = None
        for d in list(self.dirs.values()) + [self.mdir, os.path.join(self.state, "icons")]:
            os.makedirs(d, exist_ok=True)

    # ---------------------------------------------------------- moving paks about

    def area_path(self, area, name):
        return os.path.join(self.dirs[area], name)

    def to_trash(self, area, rel, why):
        """Out of the way (kept a while, in case), never deleted straight off."""
        name = rel.split("/")[-1]
        dest = "%s__%s__%s" % (name, time.strftime("%Y%m%d-%H%M%S"), why)
        src = self.area_path(area, rel)
        try:
            os.rename(src, os.path.join(self.dirs["trash"], dest))
        except OSError:
            shutil.move(src, os.path.join(self.dirs["trash"], dest))
        if area != "inbox":
            drop_manifest(self.mdir, area, name)
        return dest

    def lock_for_review(self, p):
        """A finished pak in the inbox: move it off the share and fingerprint it."""
        name = p.name
        dest = self.area_path("waiting", name)
        writers = self.writers.get(p.rel.split("/")[0], {})
        from_ip = max(writers, key=writers.get) if writers else ""
        if os.path.exists(dest):
            self.to_trash("waiting", name, "superseded")
            self.audit.add("superseded", name, "", "", "Older waiting copy replaced.")
        try:
            os.rename(p.path, dest)
        except OSError:
            shutil.move(p.path, dest)
        # read-only from here on: the tripwire notices anything that changes anyway
        for dirpath, dirs, files in os.walk(dest):
            for f in files:
                try:
                    os.chmod(os.path.join(dirpath, f), 0o444)
                except OSError:
                    pass
        top = p.rel.split("/")[0]
        if "/" in p.rel:
            try:
                os.rmdir(os.path.join(self.dirs["inbox"], top))
            except OSError:
                pass
        self.writers.pop(top, None)
        i = p.info
        self.audit.add("arrived", name, "", from_ip, "%s %s (id %s), %s. Moved from incoming to waiting."
                       % (i.get("name", ""), ver_text(i), i.get("id", ""), size_text(p.got_bytes)))
        self.queue(("fingerprint", "waiting", name, {"from_ip": from_ip, "arrived": stamp()}))

    def queue(self, job):
        with self.lock:
            if job not in self.jobs:
                self.jobs.append(job)
        self.jobs_wake.set()

    # ---------------------------------------------------------- admin actions (from the admin page)

    def find(self, area, name):
        for p in self.view.get(area, []):
            if p.rel == name:
                return p
        return None

    def approve(self, name, who, ip):
        with self.lock:
            p = self.find("waiting", name)
            if not p:
                return "FAIL - %s is not in waiting" % name
            if p.message != "review":
                return "FAIL - %s not ready: %s" % (name, p.meta.get("why", "hashing in progress"))
            m = load_manifest(self.mdir, "waiting", name)
            issues = quick_check(p.path, m[1]) if m else ["no manifest"]
            if issues:
                self.flag("waiting", name, issues)
                return "FAIL - %s modified since hashing, not approved: %s" % (name, "; ".join(issues[:3]))
            pid = p.info["id"]
            replaced = []
            if os.path.exists(self.area_path("library", name)):
                self.to_trash("library", name, "replaced")
                replaced.append(name)
            for other in self.view["library"]:
                if other.rel != name and other.info.get("id") == pid and os.path.exists(other.path):
                    self.to_trash("library", other.rel, "replaced")
                    replaced.append(other.rel)
            dest = self.area_path("library", name)
            try:
                os.rename(p.path, dest)
            except OSError:
                shutil.move(p.path, dest)
            meta, entries = m
            meta.update({"approved_by": who, "approved_ip": ip, "approved_at": stamp()})
            meta.pop("why", None)
            # moving keeps the files as they are, but another drive would give them new inodes
            entries = [(f, s, mt, ino, sha) for f, s, mt, ino, sha in entries]
            fresh = []
            for f, s, mt, ino, sha in entries:
                st = os.stat(os.path.join(dest, f))
                fresh.append((f, s, st.st_mtime_ns, st.st_ino, sha))
            save_manifest(self.mdir, "library", name, meta, fresh)
            drop_manifest(self.mdir, "waiting", name)
            self.audit.add("approved", name, who, ip, "Published to library.%s"
                           % ((" Replaced: " + ", ".join(replaced) + ".") if replaced else ""))
        self.wake.set()
        return "OK - Approved %s%s" % (name, " (replaced %s)" % ", ".join(replaced) if replaced else "")

    def reject(self, area, name, who, ip):
        with self.lock:
            p = self.find(area, name)
            if not p or not os.path.exists(p.path):
                return "FAIL - %s not found" % name
            dest = self.to_trash(area, name, "rejected" if area == "waiting" else ("removed" if area == "library" else "deleted"))
            what = {"waiting": "Rejected", "library": "Removed from library", "inbox": "Deleted from incoming"}[area]
            self.audit.add({"waiting": "rejected", "library": "removed", "inbox": "deleted"}[area], name, who, ip,
                           what + ". Moved to trash: %s" % dest)
        self.wake.set()
        return "OK - %s: %s (moved to trash)" % (what, name)

    def trust(self, area, name, who, ip):
        """An admin has looked at a changed (or unknown) pak and wants it as it is now."""
        with self.lock:
            p = self.find(area, name)
            if not p:
                return "FAIL - %s not found" % name
            self.audit.add("trusted", name, who, ip, "Changes accepted (%s). Re-hashing." % area)
            extra = {"trusted_by": who, "trusted_at": stamp()}
            if area == "library":
                extra.update({"approved_by": who, "approved_ip": ip, "approved_at": stamp()})
            old = load_manifest(self.mdir, area, name)
            if old:
                for k in ("from_ip", "arrived", "approved_by", "approved_ip", "approved_at"):
                    if k in old[0] and k not in extra:
                        extra[k] = old[0][k]
            drop_manifest(self.mdir, area, name)
            self.queue(("fingerprint", area, name, extra))
        self.wake.set()
        return "OK - Changes accepted, re-hashing %s" % name

    def check_all(self, who, ip):
        self.audit.add("check", "", who, ip, "Full verify requested.")
        self.queue(("verify", "", "", {"who": who}))
        return "OK - Full verify queued"

    def flag(self, area, name, issues):
        """The tripwire: a pak has changed since its fingerprints. Out of the Game Library until an admin looks."""
        m = load_manifest(self.mdir, area, name)
        if not m:
            return
        meta, entries = m
        if meta.get("tampered"):
            return
        meta["tampered"] = stamp()
        meta["why"] = "; ".join(issues[:5])
        save_manifest(self.mdir, area, name, meta, entries)
        self.audit.add("tampered", name, "", "", "Integrity check failed (%s): %s.%s"
                       % (area, "; ".join(issues[:5]), " Pulled from library." if area == "library" else ""))

    # ---------------------------------------------------------- the checker (fingerprints and full checks)

    def checker(self):
        while True:
            self.jobs_wake.wait(5)
            self.jobs_wake.clear()
            while True:
                with self.lock:
                    if not self.jobs:
                        break
                    job = self.jobs.popleft()
                try:
                    if job[0] == "fingerprint":
                        self.do_fingerprint(*job[1:])
                    elif job[0] == "verify":
                        self.do_verify(job[3])
                except Exception as e:
                    log("Integrity checker error:", repr(e))
                finally:
                    self.job_now = None
                self.wake.set()

    def do_fingerprint(self, area, name, extra):
        path = self.area_path(area, name)
        if not os.path.isdir(path):
            return
        self.job_now = {"what": "Fingerprinting", "pak": name, "done": 0, "total": 1}

        def progress(done, total):
            self.job_now = {"what": "Fingerprinting", "pak": name, "done": done, "total": total}
        before = folder_stamp(path)
        entries = fingerprint(path, progress)
        with self.lock:
            if not os.path.isdir(path):
                return
            if folder_stamp(path) != before:
                log("  %s changed during hashing. Hashing again." % name)
                self.queue(("fingerprint", area, name, extra))
                return
            meta = dict(extra or {})
            meta.update({"fingerprinted": stamp(), "digest": entries_digest(entries), "area": area})
            save_manifest(self.mdir, area, name, meta, entries)
            self.audit.add("fingerprinted", name, extra.get("trusted_by", ""), "",
                           "Manifest created: %d files, %s, digest %s." % (len(entries), size_text(sum(e[1] for e in entries)),
                                                                        meta["digest"][:16]))

    def do_verify(self, extra):
        todo = []
        with self.lock:
            for area in ("library", "waiting"):
                for p in self.view.get(area, []):
                    m = load_manifest(self.mdir, area, p.rel)
                    if m and not m[0].get("tampered"):
                        todo.append((area, p.rel, m[1]))
        total = sum(sum(e[1] for e in entries) for _, _, entries in todo) or 1
        done, bad = 0, []
        for area, name, entries in todo:
            self.job_now = {"what": "Checking", "pak": name, "done": done, "total": total}
            issues = full_check(self.area_path(area, name), entries)
            done += sum(e[1] for e in entries)
            if issues:
                with self.lock:
                    self.flag(area, name, issues)
                bad.append(name)
        self.last_full = stamp()
        self.audit.add("checked", "", extra.get("who", ""), "",
                       "Full verify: %d cartridge(s), %s." % (len(todo), "all OK" if not bad else "FAILED: " + ", ".join(bad)))

    # ---------------------------------------------------------- Samba's audit: who did what on the share

    def share_audit_reader(self):
        cmd = ["journalctl", "-f", "-n", "0", "-o", "cat", "-t", "smbd_audit"]
        try:
            p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, errors="replace")
        except OSError as e:
            self.share_audit = "off (unable to run journalctl: %s)" % e
            return
        self.share_audit = "on"
        for line in p.stdout:
            try:
                self.share_line(line.strip())
            except Exception as e:
                log("Unable to parse SMB audit line:", repr(e), line[:200])
        err = p.stderr.read().strip().splitlines()
        self.share_audit = "off (%s)" % (err[0] if err else "journalctl exited")
        log("SMB audit stopped:", self.share_audit)

    def share_line(self, line):
        """A line from Samba's full_audit: ip|op|ok|...|path"""
        parts = line.split("|")
        if len(parts) < 4 or parts[2] != "ok":
            return
        ip, op = parts[0], parts[1]
        if op == "openat":
            if parts[3] != "w":
                return
            what, path = "wrote", parts[-1]
        elif op == "mkdirat":
            what, path = "made folders", parts[3]
        elif op == "unlinkat":
            what, path = "deleted", parts[3]
        elif op == "renameat":
            what, path = "renamed", parts[3]
        else:
            return
        inbox = self.dirs["inbox"].rstrip("/") + "/"
        if not path.startswith(inbox):
            return
        rel = path[len(inbox):]
        top = rel.split("/")[0]
        if not top:
            return
        now = time.time()
        self.share_seen = now
        with self.lock:
            if what in ("wrote", "made folders", "renamed"):
                w = self.writers.setdefault(top, {})
                w[ip] = w.get(ip, 0) + 1
            ev = self.share_events.setdefault((ip, top, what), [0, now, now])
            ev[0] += 1
            ev[2] = now

    def flush_share_events(self, now):
        """Share activity goes in the audit log in batches: one line per PC, folder and kind of change."""
        for key in list(self.share_events):
            count, first, last = self.share_events[key]
            if now - last >= 20:
                ip, top, what = key
                del self.share_events[key]
                text = "%s: %d %s" % ({"wrote": "write", "made folders": "mkdir", "deleted": "delete", "renamed": "rename"}[what],
                                      count, "folder(s)" if what == "made folders" else "file(s)")
                self.audit.add("share", top, "", ip, text)

    # ---------------------------------------------------------- looking at everything

    def tidy_loose(self):
        """Paks from before 2.0 sat straight in the paks folder: into the inbox they go."""
        try:
            for e in os.scandir(self.base):
                if e.is_dir() and e.name not in AREAS and not e.name.startswith("."):
                    dest = os.path.join(self.dirs["inbox"], e.name)
                    if not os.path.exists(dest):
                        os.rename(e.path, dest)
                        log("Moved %s to incoming" % e.name)
        except OSError as e:
            log("Unable to tidy the paks folder:", e)

    def once(self):
        now = time.time()
        with self.lock:
            self.tidy_loose()
            self.flush_share_events(now)
            # the inbox: finished paks are locked for review once nothing's changed for a while
            inbox = find_paks("inbox", self.dirs["inbox"])
            locked = False
            for p in inbox:
                look_at(p, self.memory, now)
                if p.status == "ready":
                    stamp_now = folder_stamp(p.path)
                    seen = self.settle.get(p.rel)
                    if not seen or seen[0] != stamp_now:
                        self.settle[p.rel] = (stamp_now, now)
                        seen = self.settle[p.rel]
                    if now - seen[1] >= SETTLE_SECONDS:
                        try:
                            self.lock_for_review(p)
                            locked = True
                            continue
                        except OSError as e:
                            p.status, p.message = "problem", "Move to waiting failed: %s" % e
                    else:
                        p.status, p.message = "settling", "Complete. Moving to waiting in %d s" % max(1, int(SETTLE_SECONDS - (now - seen[1])))
            if locked:
                inbox = [p for p in inbox if os.path.isdir(p.path)]
            self.settle = {k: v for k, v in self.settle.items() if any(p.rel == k for p in inbox)}
            self.memory = {k: v for k, v in self.memory.items() if any(p.rel == k for p in inbox)}

            # waiting and library: what the fingerprints say
            waiting = find_paks("waiting", self.dirs["waiting"])
            library = find_paks("library", self.dirs["library"])
            busy = {(j[1], j[2]) for j in self.jobs if j[0] == "fingerprint"}
            if self.job_now and self.job_now["what"] == "Fingerprinting":
                busy.add(("?", self.job_now["pak"]))
            for p in waiting + library:
                look_at(p, {}, now)
                m = load_manifest(self.mdir, p.area, p.rel)
                fingerprinting = (p.area, p.rel) in busy or ("?", p.rel) in busy
                if m is None:
                    p.meta = {}
                    if fingerprinting:
                        p.message = "checking"
                    elif p.area == "waiting":
                        self.queue(("fingerprint", "waiting", p.rel, {}))
                        p.message = "checking"
                    else:
                        p.message = "unapproved"
                    continue
                p.meta, entries = m
                if not p.meta.get("tampered"):
                    issues = quick_check(p.path, entries)
                    if issues:
                        self.flag(p.area, p.rel, issues)
                        p.meta = load_manifest(self.mdir, p.area, p.rel)[0]
                if p.meta.get("tampered"):
                    p.message = "tampered"
                elif p.status != "ready":
                    p.meta["why"] = p.message or "cartridge incomplete"
                    p.message = "tampered"
                else:
                    p.message = "live" if p.area == "library" else "review"

            # one game per id in the Game Library: if there are two, the newest one
            best = {}
            for p in library:
                if p.message == "live":
                    other = best.get(p.info["id"])
                    if other is None or p.info.get("made", "") > other.info.get("made", ""):
                        best[p.info["id"]] = p
            for p in library:
                if p.message == "live" and best[p.info["id"]] is not p:
                    p.message = "hidden"
            ready = sorted(best.values(), key=lambda p: p.info.get("name", "").lower())

            # icons, for the Game Library and the admin page
            shown = ready + [p for p in waiting if p.status == "ready"]
            self.icons = {k: v for k, v in self.icons.items() if any(p.sig == k[0] for p in shown)}
            for p in shown:
                i = p.info
                if not p.games:
                    key = (p.sig, "")
                    if key not in self.icons:
                        self.icons[key] = make_icon(self.state, p, i.get("icon", ""), i.get("icon_from", ""),
                                                    "%s-%s" % (i["id"], p.sig))
                    p.icon, p.png = self.icons[key]
                for g in p.games:
                    key = (p.sig, g["id"])
                    if key not in self.icons:
                        self.icons[key] = make_icon(self.state, p, g.get("icon", ""),
                                                    g.get("icon_from", "") or g.get("exe", ""),
                                                    "%s-%s-%s" % (i["id"], g["id"], p.sig))
                    g["icon_url"], g["png"] = self.icons[key]
                if p.games and not p.png:
                    p.png = p.games[0]["png"]
                if p.sig not in self.shots:
                    self.shots[p.sig] = make_shot(self.state, p, i.get("screenshot", ""), "%s-%s" % (i["id"], p.sig))
                p.shot = self.shots[p.sig]
            self.shots = {k: v for k, v in self.shots.items() if any(p.sig == k for p in shown)}

            self.view = {"inbox": inbox, "waiting": waiting, "library": library}
            self.ready = ready
            for p in inbox + waiting + library:
                key = p.area + "/" + p.rel
                st = (p.status if p.area == "inbox" else p.message, p.message)
                if self.last.get(key) != st:
                    log("%-8s %-28s %s %s" % (p.area, p.rel, st[0], p.message if p.area == "inbox" else ""))
                    self.last[key] = st
            for k in [k for k in self.last if k not in {p.area + "/" + p.rel for p in inbox + waiting + library}]:
                del self.last[k]

            catalog = "\r\n".join(catalog_lines(ready, self.url)) + "\r\n"
            old = ""
            try:
                with open(os.path.join(self.state, "catalog.txt"), "rb") as f:
                    old = f.read().decode("cp1252", "replace")
            except OSError:
                pass
            if old.split("\r\n", 1)[-1] != catalog.split("\r\n", 1)[-1] or not old:
                write_if_changed(os.path.join(self.state, "catalog.txt"), catalog, "cp1252")
                log("Catalog updated: %d cartridge(s) in the library" % len(ready))
                keep = {p.icon[6:-4] for p in shown if p.icon}
                keep |= {g["icon_url"][6:-4] for p in shown for g in p.games if g["icon_url"]}
                keep |= {p.shot[6:-9] for p in shown if p.shot}
                tidy_icons(self.state, keep)
            key = tuple((p.area, p.rel, p.status, p.message) for p in inbox + waiting + library) + (len(ready),)
            if key != self.page_key:
                write_if_changed(os.path.join(self.state, "index.html"), public_page(self), "cp1252")
                self.page_key = key
        self.daily(now)
        return ready

    def daily(self, now):
        today = time.strftime("%Y-%m-%d")
        # every night, every pak read right through
        if time.localtime(now).tm_hour == self.settings["check_hour"] and self.last_full_day != today:
            self.last_full_day = today
            self.queue(("verify", "", "", {"who": "(scheduler)"}))
        # the trash, emptied of anything older than a week
        if self.last_purge_day != today:
            self.last_purge_day = today
            limit = time.time() - self.settings["trash_days"] * 86400
            try:
                for e in os.scandir(self.dirs["trash"]):
                    m = re.search(r"__(\d{8}-\d{6})__", e.name)
                    if m and time.mktime(time.strptime(m.group(1), "%Y%m%d-%H%M%S")) < limit:
                        for dirpath, dirs, files in os.walk(e.path):
                            os.chmod(dirpath, 0o755)
                        shutil.rmtree(e.path, ignore_errors=True)
                        log("Deleted %s from the trash" % e.name)
            except OSError:
                pass

    def startup(self):
        """Paks approved by the setup script (the ones already here when it was upgraded) get their
        fingerprints, and every approved pak is read right through once, in case anything changed
        while the service was stopped."""
        listing = os.path.join(self.state, "approve-on-start.txt")
        try:
            with open(listing, encoding="utf-8") as fh:
                names = [l.strip() for l in fh if l.strip()]
            os.remove(listing)
        except OSError:
            names = []
        for name in names:
            if os.path.isdir(self.area_path("library", name)) and not load_manifest(self.mdir, "library", name):
                self.audit.add("approved", name, "setup", "", "Pre-2.0 cartridge, approved by setup script.")
                self.queue(("fingerprint", "library", name, {"approved_by": "setup", "approved_at": stamp(),
                                                              "approved_ip": "", "arrived": "before approvals"}))
        self.queue(("verify", "", "", {"who": "(startup)"}))

    def run(self):
        log("Palms Cartridge Server %s started. Incoming: %s" % (VERSION, self.dirs["inbox"]))
        if Image is None:
            log("Pillow is not installed. Game icons are disabled (apt install python3-pil).")
        if pefile is None:
            log("pefile is not installed. Icons cannot be extracted from programs (apt install python3-pefile).")
        threading.Thread(target=self.checker, daemon=True).start()
        threading.Thread(target=self.share_audit_reader, daemon=True).start()
        start_admin(self)
        self.startup()
        while True:
            try:
                self.once()
            except Exception as e:      # keep watching whatever happens
                log("Error:", repr(e))
            self.wake.wait(SCAN_SECONDS)
            self.wake.clear()


# ------------------------------------------------------------------ pages
# Plain HTML 4 and tables, so Internet Explorer 6 on the XP PCs shows them the same as anything newer.

e = html.escape
DOCTYPE = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">'
GEN = "<!--GEN-->"      # swapped for how long the page took, when it's sent
CSS = """
body { margin: 0; background: #FFFFFF; color: #000000; font: 11px Verdana, Tahoma, Arial, Helvetica, sans-serif; }
td, th, p, div, input { font-family: Verdana, Tahoma, Arial, Helvetica, sans-serif; font-size: 11px; }
a:link { color: #0000CC; } a:visited { color: #551A8B; } a:hover { color: #CC0000; }
#hdr { background: #31616B; }
#hdr td { color: #FFFFFF; padding: 7px 8px 6px 8px; }
#hdr .t { font: bold 15px Verdana, Arial, sans-serif; }
#hdr .s { color: #BFD6DA; }
#stripe { background: #F78E21; height: 2px; font-size: 1px; line-height: 1px; }
#nav { background: #D4D0C8; border-top: 1px solid #FFFFFF; border-bottom: 1px solid #808080; padding: 4px 8px; }
#nav a, #nav b { padding: 0 5px; }
#main { padding: 8px; }
.crumb { margin: 0 0 8px 0; }
table.g { background: #7F9DA3; margin: 0 0 10px 0; }
table.g th { background: #C5D7DB; color: #000000; text-align: left; white-space: nowrap; }
table.g th.c { background: #31616B; color: #FFFFFF; }
table.g th.c a, table.g th.cw a { color: #FFFFFF; }
table.g th.cw { background: #AA2222; color: #FFFFFF; }
table.g td { background: #FFFFFF; vertical-align: top; }
table.g tr.a td { background: #EEF3F4; }
table.g tr.f td { background: #FFE1E1; }
table.g tr.n td { background: #E1F5E1; }
table.g td.k { background: #E4ECEE; font-weight: bold; white-space: nowrap; }
.m, pre { font-family: "Lucida Console", "Courier New", Courier, monospace; font-size: 11px; }
pre { margin: 0; }
.ok { color: #008000; font-weight: bold; }
.wn { color: #CC6600; font-weight: bold; }
.er { color: #CC0000; font-weight: bold; }
.dim { color: #808080; }
form.i { display: inline; margin: 0; }
input.b { background: #D4D0C8; color: #000000; border: 2px outset #E8E6E0; padding: 1px 8px; cursor: pointer; }
input.bp { font-weight: bold; }
#ft { color: #808080; border-top: 1px solid #C0C0C0; margin: 0 8px; padding: 4px 0 8px 0; font-size: 10px; }
"""

ADMIN_NAV = [("/admin/", "Overview"), ("/admin/log", "Audit Log"), ("/", "Public Index")]
PUBLIC_NAV = [("/", "Index"), ("/admin/", "Administration"), ("/catalog.txt", "catalog.txt"), ("/paks/", "/paks/")]
EVENT_WORDS = {"arrived": "RECEIVED", "fingerprinted": "HASHED", "approved": "APPROVED", "rejected": "REJECTED",
               "removed": "REMOVED", "deleted": "DELETED", "tampered": "MODIFIED", "trusted": "ACCEPTED",
               "checked": "VERIFIED", "check": "VERIFY REQ", "share": "SMB", "login-failed": "AUTH FAIL",
               "locked-out": "LOCKOUT", "superseded": "REPLACED", "admin": "ACCOUNT"}
BAD_EVENTS = ("tampered", "login-failed", "locked-out")
INBOX_WORDS = {"copying": ("COPYING", "wn"), "waiting": ("INCOMPLETE", "dim"), "settling": ("SETTLING", "wn"),
               "problem": ("ERROR", "er"), "ready": ("READY", "ok")}
JOB_WORDS = {"Fingerprinting": "Hashing", "Checking": "Verifying"}


def head(title, refresh=0, charset="utf-8"):
    return ('%s\n<html><head><meta http-equiv="Content-Type" content="text/html; charset=%s">%s<title>%s</title>'
            '<style type="text/css">%s</style></head><body>'
            % (DOCTYPE, charset, '<meta http-equiv="refresh" content="%d">' % refresh if refresh else "", e(title), CSS))


def host_of(srv):
    return (urllib.parse.urlparse(srv.url).hostname if srv else "") or "this-server"


def top(srv, sub, user=""):
    right = ("User: <b>%s</b> &nbsp;|&nbsp; " % e(user) if user else "") + "Host: %s &nbsp;|&nbsp; v%s" % (e(host_of(srv)), VERSION)
    return ('<table id="hdr" width="100%%" cellspacing="0" cellpadding="0"><tr><td><span class="t">Palms Cartridge Server</span>'
            ' &nbsp;<span class="s">%s</span></td><td align="right"><span class="s">%s</span></td></tr></table>'
            '<div id="stripe"></div>' % (e(sub), right))


def nav(items, here):
    return '<div id="nav">%s</div>' % " | ".join(
        "<b>%s</b>" % e(label) if href == here else '<a href="%s">%s</a>' % (e(href), e(label)) for href, label in items)


def foot(extra):
    return '<div id="ft">Palms Cartridge Server %s &middot; %s</div></body></html>' % (VERSION, extra)


def admin_page(srv, user, title, here, body, refresh=0):
    return (head("%s - Palms Cartridge Server" % title, refresh) + top(srv, "Administration", user) + nav(ADMIN_NAV, here)
            + '<div id="main">' + "".join(body) + "</div>" + foot("%s &middot; generated in %s" % (stamp(), GEN)))


def tag(word, cls):
    return '<span class="%s">%s</span>' % (cls, e(word))


def mono(text):
    return '<span class="m">%s</span>' % e(text) if text else "-"


def td(c):
    """A cell: html, or (attributes, html)."""
    if isinstance(c, tuple):
        return "<td%s>%s</td>" % c
    return "<td>%s</td>" % c


def grid(title, heads, rows, empty="No records.", cap="c"):
    """A table in the house style. rows: [cells] or (row class, [cells]); plain rows are striped."""
    width = max([len(heads)] + [len(r[1] if isinstance(r, tuple) else r) for r in rows] + [1])
    out = ['<table class="g" width="100%" cellspacing="1" cellpadding="3">']
    if title:
        out.append('<tr><th class="%s" colspan="%d">%s</th></tr>' % (cap, width, title))
    if heads:
        out.append("<tr>%s</tr>" % "".join("<th>%s</th>" % h for h in heads))
    for n, row in enumerate(rows):
        cls, cells = row if isinstance(row, tuple) else ("", row)
        cls = cls or ("a" if n % 2 else "")
        out.append("<tr%s>%s</tr>" % (' class="%s"' % cls if cls else "", "".join(td(c) for c in cells)))
    if not rows:
        out.append('<tr><td colspan="%d" class="dim">%s</td></tr>' % (width, empty))
    out.append("</table>")
    return "".join(out)


def kv(title, pairs, cap="c"):
    return grid(title, [], [("", [(' class="k" width="150"', k), v]) for k, v in pairs], cap=cap)


def msg_box(text):
    """OK - ... / FAIL - ... / WARNING - ..., from an action or a check."""
    if not text:
        return ""
    word, _, rest = text.partition(" - ")
    cls = {"OK": "ok", "WARNING": "wn"}.get(word, "er")
    return ('<table class="g" width="100%%" cellspacing="1" cellpadding="3"><tr><td class="k" width="60">Message:</td>'
            '<td id="msg">%s - %s</td></tr></table>' % (tag(word, cls), e(rest)))


def icon(png, size=16):
    if not png:
        return ""
    return '<img src="/%s" width="%d" height="%d" alt="" border="0">' % (e(png[:-4] + "-16.png" if size == 16 else png), size, size)


def size_of(p):
    return size_text(to_int(p.info.get("size", ""), 0)) if p.info.get("size") else "-"


def nbytes(n):
    return "{:,}".format(to_int(n, 0)) if str(n).strip() else "-"


def pak_link(p, text=None):
    return '<a href="/admin/pak/%s/%s">%s</a>' % (p.area, e(urllib.parse.quote(p.rel)), e(text or p.rel))


def name_cell(p):
    out = "<b>%s</b>" % e(p.info.get("name", "") or p.rel)
    if p.games:
        out += ' <span class="dim">[multicart: %d]</span>' % len(p.games)
    return out


def status_word(p, job=None):
    if p.message == "checking":
        pct = ""
        if job and job["pak"] == p.rel and job["total"] and job["what"] == "Fingerprinting":
            pct = " %d%%" % (100 * job["done"] // job["total"])
        return tag("HASHING" + pct, "wn")
    return {"review": tag("PENDING", "wn"), "live": tag("ACTIVE", "ok"), "hidden": tag("SUPERSEDED", "dim"),
            "tampered": tag("MODIFIED", "er"), "unapproved": tag("UNAPPROVED", "er")}.get(p.message, e(p.message.upper()))


def public_page(srv):
    """The index everyone can see (a static file nginx serves)."""
    rows = []
    for p in srv.ready:
        i = p.info
        name = "<b>%s</b>" % e(i.get("name", ""))
        for g in p.games:
            name += '<br>&nbsp;&nbsp;%s <span class="dim">%s%s</span>' % (
                icon(g.get("png", "")), e(g["name"]), (" (%s)" % e(g["players"])) if g.get("players") else "")
        rows.append([(' width="16"', icon(p.png)), name, e(ver_text(i)) or "-", e(i.get("players", "")) or "-",
                     (' nowrap', size_of(p)), '<a href="/paks/%s/">%s</a>' % (e(urllib.parse.quote(p.rel)), e(p.rel))])
    queue = []
    for p in srv.view["waiting"]:
        word = {"review": tag("PENDING APPROVAL", "wn"), "checking": tag("HASHING", "wn"),
                "tampered": tag("MODIFIED", "er")}.get(p.message, e(p.message.upper()))
        queue.append([name_cell(p), mono(p.rel), word, ""])
    for p in srv.view["library"]:
        if p.message in ("tampered", "unapproved"):
            queue.append([name_cell(p), mono(p.rel), tag("PULLED", "er"),
                          "Integrity check failed" if p.message == "tampered" else "Not approved"])
    for p in srv.view["inbox"]:
        word, cls = INBOX_WORDS.get(p.status, (p.status.upper(), ""))
        queue.append([name_cell(p) if p.info.get("name") else "-", mono(p.rel), tag(word, cls), e(p.message)])
    host = e(host_of(srv))
    body = [grid("Game Library (%d)" % len(srv.ready), ["", "Name", "Version", "Players", "Size", "Folder"], rows,
                 "Library is empty."),
            grid("Queue (%d)" % len(queue), ["Name", "Folder", "Status", "Info"], queue, "Queue is empty."),
            '<div class="dim">Upload: \\\\%s\\incoming (host approval required) &nbsp;|&nbsp; '
            'Read-only: \\\\%s\\library &nbsp;|&nbsp; Auto-refresh: 10 s</div>' % (host, host)]
    return (head("Palms Cartridge Server", 10, "windows-1252") + top(srv, "Game Library") + nav(PUBLIC_NAV, "/")
            + '<div id="main">' + "".join(body) + "</div>" + foot("last change %s" % stamp()))


def read_small(p, name):
    """A small file from a pak (pak.ini, setup.reg, notes.txt) as text, or None."""
    real = files_in(p.path).get(name.lower())
    if not real:
        return None
    try:
        if os.path.getsize(os.path.join(p.path, real)) > 512 * 1024:
            return "(not shown: %s)" % size_text(os.path.getsize(os.path.join(p.path, real)))
        return read_text(os.path.join(p.path, real))
    except OSError:
        return None


def form(action, area, name, label, srv, user, confirm, bold=False):
    js = ' onclick="return confirm(\'%s\')"' % e(confirm.replace("\\", "\\\\").replace("'", "\\'")) if confirm else ""
    return ('<form class="i" method="post" action="/admin/do"><input type="hidden" name="action" value="%s">'
            '<input type="hidden" name="area" value="%s"><input type="hidden" name="name" value="%s">'
            '<input type="hidden" name="token" value="%s"><input type="submit" class="b%s" value="%s"%s></form>'
            % (e(action), e(area), e(name), token(srv, user), " bp" if bold else "", e(label), js))


def token(srv, user, day=None):
    day = day or time.strftime("%Y%m%d")
    return hmac.new(srv.secret, ("%s|%s" % (user, day)).encode(), hashlib.sha256).hexdigest()[:24]


def audit_table(rows, title):
    out = []
    for r in rows:
        when, ev, pak, who, ip, what = (r + [""] * 6)[:6]
        word = EVENT_WORDS.get(ev, ev.upper())
        cls = "er" if ev in BAD_EVENTS else ("ok" if ev == "approved" else "")
        out.append(("f" if ev in BAD_EVENTS else "",
                    [(' nowrap', e(when)), (' nowrap', tag(word, cls) if cls else e(word)), (' nowrap', e(pak) or "-"),
                     e(who) or "-", e(ip) or "-", e(what)]))
    return grid(title, ["Time", "Event", "Cartridge", "User", "IP", "Details"], out, "No events.")


def admin_dashboard(srv, user):
    with srv.lock:
        view = {k: list(v) for k, v in srv.view.items()}
        job = srv.job_now
        queued = len(srv.jobs)
    body = [msg_box(srv.flash.pop(user, ""))]

    alerts = [p for a in ("library", "waiting") for p in view[a] if p.message in ("tampered", "unapproved")]
    if alerts:
        body.append(grid("Integrity Alerts (%d)" % len(alerts), ["Cartridge", "Area", "Detected", "Reason"],
                         [("f", [pak_link(p), p.area, (' nowrap', e(p.meta.get("tampered", "")) or "-"),
                                 e(p.meta.get("why", "")) if p.message == "tampered" else "No manifest (not approved here)"])
                          for p in alerts], cap="cw"))

    rows = []
    for p in view["waiting"]:
        rows.append(("f" if p.message == "tampered" else "",
                     [(' width="16"', icon(p.png)), name_cell(p), mono(p.rel), e(ver_text(p.info)) or "-",
                      (' nowrap', size_of(p)), e(p.meta.get("from_ip", "")) or "-", (' nowrap', e(p.meta.get("arrived", "")) or "-"),
                      (' nowrap', status_word(p, job)), pak_link(p, "Review")]))
    body.append(grid("Pending Approval (%d)" % len(rows),
                     ["", "Name", "Folder", "Version", "Size", "Source IP", "Received", "Status", "Action"], rows,
                     "No cartridges pending."))

    rows = []
    for p in view["library"]:
        by = e(p.meta.get("approved_by", "")) or "-"
        if p.meta.get("approved_ip"):
            by += ' <span class="dim">(%s)</span>' % e(p.meta["approved_ip"])
        rows.append(("f" if p.message in ("tampered", "unapproved") else "",
                     [(' width="16"', icon(p.png)), name_cell(p), mono(p.rel), mono(p.info.get("id", "")),
                      e(ver_text(p.info)) or "-", (' nowrap', size_of(p)),
                      (' nowrap', e(p.meta.get("approved_at", "")) or "-"), (' nowrap', by), (' nowrap', status_word(p, job)),
                      pak_link(p, "Details")]))
    body.append(grid("Library (%d active)" % len([p for p in view["library"] if p.message == "live"]),
                     ["", "Name", "Folder", "ID", "Version", "Size", "Approved", "By", "Status", "Action"], rows,
                     "Library is empty."))

    rows = []
    for p in view["inbox"]:
        word, cls = INBOX_WORDS.get(p.status, (p.status.upper(), ""))
        rows.append([mono(p.rel), e(p.info.get("name", "")) or "-", (' nowrap', tag(word, cls)), e(p.message) or "-",
                     form("delete", "inbox", p.rel, "Delete", srv, user, "Delete %s from incoming? (moved to trash)" % p.rel)])
    body.append(grid("Incoming: \\\\%s\\incoming (%d)" % (e(host_of(srv)), len(rows)), ["Folder", "Name", "State", "Info", "Action"],
                     rows, "No uploads."))

    if job:
        pct = 100 * job["done"] // job["total"] if job["total"] else 0
        now = "%s %s: %d%%" % (JOB_WORDS.get(job["what"], job["what"]), e(job["pak"]), pct)
    else:
        now = "Idle"
    share = srv.share_audit
    if share == "on":
        share = tag("ENABLED", "ok") + (" (last event %s)" % stamp(srv.share_seen)[11:] if srv.share_seen else " (no events yet)")
    elif share.startswith("off"):
        share = tag("DISABLED", "er") + " " + e(share[3:].strip())
    else:
        share = e(share.upper())
    body.append(kv("Integrity Checker", [
        ("Quick check", "Every %d s (size, mtime, inode)" % SCAN_SECONDS),
        ("Full verify", "Daily at %02d:00 (SHA-256, all files)" % srv.settings["check_hour"]),
        ("Last full verify", e(srv.last_full or "Never")),
        ("Current job", now),
        ("Queued jobs", str(queued)),
        ("SMB audit", share),
        ("Action", form("check", "", "", "Run Full Verify", srv, user, "Re-hash every cartridge now?"))]))

    body.append(audit_table(srv.audit.recent(15), "Recent Events (last 15)"))
    body.append('<a href="/admin/log">Full audit log &raquo;</a>')
    return admin_page(srv, user, "Overview", "/admin/", body, 15 if (job or queued) else 30)


def admin_log(srv, user):
    rows = srv.audit.recent(1000)
    return admin_page(srv, user, "Audit Log", "/admin/log", [audit_table(rows, "Audit Log (last %d)" % len(rows))])


def admin_pak(srv, user, area, name):
    with srv.lock:
        p = srv.find(area, name)
        live = [q for q in srv.view["library"] if p and q.rel != p.rel and q.info.get("id") == p.info.get("id")]
    if not p:
        return None
    i = p.info
    worries = 0
    days = srv.settings["trash_days"]

    # properties and actions
    acts = []
    if area == "waiting" and p.message == "review":
        acts.append(form("approve", area, name, "Approve", srv, user, "Approve %s? It will be published to all clients." % name, True))
    if p.message in ("tampered", "unapproved"):
        acts.append(form("trust", area, name, "Accept Changes", srv, user,
                         "Accept the current state of %s and create a new manifest?" % name))
    if area == "waiting":
        acts.append(form("reject", area, name, "Reject", srv, user, "Reject %s? (moved to trash, kept %d days)" % (name, days)))
    if area == "library":
        acts.append(form("remove", area, name, "Remove", srv, user,
                         "Remove %s from the library? (moved to trash, kept %d days)" % (name, days)))
    if p.message == "tampered":
        integrity = tag("FAILED", "er") + " " + e(p.meta.get("why", ""))
    elif p.message == "unapproved":
        integrity = tag("NO MANIFEST", "er")
    elif p.message == "checking":
        integrity = tag("HASHING", "wn")
    else:
        integrity = tag("OK", "ok") + " (manifest %s)" % e(p.meta.get("fingerprinted", ""))
    pairs = []
    if p.png:
        pairs.append(("Icon", icon(p.png, 32)))
    if p.shot:
        pairs.append(("Screenshot", '<img src="/%s" width="320" alt="" border="0">' % e(p.shot)))
    pairs += [("Name", "<b>%s</b>" % e(i.get("name", ""))),
              ("ID", mono(i.get("id", ""))),
              ("Version", e(ver_text(i)) or "-"),
              ("Players", e(i.get("players", "")) or "-"),
              ("Install folder", mono(i.get("folder", ""))),
              ("Type", "Multicart (%d games)" % len(p.games) if p.games else "Single game"),
              ("Installed size", "%s (%s files)" % (size_of(p), e(i.get("files", "?")))),
              ("Archive", "%s, %s piece(s), %s bytes" % (e(i.get("archive") or "game.7z"), e(i.get("pieces", "?")),
                                                         nbytes(i.get("archive_bytes", "")))),
              ("Made", e(i.get("made", "")) or "-")]
    if p.meta.get("arrived"):
        pairs.append(("Received", e(p.meta["arrived"]) + (" from %s" % e(p.meta["from_ip"]) if p.meta.get("from_ip") else "")))
    if p.meta.get("approved_by"):
        pairs.append(("Approved", "%s by <b>%s</b>%s" % (e(p.meta.get("approved_at", "")), e(p.meta["approved_by"]),
                                                       " from %s" % e(p.meta["approved_ip"]) if p.meta.get("approved_ip") else "")))
    pairs += [("Status", status_word(p)), ("Integrity", integrity)]
    if acts:
        pairs.append(("Action", " ".join(acts)))
    sections = [kv("Cartridge Properties: %s" % e(name), pairs)]

    # executables
    rows = []
    games = p.games or [{"name": i.get("name", ""), "players": i.get("players", ""), "exe": i.get("exe", ""),
                         "args": i.get("args", ""), "start_in": "", "config_exe": i.get("config_exe", ""),
                         "config_args": i.get("config_args", "")}]
    for g in games:
        w = program_worries(g.get("exe", ""), g.get("args", ""))
        worries += len(w)
        rows.append(("f" if w else "", [e(g.get("name", "")), e(g.get("players", "")) or "-",
                                        mono((g.get("exe", "") + " " + g.get("args", "")).strip()), mono(g.get("start_in", "")),
                                        "<br>".join(tag(x, "er") for x in w)]))
        if g.get("config_exe"):
            w = program_worries(g["config_exe"], g.get("config_args", ""))
            worries += len(w)
            rows.append(("f" if w else "", ['<span class="dim">%s (configuration)</span>' % e(g.get("name", "")), "",
                                            mono((g["config_exe"] + " " + g.get("config_args", "")).strip()), "",
                                            "<br>".join(tag(x, "er") for x in w)]))
    sections.append(grid("Executables (%d)" % len(rows), ["Title", "Players", "Command line", "Working dir", "Warnings"], rows))

    # the registry
    reg = read_small(p, i.get("reg", "")) if i.get("reg") else None
    if reg is None:
        sections.append(grid("Registry", ["#", "Line", "Warning"], [], "No registry file."))
    else:
        lines = check_reg(reg)
        flagged = len([x for x in lines if x[2]])
        worries += flagged
        sections.append(grid("Registry: %s (%d lines, %d flagged)" % (e(i.get("reg", "")), len(lines), flagged),
                             ["#", "Line", "Warning"],
                             [("f" if why else "", [(' align="right" class="dim" width="20"', str(n)), (' class="m"', e(line) or "&nbsp;"),
                                                    tag(why, "er") if why else ""]) for n, line, why in lines]))

    # notes
    notes = read_small(p, i.get("notes", "")) if i.get("notes") else None
    sections.append(grid("Setup Notes: %s" % e(i.get("notes", "")) if notes is not None else "Setup Notes", [],
                         [[("", "<pre>%s</pre>" % e(notes))]] if notes is not None else [], "No setup notes."))

    # against the active pak with the same id
    for q in live:
        rows = []
        for k in ("name", "version", "revision", "players", "folder", "exe", "args", "config_exe", "config_args", "files", "size"):
            a, b = q.info.get(k, ""), i.get(k, "")
            if a != b:
                rows.append([mono(k), mono(a), mono(b)])
        old_games = {g["id"]: g for g in q.games}
        new_games = {g["id"]: g for g in p.games}
        for gid in sorted(set(old_games) | set(new_games)):
            a, b = old_games.get(gid), new_games.get(gid)
            fa = "%s: %s %s" % (a["name"], a["exe"], a["args"]) if a else "(none)"
            fb = "%s: %s %s" % (b["name"], b["exe"], b["args"]) if b else "(removed)"
            if fa != fb:
                rows.append([mono("game " + gid), mono(fa), mono(fb)])
        old_reg = read_small(q, q.info.get("reg", "")) if q.info.get("reg") else ""
        diff = [d for d in list(difflib.unified_diff((old_reg or "").splitlines(), (reg or "").splitlines(), lineterm="", n=0))[2:]
                if not d.startswith("@@")]
        rows.append([mono("setup.reg"), ("", '<span class="dim">-</span>'), "changed (%d lines)" % len(diff) if diff else "identical"])
        sections.append(grid("Diff: %s (%s) &rarr; %s" % (e(q.rel), {"live": "active"}.get(q.message, e(q.message)), e(name)),
                             ["Field", "Active: " + pak_link(q), "This cartridge"], rows))
        if diff:
            sections.append(grid("Registry Diff", [], [("n" if d.startswith("+") else "f", [(' class="m"', e(d))]) for d in diff]))

    # the manifest
    m = load_manifest(srv.mdir, area, name)
    if m:
        sections.append(kv("Manifest", [("Cartridge digest (SHA-256)", mono(m[0].get("digest", ""))),
                                         ("Created", e(m[0].get("fingerprinted", ""))), ("Files", str(len(m[1])))]))
        sections.append(grid("", ["File", "Bytes", "SHA-256"],
                             [[mono(f), (' align="right" nowrap', nbytes(size)), mono(sha)] for f, size, _, _, sha in m[1]]))
    else:
        sections.append(grid("Manifest", [], [], "No manifest yet."))
    sections.append(grid("pak.ini", [], [[("", "<pre>%s</pre>" % e(read_small(p, "pak.ini") or ""))]]))
    sections.append(audit_table(srv.audit.recent(100, name), "Event History"))

    body = ['<div class="crumb"><a href="/admin/">Overview</a> &raquo; %s &raquo; <b>%s</b></div>'
            % ({"waiting": "Pending Approval", "library": "Library"}[area], e(name))]
    if worries and area == "waiting" and p.message == "review":
        body.append(msg_box("WARNING - %d item(s) flagged. Check the highlighted rows before approving." % worries))
    return admin_page(srv, user, name, "", body + sections)


def error_page(srv, user, code, text):
    title = {401: "401 Unauthorized", 403: "403 Forbidden", 404: "404 Not Found"}[code]
    body = [grid(title, [], [[t] for t in text], cap="cw")]
    if user:
        return admin_page(srv, user, title, "", body)
    return (head("%s - Palms Cartridge Server" % title) + top(srv, "Administration") + nav([("/", "Public Index")], "")
            + '<div id="main">' + "".join(body) + "</div>" + foot(stamp()))


# ------------------------------------------------------------------ the admin web page
# Behind nginx (http://<this box>/admin). Logins use HTTP Digest, which Internet Explorer 6
# knows: the password never crosses the network, only a one-off answer worked out from it.

SRV = None


class AdminHandler(BaseHTTPRequestHandler):
    server_version = "PalmsPaks/" + VERSION
    protocol_version = "HTTP/1.0"

    def log_message(self, fmt, *args):
        pass

    def ip(self):
        peer = self.client_address[0]
        if peer in ("127.0.0.1", "::1"):
            return self.headers.get("X-Real-IP", peer)
        return peer

    def send_page(self, code, body, extra=None):
        data = body.replace(GEN, "%.3f s" % (time.time() - getattr(self, "t0", time.time()))).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def challenge(self, stale=False):
        nonce = base64.b64encode(os.urandom(18)).decode()
        with SRV.lock:
            now = time.time()
            for n in [n for n, v in SRV.nonces.items() if now - v[0] > 12 * 3600]:
                del SRV.nonces[n]
            SRV.nonces[nonce] = [now, set()]
        hdr = 'Digest realm="%s", qop="auth", algorithm=MD5, nonce="%s", opaque="palms"%s' % (
            REALM, nonce, ", stale=TRUE" if stale else "")
        self.send_page(401, error_page(SRV, "", 401, ["A host account is required.",
                                                      "Accounts are managed on the server: " + mono("palms_paks.py --add-admin NAME")]),
                       {"WWW-Authenticate": hdr})

    def locked_out(self, ip):
        """Seconds until this PC can sign in again, or 0."""
        now = time.time()
        with SRV.lock:
            times = sorted(t for t in SRV.fails.get(ip, []) if now - t < LOCKOUT_SECONDS)
            SRV.fails[ip] = times
            if len(times) < LOCKOUT_TRIES:
                return 0
            return max(1, int(times[-LOCKOUT_TRIES] + LOCKOUT_SECONDS - now))

    def failed(self, ip, name):
        with SRV.lock:
            SRV.fails.setdefault(ip, []).append(time.time())
            n = len(SRV.fails[ip])
        SRV.audit.add("login-failed", "", name, ip, "Bad username or password.")
        if n == LOCKOUT_TRIES:
            SRV.audit.add("locked-out", "", name, ip, "%d failed logins. Locked out for %d min."
                          % (LOCKOUT_TRIES, LOCKOUT_SECONDS // 60))

    def who(self):
        """The admin signed in, or None (after sending the sign-in challenge or the locked-out page)."""
        ip = self.ip()
        wait = self.locked_out(ip)
        if wait:
            self.send_page(403, error_page(SRV, "", 403, ["Too many failed logins from %s. Retry in %d min." % (e(ip), (wait + 59) // 60)]))
            return None
        h = self.headers.get("Authorization", "")
        if not h.lower().startswith("digest "):
            self.challenge()
            return None
        f = {k.lower(): (v[1:-1] if v.startswith('"') else v)
             for k, v in re.findall(r'(\w+)\s*=\s*("(?:[^"\\]|\\.)*"|[^,\s]*)', h[7:])}
        name = f.get("username", "")
        admins = read_admins(SRV.settings["admins"])
        nonce = f.get("nonce", "")
        with SRV.lock:
            known = SRV.nonces.get(nonce)
        if known is None:
            self.challenge(stale=True)
            return None
        uri = f.get("uri", "")
        if uri.split("?")[0] != self.path.split("?")[0]:
            self.challenge()
            return None
        ha1 = admins.get(name)
        ha2 = hashlib.md5(("%s:%s" % (self.command, uri)).encode()).hexdigest()
        qop = f.get("qop", "")
        if ha1:
            if qop:
                want = hashlib.md5(("%s:%s:%s:%s:%s:%s" % (ha1, nonce, f.get("nc", ""), f.get("cnonce", ""), qop, ha2)).encode()).hexdigest()
            else:
                want = hashlib.md5(("%s:%s:%s" % (ha1, nonce, ha2)).encode()).hexdigest()
        if not ha1 or not hmac.compare_digest(want, f.get("response", "")):
            self.failed(ip, name)
            self.challenge()
            return None
        nc = f.get("nc", "")
        with SRV.lock:
            if qop and nc in known[1]:
                replay = True
            else:
                replay = False
                known[1].add(nc)
        if replay:
            self.challenge(stale=True)
            return None
        return name

    def do_GET(self):
        self.t0 = time.time()
        user = self.who()
        if not user:
            return
        path = urllib.parse.unquote(self.path.split("?")[0])
        if path in ("/admin", "/admin/"):
            self.send_page(200, admin_dashboard(SRV, user))
            return
        if path == "/admin/log":
            self.send_page(200, admin_log(SRV, user))
            return
        m = re.match(r"^/admin/pak/(waiting|library)/([^/]+)$", path)
        if m:
            page = admin_pak(SRV, user, m.group(1), m.group(2))
            if page:
                self.send_page(200, page)
                return
        self.send_page(404, error_page(SRV, user, 404, ["No such page: " + mono(path)]))

    def do_POST(self):
        self.t0 = time.time()
        user = self.who()
        if not user:
            return
        n = min(to_int(self.headers.get("Content-Length", "0"), 0), 65536)
        f = {k: v[0] for k, v in urllib.parse.parse_qs(self.rfile.read(n).decode("utf-8", "replace")).items()}
        ip = self.ip()
        good = f.get("token", "") in (token(SRV, user), token(SRV, user, time.strftime("%Y%m%d", time.localtime(time.time() - 86400))))
        if not good:
            msg = "FAIL - Form token invalid or expired. No action taken."
        else:
            action, area, name = f.get("action", ""), f.get("area", ""), f.get("name", "")
            if action == "approve":
                msg = SRV.approve(name, user, ip)
            elif action in ("reject", "remove", "delete") and area in ("waiting", "library", "inbox"):
                msg = SRV.reject(area, name, user, ip)
            elif action == "trust" and area in ("waiting", "library"):
                msg = SRV.trust(area, name, user, ip)
            elif action == "check":
                msg = SRV.check_all(user, ip)
            else:
                msg = "FAIL - Unknown action"
        SRV.flash[user] = msg
        self.send_response(302)
        self.send_header("Location", "/admin/")
        self.send_header("Content-Length", "0")
        self.end_headers()


def start_admin(srv):
    global SRV
    SRV = srv
    try:
        httpd = ThreadingHTTPServer(("127.0.0.1", srv.settings["admin_port"]), AdminHandler)
    except OSError as err:
        log("Unable to start the admin page on port %d: %s" % (srv.settings["admin_port"], err))
        return
    httpd.daemon_threads = True
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    log("Admin page listening on 127.0.0.1:%d (http://<server>/admin via nginx)" % srv.settings["admin_port"])


# ------------------------------------------------------------------ admins, from the command line

def own_by_palms(path):
    """A file root made that the service (running as palms) has to keep writing to."""
    if hasattr(os, "geteuid") and os.geteuid() == 0:
        try:
            import pwd
            u = pwd.getpwnam("palms")
            os.chown(path, u.pw_uid, u.pw_gid)
        except (KeyError, OSError, ImportError):
            pass


def admin_command(settings, args):
    path = settings["admins"]
    admins = read_admins(path)
    audit = Audit(os.path.join(settings["state"], "audit.log"))
    if args[0] == "--admins":
        print("\n".join(sorted(admins)) or "No host accounts. To add one: palms_paks.py --add-admin NAME")
        return 0
    if len(args) < 2 or not re.match(r"^[A-Za-z0-9._-]{1,32}$", args[1]):
        print("Account name required: letters, numbers, . _ or - (for example, host).")
        return 1
    name = args[1]
    if args[0] == "--remove-admin":
        if name not in admins:
            print("Host account %s does not exist." % name)
            return 1
        del admins[name]
        write_admins(path, admins)
        audit.add("admin", "", name, "", "Account removed (CLI).")
        own_by_palms(audit.path)
        print("Host account %s removed." % name)
        return 0
    while True:
        pw = getpass.getpass("Password for %s: " % name)
        if len(pw) < 6:
            print("The password must be at least 6 characters long.")
            continue
        if getpass.getpass("Confirm password: ") != pw:
            print("The passwords do not match.")
            continue
        break
    was = name in admins
    admins[name] = ha1_of(name, pw)
    write_admins(path, admins)
    try:
        audit.add("admin", "", name, "", "Password changed (CLI)." if was else "Account added (CLI).")
        own_by_palms(audit.path)
    except Exception:
        pass
    print(("Password changed for %s." if was else "Host account %s added. Log in at http://<server>/admin") % name)
    return 0


def main():
    args = [a for a in sys.argv[1:]]
    cfg = CONFIG
    if args and not args[0].startswith("-"):
        cfg = args.pop(0)
    settings = load_settings(cfg)
    if args and args[0] in ("--add-admin", "--remove-admin", "--admins"):
        sys.exit(admin_command(settings, args))
    srv = PakServer(settings)
    if "--once" in args:
        srv.once()
        print(open(os.path.join(settings["state"], "catalog.txt"), encoding="cp1252").read())
        return
    srv.run()


if __name__ == "__main__":
    main()
