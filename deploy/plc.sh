#!/bin/bash
# Palms LAN Center Server Setup: installs or upgrades the Palms Server and the Cartridge Server on Linux.
#   bash -c "$(curl -fsSL http://palms-arcade.com/plc.sh)"
# or, from a downloaded server package: bash plc.sh
# Debian, Ubuntu, Fedora, RHEL/Rocky/Alma, Arch, Alpine and openSUSE; 32-bit or 64-bit. Needs Python 3.7 or later.

PLC_VERSION="@PLC_VERSION@"
SERVER_TGZ_URL="@SERVER_TGZ_URL@"
SERVER_SHA256="@SERVER_SHA256@"
SELF_URL="@SELF_URL@"

set -e
shopt -s nullglob
ROOT="${PALMS_ROOT:-}"          # for testing the network part against a fake /etc only
case "$PLC_VERSION" in @*) PLC_VERSION="(source)" ;; esac

# ------------------------------------------------------------------ helpers

say()  { echo "$@"; }
step() { echo; echo "== $*"; }
fail() { echo; echo "$*"; exit 1; }
have() { command -v "$1" >/dev/null 2>&1; }

# ask VAR "Question" "default": reads an answer, the default if Enter is pressed
ask() {
    local __v="$1" __q="$2" __d="$3" __a
    if [ -n "$__d" ]; then
        read -r -p "   $__q [$__d]: " __a
    else
        read -r -p "   $__q: " __a
    fi
    [ -z "$__a" ] && __a="$__d"
    printf -v "$__v" '%s' "$__a"
}

# yes_no "Question" y|n: 0 for yes
yes_no() {
    local a d="$2" hint="[y/N]"
    [ "$d" = y ] && hint="[Y/n]"
    read -r -p "   $1 $hint " a
    [ -z "$a" ] && a="$d"
    case "$a" in [Yy]*) return 0 ;; *) return 1 ;; esac
}

is_ip() {
    local ip="$1" o
    [[ "$ip" =~ ^([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})$ ]] || return 1
    for o in "${BASH_REMATCH[@]:1}"; do [ "$o" -le 255 ] || return 1; done
    return 0
}

# 255.255.255.0, /24 or 24 -> 24 (nothing for anything else)
prefix_of() {
    local m="${1#/}" n=0 o bits seen0=0
    if [[ "$m" =~ ^[0-9]{1,2}$ ]]; then
        [ "$m" -ge 8 ] && [ "$m" -le 30 ] && echo "$m"
        return
    fi
    is_ip "$m" || return
    for o in ${m//./ }; do
        case "$o" in
            255) [ $seen0 = 1 ] && return; n=$((n + 8)) ;;
            0) seen0=1 ;;
            254|252|248|240|224|192|128)
                [ $seen0 = 1 ] && return
                bits=0; while [ $(( (o << bits) & 255 )) -ne 0 ]; do bits=$((bits + 1)); done
                n=$((n + bits)); seen0=1 ;;
            *) return ;;
        esac
    done
    [ "$n" -ge 8 ] && [ "$n" -le 30 ] && echo "$n"
}

mask_of() {
    local p="$1" i m=""
    for i in 1 2 3 4; do
        if [ "$p" -ge 8 ]; then m="$m.255"; p=$((p - 8))
        else m="$m.$(( 256 - (1 << (8 - p)) ))"; p=0; fi
    done
    echo "${m#.}"
}

same_net() {
    local a b p="$3" mask=0 i
    a=$(ip_num "$1"); b=$(ip_num "$2")
    for ((i = 0; i < p; i++)); do mask=$(( mask | (1 << (31 - i)) )); done
    [ $(( a & mask )) -eq $(( b & mask )) ]
}
ip_num() { local IFS=.; set -- $1; echo $(( ($1 << 24) | ($2 << 16) | ($3 << 8) | $4 )); }

# a setting from an ini file: ini_get FILE KEY
ini_get() { [ -f "$1" ] || return 0; sed -n "s/^[[:space:]]*$2[[:space:]]*=[[:space:]]*//p" "$1" | head -n1 | sed 's/[[:space:]]*$//'; }

# VERSION = '0.6.1' in a server program
py_version() { [ -f "$1" ] || return 0; sed -n "s/^VERSION = ['\"]\([^'\"]*\)['\"].*/\1/p" "$1" | head -n1; }

# ------------------------------------------------------------------ this system

INIT=""; PM=""; VIRT=""
detect_system() {
    if [ -d /run/systemd/system ]; then INIT=systemd
    elif have rc-service && have openrc-run || [ -x /sbin/openrc-run ]; then INIT=openrc
    fi
    if have apt-get; then PM=apt
    elif have dnf; then PM=dnf
    elif have yum; then PM=yum
    elif have pacman; then PM=pacman
    elif have apk; then PM=apk
    elif have zypper; then PM=zypper
    fi
    if have systemd-detect-virt; then VIRT="$(systemd-detect-virt --container 2>/dev/null || true)"
    elif grep -qa 'container=lxc' /proc/1/environ 2>/dev/null; then VIRT=lxc
    fi
    [ "$VIRT" = none ] && VIRT=""
    return 0
}

# Python 3.7 or later: sets PY, 1 if there is none
find_python() {
    local c
    for c in python3 python3.14 python3.13 python3.12 python3.11 python3.10 python3.9 python3.8 python3.7; do
        if have $c && $c -c 'import sys; sys.exit(0 if sys.version_info >= (3, 7) else 1)' 2>/dev/null; then
            PY="$(command -v $c)"; return 0
        fi
    done
    return 1
}

pkg_refresh() {
    case "$PM" in
        apt) DEBIAN_FRONTEND=noninteractive apt-get update -q ;;
        dnf|yum) true ;;
        pacman) pacman -Sy --noconfirm >/dev/null ;;
        apk) apk update -q ;;
        zypper) zypper --non-interactive --quiet refresh ;;
    esac
}

# pkg_add NAME...: installs the packages; 1 if one can't be installed
pkg_add() {
    case "$PM" in
        apt)
            # packages don't start their services while installing (the system nginx would fail without IPv6
            # or with port 80 in use); setup starts what it needs itself
            local made=0 r=0
            if [ ! -e /usr/sbin/policy-rc.d ]; then printf '#!/bin/sh\nexit 101\n' > /usr/sbin/policy-rc.d; chmod 755 /usr/sbin/policy-rc.d; made=1; fi
            DEBIAN_FRONTEND=noninteractive apt-get install -y -q "$@" || r=$?
            [ $made = 1 ] && rm -f /usr/sbin/policy-rc.d
            return $r ;;
        dnf) dnf install -y -q "$@" ;;
        yum) yum install -y -q "$@" ;;
        pacman) pacman -S --needed --noconfirm "$@" ;;
        apk) apk add -q "$@" ;;
        zypper) zypper --non-interactive --quiet install "$@" ;;
    esac
}

# pkg_first NAME...: installs the first of these that the system has
pkg_first() {
    local p
    for p in "$@"; do pkg_add "$p" >/dev/null 2>&1 && return 0; done
    return 1
}

# the package names on each system
pkg_names() {
    case "$1:$PM" in
        python:apt|python:dnf|python:yum|python:apk|python:zypper) echo python3 ;;
        python:pacman) echo python ;;
        nginx:*) echo nginx ;;
        samba:*) echo samba ;;
        7z:apt) echo "p7zip-full 7zip" ;;
        7z:dnf|7z:yum) echo "7zip p7zip-plugins p7zip" ;;
        7z:pacman) echo "7zip p7zip" ;;
        7z:apk) echo "7zip p7zip" ;;
        7z:zypper) echo "7zip p7zip-full p7zip" ;;
        pil:apt) echo python3-pil ;;
        pil:dnf|pil:yum) echo python3-pillow ;;
        pil:pacman) echo python-pillow ;;
        pil:apk) echo py3-pillow ;;
        pil:zypper) echo python3-Pillow ;;
        pefile:apt|pefile:dnf|pefile:yum|pefile:zypper) echo python3-pefile ;;
        pefile:pacman) echo python-pefile ;;
        pefile:apk) echo py3-pefile ;;
    esac
}

# ------------------------------------------------------------------ services (systemd or OpenRC)

svc_exists() {
    if [ "$INIT" = systemd ]; then systemctl cat "$1.service" >/dev/null 2>&1
    else [ -x "/etc/init.d/$1" ]; fi
}
svc_start() {   # enable at startup and (re)start
    if [ "$INIT" = systemd ]; then
        systemctl daemon-reload
        systemctl enable "$1" >/dev/null 2>&1 || true
        systemctl restart "$1"
    else
        rc-update add "$1" default >/dev/null 2>&1 || true
        rc-service "$1" restart >/dev/null 2>&1 || rc-service "$1" start
    fi
}
svc_stop_disable() {
    if [ "$INIT" = systemd ]; then
        systemctl disable --now "$1" >/dev/null 2>&1 || true
        systemctl reset-failed "$1" >/dev/null 2>&1 || true
    else
        rc-service "$1" stop >/dev/null 2>&1 || true
        rc-update del "$1" default >/dev/null 2>&1 || true
    fi
}
svc_running() {
    if [ "$INIT" = systemd ]; then systemctl is-active --quiet "$1"
    else rc-service "$1" status >/dev/null 2>&1; fi
}
svc_log_hint() {
    if [ "$INIT" = systemd ]; then echo "journalctl -u $1 -n 30"; else echo "tail -n 30 /var/log/$1.log"; fi
}

# write_service NAME "Description" USER WORKDIR PROGRAM "ARGUMENTS" [extra systemd lines]
write_service() {
    local name="$1" desc="$2" user="$3" dir="$4" cmd="$5" args="$6" extra="$7"
    if [ "$INIT" = systemd ]; then
        {
            echo "[Unit]"
            echo "Description=$desc"
            echo "After=network-online.target"
            echo "Wants=network-online.target"
            echo
            echo "[Service]"
            [ -n "$user" ] && echo "User=$user" && echo "Group=$user"
            [ -n "$extra" ] && echo "$extra"
            echo "WorkingDirectory=$dir"
            echo "ExecStart=$cmd $args"
            echo "Restart=always"
            echo "RestartSec=5"
            echo
            echo "[Install]"
            echo "WantedBy=multi-user.target"
        } > "/etc/systemd/system/$name.service"
    else
        touch "/var/log/$name.log"
        [ -n "$user" ] && chown "$user" "/var/log/$name.log"
        {
            echo "#!/sbin/openrc-run"
            echo "name=\"$desc\""
            echo "command=\"$cmd\""
            echo "command_args=\"$args\""
            [ -n "$user" ] && echo "command_user=\"$user:$user\""
            echo "directory=\"$dir\""
            echo "command_background=true"
            echo "pidfile=\"/run/$name.pid\""
            echo "output_log=\"/var/log/$name.log\""
            echo "error_log=\"/var/log/$name.log\""
            echo "depend() { need net; }"
        } > "/etc/init.d/$name"
        chmod 755 "/etc/init.d/$name"
    fi
}

# ------------------------------------------------------------------ network files

# The network setup in use: networkmanager, netplan, networkd, ifupdown, wicked, or nothing if unknown
net_kind() {
    if have nmcli && { [ "$INIT" != systemd ] || systemctl is-active --quiet NetworkManager 2>/dev/null; } && nmcli -t general status >/dev/null 2>&1; then echo networkmanager
    elif have netplan && [ -n "$(ls "$ROOT"/etc/netplan/*.yaml 2>/dev/null)" ]; then echo netplan
    elif [ "$INIT" = systemd ] && systemctl is-active --quiet systemd-networkd 2>/dev/null && [ ! -f "$ROOT/etc/network/interfaces" ]; then echo networkd
    elif [ "$INIT" = systemd ] && systemctl is-active --quiet wicked 2>/dev/null; then echo wicked
    elif [ -f "$ROOT/etc/network/interfaces" ]; then echo ifupdown
    fi
}

# write_network KIND IFACE IP PREFIX GATEWAY DNS (gateway and DNS may be blank)
write_network() {
    local kind="$1" ifc="$2" ip="$3" p="$4" gw="$5" dns="$6" stamp f
    stamp="$(date +%Y%m%d-%H%M%S)"
    case "$kind" in
    ifupdown)
        mkdir -p "$ROOT/etc/network/interfaces.d"
        cp "$ROOT/etc/network/interfaces" "$ROOT/etc/network/interfaces.before-palms-$stamp"
        # the card's old settings are commented out (other cards stay as they are), the new ones added at the end
        awk -v ifc="$ifc" '
            /^[[:space:]]*(iface|auto|allow-[a-z]+|mapping|source|source-directory)[[:space:]]/ { inside = 0 }
            /^# Written by Palms/ { next }
            ($1 == "auto" || $1 ~ /^allow-/) && NF == 2 && $2 == ifc { print "# (Palms Setup) " $0; next }
            $1 == "iface" && $2 == ifc { inside = 1 }
            inside && NF > 0 { print "# (Palms Setup) " $0; next }
            { print }
        ' "$ROOT/etc/network/interfaces.before-palms-$stamp" > "$ROOT/etc/network/interfaces"
        # another file for the same card would fight the new settings
        for f in "$ROOT"/etc/network/interfaces.d/*; do
            if grep -Eq "^[[:space:]]*(auto|allow-hotplug|iface)[[:space:]]+$ifc([[:space:]]|$)" "$f"; then
                mv "$f" "$ROOT/etc/network/$(basename "$f").before-palms-$stamp"
                say "   Moved $(basename "$f") out of interfaces.d (it also set up $ifc)."
            fi
        done
        {
            echo
            echo "# Written by Palms Setup on $(date '+%Y-%m-%d %H:%M'). Previous: interfaces.before-palms-$stamp"
            echo "auto $ifc"
            echo "iface $ifc inet static"
            echo "    address $ip/$p"
            if [ -n "$gw" ]; then echo "    gateway $gw"; fi
        } >> "$ROOT/etc/network/interfaces"
        if [ -n "$dns" ]; then
            [ -f "$ROOT/etc/resolv.conf" ] && cp -L "$ROOT/etc/resolv.conf" "$ROOT/etc/resolv.conf.before-palms-$stamp"
            rm -f "$ROOT/etc/resolv.conf"
            { echo "# Written by Palms Setup"; for d in $dns; do echo "nameserver $d"; done; } > "$ROOT/etc/resolv.conf"
        fi
        ;;
    netplan)
        for f in "$ROOT"/etc/netplan/*.yaml; do
            if grep -q "$ifc" "$f" && [ "$(basename "$f")" != "60-palms.yaml" ]; then
                mv "$f" "$f.before-palms-$stamp"
                say "   Moved $(basename "$f") aside (it also set up $ifc)."
            fi
        done
        {
            echo "# Written by Palms Setup"
            echo "network:"
            echo "  version: 2"
            echo "  ethernets:"
            echo "    $ifc:"
            echo "      dhcp4: false"
            echo "      addresses: [$ip/$p]"
            if [ -n "$gw" ]; then echo "      routes: [{to: default, via: $gw}]"; fi
            if [ -n "$dns" ]; then echo "      nameservers: {addresses: [$(echo $dns | sed 's/ /, /g')]}"; fi
        } > "$ROOT/etc/netplan/60-palms.yaml"
        chmod 600 "$ROOT/etc/netplan/60-palms.yaml"
        if [ -z "$ROOT" ] && ! netplan generate 2>/dev/null; then
            # older netplan (Ubuntu 20.04) has no "to: default"
            sed -i "s/      routes: \[{to: default, via: \(.*\)}\]/      gateway4: \1/" "$ROOT/etc/netplan/60-palms.yaml"
            netplan generate
        fi
        ;;
    networkd)
        mkdir -p "$ROOT/etc/systemd/network"
        for f in "$ROOT"/etc/systemd/network/*.network; do
            if grep -Eq "^Name=$ifc\$" "$f" && [ "$(basename "$f")" != "10-palms-$ifc.network" ]; then mv "$f" "$f.before-palms-$stamp"; fi
        done
        {
            echo "# Written by Palms Setup"
            echo "[Match]"
            echo "Name=$ifc"
            echo
            echo "[Network]"
            echo "Address=$ip/$p"
            if [ -n "$gw" ]; then echo "Gateway=$gw"; fi
            for d in $dns; do echo "DNS=$d"; done
        } > "$ROOT/etc/systemd/network/10-palms-$ifc.network"
        ;;
    wicked)
        f="$ROOT/etc/sysconfig/network/ifcfg-$ifc"
        [ -f "$f" ] && cp "$f" "$f.before-palms-$stamp"
        printf "# Written by Palms Setup\nBOOTPROTO='static'\nSTARTMODE='auto'\nIPADDR='%s/%s'\n" "$ip" "$p" > "$f"
        if [ -n "$gw" ]; then
            [ -f "$ROOT/etc/sysconfig/network/routes" ] && cp "$ROOT/etc/sysconfig/network/routes" "$ROOT/etc/sysconfig/network/routes.before-palms-$stamp"
            echo "default $gw - $ifc" > "$ROOT/etc/sysconfig/network/routes"
        fi
        if [ -n "$dns" ] && [ -f "$ROOT/etc/sysconfig/network/config" ]; then
            sed -i "s/^NETCONFIG_DNS_STATIC_SERVERS=.*/NETCONFIG_DNS_STATIC_SERVERS=\"$dns\"/" "$ROOT/etc/sysconfig/network/config"
        fi
        ;;
    networkmanager)
        local con
        con="$(nmcli -g GENERAL.CONNECTION device show "$ifc" 2>/dev/null | head -n1)"
        if [ -z "$con" ] || [ "$con" = "--" ]; then
            con="palms-$ifc"
            nmcli con add type ethernet ifname "$ifc" con-name "$con" >/dev/null
        fi
        nmcli con mod "$con" ipv4.method manual ipv4.addresses "$ip/$p" ipv4.gateway "${gw:-}" \
            ipv4.dns "${dns:-}" ipv6.method ignore connection.autoconnect yes
        ;;
    esac
    # cloud-init (cloud images, some VM templates) would put its own settings back at startup
    if [ -d "$ROOT/etc/cloud/cloud.cfg.d" ]; then
        echo "network: {config: disabled}" > "$ROOT/etc/cloud/cloud.cfg.d/99-palms-disable-network-config.cfg"
        say "   cloud-init network setup turned off (it would undo this address at startup)."
    fi
}

write_hostname() {
    local name="$1"
    echo "$name" > "$ROOT/etc/hostname"
    if [ -f "$ROOT/etc/hosts" ] && grep -q '^127\.0\.1\.1' "$ROOT/etc/hosts"; then
        sed "s/^127\.0\.1\.1.*/127.0.1.1\t$name/" "$ROOT/etc/hosts" > "$ROOT/etc/hosts.palms-new"
        cat "$ROOT/etc/hosts.palms-new" > "$ROOT/etc/hosts"; rm -f "$ROOT/etc/hosts.palms-new"
    else
        printf '127.0.1.1\t%s\n' "$name" >> "$ROOT/etc/hosts"
    fi
    if [ -z "$ROOT" ]; then
        if have hostnamectl && [ "$INIT" = systemd ]; then hostnamectl set-hostname "$name" 2>/dev/null || true
        else hostname "$name" 2>/dev/null || true; fi
    fi
}

# only the network part, against a fake root: PALMS_ROOT=/tmp/x plc.sh --test-network ifupdown eth0 10.0.0.5 24 10.0.0.1 "1.1.1.1"
if [ "$1" = "--test-network" ]; then
    write_network "$2" "$3" "$4" "$5" "$6" "$7"
    write_hostname "${8:-palms-test}"
    exit 0
fi

# ------------------------------------------------------------------ start
# All of setup is in main, so that bash has read the whole script before it starts
# (it may be arriving through a pipe).
main() {

echo
echo "Palms LAN Center Server Setup $PLC_VERSION"
echo "=================================================="

# the questions come from the keyboard, even when the script itself arrived through a pipe (curl ... | bash)
if [ ! -t 0 ]; then
    if { true < /dev/tty; } 2>/dev/null; then exec < /dev/tty; else fail "Run setup in a terminal (it asks questions)."; fi
fi
# run as root (with sudo if there is one)
if [ "$(id -u)" != 0 ]; then
    if have sudo; then
        me="$(mktemp /tmp/plc-setup.XXXXXX)"
        if [ -f "$0" ] && grep -q "Palms LAN Center Server Setup" "$0" 2>/dev/null; then cat "$0" > "$me"
        elif [ -n "${BASH_EXECUTION_STRING:-}" ]; then printf '%s\n' "$BASH_EXECUTION_STRING" > "$me"
        else
            case "$SELF_URL" in @*) fail "Setup needs root. Log in as root (or type: su -) and run it again." ;; esac
            if have curl; then curl -fsSL -o "$me" "$SELF_URL"; else wget -q -O "$me" "$SELF_URL"; fi
        fi
        say "   Setup needs root. Asking sudo..."
        exec sudo bash "$me" "$@"
    fi
    fail "Setup needs root. Log in as root (or type: su -) and run it again."
fi

detect_system
[ -n "$PM" ] || fail "This Linux has no package manager setup knows (apt, dnf, yum, pacman, apk or zypper)."
[ -n "$INIT" ] || fail "This Linux starts programs with something other than systemd or OpenRC. Setup can't install the services."
. /etc/os-release 2>/dev/null || true
say "   This machine:   ${PRETTY_NAME:-Linux}, $(uname -m), $INIT"

# ------------------------------------------------------------------ the server programs

WORK="$(mktemp -d /tmp/plc-files.XXXXXX)"
trap 'rm -rf "$WORK"' EXIT
HERE=""
[ -f "$0" ] && HERE="$(cd "$(dirname "$0")" && pwd)"
if [ -n "$HERE" ] && [ -f "$HERE/palms_server.py" ] && [ -f "$HERE/palms_paks.py" ]; then
    cp "$HERE/palms_server.py" "$HERE/palms_paks.py" "$WORK/"
elif [ -n "$HERE" ] && [ -f "$HERE/../server/palms_server.py" ] && [ -f "$HERE/../cartridge-server/palms_paks.py" ]; then
    cp "$HERE/../server/palms_server.py" "$HERE/../cartridge-server/palms_paks.py" "$WORK/"
else
    case "$SERVER_TGZ_URL" in @*) fail "palms_server.py and palms_paks.py not found next to this script." ;; esac
    say "   Downloading:    $SERVER_TGZ_URL"
    if have curl; then curl -fsSL -o "$WORK/server.tar.gz" "$SERVER_TGZ_URL" || fail "Download failed. Check the machine's internet connection."
    elif have wget; then wget -q -O "$WORK/server.tar.gz" "$SERVER_TGZ_URL" || fail "Download failed. Check the machine's internet connection."
    else fail "Install curl or wget, then run setup again."; fi
    if have sha256sum; then
        [ "$(sha256sum "$WORK/server.tar.gz" | awk '{print $1}')" = "$SERVER_SHA256" ] || fail "The download is damaged (checksum mismatch). Run setup again."
    fi
    tar -xzf "$WORK/server.tar.gz" -C "$WORK"
    [ -f "$WORK/palms_server.py" ] && [ -f "$WORK/palms_paks.py" ] || fail "The download is incomplete. Run setup again."
fi
NEW_SRV="$(py_version "$WORK/palms_server.py")"
NEW_PAKS="$(py_version "$WORK/palms_paks.py")"

# ------------------------------------------------------------------ already installed: upgrade?

OLD_SRV="$(py_version /opt/palms/palms_server.py)"
OLD_PAKS="$(py_version /opt/palms-paks/palms_paks.py)"
MODE=new
if [ -n "$OLD_SRV$OLD_PAKS" ]; then
    step "Already installed"
    [ -n "$OLD_SRV" ] && say "   Palms Server $OLD_SRV       (this setup: $NEW_SRV)"
    [ -n "$OLD_PAKS" ] && say "   Cartridge Server $OLD_PAKS   (this setup: $NEW_PAKS)"
    say
    say "   1  Upgrade (keeps every setting, chat, schedule and cartridge)"
    say "   2  Set up again (asks every question)"
    say "   3  Quit"
    A=""
    while ! [[ "$A" =~ ^[123]$ ]]; do ask A "Choose" "1"; done
    [ "$A" = 3 ] && { echo "Nothing was changed."; exit 0; }
    [ "$A" = 1 ] && MODE=upgrade
fi

if [ $MODE = upgrade ]; then
    step "Upgrading"
    stamp="$(date +%Y%m%d-%H%M%S)"
    if [ -n "$OLD_SRV" ]; then
        mkdir -p /opt/palms/old
        cp /opt/palms/palms_server.py "/opt/palms/old/palms_server-$OLD_SRV-$stamp.py"
        install -m 755 "$WORK/palms_server.py" /opt/palms/palms_server.py
        chown -R palms:palms /opt/palms/old /opt/palms/palms_server.py 2>/dev/null || true
        svc_start palms
        say "   Palms Server $OLD_SRV -> $NEW_SRV (the old one is in /opt/palms/old)"
    fi
    if [ -n "$OLD_PAKS" ]; then
        mkdir -p /opt/palms-paks/old
        cp /opt/palms-paks/palms_paks.py "/opt/palms-paks/old/palms_paks-$OLD_PAKS-$stamp.py"
        install -m 755 "$WORK/palms_paks.py" /opt/palms-paks/palms_paks.py
        svc_start palms-paks
        say "   Cartridge Server $OLD_PAKS -> $NEW_PAKS (the old one is in /opt/palms-paks/old)"
    fi
    sleep 3
    OK=1
    for s in $( [ -n "$OLD_PAKS" ] && echo palms-paks ) $( [ -n "$OLD_SRV" ] && echo palms ); do
        if svc_running "$s"; then say "   $s: running"; else say "   $s: NOT RUNNING. For details: $(svc_log_hint "$s")"; OK=0; fi
    done
    echo
    if [ $OK = 1 ]; then echo "Upgrade complete."; else echo "Upgrade finished with problems (see above)."; fi
    exit 0
fi

# ------------------------------------------------------------------ questions

step "1. Role"
say "   1  Palms Server       (chat, schedule, tournaments, projector page)"
say "   2  Cartridge Server   (Game Library cartridges, upload share, approval)"
say "   3  Both on this machine"
ROLE=""
DEF_ROLE=1; [ -n "$OLD_PAKS" ] && DEF_ROLE=2; [ -n "$OLD_PAKS" ] && [ -n "$OLD_SRV" ] && DEF_ROLE=3
while ! [[ "$ROLE" =~ ^[123]$ ]]; do ask ROLE "Install which" "$DEF_ROLE"; done
WANT_PALMS=0; WANT_PAKS=0
[ "$ROLE" = 1 ] || [ "$ROLE" = 3 ] && WANT_PALMS=1
[ "$ROLE" = 2 ] || [ "$ROLE" = 3 ] && WANT_PAKS=1
case "$ROLE" in 1) DEF_NAME=palms-server ;; 2) DEF_NAME=palms-carts ;; *) DEF_NAME=palms-lan ;; esac

step "2. Network"
have ip || pkg_add iproute2 >/dev/null 2>&1 || true
have ip || fail "The ip command is missing. Install iproute2, then run setup again."
# the card with the default route, otherwise the first one that isn't lo
IFACE="$(ip -o route show default 2>/dev/null | awk '{for (i = 1; i < NF; i++) if ($i == "dev") {print $(i + 1); exit}}')"
[ -n "$IFACE" ] || IFACE="$(ip -o link show | awk -F': ' '$2 != "lo" {sub(/@.*/, "", $2); print $2; exit}')"
CUR="$(ip -o -4 addr show dev "$IFACE" 2>/dev/null | awk '{print $4; exit}')"
CUR_IP="${CUR%/*}"
SET_NET=0
NET="$(net_kind)"
if [ "$VIRT" = lxc ]; then
    say "   This is a Proxmox (LXC) container. Its address and name are set in Proxmox, not here:"
    say "   anything set inside the container is replaced when it starts."
    say
    say "   On the Proxmox host, open the container > Network > net0 > Edit and set the IPv4 address"
    say "   (for example 10.0.200.71/24) and gateway, and set the name under DNS > Hostname."
    say "   Or run these on the Proxmox host (<CTID> is the container's number, for example 105):"
    say "     pct set <CTID> -net0 name=eth0,bridge=vmbr0,ip=10.0.200.71/24,gw=10.0.200.1"
    say "     pct set <CTID> -hostname $DEF_NAME"
    say "     pct reboot <CTID>"
    say
    say "   This container's address now: ${CUR:-none} on ${IFACE:-no network card}. Name: $(hostname)."
    if ! yes_no "Is that the address and name it should keep?" n; then
        echo
        echo "Set the address and name in Proxmox, restart the container, then run setup again."
        exit 0
    fi
    [ -n "$CUR_IP" ] || fail "The container has no IPv4 address. Set one in Proxmox first."
    IP="$CUR_IP"; HOSTNAME_NEW="$(hostname)"
else
    say "   Network cards:"
    for c in $(ip -o link show | awk -F': ' '$2 != "lo" {sub(/@.*/, "", $2); print $2}'); do
        a="$(ip -o -4 addr show dev "$c" 2>/dev/null | awk '{print $4; exit}')"
        printf "     %-12s %s\n" "$c" "${a:-(no address)}"
    done
    while :; do
        ask IFACE "Network card for the LAN" "$IFACE"
        ip link show "$IFACE" >/dev/null 2>&1 && break
        say "   There is no network card called $IFACE."
    done
    CUR="$(ip -o -4 addr show dev "$IFACE" 2>/dev/null | awk '{print $4; exit}')"
    CUR_IP="${CUR%/*}"
    say "   Address now: ${CUR:-none}"
    if [ -z "$NET" ]; then
        say "   Setup can't change the address on this Linux. Set a fixed address yourself if it needs one."
    elif yes_no "Set a fixed (static) address on $IFACE?" y; then
        SET_NET=1
        while :; do ask IP "IP address" "${CUR_IP:-10.0.200.71}"; is_ip "$IP" && break; say "   Enter an address like 10.0.200.71."; done
        CUR_P="${CUR#*/}"; [ "$CUR_P" = "$CUR" ] && CUR_P=24
        while :; do
            ask MASK "Subnet mask" "$(mask_of "$CUR_P")"
            PREFIX="$(prefix_of "$MASK")"
            [ -n "$PREFIX" ] && break
            say "   Enter a mask like 255.255.255.0 (or /24)."
        done
        CUR_GW="$(ip -o route show default dev "$IFACE" 2>/dev/null | awk '{print $3; exit}')"
        while :; do
            ask GW "Default gateway (Enter - for none, on a LAN without a router)" "${CUR_GW:--}"
            [ "$GW" = - ] && GW="" && break
            if is_ip "$GW" && same_net "$IP" "$GW" "$PREFIX"; then break; fi
            say "   Enter an address on the same network as $IP/$PREFIX, or - for none."
        done
        CUR_DNS="$(awk '/^nameserver/ && $2 !~ /^127\./ {printf "%s ", $2}' /etc/resolv.conf 2>/dev/null | sed 's/ $//')"
        [ -z "$CUR_DNS" ] && [ -n "$CUR_GW" ] && CUR_DNS="$CUR_GW"
        while :; do
            ask DNS "DNS servers, separated by spaces (- for none)" "${CUR_DNS:--}"
            [ "$DNS" = - ] && DNS="" && break
            ok=1; for d in $DNS; do is_ip "$d" || ok=0; done
            [ $ok = 1 ] && break
            say "   Enter addresses like 10.0.200.1 8.8.8.8, or - for none."
        done
    fi
    if [ $SET_NET = 0 ]; then
        [ -n "$CUR_IP" ] || fail "$IFACE has no IPv4 address. Give it one, then run setup again."
        IP="$CUR_IP"
        say "   Keeping $CUR. Make sure this address does not change (a DHCP reservation, for example)."
    fi
    while :; do
        ask HOSTNAME_NEW "Computer name" "$( [ -n "$OLD_SRV$OLD_PAKS" ] && hostname || echo "$DEF_NAME")"
        [[ "$HOSTNAME_NEW" =~ ^[A-Za-z0-9]([A-Za-z0-9-]{0,13}[A-Za-z0-9])?$ ]] && break
        say "   1 to 15 letters, numbers and - (Windows 98 can't see longer names)."
    done
fi

step "3. Settings"
if [ $WANT_PALMS = 1 ]; then
    OLD_INI=/opt/palms/palms.ini
    d="$(ini_get $OLD_INI name || true)"
    ask SRV_NAME "Palms Server name (shown in Palms LAN Center)" "${d:-PALMS-SRV}"
    SRV_NAME="$(echo "$SRV_NAME" | tr -cd 'A-Za-z0-9 _-' | cut -c1-24)"
    [ -n "$SRV_NAME" ] || SRV_NAME="PALMS-SRV"
    OLD_PW="$(ini_get $OLD_INI admin_password || true)"
    while :; do
        if [ -n "$OLD_PW" ]; then ask HOST_PW "Host password for Palms LAN Center" "$OLD_PW"
        else ask HOST_PW "Host password for Palms LAN Center (Enter for a random one)" ""; fi
        [[ "$HOST_PW" =~ ^[^|[:space:]]{0,32}$ ]] && break
        say "   Up to 32 characters, without spaces or |."
    done
    if [ $WANT_PAKS = 1 ]; then
        PAKS_HOST="$IP"
    else
        OLD_PAKS_HOST="$(ini_get $OLD_INI paks_url | sed 's#^http://##; s#/.*##' || true)"
        while :; do
            ask PAKS_HOST "Cartridge Server address" "${OLD_PAKS_HOST:-10.0.200.73}"
            is_ip "$PAKS_HOST" && break
            [[ "$PAKS_HOST" =~ ^[A-Za-z0-9.-]+$ ]] && break
            say "   Enter an address like 10.0.200.73."
        done
    fi
    FILE_PAGE=off
    d=n; [ "$(ini_get $OLD_INI file_page || true)" = on ] && d=y
    if yes_no "Turn on the file page (http://$IP:8081/, everyone on the LAN can download and upload files)?" $d; then
        FILE_PAGE=on
    fi
fi
if [ $WANT_PAKS = 1 ]; then
    while :; do
        d="$(ini_get /etc/palms-paks.ini folder || true)"
        ask PAKS "Cartridge folder (needs space for every cartridge)" "${d:-/srv/paks}"
        [[ "$PAKS" =~ ^/[A-Za-z0-9/._-]+$ ]] && [ "$PAKS" != / ] && break
        say "   Enter a full path like /srv/paks or /mnt/games/paks."
    done
    PAKS="${PAKS%/}"
fi

step "4. Summary"
case "$ROLE" in 1) say "   Role:              Palms Server" ;; 2) say "   Role:              Cartridge Server" ;; *) say "   Role:              Palms Server and Cartridge Server" ;; esac
if [ "$VIRT" = lxc ]; then
    say "   Network:           $IP (set in Proxmox), name $HOSTNAME_NEW"
elif [ $SET_NET = 1 ]; then
    say "   Network card:      $IFACE ($NET)"
    say "   Address:           $IP, mask $(mask_of "$PREFIX"), gateway ${GW:-none}, DNS ${DNS:-none}"
    say "   Computer name:     $HOSTNAME_NEW"
else
    say "   Address:           $IP on $IFACE (kept as it is), name $HOSTNAME_NEW"
fi
if [ $WANT_PALMS = 1 ]; then
    say "   Palms Server $NEW_SRV: \"$SRV_NAME\", /opt/palms, host password ${HOST_PW:-(random)}"
    say "                      Game Library from the Cartridge Server at http://$PAKS_HOST/"
    say "                      Projector page http://$IP:8080/, file page $FILE_PAGE"
fi
[ $WANT_PAKS = 1 ] && say "   Cartridge Server $NEW_PAKS: cartridges in $PAKS, address http://$IP/"
say "   Packages from:     $PM (the machine needs internet access for this)"
say
yes_no "Install now?" y || { echo "Nothing was changed."; exit 0; }

# ------------------------------------------------------------------ the network (applied at the end)

if [ "$VIRT" != lxc ]; then
    step "5. Computer name and network settings"
    write_hostname "$HOSTNAME_NEW"
    say "   Computer name: $HOSTNAME_NEW"
    if [ $SET_NET = 1 ]; then
        write_network "$NET" "$IFACE" "$IP" "$PREFIX" "$GW" "$DNS"
        say "   $IP/$PREFIX on $IFACE is used from the next restart."
    fi
fi

# ------------------------------------------------------------------ packages

step "6. Installing packages"
if ! pkg_refresh; then
    say "   Unable to download the package lists. The machine needs internet access for this step"
    say "   (a gateway and DNS server). Connect it, then run setup again."
    exit 1
fi
# Rocky, Alma and other RHEL copies have 7-Zip and pefile in EPEL
if [ "$PM" = dnf ] || [ "$PM" = yum ]; then
    case " ${ID:-} ${ID_LIKE:-} " in *" fedora "*) [ "${ID:-}" = fedora ] || pkg_add epel-release >/dev/null 2>&1 || true ;; esac
fi
pkg_add $(pkg_names python) || fail "Unable to install Python. See the messages above."
# older Ubuntu has a newer Python beside its own
find_python || { pkg_first python3.8 python3.7 python38 >/dev/null 2>&1 || true; find_python; } || \
    fail "This Linux's Python is older than 3.7, and there is no newer one. The Palms servers need Python 3.7 or later.
Use a newer Linux (for example Debian 12, which runs on 32-bit PCs from the Pentium 4 on)."
say "   Python $("$PY" -c 'import sys; print("%d.%d" % sys.version_info[:2])')"
if [ $WANT_PAKS = 1 ]; then
    pkg_add nginx || fail "Unable to install nginx. See the messages above."
    pkg_add samba || fail "Unable to install Samba. See the messages above."
    # shellcheck disable=SC2046
    pkg_first $(pkg_names 7z) || say "   7-Zip could not be installed. Cartridges can't be checked without it."
    if ! have 7z && have 7zz; then ln -sf "$(command -v 7zz)" /usr/local/bin/7z; fi
    pkg_first $(pkg_names pil) || say "   Pillow (pictures) is not available here. Icons and screenshots are left out."
    pkg_first $(pkg_names pefile) || say "   pefile is not available here. Icons can't be taken from programs."
    [ "$PM" = apt ] && pkg_add samba-vfs-modules >/dev/null 2>&1 || true
    [ "$PM" = apk ] && pkg_add samba-server >/dev/null 2>&1 || true
    if have getenforce && [ "$(getenforce 2>/dev/null)" = Enforcing ]; then
        pkg_first policycoreutils-python-utils policycoreutils-python >/dev/null 2>&1 || true
    fi
fi

# the palms user runs the services
NOLOGIN=/usr/sbin/nologin; [ -x $NOLOGIN ] || NOLOGIN=/sbin/nologin
if ! id palms >/dev/null 2>&1; then
    if have useradd; then useradd --system --user-group --home-dir /opt/palms --shell $NOLOGIN palms
    else addgroup -S palms; adduser -S -D -H -h /opt/palms -s $NOLOGIN -G palms palms; fi
fi

# ------------------------------------------------------------------ Cartridge Server

if [ $WANT_PAKS = 1 ]; then
    step "7. Cartridge Server"
    STATE=/var/lib/palms-paks
    APP=/opt/palms-paks
    mkdir -p "$PAKS" "$STATE/icons" "$STATE/manifests" "$STATE/nginx" "$APP"
    for d in inbox waiting library trash; do mkdir -p "$PAKS/$d"; done
    chown palms:palms "$PAKS" "$PAKS/inbox" "$PAKS/waiting" "$PAKS/library" "$PAKS/trash" "$STATE" "$STATE/icons" "$STATE/manifests" "$STATE/nginx"
    chmod 755 "$PAKS" "$PAKS/waiting" "$PAKS/library" "$PAKS/trash"
    chmod 2775 "$PAKS/inbox"
    install -m 755 "$WORK/palms_paks.py" "$APP/palms_paks.py"
    cat > /etc/palms-paks.ini <<EOF
; Palms Cartridge Server settings. After making changes, restart the palms-paks service.
[paks]
; cartridge folder: contains inbox (the incoming share), waiting, library and trash
folder = $PAKS
; state folder: catalog, index page, icons, manifests and audit log
state = $STATE
; web address of this server, as used by Palms LAN Center
url = http://$IP/
EOF
    say "   Settings: /etc/palms-paks.ini"

    mkdir -p /etc/palms-paks
    chown root:palms /etc/palms-paks
    chmod 750 /etc/palms-paks
    if [ -s /etc/palms-paks/admins ] && grep -q '^[^#].*:' /etc/palms-paks/admins; then
        say "   Host accounts: $("$PY" "$APP/palms_paks.py" --admins | tr '\n' ' ')"
    else
        say "   Cartridges are approved by the host at http://$IP/admin. Create the first host account."
        NAME=""
        while ! echo "$NAME" | grep -Eq '^[A-Za-z0-9._-]{1,32}$'; do
            read -r -p "   Account name (for example, host): " NAME
        done
        "$PY" "$APP/palms_paks.py" --add-admin "$NAME"
    fi
    [ -f /etc/palms-paks/admins ] && chown root:palms /etc/palms-paks/admins && chmod 640 /etc/palms-paks/admins

    # The web server: nginx with its own settings and service (palms-web), the same on every Linux.
    # The system's own nginx service is turned off: it would want port 80 too.
    MIME=/etc/nginx/mime.types
    cat > /etc/palms-paks/nginx.conf <<EOF
# Palms Cartridge Server web server: index page, catalog, icons, approved cartridges and the admin page.
# Started by the palms-web service.
user palms palms;
worker_processes 1;
pid /run/palms-web-nginx.pid;
error_log /var/log/palms-web-error.log warn;
events { worker_connections 256; }
http {
    include $MIME;
    default_type application/octet-stream;
    access_log off;
    client_body_temp_path $STATE/nginx/body;
    proxy_temp_path $STATE/nginx/proxy;
    fastcgi_temp_path $STATE/nginx/fastcgi;
    uwsgi_temp_path $STATE/nginx/uwsgi;
    scgi_temp_path $STATE/nginx/scgi;
    # send large files directly from disk
    sendfile on;
    tcp_nopush on;
    sendfile_max_chunk 2m;
    server {
        listen 80 default_server;
        server_name _;
        root $STATE;
        index index.html;
        # only these files from the state folder (not the manifests or the audit log)
        location = / { }
        location = /index.html  { add_header Cache-Control "no-cache"; }
        location = /catalog.txt { add_header Cache-Control "no-cache"; }
        location /icons/ { }
        location / { return 404; }
        # approved cartridges only
        location /paks/ {
            alias $PAKS/library/;
            autoindex on;
            autoindex_exact_size off;
            autoindex_localtime on;
        }
        # admin page (served by the palms-paks service)
        location /admin {
            proxy_pass http://127.0.0.1:8081;
            proxy_set_header X-Real-IP \$remote_addr;
            proxy_set_header Host \$host;
            client_max_body_size 64k;
        }
    }
}
EOF
    NGINX="$(command -v nginx || echo /usr/sbin/nginx)"
    if ! "$NGINX" -t -c /etc/palms-paks/nginx.conf >/dev/null 2>&1; then
        "$NGINX" -t -c /etc/palms-paks/nginx.conf 2>&1 | sed 's/^/   /'
        fail "The web server settings are wrong (see above)."
    fi
    svc_exists nginx && svc_stop_disable nginx
    rm -f /etc/nginx/sites-enabled/palms-paks      # from Setup 1.x
    write_service palms-web "Palms Cartridge Server web server" "" / "$NGINX" "-c /etc/palms-paks/nginx.conf -g 'daemon off;'"
    say "   Web server: palms-web (nginx) on port 80"

    [ -f /etc/samba/smb.conf ] && [ ! -f /etc/samba/smb.conf.before-palms ] && cp /etc/samba/smb.conf /etc/samba/smb.conf.before-palms
    mkdir -p /etc/samba /var/log/samba
    AUDIT=""
    if have journalctl && ls /usr/lib*/samba/vfs/full_audit.so /usr/lib/*/samba/vfs/full_audit.so >/dev/null 2>&1; then
        AUDIT="   # writes, renames and deletions (with the client address) are logged to the system log
   # and read by the palms-paks service for the audit log
   vfs objects = full_audit
   full_audit:prefix = %I
   full_audit:success = mkdirat renameat unlinkat openat
   full_audit:failure = none
   full_audit:facility = LOCAL5
   full_audit:priority = NOTICE"
    else
        say "   Share activity is not recorded in the audit log on this Linux."
    fi
    cat > /etc/samba/smb.conf <<EOF
# Palms Cartridge Server: "incoming" is the upload share, open to everyone on the network;
# "library" contains the approved cartridges (read-only). The previous file is smb.conf.before-palms.
[global]
   workgroup = WORKGROUP
   server string = Palms Cartridge Server
   netbios name = $HOSTNAME_NEW
   server role = standalone server
   # Windows 98 and XP support SMB1 only, which Samba disables by default
   server min protocol = NT1
   ntlm auth = ntlmv1-permitted
   # no password required: unknown users connect as guest
   map to guest = Bad User
   # printing disabled
   load printers = no
   printing = bsd
   printcap name = /dev/null
   disable spoolss = yes
   # faster transfers of large files
   use sendfile = yes
   log file = /var/log/samba/log.%m
   max log size = 1000
   logging = file

[incoming]
   comment = Palms Cartridge Server - upload (host approval required)
   path = $PAKS/inbox
   browseable = yes
   read only = no
   guest ok = yes
   # all files are owned by the palms user
   force user = palms
   force group = palms
   create mask = 0664
   directory mask = 2775
   # files created by Windows Explorer
   veto files = /Thumbs.db/desktop.ini/.DS_Store/
   delete veto files = yes
$AUDIT

[library]
   comment = Palms Cartridge Server - approved cartridges (read-only)
   path = $PAKS/library
   browseable = yes
   read only = yes
   guest ok = yes
EOF
    if have testparm && ! testparm -s >/dev/null 2>&1; then echo "Samba configuration error:"; testparm -s; exit 1; fi
    SMB_SERVICES=""
    for s in smbd nmbd smb nmb samba; do svc_exists $s && SMB_SERVICES="$SMB_SERVICES $s"; done
    case "$SMB_SERVICES" in *smbd*|*smb*) SMB_SERVICES="${SMB_SERVICES/ samba/}" ;; esac
    say "   Shares: \\\\$IP\\incoming and \\\\$IP\\library"

    JOURNAL=""
    getent group systemd-journal >/dev/null 2>&1 && JOURNAL="# required to read Samba audit entries from the system log
SupplementaryGroups=systemd-journal"
    write_service palms-paks "Palms Cartridge Server" palms "$APP" "$PY" "$APP/palms_paks.py" "$JOURNAL"
fi

# ------------------------------------------------------------------ Palms Server

if [ $WANT_PALMS = 1 ]; then
    step "8. Palms Server"
    mkdir -p /opt/palms/data /opt/palms/files
    install -m 755 "$WORK/palms_server.py" /opt/palms/palms_server.py
    [ -n "$HOST_PW" ] || HOST_PW="$(tr -dc 'a-z0-9' < /dev/urandom | head -c 8)"
    [ -f /opt/palms/palms.ini ] && cp /opt/palms/palms.ini /opt/palms/palms.ini.before-setup
    cat > /opt/palms/palms.ini <<EOF
# Palms Server settings. Restart the palms service after making changes.
[server]
name = $SRV_NAME
tcp_port = 4098
beacon_port = 4099
beacon_seconds = 3
admin_password = $HOST_PW
; the projector page for the big screen (0 = off)
projector_port = 8080
; the file page: on or off, and its port
file_page = $FILE_PAGE
http_port = 8081
paks_url = http://$PAKS_HOST/
EOF
    chown -R palms:palms /opt/palms
    chmod 640 /opt/palms/palms.ini
    write_service palms "Palms Server" palms /opt/palms "$PY" "/opt/palms/palms_server.py"
    say "   Program and settings: /opt/palms"
fi

# ------------------------------------------------------------------ firewall and SELinux

PORTS_TCP=""; PORTS_UDP=""
[ $WANT_PALMS = 1 ] && PORTS_TCP="4098 8080" && PORTS_UDP="4099"
[ $WANT_PALMS = 1 ] && [ "$FILE_PAGE" = on ] && PORTS_TCP="$PORTS_TCP 8081"
[ $WANT_PAKS = 1 ] && PORTS_TCP="$PORTS_TCP 80 139 445" && PORTS_UDP="$PORTS_UDP 137 138"
if have firewall-cmd && firewall-cmd --state >/dev/null 2>&1; then
    for p in $PORTS_TCP; do firewall-cmd -q --permanent --add-port=$p/tcp; done
    for p in $PORTS_UDP; do firewall-cmd -q --permanent --add-port=$p/udp; done
    firewall-cmd -q --reload
    say "   Firewall (firewalld): opened TCP $PORTS_TCP, UDP $PORTS_UDP"
elif have ufw && ufw status 2>/dev/null | grep -q "Status: active"; then
    for p in $PORTS_TCP; do ufw allow $p/tcp >/dev/null; done
    for p in $PORTS_UDP; do ufw allow $p/udp >/dev/null; done
    say "   Firewall (ufw): opened TCP $PORTS_TCP, UDP $PORTS_UDP"
fi
if [ $WANT_PAKS = 1 ] && have getenforce && [ "$(getenforce 2>/dev/null)" = Enforcing ]; then
    if have semanage; then
        semanage fcontext -a -t public_content_rw_t "$PAKS(/.*)?" 2>/dev/null || semanage fcontext -m -t public_content_rw_t "$PAKS(/.*)?" 2>/dev/null || true
        semanage fcontext -a -t httpd_sys_rw_content_t "$STATE(/.*)?" 2>/dev/null || semanage fcontext -m -t httpd_sys_rw_content_t "$STATE(/.*)?" 2>/dev/null || true
        restorecon -R "$PAKS" "$STATE" 2>/dev/null || true
        setsebool -P httpd_can_network_connect 1 smbd_anon_write 1 2>/dev/null || true
        say "   SELinux: cartridge folders labelled for the web server and Samba"
    else
        say "   SELinux is on and semanage is missing. If cartridges don't download, run: setenforce 0"
    fi
fi

# ------------------------------------------------------------------ start and check

step "9. Starting"
SERVICES=""
[ $WANT_PAKS = 1 ] && SERVICES="palms-paks palms-web $SMB_SERVICES"
[ $WANT_PALMS = 1 ] && SERVICES="$SERVICES palms"
for s in $SERVICES; do svc_start "$s" >/dev/null 2>&1 || true; done
sleep 3
OK=1
for s in $SERVICES; do
    if svc_running "$s"; then
        say "   $s: running"
    elif [ "$s" = nmbd ] || [ "$s" = nmb ]; then
        say "   $s: not running. \\\\$HOSTNAME_NEW may not be found by name; \\\\$IP\\incoming still works."
    else
        say "   $s: NOT RUNNING. For details: $(svc_log_hint "$s")"
        OK=0
    fi
done

echo
if [ $OK = 1 ]; then echo "Setup complete."; else echo "Setup finished with problems (see above)."; fi
echo
if [ $WANT_PALMS = 1 ]; then
    echo "   Palms Server:      $IP, TCP port 4098 (PCs find it by themselves)"
    echo "   Projector:         http://$IP:8080/   (open it in Internet Explorer, press F11)"
    if [ "$FILE_PAGE" = on ]; then echo "   File page:         http://$IP:8081/"; fi
    echo "   Host password:     $HOST_PW   (in /opt/palms/palms.ini)"
fi
if [ $WANT_PAKS = 1 ]; then
    echo "   Upload share:      \\\\$IP\\incoming   (host approval required)"
    echo "   Administration:    http://$IP/admin"
    echo "   Cartridge index:   http://$IP/"
    echo "   Add a host:        $PY /opt/palms-paks/palms_paks.py --add-admin NAME"
    echo "   Free space:        $(df -h "$PAKS" | awk 'NR==2 {print $4 " of " $2}')"
fi
echo "   Upgrade later:     run this setup again"
echo
if [ "$VIRT" != lxc ] && [ $SET_NET = 1 ]; then
    echo "The new network settings ($IP) apply after a restart."
    if yes_no "Restart now?" y; then
        echo "Restarting. Connect again with: ssh root@$IP"
        sleep 2
        reboot
    else
        echo "Restart later with: reboot"
    fi
fi
}

main "$@"
