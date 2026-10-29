#!/bin/bash
# Palms Arcade Server Setup: installs the Palms Server, the Cartridge Server, or both, on a new Debian 12 machine.
set -e
shopt -s nullglob
HERE="$(cd "$(dirname "$0")" && pwd)"
# the server programs: beside this script, or in the source tree's server folders
SERVER_PY="$HERE/palms_server.py"; [ -f "$SERVER_PY" ] || SERVER_PY="$HERE/../server/palms_server.py"
PAKS_PY="$HERE/palms_paks.py"; [ -f "$PAKS_PY" ] || PAKS_PY="$HERE/../cartridge-server/palms_paks.py"
ROOT="${PALMS_ROOT:-}"          # for testing the network part against a fake /etc only

# ------------------------------------------------------------------ helpers

say()  { echo "$@"; }
step() { echo; echo "== $*"; }
fail() { echo; echo "$*"; exit 1; }

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

# yes "Question" y|n: 0 for yes
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

# the same network: is_same_net 10.0.200.71 10.0.200.1 24
same_net() {
    local a b p="$3" mask=0 i
    a=$(ip_num "$1"); b=$(ip_num "$2")
    for ((i = 0; i < p; i++)); do mask=$(( mask | (1 << (31 - i)) )); done
    [ $(( a & mask )) -eq $(( b & mask )) ]
}
ip_num() { local IFS=.; set -- $1; echo $(( ($1 << 24) | ($2 << 16) | ($3 << 8) | $4 )); }

# ------------------------------------------------------------------ network files

# The network manager in use: networkmanager, networkd or ifupdown
net_kind() {
    if systemctl is-active --quiet NetworkManager 2>/dev/null; then echo networkmanager
    elif systemctl is-active --quiet systemd-networkd 2>/dev/null && [ ! -f "$ROOT/etc/network/interfaces" ]; then echo networkd
    else echo ifupdown; fi
}

# write_network KIND IFACE IP PREFIX GATEWAY DNS (gateway and DNS may be blank)
write_network() {
    local kind="$1" ifc="$2" ip="$3" p="$4" gw="$5" dns="$6" stamp
    stamp="$(date +%Y%m%d-%H%M%S)"
    case "$kind" in
    ifupdown)
        mkdir -p "$ROOT/etc/network/interfaces.d"
        touch "$ROOT/etc/network/interfaces"
        cp "$ROOT/etc/network/interfaces" "$ROOT/etc/network/interfaces.before-palms-$stamp"
        # the card's old settings in /etc/network/interfaces are commented out (other cards stay as they are)
        awk -v ifc="$ifc" '
            /^[[:space:]]*(iface|auto|allow-[a-z]+|mapping|source|source-directory)[[:space:]]/ { inside = 0 }
            ($1 == "auto" || $1 ~ /^allow-/) && NF == 2 && $2 == ifc { print "# (Palms Arcade Server Setup) " $0; next }
            $1 == "iface" && $2 == ifc { inside = 1 }
            inside && NF > 0 { print "# (Palms Arcade Server Setup) " $0; next }
            { print }
        ' "$ROOT/etc/network/interfaces.before-palms-$stamp" > "$ROOT/etc/network/interfaces"
        grep -Eq '^[[:space:]]*source[[:space:]]+/etc/network/interfaces\.d' "$ROOT/etc/network/interfaces" || \
            printf '\nsource /etc/network/interfaces.d/*\n' >> "$ROOT/etc/network/interfaces"
        # another file for the same card would fight the new one
        for f in "$ROOT"/etc/network/interfaces.d/*; do
            if grep -Eq "^[[:space:]]*(auto|allow-hotplug|iface)[[:space:]]+$ifc([[:space:]]|$)" "$f"; then
                mv "$f" "$ROOT/etc/network/$(basename "$f").before-palms-$stamp"
                say "   Moved $(basename "$f") out of interfaces.d (it also set up $ifc)."
            fi
        done
        {
            echo "# Written by Palms Arcade Server Setup on $(date '+%Y-%m-%d %H:%M')."
            echo "# The previous settings: /etc/network/interfaces.before-palms-$stamp"
            echo "auto $ifc"
            echo "iface $ifc inet static"
            echo "    address $ip/$p"
            if [ -n "$gw" ]; then echo "    gateway $gw"; fi
        } > "$ROOT/etc/network/interfaces.d/palms-$ifc"
        if [ -n "$dns" ]; then
            [ -f "$ROOT/etc/resolv.conf" ] && cp -L "$ROOT/etc/resolv.conf" "$ROOT/etc/resolv.conf.before-palms-$stamp"
            rm -f "$ROOT/etc/resolv.conf"
            { echo "# Written by Palms Arcade Server Setup"; for d in $dns; do echo "nameserver $d"; done; } > "$ROOT/etc/resolv.conf"
        fi
        ;;
    networkd)
        mkdir -p "$ROOT/etc/systemd/network"
        for f in "$ROOT"/etc/systemd/network/*.network; do
            if grep -Eq "^Name=$ifc\$" "$f"; then mv "$f" "$f.before-palms-$stamp"; fi
        done
        {
            echo "# Written by Palms Arcade Server Setup"
            echo "[Match]"
            echo "Name=$ifc"
            echo
            echo "[Network]"
            echo "Address=$ip/$p"
            if [ -n "$gw" ]; then echo "Gateway=$gw"; fi
            for d in $dns; do echo "DNS=$d"; done
        } > "$ROOT/etc/systemd/network/10-palms-$ifc.network"
        ;;
    networkmanager)
        local con
        con="$(nmcli -g GENERAL.CONNECTION device show "$ifc" 2>/dev/null | head -n1)"
        if [ -z "$con" ]; then
            con="palms-$ifc"
            nmcli con add type ethernet ifname "$ifc" con-name "$con" >/dev/null
        fi
        nmcli con mod "$con" ipv4.method manual ipv4.addresses "$ip/$p" ipv4.gateway "${gw:-}" \
            ipv4.dns "${dns:-}" ipv6.method ignore connection.autoconnect yes
        ;;
    esac
    # cloud-init (cloud images, some Proxmox VM templates) would put its own settings back at startup
    if [ -d "$ROOT/etc/cloud/cloud.cfg.d" ]; then
        echo "network: {config: disabled}" > "$ROOT/etc/cloud/cloud.cfg.d/99-palms-disable-network-config.cfg"
        say "   cloud-init network setup turned off (it would undo this address at startup)."
    fi
}

write_hostname() {
    local name="$1"
    echo "$name" > "$ROOT/etc/hostname"
    if [ -f "$ROOT/etc/hosts" ] && grep -q '^127\.0\.1\.1' "$ROOT/etc/hosts"; then
        sed -i "s/^127\.0\.1\.1.*/127.0.1.1\t$name/" "$ROOT/etc/hosts"
    else
        printf '127.0.1.1\t%s\n' "$name" >> "$ROOT/etc/hosts"
    fi
    [ -z "$ROOT" ] && hostname "$name" 2>/dev/null || true
}

# only the network part, against a fake root: PALMS_ROOT=/tmp/x palms-setup.sh --test-network ifupdown eth0 10.0.0.5 24 10.0.0.1 "1.1.1.1"
if [ "$1" = "--test-network" ]; then
    write_network "$2" "$3" "$4" "$5" "$6" "$7"
    write_hostname "${8:-palms-test}"
    exit 0
fi

# ------------------------------------------------------------------ checks

echo
echo "Palms Arcade Server Setup"
echo "========================="
[ "$(id -u)" = 0 ] || fail "This script must be run as root."
[ -t 0 ] || fail "Run this script in a terminal (it asks questions)."
if [ -f /etc/os-release ]; then
    . /etc/os-release
    if [ "$ID" != debian ] || [ "${VERSION_ID%%.*}" != 12 ]; then
        say "   This script is made for Debian 12. This machine runs ${PRETTY_NAME:-an unknown system}."
        yes_no "Continue anyway?" n || exit 1
    fi
fi
VIRT="$(systemd-detect-virt --container 2>/dev/null || true)"

# ------------------------------------------------------------------ questions

step "1. Role"
say "   1  Palms Server       (chat, schedule, tournaments, file page)"
say "   2  Cartridge Server   (Game Library cartridges, upload share, approval)"
say "   3  Both on this machine"
ROLE=""
while ! [[ "$ROLE" =~ ^[123]$ ]]; do ask ROLE "Install which" "1"; done
WANT_PALMS=0; WANT_PAKS=0
[ "$ROLE" = 1 ] || [ "$ROLE" = 3 ] && WANT_PALMS=1
[ "$ROLE" = 2 ] || [ "$ROLE" = 3 ] && WANT_PAKS=1
if [ $WANT_PALMS = 1 ] && [ ! -f "$SERVER_PY" ]; then
    fail "palms_server.py not found. Copy it to the same folder as this script ($HERE)."
fi
if [ $WANT_PAKS = 1 ] && [ ! -f "$PAKS_PY" ]; then
    fail "palms_paks.py not found. Copy it to the same folder as this script ($HERE)."
fi
case "$ROLE" in 1) DEF_NAME=palms-server ;; 2) DEF_NAME=palms-carts ;; *) DEF_NAME=palms-lan ;; esac

step "2. Network"
# the card with the default route, otherwise the first one that isn't lo
IFACE="$(ip -o route show default 2>/dev/null | awk '{for (i = 1; i < NF; i++) if ($i == "dev") {print $(i + 1); exit}}')"
[ -n "$IFACE" ] || IFACE="$(ip -o link show | awk -F': ' '$2 != "lo" {sub(/@.*/, "", $2); print $2; exit}')"
CUR="$(ip -o -4 addr show dev "$IFACE" 2>/dev/null | awk '{print $4; exit}')"
CUR_IP="${CUR%/*}"
SET_NET=0
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
        echo "Set the address and name in Proxmox, restart the container, then run this script again."
        exit 0
    fi
    [ -n "$CUR_IP" ] || fail "The container has no IPv4 address. Set one in Proxmox first."
    IP="$CUR_IP"; HOSTNAME_NEW="$(hostname)"
else
    say "   Network cards:"
    ip -br -4 addr show | awk '$1 != "lo" {printf "     %-12s %-8s %s\n", $1, $2, $3}'
    for c in $(ip -o link show | awk -F': ' '$2 != "lo" {sub(/@.*/, "", $2); print $2}'); do
        ip -o -4 addr show dev "$c" | grep -q . || printf "     %-12s (no address)\n" "$c"
    done
    while :; do
        ask IFACE "Network card for the LAN" "$IFACE"
        ip link show "$IFACE" >/dev/null 2>&1 && break
        say "   There is no network card called $IFACE."
    done
    CUR="$(ip -o -4 addr show dev "$IFACE" 2>/dev/null | awk '{print $4; exit}')"
    CUR_IP="${CUR%/*}"
    say "   Address now: ${CUR:-none}"
    if yes_no "Set a fixed (static) address on $IFACE?" y; then
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
        CUR_DNS="$(awk '/^nameserver/ {printf "%s ", $2}' /etc/resolv.conf 2>/dev/null | sed 's/ $//')"
        while :; do
            ask DNS "DNS servers, separated by spaces (- for none)" "${CUR_DNS:--}"
            [ "$DNS" = - ] && DNS="" && break
            ok=1; for d in $DNS; do is_ip "$d" || ok=0; done
            [ $ok = 1 ] && break
            say "   Enter addresses like 10.0.200.1 8.8.8.8, or - for none."
        done
    else
        [ -n "$CUR_IP" ] || fail "$IFACE has no IPv4 address. Run the script again and set a fixed address."
        IP="$CUR_IP"
        say "   Keeping $CUR. Make sure this address does not change (a DHCP reservation, for example)."
    fi
    while :; do
        ask HOSTNAME_NEW "Computer name" "$DEF_NAME"
        [[ "$HOSTNAME_NEW" =~ ^[A-Za-z0-9]([A-Za-z0-9-]{0,13}[A-Za-z0-9])?$ ]] && break
        say "   1 to 15 letters, numbers and - (Windows 98 can't see longer names)."
    done
fi

step "3. Settings"
if [ $WANT_PALMS = 1 ]; then
    ask SRV_NAME "Palms Server name (shown in Palms LAN Center)" "PALMS-SRV"
    SRV_NAME="$(echo "$SRV_NAME" | tr -cd 'A-Za-z0-9 _-' | cut -c1-24)"
    [ -n "$SRV_NAME" ] || SRV_NAME="PALMS-SRV"
    while :; do
        ask HOST_PW "Host password for Palms LAN Center (Enter for a random one)" ""
        [[ "$HOST_PW" =~ ^[^|[:space:]]{0,32}$ ]] && break
        say "   Up to 32 characters, without spaces or |."
    done
    if [ $WANT_PAKS = 1 ]; then
        PAKS_HOST="$IP"
    else
        while :; do
            ask PAKS_HOST "Cartridge Server address" "10.0.200.73"
            is_ip "$PAKS_HOST" && break
            [[ "$PAKS_HOST" =~ ^[A-Za-z0-9.-]+$ ]] && break
            say "   Enter an address like 10.0.200.73."
        done
    fi
fi
if [ $WANT_PAKS = 1 ]; then
    while :; do
        ask PAKS "Cartridge folder (needs space for every cartridge)" "/srv/paks"
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
    say "   Network card:      $IFACE ($(net_kind))"
    say "   Address:           $IP, mask $(mask_of "$PREFIX"), gateway ${GW:-none}, DNS ${DNS:-none}"
    say "   Computer name:     $HOSTNAME_NEW"
else
    say "   Address:           $IP on $IFACE (kept as it is), name $HOSTNAME_NEW"
fi
if [ $WANT_PALMS = 1 ]; then
    say "   Palms Server:      \"$SRV_NAME\", /opt/palms, host password ${HOST_PW:-(random)}"
    say "                      Game Library from the Cartridge Server at http://$PAKS_HOST/"
fi
[ $WANT_PAKS = 1 ] && say "   Cartridge Server:  cartridges in $PAKS, address http://$IP/"
say
yes_no "Install now?" y || { echo "Nothing was changed."; exit 0; }

# ------------------------------------------------------------------ the network (applied at the end)

if [ "$VIRT" != lxc ]; then
    step "5. Computer name and network settings"
    write_hostname "$HOSTNAME_NEW"
    say "   Computer name: $HOSTNAME_NEW"
    if [ $SET_NET = 1 ]; then
        write_network "$(net_kind)" "$IFACE" "$IP" "$PREFIX" "$GW" "$DNS"
        say "   $IP/$PREFIX on $IFACE is used from the next restart."
    fi
fi

# ------------------------------------------------------------------ packages

step "6. Installing packages"
export DEBIAN_FRONTEND=noninteractive
PKGS="python3"
[ $WANT_PAKS = 1 ] && PKGS="$PKGS nginx samba p7zip-full python3-pil python3-pefile"
if ! apt-get update -q; then
    say "   Unable to download the package lists. The machine needs internet access for this step"
    say "   (a gateway and DNS server). Connect it, then run the script again."
    exit 1
fi
# shellcheck disable=SC2086
apt-get install -y -q $PKGS
[ $WANT_PAKS = 1 ] && { apt-get install -y -q samba-vfs-modules >/dev/null 2>&1 || true; }
id palms >/dev/null 2>&1 || useradd --system --home-dir /opt/palms --shell /usr/sbin/nologin palms

# ------------------------------------------------------------------ Cartridge Server

if [ $WANT_PAKS = 1 ]; then
    step "7. Cartridge Server"
    STATE=/var/lib/palms-paks
    APP=/opt/palms-paks
    mkdir -p "$PAKS" "$STATE/icons" "$STATE/manifests" "$APP"
    for d in inbox waiting library trash; do mkdir -p "$PAKS/$d"; done
    chown palms:palms "$PAKS" "$PAKS/inbox" "$PAKS/waiting" "$PAKS/library" "$PAKS/trash" "$STATE" "$STATE/icons" "$STATE/manifests"
    chmod 755 "$PAKS" "$PAKS/waiting" "$PAKS/library" "$PAKS/trash"
    chmod 2775 "$PAKS/inbox"
    install -m 755 "$PAKS_PY" "$APP/palms_paks.py"
    cat > /etc/palms-paks.ini <<EOF
; Palms Cartridge Server settings. After making changes, run: systemctl restart palms-paks
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
        say "   Host accounts: $(python3 "$APP/palms_paks.py" --admins | tr '\n' ' ')"
    else
        say "   Cartridges are approved by the host at http://$IP/admin. Create the first host account."
        NAME=""
        while ! echo "$NAME" | grep -Eq '^[A-Za-z0-9._-]{1,32}$'; do
            read -r -p "   Account name (for example, host): " NAME
        done
        python3 "$APP/palms_paks.py" --add-admin "$NAME"
    fi
    [ -f /etc/palms-paks/admins ] && chown root:palms /etc/palms-paks/admins && chmod 640 /etc/palms-paks/admins

    cat > /etc/nginx/sites-available/palms-paks <<EOF
# Palms Cartridge Server: index page, catalog, icons, approved cartridges and the admin page
server {
    listen 80 default_server;
    server_name _;
    # send large files directly from disk
    sendfile on;
    tcp_nopush on;
    sendfile_max_chunk 2m;
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
EOF
    ln -sf /etc/nginx/sites-available/palms-paks /etc/nginx/sites-enabled/palms-paks
    rm -f /etc/nginx/sites-enabled/default
    nginx -t 2>&1 | sed 's/^/   /'
    systemctl enable nginx >/dev/null 2>&1 || true
    systemctl restart nginx
    say "   Web server: nginx on port 80"

    [ -f /etc/samba/smb.conf ] && [ ! -f /etc/samba/smb.conf.debian ] && cp /etc/samba/smb.conf /etc/samba/smb.conf.debian
    AUDIT=""
    if ls /usr/lib/*/samba/vfs/full_audit.so /usr/lib/samba/vfs/full_audit.so >/dev/null 2>&1; then
        AUDIT="   # writes, renames and deletions (with the client address) are logged to the system log
   # and read by the palms-paks service for the audit log
   vfs objects = full_audit
   full_audit:prefix = %I
   full_audit:success = mkdirat renameat unlinkat openat
   full_audit:failure = none
   full_audit:facility = LOCAL5
   full_audit:priority = NOTICE"
    else
        say "   The Samba full_audit module is not installed. Share activity is not recorded in the audit log."
    fi
    cat > /etc/samba/smb.conf <<EOF
# Palms Cartridge Server: "incoming" is the upload share, open to everyone on the network;
# "library" contains the approved cartridges (read-only). The Debian default is in smb.conf.debian.
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
    testparm -s >/dev/null 2>&1 || { echo "Samba configuration error:"; testparm -s; exit 1; }
    systemctl enable smbd nmbd >/dev/null 2>&1 || true
    systemctl restart smbd
    systemctl restart nmbd 2>/dev/null || say "   nmbd did not start. \\\\$HOSTNAME_NEW may not be found by name; \\\\$IP\\incoming still works."
    say "   Shares: \\\\$IP\\incoming and \\\\$IP\\library"

    cat > /etc/systemd/system/palms-paks.service <<EOF
[Unit]
Description=Palms Cartridge Server
After=network-online.target

[Service]
User=palms
Group=palms
# required to read Samba audit entries from the system log
SupplementaryGroups=systemd-journal
ExecStart=/usr/bin/python3 $APP/palms_paks.py
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF
    systemctl daemon-reload
    systemctl enable palms-paks >/dev/null 2>&1
    systemctl restart palms-paks
    say "   Service: palms-paks"
fi

# ------------------------------------------------------------------ Palms Server

if [ $WANT_PALMS = 1 ]; then
    step "8. Palms Server"
    mkdir -p /opt/palms/data /opt/palms/files/client
    install -m 755 "$SERVER_PY" /opt/palms/palms_server.py
    [ -n "$HOST_PW" ] || HOST_PW="$(tr -dc 'a-z0-9' < /dev/urandom | head -c 8)"
    cat > /opt/palms/palms.ini <<EOF
# Palms Server settings. Restart the server after making changes: systemctl restart palms
[server]
name = $SRV_NAME
tcp_port = 4098
beacon_port = 4099
beacon_seconds = 3
admin_password = $HOST_PW
http_port = 8080
paks_url = http://$PAKS_HOST/
EOF
    chmod 640 /opt/palms/palms.ini
    python3 /opt/palms/palms_server.py --install-service | sed 's/^/   /'
    chgrp palms /opt/palms/palms.ini
fi

# ------------------------------------------------------------------ done

step "9. Checking"
sleep 3
OK=1
for s in $( [ $WANT_PAKS = 1 ] && echo palms-paks nginx smbd ) $( [ $WANT_PALMS = 1 ] && echo palms ); do
    if systemctl is-active --quiet "$s"; then
        say "   $s: running"
    else
        say "   $s: NOT RUNNING. For details, run: journalctl -u $s -n 30"
        OK=0
    fi
done

echo
if [ $OK = 1 ]; then echo "Setup complete."; else echo "Setup finished with problems (see above)."; fi
echo
if [ $WANT_PALMS = 1 ]; then
    echo "   Palms Server:      $IP, TCP port 4098 (PCs find it by themselves)"
    echo "   File page:         http://$IP:8080/   (upload PalmsLAN.exe here to update the PCs)"
    echo "   Host password:     $HOST_PW   (in /opt/palms/palms.ini)"
    echo "   Log:               journalctl -u palms -f"
fi
if [ $WANT_PAKS = 1 ]; then
    echo "   Upload share:      \\\\$IP\\incoming   (host approval required)"
    echo "   Administration:    http://$IP/admin"
    echo "   Cartridge index:   http://$IP/"
    echo "   Log:               journalctl -u palms-paks -f"
    echo "   Add a host:        python3 /opt/palms-paks/palms_paks.py --add-admin NAME"
    echo "   Free space:        $(df -h "$PAKS" | awk 'NR==2 {print $4 " of " $2}')"
fi
echo
if [ "$VIRT" != lxc ] && [ $SET_NET = 1 ]; then
    echo "The new network settings ($IP) apply after a restart."
    if yes_no "Restart now?" y; then
        echo "Restarting. Connect again with: ssh root@$IP"
        sleep 2
        systemctl reboot
    else
        echo "Restart later with: systemctl reboot"
    fi
fi
