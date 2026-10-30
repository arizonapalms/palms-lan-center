Palms LAN Center Server Setup
=============================

Installs or upgrades the Palms Server, the Cartridge Server, or both, on a Linux PC,
virtual machine or Proxmox container.

  plc.sh              The setup. Asks for everything, shows a summary, then installs.
                      Run it again later to upgrade (settings, chat, schedule and
                      cartridges are kept).
  palms_server.py     Palms Server
  palms_paks.py       Palms Cartridge Server

Works on Debian, Ubuntu, Fedora, RHEL/Rocky/Alma, Arch, Alpine and openSUSE,
32-bit or 64-bit, with systemd or OpenRC. Needs Python 3.7 or later (Debian 10,
Ubuntu 18.04 and later). A Pentium 4 with Debian 12 32-bit is plenty.
Alpine: apk add bash curl first.


WITH INTERNET ACCESS
--------------------
Log in as root (or a user that can use sudo) and type:

  bash -c "$(curl -fsSL http://palms-arcade.com/plc.sh)"


FROM THIS PACKAGE
-----------------
  tar xzf palms-server-<version>.tar.gz
  bash plc.sh

The machine still needs internet access while setup installs its Linux packages
(Python, and nginx, Samba and 7-Zip for the Cartridge Server).


WHAT IT ASKS
------------
  Role                Palms Server, Cartridge Server, or both on one machine.
  Network card        The card on the party LAN.
  Address             A fixed address (IP, subnet mask, gateway, DNS), or the
                      one it has now. Gateway and DNS can be left out on a LAN
                      without a router.
  Computer name       Up to 15 characters (Windows 98 can't see longer names).
  Palms Server        Its name, the host password for Palms LAN Center, the
                      Cartridge Server's address, and whether to turn on the
                      file page (port 8081).
  Cartridge Server    The folder for the cartridges (default /srv/paks), and the
                      first host account for the administration page.

A Proxmox container gets its address and name from Proxmox. Setup shows what to
set there, and stops until it has been set.

The new address applies when the machine restarts. Setup offers to restart.


AFTERWARDS
----------
  Palms Server        /opt/palms/palms.ini          service: palms
  Cartridge Server    /etc/palms-paks.ini           service: palms-paks
  Web server          /etc/palms-paks/nginx.conf    service: palms-web (port 80)

  systemctl status palms           running?     (OpenRC: rc-service palms status)
  journalctl -u palms -f           log          (OpenRC: /var/log/palms.log)
  systemctl restart palms          after changing palms.ini

  palms.ini: file_page = on turns on the file page (port 8081),
             projector_port = 0 turns off the projector page.
  python3 /opt/palms-paks/palms_paks.py --add-admin NAME
                                   add a Cartridge Server host account, or change its password

The Palms Server finds the Cartridge Server through paks_url in palms.ini.
