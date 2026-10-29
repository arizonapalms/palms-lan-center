Palms Arcade Server Setup 1.0
=============================

Installs the Palms Server, the Cartridge Server, or both, on a new Debian 12
machine: a PC, a virtual machine or a Proxmox container.

  palms-setup.sh      The setup script. Asks for everything, shows a summary,
                      then installs.
  palms_server.py     Palms Server 0.5.2
  palms_paks.py       Palms Cartridge Server 2.4

Do not open palms-setup.sh in Notepad and save it: Linux can't run it with
Windows line endings.


WHAT IT ASKS
------------
  Role                Palms Server, Cartridge Server, or both on one machine.
  Network card        The card on the party LAN.
  Address             A fixed address (IP, subnet mask, gateway, DNS), or the
                      one it has now. The gateway and DNS can be left out on a
                      LAN without a router.
  Computer name       Up to 15 characters (Windows 98 can't see longer names).
  Palms Server        Its name, the host password for Palms LAN Center, and the
                      Cartridge Server's address (the same machine for "both").
  Cartridge Server    The folder for the cartridges (default /srv/paks), and the
                      first host account for the administration page.

A Proxmox container gets its address and name from Proxmox. The script shows
what to set there, and stops until it has been set.

The new address applies when the machine restarts. The script offers to
restart at the end.

The machine needs internet access while the script runs (to download nginx,
Samba and so on). After that it doesn't.


STEP BY STEP
------------
1. Install Debian 12 with "SSH server" and "standard system utilities" (no
   desktop). For a Proxmox container, use the Debian 12 template.

2. On your PC, unzip this file. You get a folder called palms-setup.

3. Copy the folder to the new machine. In a Command Prompt, in the folder
   that contains palms-setup (use the machine's address now):

     scp -r palms-setup root@10.0.200.50:/root/

   Debian doesn't allow root over SSH unless that was set up. If it says
   "Permission denied", copy it as the user made during the Debian install:

     scp -r palms-setup user@10.0.200.50:/home/user/

4. Log in to the new machine:

     ssh root@10.0.200.50

   (or ssh user@10.0.200.50, then type su - and the root password)

5. Run the script:

     bash /root/palms-setup/palms-setup.sh

   (or bash /home/user/palms-setup/palms-setup.sh)

6. Answer the questions. Press Enter to take the suggestion in [brackets].
   Check the summary, then type Y.

7. When it asks, restart. Log in again at the new address:

     ssh root@10.0.200.71

8. Palms Server: upload PalmsLAN.exe and the other LAN Center Files on the
   file page (http://<address>:8080/), the same as on the old server.
   Cartridge Server: copy cartridges to \\<address>\incoming and approve them
   at http://<address>/admin.


AFTERWARDS
----------
  systemctl status palms            Palms Server running?
  journalctl -u palms -f            Palms Server log (Ctrl+C to stop)
  /opt/palms/palms.ini              Palms Server settings (then: systemctl restart palms)

  systemctl status palms-paks       Cartridge Server running?
  journalctl -u palms-paks -f       Cartridge Server log
  /etc/palms-paks.ini               Cartridge Server settings (then: systemctl restart palms-paks)
  python3 /opt/palms-paks/palms_paks.py --add-admin NAME
                                    Add a host account, or change its password

The Palms Server finds the Cartridge Server through paks_url in
/opt/palms/palms.ini. If the Cartridge Server's address changes, change it
there and restart the Palms Server.
