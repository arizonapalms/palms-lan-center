# Palms LAN Center

A LAN party program for Windows 98, 2000 and XP, made for [Palms Arcade](http://palms-arcade.com) retro LAN parties.
Everyone at the party runs it on their PC; it talks to a small server on the LAN.

- **Chat** across the room, with name alerts and announcements from the host
- **What's On**: the day's schedule, reminders, next-game votes, game suggestions and the food order
- **Game Library**: install games from the Cartridge Server, and make new cartridges with the Make Cartridge Wizard
- **Who's Here**: who's signed in, their seat on the seat map, and what PC they brought
- **Tournaments**: sign-ups, single-elimination brackets and results
- **Projector**: slides for the big screen (now playing, schedule, brackets, the vote), full screen in the program or as a web page for Internet Explorer 5
- Pop-up notifications

It is written to run well on a Pentium III with an 800 x 600 screen.

## What's here

| Folder | What it is | Written in |
|---|---|---|
| `client/` | Palms LAN Center, the program on each PC | Visual Basic 6.0 |
| `palmsget/` | `palmsget.exe`, the small downloader the Game Library uses | C (MinGW) |
| `server/` | Palms Server: chat, schedule, votes, tournaments, the projector page | Python 3 |
| `cartridge-server/` | Palms Cartridge Server: stores the cartridges, host approval | Python 3 |
| `deploy/` | `plc.sh`, installs or upgrades either server (or both) on Linux | Bash |
| `setup/` | The Windows installer script | NSIS 3 |

## Building the client

1. Open `client/PalmsLAN.vbp` in Visual Basic 6.0 (Service Pack 6).
2. File > Make PalmsLAN.exe.

Next to `PalmsLAN.exe`, the program needs:

- `palmsget.exe`: build it with `palmsget/build.sh` (needs `i686-w64-mingw32-gcc`)
- `7za.exe`: 7-Zip 9.20 command line version (`7za920.zip` from 7-zip.org)
- `MSWINSCK.OCX`: the Microsoft Winsock Control, registered on the PC
- the `.wav` files from `client/`

`7za.exe`, `MSWINSCK.OCX` and the Visual Basic 6.0 runtime are not part of this repository.

## Running the servers

On any Linux PC (Debian, Ubuntu, Fedora, RHEL/Rocky/Alma, Arch, Alpine, openSUSE; 32-bit or 64-bit; Python 3.7 or later), as root:

```
bash deploy/plc.sh
```
 It asks which server to install, the network address and the computer name,
shows a summary, then installs. Run it again to upgrade. See `deploy/README.txt`.

To try the Palms Server on any machine with Python 3: `python3 server/palms_server.py`.

## Building the installer

```
makensis -DVERSION=0.9.3-alpha -DNUMVER=0.9.3 -DSRC=<folder with the release files> -DOUT=PalmsLAN-0.9.3-alpha-setup.exe setup/PalmsLAN.nsi
```

The release folder holds `PalmsLAN.exe`, `palmsget.exe`, `7za.exe`, `MSWINSCK.OCX`, the `.wav` files and
`palms-clean.vbs`, and optionally `vbrun60sp6.exe` (Microsoft's Visual Basic 6.0 runtime) for Windows 98 PCs.

## Licence

Copyright (C) 2026 Patrik Perkovic, Palms Arcade.

This program is free software: you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation, either version 3 of the License,
or (at your option) any later version. See [LICENSE](LICENSE).
