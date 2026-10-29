#!/bin/sh
# Builds palmsget.exe with MinGW: no C runtime, no console window, runs on Windows 95 and later.
set -e
cd "$(dirname "$0")"
i686-w64-mingw32-gcc -Os -march=i586 -ffreestanding -fno-builtin -fno-stack-protector -fno-asynchronous-unwind-tables \
  -fno-ident -Wall -Wno-unused-parameter -nostdlib -e _start -s \
  -Wl,--subsystem,windows:4.0 -Wl,--major-os-version,4 -Wl,--minor-os-version,0 \
  -Wl,--major-image-version,1 -Wl,--minor-image-version,0 \
  -o palmsget.exe palmsget.c -lwininet -luser32 -lkernel32 -lgcc
ls -l palmsget.exe
