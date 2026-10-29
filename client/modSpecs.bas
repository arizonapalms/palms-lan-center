Attribute VB_Name = "modSpecs"
' Palms LAN Center - Finds this PC's processor, memory, graphics, sound and network card for Who's Here.
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Type DISPLAY_DEVICE
    cb As Long
    DeviceName As String * 32
    DeviceString As String * 128
    StateFlags As Long
    DeviceID As String * 128
    DeviceKey As String * 128
End Type

Private Type WAVEOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * 32
    dwFormats As Long
    wChannels As Integer
    wReserved1 As Integer
    dwSupport As Long
End Type

Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long
Private Declare Function GetBestInterface Lib "iphlpapi.dll" (ByVal dwDestAddr As Long, pdwBestIfIndex As Long) As Long
Private Declare Function GetIfEntry Lib "iphlpapi.dll" (pIfRow As Any) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesA" (ByVal lpDevice As Long, ByVal iDevNum As Long, lpDisplayDevice As DISPLAY_DEVICE, ByVal dwFlags As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const KEY_READ As Long = &H20019
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const DISPLAY_DEVICE_PRIMARY_DEVICE As Long = &H4
Private Const HORZRES As Long = 8
Private Const VERTRES As Long = 10
Private Const BITSPIXEL As Long = 12
Private Const PLANES As Long = 14
Private Const VREFRESH As Long = 116
Private Const CPU_KEY As String = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"

' Everything for the SPECS message, with "|" taken out. serverIp picks the network card
' that talks to the server, like 10.0.200.71.
Public Function SpecsLine(ByVal serverIp As String) As String
    Dim nic As String, mbps As Long
    NetworkCard serverIp, nic, mbps
    SpecsLine = NoBar(CpuName()) & "|" & NoBar(MemoryText()) & "|" & NoBar(GraphicsName()) & "|" & NoBar(WindowsName()) & "|" & _
                NoBar(ScreenText()) & "|" & NoBar(SoundName()) & "|" & NoBar(nic) & "|" & mbps
End Function

Private Function NoBar(ByVal s As String) As String
    NoBar = Trim$(Replace(s, "|", "/"))
End Function

' Like "AMD Athlon XP 2400+" or "Intel Pentium III processor, 800 MHz"
Public Function CpuName() As String
    Dim s As String, mhz As Long, vendor As String
    On Error Resume Next
    s = RegText(CPU_KEY, "ProcessorNameString")
    If Len(s) = 0 Then
        ' Windows 98 only knows the family, like "x86 Family 6 Model 8 Stepping 3"
        vendor = RegText(CPU_KEY, "VendorIdentifier")
        Select Case vendor
        Case "GenuineIntel": vendor = "Intel"
        Case "AuthenticAMD": vendor = "AMD"
        Case "CyrixInstead": vendor = "Cyrix"
        End Select
        s = Trim$(vendor & " " & RegText(CPU_KEY, "Identifier"))
    End If
    s = Replace(s, "(R)", "")
    s = Replace(s, "(r)", "")
    s = Replace(s, "(TM)", "")
    s = Replace(s, "(tm)", "")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    mhz = RegNumber(CPU_KEY, "~MHz")
    If mhz > 0 And InStr(1, s, "Hz", vbTextCompare) = 0 Then
        If Len(s) > 0 Then s = s & ", "
        s = s & mhz & " MHz"
    End If
    CpuName = s
End Function

' Like "512 MB" or "1.5 GB"
Public Function MemoryText() As String
    Dim ms As MEMORYSTATUS, bytes As Double, mb As Long
    On Error Resume Next
    ms.dwLength = Len(ms)
    GlobalMemoryStatus ms
    bytes = ms.dwTotalPhys
    If bytes < 0 Then bytes = bytes + 4294967296#
    If bytes <= 0 Then Exit Function
    ' Windows keeps a little back (and shared graphics take some), so round up to the next 16 MB
    mb = -Int(-(bytes / 1048576#) / 16) * 16
    If mb < 1024 Then
        MemoryText = mb & " MB"
    ElseIf mb Mod 1024 = 0 Then
        MemoryText = (mb \ 1024) & " GB"
    Else
        MemoryText = Format$(mb / 1024, "0.0") & " GB"
    End If
End Function

' Like "NVIDIA GeForce4 Ti 4200"
Public Function GraphicsName() As String
    Dim dd As DISPLAY_DEVICE, i As Long, s As String, first As String
    On Error GoTo Done
    For i = 0 To 9
        dd.cb = Len(dd)
        If EnumDisplayDevices(0, i, dd, 0) = 0 Then Exit For
        s = Trim$(CutNull(dd.DeviceString))
        If Len(first) = 0 Then first = s
        If (dd.StateFlags And DISPLAY_DEVICE_PRIMARY_DEVICE) <> 0 And Len(s) > 0 Then
            GraphicsName = s
            Exit Function
        End If
    Next
Done:
    GraphicsName = first
End Function

' Like "Windows XP, Service Pack 3" or "Windows 98 SE"
Public Function WindowsName() As String
    Dim v As OSVERSIONINFO, s As String, sp As String
    On Error Resume Next
    v.dwOSVersionInfoSize = Len(v)
    If GetVersionEx(v) = 0 Then Exit Function
    sp = Trim$(CutNull(v.szCSDVersion))
    If v.dwPlatformId = 1 Then
        ' Windows 95, 98 and Me
        Select Case v.dwMinorVersion
        Case 0: s = "Windows 95"
        Case 10
            If (v.dwBuildNumber And &HFFFF&) >= 2222 Then
                s = "Windows 98 SE"
            Else
                s = "Windows 98"
            End If
        Case 90: s = "Windows Me"
        Case Else: s = "Windows 4." & v.dwMinorVersion
        End Select
        sp = ""
    Else
        Select Case v.dwMajorVersion * 100 + v.dwMinorVersion
        Case 400: s = "Windows NT 4.0"
        Case 500: s = "Windows 2000"
        Case 501: s = "Windows XP"
        Case 502: s = "Windows XP x64 or Server 2003"
        Case 600: s = "Windows Vista"
        Case 601: s = "Windows 7"
        Case Else: s = "Windows " & v.dwMajorVersion & "." & v.dwMinorVersion
        End Select
    End If
    If Len(sp) > 0 Then s = s & ", " & sp
    WindowsName = s
End Function

' Like "1024 x 768, 32-bit, 85 Hz"
Public Function ScreenText() As String
    Dim dc As Long, w As Long, h As Long, bits As Long, hz As Long, s As String
    On Error Resume Next
    dc = GetDC(0)
    If dc = 0 Then Exit Function
    w = GetDeviceCaps(dc, HORZRES)
    h = GetDeviceCaps(dc, VERTRES)
    bits = GetDeviceCaps(dc, BITSPIXEL) * GetDeviceCaps(dc, PLANES)
    hz = GetDeviceCaps(dc, VREFRESH)
    ReleaseDC 0, dc
    s = w & " x " & h
    If bits > 0 Then s = s & ", " & bits & "-bit"
    If hz > 1 Then s = s & ", " & hz & " Hz"     ' Windows 98 says 0 or 1 here, meaning "the usual"
    ScreenText = s
End Function

' The first sound card, like "SB Live! Wave Device"
Public Function SoundName() As String
    Dim caps As WAVEOUTCAPS
    On Error GoTo Done
    If waveOutGetNumDevs() < 1 Then
        SoundName = "None"
        Exit Function
    End If
    If waveOutGetDevCaps(0, caps, Len(caps)) = 0 Then SoundName = Trim$(CutNull(caps.szPname))
Done:
End Function

' The network card used to reach the server, and how fast its link is (10, 100 or 1000 Mb).
' Windows keeps these in a 860-byte record (MIB_IFROW): the speed is at byte 524,
' the description's length at 600 and the description itself from 604.
Public Sub NetworkCard(ByVal serverIp As String, nic As String, mbps As Long)
    Dim addr As Long, idx As Long, row(0 To 859) As Byte, speed As Long, n As Long, i As Long, bps As Double
    nic = ""
    mbps = 0
    On Error GoTo Done
    addr = inet_addr(serverIp)
    If addr = -1 Or addr = 0 Then Exit Sub
    If GetBestInterface(addr, idx) <> 0 Then Exit Sub
    CopyMemory row(512), idx, 4
    If GetIfEntry(row(0)) <> 0 Then Exit Sub
    CopyMemory speed, row(524), 4
    bps = speed
    If bps < 0 Then bps = bps + 4294967296#
    mbps = CLng(bps / 1000000#)
    CopyMemory n, row(600), 4
    If n > 256 Then n = 256
    For i = 0 To n - 1
        If row(604 + i) = 0 Then Exit For
        nic = nic & Chr$(row(604 + i))
    Next
    nic = Trim$(nic)
Done:
End Sub

Private Function CutNull(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, vbNullChar)
    If p > 0 Then
        CutNull = Left$(s, p - 1)
    Else
        CutNull = s
    End If
End Function

Private Function RegText(ByVal keyPath As String, ByVal valueName As String) As String
    Dim hk As Long, typ As Long, n As Long, buf As String
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, hk) <> 0 Then Exit Function
    n = 256
    buf = String$(n, vbNullChar)
    If RegQueryValueEx(hk, valueName, 0, typ, ByVal buf, n) = 0 Then
        If typ = REG_SZ And n > 0 Then RegText = Trim$(CutNull(Left$(buf, n)))
    End If
    RegCloseKey hk
End Function

Private Function RegNumber(ByVal keyPath As String, ByVal valueName As String) As Long
    Dim hk As Long, typ As Long, n As Long, v As Long
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, hk) <> 0 Then Exit Function
    n = 4
    If RegQueryValueEx(hk, valueName, 0, typ, v, n) = 0 Then
        If typ = REG_DWORD Then RegNumber = v
    End If
    RegCloseKey hk
End Function
