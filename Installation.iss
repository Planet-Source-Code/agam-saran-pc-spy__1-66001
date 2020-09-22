; -- Example1.iss --
; Demonstrates copying 3 files and creating an icon.

; SEE THE DOCUMENTATION FOR DETAILS ON CREATING .ISS SCRIPT FILES!

[Setup]
AppName=PC Spy
AppVerName=PC Spy
DefaultDirName={pf}\PC Spy
Compression=lzma
SolidCompression=yes

[Files]
Source: "Spy.exe"; DestDir: "{app}"
Source: "pc.log"; DestDir: "{app}"
Source: "Hand.cur"; DestDir: "{app}"
Source: "Spy.ini"; DestDir: "{app}"
Source: "Help.chm"; DestDir: "{app}"
Source: "readme.txt"; DestDir: "{app}"; flags: isreadme


[Run]
Filename: "{app}\spy.exe"; Flags: nowait

[Registry]
Root: HKLM; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueName:"PC Spy"; Flags: uninsdeletevalue

