[Setup]
AppName=Outlook Delete Empty Inbox Folders
AppVersion=0.0.7
WizardStyle=modern
DefaultDirName={autopf}\Outlook Utilities
DefaultGroupName=Outlook Utilities
UninstallDisplayIcon={app}\DeleteEmptyFolders.exe
Compression=lzma2
SolidCompression=yes
OutputDir=installer
OutputBaseFilename=DeleteEmptyFoldersSetup
SetupIconFile=DeleteEmptyFolders_Icon.ico

[Files]
Source: "Win64\Release\DeleteEmptyFolders.exe"; DestDir: "{app}"

[Icons]
Name: "{group}\Outlook Delete Empty Folders"; Filename: "{app}\DeleteEmptyFolders.exe"
