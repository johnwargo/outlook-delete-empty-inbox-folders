[Setup]
AppName=Outlook Delete Empty Inbox Folders
AppVersion=.05
WizardStyle=modern
DefaultDirName={autopf}\johnwargo
DefaultGroupName=Outlook Utilities
UninstallDisplayIcon={app}\DeleteEmptyFolders.exe
Compression=lzma2
SolidCompression=yes
OutputDir=Output
OutputBaseFilename=SetupDeleteEmptyFolders

[Files]
Source: "Win64\Release\DeleteEmptyFolders.exe"; DestDir: "{app}"

[Icons]
Name: "{group}\Outlook Delete Empty Folders"; Filename: "{app}\DeleteEmptyFolders.exe"
