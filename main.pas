(* ************************************************************************
  * Delete Empty Outlook Inbox Folders
  *
  * By John M. Wargo
  * https://johnwargo.com
  * https://github.com/johnwargo/outlook-delete-empty-inbox-folders-delphi
  ************************************************************************* *)
unit main;

interface

uses

  // local log unit
  Log,

  // Outlook stuff
  ComObj, Outlook2010,

  CodeSiteLogging,

  // Raize Components
  RzPanel, RzButton, RzEdit, RzLstBox, RzSplit, RzLaunch, RzStatus,

  // System stuff
  ShellApi, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls,
  Vcl.StdCtrls, Winapi.ActiveX, Winapi.Windows, Winapi.Messages, RzCommon,
  RzForms, Vcl.Imaging.pngimage;

type
  TfrmMain = class(TForm)
    StatusBar: TRzStatusBar;
    Toolbar: TRzToolbar;
    RzSpacer1: TRzSpacer;
    btnClose: TRzToolButton;
    btnDelete: TRzToolButton;
    RzSpacer2: TRzSpacer;
    btnViewLog: TRzToolButton;
    Launcher: TRzLauncher;
    lblFolderCount: TRzStatusPane;
    VersionInfoStatus: TRzVersionInfoStatus;
    VersionInfo: TRzVersionInfo;
    btnAnalyze: TRzToolButton;
    lstFolders: TRzListBox;
    PanelHeader: TRzPanel;
    Label1: TLabel;
    AuthorNotice: TRzStatusPane;
    btnHelp: TRzToolButton;
    SplashPanel: TRzPanel;
    SplashImage: TImage;
    SplashTimer: TTimer;
    RzRegIniFile: TRzRegIniFile;
    RzFormState: TRzFormState;
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnViewLogClick(Sender: TObject);
    procedure btnAnalyzeClick(Sender: TObject);
    procedure btnHelpClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SplashTimerTimer(Sender: TObject);
    procedure AuthorNoticeClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

const
  HelpURL: string =
    'https://github.com/johnwargo/outlook-delete-empty-inbox-folders/wiki/Delete-Empty-Inbox-Folders-Help';

var
  frmMain: TfrmMain;
  folderCounter: integer;

implementation

{$R *.dfm}

procedure doLog(msg: String; doLog: Boolean = false);
begin
  frmMain.lstFolders.Add(msg);
  if doLog then
    LogMessage(msg);
end;

procedure SetStartButtonState(state: Boolean);
begin
  frmMain.btnDelete.Enabled := state;
end;

procedure UpdateFolderCounter();
begin
  Inc(folderCounter);
  frmMain.lblFolderCount.caption := IntToStr(folderCounter) + ' folders';
end;

function CanDeleteFolder(Folder: OleVariant): Boolean;
begin
  Result := (Folder.Folders.Count < 1) and (Folder.Items.Count < 1);
end;

procedure DeleteFolder(Folder: OleVariant; path: string; doDelete: Boolean;
  AddToList: Boolean);
begin
  if AddToList then begin
    doLog(path, false);
    UpdateFolderCounter();
    LogMessage('Deletable: ' + path);
  end;

  if doDelete then begin
    doLog('Deleting ' + path, True);
    try
      Folder.Delete;
    except
      on E: Exception do begin
        MessageDlg(E.Message, mtError, [mbOk], 0, mbOk);
        LogMessage('ERROR: ' + E.Message);
        CloseLog();
        Application.Terminate;
      end;
    end;
  end;
end;

procedure ProcessFolders(Folders: OleVariant; path: string; doDelete: Boolean);
var
  i: integer;
  newPath: string;
  Folder: OleVariant;
begin
  LogMessage('Folder: ' + path + ' (' + IntToStr(Folders.Count) + ')');
  for i := Folders.Count downto 1 do begin
    Folder := Folders.item[i];
    // get the nth folder
    newPath := path + Folder.name + '\';
    LogMessage('Folder: ' + newPath);
    // does it have sub-folders?
    if Folder.Folders.Count < 1 then begin
      // can I delete it (no sub-folders and no mail items)
      if CanDeleteFolder(Folder) then begin
        DeleteFolder(Folder, newPath, doDelete, True);
      end;
    end else begin
      // folder has subfolders, so go ahead and process them...
      ProcessFolders(Folder.Folders, newPath, doDelete);
    end;
    // check again to see if the Folder has subfolders
    // do the the possible deletion during the recursive call to ProcessFolders
    if CanDeleteFolder(Folder) then begin
      // at this point, the path is already in the list, so set AddToList
      // to false (last parameter)
      DeleteFolder(Folder, newPath, doDelete, false);
    end;
  end;
end;

procedure LoadOutlookFolders(doDelete: Boolean);
const
  olFolderInbox = 6;

var
  inbox, nameSpace, outlook: OleVariant;

begin
  LogMessage('Loading Outlook folders');
  if doDelete then begin
    LogMessage('Delete enabled');
  end else begin
    LogMessage('Delete disabled');
  end;

  frmMain.lstFolders.Clear;
  // frmMain.lblFolderCount.caption := 'No Empty Folders';
  folderCounter := -1;
  UpdateFolderCounter();
  Screen.Cursor := crHourGlass;
  outlook := CreateOleObject('Outlook.Application');
  nameSpace := outlook.GetNameSpace('MAPI');
  inbox := nameSpace.GetDefaultFolder(olFolderInbox);
  ProcessFolders(inbox.Folders, 'Inbox\', doDelete);
  outlook := Unassigned;
  Screen.Cursor := crDefault;
  if frmMain.lstFolders.Count < 1 then begin
    MessageDlg('No empty folders in the Outlook Inbox', mtInformation,
      [mbOk], 0, mbOk);
  end else begin
    // enable the Delete button
    SetStartButtonState(True);
  end;
end;

function GetOutlookApplication: OutlookApplication;
var
  ActiveObject: IUnknown;
begin
  if Succeeded(GetActiveObject(OutlookApplication, nil, ActiveObject)) then
    Result := ActiveObject as OutlookApplication
  else
    Result := CoOutlookApplication.Create;
end;

// =====================================================================

procedure TfrmMain.AuthorNoticeClick(Sender: TObject);
begin
  // open the wiki page for the repo
  ShellExecute(self.WindowHandle, 'open', 'https://johnwargo.com', nil, nil,
    SW_SHOWNORMAL);
end;

procedure TfrmMain.btnAnalyzeClick(Sender: TObject);
begin
  LogMessage('Analyze button clicked');
  LoadOutlookFolders(false);
end;

procedure TfrmMain.btnCloseClick(Sender: TObject);
begin
  LogMessage('Closing application');
  CloseLog;
  Application.Terminate();
end;

procedure TfrmMain.btnHelpClick(Sender: TObject);
begin
  // open the wiki page for the repo
  ShellExecute(self.WindowHandle, 'open', PChar(HelpURL), nil, nil,
    SW_SHOWNORMAL);
end;

procedure TfrmMain.btnViewLogClick(Sender: TObject);
begin
  LogMessage('View Log button clicked');
  Launcher.FileName := getLogFilePath;
  Launcher.Execute;
end;

procedure TfrmMain.btnDeleteClick(Sender: TObject);
begin
  LogMessage('Start button clicked');
  if MessageDlg('Are you sure you want to delete all empty Inbox folders? ' +
    'This is not reversable!', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes
  then begin
    LogMessage('User confirmed delete');
    LoadOutlookFolders(True);
  end else begin
    LogMessage('User cancelled delete');
  end;
end;

// =====================================================================

procedure TfrmMain.FormActivate(Sender: TObject);
begin
  VersionInfo.FilePath := Application.ExeName;
  OpenLog;
  LogMessage('Form activated');
  SetStartButtonState(false);
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  LogMessage('Closing application');
  CloseLog;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  RzRegIniFile.Path := 'Software\John Wargo\Outlook Delete Empty Inbox Folders';
  // Hide the existing UI
  Toolbar.Visible := false;
  StatusBar.Visible := false;
  PanelHeader.Visible := false;
  // Make the Splash panel full size
  SplashPanel.Align := alClient;
  SplashImage.Align := alClient;
  // Enable the timer that hides the splash panel
  SplashTimer.Enabled := True;
end;

procedure TfrmMain.SplashTimerTimer(Sender: TObject);
begin
  // Turn off the timer (we only want it to fire once)
  SplashTimer.Enabled := false;
  // Hide the Splash panel
  SplashPanel.Visible := false;
  // Unhide the app UX
  Toolbar.Visible := True;
  StatusBar.Visible := True;
  PanelHeader.Visible := True;
end;

end.


