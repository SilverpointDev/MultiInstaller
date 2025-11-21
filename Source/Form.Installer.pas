unit Form.Installer;

interface

{$BOOLEVAL OFF}       // Unit depends on short-circuit boolean evaluation

{$IFDEF DEBUG}
{$DEFINE SPDEBUGMODE} // Uncomment to debug
{$ENDIF}


uses
  Windows,
  Messages,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  ComCtrls,
  StdCtrls,
  ExtCtrls,
  ActnList,
  CheckLst,
  Actions,
  SpComponentInstaller;

type
  TFormInstall = class(TForm)
    PageControl1: TPageControl;
    tshSelectComponents: TTabSheet;
    tshSelectIde: TTabSheet;
    tshInstallation: TTabSheet;
    pnlTop: TPanel;
    pnlBottom: TPanel;
    btnNext: TButton;
    btnBack: TButton;
    btnCancel: TButton;
    btnFinish: TButton;
    btnSaveLog: TButton;
    btnInstallFolder: TButton;
    lblTitle: TLabel;
    lblSelectComponents: TLabel;
    clbSelectComponents: TCheckListBox;
    chkGetFromGit: TCheckBox;
    edtInstallFolder: TEdit;
    lblInstallfolder: TLabel;
    ActionList1: TActionList;
    aBack: TAction;
    aNext: TAction;
    aCancel: TAction;
    aBrowse: TAction;
    aSaveLog: TAction;
    aFinish: TAction;
    rgSelectIde: TRadioGroup;
    lblInstallation: TLabel;
    lblInstallationFinished: TLabel;
    bvlTop: TBevel;
    memInstallationLog: TMemo;
    bvlBottom: TBevel;
    pbxVersionInfo: TPaintBox;
    imgLogo: TImage;
    Timer1: TTimer;
    SaveDialog1: TSaveDialog;
    chkSelectAllNone: TCheckBox;
    procedure aBrowseExecute(Sender: TObject);
    procedure aBackExecute(Sender: TObject);
    procedure aCancelExecute(Sender: TObject);
    procedure aFinishExecute(Sender: TObject);
    procedure aNextExecute(Sender: TObject);
    procedure aSaveLogExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure clbSelectComponentsDrawItem(Control: TWinControl; Index: Integer;
      Rect: TRect; State: TOwnerDrawState);
    procedure clbSelectComponentsMeasureItem(Control: TWinControl;
      Index: Integer; var Height: Integer);
    procedure clbSelectComponentsClickCheck(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure pbxVersionInfoPaint(Sender: TObject);
    procedure pbxVersionInfoClick(Sender: TObject);
    procedure chkSelectAllNoneClick(Sender: TObject);
  private
    FAppPath  : string;
    FIniPath  : string;
    FAutoStart: Boolean;
    FInstaller: TSpMultiInstaller;

    function ChangePage(Next: Boolean): Boolean;
    function Install: Boolean;
    function ValidateCheckListBox: Boolean;

    procedure CloseDelphi;
    procedure CreateInstaller;
    procedure FillCheckListBox;
    procedure FillRadioGroup;
    procedure ShowNavigationActions(AFinalStep: Boolean);
    procedure WMDROPFILES(var Msg: TWMDropFiles); message WM_DROPFILES;
  end;

var
  FormInstall: TFormInstall;

implementation

{$R *.dfm}


uses
  Winapi.ShellAPI,
  System.StrUtils,
  System.SysUtils,
  System.UITypes;

const
  rvMultiInstallerVersion = 'Silverpoint MultiInstaller 3.6.0';
  rvMultiInstallerLink    = 'http://www.silverpointdevelopment.com';
  rvSetupIni              = 'Setup.Ini';
  crIDC_HAND              = 32649;

resourcestring
  SWelcomeTitle = 'Welcome to the Silverpoint MultiInstaller Setup Wizard';
  SDestinationTitle = 'Select Destination Folder';
  SInstallingTitle = 'Installing...';
  SFinishTitle = 'Completing the MultiInstaller Setup Wizard';

  SCloseDelphi = 'Close Delphi to continue.';
  SErrorLabel = 'There were errors found in the setup, check the log.';
  SErrorInvalidBasePath = 'The directory doesn''t exist.';

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Form UI }

procedure TFormInstall.FormCreate(Sender: TObject);
var
  LAutoStart: string;
  LSetupIni : string;
begin
  Screen.Cursors[crIDC_HAND] := LoadCursor(0, IDC_HAND);
  pbxVersionInfo.Cursor := crIDC_HAND;
  FAppPath := IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName));

  DragAcceptFiles(Handle, True);

  PageControl1.ActivePageIndex := 0;
  lblTitle.Caption := SWelcomeTitle;
  SaveDialog1.InitialDir := FAppPath;

  // Allow to turn the autostart off via command line. Default:on
  if FindCmdLineSwitch('A', LAutoStart) then
    FAutoStart := MatchText(LAutoStart, ['Yes', 'True', '1'])
  else
    FAutoStart := True;

  // Allow to pass the ini file via command line
  if FindCmdLineSwitch('I', LSetupIni) then
    FIniPath := LSetupIni
  else
    FIniPath := FAppPath + rvSetupIni;
  CreateInstaller;

{$IFDEF SPDEBUGMODE}
  ReportMemoryLeaksOnShutdown := True;
{$ENDIF}
end;

procedure TFormInstall.FormDestroy(Sender: TObject);
begin
  FInstaller.Free;
end;

procedure TFormInstall.FormShow(Sender: TObject);
begin
  CloseDelphi;

  if DirectoryExists(FInstaller.ComponentPackages.DefaultInstallFolder) then
    begin
      edtInstallFolder.Text := FInstaller.ComponentPackages.DefaultInstallFolder;
      if FAutoStart then
        begin
          PageControl1.ActivePageIndex := PageControl1.PageCount - 1;
          Timer1.Enabled := True; // Delay it a little for UI responsiveness
        end;
    end;
end;

procedure TFormInstall.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  Install;
end;

function TFormInstall.ChangePage(Next: Boolean): Boolean;
var
  I, C: Integer;
begin
  Result := False;
  I := PageControl1.ActivePageIndex;
  C := PageControl1.PageCount - 1;

  if Next then
    begin
      if I = C then
        Exit
      else
        if I = 1 then
        if not DirectoryExists(edtInstallFolder.Text) then
          begin
            MessageDlg(SErrorInvalidBasePath, mtWarning, [mbOK], 0);
            Exit;
          end;
    end
  else
    if I = 0 then
    Exit;

  Result := True;
  if Next then
    Inc(I)
  else
    Dec(I);
  PageControl1.ActivePageIndex := I;

  btnBack.Enabled := I > 0;
  case I of
    0:
      begin
        lblTitle.Caption := SWelcomeTitle;
        CreateInstaller;
      end;
    1:
      lblTitle.Caption := SDestinationTitle;
    2:
      begin
        lblTitle.Caption := SInstallingTitle;
        Timer1.Enabled := True; // Delay it a little for UI responsiveness
      end;
  else
    lblTitle.Caption := '';
  end;
end;

procedure TFormInstall.chkSelectAllNoneClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to clbSelectComponents.Count - 1 do
    clbSelectComponents.Checked[I] := chkSelectAllNone.Checked;
end;

procedure TFormInstall.CreateInstaller;
var
  I: Integer;
begin
  FreeAndNil(FInstaller);

  FInstaller := TSpMultiInstaller.Create(FIniPath);
  FillCheckListBox;
  FillRadioGroup;
  ValidateCheckListBox;

  // Enable Get from Git only when at least one component has a Git URL
  for I := 0 to FInstaller.ComponentPackages.Count - 1 do
    if not FInstaller.ComponentPackages[I].Git.IsEmpty then
      begin
        chkGetFromGit.Enabled := True;
        chkGetFromGit.Checked := True;
        Break;
      end;

  if DirectoryExists(FInstaller.ComponentPackages.DefaultInstallFolder) then
    edtInstallFolder.Text := FInstaller.ComponentPackages.DefaultInstallFolder;
end;

procedure TFormInstall.FillCheckListBox;
var
  I, G, P: Integer;
begin
  clbSelectComponents.Clear;
  clbSelectComponents.ItemIndex := -1;

  for I := 0 to FInstaller.ComponentPackages.Count - 1 do
    begin
      P := -1;
      G := FInstaller.ComponentPackages[I].GroupIndex;
      if G > 0 then
        begin
          P := clbSelectComponents.Items.IndexOfObject(Pointer(G));
          if P > -1 then
            clbSelectComponents.Items[P] := clbSelectComponents.Items[P] + sLineBreak +
              FInstaller.ComponentPackages[I].Name;
        end;

      if P = -1 then
        begin
          P := clbSelectComponents.Items.AddObject(FInstaller.ComponentPackages[I].Name, Pointer(G));
          clbSelectComponents.Checked[P] := True;
          if FInstaller.ComponentPackages[I].Git <> '' then
            clbSelectComponents.Items[P] := clbSelectComponents.Items[P] + sLineBreak +
              'GIT: ' + FInstaller.ComponentPackages[I].Git;
        end;
    end;
  chkSelectAllNone.Checked := True;
end;

procedure TFormInstall.FillRadioGroup;
var
  IDE: TSpIDEType;
begin
  rgSelectIde.Items.Clear;
  rgSelectIde.ItemIndex := -1;

  for IDE := Low(TSpIDEType) to High(TSpIDEType) do
    if IDE >= FInstaller.ComponentPackages.MinimumIDE then
      if TSpDelphiIDE.Installed(IDE) then
        begin
          rgSelectIde.Items.AddObject(IDETypes[IDE].IDEName, Pointer(Ord(IDE)));
          if IDE = FInstaller.ComponentPackages.DefaultInstallIDE then
            rgSelectIde.ItemIndex := rgSelectIde.Items.Count - 1;
        end;

  if rgSelectIde.ItemIndex = -1 then
    rgSelectIde.ItemIndex := rgSelectIde.Items.Count - 1;
end;

function TFormInstall.ValidateCheckListBox: Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 0 to clbSelectComponents.Count - 1 do
    if clbSelectComponents.Checked[I] then
      begin
        Result := True;
        Break;
      end;

  btnNext.Enabled := Result;
end;

procedure TFormInstall.WMDROPFILES(var Msg: TWMDropFiles);
var
  LDropHandle   : HDROP;
  FileNameLength: Integer;
  LCount        : Integer;
  LFileName     : string;
  LFileExt      : string;
begin
  inherited;

  LDropHandle := Msg.Drop;
  LCount := DragQueryFile(LDropHandle, $FFFFFFFF, nil, 0);

  try
    if LCount = 1 then
      begin
        FileNameLength := DragQueryFile(LDropHandle, 0, nil, 0);
        SetLength(LFileName, FileNameLength);
        DragQueryFile(LDropHandle, 0, PChar(LFileName), FileNameLength + 1);
        LFileExt := ExtractFileExt(LFileName);
        if SameText(LFileExt, '.ini') then
          begin
            FIniPath := LFileName;
            CreateInstaller;
          end;
      end;

  finally
    DragFinish(LDropHandle);
    Msg.Result := 0;
  end;
end;

procedure TFormInstall.clbSelectComponentsClickCheck(Sender: TObject);
begin
  ValidateCheckListBox;
end;

procedure TFormInstall.clbSelectComponentsMeasureItem(Control: TWinControl;
  Index: Integer; var Height: Integer);
var
  R: TRect;
begin
  if Index > -1 then
    Height := DrawText(clbSelectComponents.Canvas.Handle, PChar(clbSelectComponents.Items[Index]), -1, R, DT_CALCRECT) + 4;
end;

procedure TFormInstall.clbSelectComponentsDrawItem(Control: TWinControl;
  Index: Integer; Rect: TRect; State: TOwnerDrawState);
begin
  if Index > -1 then
    begin
      clbSelectComponents.Canvas.FillRect(Rect);
      OffsetRect(Rect, 8, 2);
      DrawText(clbSelectComponents.Canvas.Handle, PChar(clbSelectComponents.Items[Index]), -1, Rect, 0);
    end;
end;

procedure TFormInstall.pbxVersionInfoPaint(Sender: TObject);
var
  C: TCanvas;
begin
  C := pbxVersionInfo.Canvas;
  C.Brush.Style := bsClear;
  C.Font.Color := clBtnHighlight;
  C.TextOut(1, 1, rvMultiInstallerVersion);
  C.Font.Color := clBtnShadow;
  C.TextOut(0, 0, rvMultiInstallerVersion);
end;

procedure TFormInstall.pbxVersionInfoClick(Sender: TObject);
begin
  SpOpenLink(rvMultiInstallerLink);
end;

procedure TFormInstall.ShowNavigationActions(AFinalStep: Boolean);
begin
  aFinish.Visible := AFinalStep;
  aSaveLog.Visible := AFinalStep;
  aNext.Visible := not AFinalStep;
  aCancel.Visible := not AFinalStep;
end;


//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{$REGION 'Actions'}


procedure TFormInstall.aBackExecute(Sender: TObject);
begin
  ChangePage(False);
  ShowNavigationActions(False);
  lblInstallationFinished.Visible := False;
end;

procedure TFormInstall.aNextExecute(Sender: TObject);
begin
  ChangePage(True);
end;

procedure TFormInstall.aCancelExecute(Sender: TObject);
begin
  Close;
end;

procedure TFormInstall.aFinishExecute(Sender: TObject);
begin
  Close;
end;

procedure TFormInstall.aSaveLogExecute(Sender: TObject);
begin
  if SaveDialog1.Execute then
    memInstallationLog.Lines.SaveToFile(SaveDialog1.FileName);
end;

procedure TFormInstall.aBrowseExecute(Sender: TObject);
var
  D: string;
begin
  if SpSelectDirectory('', D) then
    edtInstallFolder.Text := D;
end;

{$ENDREGION 'Actions'}

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM

{ Install }

procedure TFormInstall.CloseDelphi;
var
  Cancel: Boolean;
begin
{$IFDEF SPDEBUGMODE}
  Exit;
{$ENDIF}
  Cancel := False;
  while not Cancel and ((FindWindow('TAppBuilder', nil) <> 0) or (FindWindow('TAppBuilder', nil) <> 0)) do
    Cancel := MessageDlg(SCloseDelphi, mtWarning, [mbOK, mbCancel], 0) = mrCancel;
  if Cancel then
    Close;
end;

function TFormInstall.Install: Boolean;
var
  I, J, G: Integer;
  IDE    : TSpIDEType;
begin
  Result := False;
  CloseDelphi;

  // Get IDE version
  IDE := ideNone;
  I := rgSelectIde.ItemIndex;
  if (I > -1) and Assigned(rgSelectIde.Items.Objects[I]) then
    IDE := TSpIDEType(rgSelectIde.Items.Objects[I]);

  // Delete unchecked components from the ComponentPackages list
  for I := 0 to clbSelectComponents.Count - 1 do
    if not clbSelectComponents.Checked[I] then
      begin
        G := Integer(clbSelectComponents.Items.Objects[I]);
        for J := FInstaller.ComponentPackages.Count - 1 downto 0 do
          if (G > 0) and (FInstaller.ComponentPackages[J].GroupIndex = G) then
            FInstaller.ComponentPackages.Delete(J)
          else
            if clbSelectComponents.Items[I].Contains(FInstaller.ComponentPackages[J].Name) then
            FInstaller.ComponentPackages.Delete(J);
      end;

  // Prioritize GIT over ZIP
  if chkGetFromGit.Checked then
    begin
      for J := 0 to FInstaller.ComponentPackages.Count - 1 do
        if not FInstaller.ComponentPackages[J].Git.IsEmpty then
          FInstaller.ComponentPackages[J].ZipFile := '';
    end;

  ShowNavigationActions(True);
  aBack.Enabled := False;
  aFinish.Enabled := False;
  aSaveLog.Enabled := False;
  lblInstallationFinished.Visible := False;

  Application.ProcessMessages;
  try
    // Check, Unzip, Patch, Compile, Install
    if FInstaller.Install(FAppPath, edtInstallFolder.Text, IDE, memInstallationLog.Lines) then
      Result := True;
  finally
    lblTitle.Caption := SFinishTitle;
    aBack.Enabled := True;
    aFinish.Enabled := True;
    aSaveLog.Enabled := True;
    lblInstallationFinished.Visible := True;
    if not Result then
      begin
        lblInstallationFinished.Font.Color := clRed;
        lblInstallationFinished.Caption := SErrorLabel;
      end;
    lblInstallationFinished.Visible := True;
  end;
end;

end.
