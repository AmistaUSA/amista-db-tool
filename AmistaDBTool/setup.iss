[Setup]
AppName=AmistaDBTool
AppVersion=1.0.0
DefaultDirName={autopf}\AmistaDBTool
DefaultGroupName=AmistaDBTool
OutputDir=Output
OutputBaseFilename=AmistaDBToolSetup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin

[Files]
Source: "publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Files for testing connection during install (extract all needed files)
Source: "publish\AmistaDBTool.exe"; DestDir: "{tmp}"; Flags: dontcopy
Source: "publish\AmistaDBTool.dll"; DestDir: "{tmp}"; Flags: dontcopy
Source: "publish\AmistaDBTool.runtimeconfig.json"; DestDir: "{tmp}"; Flags: dontcopy
Source: "publish\AmistaDBTool.deps.json"; DestDir: "{tmp}"; Flags: dontcopy
Source: "publish\SAPBusinessOneSDK.dll"; DestDir: "{tmp}"; Flags: dontcopy
Source: "publish\*.dll"; DestDir: "{tmp}"; Flags: dontcopy; Excludes: "AmistaDBTool.dll,SAPBusinessOneSDK.dll"

[Code]
var
  // DB Connection Page controls
  DBPage: TWizardPage;
  edtServer: TNewEdit;
  cmbDbServerType: TNewComboBox;
  edtDBUserName: TNewEdit;
  edtDBPassword: TPasswordEdit;
  edtCompanyDB: TNewEdit;

  // SAP Connection Page
  SAPPage: TInputQueryWizardPage;

procedure InitializeWizard;
var
  lblServer, lblDbType, lblDBUser, lblDBPass, lblCompanyDB: TNewStaticText;
  TopPos: Integer;
begin
  // Page 1: Database Connection Credentials (Custom page with combo box)
  DBPage := CreateCustomPage(wpWelcome,
    'DB Connection Credentials',
    'Enter database connection details. These settings are used to connect to the SAP Business One database.');

  TopPos := 8;

  // Server
  lblServer := TNewStaticText.Create(DBPage);
  lblServer.Parent := DBPage.Surface;
  lblServer.Caption := 'Server:';
  lblServer.Left := 0;
  lblServer.Top := TopPos;

  edtServer := TNewEdit.Create(DBPage);
  edtServer.Parent := DBPage.Surface;
  edtServer.Left := 0;
  edtServer.Top := TopPos + 16;
  edtServer.Width := DBPage.SurfaceWidth;

  TopPos := TopPos + 48;

  // DB Server Type (Dropdown)
  lblDbType := TNewStaticText.Create(DBPage);
  lblDbType.Parent := DBPage.Surface;
  lblDbType.Caption := 'DB Server Type:';
  lblDbType.Left := 0;
  lblDbType.Top := TopPos;

  cmbDbServerType := TNewComboBox.Create(DBPage);
  cmbDbServerType.Parent := DBPage.Surface;
  cmbDbServerType.Left := 0;
  cmbDbServerType.Top := TopPos + 16;
  cmbDbServerType.Width := DBPage.SurfaceWidth;
  cmbDbServerType.Style := csDropDownList;
  cmbDbServerType.Items.Add('dst_MSSQL');
  cmbDbServerType.Items.Add('dst_DB_2');
  cmbDbServerType.Items.Add('dst_SYBASE');
  cmbDbServerType.Items.Add('dst_MSSQL2005');
  cmbDbServerType.Items.Add('dst_MAXDB');
  cmbDbServerType.Items.Add('dst_MSSQL2008');
  cmbDbServerType.Items.Add('dst_MSSQL2012');
  cmbDbServerType.Items.Add('dst_MSSQL2014');
  cmbDbServerType.Items.Add('dst_HANADB');
  cmbDbServerType.Items.Add('dst_MSSQL2016');
  cmbDbServerType.Items.Add('dst_MSSQL2017');
  cmbDbServerType.Items.Add('dst_MSSQL2019');
  cmbDbServerType.ItemIndex := 8; // Default to dst_HANADB

  TopPos := TopPos + 48;

  // DB User Name
  lblDBUser := TNewStaticText.Create(DBPage);
  lblDBUser.Parent := DBPage.Surface;
  lblDBUser.Caption := 'DB User Name:';
  lblDBUser.Left := 0;
  lblDBUser.Top := TopPos;

  edtDBUserName := TNewEdit.Create(DBPage);
  edtDBUserName.Parent := DBPage.Surface;
  edtDBUserName.Left := 0;
  edtDBUserName.Top := TopPos + 16;
  edtDBUserName.Width := DBPage.SurfaceWidth;

  TopPos := TopPos + 48;

  // DB Password
  lblDBPass := TNewStaticText.Create(DBPage);
  lblDBPass.Parent := DBPage.Surface;
  lblDBPass.Caption := 'DB Password:';
  lblDBPass.Left := 0;
  lblDBPass.Top := TopPos;

  edtDBPassword := TPasswordEdit.Create(DBPage);
  edtDBPassword.Parent := DBPage.Surface;
  edtDBPassword.Left := 0;
  edtDBPassword.Top := TopPos + 16;
  edtDBPassword.Width := DBPage.SurfaceWidth;

  TopPos := TopPos + 48;

  // Company DB
  lblCompanyDB := TNewStaticText.Create(DBPage);
  lblCompanyDB.Parent := DBPage.Surface;
  lblCompanyDB.Caption := 'Company DB:';
  lblCompanyDB.Left := 0;
  lblCompanyDB.Top := TopPos;

  edtCompanyDB := TNewEdit.Create(DBPage);
  edtCompanyDB.Parent := DBPage.Surface;
  edtCompanyDB.Left := 0;
  edtCompanyDB.Top := TopPos + 16;
  edtCompanyDB.Width := DBPage.SurfaceWidth;

  // Page 2: SAP Connection Credentials (appears after DBPage)
  SAPPage := CreateInputQueryPage(DBPage.ID,
    'SAP Connection Credentials', 'Enter SAP Business One user credentials',
    'These settings are used to authenticate with SAP Business One.');

  SAPPage.Add('SAP User Name:', False);
  SAPPage.Add('SAP Password:', True);
  SAPPage.Add('License Server:', False);
  SAPPage.Add('SLD Server:', False);
end;

function EscapeJsonString(const S: String): String;
var
  I: Integer;
  C: Char;
begin
  Result := '';
  for I := 1 to Length(S) do
  begin
    C := S[I];
    if C = '"' then
      Result := Result + '\"'
    else if C = '\' then
      Result := Result + '\\'
    else if C = #8 then
      Result := Result + '\b'
    else if C = #9 then
      Result := Result + '\t'
    else if C = #10 then
      Result := Result + '\n'
    else if C = #12 then
      Result := Result + '\f'
    else if C = #13 then
      Result := Result + '\r'
    else
      Result := Result + C;
  end;
end;

function BuildJsonConfig: String;
var
  DbServerType: String;
begin
  if cmbDbServerType.ItemIndex >= 0 then
    DbServerType := cmbDbServerType.Items[cmbDbServerType.ItemIndex]
  else
    DbServerType := 'dst_HANADB';

  Result := '{' + #13#10 +
    '  "SapConnection": {' + #13#10 +
    '    "Server": "' + EscapeJsonString(edtServer.Text) + '",' + #13#10 +
    '    "DbServerType": "' + DbServerType + '",' + #13#10 +
    '    "DBUserName": "' + EscapeJsonString(edtDBUserName.Text) + '",' + #13#10 +
    '    "DBPassword": "' + EscapeJsonString(edtDBPassword.Text) + '",' + #13#10 +
    '    "CompanyDB": "' + EscapeJsonString(edtCompanyDB.Text) + '",' + #13#10 +
    '    "UserName": "' + EscapeJsonString(SAPPage.Values[0]) + '",' + #13#10 +
    '    "Password": "' + EscapeJsonString(SAPPage.Values[1]) + '",' + #13#10 +
    '    "LicenseServer": "' + EscapeJsonString(SAPPage.Values[2]) + '",' + #13#10 +
    '    "SLDServer": "' + EscapeJsonString(SAPPage.Values[3]) + '"' + #13#10 +
    '  },' + #13#10 +
    '  "Logging": { "LogLevel": { "Default": "Information", "Microsoft.Hosting.Lifetime": "Information" } }' + #13#10 +
    '}';
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  JsonContent: String;
  TempConfigPath: String;
  ExePath: String;
  ResultCode: Integer;
begin
  Result := True;

  // Validate DB Connection page
  if CurPageID = DBPage.ID then
  begin
    if (edtServer.Text = '') or (edtDBUserName.Text = '') or (edtCompanyDB.Text = '') then
    begin
      MsgBox('Please fill in all required fields (Server, DB User, Company DB).', mbError, MB_OK);
      Result := False;
      Exit;
    end;
  end;

  // Test connection after SAP Credentials page
  if CurPageID = SAPPage.ID then
  begin
    // Build JSON config
    JsonContent := BuildJsonConfig;

    TempConfigPath := ExpandConstant('{tmp}\appsettings.json');
    SaveStringToFile(TempConfigPath, JsonContent, False);

    // Extract all required files for connection test
    try
      ExtractTemporaryFiles('*.dll');
      ExtractTemporaryFile('AmistaDBTool.exe');
      ExtractTemporaryFile('AmistaDBTool.runtimeconfig.json');
      ExtractTemporaryFile('AmistaDBTool.deps.json');
    except
      MsgBox('Failed to extract helper files. ' + GetExceptionMessage, mbError, MB_OK);
      Result := False;
      Exit;
    end;

    ExePath := ExpandConstant('{tmp}\AmistaDBTool.exe');

    // Test the connection (use SW_SHOWNORMAL to see output for debugging)
    if Exec(ExePath, '--test-connection "' + TempConfigPath + '"', ExpandConstant('{tmp}'), SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode) then
    begin
      if ResultCode = 0 then
      begin
        MsgBox('Connection Successful! Proceeding with installation.', mbInformation, MB_OK);
      end
      else
      begin
        if MsgBox('Connection Failed (Exit code: ' + IntToStr(ResultCode) + ').' + #13#10 + #13#10 +
                  'Config path: ' + TempConfigPath + #13#10 + #13#10 +
                  'Do you want to try again? Click No to proceed anyway.', mbConfirmation, MB_YESNO) = IDYES then
        begin
          Result := False; // Stay on page
        end;
      end;
    end
    else
    begin
      MsgBox('Failed to run connection test tool.' + #13#10 +
             'Exe path: ' + ExePath + #13#10 +
             'Config path: ' + TempConfigPath, mbError, MB_OK);
      Result := False;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  AppConfigPath: String;
  JsonContent: String;
begin
  if CurStep = ssPostInstall then
  begin
    // Write config to installed location
    JsonContent := BuildJsonConfig;
    AppConfigPath := ExpandConstant('{app}\appsettings.json');
    SaveStringToFile(AppConfigPath, JsonContent, False);
  end;
end;
