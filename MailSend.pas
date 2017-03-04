unit MailSend;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, sPageControl, sSpinEdit, StdCtrls, sEdit,
  IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase,
  IdMessageClient, IdSMTPBase, IdSMTP, IdComponent, IdIOHandler,
  IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL,
  IdBaseComponent, IdMessage, sMemo, Buttons, sSpeedButton, sListBox,
  sGroupBox, sComboBox, IdAttachmentFile, IdCoderHeader, sListView,
  ShellApi, acAlphaImageList, ImgList, DB, IBC, IdText, Menus, sRichEdit,
  acPNG, ExtCtrls, sFontCtrls, sPanel, sComboBoxes, NxEdit, sLabel, sGauge,
  IniFiles, NxScrollControl, NxCustomGridControl, NxCustomGrid, NxGrid,
  EncdDecd, MapiEmail;

type
  TFormMailSender = class(TForm)
    sPageControl: TsPageControl;
    tabSendMail: TsTabSheet;
    tabParams: TsTabSheet;
    IdMessage: TIdMessage;
    IdSSLIOHandlerSocketOpenSSL: TIdSSLIOHandlerSocketOpenSSL;
    IdSMTP: TIdSMTP;
    editSubj: TsEdit;
    btnAttachAdd: TsSpeedButton;
    btnSend: TsSpeedButton;
    btnEmailAdd: TsSpeedButton;
    btnEmailDelete: TsSpeedButton;
    btnEmailClear: TsSpeedButton;
    editTo: TsEdit;
    sGroupBox: TsGroupBox;
    editHost: TsEdit;
    editPort: TsSpinEdit;
    editDomen: TsEdit;
    editLogin: TsEdit;
    editPassword: TsEdit;
    editProfile: TsComboBox;
    btnProfileCreate: TsSpeedButton;
    btnProfileDelete: TsSpeedButton;
    btnProfileEdit: TsSpeedButton;
    editFrom: TsComboBox;
    AttachmentDialog: TOpenDialog;
    btnLog: TsSpeedButton;
    editUName: TsEdit;
    editReplyTo: TsEdit;
    editAttach: TsListView;
    AlphaImageList: TsAlphaImageList;
    btnAttachDelete: TsSpeedButton;
    popupPriority: TPopupMenu;
    nPriorityHigh: TMenuItem;
    nPriorityMed: TMenuItem;
    nPriorityLow: TMenuItem;
    btnPriority: TsSpeedButton;
    editSignature: TsMemo;
    btnCancel: TsSpeedButton;
    Gauge: TsGauge;
    lblGauge: TsLabel;
    editEmailList: TsListBox;
    editMessage: TsMemo;
    editSSLMethod: TsComboBox;
    editTLSMethod: TsComboBox;
    procedure ClearEdits(ClearList: Boolean);
    procedure LoadProfiles;
    procedure IniLoadMailSender;
    procedure btnProfileCreateClick(Sender: TObject);
    procedure editProfileChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnProfileEditClick(Sender: TObject);
    procedure btnProfileDeleteClick(Sender: TObject);
    procedure btnAttachAddClick(Sender: TObject);
    procedure btnEmailDeleteClick(Sender: TObject);
    procedure btnEmailClearClick(Sender: TObject);
    procedure btnLogClick(Sender: TObject);
    procedure IdSMTPStatus(ASender: TObject; const AStatus: TIdStatus; const AStatusText: string);
    procedure IdMessageInitializeISO(var VHeaderEncoding: Char; var VCharSet: string);
    procedure LV_InsertFiles(nFile: string; ListView: TsListView; ImageList: TsAlphaImageList);
    function IsValidEmail(const Value: string): Boolean;
    function IsValidWeb(const Value: string): Boolean;
    procedure GetEmailList(param: integer; Query: TIBCQuery; ID: string; List: TsListBox; ClearList: Boolean);
    procedure GetFirmList(param: integer; Query: TIBCQuery; ID: string; List: TsListBox);
    procedure btnAttachDeleteClick(Sender: TObject);
    procedure nPriorityHighClick(Sender: TObject);
    procedure btnPriorityClick(Sender: TObject);
    procedure editToKeyPress(Sender: TObject; var Key: Char);
    procedure editEmailListKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btnEmailAddClick(Sender: TObject);
    procedure IdSMTPFailedRecipient(Sender: TObject; const AAddress, ACode, AText: string; var VContinue: Boolean);
    procedure btnCancelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SendEmail(Sender: TObject);
    procedure SendRegInfoCheck(Sender: TObject);

  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams); override;
    { Public declarations }
  end;

var
  FormMailSender: TFormMailSender;
  memoLog: TsMemo;
  REFirmInfo: TsRichEdit;
  MailSend_Break: Boolean = False;
  editEmailList_tmp: TsComboBox;

implementation

uses Main, Editor, Helpers;

{$R *.dfm}

procedure TFormMailSender.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.ExStyle := Params.ExStyle or WS_Ex_AppWindow;
end;

procedure TFormMailSender.FormCreate(Sender: TObject);
begin
  LoadProfiles;
  IniLoadMailSender;
  memoLog := TsMemo.Create(tabSendMail);
  memoLog.Parent := tabSendMail;
  memoLog.Left := 6;
  memoLog.Top := 220;
  memoLog.Width := 532;
  memoLog.Height := 225;
  memoLog.Visible := False;
  memoLog.Clear;
  memoLog.ScrollBars := ssVertical;
  memoLog.ReadOnly := True;
  editEmailList_tmp := TsComboBox.Create(FormMailSender);
  editEmailList_tmp.Parent := FormMailSender;
  editEmailList_tmp.Visible := False;
  REFirmInfo := TsRichEdit.Create(FormMailSender);
  REFirmInfo.Parent := FormMailSender;
  REFirmInfo.Visible := False;
  REFirmInfo.Clear;
  REFirmInfo.WordWrap := False;
end;

procedure TFormMailSender.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  MailSend_Break := True;
end;

procedure TFormMailSender.IniLoadMailSender;
begin
  IniMain := TIniFile.Create(AppPath + 'office.ini');
  if IniMain.ReadInteger('mailsender', 'fromIndex', 0) <= editFrom.Items.Count - 1 then
    editFrom.ItemIndex := IniMain.ReadInteger('mailsender', 'fromIndex', 0)
  else if editFrom.Items.Count > 0 then
    editFrom.ItemIndex := 0;
  IniMain.Free;
end;

procedure TFormMailSender.ClearEdits(ClearList: Boolean);
begin
  if FormMailSender.Caption = 'Рассылка почты' then
  begin
    editTo.Enabled := True;
    btnEmailAdd.Enabled := True;
    btnEmailDelete.Enabled := True;
    btnSend.OnClick := SendEmail;
  end
  else
  begin
    editTo.Enabled := False;
    btnEmailAdd.Enabled := False;
    btnEmailDelete.Enabled := False;
    btnSend.OnClick := SendRegInfoCheck;
  end;
  sPageControl.ActivePageIndex := 0;
  editSubj.Clear;
  editAttach.Clear;
  editMessage.Clear;
  memoLog.Clear;
  editTo.Clear;
  if ClearList then
    editEmailList.Clear;
  memoLog.Visible := False;
  btnLog.Caption := 'Показать лог';
  btnPriority.ImageIndex := 6;
  nPriorityHigh.Checked := False;
  nPriorityMed.Checked := True;
  nPriorityLow.Checked := False;
  editEmailList_tmp.Clear;
end;

procedure TFormMailSender.LoadProfiles;
var
  i: integer;
begin
  editProfile.ItemIndex := -1;
  editProfile.Clear;
  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text := 'select * from ACCOUNTS';
  FormMain.IBQuery1.Open;
  FormMain.IBQuery1.FetchAll := True;
  if FormMain.IBQuery1.RecordCount > 0 then
    for i := 0 to FormMain.IBQuery1.RecordCount - 1 do
    begin
      editProfile.Items.Add(FormMain.IBQuery1.FieldValues['NAME']);
      editFrom.Items.Add(FormMain.IBQuery1.FieldValues['DOMEN'] + ' (' + FormMain.IBQuery1.FieldValues['NAME'] + ')');
      FormMain.IBQuery1.Next;
    end;
  if editProfile.Items.Count > 0 then
    editProfile.ItemIndex := 0;
  if editProfile.Items.Count > 0 then
    editProfileChange(editProfile);
  FormMain.IBQuery1.Close;
end;

procedure TFormMailSender.LV_InsertFiles(nFile: string; ListView: TsListView; ImageList: TsAlphaImageList);

  function GetFileSize(FileName: string): Longint;
  var
    SearchRec: TSearchRec;
  begin
    if FindFirst(FileName, faAnyFile, SearchRec) = 0 then
      Result := SearchRec.Size
    else
      Result := -1;
  end;

  function FormatByteSize(const bytes: Longint): string;
  const
    B = 1;
    KB = 1024 * B;
    MB = 1024 * KB;
    GB = 1024 * MB;
  begin
    if bytes = 0 then
    begin
      Result := '0 байт';
      exit;
    end;
    if bytes > GB then
      Result := FormatFloat('#.## ГБ', bytes / GB)
    else if bytes > MB then
      Result := FormatFloat('#.## МБ', bytes / MB)
    else if bytes > KB then
      Result := FormatFloat('#.## КБ', bytes / KB)
    else
      Result := FormatFloat('#.## байт', bytes);
  end;

var
  i: integer;
  Icon: TIcon;
  ListItem: TListItem;
  FileInfo: SHFILEINFO;
  capt, fPath, fSize: string;
begin
  SHGetFileInfo(PChar(nFile), 0, FileInfo, SizeOf(FileInfo), SHGFI_DISPLAYNAME);
  fSize := FormatByteSize(GetFileSize(nFile));
  fPath := ExtractFilePath(nFile);
  capt := ExtractFileName(nFile) + ' (' + fSize + ')';
  for i := 0 to ListView.Items.Count - 1 do
  begin
    if (AnsiLowerCase(ListView.Items[i].Caption) = AnsiLowerCase(capt)) and
      (AnsiLowerCase(ListView.Items[i].SubItems[0]) = AnsiLowerCase(fPath)) then
    begin
      MessageBox(handle, 'Этот файл уже находится в списке', 'Информация', MB_OK or MB_ICONINFORMATION);
      exit;
    end;
  end;
  Icon := TIcon.Create;
  ListView.Items.BeginUpdate;
  try
    ListItem := ListView.Items.Add;
    ListItem.Caption := capt;
    ListItem.SubItems.Add(fPath);
    SHGetFileInfo(PChar(nFile), 0, FileInfo, SizeOf(FileInfo), SHGFI_ICON or SHGFI_SMALLICON);
    Icon.handle := FileInfo.hIcon;
    ListItem.ImageIndex := ImageList.AddIcon(Icon);
    DestroyIcon(FileInfo.hIcon);
  finally
    Icon.Free;
    ListView.Items.EndUpdate;
  end;
end;

procedure TFormMailSender.IdSMTPStatus(ASender: TObject; const AStatus: TIdStatus; const AStatusText: string);
begin
  memoLog.Lines.Insert(0, '[' + TimeToStr(Now) + '] Status: ' + AStatusText);
end;

procedure TFormMailSender.IdSMTPFailedRecipient(Sender: TObject; const AAddress, ACode, AText: string; var VContinue: Boolean);
begin
  { i := editEmailList.Items.IndexOf(AAddress);
    if i <> -1 then editEmailList.Items.Delete(i); }
  memoLog.Lines.Insert(0, '[' + TimeToStr(Now) + '] ERROR: ' + AText);
  VContinue := True;
end;

procedure TFormMailSender.IdMessageInitializeISO(var VHeaderEncoding: Char; var VCharSet: string);
begin
  VHeaderEncoding := 'B';
  VCharSet := 'Windows-1251';
end;

procedure TFormMailSender.editProfileChange(Sender: TObject);
begin
  if editProfile.ItemIndex < 0 then
    exit;
  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text := 'select * from ACCOUNTS where NAME = :NAME';
  FormMain.IBQuery1.ParamByName('NAME').AsString := editProfile.Text;
  FormMain.IBQuery1.Open;
  if FormMain.IBQuery1.RecordCount > 0 then
  begin
    editHost.Text := FormMain.IBQuery1.FieldValues['HOST'];
    editPort.Text := FormMain.IBQuery1.FieldValues['PORT'];
    editUName.Text := FormMain.IBQuery1.FieldValues['UNAME'];
    editReplyTo.Text := FormMain.IBQuery1.FieldValues['REPLYTO'];
    editDomen.Text := FormMain.IBQuery1.FieldValues['DOMEN'];
    editLogin.Text := FormMain.IBQuery1.FieldValues['LOGIN'];
    editPassword.Text := DecodeString(FormMain.IBQuery1.FieldValues['PASS']);
    editSignature.Text := FormMain.IBQuery1.FieldValues['SIGNATURE'];
    editSSLMethod.ItemIndex := FormMain.IBQuery1.FieldValues['SSL_METHOD'];
    editTLSMethod.ItemIndex := FormMain.IBQuery1.FieldValues['TLS_METHOD'];
  end;
  FormMain.IBQuery1.Close;
end;

procedure TFormMailSender.btnProfileCreateClick(Sender: TObject);
var
  profile, host, port, uname, replyto, domen, login, pass: string;
  pass_encoded: string;
begin
  if InputQuery('Профиль', 'Название:', profile) then
    if Trim(profile) <> '' then
    begin
      if editProfile.Items.IndexOf(Trim(profile)) <> -1 then
      begin
        MessageBox(handle, 'Профиль с таким именем уже существует', 'Предупреждение', MB_OK or MB_ICONWARNING);
        exit;
      end;
      if not InputQuery(profile, 'Сервер отправки почти (SMTP server):', host) then
        exit;
      if Trim(host) = '' then
        exit;
      if not InputQuery(profile, 'Порт:', port) then
        exit;
      if Trim(port) = '' then
        exit;
      InputQuery(profile, 'Имя отправителя (не обязательно):', uname);
      InputQuery(profile, 'Обратный адрес (не обязательно):', replyto);
      if not InputQuery(profile, 'Домен (адрес электронной почты):', domen) then
        exit;
      if Trim(domen) = '' then
        exit;
      if not InputQuery(profile, 'Логин:', login) then
        exit;
      if Trim(login) = '' then
        exit;
      if not InputQuery(profile, 'Пароль:', pass) then
        exit;
      if Trim(pass) = '' then
        exit;

      FormMain.IBQuery1.Close;
      FormMain.IBQuery1.SQL.Text :=
        'insert into ACCOUNTS (NAME,HOST,PORT,UNAME,REPLYTO,DOMEN,LOGIN,PASS,SIGNATURE,SSL_METHOD,TLS_METHOD) values' +
        '(:NAME,:HOST,:PORT,:UNAME,:REPLYTO,:DOMEN,:LOGIN,:PASS,:SIGNATURE,:SSL_METHOD,:TLS_METHOD)';
      FormMain.IBQuery1.ParamByName('NAME').AsString := Trim(profile);
      FormMain.IBQuery1.ParamByName('HOST').AsString := Trim(host);
      FormMain.IBQuery1.ParamByName('PORT').AsString := Trim(port);
      FormMain.IBQuery1.ParamByName('UNAME').AsString := Trim(uname);
      FormMain.IBQuery1.ParamByName('REPLYTO').AsString := Trim(replyto);
      FormMain.IBQuery1.ParamByName('DOMEN').AsString := Trim(domen);
      FormMain.IBQuery1.ParamByName('LOGIN').AsString := Trim(login);
      pass := Trim(pass);
      pass_encoded := EncodeString(pass);
      FormMain.IBQuery1.ParamByName('PASS').AsString := pass_encoded;
      FormMain.IBQuery1.ParamByName('SIGNATURE').AsString := '';
      FormMain.IBQuery1.ParamByName('SSL_METHOD').AsInteger := 0;
      if Trim(port) = '465' then
        FormMain.IBQuery1.ParamByName('TLS_METHOD').AsInteger := 1
      else
        FormMain.IBQuery1.ParamByName('TLS_METHOD').AsInteger := 0;
      try
        FormMain.IBQuery1.Execute;
      except
        on E: Exception do
        begin
          WriteLog('TFormMailSender.btnProfileCreateClick' + #13 + 'Ошибка: ' + E.Message);
          MessageBox(handle, PChar('Ошибка при создании профиля.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
          exit;
        end;
      end;
      FormMain.IBTransaction1.CommitRetaining;

      editProfile.Items.Add(profile);
      editFrom.Items.Add(domen + ' (' + profile + ')');
      editProfile.ItemIndex := editProfile.Items.Count - 1;
      editHost.Text := host;
      editPort.Text := port;
      editReplyTo.Text := replyto;
      editUName.Text := uname;
      editDomen.Text := domen;
      editLogin.Text := login;
      editPassword.Text := pass;
      editSignature.Clear;
      editSSLMethod.ItemIndex := 0;
      editTLSMethod.ItemIndex := integer(Trim(port) = '465');
    end;
end;

procedure TFormMailSender.btnProfileEditClick(Sender: TObject);
var
  pass, pass_encoded: string;
begin
  if editProfile.ItemIndex < 0 then
    exit;
  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text := 'update ACCOUNTS set HOST=:HOST,PORT=:PORT,' +
    'UNAME=:UNAME,REPLYTO=:REPLYTO,DOMEN=:DOMEN,LOGIN=:LOGIN,PASS=:PASS,SIGNATURE=:SIGNATURE,SSL_METHOD=:SSL_METHOD,TLS_METHOD=:TLS_METHOD where NAME = :NAME';
  FormMain.IBQuery1.ParamByName('NAME').AsString := Trim(editProfile.Text);
  FormMain.IBQuery1.ParamByName('HOST').AsString := Trim(editHost.Text);
  FormMain.IBQuery1.ParamByName('PORT').AsString := Trim(editPort.Text);
  FormMain.IBQuery1.ParamByName('UNAME').AsString := Trim(editUName.Text);
  FormMain.IBQuery1.ParamByName('REPLYTO').AsString := Trim(editReplyTo.Text);
  FormMain.IBQuery1.ParamByName('DOMEN').AsString := Trim(editDomen.Text);
  FormMain.IBQuery1.ParamByName('LOGIN').AsString := Trim(editLogin.Text);
  pass := Trim(editPassword.Text);
  pass_encoded := EncodeString(pass);
  FormMain.IBQuery1.ParamByName('PASS').AsString := pass_encoded;
  FormMain.IBQuery1.ParamByName('SIGNATURE').AsMemo := Trim(editSignature.Text);
  FormMain.IBQuery1.ParamByName('SSL_METHOD').AsInteger := editSSLMethod.ItemIndex;
  FormMain.IBQuery1.ParamByName('TLS_METHOD').AsInteger := editTLSMethod.ItemIndex;
  try
    FormMain.IBQuery1.Execute;
  except
    on E: Exception do
    begin
      WriteLog('TFormMailSender.btnProfileEditClick' + #13 + 'Ошибка: ' + E.Message);
      MessageBox(handle, PChar('Ошибка при редактировании профиля.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      exit;
    end;
  end;
  FormMain.IBTransaction1.CommitRetaining;
  MessageBox(handle, 'Профиль сохранен', 'Информация', MB_OK or MB_ICONINFORMATION);
end;

procedure TFormMailSender.btnProfileDeleteClick(Sender: TObject);
var
  i: integer;
begin
  if editProfile.ItemIndex < 0 then
    exit;
  if MessageBox(handle, 'Профиль будет удален. Продолжить?', 'Подтверждение', MB_YESNO or MB_ICONQUESTION) = MRNO then
    exit;
  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text := 'delete from ACCOUNTS where NAME = :NAME';
  FormMain.IBQuery1.ParamByName('NAME').AsString := editProfile.Text;
  try
    FormMain.IBQuery1.Execute;
  except
    on E: Exception do
    begin
      WriteLog('TFormMailSender.btnProfileDeleteClick' + #13 + 'Ошибка: ' + E.Message);
      MessageBox(handle, PChar('Ошибка при удалении профиля.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      exit;
    end;
  end;
  FormMain.IBTransaction1.CommitRetaining;

  for i := 0 to editFrom.Items.Count - 1 do
    if Pos('(' + editProfile.Text + ')', editFrom.Items[i]) > 0 then
      editFrom.Items.Delete(i);
  editProfile.DeleteSelected;
  if editProfile.Items.Count > 0 then
  begin
    editProfile.ItemIndex := 0;
    editProfileChange(editProfile);
  end
  else
  begin
    editProfile.Clear;
    editHost.Clear;
    editPort.Clear;
    editUName.Clear;
    editReplyTo.Clear;
    editDomen.Clear;
    editLogin.Clear;
    editPassword.Clear;
    editSignature.Clear;
    editSSLMethod.ItemIndex := -1;
    editTLSMethod.ItemIndex := -1;
  end;
end;

procedure TFormMailSender.btnAttachAddClick(Sender: TObject);
begin
  if AttachmentDialog.Execute then
    LV_InsertFiles(AttachmentDialog.FileName, editAttach, AlphaImageList);
end;

procedure TFormMailSender.btnAttachDeleteClick(Sender: TObject);
begin
  if editAttach.SelCount = 0 then
    exit;
  editAttach.Selected.Delete;
  if editAttach.Items.Count > 0 then
    editAttach.Items[0].Selected := True;
end;

procedure TFormMailSender.editToKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  begin
    Key := #0;
    btnEmailAddClick(btnEmailAdd);
  end;
end;

procedure TFormMailSender.editEmailListKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if FormMailSender.Caption = 'Рассылка почты' then
    if Key = 46 then
    begin
      Key := 0;
      btnEmailDeleteClick(btnEmailDelete);
    end;
end;

procedure TFormMailSender.btnEmailAddClick(Sender: TObject);
var
  i: integer;
begin
  if Trim(editTo.Text) = '' then
    exit;
  if not IsValidEmail(editTo.Text) then
  begin
    MessageBox(handle, 'Неверно указан адрес получателя', 'Информация', MB_OK or MB_ICONINFORMATION);
    exit;
  end;
  for i := 0 to editEmailList.Items.Count - 1 do
    if editEmailList.Items[i] = Trim(editTo.Text) then
      exit;
  editEmailList.Items.Add(Trim(editTo.Text));
  editEmailList.MultiSelect := False;
  editEmailList.ItemIndex := editEmailList.Items.Count - 1;
  editEmailList.Selected[editEmailList.Items.Count - 1] := True;
  editEmailList.MultiSelect := False;
  editTo.Clear;
end;

procedure TFormMailSender.btnEmailDeleteClick(Sender: TObject);
var
  i: integer;
begin
  if editEmailList.ItemIndex = -1 then
    exit;
  i := editEmailList.ItemIndex;
  editEmailList.DeleteSelected;
  if (i > -1) and (i <= editEmailList.Items.Count - 1) then
    editEmailList.Selected[i] := True;
  if (editEmailList.ItemIndex = -1) and (editEmailList.Items.Count > 0) then
    editEmailList.ItemIndex := editEmailList.Items.Count - 1;
end;

procedure TFormMailSender.btnEmailClearClick(Sender: TObject);
begin
  if editEmailList.Items.Count = 0 then
    exit;
  if MessageBox(handle, 'Очистить список адресов?', 'Подтверждение', MB_YESNO or MB_ICONQUESTION) = MRYES then
  begin
    editEmailList.Clear;
    editEmailList_tmp.Clear;
  end;
end;

procedure TFormMailSender.nPriorityHighClick(Sender: TObject);
begin
  if TMenuItem(Sender).Name = 'nPriorityHigh' then
  begin
    btnPriority.ImageIndex := 5;
    nPriorityHigh.Checked := True;
    nPriorityMed.Checked := False;
    nPriorityLow.Checked := False;
  end;
  if TMenuItem(Sender).Name = 'nPriorityMed' then
  begin
    btnPriority.ImageIndex := 6;
    nPriorityHigh.Checked := False;
    nPriorityMed.Checked := True;
    nPriorityLow.Checked := False;
  end;
  if TMenuItem(Sender).Name = 'nPriorityLow' then
  begin
    btnPriority.ImageIndex := 7;
    nPriorityHigh.Checked := False;
    nPriorityMed.Checked := False;
    nPriorityLow.Checked := True;
  end;
end;

procedure TFormMailSender.btnPriorityClick(Sender: TObject);
begin
  popupPriority.Popup(mouse.CursorPos.X - 5, mouse.CursorPos.Y - 5);
end;

procedure TFormMailSender.btnCancelClick(Sender: TObject);
begin
  MailSend_Break := True;
  FormMailSender.Close;
end;

procedure TFormMailSender.btnLogClick(Sender: TObject);
begin
  if memoLog.Visible then
  begin
    btnLog.Caption := 'Показать лог';
    memoLog.Visible := False;
  end
  else
  begin
    btnLog.Caption := 'Скрыть лог';
    memoLog.Visible := True;
  end;
end;

function TFormMailSender.IsValidEmail(const Value: string): Boolean;
  function CheckAllowed(const s: string): Boolean;
  var
    i: integer;
  begin
    Result := False;
    for i := 1 to Length(s) do
    begin
      if not CharInSet(s[i], ['a' .. 'z', 'A' .. 'Z', '0' .. '9', '_', '-', '.', '&']) then
        exit;
    end;
    Result := True;
  end;

var
  i: integer;
  namePart, serverPart: string;
begin
  Result := False;
  i := Pos('@', Value);
  if i = 0 then
    exit;
  namePart := Copy(Value, 1, i - 1);
  serverPart := Copy(Value, i + 1, Length(Value));
  if (Length(namePart) = 0) or ((Length(serverPart) < 4)) then
    exit;
  i := Pos('.', serverPart);
  if (i = 0) or (i > (Length(serverPart) - 2)) then
    exit;
  Result := CheckAllowed(namePart) and CheckAllowed(serverPart);
end;

function TFormMailSender.IsValidWeb(const Value: string): Boolean;
  function CheckAllowed(const s: string): Boolean;
  var
    i: integer;
  begin
    Result := False;
    for i := 1 to Length(s) do
    begin
      if not CharInSet(s[i], ['a' .. 'z', 'A' .. 'Z', '0' .. '9', '_', '-', '.']) then
        exit;
    end;
    Result := True;
  end;

var
  i: integer;
  wwwPart, domenPart: string;
begin
  Result := False;
  i := Pos('.', Value);
  if i = 0 then
    exit;
  wwwPart := Copy(Value, 1, i - 1);
  domenPart := Copy(Value, i + 1, Length(Value));
  if (AnsiLowerCase(wwwPart) <> 'www') or ((Length(domenPart) < 5)) then
    exit;
  i := Pos('.', domenPart);
  if (i = 0) or (i > (Length(domenPart) - 2)) then
    exit;
  Result := CheckAllowed(wwwPart) and CheckAllowed(domenPart);
end;

procedure TFormMailSender.GetEmailList(param: integer; Query: TIBCQuery; ID: string; List: TsListBox; ClearList: Boolean);
var
  i: integer;
  email, tmp: string;
  inList: Boolean;
  Q: TIBCQuery;
begin
  if param = 0 then // работаем с Query
  begin
    if ClearList then
      List.Clear;
    if Query.RecordCount = 0 then
      exit;
    Query.Last;
    Query.First;
    for i := 0 to Query.RecordCount - 1 do
    begin
      email := '';
      if Query.FieldValues['EMAIL'] <> null then
        email := Query.FieldValues['EMAIL'];
      if Length(email) > 0 then
      begin
        if email[Length(email)] <> ',' then
          email := email + ',';
        while Pos(',', email) > 0 do
        begin
          inList := False;
          tmp := Copy(email, 0, Pos(',', email));
          Delete(email, 1, Length(tmp));
          tmp := Trim(tmp);
          Delete(tmp, Length(tmp), 1);
          if not IsValidEmail(tmp) then
            inList := True;
          if List.Items.IndexOf(tmp) > -1 then
            inList := True;
          if not inList then
            List.Items.Add(tmp);
        end;
      end;
      Query.Next;
    end;
  end; { if param = 0 then }
  if param = 1 then // работаем с ID
  begin
    if ClearList then
      List.Clear;
    Q := QueryCreate;
    Q.Close;
    Q.SQL.Text := 'select * from BASE where ID = :ID';
    Q.ParamByName('ID').AsString := ID;
    Q.Open;
    if Q.RecordCount = 0 then
      exit;
    email := '';
    if Q.FieldValues['EMAIL'] <> null then
      email := Q.FieldValues['EMAIL'];
    if Length(email) > 0 then
    begin
      if email[Length(email)] <> ',' then
        email := email + ',';
      while Pos(',', email) > 0 do
      begin
        inList := False;
        tmp := Copy(email, 0, Pos(',', email));
        Delete(email, 1, Length(tmp));
        tmp := Trim(tmp);
        Delete(tmp, Length(tmp), 1);
        if not IsValidEmail(tmp) then
          inList := True;
        if List.Items.IndexOf(tmp) > -1 then
          inList := True;
        if not inList then
          List.Items.Add(tmp);
      end;
    end;
    Q.Close;
    Q.Free;
  end; { if param = 1 then }
  if FormMain.IBDatabase1.Connected then
    FormMain.IBDatabase1.Close;
end;

procedure TFormMailSender.GetFirmList(param: integer; Query: TIBCQuery; ID: string; List: TsListBox);
var
  i: integer;
  Q: TIBCQuery;
begin
  if param = 0 then // работаем с Query
  begin
    List.Clear;
    editEmailList_tmp.Clear;
    if Query.RecordCount = 0 then
      exit;
    Query.Last;
    Query.First;
    for i := 0 to Query.RecordCount - 1 do
    begin
      if editEmailList_tmp.Items.IndexOf(Query.FieldByName('ID').AsString) = -1 then
        editEmailList_tmp.Items.Add(Query.FieldByName('ID').AsString);
      editEmailList.Items.Add(Query.FieldByName('NAME').AsString);
      Query.Next;
    end;
  end; { if param = 0 then }
  if param = 1 then // работаем с ID
  begin
    List.Clear;
    editEmailList_tmp.Clear;
    Q := QueryCreate;
    Q.Close;
    Q.SQL.Text := 'select * from BASE where ID = :ID';
    Q.ParamByName('ID').AsString := ID;
    Q.Open;
    if Q.RecordCount = 0 then
      exit;
    if editEmailList_tmp.Items.IndexOf(Q.FieldByName('ID').AsString) = -1 then
      editEmailList_tmp.Items.Add(Q.FieldByName('ID').AsString);
    editEmailList.Items.Add(Q.FieldByName('NAME').AsString);
    Q.Close;
    Q.Free;
  end; { if param = 1 then }
  if FormMain.IBDatabase1.Connected then
    FormMain.IBDatabase1.Close;
end;

procedure TFormMailSender.SendEmail(Sender: TObject);

  procedure SendViaClient;
  var
    MAPISendMail: TMAPISendMail;
    i: integer;
  begin
    MAPISendMail := TMAPISendMail.Create;
    { if MAPISendMail.Prerequisites.IsMapiAvailable then showmessage('MapiAvailable');
      if MAPISendMail.Prerequisites.IsClientAvailable then showmessage('ClientAvailable'); }
    try
      for i := 0 to editEmailList.Count - 1 do
        MAPISendMail.AddRecipient(editEmailList.Items[i]);
      MAPISendMail.SendMail;
    finally
      MAPISendMail.Free;
    end;
  end;

var
  Name, host, port, uname, replyto, domen, login, pass, signature: string;
  i, n, SSL_Method, TLS_Method: integer;
  isError: Boolean;
  txtpart: TIdText;
  tmp, attachCapt: string;
  attach: TIdAttachmentFile;
begin
  if editFrom.ItemIndex < 0 then
  begin
    MessageBox(handle, 'Необходимо указать отправителя', 'Информация', MB_OK or MB_ICONINFORMATION);
    exit;
  end;
  if Trim(editTo.Text) <> '' then
    btnEmailAddClick(btnEmailAdd);
  if editEmailList.Items.Count = 0 then
  begin
    MessageBox(handle, 'Необходимо указать получателей', 'Информация', MB_OK or MB_ICONINFORMATION);
    exit;
  end;
  { if Trim(editMessage.Text) = '' then
    if MessageBox(handle,'Текст сообщения пуст. Продолжить?',
    'Подтверждение',MB_YESNO or MB_ICONQUESTION) = mrNO then exit; }
  if editEmailList.Items.Count > 1 then
    if MessageBox(handle, PChar(IntToStr(editEmailList.Items.Count) + ' адреса(ов) подготовлено для рассылки. Продолжить?'),
      'Подтверждение', MB_YESNO or MB_ICONQUESTION) = MRNO then
      exit;

  SendViaClient;
  FormMailSender.Close;
  exit;

  for i := Length(editFrom.Text) downto 0 do
    if editFrom.Text[i] = '(' then
    begin
      name := Copy(editFrom.Text, i + 1, Length(editFrom.Text));
      Delete(name, Length(name), 1);
      break;
    end;
  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text := 'select * from ACCOUNTS where NAME = :NAME';
  FormMain.IBQuery1.ParamByName('NAME').AsString := name;
  FormMain.IBQuery1.Open;
  if FormMain.IBQuery1.RecordCount = 0 then
  begin
    btnLog.Caption := 'Скрыть лог';
    memoLog.Visible := True;
    memoLog.Lines.Insert(0, '[' + TimeToStr(Now) + '] ERROR: Cannot retrieve profile data');
    exit;
  end;
  host := FormMain.IBQuery1.FieldValues['HOST'];
  port := FormMain.IBQuery1.FieldValues['PORT'];
  uname := FormMain.IBQuery1.FieldValues['UNAME'];
  replyto := FormMain.IBQuery1.FieldValues['REPLYTO'];
  domen := FormMain.IBQuery1.FieldValues['DOMEN'];
  login := FormMain.IBQuery1.FieldValues['LOGIN'];
  pass := DecodeString(FormMain.IBQuery1.FieldValues['PASS']);
  signature := FormMain.IBQuery1.FieldValues['SIGNATURE'];
  SSL_Method := FormMain.IBQuery1.FieldValues['SSL_METHOD'];
  TLS_Method := FormMain.IBQuery1.FieldValues['TLS_METHOD'];
  FormMain.IBQuery1.Close;
  memoLog.Clear;
  { формируем тело сообщения }
  IdMessage.Clear;
  IdMessage.MessageParts.Clear;
  IdMessage.Priority := mpNormal;
  IdMessage.ContentType := 'multipart/mixed';
  IdMessage.From.Address := domen;
  IdMessage.From.Name := uname;
  IdMessage.replyto.EMailAddresses := replyto;
  IdMessage.Subject := editSubj.Text;
  IdMessage.Date := Now;
  if nPriorityHigh.Checked then
    IdMessage.Priority := mpHigh;
  if nPriorityMed.Checked then
    IdMessage.Priority := mpNormal;
  if nPriorityLow.Checked then
    IdMessage.Priority := mpLow;
  { for i := 0 to editEmailList.Items.Count - 1 do
    IdMessage.Recipients.Add.Text := editEmailList.Items[i]; }

  txtpart := TIdText.Create(IdMessage.MessageParts);
  txtpart.ContentType := 'text/plain';
  txtpart.CharSet := 'Windows-1251';
  txtpart.Body.Text := Trim(editMessage.Text);
  if Length(signature) > 0 then
    txtpart.Body.Text := txtpart.Body.Text + #13 + signature;

  for i := 0 to editAttach.Items.Count - 1 do
  begin
    tmp := editAttach.Items[i].Caption;
    attachCapt := '';
    for n := Length(tmp) downto 0 do
      if tmp[n] = '(' then
        attachCapt := Copy(tmp, 0, n - 2);
    if FileExists(editAttach.Items[i].SubItems[0] + attachCapt) then
    begin
      attach := TIdAttachmentFile.Create(IdMessage.MessageParts, editAttach.Items[i].SubItems[0] + attachCapt);
      attach.FileName := '=?windows-1251?B?' + EncodeString(attach.FileName) + '?=';
    end;
  end;
  { настройка компонентов перед отправкой }
  IdSMTP.host := host;
  try
    IdSMTP.port := StrToInt(port); // обычно при использование ssl 495, 587 или стандартный 25
  except
    on E: Exception do
    begin
      WriteLog('TFormMailSender.btnSendClick' + #13 + 'Ошибка: Wrong port number: ' + E.Message);
      btnLog.Caption := 'Скрыть лог';
      memoLog.Visible := True;
      memoLog.Lines.Insert(0, '[' + TimeToStr(Now) + '] ERROR: Wrong port number (' + E.Message + ')');
      exit;
    end;
  end;
  IdSMTP.Username := login;
  IdSMTP.Password := pass;
  // IdSMTP.AuthType := atDefault;

  { это необходимо использовать для SSL }
  IdSSLIOHandlerSocketOpenSSL.Destination := IdSMTP.host + ':' + IntToStr(IdSMTP.port);
  IdSSLIOHandlerSocketOpenSSL.host := IdSMTP.host;
  IdSSLIOHandlerSocketOpenSSL.port := IdSMTP.port;
  IdSSLIOHandlerSocketOpenSSL.DefaultPort := 0;
  case SSL_Method of
    0:
      IdSSLIOHandlerSocketOpenSSL.SSLOptions.Method := sslvTLSv1;
    1:
      IdSSLIOHandlerSocketOpenSSL.SSLOptions.Method := sslvSSLv2;
  end;
  IdSSLIOHandlerSocketOpenSSL.SSLOptions.Mode := sslmUnassigned;

  IdSMTP.IOHandler := IdSSLIOHandlerSocketOpenSSL;
  case TLS_Method of
    0:
      IdSMTP.UseTLS := utUseExplicitTLS;
    1:
      IdSMTP.UseTLS := utUseImplicitTLS;
    2:
      IdSMTP.UseTLS := utNoTLSSupport;
  end;

  { отправляем письмо }
  // idSMTP.ConnectTimeout := 10;
  btnSend.Enabled := False;
  btnLog.Caption := 'Скрыть лог';
  memoLog.Visible := True;
  try
    isError := False;
    MailSend_Break := False;
    btnCancel.Caption := 'Отмена';
    Gauge.Progress := 0;
    Gauge.MaxValue := editEmailList.Items.Count;
    lblGauge.Caption := '';
    lblGauge.Visible := True;
    Gauge.Visible := True;
    for i := 0 to editEmailList.Items.Count - 1 do
    begin
      if MailSend_Break then
        break;
      lblGauge.Caption := IntToStr(i + 1) + ' из ' + IntToStr(editEmailList.Items.Count);
      Application.ProcessMessages;
      try
        IdMessage.Recipients.Clear;
        IdMessage.Recipients.Add.Text := editEmailList.Items[i];
        if not IdSMTP.Connected then
          IdSMTP.Connect;
        IdSMTP.Send(IdMessage);
        Gauge.Progress := i + 1;
        Application.ProcessMessages;
      except
        on E: Exception do
        begin
          isError := True;
          WriteLog('TFormMailSender.btnSendClick' + #13 + 'Ошибка: IdSMTP.Send(IdMessage): ' + E.Message);
          memoLog.Lines.Insert(0, '[' + TimeToStr(Now) + '] ERROR: ' + E.Message);
        end;
      end;
    end;
  finally
    if IdSMTP.Connected then
      IdSMTP.Disconnect;
    if Gauge.Visible then
      Gauge.Visible := False;
    if lblGauge.Visible then
      lblGauge.Visible := False;
    btnCancel.Caption := 'Закрыть';
  end;
  btnSend.Enabled := True;
  if not isError then
  begin
    btnLog.Caption := 'Показать лог';
    memoLog.Visible := False;
    MessageBox(handle, 'Все сообщения были успешно отправлены.', 'Информация', MB_OK or MB_ICONINFORMATION);
  end
  else
  begin
    MessageBox(handle, 'Рассылка завершена. Некоторые сообщения отправить не удалось.', 'Предупреждение', MB_OK or MB_ICONWARNING);
  end;
end;

procedure TFormMailSender.SendRegInfoCheck(Sender: TObject);
var
  Name, host, port, uname, replyto, domen, login, pass, signature: string;
  i, n, z, SSL_Method, TLS_Method: integer;
  isError: Boolean;
  txtpart: TIdText;
  tmp, attachCapt: string;
  attach: TIdAttachmentFile;
  EmailsList: TStringList;

  function GetFirmData(Firm_ID: string): Boolean;
  var
    Q, Q1: TIBCQuery;
    email: string;
    inList: Boolean;
    X: integer;
    Rubr, tmp, adres, country_str, region_str, city_str, ofType, zip_str: string;
    phones: WideString;
    List, list2: TStringList;

    procedure AddColoredLine(AText: string; AColor: TColor; AFontSize: integer; AFontName: TFontName; AFontStyle: TFontStyles);
    begin
      REFirmInfo.SelStart := Length(REFirmInfo.Text);
      REFirmInfo.SelAttributes.Color := AColor;
      REFirmInfo.SelAttributes.Size := AFontSize;
      REFirmInfo.SelAttributes.Name := AFontName;
      REFirmInfo.SelAttributes.Style := AFontStyle;
      REFirmInfo.Lines.Add(AText);
    end;

  begin
    Result := False;
    EmailsList.Clear;
    REFirmInfo.Clear;
    Q := QueryCreate;
    Q.Close;
    Q.SQL.Text := 'select * from BASE where ID = :ID';
    Q.ParamByName('ID').AsString := Firm_ID;
    Q.Open;
    if Q.RecordCount = 0 then
      exit;
    if Q.FieldValues['EMAIL'] <> null then
      email := Q.FieldValues['EMAIL'];
    if Length(email) > 0 then // получаем все emails для фирмы
    begin
      if email[Length(email)] <> ',' then
        email := email + ',';
      while Pos(',', email) > 0 do
      begin
        inList := False;
        tmp := Copy(email, 0, Pos(',', email));
        Delete(email, 1, Length(tmp));
        tmp := Trim(tmp);
        Delete(tmp, Length(tmp), 1);
        if (Pos('@', tmp) = 0) or (Pos('.', tmp) = 0) { можно делать IsValidEMail }
          or (Pos(',', tmp) > 0) then
          inList := True;
        if EmailsList.IndexOf(tmp) <> -1 then
          inList := True;
        if not inList then
          EmailsList.Add(tmp);
      end;
    end;
    if EmailsList.Count > 0 then
    begin // если emails найдены получаем инфо фирмы
      if Q.FieldValues['NAME'] <> null then
      begin
        AddColoredLine(Q.FieldValues['NAME'], clRed, 16, 'Tahoma', []);
        AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);
      end;

      AddColoredLine('Рубрики:', clMaroon, 10, 'Tahoma', [fsBold]);
      if Q.FieldValues['RUBR'] <> null then
        Rubr := Q.FieldValues['RUBR'];
      while Pos('$', Rubr) > 0 do
      begin
        tmp := Copy(Rubr, 0, Pos('$', Rubr));
        Delete(Rubr, 1, Length(tmp));
        Delete(tmp, 1, 1);
        Delete(tmp, Length(tmp), 1);
        Q1 := QueryCreate;
        Q1.Close;
        Q1.SQL.Text := 'select * from RUBRIKATOR where ID = :ID';
        Q1.ParamByName('ID').AsString := tmp;
        Q1.Open;
        AddColoredLine(Q1.FieldValues['NAME'], clWindowText, 10, 'Tahoma', []);
        Q1.Close;
        Q1.Free;
      end;
      AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);

      AddColoredLine('Направления:', clMaroon, 10, 'Tahoma', [fsBold]);
      if Q.FieldValues['NAPRAVLENIE'] <> null then
        Rubr := Q.FieldValues['NAPRAVLENIE'];
      while Pos('$', Rubr) > 0 do
      begin
        tmp := Copy(Rubr, 0, Pos('$', Rubr));
        Delete(Rubr, 1, Length(tmp));
        Delete(tmp, 1, 1);
        Delete(tmp, Length(tmp), 1);
        Q1 := QueryCreate;
        Q1.Close;
        Q1.SQL.Text := 'select * from NAPRAVLENIE where ID = :ID';
        Q1.ParamByName('ID').AsString := tmp;
        Q1.Open;
        AddColoredLine(Q1.FieldValues['NAME'], clWindowText, 10, 'Tahoma', []);
        Q1.Close;
        Q1.Free;
      end;
      AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);

      if Q.FieldValues['EMAIL'] <> null then
        if Trim(Q.FieldValues['EMAIL']) <> '' then
        begin
          AddColoredLine('Эл. почта:', clMaroon, 10, 'Tahoma', [fsBold]);
          AddColoredLine(Q.FieldValues['EMAIL'], clWindowText, 10, 'Tahoma', []);
          AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);
        end;

      if Q.FieldValues['WEB'] <> null then
        if Trim(Q.FieldValues['WEB']) <> '' then
        begin
          AddColoredLine('Сайт:', clMaroon, 10, 'Tahoma', [fsBold]);
          AddColoredLine(Q.FieldValues['WEB'], clWindowText, 10, 'Tahoma', []);
          AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);
        end;

      AddColoredLine('Адреса:', clMaroon, 10, 'Tahoma', [fsBold]);
      List := TStringList.Create;
      List.Text := Q.FieldValues['ADRES'];
      phones := Q.FieldValues['PHONES'];
      for X := 0 to List.Count - 1 do
      begin
        list2 := FormEditor.ParseAdresFieldToEntriesList(List[X]);

        tmp := Copy(phones, 0, Pos('$', phones));
        Delete(phones, 1, Length(tmp));
        tmp := Copy(phones, 0, Pos('$', phones));
        Delete(phones, 1, Length(tmp));
        Delete(tmp, 1, 1);
        Delete(tmp, Length(tmp), 1);

        if list2[0] = '1' then
        begin
          // list2[0] = CBAdres; list2[1] = NO; list2[2] = OfficeType; list2[3] = ZIP;
          // list2[4] = Street; list2[5] = Country; list2[6] = Region; list2[7] = City;
          ofType := list2[2];
          zip_str := list2[3];
          country_str := list2[5];
          region_str := list2[6];
          city_str := list2[7];
          if Trim(ofType) <> '' then
            ofType := GetNameByID('OFFICETYPE', ofType) + ' - ';
          if Trim(zip_str) <> '' then
            zip_str := zip_str + ', ';
          if Trim(country_str) <> '' then
            country_str := GetNameByID('COUNTRY', country_str) + ', ';
          if Trim(city_str) <> '' then
            city_str := GetNameByID('CITY', city_str) + ', ';
          adres := ofType + zip_str + country_str + city_str + list2[4];
          { officetype - zip, country, city, street }
          if Trim(adres) <> '' then
            AddColoredLine(adres, clWindowText, 10, 'Tahoma', []);
          if Trim(tmp) <> '' then
            AddColoredLine(tmp, clNavy, 10, 'Tahoma', []);
        end;

        list2.Free;
      end;
      List.Free;
    end; // if EmailsList.Count > 0
    Q.Close;
    Q.Free;
    Result := EmailsList.Count > 0;
  end;

begin
  if editFrom.ItemIndex < 0 then
  begin
    MessageBox(handle, 'Необходимо указать отправителя', 'Информация', MB_OK or MB_ICONINFORMATION);
    exit;
  end;
  if editEmailList_tmp.Items.Count = 0 then
  begin
    MessageBox(handle, 'Список фирм для сверки пуст', 'Информация', MB_OK or MB_ICONINFORMATION);
    exit;
  end;
  if Trim(editMessage.Text) = '' then
    if MessageBox(handle, 'Текст сообщения пуст. Продолжить?', 'Подтверждение', MB_YESNO or MB_ICONQUESTION) = MRNO then
      exit;
  if editEmailList_tmp.Items.Count > 1 then
    if MessageBox(handle, PChar(IntToStr(editEmailList_tmp.Items.Count) + ' фирм(ы) подготовлено для рассылки. Продолжить?'),
      'Подтверждение', MB_YESNO or MB_ICONQUESTION) = MRNO then
      exit;
  for i := Length(editFrom.Text) downto 0 do
    if editFrom.Text[i] = '(' then
    begin
      name := Copy(editFrom.Text, i + 1, Length(editFrom.Text));
      Delete(name, Length(name), 1);
      break;
    end;
  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text := 'select * from ACCOUNTS where NAME = :NAME';
  FormMain.IBQuery1.ParamByName('NAME').AsString := name;
  FormMain.IBQuery1.Open;
  if FormMain.IBQuery1.RecordCount = 0 then
  begin
    btnLog.Caption := 'Скрыть лог';
    memoLog.Visible := True;
    memoLog.Lines.Insert(0, '[' + TimeToStr(Now) + '] ERROR: Cannot retrieve profile data');
    exit;
  end;
  host := FormMain.IBQuery1.FieldValues['HOST'];
  port := FormMain.IBQuery1.FieldValues['PORT'];
  uname := FormMain.IBQuery1.FieldValues['UNAME'];
  replyto := FormMain.IBQuery1.FieldValues['REPLYTO'];
  domen := FormMain.IBQuery1.FieldValues['DOMEN'];
  login := FormMain.IBQuery1.FieldValues['LOGIN'];
  pass := DecodeString(FormMain.IBQuery1.FieldValues['PASS']);
  signature := FormMain.IBQuery1.FieldValues['SIGNATURE'];
  SSL_Method := FormMain.IBQuery1.FieldValues['SSL_METHOD'];
  TLS_Method := FormMain.IBQuery1.FieldValues['TLS_METHOD'];
  FormMain.IBQuery1.Close;
  memoLog.Clear;

  { настройка компонентов перед отправкой }
  IdSMTP.host := host;
  try
    IdSMTP.port := StrToInt(port); // обычно при использование ssl 495, 587 или стандартный 25
  except
    on E: Exception do
    begin
      WriteLog('TFormMailSender.btnSendClick' + #13 + 'Ошибка: Wrong port number: ' + E.Message);
      btnLog.Caption := 'Скрыть лог';
      memoLog.Visible := True;
      memoLog.Lines.Insert(0, '[' + TimeToStr(Now) + '] ERROR: Wrong port number (' + E.Message + ')');
      exit;
    end;
  end;
  IdSMTP.Username := login;
  IdSMTP.Password := pass;
  // IdSMTP.AuthType := atDefault;

  { это необходимо использовать для SSL }
  IdSSLIOHandlerSocketOpenSSL.Destination := IdSMTP.host + ':' + IntToStr(IdSMTP.port);
  IdSSLIOHandlerSocketOpenSSL.host := IdSMTP.host;
  IdSSLIOHandlerSocketOpenSSL.port := IdSMTP.port;
  IdSSLIOHandlerSocketOpenSSL.DefaultPort := 0;
  case SSL_Method of
    0:
      IdSSLIOHandlerSocketOpenSSL.SSLOptions.Method := sslvTLSv1;
    1:
      IdSSLIOHandlerSocketOpenSSL.SSLOptions.Method := sslvSSLv2;
  end;
  IdSSLIOHandlerSocketOpenSSL.SSLOptions.Mode := sslmUnassigned;

  IdSMTP.IOHandler := IdSSLIOHandlerSocketOpenSSL;
  case TLS_Method of
    0:
      IdSMTP.UseTLS := utUseExplicitTLS;
    1:
      IdSMTP.UseTLS := utUseImplicitTLS;
    2:
      IdSMTP.UseTLS := utNoTLSSupport;
  end;

  btnSend.Enabled := False;
  btnLog.Caption := 'Скрыть лог';
  memoLog.Visible := True;
  EmailsList := TStringList.Create;
  try
    isError := False;
    MailSend_Break := False;
    btnCancel.Caption := 'Отмена';
    Gauge.Progress := 0;
    Gauge.MaxValue := editEmailList_tmp.Items.Count;
    lblGauge.Caption := '';
    lblGauge.Visible := True;
    Gauge.Visible := True;
    for z := 0 to editEmailList_tmp.Items.Count - 1 do
    begin
      if MailSend_Break then
        break;
      lblGauge.Caption := IntToStr(z + 1) + ' из ' + IntToStr(editEmailList_tmp.Items.Count);
      Gauge.Progress := z + 1;
      Application.ProcessMessages;
      if not(GetFirmData(editEmailList_tmp.Items[z])) then
        Continue;
      { формируем тело сообщения }
      IdMessage.Clear;
      IdMessage.MessageParts.Clear;
      IdMessage.Priority := mpNormal;
      IdMessage.ContentType := 'multipart/mixed';
      IdMessage.From.Address := domen;
      IdMessage.From.Name := uname;
      IdMessage.replyto.EMailAddresses := replyto;
      IdMessage.Subject := editSubj.Text;
      IdMessage.Date := Now;
      if nPriorityHigh.Checked then
        IdMessage.Priority := mpHigh;
      if nPriorityMed.Checked then
        IdMessage.Priority := mpNormal;
      if nPriorityLow.Checked then
        IdMessage.Priority := mpLow;

      txtpart := TIdText.Create(IdMessage.MessageParts);
      txtpart.ContentType := 'text/plain';
      txtpart.CharSet := 'Windows-1251';
      txtpart.Body.Text := Trim(editMessage.Text);
      if Length(Trim(REFirmInfo.Text)) > 0 then
        txtpart.Body.Text := txtpart.Body.Text + #13 + Trim(REFirmInfo.Text);
      if Length(signature) > 0 then
        txtpart.Body.Text := txtpart.Body.Text + #13 + signature;

      for i := 0 to editAttach.Items.Count - 1 do
      begin
        tmp := editAttach.Items[i].Caption;
        for n := Length(tmp) downto 0 do
          if tmp[n] = '(' then
            attachCapt := Copy(tmp, 0, n - 2);
        if FileExists(editAttach.Items[i].SubItems[0] + attachCapt) then
        begin
          attach := TIdAttachmentFile.Create(IdMessage.MessageParts, editAttach.Items[i].SubItems[0] + attachCapt);
          attach.FileName := '=?windows-1251?B?' + EncodeString(attach.FileName) + '?=';
        end;
      end;

      { отправляем письмо }
      for i := 0 to EmailsList.Count - 1 do
      begin
        try
          IdMessage.Recipients.Clear;
          IdMessage.Recipients.Add.Text := EmailsList[i];
          if not IdSMTP.Connected then
            IdSMTP.Connect;
          IdSMTP.Send(IdMessage);
        except
          on E: Exception do
          begin
            isError := True;
            WriteLog('TFormMailSender.btnSendClick' + #13 + 'Ошибка: IdSMTP.Send(IdMessage): ' + E.Message);
            memoLog.Lines.Insert(0, '[' + TimeToStr(Now) + '] ERROR: ' + E.Message);
          end;
        end;
      end;
    end;
  finally
    if IdSMTP.Connected then
      IdSMTP.Disconnect;
    if Gauge.Visible then
      Gauge.Visible := False;
    if lblGauge.Visible then
      lblGauge.Visible := False;
    btnCancel.Caption := 'Закрыть';
    EmailsList.Free;
    FormMain.IBDatabase1.Close;
  end;
  btnSend.Enabled := True;
  if not isError then
  begin
    btnLog.Caption := 'Показать лог';
    memoLog.Visible := False;
    MessageBox(handle, 'Все сообщения были успешно отправлены.', 'Информация', MB_OK or MB_ICONINFORMATION);
  end
  else
  begin
    MessageBox(handle, 'Рассылка завершена. Некоторые сообщения отправить не удалось.', 'Предупреждение', MB_OK or MB_ICONWARNING);
  end;
end;

end.
