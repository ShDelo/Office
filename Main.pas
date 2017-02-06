unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DateUtils, psAPI, sSkinProvider, sSkinManager, ComCtrls, ToolWin, sToolBar,
  sTreeView, sListView, IBDatabase, DB, IBCustomDataSet, IBQuery, StdCtrls,
  acProgressBar, ExtCtrls, sStatusBar, ImgList, acAlphaImageList,
  sPageControl, NxScrollControl, NxCustomGridControl, NxCustomGrid, NxGrid,
  NxColumns, NxColumnClasses, Menus, sSpeedButton, sPanel, Buttons,
  ComObj, sRichEdit, sEdit, sComboBoxes, sComboBox, NxEdit, acCoolBar,
  StrUtils, IniFiles, sCheckBox, sButton, JvExStdCtrls, JvRichEdit, ShellApi,
  sMemo;

procedure debug(Text: string; Params: array of TVarRec);
procedure WriteLog(Text: string);
function QueryCreate: TIBQuery;

type
  TFormMain = class(TForm)
    sStatusBar1: TsStatusBar;
    TimerMemory: TTimer;
    ImageList32: TsAlphaImageList;
    ImageList1: TImageList;
    Splitter1: TSplitter;
    PopupMenu_Tabs: TPopupMenu;
    nCloseTab: TMenuItem;
    nCloseAllButThisTab: TMenuItem;
    nCloseAllTab: TMenuItem;
    PopupMenu_Rubr: TPopupMenu;
    nAddRec: TMenuItem;
    nDeleteRec: TMenuItem;
    nEditRec: TMenuItem;
    IBDatabase1: TIBDatabase;
    IBTransaction1: TIBTransaction;
    IBQuery1: TIBQuery;
    nSeparator1: TMenuItem;
    nSendEmail: TMenuItem;
    imgMenus: TImageList;
    panelRubrikator: TsPanel;
    TVRubrikator: TsTreeView;
    panelSearchFast: TsPanel;
    editSearchFast: TNxEdit;
    panelGeneral: TsPanel;
    sPageControl1: TsPageControl;
    sTabSheet1: TsTabSheet;
    SGGeneral: TNextGrid;
    NxImageColumn1: TNxImageColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    NxTextColumn7: TNxTextColumn;
    NxTextColumn8: TNxTextColumn;
    NxTextColumn1: TNxTextColumn;
    sCoolBar1: TsCoolBar;
    sToolBar1: TsToolBar;
    BtnRecAdd: TsSpeedButton;
    BtnSendEmail: TsSpeedButton;
    btnDivider2: TsSpeedButton;
    BtnReportSimple: TsSpeedButton;
    btnDivider3: TsSpeedButton;
    PopupMenu_Emails: TPopupMenu;
    nBtnAccounts: TMenuItem;
    btnDivider4: TsSpeedButton;
    BtnDirectory: TsSpeedButton;
    nSeparator2: TMenuItem;
    nRubtToReport: TMenuItem;
    NxTextColumn9: TNxTextColumn;
    NxTextColumn10: TNxTextColumn;
    PopupMenu_Columns: TPopupMenu;
    nCol_Name: TMenuItem;
    nCol_Curator: TMenuItem;
    nCol_Added: TMenuItem;
    nCol_Edited: TMenuItem;
    nCol_WEB: TMenuItem;
    nCol_EMAIL: TMenuItem;
    nCol_Type: TMenuItem;
    nCol_FIO: TMenuItem;
    nCol_Rubr: TMenuItem;
    BtnReportAnvanced: TsSpeedButton;
    nSeparator3: TMenuItem;
    nLoadSGonRubrChange: TMenuItem;
    nRubtToReportSimple: TMenuItem;
    nRubtToReportAdvanced: TMenuItem;
    nRelations: TMenuItem;
    nRegInfoCheck: TMenuItem;
    nSendEmailNew: TMenuItem;
    nSendEmailAddTo: TMenuItem;
    sSkinManager1: TsSkinManager;
    sSkinProvider1: TsSkinProvider;
    nRubrInfo: TMenuItem;
    NxTextColumn11: TNxTextColumn;
    nCol_Relevance: TMenuItem;
    memoDebug: TsMemo;
    function CurrentProcessMemory: Cardinal;
    function SearchNode(component: TsTreeView; id: integer; itemlevel: integer): TTreeNode;
    procedure DisableAllForms(StayActive: string);
    procedure EnableAllForms(StayNotActive: string);
    procedure ReloadDataGlobal(Sender: TObject);
    procedure LoadTempTables;
    procedure BuildRubrikatorTree(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure IniCreateMain;
    procedure IniLoadMain;
    procedure IniSaveMain;
    procedure ClearSearhcAndFilter;
    procedure TimerMemoryTimer(Sender: TObject);
    procedure TVRubrikatorChange(Sender: TObject; Node: TTreeNode);
    procedure SGAfterSort(Sender: TObject; ACol: integer);
    procedure TVRubrikatorDblClick(Sender: TObject);
    procedure nCloseTabClick(Sender: TObject);
    procedure nCloseAllButThisTabClick(Sender: TObject);
    procedure nCloseAllTabClick(Sender: TObject);
    procedure sPageControl1MouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: integer);
    procedure TVRubrikatorContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure nDeleteRecClick(Sender: TObject);
    procedure nEditRecClick(Sender: TObject);
    procedure nAddRecClick(Sender: TObject);
    procedure PrintTabData(Sender: TObject);
    function CloseTabByID(id: string): Boolean;
    procedure OpenTabByID(id: string);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure nSendEmailNewClick(Sender: TObject);
    procedure editSearchFastEnter(Sender: TObject);
    procedure editSearchFastExit(Sender: TObject);
    procedure editSearchFastKeyPress(Sender: TObject; var Key: Char);
    procedure SGAddRow(Grid: TNextGrid; Activity, Relevance: integer; Name, Curator, DateAdded, DateEdited, Web, Email, fType, id, FIO,
      Rubr: string);
    procedure SearchFast(Sender: TObject);
    procedure nBtnAccountsClick(Sender: TObject);
    procedure BtnSendEmailClick(Sender: TObject);
    procedure BtnReportSimpleClick(Sender: TObject);
    procedure BtnReportAnvancedClick(Sender: TObject);
    procedure BtnDirectoryClick(Sender: TObject);
    procedure BtnRecAddClick(Sender: TObject);
    procedure SGGeneralMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: integer);
    procedure nCol_NameClick(Sender: TObject);
    procedure editSearchFastGlyphClick(Sender: TObject);
    procedure nLoadSGonRubrChangeClick(Sender: TObject);
    procedure nRubtToReportAdvancedClick(Sender: TObject);
    procedure nRubtToReportSimpleClick(Sender: TObject);
    procedure nRelationsClick(Sender: TObject);
    procedure RichEditURLClick(Sender: TObject; const URLText: string; Button: TMouseButton);
    procedure nRegInfoCheckClick(Sender: TObject);
    procedure nSendEmailAddToClick(Sender: TObject);
    procedure nRubrInfoClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure WMGetMinMaxInfo(var M: TWMGetMinMaxInfo); message WM_GetMinMaxInfo;
  end;

var
  FormMain: TFormMain;
  AppPath: string;
  sgCurator_tmp, sgType_tmp, sgRubr_tmp, sgNapr_tmp: TNextGrid;
  IniMain: TIniFile;
  LoadSGonRubrChange: Boolean;

implementation

uses Editor, Logo, MailSend, Report, Directory, ReportSimple, Relations,
  Dublicate;

const
  bDebug: boolean = true;

{$R *.dfm}
  { #BACKUP [changes].txt }
  { #BACKUP office.ini }
  { #BACKUP MapiEmail.pas }

procedure debug(Text: string; Params: array of TVarRec);
begin
  if bDebug then
  begin
    FormMain.memoDebug.Lines.Add(DateTimeToStr(now) + ': ' + Format(Text, Params));
  end;
end;

function QueryCreate: TIBQuery;
var
  Query: TIBQuery;
begin
  Query := TIBQuery.Create(nil);
  Query.Database := FormMain.IBDatabase1;
  Query.Transaction := FormMain.IBTransaction1;
  result := Query;
end;

procedure TFormMain.WMGetMinMaxInfo(var M: TWMGetMinMaxInfo);
begin
  M.MinMaxInfo^.PTMaxPosition.X := 0;
  M.MinMaxInfo^.PTMinTrackSize.X := 800;
  M.MinMaxInfo^.PTMaxPosition.Y := 0;
  M.MinMaxInfo^.PTMinTrackSize.Y := 600;
end;

procedure WriteLog(Text: string);
var
  LogFile, d: string;
  f: TextFile;
begin
  d := DateToStr(now);
  delete(d, 1, 3);
  LogFile := AppPath + 'log_' + d + '.txt';
  AssignFile(f, LogFile);
  try
    if not FileExists(LogFile) then
      Rewrite(f)
    else
      Append(f);
    Writeln(f, DateTimeToStr(now) + ': ' + Text);
  finally
    CloseFile(f);
  end;
end;

procedure TFormMain.DisableAllForms(StayActive: string);
begin
  StayActive := AnsiLowerCase(StayActive);
  if StayActive <> 'formmain' then
    FormMain.Enabled := False;
  if StayActive <> 'formeditor' then
    FormEditor.Enabled := False;
  if StayActive <> 'formreport' then
    FormReport.Enabled := False;
  if StayActive <> 'formmailsender' then
    FormMailSender.Enabled := False;
  if StayActive <> 'formdirectory' then
    FormDirectory.Enabled := False;
  if StayActive <> 'formreportsimple' then
    FormReportSimple.Enabled := False;
  if StayActive <> 'formrelations' then
    FormRelations.Enabled := False;
  if StayActive <> 'formdublicare' then
    FormDublicate.Enabled := False;
end;

procedure TFormMain.EnableAllForms(StayNotActive: string);
begin
  StayNotActive := AnsiLowerCase(StayNotActive);
  if StayNotActive <> 'formmain' then
    FormMain.Enabled := True;
  if StayNotActive <> 'formeditor' then
    FormEditor.Enabled := True;
  if StayNotActive <> 'formreport' then
    FormReport.Enabled := True;
  if StayNotActive <> 'formmailsender' then
    FormMailSender.Enabled := True;
  if StayNotActive <> 'formdirectory' then
    FormDirectory.Enabled := True;
  if StayNotActive <> 'formreportsimple' then
    FormReportSimple.Enabled := True;
  if StayNotActive <> 'formrelations' then
    FormRelations.Enabled := True;
  if StayNotActive <> 'formdublicare' then
    FormDublicate.Enabled := True;
end;

function TFormMain.CurrentProcessMemory: Cardinal;
var
  MemCounters: TProcessMemoryCounters;
begin
  MemCounters.cb := SizeOf(MemCounters);
  if GetProcessMemoryInfo(GetCurrentProcess, @MemCounters, SizeOf(MemCounters)) then
    Result := MemCounters.WorkingSetSize
  else
  begin
    RaiseLastOSError;
    Result := 0
  end;
end;

procedure TFormMain.TimerMemoryTimer(Sender: TObject);
var
  e: extended;
begin
  e := FormMain.CurrentProcessMemory / 1024;
  sStatusBar1.Panels[0].Text := FloatToStr(e) + 'kb';
end;

function CustomSortProc(Node1, Node2: TTreeNode; iUpToThisLevel: integer): integer; stdcall;
begin
  Result := AnsiStrIComp(PChar(Node1.Text), PChar(Node2.Text));
end;

procedure TFormMain.sPageControl1MouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: integer);
begin
  if Button = mbRight then
  begin
    if Y <= 19 then // учим PageControl понимать правый клик только по шапке
      PopupMenu_Tabs.Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y)
    else
      exit;
  end;
end;

procedure TFormMain.TVRubrikatorContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
var
  treeNode: TTreeNode;
begin
  treeNode := TVRubrikator.GetNodeAt(MousePos.X, MousePos.Y);
  if Assigned(treeNode) then // учим TreeView понимать правый клик
    TVRubrikator.Selected := treeNode
  else
    Handled := True;
end;

function TFormMain.SearchNode(component: TsTreeView; id: integer; itemlevel: integer): TTreeNode;
var
  Searching: Boolean;
  Noddy: TTreeNode;
begin
  Result := nil;
  if component.Items.Count = 0 then
    exit;
  Noddy := component.Items[0];
  Searching := True;
  while (Searching) and (Noddy <> nil) do
    if (integer(Noddy.Data) = id) and (Noddy.Level = itemlevel) then
    begin
      Searching := False;
      Result := Noddy;
    end
    else
      Noddy := Noddy.GetNext;
end;

procedure TFormMain.ReloadDataGlobal(Sender: TObject);
begin
  if TForm(Sender).Name <> 'FormEditor' then
    FormEditor.Close;
  if TForm(Sender).Name <> 'FormReport' then
    FormReport.Close;
  if TForm(Sender).Name <> 'FormMailSender' then
    FormMailSender.Close;
  if TForm(Sender).Name <> 'FormDirectory' then
    FormDirectory.Close;
  if TForm(Sender).Name <> 'FormReportSimple' then
    FormReportSimple.Close;
  if TForm(Sender).Name <> 'FormRelations' then
    FormRelations.Close;
  if TForm(Sender).Name <> 'FormDublicate' then
    FormDublicate.Close;
  DisableAllForms('');
  TVRubrikator.OnChange := nil;
  TVRubrikator.Items.BeginUpdate;
  SGGeneral.BeginUpdate;
  FormLogo := TFormLogo.Create(Application);
  FormLogo.Show;
  FormLogo.Update;
  LoadTempTables;
  BuildRubrikatorTree(FormMain);
  FormEditor.LoadDataEditor;
  FormDirectory.LoadDataDirectory;
  FormLogo.Hide;
  FormLogo.Free;
  TVRubrikator.Items.EndUpdate;
  ClearSearhcAndFilter;
  SGGeneral.EndUpdate;
  TVRubrikator.OnChange := TVRubrikatorChange;
  TVRubrikatorChange(TVRubrikator, TVRubrikator.Selected);
  IBDatabase1.Close;
  Directory.isDataEdited := False; // сбрасуем релоад
  EnableAllForms('');
end;

procedure TFormMain.LoadTempTables;
var
  Q_LTT: TIBQuery;
  i: integer;
begin
  FormLogo.sGauge1.MinValue := 0;
  FormLogo.sLabel1.Caption := 'Подключение временных таблиц ...';
  if not Assigned(sgCurator_tmp) then
  begin
    sgCurator_tmp := TNextGrid.Create(FormMain);
    sgCurator_tmp.Parent := FormMain;
    sgCurator_tmp.Visible := False;
    sgCurator_tmp.Columns.Add(TNxTextColumn, 'Название');
    sgCurator_tmp.Columns.Add(TNxTextColumn, 'ID');
    sgCurator_tmp.Columns[0].SortType := stAlphabetic;
    sgCurator_tmp.Columns[0].Sorted := True;
  end;
  sgCurator_tmp.ClearRows;
  if not Assigned(sgRubr_tmp) then
  begin
    sgRubr_tmp := TNextGrid.Create(FormMain);
    sgRubr_tmp.Parent := FormMain;
    sgRubr_tmp.Visible := False;
    sgRubr_tmp.Columns.Add(TNxTextColumn, 'Название');
    sgRubr_tmp.Columns.Add(TNxTextColumn, 'ID');
    sgRubr_tmp.Columns[0].SortType := stAlphabetic;
    sgRubr_tmp.Columns[0].Sorted := True;
  end;
  sgRubr_tmp.ClearRows;
  if not Assigned(sgType_tmp) then
  begin
    sgType_tmp := TNextGrid.Create(FormMain);
    sgType_tmp.Parent := FormMain;
    sgType_tmp.Visible := False;
    sgType_tmp.Columns.Add(TNxTextColumn, 'Название');
    sgType_tmp.Columns.Add(TNxTextColumn, 'ID');
    sgType_tmp.Columns[0].SortType := stAlphabetic;
    sgType_tmp.Columns[0].Sorted := True;
  end;
  sgType_tmp.ClearRows;
  if not Assigned(sgNapr_tmp) then
  begin
    sgNapr_tmp := TNextGrid.Create(FormMain);
    sgNapr_tmp.Parent := FormMain;
    sgNapr_tmp.Visible := False;
    sgNapr_tmp.Columns.Add(TNxTextColumn, 'Название');
    sgNapr_tmp.Columns.Add(TNxTextColumn, 'ID');
    sgNapr_tmp.Columns[0].SortType := stAlphabetic;
    sgNapr_tmp.Columns[0].Sorted := True;
  end;
  sgNapr_tmp.ClearRows;
  Q_LTT := QueryCreate;
  Q_LTT.Close;
  Q_LTT.SQL.Text := 'select * from RUBRIKATOR';
  Q_LTT.Open;
  Q_LTT.FetchAll;
  // Получаю рубрики первыми чтобы сформировать ProgressBar
  // 4 = шаги TempTables + кол-во рубрик + 8 = шаги из Editor.LoadDataEditor + 8 = шаги из Directory.LoadDataDirectory; 
  FormLogo.sGauge1.MaxValue := 4 + Q_LTT.RecordCount + 8 + 8;
  sgRubr_tmp.BeginUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  for i := 1 to Q_LTT.RecordCount do
  begin
    sgRubr_tmp.AddRow;
    sgRubr_tmp.Cells[0, sgRubr_tmp.LastAddedRow] := Q_LTT.FieldValues['NAME'];
    sgRubr_tmp.Cells[1, sgRubr_tmp.LastAddedRow] := Q_LTT.FieldValues['ID'];
    Q_LTT.Next;
    Application.ProcessMessages;
  end;
  sgRubr_tmp.Resort;
  sgRubr_tmp.EndUpdate;
  Q_LTT.Close;
  Q_LTT.SQL.Text := 'select * from CURATOR';
  Q_LTT.Open;
  Q_LTT.FetchAll;
  sgCurator_tmp.BeginUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  for i := 1 to Q_LTT.RecordCount do
  begin
    sgCurator_tmp.AddRow;
    sgCurator_tmp.Cells[0, sgCurator_tmp.LastAddedRow] := Q_LTT.FieldValues['NAME'];
    sgCurator_tmp.Cells[1, sgCurator_tmp.LastAddedRow] := Q_LTT.FieldValues['ID'];
    Q_LTT.Next;
    Application.ProcessMessages;
  end;
  sgCurator_tmp.Resort;
  sgCurator_tmp.EndUpdate;
  Q_LTT.Close;
  Q_LTT.SQL.Text := 'select * from TYPE';
  Q_LTT.Open;
  Q_LTT.FetchAll;
  sgType_tmp.BeginUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  for i := 1 to Q_LTT.RecordCount do
  begin
    sgType_tmp.AddRow;
    sgType_tmp.Cells[0, sgType_tmp.LastAddedRow] := Q_LTT.FieldValues['NAME'];
    sgType_tmp.Cells[1, sgType_tmp.LastAddedRow] := Q_LTT.FieldValues['ID'];
    Q_LTT.Next;
    Application.ProcessMessages;
  end;
  sgType_tmp.Resort;
  sgType_tmp.EndUpdate;
  Q_LTT.Close;
  Q_LTT.SQL.Text := 'select * from NAPRAVLENIE';
  Q_LTT.Open;
  Q_LTT.FetchAll;
  sgNapr_tmp.BeginUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  for i := 1 to Q_LTT.RecordCount do
  begin
    sgNapr_tmp.AddRow;
    sgNapr_tmp.Cells[0, sgNapr_tmp.LastAddedRow] := Q_LTT.FieldValues['NAME'];
    sgNapr_tmp.Cells[1, sgNapr_tmp.LastAddedRow] := Q_LTT.FieldValues['ID'];
    Q_LTT.Next;
    Application.ProcessMessages;
  end;
  sgNapr_tmp.Resort;
  sgNapr_tmp.EndUpdate;
  Q_LTT.Close;
  Q_LTT.Free;
  IBDatabase1.Close;
end;

procedure TFormMain.BuildRubrikatorTree(Sender: TObject);
var
  Q: TIBQuery;
  i, n: integer;
  Node, tmp: TTreeNode;
  // T : TDateTime;
  msg: string;
begin
  // T := now;
  TVRubrikator.Items.Clear;
  IBQuery1.Close;
  IBQuery1.SQL.Text := 'select * from RUBRIKATOR';
  IBQuery1.Open;
  IBQuery1.FetchAll;
  TVRubrikator.Items.BeginUpdate;
  for i := 1 to IBQuery1.RecordCount do
  begin
    if TForm(Sender).Name = 'FormMain' then
    begin
      FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
      msg := 'Загрузка данных ... (' + IntToStr(i) + '/' + IntToStr(IBQuery1.RecordCount) + ')';
      FormLogo.sLabel1.Caption := msg;
      FormLogo.sLabel1.Update;
    end;

    tmp := TVRubrikator.Items.AddChildObject(nil, IBQuery1.FieldByName('Name').AsString, Pointer(IBQuery1.FieldByName('ID').AsInteger));
    tmp.ImageIndex := 0;
    tmp.SelectedIndex := 0;

    Q := TIBQuery.Create(FormMain);
    Q.Database := IBDatabase1;
    Q.Transaction := IBTransaction1;
    Q.Close;
    Q.SQL.Text := 'select ID,ACTIVITY,NAME from BASE where RUBR like :RUBR';
    Q.ParamByName('RUBR').AsString := '%#' + IBQuery1.FieldByName('ID').AsString + '$%';
    Q.Open;
    Q.FetchAll;
    Node := SearchNode(TVRubrikator, IBQuery1.FieldByName('ID').AsInteger, 0);
    for n := 1 to Q.RecordCount do
    begin
      tmp := TVRubrikator.Items.AddChildObject(Node, Q.FieldByName('Name').AsString, Pointer(Q.FieldByName('ID').AsInteger));
      if Q.FieldByName('ACTIVITY').AsInteger = -1 then
      begin
        tmp.ImageIndex := 1;
        tmp.SelectedIndex := 1;
      end
      else
      begin
        tmp.ImageIndex := 2;
        tmp.SelectedIndex := 2;
      end;
      Q.Next;
    end;
    Q.Close;
    Q.Free;
    IBQuery1.Next;
  end;
  IBQuery1.Close;
  TVRubrikator.CustomSort(@CustomSortProc, 0, True);
  if TVRubrikator.Items.Count > 0 then
    TVRubrikator.Items[0].Selected := True;
  TVRubrikator.Items.EndUpdate;
  IBDatabase1.Close;
  // sLabel1.Caption := 'BuildTree: '+IntToStr(MilliSecondsBetween(T, NOW))+' Ms';
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  memoDebug.Visible := bDebug;
  AppPath := ExtractFilePath(Application.ExeName);
  IBDatabase1.DatabaseName := AppPath + 'usrdt.msq';
  WriteLog('* Запуск программы *');
  try
    LoadTempTables;
    BuildRubrikatorTree(Sender);
  except
    on e: Exception do
    begin
      WriteLog('TFormMain.FormCreate' + #13 + 'Ошибка подключения файлов баз данных' + #13 + e.Message);
      MessageBox(FormLogo.Handle, PChar('Ошибка подключения файлов баз данных' + #13 + e.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      FormLogo.Free;
      Halt;
    end;
  end;
  IniCreateMain;
  IniLoadMain;
  sStatusBar1.Panels[1].Text := 'Фирм в базе: ' + FormEditor.GetFirmCount;
end;

procedure TFormMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  IBDatabase1.Connected := False;
  // с формы FormEditor
  Editor.list_GC_IDs.Free;
  // с формы FormRelations
  Relations.listRubrRelations.Free;
  // с формы FormEmailSander
  MailSend.memoLog.Free;
  MailSend.REFirmInfo.Free;
  MailSend.editEmailList_tmp.Free;
  // с формы FormMain
  sgCurator_tmp.Free;
  sgNapr_tmp.Free;
  sgRubr_tmp.Free;
  sgType_tmp.Free;
  IniSaveMain;
  WriteLog('* Завершение программы *');
end;

procedure TFormMain.IniCreateMain;
begin
  if FileExists(AppPath + 'office.ini') then
    exit;
  IniMain := TIniFile.Create(AppPath + 'office.ini');
  { **************** FormMain ***************** }
  IniMain.WriteInteger('main', 'wndstate', 0);
  IniMain.WriteInteger('main', 'wndwidth', 1024);
  IniMain.WriteInteger('main', 'wndheight', 738);
  IniMain.WriteInteger('main', 'panelRubr', 312);
  IniMain.WriteInteger('main', 'loadSGonRubrChange', 0);
  { **************** SGGeneral ***************** }
  IniMain.WriteInteger('sggeneral', 'col_name_vis', 1);
  IniMain.WriteInteger('sggeneral', 'col_name_pos', 2);
  IniMain.WriteInteger('sggeneral', 'col_name_width', 275);
  IniMain.WriteInteger('sggeneral', 'col_curator_vis', 1);
  IniMain.WriteInteger('sggeneral', 'col_curator_pos', 3);
  IniMain.WriteInteger('sggeneral', 'col_curator_width', 150);
  IniMain.WriteInteger('sggeneral', 'col_added_vis', 1);
  IniMain.WriteInteger('sggeneral', 'col_added_pos', 4);
  IniMain.WriteInteger('sggeneral', 'col_added_width', 100);
  IniMain.WriteInteger('sggeneral', 'col_edited_vis', 1);
  IniMain.WriteInteger('sggeneral', 'col_edited_pos', 5);
  IniMain.WriteInteger('sggeneral', 'col_edited_width', 100);
  IniMain.WriteInteger('sggeneral', 'col_web_vis', 1);
  IniMain.WriteInteger('sggeneral', 'col_web_pos', 6);
  IniMain.WriteInteger('sggeneral', 'col_web_width', 100);
  IniMain.WriteInteger('sggeneral', 'col_email_vis', 1);
  IniMain.WriteInteger('sggeneral', 'col_email_pos', 7);
  IniMain.WriteInteger('sggeneral', 'col_email_width', 100);
  IniMain.WriteInteger('sggeneral', 'col_type_vis', 1);
  IniMain.WriteInteger('sggeneral', 'col_type_pos', 8);
  IniMain.WriteInteger('sggeneral', 'col_type_width', 100);
  IniMain.WriteInteger('sggeneral', 'col_fio_vis', 1);
  IniMain.WriteInteger('sggeneral', 'col_fio_pos', 9);
  IniMain.WriteInteger('sggeneral', 'col_fio_width', 150);
  IniMain.WriteInteger('sggeneral', 'col_rubr_vis', 1);
  IniMain.WriteInteger('sggeneral', 'col_rubr_pos', 10);
  IniMain.WriteInteger('sggeneral', 'col_rubr_width', 150);
  IniMain.WriteInteger('sggeneral', 'col_relevance_vis', 1);
  IniMain.WriteInteger('sggeneral', 'col_relevance_pos', 11);
  IniMain.WriteInteger('sggeneral', 'col_relevance_width', 115);
  IniMain.WriteInteger('sggeneral', 'col_sorted_num', 2);
  IniMain.WriteInteger('sggeneral', 'col_sorted_kind', 0);
  { **************** FormMailSender ***************** }
  IniMain.WriteInteger('mailsender', 'fromIndex', -1);
  { **************** FormReport ***************** }
  IniMain.WriteInteger('report', 'cbLocGeneral', 1);
  IniMain.WriteInteger('report', 'cbLocWord', 0);
  IniMain.WriteInteger('report', 'cbLocExcel_List', 0);
  IniMain.WriteInteger('report', 'cbLocExcel_Report', 0);
  IniMain.WriteInteger('report', 'formatDoc0', 1);
  IniMain.WriteInteger('report', 'formatDoc1', 0);
  IniMain.WriteInteger('report', 'formatDoc2', 0);
  IniMain.WriteInteger('report', 'formatDoc3', 0);
  IniMain.WriteInteger('report', 'formatDoc4', 0);
  IniMain.WriteInteger('report', 'formatDoc5', 1);
  IniMain.WriteInteger('report', 'formatDoc6', 1);
  IniMain.WriteInteger('report', 'formatDoc7', 1);
  IniMain.WriteInteger('report', 'formatDoc8', 0);
  IniMain.WriteInteger('report', 'formatDoc9', 0);
  { **************** FormReportSimple ***************** }
  IniMain.WriteInteger('reportsimple', 'cbLocGeneral', 1);
  IniMain.WriteInteger('reportsimple', 'cbLocWord', 0);
  IniMain.Free;
end;

procedure TFormMain.IniLoadMain;
var
  sorted_num, sorted_kind: integer;
begin
  IniMain := TIniFile.Create(AppPath + 'office.ini');
  { **************** FormMain ***************** }
  FormMain.Left := IniMain.ReadInteger('main', 'wndleft', 0);
  FormMain.Top := IniMain.ReadInteger('main', 'wndtop', 0);
  FormMain.Width := IniMain.ReadInteger('main', 'wndwidth', 0);
  FormMain.Height := IniMain.ReadInteger('main', 'wndheight', 0);
  case IniMain.ReadInteger('main', 'wndstate', 0) of
    0:
      FormMain.WindowState := wsNormal;
    1:
      FormMain.WindowState := wsMaximized;
  end;
  panelRubrikator.Width := IniMain.ReadInteger('main', 'panelRubr', 0);
  LoadSGonRubrChange := IniMain.ReadBool('main', 'loadSGonRubrChange', False);
  nLoadSGonRubrChange.Checked := IniMain.ReadBool('main', 'loadSGonRubrChange', False);
  { **************** SGGeneral ***************** }
  SGGeneral.Columns[2].Visible := IniMain.ReadBool('sggeneral', 'col_name_vis', False);
  SGGeneral.Columns[2].Position := IniMain.ReadInteger('sggeneral', 'col_name_pos', 0);
  SGGeneral.Columns[2].Width := IniMain.ReadInteger('sggeneral', 'col_name_width', 0);
  nCol_Name.Checked := IniMain.ReadBool('sggeneral', 'col_name_vis', False);
  SGGeneral.Columns[3].Visible := IniMain.ReadBool('sggeneral', 'col_curator_vis', False);
  SGGeneral.Columns[3].Position := IniMain.ReadInteger('sggeneral', 'col_curator_pos', 0);
  SGGeneral.Columns[3].Width := IniMain.ReadInteger('sggeneral', 'col_curator_width', 0);
  nCol_Curator.Checked := IniMain.ReadBool('sggeneral', 'col_curator_vis', False);
  SGGeneral.Columns[4].Visible := IniMain.ReadBool('sggeneral', 'col_added_vis', False);
  SGGeneral.Columns[4].Position := IniMain.ReadInteger('sggeneral', 'col_added_pos', 0);
  SGGeneral.Columns[4].Width := IniMain.ReadInteger('sggeneral', 'col_added_width', 0);
  nCol_Added.Checked := IniMain.ReadBool('sggeneral', 'col_added_vis', False);
  SGGeneral.Columns[5].Visible := IniMain.ReadBool('sggeneral', 'col_edited_vis', False);
  SGGeneral.Columns[5].Position := IniMain.ReadInteger('sggeneral', 'col_edited_pos', 0);
  SGGeneral.Columns[5].Width := IniMain.ReadInteger('sggeneral', 'col_edited_width', 0);
  nCol_Edited.Checked := IniMain.ReadBool('sggeneral', 'col_edited_vis', False);
  SGGeneral.Columns[6].Visible := IniMain.ReadBool('sggeneral', 'col_web_vis', False);
  SGGeneral.Columns[6].Position := IniMain.ReadInteger('sggeneral', 'col_web_pos', 0);
  SGGeneral.Columns[6].Width := IniMain.ReadInteger('sggeneral', 'col_web_width', 0);
  nCol_WEB.Checked := IniMain.ReadBool('sggeneral', 'col_web_vis', False);
  SGGeneral.Columns[7].Visible := IniMain.ReadBool('sggeneral', 'col_email_vis', False);
  SGGeneral.Columns[7].Position := IniMain.ReadInteger('sggeneral', 'col_email_pos', 0);
  SGGeneral.Columns[7].Width := IniMain.ReadInteger('sggeneral', 'col_email_width', 0);
  nCol_EMAIL.Checked := IniMain.ReadBool('sggeneral', 'col_email_vis', False);
  SGGeneral.Columns[8].Visible := IniMain.ReadBool('sggeneral', 'col_type_vis', False);
  SGGeneral.Columns[8].Position := IniMain.ReadInteger('sggeneral', 'col_type_pos', 0);
  SGGeneral.Columns[8].Width := IniMain.ReadInteger('sggeneral', 'col_type_width', 0);
  nCol_Type.Checked := IniMain.ReadBool('sggeneral', 'col_type_vis', False);
  SGGeneral.Columns[9].Visible := IniMain.ReadBool('sggeneral', 'col_fio_vis', False);
  SGGeneral.Columns[9].Position := IniMain.ReadInteger('sggeneral', 'col_fio_pos', 0);
  SGGeneral.Columns[9].Width := IniMain.ReadInteger('sggeneral', 'col_fio_width', 0);
  nCol_FIO.Checked := IniMain.ReadBool('sggeneral', 'col_fio_vis', False);
  SGGeneral.Columns[10].Visible := IniMain.ReadBool('sggeneral', 'col_rubr_vis', False);
  SGGeneral.Columns[10].Position := IniMain.ReadInteger('sggeneral', 'col_rubr_pos', 0);
  SGGeneral.Columns[10].Width := IniMain.ReadInteger('sggeneral', 'col_rubr_width', 0);
  nCol_Rubr.Checked := IniMain.ReadBool('sggeneral', 'col_rubr_vis', False);
  SGGeneral.Columns[11].Visible := IniMain.ReadBool('sggeneral', 'col_relevance_vis', False);
  SGGeneral.Columns[11].Position := IniMain.ReadInteger('sggeneral', 'col_relevance_pos', 0);
  SGGeneral.Columns[11].Width := IniMain.ReadInteger('sggeneral', 'col_relevance_width', 0);
  nCol_Relevance.Checked := IniMain.ReadBool('sggeneral', 'col_relevance_vis', False);
  sorted_num := IniMain.ReadInteger('sggeneral', 'col_sorted_num', 0);
  sorted_kind := IniMain.ReadInteger('sggeneral', 'col_sorted_kind', 0);
  SGGeneral.Columns[sorted_num].Sorted := True;
  if sorted_kind = 0 then
    SGGeneral.Columns[sorted_num].SortKind := skAscending;
  if sorted_kind = 1 then
    SGGeneral.Columns[sorted_num].SortKind := skDescending;
  IniMain.Free;
end;

procedure TFormMain.IniSaveMain;
var
  i: integer;
begin
  IniMain := TIniFile.Create(AppPath + 'office.ini');
  { **************** FormMain ***************** }
  case FormMain.WindowState of
    wsNormal:
      begin
        IniMain.WriteInteger('main', 'wndstate', 0);
        IniMain.WriteInteger('main', 'wndleft', FormMain.Left);
        IniMain.WriteInteger('main', 'wndtop', FormMain.Top);
        IniMain.WriteInteger('main', 'wndwidth', FormMain.Width);
        IniMain.WriteInteger('main', 'wndheight', FormMain.Height);
      end;
    wsMaximized:
      IniMain.WriteInteger('main', 'wndstate', 1);
  end;
  IniMain.WriteInteger('main', 'panelRubr', panelRubrikator.Width);
  IniMain.WriteBool('main', 'loadSGonRubrChange', nLoadSGonRubrChange.Checked);
  { ******************* SGGeneral ******************* }
  IniMain.WriteBool('sggeneral', 'col_name_vis', SGGeneral.Columns[2].Visible);
  IniMain.WriteInteger('sggeneral', 'col_name_pos', SGGeneral.Columns[2].Position);
  IniMain.WriteInteger('sggeneral', 'col_name_width', SGGeneral.Columns[2].Width);
  IniMain.WriteBool('sggeneral', 'col_curator_vis', SGGeneral.Columns[3].Visible);
  IniMain.WriteInteger('sggeneral', 'col_curator_pos', SGGeneral.Columns[3].Position);
  IniMain.WriteInteger('sggeneral', 'col_curator_width', SGGeneral.Columns[3].Width);
  IniMain.WriteBool('sggeneral', 'col_added_vis', SGGeneral.Columns[4].Visible);
  IniMain.WriteInteger('sggeneral', 'col_added_pos', SGGeneral.Columns[4].Position);
  IniMain.WriteInteger('sggeneral', 'col_added_width', SGGeneral.Columns[4].Width);
  IniMain.WriteBool('sggeneral', 'col_edited_vis', SGGeneral.Columns[5].Visible);
  IniMain.WriteInteger('sggeneral', 'col_edited_pos', SGGeneral.Columns[5].Position);
  IniMain.WriteInteger('sggeneral', 'col_edited_width', SGGeneral.Columns[5].Width);
  IniMain.WriteBool('sggeneral', 'col_web_vis', SGGeneral.Columns[6].Visible);
  IniMain.WriteInteger('sggeneral', 'col_web_pos', SGGeneral.Columns[6].Position);
  IniMain.WriteInteger('sggeneral', 'col_web_width', SGGeneral.Columns[6].Width);
  IniMain.WriteBool('sggeneral', 'col_email_vis', SGGeneral.Columns[7].Visible);
  IniMain.WriteInteger('sggeneral', 'col_email_pos', SGGeneral.Columns[7].Position);
  IniMain.WriteInteger('sggeneral', 'col_email_width', SGGeneral.Columns[7].Width);
  IniMain.WriteBool('sggeneral', 'col_type_vis', SGGeneral.Columns[8].Visible);
  IniMain.WriteInteger('sggeneral', 'col_type_pos', SGGeneral.Columns[8].Position);
  IniMain.WriteInteger('sggeneral', 'col_type_width', SGGeneral.Columns[8].Width);
  IniMain.WriteBool('sggeneral', 'col_fio_vis', SGGeneral.Columns[9].Visible);
  IniMain.WriteInteger('sggeneral', 'col_fio_pos', SGGeneral.Columns[9].Position);
  IniMain.WriteInteger('sggeneral', 'col_fio_width', SGGeneral.Columns[9].Width);
  IniMain.WriteBool('sggeneral', 'col_rubr_vis', SGGeneral.Columns[10].Visible);
  IniMain.WriteInteger('sggeneral', 'col_rubr_pos', SGGeneral.Columns[10].Position);
  IniMain.WriteInteger('sggeneral', 'col_rubr_width', SGGeneral.Columns[10].Width);
  IniMain.WriteBool('sggeneral', 'col_relevance_vis', SGGeneral.Columns[11].Visible);
  IniMain.WriteInteger('sggeneral', 'col_relevance_pos', SGGeneral.Columns[11].Position);
  IniMain.WriteInteger('sggeneral', 'col_relevance_width', SGGeneral.Columns[11].Width);
  for i := 2 to 11 do
    if SGGeneral.Columns[i].Sorted then
    begin
      if SGGeneral.Columns[i].SortKind = skAscending then
        IniMain.WriteInteger('sggeneral', 'col_sorted_kind', 0);
      if SGGeneral.Columns[i].SortKind = skDescending then
        IniMain.WriteInteger('sggeneral', 'col_sorted_kind', 1);
      IniMain.WriteInteger('sggeneral', 'col_sorted_num', i);
      break;
    end;
  { **************** FormMailSender ***************** }
  IniMain.WriteInteger('mailsender', 'fromIndex', FormMailSender.editFrom.ItemIndex);
  { **************** FormReport ***************** }
  IniMain.WriteBool('report', 'cbLocGeneral', FormReport.cbLocGeneral.Checked);
  IniMain.WriteBool('report', 'cbLocWord', FormReport.cbLocWord.Checked);
  IniMain.WriteBool('report', 'cbLocExcel_List', FormReport.cbLocExcel_List.Checked);
  IniMain.WriteBool('report', 'cbLocExcel_Report', FormReport.cbLocExcel_Report.Checked);
  IniMain.WriteBool('report', 'formatDoc0', FormReport.editFormatDoc.Checked[0]);
  IniMain.WriteBool('report', 'formatDoc1', FormReport.editFormatDoc.Checked[1]);
  IniMain.WriteBool('report', 'formatDoc2', FormReport.editFormatDoc.Checked[2]);
  IniMain.WriteBool('report', 'formatDoc3', FormReport.editFormatDoc.Checked[3]);
  IniMain.WriteBool('report', 'formatDoc4', FormReport.editFormatDoc.Checked[4]);
  IniMain.WriteBool('report', 'formatDoc5', FormReport.editFormatDoc.Checked[5]);
  IniMain.WriteBool('report', 'formatDoc6', FormReport.editFormatDoc.Checked[6]);
  IniMain.WriteBool('report', 'formatDoc7', FormReport.editFormatDoc.Checked[7]);
  IniMain.WriteBool('report', 'formatDoc8', FormReport.editFormatDoc.Checked[8]);
  IniMain.WriteBool('report', 'formatDoc9', FormReport.editFormatDoc.Checked[9]);
  { **************** FormReportSimple ***************** }
  IniMain.WriteBool('reportsimple', 'cbLocGeneral', FormReportSimple.cbLocGeneral.Checked);
  IniMain.WriteBool('reportsimple', 'cbLocWord', FormReportSimple.cbLocWord.Checked);
  IniMain.Free;
end;

procedure TFormMain.SGAddRow(Grid: TNextGrid; Activity, Relevance: integer; Name, Curator, DateAdded, DateEdited, Web, Email, fType, id,
  FIO, Rubr: string);
var
  s, Cur, Typ, Rub: string;
begin
  while pos('$', Curator) > 0 do
  begin
    s := copy(Curator, 0, pos('$', Curator));
    delete(Curator, 1, length(s));
    delete(s, 1, 1);
    delete(s, length(s), 1);
    if sgCurator_tmp.FindText(1, s, [soCaseInsensitive, soExactMatch]) then
      Cur := Cur + ', ' + sgCurator_tmp.Cells[0, sgCurator_tmp.SelectedRow];
    if length(Cur) > 0 then
      if Cur[1] = ',' then
        delete(Cur, 1, 2);
  end;
  while pos('$', fType) > 0 do
  begin
    s := copy(fType, 0, pos('$', fType));
    delete(fType, 1, length(s));
    delete(s, 1, 1);
    delete(s, length(s), 1);
    if sgType_tmp.FindText(1, s, [soCaseInsensitive, soExactMatch]) then
      Typ := Typ + ', ' + sgType_tmp.Cells[0, sgType_tmp.SelectedRow];
    if length(Typ) > 0 then
      if Typ[1] = ',' then
        delete(Typ, 1, 2);
  end;
  while pos('$', Rubr) > 0 do
  begin
    s := copy(Rubr, 0, pos('$', Rubr));
    delete(Rubr, 1, length(s));
    delete(s, 1, 1);
    delete(s, length(s), 1);
    if sgRubr_tmp.FindText(1, s, [soCaseInsensitive, soExactMatch]) then
      Rub := Rub + ', ' + sgRubr_tmp.Cells[0, sgRubr_tmp.SelectedRow];
    if length(Rub) > 0 then
      if Rub[1] = ',' then
        delete(Rub, 1, 2);
  end;
  Grid.AddRow;
  Grid.Cells[0, Grid.LastAddedRow] := id;
  if Activity = -1 then
    Grid.Cells[1, Grid.LastAddedRow] := '1'
  else
    Grid.Cells[1, Grid.LastAddedRow] := '2';
  Grid.Cells[2, Grid.LastAddedRow] := Name;
  Grid.Cells[3, Grid.LastAddedRow] := Cur;
  Grid.Cells[4, Grid.LastAddedRow] := DateAdded;
  Grid.Cells[5, Grid.LastAddedRow] := DateEdited;
  Grid.Cells[6, Grid.LastAddedRow] := Web;
  Grid.Cells[7, Grid.LastAddedRow] := Email;
  Grid.Cells[8, Grid.LastAddedRow] := Typ;
  Grid.Cells[9, Grid.LastAddedRow] := FIO;
  Grid.Cells[10, Grid.LastAddedRow] := Rub;
  if Relevance = -1 then
    Grid.Cells[11, Grid.LastAddedRow] := 'Да'
  else
    Grid.Cells[11, Grid.LastAddedRow] := 'Нет';
end;

procedure TFormMain.TVRubrikatorChange(Sender: TObject; Node: TTreeNode);
var
  i: integer;
  // t1, t2 : TDateTime;
begin
  if not LoadSGonRubrChange then
    exit;
  // t1 := now;
  if TVRubrikator.Selected = nil then
    exit;
  // sLabel1.Caption := 'TVRUBR_ID: ' + IntToStr(Integer(Pointer(TVRubrikator.Selected.Data)));
  // if sPageControl1.ActivePageIndex <> 0 then sPageControl1.ActivePageIndex := 0;
  if TVRubrikator.Selected.Level = 0 then // Рубрика
  begin
    IBQuery1.Close;
    IBQuery1.SQL.Text := 'select * from BASE where RUBR like :RUBR';
    IBQuery1.ParamByName('RUBR').AsString := '%#' + IntToStr(integer(Pointer(TVRubrikator.Selected.Data))) + '$%';
    IBQuery1.Open;
    IBQuery1.FetchAll;
    if IBQuery1.RecordCount = 0 then
    begin
      SGGeneral.ClearRows;
      IBQuery1.Close;
      IBDatabase1.Close;
      exit;
    end;
    SGGeneral.BeginUpdate;
    SGGeneral.ClearRows;
    for i := 1 to IBQuery1.RecordCount do
    begin
      SGAddRow(SGGeneral, IBQuery1.FieldByName('ACTIVITY').AsInteger, IBQuery1.FieldByName('RELEVANCE').AsInteger,
        IBQuery1.FieldValues['NAME'], IBQuery1.FieldValues['CURATOR'], IBQuery1.FieldValues['DATE_ADDED'],
        IBQuery1.FieldValues['DATE_EDITED'], IBQuery1.FieldValues['WEB'], IBQuery1.FieldValues['EMAIL'], IBQuery1.FieldValues['TYPE'],
        IBQuery1.FieldValues['ID'], IBQuery1.FieldValues['FIO'], IBQuery1.FieldValues['RUBR']);
      IBQuery1.Next;
    end;
    SGGeneral.Resort;
    SGGeneral.EndUpdate;
  end;
  IBQuery1.Close;
  IBDatabase1.Close;
  { t2 := now;
    sStatusBar1.Panels[2].Text := 'rubrChange: '+IntToStr(MilliSecondsBetween(t1, t2))+' Ms'; }
end;

procedure TFormMain.SGAfterSort(Sender: TObject; ACol: integer);
var
  i, z: integer;
  C: TColor;
begin
  with Sender as TNextGrid do
  begin
    BeginUpdate;
    for i := 0 to RowCount - 1 do
    begin
      if not Odd(i) then
        C := $00EDE9EB { clMenuBar }
      else
        C := clWindow;
      for z := 0 to Columns.Count - 1 do
        Cell[z, i].Color := C;
    end;
    EndUpdate;
  end;
end;

procedure TFormMain.TVRubrikatorDblClick(Sender: TObject);
var
  id: string;
  i: integer;
begin
  if TsTreeView(Sender).Name = 'TVRubrikator' then
  begin
    if TVRubrikator.Selected = nil then
      exit;
    if TVRubrikator.Selected.Level = 0 then
      exit;
    id := IntToStr(integer(Pointer(TVRubrikator.Selected.Data)));
    for i := 0 to sPageControl1.PageCount - 1 do
    begin
      if sPageControl1.Pages[i].Name = 'Tab' + id then
      begin
        sPageControl1.Pages[i].Show;
        exit;
      end;
    end;
  end;
  if TNextGrid(Sender).Name = 'SGGeneral' then
  begin
    if SGGeneral.SelectedCount = 0 then
      exit;
    id := SGGeneral.Cells[0, SGGeneral.SelectedRow];
    for i := 0 to sPageControl1.PageCount - 1 do
    begin
      if sPageControl1.Pages[i].Name = 'Tab' + id then
      begin
        sPageControl1.Pages[i].Show;
        exit;
      end;
    end;
  end;
  OpenTabByID(id);
end;

procedure TFormMain.SGGeneralMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: integer);
begin
  if Button = mbRight then
  begin
    if (X >= 0) and (X <= screen.Width) and (Y >= 0) and (Y <= 22) then
      PopupMenu_Columns.Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y)
    else
      PopupMenu_Rubr.Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
  end;
end;

procedure TFormMain.nCloseTabClick(Sender: TObject);
begin
  if sPageControl1.ActivePageIndex = 0 then
    exit;
  sPageControl1.ActivePage.Free;
end;

procedure TFormMain.nCloseAllButThisTabClick(Sender: TObject);
var
  curr: string;
begin
  if sPageControl1.ActivePageIndex = 0 then
  begin
    while sPageControl1.PageCount > 1 do
      sPageControl1.Pages[1].Free;
    exit;
  end;
  curr := sPageControl1.ActivePage.Name;
  while sPageControl1.PageCount <> 2 do
  begin
    if not (sPageControl1.Pages[1].Name = curr) then
      sPageControl1.Pages[1].Free
    else if sPageControl1.PageCount > 2 then
      sPageControl1.Pages[2].Free;
  end;
  sPageControl1.Pages[1].Show;
end;

procedure TFormMain.nCloseAllTabClick(Sender: TObject);
begin
  if sPageControl1.PageCount = 1 then
    exit;
  while sPageControl1.PageCount > 1 do
    sPageControl1.Pages[1].Free;
end;

procedure TFormMain.ClearSearhcAndFilter;
begin
  editSearchFast.Text := 'Поиск ...';
end;

procedure TFormMain.BtnRecAddClick(Sender: TObject);
begin
  if FormDublicate.Visible then
  begin
    FormDublicate.SetFocus;
    MessageBox(FormDublicate.Handle, 'Необходимо завершить работу с текущей фирмой', 'Информация', MB_OK or MB_ICONINFORMATION);
    exit;
  end;
  WriteLog('TFormMain.BtnRecAddClick: добавление записи');
  FormEditor.Caption := 'Добавить фирму';
  FormEditor.BtnOK.Caption := 'Добавить';
  FormEditor.BtnOK.OnClick := FormEditor.AddRecord;
  FormEditor.ClearEdits;
  Editor.IsDublicate := False;
  FormEditor.Show;
  FormEditor.EditName.SetFocus;
end;

procedure TFormMain.nAddRecClick(Sender: TObject);
begin
  if FormDublicate.Visible then
  begin
    FormDublicate.SetFocus;
    MessageBox(FormDublicate.Handle, 'Необходимо завершить работу с текущей фирмой', 'Информация', MB_OK or MB_ICONINFORMATION);
    exit;
  end;
  WriteLog('TFormMain.nAddRecClick: добавление записи');
  if (TVRubrikator.Focused) and (TVRubrikator.Selected.Level = 0) then
  begin
    FormEditor.Caption := 'Добавить фирму';
    FormEditor.BtnOK.Caption := 'Добавить';
    FormEditor.BtnOK.OnClick := FormEditor.AddRecord;
    FormEditor.ClearEdits;
    FormEditor.SGRubr.AddRow;
    FormEditor.SGRubr.Cells[0, FormEditor.SGRubr.LastAddedRow] := TVRubrikator.Selected.Text;
    FormEditor.SGRubr.Cells[1, FormEditor.SGRubr.LastAddedRow] := IntToStr(integer(TVRubrikator.Selected.Data));
  end
  else
  begin
    FormEditor.Caption := 'Добавить фирму';
    FormEditor.BtnOK.Caption := 'Добавить';
    FormEditor.BtnOK.OnClick := FormEditor.AddRecord;
    FormEditor.ClearEdits;
  end;
  Editor.IsDublicate := False;
  FormEditor.Show;
end;

procedure TFormMain.nEditRecClick(Sender: TObject);
var
  id: string;
begin
  if FormDublicate.Visible then
  begin
    FormDublicate.SetFocus;
    MessageBox(FormDublicate.Handle, 'Необходимо завершить работу с текущей фирмой', 'Информация', MB_OK or MB_ICONINFORMATION);
    exit;
  end;
  if Sender.ClassName = 'TsSpeedButton' then
  begin
    if TsSpeedButton(Sender).Caption = 'Редактировать' then
      id := TsSpeedButton(Sender).Hint;
  end
  else if SGGeneral.Focused then
  begin
    if SGGeneral.SelectedCount = 0 then
      exit;
    id := SGGeneral.Cells[0, SGGeneral.SelectedRow];
  end
  else if TVRubrikator.Focused then
  begin
    if TVRubrikator.Selected = nil then
      exit;
    if TVRubrikator.Selected.Level = 0 then
      exit;
    id := IntToStr(integer(TVRubrikator.Selected.Data));
  end;
  if Trim(id) = '' then
    exit;
  WriteLog('TFormMain.nEditRecClick: редактирование записи ' + id);
  FormEditor.Caption := 'Редактировать фирму';
  FormEditor.BtnOK.Caption := 'Изменить';
  FormEditor.BtnOK.OnClick := FormEditor.EditRecord;
  FormEditor.ClearEdits;
  FormEditor.PrepareEditRecord(id);
  FormEditor.Show;
  FormEditor.EditName.SetFocus;
end;

procedure TFormMain.nRelationsClick(Sender: TObject);
var
  id: string;
begin
  if SGGeneral.Focused then
  begin
    if SGGeneral.SelectedCount = 0 then
      exit;
    id := SGGeneral.Cells[0, SGGeneral.SelectedRow];
  end;
  if TVRubrikator.Focused then
  begin
    if TVRubrikator.Selected = nil then
      exit;
    if TVRubrikator.Selected.Level = 0 then
      exit;
    id := IntToStr(integer(TVRubrikator.Selected.Data));
  end;
  if Trim(id) = '' then
    exit;
  FormRelations.ClearEdits;
  FormRelations.PrepareRecord(id);
  FormRelations.Show;
end;

procedure TFormMain.nDeleteRecClick(Sender: TObject);
var
  id: string;
begin
  if SGGeneral.Focused then
  begin
    if SGGeneral.SelectedCount = 0 then
      exit;
    id := SGGeneral.Cells[0, SGGeneral.SelectedRow];
  end;
  if TVRubrikator.Focused then
  begin
    if TVRubrikator.Selected = nil then
      exit;
    if TVRubrikator.Selected.Level = 0 then
      exit;
    id := IntToStr(integer(TVRubrikator.Selected.Data));
  end;
  WriteLog('TFormMain.nDeleteRecClick: удаление записи ' + id);
  FormEditor.DeleteRecord(id);
end;

procedure TFormMain.nBtnAccountsClick(Sender: TObject);
begin
  if FormMailSender.Visible then
  begin
    FormMailSender.sPageControl.ActivePageIndex := 1;
    FormMailSender.SetFocus;
    FormMailSender.editProfile.SetFocus;
    exit;
  end;
  FormMailSender.Caption := 'Рассылка почты';
  FormMailSender.ClearEdits(True);
  FormMailSender.sPageControl.ActivePageIndex := 1;
  FormMailSender.Show;
  FormMailSender.editProfile.SetFocus;
end;

procedure TFormMain.BtnSendEmailClick(Sender: TObject);
begin
  if (FormMailSender.Visible) and (FormMailSender.Caption = 'Рассылка почты') then
  begin
    FormMailSender.SetFocus;
    exit;
  end;
  FormMailSender.Caption := 'Рассылка почты';
  FormMailSender.ClearEdits(True);
  FormMailSender.Show;
end;

procedure TFormMain.nSendEmailNewClick(Sender: TObject);
var
  id: string;
  Q: TIBQuery;
begin
  if Sender.ClassName = 'TsSpeedButton' then
    if TsSpeedButton(Sender).Caption = 'Рассылка почты' then
    begin
      id := TsSpeedButton(Sender).Hint;
      FormMailSender.Caption := 'Рассылка почты';
      FormMailSender.ClearEdits(True);
      FormMailSender.GetEmailList(1, IBQuery1, id, FormMailSender.editEmailList, True);
      FormMailSender.Show;
      FormMailSender.editTo.SetFocus;
      exit;
    end;
  if SGGeneral.Focused then
  begin
    if SGGeneral.SelectedCount = 0 then
      exit;
    id := SGGeneral.Cells[0, SGGeneral.SelectedRow];
    FormMailSender.Caption := 'Рассылка почты';
    FormMailSender.ClearEdits(True);
    FormMailSender.GetEmailList(1, IBQuery1, id, FormMailSender.editEmailList, True);
    FormMailSender.Show;
    FormMailSender.editTo.SetFocus;
  end;
  if (TVRubrikator.Focused) and (TVRubrikator.Selected <> nil) then
  begin
    if TVRubrikator.Selected.Level = 0 then
    begin
      if not TVRubrikator.Selected.HasChildren then
      begin
        MessageBox(Handle, 'Не было найдено ни одного адреса для рассылки', 'Информация', MB_OK or MB_ICONINFORMATION);
        exit;
      end;
      id := IntToStr(integer(TVRubrikator.Selected.Data));
      Q := TIBQuery.Create(FormMain);
      Q.Database := IBDatabase1;
      Q.Transaction := IBTransaction1;
      Q.Close;
      Q.SQL.Text := 'select * from BASE where RUBR like :RUBR order by lower(NAME)';
      Q.ParamByName('RUBR').AsString := '%#' + id + '$%';
      Q.Open;
      Q.FetchAll;
      FormMailSender.Caption := 'Рассылка почты';
      FormMailSender.ClearEdits(True);
      FormMailSender.GetEmailList(0, Q, '', FormMailSender.editEmailList, True);
      FormMailSender.Show;
      FormMailSender.editTo.SetFocus;
      Q.Close;
      Q.Free;
      IBDatabase1.Close;
    end;
    if TVRubrikator.Selected.Level > 0 then
    begin
      id := IntToStr(integer(TVRubrikator.Selected.Data));
      FormMailSender.Caption := 'Рассылка почты';
      FormMailSender.ClearEdits(True);
      FormMailSender.GetEmailList(1, IBQuery1, id, FormMailSender.editEmailList, True);
      FormMailSender.Show;
      FormMailSender.editTo.SetFocus;
    end;
  end;
end;

procedure TFormMain.nSendEmailAddToClick(Sender: TObject);
var
  id: string;
  Q: TIBQuery;
begin
  if FormMailSender.Caption <> 'Рассылка почты' then
    FormMailSender.ClearEdits(True);
  if SGGeneral.Focused then
  begin
    if SGGeneral.SelectedCount = 0 then
      exit;
    id := SGGeneral.Cells[0, SGGeneral.SelectedRow];
    FormMailSender.Caption := 'Рассылка почты';
    FormMailSender.ClearEdits(False);
    FormMailSender.GetEmailList(1, IBQuery1, id, FormMailSender.editEmailList, False);
    FormMailSender.Show;
    FormMailSender.editTo.SetFocus;
  end;
  if (TVRubrikator.Focused) and (TVRubrikator.Selected <> nil) then
  begin
    if TVRubrikator.Selected.Level = 0 then
    begin
      if not TVRubrikator.Selected.HasChildren then
      begin
        MessageBox(Handle, 'Не было найдено ни одного адреса для рассылки', 'Информация', MB_OK or MB_ICONINFORMATION);
        exit;
      end;
      id := IntToStr(integer(TVRubrikator.Selected.Data));
      Q := TIBQuery.Create(FormMain);
      Q.Database := IBDatabase1;
      Q.Transaction := IBTransaction1;
      Q.Close;
      Q.SQL.Text := 'select * from BASE where RUBR like :RUBR order by lower(NAME)';
      Q.ParamByName('RUBR').AsString := '%#' + id + '$%';
      Q.Open;
      Q.FetchAll;
      FormMailSender.Caption := 'Рассылка почты';
      FormMailSender.ClearEdits(False);
      FormMailSender.GetEmailList(0, Q, '', FormMailSender.editEmailList, False);
      FormMailSender.Show;
      FormMailSender.editTo.SetFocus;
      Q.Close;
      Q.Free;
      IBDatabase1.Close;
    end;
    if TVRubrikator.Selected.Level > 0 then
    begin
      id := IntToStr(integer(TVRubrikator.Selected.Data));
      FormMailSender.Caption := 'Рассылка почты';
      FormMailSender.ClearEdits(False);
      FormMailSender.GetEmailList(1, IBQuery1, id, FormMailSender.editEmailList, False);
      FormMailSender.Show;
      FormMailSender.editTo.SetFocus;
    end;
  end;
end;

procedure TFormMain.nRegInfoCheckClick(Sender: TObject);
var
  id: string;
  Q: TIBQuery;
begin
  if SGGeneral.Focused then
  begin
    if SGGeneral.SelectedCount = 0 then
      exit;
    id := SGGeneral.Cells[0, SGGeneral.SelectedRow];
    FormMailSender.Caption := 'Сверка регистрационных данных';
    FormMailSender.ClearEdits(True);
    FormMailSender.GetFirmList(1, IBQuery1, id, FormMailSender.editEmailList);
    FormMailSender.Show;
  end;
  if (TVRubrikator.Focused) and (TVRubrikator.Selected <> nil) then
  begin
    if TVRubrikator.Selected.Level = 0 then
    begin
      if not TVRubrikator.Selected.HasChildren then
      begin
        MessageBox(Handle, 'Не было найдено ни одной фирмы для сверки данных', 'Информация', MB_OK or MB_ICONINFORMATION);
        exit;
      end;
      id := IntToStr(integer(TVRubrikator.Selected.Data));
      Q := TIBQuery.Create(FormMain);
      Q.Database := IBDatabase1;
      Q.Transaction := IBTransaction1;
      Q.Close;
      Q.SQL.Text := 'select * from BASE where RUBR like :RUBR order by lower(NAME)';
      Q.ParamByName('RUBR').AsString := '%#' + id + '$%';
      Q.Open;
      Q.FetchAll;
      FormMailSender.Caption := 'Сверка регистрационных данных';
      FormMailSender.ClearEdits(True);
      FormMailSender.GetFirmList(0, Q, '', FormMailSender.editEmailList);
      FormMailSender.Show;
      Q.Close;
      Q.Free;
      IBDatabase1.Close;
    end;
    if TVRubrikator.Selected.Level > 0 then
    begin
      id := IntToStr(integer(TVRubrikator.Selected.Data));
      FormMailSender.Caption := 'Сверка регистрационных данных';
      FormMailSender.ClearEdits(True);
      FormMailSender.GetFirmList(1, IBQuery1, id, FormMailSender.editEmailList);
      FormMailSender.Show;
    end;
  end;
end;

procedure TFormMain.nRubrInfoClick(Sender: TObject);
var
  id, msg: string;
  i, noactive: integer;
begin
  if not TVRubrikator.Focused then
    exit;
  if TVRubrikator.Selected = nil then
    exit;
  if TVRubrikator.Selected.Level <> 0 then
    exit;
  id := IntToStr(integer(TVRubrikator.Selected.Data));
  noactive := 0;
  IBQuery1.Close;
  IBQuery1.SQL.Text := 'select ACTIVITY from BASE where RUBR like :RUBR';
  IBQuery1.Params[0].Value := '%#' + id + '$%';
  IBQuery1.Open;
  IBQuery1.FetchAll;
  for i := 1 to IBQuery1.RecordCount do
  begin
    if IBQuery1.FieldValues['ACTIVITY'] = '0' then
      inc(noactive);
    IBQuery1.Next;
  end;
  msg := 'Рубрика: ' + TVRubrikator.Selected.Text + #13 + #13;
  msg := msg + 'Всего фирм: ' + IntToStr(IBQuery1.RecordCount) + #13;
  msg := msg + 'Активных: ' + IntToStr(IBQuery1.RecordCount - noactive) + #13;
  msg := msg + 'Не активных: ' + IntToStr(noactive);
  MessageBox(Handle, PChar(msg), 'Информация о рубрике', MB_OK or MB_ICONINFORMATION);
  IBQuery1.Close;
  IBDatabase1.Close;
end;

procedure TFormMain.PrintTabData(Sender: TObject);
var
  WordApp: OleVariant;
  tc: TWinControl;
  RE: TJvRichEdit;
begin
  if not FormReport.IsWordInstalled then
  begin
    MessageBox(Handle, 'В операционной системе не установлена программа для просмотра файлов Microsoft Office.' + #13 +
      'Установите программу и повторите попытку.', 'Предупреждение', MB_OK or MB_ICONWARNING);
    exit;
  end;
  try
    tc := TsSpeedButton(Sender).Parent;
    tc := tc.Parent;
    try
      WordApp := GetActiveOleObject('Word.Application');
      WordApp.Documents.Close(0);
      WordApp.Quit;
    except
    end;
    RE := tc.Components[1] as (TJvRichEdit);
    RE.Lines.SaveToFile(AppPath + '\tmp.rtf');
    WordApp := CreateOLEObject('Word.Application');
    WordApp.Documents.Add(AppPath + '\tmp.rtf');
    WordApp.Visible := True;
    { ShellExecute(handle,'open',PChar(AppPath+'\tmp.rtf'),nil,nil,SW_SHOWNORMAL); }
  except
    on e: Exception do
    begin
      WriteLog('TFormMain.PrintTabData' + #13 + 'Произошел сбой при генерации отчета' + #13 + e.Message);
      MessageBox(Handle, PChar('Произошел сбой при генерации отчета' + #13 + e.Message), 'Ошибка', MB_OK or MB_ICONERROR);
    end;
  end;
end;

function TFormMain.CloseTabByID(id: string): Boolean;
var
  i: integer;
begin
  Result := False;
  for i := 0 to sPageControl1.PageCount - 1 do
  begin
    if sPageControl1.Pages[i].Name = 'Tab' + id then
    begin
      sPageControl1.Pages[i].Free;
      Result := True;
      break;
    end;
  end;
end;

procedure TFormMain.OpenTabByID(id: string);
var
  ts: TsTabSheet;
  RE: TJvRichEdit;
  Pnl: TsPanel;
  BtnEdit, BtnPrint, BtnSend: TsSpeedButton;
  i, RE_TextLength: integer;
  Rubr, tmp, adres, Cur, city_str, country_str, ofType, zip_str: string;
  phones: WideString;
  Q: TIBQuery;
  list, list2: TStrings;

  procedure AddColoredLine(AText: string; AColor: TColor; AFontSize: integer; AFontName: TFontName; AFontStyle: TFontStyles);
  begin
    RE_TextLength := RE_TextLength + length(AText) + 2;
    RE.SelStart := RE_TextLength - length(AText) - 2;
    // RE.SelStart := Length(RE.Text);
    RE.SelAttributes.Color := AColor;
    RE.SelAttributes.Size := AFontSize;
    RE.SelAttributes.Name := AFontName;
    RE.SelAttributes.Style := AFontStyle;
    RE.Lines.Add(AText);
  end;

begin
  ts := TsTabSheet.Create(sPageControl1);
  ts.Name := 'Tab' + id;
  ts.Caption := '';
  // ts.Visible := True;
  ts.PageControl := sPageControl1;

  Pnl := TsPanel.Create(ts);
  Pnl.Name := 'Pnl' + id;
  Pnl.Caption := '';
  Pnl.Parent := ts;
  Pnl.Height := 30;
  Pnl.Align := alBottom;

  BtnEdit := TsSpeedButton.Create(Pnl);
  BtnEdit.Name := 'BtnEdit' + id;
  BtnEdit.Hint := id;
  BtnEdit.Caption := 'Редактировать';
  BtnEdit.Anchors := [akTop, akRight];
  BtnEdit.Parent := Pnl;
  BtnEdit.Height := 24;
  BtnEdit.Width := 130;
  BtnEdit.Left := Pnl.Width - 140;
  BtnEdit.Top := 3;
  BtnEdit.Images := imgMenus;
  BtnEdit.ImageIndex := 1;
  BtnEdit.OnClick := nEditRecClick;

  BtnPrint := TsSpeedButton.Create(Pnl);
  BtnPrint.Name := 'BtnPrint' + id;
  BtnPrint.Caption := 'Печать';
  BtnPrint.Anchors := [akTop, akRight];
  BtnPrint.Parent := Pnl;
  BtnPrint.Height := 24;
  BtnPrint.Width := 90;
  BtnPrint.Left := BtnEdit.Left - 95;
  BtnPrint.Top := 3;
  BtnPrint.Images := imgMenus;
  BtnPrint.ImageIndex := 4;
  BtnPrint.OnClick := PrintTabData;

  BtnSend := TsSpeedButton.Create(Pnl);
  BtnSend.Name := 'BtnSend' + id;
  BtnSend.Hint := id;
  BtnSend.Caption := 'Рассылка почты';
  BtnSend.Anchors := [akTop, akRight];
  BtnSend.Parent := Pnl;
  BtnSend.Height := 24;
  BtnSend.Width := 120;
  BtnSend.Left := BtnPrint.Left - 125;
  BtnSend.Top := 3;
  BtnSend.Images := imgMenus;
  BtnSend.ImageIndex := 3;
  BtnSend.OnClick := nSendEmailNewClick;

  RE := TJvRichEdit.Create(ts);
  RE.Parent := ts;
  RE.Name := 'RichEdit' + id;
  RE.Align := alClient;
  RE.AutoURLDetect := True;
  RE.OnURLClick := RichEditURLClick;
  RE.ReadOnly := True;
  RE.WordWrap := True;
  RE.BorderWidth := 5;
  RE.ScrollBars := ssVertical;
  RE.Clear;

  IBQuery1.Close;
  IBQuery1.SQL.Text := 'select * from BASE where ID = :ID';
  IBQuery1.ParamByName('ID').AsString := id;
  IBQuery1.Open;
  if IBQuery1.FieldValues['NAME'] <> null then
  begin
    ts.Caption := IBQuery1.FieldValues['NAME'];
    AddColoredLine(IBQuery1.FieldValues['NAME'], clRed, 16, 'Tahoma', []);
    AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);
  end;

  AddColoredLine('Куратор:', clMaroon, 10, 'Tahoma', [fsBold]);
  if IBQuery1.FieldValues['CURATOR'] <> null then
  begin
    Rubr := IBQuery1.FieldValues['CURATOR'];
    Cur := '';
    while pos('$', Rubr) > 0 do
    begin
      tmp := copy(Rubr, 0, pos('$', Rubr));
      delete(Rubr, 1, length(tmp));
      delete(tmp, 1, 1);
      delete(tmp, length(tmp), 1);
      if sgCurator_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
        Cur := Cur + ', ' + sgCurator_tmp.Cells[0, sgCurator_tmp.SelectedRow];
    end;
    if length(Cur) > 0 then
      if Cur[1] = ',' then
        delete(Cur, 1, 2);
    if Trim(Cur) <> '' then
      AddColoredLine(Cur, clWindowText, 10, 'Tahoma', []);
  end;
  AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);

  AddColoredLine('Рубрики:', clMaroon, 10, 'Tahoma', [fsBold]);
  if IBQuery1.FieldValues['RUBR'] <> null then
    Rubr := IBQuery1.FieldValues['RUBR'];
  while pos('$', Rubr) > 0 do
  begin
    tmp := copy(Rubr, 0, pos('$', Rubr));
    delete(Rubr, 1, length(tmp));
    delete(tmp, 1, 1);
    delete(tmp, length(tmp), 1);
    Q := TIBQuery.Create(FormMain);
    Q.Database := IBDatabase1;
    Q.Transaction := IBTransaction1;
    Q.Close;
    Q.SQL.Text := 'select * from RUBRIKATOR where ID = :ID';
    Q.ParamByName('ID').AsString := tmp;
    Q.Open;
    AddColoredLine(Q.FieldValues['NAME'], clWindowText, 10, 'Tahoma', []);
    Q.Close;
    Q.Free;
  end;
  AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);

  AddColoredLine('Направления:', clMaroon, 10, 'Tahoma', [fsBold]);
  if IBQuery1.FieldValues['NAPRAVLENIE'] <> null then
    Rubr := IBQuery1.FieldValues['NAPRAVLENIE'];
  while pos('$', Rubr) > 0 do
  begin
    tmp := copy(Rubr, 0, pos('$', Rubr));
    delete(Rubr, 1, length(tmp));
    delete(tmp, 1, 1);
    delete(tmp, length(tmp), 1);
    Q := TIBQuery.Create(FormMain);
    Q.Database := IBDatabase1;
    Q.Transaction := IBTransaction1;
    Q.Close;
    Q.SQL.Text := 'select * from NAPRAVLENIE where ID = :ID';
    Q.ParamByName('ID').AsString := tmp;
    Q.Open;
    AddColoredLine(Q.FieldValues['NAME'], clWindowText, 10, 'Tahoma', []);
    Q.Close;
    Q.Free;
  end;
  AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);

  if IBQuery1.FieldValues['EMAIL'] <> null then
    if Trim(IBQuery1.FieldValues['EMAIL']) <> '' then
    begin
      AddColoredLine('Эл. почта:', clMaroon, 10, 'Tahoma', [fsBold]);
      AddColoredLine(IBQuery1.FieldValues['EMAIL'], clWindowText, 10, 'Tahoma', []);
      AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);
      { Rubr := IBQuery1.FieldValues['EMAIL'];
        if Length(Rubr) > 0 then
        begin
        if Rubr[Length(Rubr)] <> ',' then Rubr := Rubr + ',';
        while pos(',',Rubr) > 0 do
        begin
        tmp := copy(Rubr ,0 , pos(',',Rubr));
        delete(Rubr, 1 , length(tmp));
        tmp := Trim(tmp);
        delete(tmp,length(tmp),1);
        AddColoredLine(tmp,clBlue,12,'Times New Roman',[fsBold]);
        end;
        end; }
    end;

  if IBQuery1.FieldValues['WEB'] <> null then
    if Trim(IBQuery1.FieldValues['WEB']) <> '' then
    begin
      AddColoredLine('Сайт:', clMaroon, 10, 'Tahoma', [fsBold]);
      AddColoredLine(IBQuery1.FieldValues['WEB'], clWindowText, 10, 'Tahoma', []);
      AddColoredLine('', clNone, 10, 'Times New Roman', [fsBold]);
    end;

  AddColoredLine('Адреса:', clMaroon, 10, 'Tahoma', [fsBold]);
  list := TStringList.Create;
  list2 := TStringList.Create;
  list.Text := IBQuery1.FieldValues['ADRES'];
  phones := IBQuery1.FieldValues['PHONES'];
  for i := 0 to list.Count - 1 do
  begin
    Rubr := list[i];
    { Rubr = #CBAdres$#NUM$#OfficeType$#ZIP$#Street$#Country$#City$ }
    list2.Clear;
    while pos('$', Rubr) > 0 do
    begin
      tmp := copy(Rubr, 0, pos('$', Rubr));
      delete(Rubr, 1, length(tmp));
      delete(tmp, 1, 1);
      delete(tmp, length(tmp), 1);
      list2.Add(tmp);
      // list2[0] = CBAdres; list2[1] = NO; list2[2] = OfficeType; list2[3] = ZIP;
      // list2[4] = Street; list2[5] = Country; list2[6] = City;
    end;

    tmp := copy(phones, 0, pos('$', phones));
    delete(phones, 1, length(tmp));
    tmp := copy(phones, 0, pos('$', phones));
    delete(phones, 1, length(tmp));
    delete(tmp, 1, 1);
    delete(tmp, length(tmp), 1);
    // в TMP сейчас хранятся все телефоны в формате МЕМО

    if list2[0] = '1' then
    begin
      // такая же процедура в Report.GenerateReport и ReportSimple.GenerateReport и Editor.PrepareEdit
      city_str := list2[6];
      if city_str[1] = '^' then
        delete(city_str, 1, 1);
      country_str := list2[5];
      if country_str[1] = '&' then
        delete(country_str, 1, 1);
      ofType := list2[2];
      if ofType[1] = '@' then
        delete(ofType, 1, 1);
      zip_str := list2[3];
      if Trim(ofType) <> '' then
        ofType := FormEditor.GetNameByID('OFFICETYPE', ofType) + ' - ';
      if Trim(zip_str) <> '' then
        zip_str := zip_str + ', ';
      if Trim(country_str) <> '' then
        country_str := FormEditor.GetNameByID('COUNTRY', country_str) + ', ';
      if Trim(city_str) <> '' then
        city_str := FormEditor.GetNameByID('GOROD', city_str) + ', ';
      adres := ofType + zip_str + country_str + city_str + list2[4];
      { officetype - zip, country, city, street }
      if Trim(adres) <> '' then
        AddColoredLine(adres, clWindowText, 10, 'Tahoma', []);
      if Trim(tmp) <> '' then
        AddColoredLine(tmp, clNavy, 10, 'Tahoma', []);
    end;
  end; // for i := 0 to list.Count - 1 do
  list.Free;
  list2.Free;

  RE.SelStart := RE.Perform(EM_LINEINDEX, 0, 0);
  Application.ProcessMessages;
  ts.Visible := True;
  RE.SetFocus;
  ts.Show; // sPageControl1.SetFocus;
  if IBDatabase1.Connected then
    IBDatabase1.Close;
end;

procedure TFormMain.RichEditURLClick(Sender: TObject; const URLText: string; Button: TMouseButton);
begin
  ShellExecute(Handle, 'Open', PChar(URLText), nil, nil, SW_NORMAL);
end;

procedure TFormMain.editSearchFastEnter(Sender: TObject);
begin
  if Trim(editSearchFast.Text) = 'Поиск ...' then
    editSearchFast.Clear;
end;

procedure TFormMain.editSearchFastExit(Sender: TObject);
begin
  if Trim(editSearchFast.Text) = '' then
    editSearchFast.Text := 'Поиск ...';
end;

procedure TFormMain.editSearchFastKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  begin
    Key := #0;
    SearchFast(editSearchFast);
  end;
end;

procedure TFormMain.editSearchFastGlyphClick(Sender: TObject);
begin
  SearchFast(editSearchFast);
end;

procedure TFormMain.SearchFast(Sender: TObject);
var
  SS1: string;
  i: integer;
  QuerySearch: TIBQuery;
begin
  SS1 := AnsiLowerCase(Trim(editSearchFast.Text));
  if (SS1 = 'поиск ...') or (SS1 = '') then
    exit;
  QuerySearch := TIBQuery.Create(FormMain);
  QuerySearch.Database := IBDatabase1;
  QuerySearch.Transaction := IBTransaction1;
  try
    QuerySearch.Close;
    if SS1[1] = '+' then
    begin
      QuerySearch.SQL.Text := 'select * from BASE where PHONES like :PHONES';
      QuerySearch.ParamByName('PHONES').AsString := '%' + SS1 + '%';
    end
    else
    begin
      QuerySearch.SQL.Text := 'select * from BASE where lower(NAME) like :NAME';
      QuerySearch.ParamByName('NAME').AsString := '%' + SS1 + '%';
    end;
    QuerySearch.Open;
    QuerySearch.FetchAll;
    if QuerySearch.RecordCount = 0 then
    begin
      QuerySearch.Close;
      QuerySearch.Free;
      IBDatabase1.Close;
      SGGeneral.ClearRows;
      exit;
    end;
    SGGeneral.BeginUpdate;
    SGGeneral.ClearRows;
    for i := 1 to QuerySearch.RecordCount do
    begin
      SGAddRow(SGGeneral, QuerySearch.FieldByName('ACTIVITY').AsInteger, QuerySearch.FieldByName('RELEVANCE').AsInteger,
        QuerySearch.FieldValues['NAME'], QuerySearch.FieldValues['CURATOR'], QuerySearch.FieldValues['DATE_ADDED'],
        QuerySearch.FieldValues['DATE_EDITED'], QuerySearch.FieldValues['WEB'], QuerySearch.FieldValues['EMAIL'],
        QuerySearch.FieldValues['TYPE'], QuerySearch.FieldValues['ID'], QuerySearch.FieldValues['FIO'], QuerySearch.FieldValues['RUBR']);
      QuerySearch.Next;
    end;
    SGGeneral.Resort;
    SGGeneral.EndUpdate;
    sPageControl1.ActivePageIndex := 0;
  except
    on e: Exception do
    begin
      WriteLog('TFormMain.SearchFast' + #13 + 'Произошел сбой при поиске' + #13 + e.Message);
      MessageBox(Handle, PChar('Произошел сбой при поиске' + #13 + e.Message), 'Ошибка', MB_OK or MB_ICONERROR);
    end;
  end;
  QuerySearch.Close;
  QuerySearch.Free;
  IBDatabase1.Close;
end;

procedure TFormMain.BtnReportSimpleClick(Sender: TObject);
begin
  ReportSimple.Q_GEN_SIMPLE_BREAK := False;
  if FormReportSimple.Visible then
  begin
    FormReportSimple.SetFocus;
    exit;
  end;
  FormReport.IsWordAviable(FormReportSimple.cbLocGeneral, FormReportSimple.cbLocWord);
  FormReportSimple.ClearEdits;
  FormReportSimple.Show;
  FormReportSimple.editFilter.SetFocus;
  FormReportSimple.editFilter.DroppedDown := True;
end;

procedure TFormMain.BtnReportAnvancedClick(Sender: TObject);
begin
  Report.Q_GEN_BREAK := False;
  if FormReport.Visible then
  begin
    FormReport.SetFocus;
    exit;
  end;
  FormReport.IsWordAviable(FormReport.cbLocGeneral, FormReport.cbLocWord);
  FormReport.ClearEdits;
  FormReport.Show;
  FormReport.editSelect1.SetFocus;
end;

procedure TFormMain.BtnDirectoryClick(Sender: TObject);
begin
  WriteLog('TFormMain.BtnDirectoryClick: редактирование директорий');
  DisableAllForms('FormDirectory');
  FormDirectory.Show;
end;

procedure TFormMain.nRubtToReportAdvancedClick(Sender: TObject);
var
  i: integer;
begin
  if TVRubrikator.Selected = nil then
    exit;
  if (TVRubrikator.Focused) and (TVRubrikator.Selected.Level = 0) then
  begin
    if FormReport.Visible then
      FormReport.Close;
    Report.Q_GEN_BREAK := False;
    FormReport.IsWordAviable(FormReport.cbLocGeneral, FormReport.cbLocWord);
    FormReport.ClearEdits;
    FormReport.editSelect1.ItemIndex := 6; { Рубрика }
    FormReport.editSelect1Change(FormReport.editSelect1);
    for i := 0 to FormReport.editSelect2.Items.Count - 1 do
      if integer(FormReport.editSelect2.Items.Objects[i]) = integer(TVRubrikator.Selected.Data) then
      begin
        FormReport.editSelect2.ItemIndex := i;
        break;
      end;
    FormReport.Show;
    FormReport.editSelect1.SetFocus;
  end;
end;

procedure TFormMain.nRubtToReportSimpleClick(Sender: TObject);
var
  i: integer;
begin
  if TVRubrikator.Selected = nil then
    exit;
  if (TVRubrikator.Focused) and (TVRubrikator.Selected.Level = 0) then
  begin
    if FormReportSimple.Visible then
      FormReportSimple.Close;
    ReportSimple.Q_GEN_SIMPLE_BREAK := False;
    FormReport.IsWordAviable(FormReportSimple.cbLocGeneral, FormReportSimple.cbLocWord);
    FormReportSimple.ClearEdits;
    FormReportSimple.editFilter.ItemIndex := 0; { Рубрика }
    for i := 0 to sgRubr_tmp.RowCount - 1 do
    begin
      FormReportSimple.editFilterData.AddItem(sgRubr_tmp.Cells[0, i], Pointer(StrToInt(sgRubr_tmp.Cells[1, i])));
    end;
    for i := 0 to FormReportSimple.editFilterData.Items.Count - 1 do
      if integer(FormReportSimple.editFilterData.Items.Objects[i]) = integer(TVRubrikator.Selected.Data) then
      begin
        FormReportSimple.editFilterData.ItemIndex := i;
        break;
      end;
    FormReportSimple.Show;
    FormReportSimple.editFilterData.SetFocus;
  end;
end;

procedure TFormMain.nCol_NameClick(Sender: TObject);
begin
  if TMenuItem(Sender).Name = 'nCol_Name' then
  begin
    nCol_Name.Checked := not nCol_Name.Checked;
    SGGeneral.Columns[2].Visible := nCol_Name.Checked;
  end
  else if TMenuItem(Sender).Name = 'nCol_Curator' then
  begin
    nCol_Curator.Checked := not nCol_Curator.Checked;
    SGGeneral.Columns[3].Visible := nCol_Curator.Checked;
  end
  else if TMenuItem(Sender).Name = 'nCol_Added' then
  begin
    nCol_Added.Checked := not nCol_Added.Checked;
    SGGeneral.Columns[4].Visible := nCol_Added.Checked;
  end
  else if TMenuItem(Sender).Name = 'nCol_Edited' then
  begin
    nCol_Edited.Checked := not nCol_Edited.Checked;
    SGGeneral.Columns[5].Visible := nCol_Edited.Checked;
  end
  else if TMenuItem(Sender).Name = 'nCol_WEB' then
  begin
    nCol_WEB.Checked := not nCol_WEB.Checked;
    SGGeneral.Columns[6].Visible := nCol_WEB.Checked;
  end
  else if TMenuItem(Sender).Name = 'nCol_EMAIL' then
  begin
    nCol_EMAIL.Checked := not nCol_EMAIL.Checked;
    SGGeneral.Columns[7].Visible := nCol_EMAIL.Checked;
  end
  else if TMenuItem(Sender).Name = 'nCol_Type' then
  begin
    nCol_Type.Checked := not nCol_Type.Checked;
    SGGeneral.Columns[8].Visible := nCol_Type.Checked;
  end
  else if TMenuItem(Sender).Name = 'nCol_FIO' then
  begin
    nCol_FIO.Checked := not nCol_FIO.Checked;
    SGGeneral.Columns[9].Visible := nCol_FIO.Checked;
  end
  else if TMenuItem(Sender).Name = 'nCol_Rubr' then
  begin
    nCol_Rubr.Checked := not nCol_Rubr.Checked;
    SGGeneral.Columns[10].Visible := nCol_Rubr.Checked;
  end
  else if TMenuItem(Sender).Name = 'nCol_Relevance' then
  begin
    nCol_Relevance.Checked := not nCol_Relevance.Checked;
    SGGeneral.Columns[11].Visible := nCol_Relevance.Checked;
  end;
end;

procedure TFormMain.nLoadSGonRubrChangeClick(Sender: TObject);
begin
  nLoadSGonRubrChange.Checked := not nLoadSGonRubrChange.Checked;
  LoadSGonRubrChange := nLoadSGonRubrChange.Checked;
end;

end.
