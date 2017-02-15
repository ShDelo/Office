unit Directory;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, sPageControl, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, StdCtrls, sEdit, Buttons, sSpeedButton, ExtCtrls,
  sPanel, IBC, NxColumns, NxColumnClasses, sComboBoxes;

type
  TFormDirectory = class(TForm)
    PageControl: TsPageControl;
    tabCurator: TsTabSheet;
    tabRubr: TsTabSheet;
    tabFirmType: TsTabSheet;
    tabNapr: TsTabSheet;
    tabOfficeType: TsTabSheet;
    panelCurator: TsPanel;
    btnCuratorCreate: TsSpeedButton;
    btnCuratorEdit: TsSpeedButton;
    btnCuratorDelete: TsSpeedButton;
    SGCurator: TNextGrid;
    panelRubr: TsPanel;
    btnRubrCreate: TsSpeedButton;
    btnRubrEdit: TsSpeedButton;
    btnRubrDelete: TsSpeedButton;
    panelFirmType: TsPanel;
    btnFirmTypeCreate: TsSpeedButton;
    btnFirmTypeEdit: TsSpeedButton;
    btnFirmTypeDelete: TsSpeedButton;
    panelNapr: TsPanel;
    btnNaprCreate: TsSpeedButton;
    btnNaprEdit: TsSpeedButton;
    btnNaprDelete: TsSpeedButton;
    panelOfficeType: TsPanel;
    btnOfficeTypeCreate: TsSpeedButton;
    btnOfficeTypeEdit: TsSpeedButton;
    btnOfficeTypeDelete: TsSpeedButton;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    SGRubr: TNextGrid;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    SGFirmType: TNextGrid;
    NxTextColumn5: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    SGNapr: TNextGrid;
    NxTextColumn7: TNxTextColumn;
    NxTextColumn8: TNxTextColumn;
    SGOfficeType: TNextGrid;
    NxTextColumn9: TNxTextColumn;
    NxTextColumn10: TNxTextColumn;
    tabCountry: TsTabSheet;
    tabCity: TsTabSheet;
    panelCountry: TsPanel;
    btnCountryCreate: TsSpeedButton;
    btnCountryEdit: TsSpeedButton;
    btnCountryDelete: TsSpeedButton;
    SGCountry: TNextGrid;
    NxTextColumn11: TNxTextColumn;
    NxTextColumn12: TNxTextColumn;
    panelCity: TsPanel;
    btnCityCreate: TsSpeedButton;
    btnCityEdit: TsSpeedButton;
    btnCityDelete: TsSpeedButton;
    SGCity: TNextGrid;
    NxTextColumn13: TNxTextColumn;
    NxTextColumn14: TNxTextColumn;
    editCurator: TsEdit;
    editCity: TsEdit;
    editRubr: TsEdit;
    editFirmType: TsEdit;
    editNapr: TsEdit;
    editOfficeType: TsEdit;
    editCountry: TsEdit;
    sTabSheet1: TsTabSheet;
    panelPhoneType: TsPanel;
    btnPhoneTypeCreate: TsSpeedButton;
    btnPhoneTypeEdit: TsSpeedButton;
    btnPhoneTypeDelete: TsSpeedButton;
    editPhoneType: TsEdit;
    SGPhoneType: TNextGrid;
    NxTextColumn15: TNxTextColumn;
    NxTextColumn16: TNxTextColumn;
    tabOblast: TsTabSheet;
    panelOblast: TsPanel;
    btnOblastCreate: TsSpeedButton;
    btnOblastEdit: TsSpeedButton;
    btnOblastDelete: TsSpeedButton;
    editOblast: TsEdit;
    SGOblast: TNextGrid;
    NxTextColumn17: TNxTextColumn;
    NxTextColumn18: TNxTextColumn;
    procedure LoadDataDirectory;
    procedure FormShow(Sender: TObject);
    procedure btnCreateClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure editCuratorChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    function Directory_ADD(DirCode: integer; Values: array of string; var NewRecord_ID: string): boolean;
    function Directory_EDIT(DirCode: integer; ID_Directory: string; Values: array of string): boolean;
  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams); override;
    { Public declarations }
  end;

  TDirectoryData = packed record
    Name1: string[255];
    Name2: string[255];
    ID_oblast: integer;
    constructor Create(Name1, Name2: string; ID_oblast: integer);
  end;

  TDirectoryContainer = class
    SG_Main: Pointer; // temp grids e.g main.sgCurator_tmp
    SG_Directory: Pointer; // SG controls on FormDirectory
    Edit_Directory: Pointer; // editors from FormEditor
    Edit_DirectoryQuery: Pointer; // edits from DirectoryQuery
    constructor Create(DirCode: integer);
  end;

const
  DIR_CODE_TOTAL = 8; // 0-based total number of codes
  DIR_CODE_TO_TABLE: array [0 .. DIR_CODE_TOTAL] of string = ('curator', 'rubrikator', 'type', 'napravlenie', 'officetype', 'country',
    'oblast', 'gorod', 'phonetype');
  DIR_CODE_CURATOR = 0;
  DIR_CODE_RUBRIKA = 1;
  DIR_CODE_FIRMTYPE = 2;
  DIR_CODE_NAPRAVLENIE = 3;
  DIR_CODE_OFFICETYPE = 4;
  DIR_CODE_COUNTRY = 5;
  DIR_CODE_OBLAST = 6;
  DIR_CODE_CITY = 7;
  DIR_CODE_PHONETYPE = 8;

var
  FormDirectory: TFormDirectory;
  isDataEdited: Boolean = False;

implementation

uses Main, Editor, Report, Logo, DirectoryQuery;

{$R *.dfm}
{ TDirectoryData }

constructor TDirectoryData.Create(Name1, Name2: string; ID_oblast: integer);
begin
  self.Name1 := Name1;
  self.Name2 := Name2;
  self.ID_oblast := ID_oblast;
end;

{ TDirectoryContainer }

constructor TDirectoryContainer.Create(DirCode: integer);
begin
  if not(DirCode in [0 .. DIR_CODE_TOTAL]) then
    exit;

  case DirCode of
    DIR_CODE_CURATOR:
      begin
        self.SG_Main := main.sgCurator_tmp;
        self.SG_Directory := FormDirectory.SGCurator;
        self.Edit_Directory := FormEditor.EditCurator;
        self.Edit_DirectoryQuery := nil;
      end;
    DIR_CODE_RUBRIKA:
      begin
        self.SG_Main := main.sgRubr_tmp;
        self.SG_Directory := FormDirectory.SGRubr;
        self.Edit_Directory := FormEditor.EditRubr;
        self.Edit_DirectoryQuery := nil;
      end;
    DIR_CODE_FIRMTYPE:
      begin
        self.SG_Main := main.sgCurator_tmp;
        self.SG_Directory := FormDirectory.SGFirmType;
        self.Edit_Directory := FormEditor.EditType;
        self.Edit_DirectoryQuery := nil;
      end;
    DIR_CODE_NAPRAVLENIE:
      begin
        self.SG_Main := main.sgNapr_tmp;
        self.SG_Directory := FormDirectory.SGNapr;
        self.Edit_Directory := FormEditor.EditNapravlenie;
        self.Edit_DirectoryQuery := nil;
      end;
    DIR_CODE_OFFICETYPE:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGOfficeType;
        self.Edit_Directory := FormEditor.EditOfficeType1;
        self.Edit_DirectoryQuery := nil;
      end;
    DIR_CODE_COUNTRY:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGCountry;
        self.Edit_Directory := FormEditor.EditCountry1;
        self.Edit_DirectoryQuery := nil;
      end;
    DIR_CODE_OBLAST:
      begin
         self.SG_Main := nil;
         self.SG_Directory := FormDirectory.SGOblast;
         self.Edit_Directory := FormEditor.EditOblast1;
         self.Edit_DirectoryQuery := nil;
      end;
    DIR_CODE_CITY:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGCity;
        self.Edit_Directory := FormEditor.EditCity1;
        self.Edit_DirectoryQuery := nil;
      end;
    DIR_CODE_PHONETYPE:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGPhoneType;
        self.Edit_Directory := FormEditor.EditPhoneType1;
        self.Edit_DirectoryQuery := nil;
      end;
  end;
end;

{ TFormDirectory }

procedure TFormDirectory.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.ExStyle := Params.ExStyle or WS_Ex_AppWindow;
end;

procedure TFormDirectory.FormCreate(Sender: TObject);
begin
  LoadDataDirectory;
end;

procedure TFormDirectory.FormShow(Sender: TObject);
begin
  editCurator.Clear;
  editRubr.Clear;
  editFirmType.Clear;
  editNapr.Clear;
  editOfficeType.Clear;
  editCountry.Clear;
  editCity.Clear;
  editPhoneType.Clear;
end;

procedure TFormDirectory.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  // #TODO1: Need to rework data updating without GLOBAL_RELOAD
  WriteLog(Format('TFormDirectory.FormClose: редактирование директорий (isDataEdited = %s)', [BoolToStr(isDataEdited, True)]));
  try
    if isDataEdited then
      FormMain.ReloadDataGlobal(FormDirectory);
  finally
    FormMain.EnableAllForms('');
  end;
end;

procedure TFormDirectory.LoadDataDirectory;
var
  Q_Dir: TIBCQuery;
  i: integer;
begin
  FormLogo.sLabel1.Caption := 'Подключение директорий ...';
  SGCurator.BeginUpdate;
  SGCurator.ClearRows;
  for i := 0 to Main.sgCurator_tmp.RowCount - 1 do
  begin
    SGCurator.AddRow; // КУРАТОРЫ
    SGCurator.Cells[0, SGCurator.LastAddedRow] := Main.sgCurator_tmp.Cells[0, i];
    SGCurator.Cells[1, SGCurator.LastAddedRow] := Main.sgCurator_tmp.Cells[1, i];
  end;
  SGCurator.EndUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  SGRubr.BeginUpdate;
  SGRubr.ClearRows;
  for i := 0 to Main.sgRubr_tmp.RowCount - 1 do
  begin
    SGRubr.AddRow; // РУБРИКИ
    SGRubr.Cells[0, SGRubr.LastAddedRow] := Main.sgRubr_tmp.Cells[0, i];
    SGRubr.Cells[1, SGRubr.LastAddedRow] := Main.sgRubr_tmp.Cells[1, i];
  end;
  SGRubr.EndUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  SGFirmType.BeginUpdate;
  SGFirmType.ClearRows;
  for i := 0 to Main.sgType_tmp.RowCount - 1 do
  begin
    SGFirmType.AddRow; // ТИПЫ ФИРМ
    SGFirmType.Cells[0, SGFirmType.LastAddedRow] := Main.sgType_tmp.Cells[0, i];
    SGFirmType.Cells[1, SGFirmType.LastAddedRow] := Main.sgType_tmp.Cells[1, i];
  end;
  SGFirmType.EndUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  SGNapr.BeginUpdate;
  SGNapr.ClearRows;
  for i := 0 to Main.sgNapr_tmp.RowCount - 1 do
  begin
    SGNapr.AddRow; // НАПРАВЛЕНИЯ
    SGNapr.Cells[0, SGNapr.LastAddedRow] := Main.sgNapr_tmp.Cells[0, i];
    SGNapr.Cells[1, SGNapr.LastAddedRow] := Main.sgNapr_tmp.Cells[1, i];
  end;
  SGNapr.EndUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  Q_Dir := QueryCreate;
  Q_Dir.Close;
  Q_Dir.SQL.Text := 'select * from OFFICETYPE order by lower(NAME)';
  Q_Dir.Open;
  Q_Dir.FetchAll := True;
  SGOfficeType.BeginUpdate;
  SGOfficeType.ClearRows;
  if Q_Dir.RecordCount > 0 then
    for i := 1 to Q_Dir.RecordCount do
    begin
      SGOfficeType.AddRow; // ТИПЫ ОФФИСА
      SGOfficeType.Cells[0, SGOfficeType.LastAddedRow] := Q_Dir.FieldValues['NAME'];
      SGOfficeType.Cells[1, SGOfficeType.LastAddedRow] := Q_Dir.FieldValues['ID'];
      Q_Dir.Next;
    end;
  SGOfficeType.EndUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  Q_Dir.Close;
  Q_Dir.SQL.Text := 'select * from COUNTRY order by lower(NAME)';
  Q_Dir.Open;
  Q_Dir.FetchAll := True;
  SGCountry.BeginUpdate;
  SGCountry.ClearRows;
  if Q_Dir.RecordCount > 0 then
    for i := 1 to Q_Dir.RecordCount do
    begin
      SGCountry.AddRow; // СТРАНЫ
      SGCountry.Cells[0, SGCountry.LastAddedRow] := Q_Dir.FieldValues['NAME'];
      SGCountry.Cells[1, SGCountry.LastAddedRow] := Q_Dir.FieldValues['ID'];
      Q_Dir.Next;
    end;
  SGCountry.EndUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  Q_Dir.Close;
  Q_Dir.SQL.Text := 'select * from GOROD order by lower(NAME)';
  Q_Dir.Open;
  Q_Dir.FetchAll := True;
  SGCity.BeginUpdate;
  SGCity.ClearRows;
  if Q_Dir.RecordCount > 0 then
    for i := 1 to Q_Dir.RecordCount do
    begin
      SGCity.AddRow; // ГОРОДА
      SGCity.Cells[0, SGCity.LastAddedRow] := Q_Dir.FieldValues['NAME'];
      SGCity.Cells[1, SGCity.LastAddedRow] := Q_Dir.FieldValues['ID'];
      Q_Dir.Next;
    end;
  SGCity.EndUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  Q_Dir.Close;
  Q_Dir.SQL.Text := 'select * from PHONETYPE order by lower(NAME)';
  Q_Dir.Open;
  Q_Dir.FetchAll := True;
  SGPhoneType.BeginUpdate;
  SGPhoneType.ClearRows;
  if Q_Dir.RecordCount > 0 then
    for i := 1 to Q_Dir.RecordCount do
    begin
      SGPhoneType.AddRow; // ТИПЫ ТЕЛЕФОНОВ
      SGPhoneType.Cells[0, SGPhoneType.LastAddedRow] := Q_Dir.FieldValues['NAME'];
      SGPhoneType.Cells[1, SGPhoneType.LastAddedRow] := Q_Dir.FieldValues['ID'];
      Q_Dir.Next;
    end;
  SGPhoneType.EndUpdate;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  Q_Dir.Close;
  Q_Dir.Free;
end;

procedure TFormDirectory.editCuratorChange(Sender: TObject);
var
  s, t: string;
  RowVisible: Boolean;
  SG: TNextGrid;
  i: integer;
begin
  SG := nil;
  s := AnsiUpperCase(copy(TsEdit(Sender).Text, 0, TsEdit(Sender).SelStart));
  if TsEdit(Sender).Name = 'editCurator' then
    SG := SGCurator;
  if TsEdit(Sender).Name = 'editRubr' then
    SG := SGRubr;
  if TsEdit(Sender).Name = 'editFirmType' then
    SG := SGFirmType;
  if TsEdit(Sender).Name = 'editNapr' then
    SG := SGNapr;
  if TsEdit(Sender).Name = 'editOfficeType' then
    SG := SGOfficeType;
  if TsEdit(Sender).Name = 'editCountry' then
    SG := SGCountry;
  if TsEdit(Sender).Name = 'editCity' then
    SG := SGCity;
  if TsEdit(Sender).Name = 'editPhoneType' then
    SG := SGPhoneType;
  SG.BeginUpdate;
  for i := 0 to SG.RowCount - 1 do
  begin
    t := AnsiUpperCase(copy(SG.Cell[0, i].AsString, 0, length(s)));
    RowVisible := (s = '') or (t = s);
    SG.RowVisible[i] := RowVisible;
  end;
  SG.EndUpdate;
end;

procedure TFormDirectory.btnCreateClick(Sender: TObject);
var
  NewRecord_ID: string;
  QueryValues: array [0 .. 2] of string;
  DirCode: integer;
  DirContrainer: TDirectoryContainer;
begin
  // #TODO1: ADD/EDIT/DELETE functions need to be separated and need to implement helper function to update every control that stores
  // directory entries (e.g: main.TEMP_SG; editor.edits; directory.SG; directoryQuery.edits). Rubr changes need to be updated manually.
  // This way there should be no need for GLOBAL_RELOAD any more.
  DirCode := TsTabSheet(TsSpeedButton(Sender).Parent.Parent).PageIndex;
  if FormDirectoryQuery.ShowDirectoryQuery(DirCode, 'add', TDirectoryData.Create(EmptyStr, EmptyStr, -1), QueryValues) = True then
  begin
    if not Directory_ADD(DirCode, QueryValues, NewRecord_ID) then
      exit;

    // isDataEdited := True;

    DirContrainer := TDirectoryContainer.Create(DirCode);

    if Assigned(DirContrainer.SG_Main) then
      with TNextGrid(DirContrainer.SG_Main) do
      begin
        AddRow;
        Cells[0, LastAddedRow] := QueryValues[0];
        Cells[1, LastAddedRow] := NewRecord_ID;
      end;

    if Assigned(DirContrainer.SG_Directory) then
      with TNextGrid(DirContrainer.SG_Directory) do
      begin
        AddRow;
        Cells[0, LastAddedRow] := QueryValues[0];
        Cells[1, LastAddedRow] := NewRecord_ID;
      end;

    if Assigned(DirContrainer.Edit_Directory) then
      with TsComboBoxEx(DirContrainer.Edit_Directory) do
      begin
        AddItem(QueryValues[0], TObject(StrToInt(NewRecord_ID)));
      end;

    if Assigned(DirContrainer.Edit_DirectoryQuery) then
      with TsComboBoxEx(DirContrainer.Edit_DirectoryQuery) do
      begin
        AddItem(QueryValues[0], TObject(StrToInt(NewRecord_ID)));
      end;

    DirContrainer.Free;
  end;
end;

procedure TFormDirectory.btnEditClick(Sender: TObject);
var
  s, req, id: string;
  Q_Dir: TIBCQuery;
  grid: TNextGrid;
  QueryValues: array [0 .. 2] of string;
  DirContainer: TDirectoryContainer;
  DirCode: integer;
begin
  grid := nil;
  DirCode := TsTabSheet(TsSpeedButton(Sender).Parent.Parent).PageIndex;
  DirContainer := TDirectoryContainer.Create(DirCode);

  if Assigned(DirContainer.SG_Directory) then
    with TNextGrid(DirContainer.SG_Directory) do
    begin
      if SelectedCount = 0 then
        exit;
      s := Cells[0, SelectedRow];
      id := Cells[1, SelectedRow];
      grid := TNextGrid(DirContainer.SG_Directory);
    end;

  { if FormDirectoryQuery.ShowDirectoryQuery(DirCode, 'edit', TDirectoryData.Create(s, 'unknown', StrToInt(id)),
    QueryValues); }

  if InputQuery('Редактировать', 'Управление директориями', s) then
    if Trim(s) <> '' then
    begin
      s := Trim(s);
      case PageControl.ActivePageIndex of
        0:
          req := 'select ID from CURATOR where lower(NAME) = :NAME';
        1:
          req := 'select ID from RUBRIKATOR where lower(NAME) = :NAME';
        2:
          req := 'select ID from TYPE where lower(NAME) = :NAME';
        3:
          req := 'select ID from NAPRAVLENIE where lower(NAME) = :NAME';
        4:
          req := 'select ID from OFFICETYPE where lower(NAME) = :NAME';
        5:
          req := 'select ID from COUNTRY where lower(NAME) = :NAME';
        6:
          req := 'select ID from GOROD where lower(NAME) = :NAME';
        7:
          req := 'select ID from PHONETYPE where lower(NAME) = :NAME';
      end;
      Q_Dir := QueryCreate;
      Q_Dir.Close;
      Q_Dir.SQL.Text := req;
      Q_Dir.Params[0].AsString := AnsiLowerCase(s);
      Q_Dir.Open; // Q_Dir.FetchAll := True;
      if Q_Dir.RecordCount > 0 then
      begin
        MessageBox(handle, 'Запись с таким названием уже существует', 'Информация', MB_OK or MB_ICONINFORMATION);
        Q_Dir.Close;
        Q_Dir.Free;
        exit;
      end;
      case PageControl.ActivePageIndex of
        0:
          req := 'update CURATOR set NAME = :NAME where ID = :ID';
        1:
          req := 'update RUBRIKATOR set NAME = :NAME where ID = :ID';
        2:
          req := 'update TYPE set NAME = :NAME where ID = :ID';
        3:
          req := 'update NAPRAVLENIE set NAME = :NAME where ID = :ID';
        4:
          req := 'update OFFICETYPE set NAME = :NAME where ID = :ID';
        5:
          req := 'update COUNTRY set NAME = :NAME where ID = :ID';
        6:
          req := 'update GOROD set NAME = :NAME where ID = :ID';
        7:
          req := 'update PHONETYPE set NAME = :NAME where ID = :ID';
      end;
      Q_Dir.Close;
      Q_Dir.SQL.Text := req;
      s := UpperFirst(s);
      Q_Dir.ParamByName('NAME').AsString := s;
      Q_Dir.ParamByName('ID').AsString := id;
      try
        Q_Dir.Execute;
      except
        on E: Exception do
        begin
          WriteLog('TFormDirectory.btnEditClick' + #13 + 'Ошибка: ' + E.Message);
          MessageBox(handle, PChar('Ошибка при редактировании директории.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
          Q_Dir.Close;
          Q_Dir.Free;
          exit;
        end;
      end;
      FormMain.IBTransaction1.CommitRetaining;
      isDataEdited := True;
      if Assigned(grid) then
        grid.Cells[0, grid.SelectedRow] := s;
      Q_Dir.Close;
      Q_Dir.Free;
      DirContainer.Free;
    end;
end;

procedure TFormDirectory.btnDeleteClick(Sender: TObject);

  function CheckInUse(field, match: string; var count: integer): Boolean;
  var
    Q: TIBCQuery;
  begin
    result := False;
    count := 0;
    if (Trim(field) = '') or (Trim(match) = '') then
      exit;
    Q := QueryCreate;
    Q.Close;
    Q.SQL.Text := 'select ID from BASE where lower(' + field + ') like :STR';
    Q.ParamByName('STR').AsString := match;
    Q.Open;
    Q.FetchAll := True;
    if Q.RecordCount > 0 then
    begin
      count := Q.RecordCount;
      result := True;
    end;
    Q.Close;
    Q.Free;
    FormMain.IBDatabase1.Close;
  end;

var
  req, id: string;
  Q_Dir: TIBCQuery;
  grid: TNextGrid;
  c: integer;
begin
  grid := nil;
  if PageControl.ActivePageIndex = 0 then
  begin
    if SGCurator.SelectedCount = 0 then
      exit;
    id := SGCurator.Cells[1, SGCurator.SelectedRow];
    grid := SGCurator;
    if CheckInUse('CURATOR', '%#' + id + '$%', c) then
    begin
      MessageBox(handle, PChar('Удаление выбранного куратора невозможно т.к он курирует ' + IntToStr(c) + ' фирм.'), 'Уведомление',
        MB_OK or MB_ICONWARNING);
      exit;
    end;
  end
  else if PageControl.ActivePageIndex = 1 then
  begin
    if SGRubr.SelectedCount = 0 then
      exit;
    id := SGRubr.Cells[1, SGRubr.SelectedRow];
    grid := SGRubr;
    if CheckInUse('RUBR', '%#' + id + '$%', c) then
    begin
      MessageBox(handle, PChar('Удаление выбранной рубрики невозможно т.к за ней закрепленно ' + IntToStr(c) + ' фирм.'), 'Уведомление',
        MB_OK or MB_ICONWARNING);
      exit;
    end;
  end
  else if PageControl.ActivePageIndex = 2 then
  begin
    if SGFirmType.SelectedCount = 0 then
      exit;
    id := SGFirmType.Cells[1, SGFirmType.SelectedRow];
    grid := SGFirmType;
    if CheckInUse('TYPE', '%#' + id + '$%', c) then
    begin
      MessageBox(handle, PChar('Удаление выбранного типа фирмы невозможно т.к он используется ' + IntToStr(c) + ' фирмами.'), 'Уведомление',
        MB_OK or MB_ICONWARNING);
      exit;
    end;
  end
  else if PageControl.ActivePageIndex = 3 then
  begin
    if SGNapr.SelectedCount = 0 then
      exit;
    id := SGNapr.Cells[1, SGNapr.SelectedRow];
    grid := SGNapr;
    if CheckInUse('NAPRAVLENIE', '%#' + id + '$%', c) then
    begin
      MessageBox(handle, PChar('Удаление выбранного вида деятельности невозможно т.к он используется ' + IntToStr(c) + ' фирмами.'),
        'Уведомление', MB_OK or MB_ICONWARNING);
      exit;
    end;
  end
  else if PageControl.ActivePageIndex = 4 then
  begin
    if SGOfficeType.SelectedCount = 0 then
      exit;
    id := SGOfficeType.Cells[1, SGOfficeType.SelectedRow];
    grid := SGOfficeType;
    if CheckInUse('ADRES', '%#@' + id + '$%', c) then
    begin
      MessageBox(handle, PChar('Удаление выбранного типа адреса невозможно т.к он используется ' + IntToStr(c) + ' фирмами.'),
        'Уведомление', MB_OK or MB_ICONWARNING);
      exit;
    end;
  end
  else if PageControl.ActivePageIndex = 5 then
  begin
    if SGCountry.SelectedCount = 0 then
      exit;
    id := SGCountry.Cells[1, SGCountry.SelectedRow];
    grid := SGCountry;
    if CheckInUse('ADRES', '%#&' + id + '$%', c) then
    begin
      MessageBox(handle, PChar('Удаление выбранной страны невозможно т.к она используется ' + IntToStr(c) + ' фирмами.'), 'Уведомление',
        MB_OK or MB_ICONWARNING);
      exit;
    end;
  end
  else if PageControl.ActivePageIndex = 6 then
  begin
    if SGCity.SelectedCount = 0 then
      exit;
    id := SGCity.Cells[1, SGCity.SelectedRow];
    grid := SGCity;
    if CheckInUse('ADRES', '%#^' + id + '$%', c) then
    begin
      MessageBox(handle, PChar('Удаление выбранного города невозможно т.к он используется ' + IntToStr(c) + ' фирмами.'), 'Уведомление',
        MB_OK or MB_ICONWARNING);
      exit;
    end;
  end
  else if PageControl.ActivePageIndex = 7 then
  begin
    if SGPhoneType.SelectedCount = 0 then
      exit;
    id := SGPhoneType.Cells[1, SGPhoneType.SelectedRow];
    grid := SGPhoneType;
    // ДЛЯ 7ой страницы (ТИПЫ ТЕЛЕФОНОВ) проверка (CheckInUse) не делается
  end;
  if MessageBox(handle, 'Вы уверены что хотите удалить выбранный объект?', 'Удалить', MB_YESNO or MB_ICONQUESTION) = MRYES then
  begin
    case PageControl.ActivePageIndex of
      0:
        req := 'delete from CURATOR where ID = :ID';
      1:
        req := 'delete from RUBRIKATOR where ID = :ID';
      2:
        req := 'delete from TYPE where ID = :ID';
      3:
        req := 'delete from NAPRAVLENIE where ID = :ID';
      4:
        req := 'delete from OFFICETYPE where ID = :ID';
      5:
        req := 'delete from COUNTRY where ID = :ID';
      6:
        req := 'delete from GOROD where ID = :ID';
      7:
        req := 'delete from PHONETYPE where ID = :ID';
    end;
    Q_Dir := QueryCreate;
    Q_Dir.Close;
    Q_Dir.SQL.Text := req;
    Q_Dir.Params[0].AsString := id;
    try
      Q_Dir.Execute;
    except
      on E: Exception do
      begin
        WriteLog('TFormDirectory.btnDeleteClick' + #13 + 'Ошибка: ' + E.Message);
        MessageBox(handle, PChar('Ошибка при удалении директории.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
        Q_Dir.Close;
        Q_Dir.Free;
        exit;
      end;
    end;
    FormMain.IBTransaction1.CommitRetaining;
    isDataEdited := True;
    if Assigned(grid) then
      grid.DeleteRow(grid.SelectedRow);
    Q_Dir.Close;
    Q_Dir.Free;
  end;
end;

function TFormDirectory.Directory_ADD(DirCode: integer; Values: array of string; var NewRecord_ID: string): boolean;
var
  Query: TIBCQuery;
  Name1, Name2, ID_Oblast, SQL_select, SQL_insert: string;
begin
  Result := False;
  NewRecord_ID := EmptyStr;

  if not DirCode in [0 .. DIR_CODE_TOTAL] then
    exit;

  Name1 := Values[0];
  Name2 := Values[1];
  ID_Oblast := Values[2];

  if Trim(Name1) = EmptyStr then
    exit;

  if DirCode in [DIR_CODE_CITY] then
  begin
    SQL_select := Format('select ID from %s where (lower(NAME) = :NAME and ID_OBLAST = :ID_OBLAST)', [DIR_CODE_TO_TABLE[DirCode]]);
    SQL_insert := Format('insert into %s (NAME, NAME_ALT, ID_OBLAST) values (:NAME, :NAME_ALT, :ID_OBLAST) returning ID',
      [DIR_CODE_TO_TABLE[DirCode]])
  end
  else
  begin
    SQL_select := Format('select ID from %s where lower(NAME) = :NAME', [DIR_CODE_TO_TABLE[DirCode]]);
    SQL_insert := Format('insert into %s (NAME) values (:NAME) returning ID', [DIR_CODE_TO_TABLE[DirCode]]);
  end;

  // Check if record already exist
  Query := QueryCreate;
  Query.SQL.Text := SQL_select;
  Query.ParamByName('NAME').AsString := AnsiLowerCase(Name1);
  if DirCode in [DIR_CODE_CITY] then
    Query.ParamByName('ID_OBLAST').AsString := ID_Oblast;
  Query.Open;
  if Query.RecordCount > 0 then
  begin
    MessageBox(handle, 'Запись с таким названием уже существует', 'Информация', MB_OK or MB_ICONINFORMATION);
    Query.Close;
    Query.Free;
    exit;
  end;

  // insert new record
  Query.Close;
  Query.Params.Clear;
  Query.SQL.Text := SQL_insert;
  Query.ParamByName('NAME').AsString := Name1;
  if DirCode in [DIR_CODE_CITY] then
  begin
    Query.ParamByName('NAME_ALT').AsString := Name2;
    Query.ParamByName('ID_OBLAST').AsString := ID_Oblast;
  end;
  try
    Query.Execute;

    NewRecord_ID := Query.ParamByName('RET_ID').AsString;
  except
    on E: Exception do
    begin
      WriteLog('TFormDirectory.Directory_ADD' + #13 + 'Ошибка: ' + E.Message);
      MessageBox(handle, PChar('Ошибка при создании директории.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      Query.Close;
      Query.Free;
      exit;
    end;
  end;

  FormMain.IBTransaction1.CommitRetaining;
  Query.Close;
  Query.Free;

  Result := True;
end;

// #TODO1: IMPORTANT: need to think of a way to update opened elements that using directory entries while they being updated.
// say I have add firm dialog opened with already entered city and then I decided to update that city from directory editor window...?
// also opened tabs currently doesn't get updated even by GLOBAL_RELOAD.
function TFormDirectory.Directory_EDIT(DirCode: integer; ID_Directory: string; Values: array of string): boolean;
var
  Query: TIBCQuery;
  Name1, Name2, ID_Oblast, SQL_select, SQL_update: string;
begin
  Result := False;

  if not DirCode in [0 .. DIR_CODE_TOTAL] then
    exit;

  if Trim(ID_Directory) = EmptyStr then
    exit;

  Name1 := Values[0];
  Name2 := Values[1];
  ID_Oblast := Values[2];

  if Trim(Name1) = EmptyStr then
    exit;

  if DirCode in [DIR_CODE_CITY] then
  begin
    SQL_select := Format('select ID from %s where (lower(NAME) = :NAME and ID_OBLAST = :ID_OBLAST)', [DIR_CODE_TO_TABLE[DirCode]]);
    SQL_update := Format('update %s set NAME = :NAME, NAME_ALT = :NAME_ALT, ID_OBLAST = :ID_OBLAST where ID = :ID',
      [DIR_CODE_TO_TABLE[DirCode]]);
  end
  else
  begin
    SQL_select := Format('select ID from %s where lower(NAME) = :NAME', [DIR_CODE_TO_TABLE[DirCode]]);
    SQL_update := Format('update %s set NAME = :NAME where ID = :ID', [DIR_CODE_TO_TABLE[DirCode]]);
  end;

  // Check if record already exist
  Query := QueryCreate;
  Query.SQL.Text := SQL_select;
  Query.ParamByName('NAME').AsString := AnsiLowerCase(Name1);
  if DirCode in [DIR_CODE_CITY] then
    Query.ParamByName('ID_OBLAST').AsString := ID_Oblast;
  Query.Open;
  if Query.RecordCount > 0 then
  begin
    MessageBox(handle, 'Запись с таким названием уже существует', 'Информация', MB_OK or MB_ICONINFORMATION);
    Query.Close;
    Query.Free;
    exit;
  end;

  // update record
  Query.Close;
  Query.Params.Clear;
  Query.SQL.Text := SQL_update;
  Query.ParamByName('NAME').AsString := Name1;
  Query.ParamByName('ID').AsString := ID_Directory;
  if DirCode in [DIR_CODE_CITY] then
  begin
    Query.ParamByName('NAME_ALT').AsString := Name2;
    Query.ParamByName('ID_OBLAST').AsString := ID_Oblast;
  end;
  try
    Query.Execute;
  except
    on E: Exception do
    begin
      WriteLog('TFormDirectory.Directory_EDIT' + #13 + 'Ошибка: ' + E.Message);
      MessageBox(handle, PChar('Ошибка при редактировании директории.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      Query.Close;
      Query.Free;
      exit;
    end;
  end;

  FormMain.IBTransaction1.CommitRetaining;
  Query.Close;
  Query.Free;

  Result := True;
end;

end.
