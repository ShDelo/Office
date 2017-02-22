{ #TODO1: PLAN :
  1. DONE: Implement rubrik directory creation live update
  2. Implement region > city edit control interaction
  3. Implement id_region updating when city edited to new region via directory window
  4. Implement edit controls behavior when entered data is not in the list (not valid values for adres edit controls }
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
    tabRegion: TsTabSheet;
    panelRegion: TsPanel;
    btnRegionCreate: TsSpeedButton;
    btnRegionEdit: TsSpeedButton;
    btnRegionDelete: TsSpeedButton;
    editRegion: TsEdit;
    SGRegion: TNextGrid;
    NxTextColumn17: TNxTextColumn;
    NxTextColumn18: TNxTextColumn;
    NxTextColumn19: TNxTextColumn;
    NxTextColumn20: TNxTextColumn;
    NxTextColumn21: TNxTextColumn;
    procedure LoadDataDirectory;
    procedure FormShow(Sender: TObject);
    procedure btnCreateClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure editCuratorChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    function Directory_CREATE(DirCode: integer; Values: array of string): integer;
    function Directory_EDIT(DirCode: integer; ID_Directory: string; Values: array of string): integer;
    function Directory_DELETE(DirCode: integer; ID_Directory: string): integer;
  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams); override;
    { Public declarations }
  end;

  TDirectoryData = packed record
    Name1: string[255];
    Name2: string[255];
    ID_Region: integer;
    constructor Create(Name1, Name2: string; ID_Region: integer);
  end;

  TDirectoryContainer = class
    SG_Main: Pointer; // temp grid e.g main.sgCurator_tmp
    SG_Directory: Pointer; // SG control on FormDirectory
    Edit_Editor: Pointer; // edit on FormEditor
    Edit_DirectoryQuery: Pointer; // edit on DirectoryQuery
    IsAdresEdits: Boolean;
    constructor Create(DirCode: integer);
  end;

const
  DIR_CODE_TOTAL = 8; // 0-based total number of codes
  DIR_CODE_TO_TABLE: array [0 .. DIR_CODE_TOTAL] of string = ('curator', 'rubrikator', 'firmtype', 'napravlenie', 'officetype', 'country',
    'region', 'city', 'phonetype');
  DIR_CODE_CURATOR = 0;
  DIR_CODE_RUBRIKA = 1;
  DIR_CODE_FIRMTYPE = 2;
  DIR_CODE_NAPRAVLENIE = 3;
  DIR_CODE_OFFICETYPE = 4;
  DIR_CODE_COUNTRY = 5;
  DIR_CODE_REGION = 6;
  DIR_CODE_CITY = 7;
  DIR_CODE_PHONETYPE = 8;

var
  FormDirectory: TFormDirectory;

implementation

uses Main, Editor, Report, Logo, DirectoryQuery, Helpers;

{$R *.dfm}
{ TDirectoryData }

constructor TDirectoryData.Create(Name1, Name2: string; ID_Region: integer);
begin
  self.Name1 := Name1;
  self.Name2 := Name2;
  self.ID_Region := ID_Region;
end;

{ TDirectoryContainer }

constructor TDirectoryContainer.Create(DirCode: integer);
begin
  if not(DirCode in [0 .. DIR_CODE_TOTAL]) then
  begin
    self.SG_Main := nil;
    self.SG_Directory := nil;
    self.Edit_Editor := nil;
    self.Edit_DirectoryQuery := nil;
    self.IsAdresEdits := False;
  end;

  case DirCode of
    DIR_CODE_CURATOR:
      begin
        self.SG_Main := main.sgCurator_tmp;
        self.SG_Directory := FormDirectory.SGCurator;
        self.Edit_Editor := FormEditor.EditCurator;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := False;
      end;
    DIR_CODE_RUBRIKA:
      begin
        self.SG_Main := main.sgRubr_tmp;
        self.SG_Directory := FormDirectory.SGRubr;
        self.Edit_Editor := FormEditor.EditRubr;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := False;
      end;
    DIR_CODE_FIRMTYPE:
      begin
        self.SG_Main := main.sgFirmType_tmp;
        self.SG_Directory := FormDirectory.SGFirmType;
        self.Edit_Editor := FormEditor.EditFirmType;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := False;
      end;
    DIR_CODE_NAPRAVLENIE:
      begin
        self.SG_Main := main.sgNapr_tmp;
        self.SG_Directory := FormDirectory.SGNapr;
        self.Edit_Editor := FormEditor.EditNapravlenie;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := False;
      end;
    DIR_CODE_OFFICETYPE:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGOfficeType;
        self.Edit_Editor := FormEditor.EditOfficeType1;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := True;
      end;
    DIR_CODE_COUNTRY:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGCountry;
        self.Edit_Editor := FormEditor.EditCountry1;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := True;
      end;
    DIR_CODE_REGION:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGRegion;
        self.Edit_Editor := FormEditor.EditRegion1;
        self.Edit_DirectoryQuery := FormDirectoryQuery.Edit3;
        self.IsAdresEdits := True;
      end;
    DIR_CODE_CITY:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGCity;
        self.Edit_Editor := FormEditor.EditCity1;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := True;
      end;
    DIR_CODE_PHONETYPE:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGPhoneType;
        self.Edit_Editor := FormEditor.EditPhoneType1;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := True;
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
  editRegion.Clear;
  editCity.Clear;
  editPhoneType.Clear;
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
  for i := 0 to Main.sgFirmType_tmp.RowCount - 1 do
  begin
    SGFirmType.AddRow; // ТИПЫ ФИРМ
    SGFirmType.Cells[0, SGFirmType.LastAddedRow] := Main.sgFirmType_tmp.Cells[0, i];
    SGFirmType.Cells[1, SGFirmType.LastAddedRow] := Main.sgFirmType_tmp.Cells[1, i];
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
  try
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
    Q_Dir.SQL.Text := 'select * from REGION order by lower(NAME)';
    Q_Dir.Open;
    Q_Dir.FetchAll := True;
    SGRegion.BeginUpdate;
    SGRegion.ClearRows;
    if Q_Dir.RecordCount > 0 then
      for i := 1 to Q_Dir.RecordCount do
      begin
        SGRegion.AddRow; // ОБЛАСТИ
        SGRegion.Cells[0, SGRegion.LastAddedRow] := Q_Dir.FieldValues['NAME'];
        SGRegion.Cells[1, SGRegion.LastAddedRow] := Q_Dir.FieldValues['ID'];
        Q_Dir.Next;
      end;
    SGRegion.EndUpdate;
    FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
    Application.ProcessMessages;

    Q_Dir.Close;
    Q_Dir.SQL.Text := 'select g.ID, g.NAME, g.NAME_ALT, g.ID_REGION, ' +
      '(select o.NAME as REGION_NAME from REGION o where g.ID_REGION = o.ID) from CITY g order by lower(NAME)';
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
        SGCity.Cells[2, SGCity.LastAddedRow] := VarToStr(Q_Dir.FieldValues['NAME_ALT']);
        SGCity.Cells[3, SGCity.LastAddedRow] := VarToStr(Q_Dir.FieldValues['REGION_NAME']);
        SGCity.Cells[4, SGCity.LastAddedRow] := Q_Dir.FieldValues['ID_REGION'];
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

  finally
    Q_Dir.Free;
  end;
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
  if TsEdit(Sender).Name = 'editRegion' then
    SG := SGRegion;
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
  QueryValues: array [0 .. 2] of string;
  DirCode: integer;
begin
  // #TODO1: ADD/EDIT/DELETE functions need to be separated and need to implement helper function to update every control that stores
  // directory entries (e.g: main.TEMP_SG; editor.edits; directory.SG; directoryQuery.edits). Rubr changes need to be updated manually.
  // This way there should be no need for GLOBAL_RELOAD any more.
  DirCode := TsTabSheet(TsSpeedButton(Sender).Parent.Parent).PageIndex;
  if FormDirectoryQuery.ShowDirectoryQuery(DirCode, 'add', TDirectoryData.Create(EmptyStr, EmptyStr, -1), QueryValues) = True then
    Directory_CREATE(DirCode, QueryValues);
end;

procedure TFormDirectory.btnEditClick(Sender: TObject);
var
  Name_First, Name_ALT, ID_Region, ID_City: string;
  QueryValues: array [0 .. 2] of string;
  DirContainer: TDirectoryContainer;
  DirCode: integer;
begin
  DirCode := TsTabSheet(TsSpeedButton(Sender).Parent.Parent).PageIndex;
  DirContainer := TDirectoryContainer.Create(DirCode);
  try
    if Assigned(DirContainer.SG_Directory) then
      with TNextGrid(DirContainer.SG_Directory) do
      begin
        if SelectedCount = 0 then
          exit;
        Name_First := Cells[0, SelectedRow];
        ID_City := Cells[1, SelectedRow];
        if DirCode in [DIR_CODE_CITY] then
        begin
          Name_ALT := Cells[2, SelectedRow];
          ID_Region := Cells[4, SelectedRow];
        end;
      end;

    if FormDirectoryQuery.ShowDirectoryQuery(DirCode, 'edit', TDirectoryData.Create(Name_First, Name_ALT, StrToIntDef(ID_Region, -1)),
      QueryValues) = True then
      Directory_EDIT(DirCode, ID_City, QueryValues);
  finally
    DirContainer.Free;
  end;
end;

procedure TFormDirectory.btnDeleteClick(Sender: TObject);
var
  ID_Directory: string;
  DirContainer: TDirectoryContainer;
  DirCode: integer;
begin
  DirCode := TsTabSheet(TsSpeedButton(Sender).Parent.Parent).PageIndex;
  DirContainer := TDirectoryContainer.Create(DirCode);
  try
    if Assigned(DirContainer.SG_Directory) then
      with TNextGrid(DirContainer.SG_Directory) do
      begin
        if SelectedCount = 0 then
          exit;
        ID_Directory := Cells[1, SelectedRow];
      end;
    Directory_DELETE(DirCode, ID_Directory);
  finally
    DirContainer.Free;
  end;
end;

{ #TODO3: LOG : write log for directory create/edit/delete actions }
function TFormDirectory.Directory_CREATE(DirCode: integer; Values: array of string): integer;
var
  Query: TIBCQuery;
  EditControlName, Name1, Name2, ID_Region, SQL_select, SQL_insert, ID_NewRecord: string;
  i: integer;
  DirContainer: TDirectoryContainer;
  New_Node: TTreeNode;
begin
  Result := -1;

  if not DirCode in [0 .. DIR_CODE_TOTAL] then
    exit;

  { DATABASE INSERT part }

  Name1 := Values[0];
  Name2 := Values[1];
  ID_Region := Values[2];

  if Name1 = EmptyStr then
    exit;

  if DirCode in [DIR_CODE_CITY] then
  begin
    SQL_select := Format('select ID from %s where (lower(NAME) = :NAME and ID_REGION = :ID_REGION)', [DIR_CODE_TO_TABLE[DirCode]]);
    SQL_insert := Format('insert into %s (NAME, NAME_ALT, ID_REGION) values (:NAME, :NAME_ALT, :ID_REGION) returning ID',
      [DIR_CODE_TO_TABLE[DirCode]])
  end
  else
  begin
    SQL_select := Format('select ID from %s where lower(NAME) = :NAME', [DIR_CODE_TO_TABLE[DirCode]]);
    SQL_insert := Format('insert into %s (NAME) values (:NAME) returning ID', [DIR_CODE_TO_TABLE[DirCode]]);
  end;

  Query := QueryCreate;
  try
    // Check if record already exist
    Query.SQL.Text := SQL_select;
    Query.ParamByName('NAME').AsString := AnsiLowerCase(Name1);
    if DirCode in [DIR_CODE_CITY] then
      Query.ParamByName('ID_REGION').AsString := ID_Region;
    Query.Open;
    if Query.RecordCount > 0 then
    begin
      MessageBox(handle, 'Запись с таким названием уже существует', 'Информация', MB_OK or MB_ICONINFORMATION);
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
      Query.ParamByName('ID_REGION').AsString := ID_Region;
    end;
    try
      Query.Execute;
      Query.Transaction.CommitRetaining;

      ID_NewRecord := Query.ParamByName('RET_ID').AsString;
    except
      on E: Exception do
      begin
        Query.Transaction.RollbackRetaining;
        WriteLog('TFormDirectory.Directory_ADD' + #13 + 'Ошибка: ' + E.Message);
        MessageBox(handle, PChar('Ошибка при создании директории.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
        exit;
      end;
    end;
  finally
    Query.Free;
  end;

  { VISUAL DATA UPDATE part }

  DirContainer := TDirectoryContainer.Create(DirCode);
  try
    if Assigned(DirContainer.SG_Main) then
      with TNextGrid(DirContainer.SG_Main) do
      begin
        AddRow;
        Cells[0, LastAddedRow] := Name1;
        Cells[1, LastAddedRow] := ID_NewRecord;
      end;

    if Assigned(DirContainer.SG_Directory) then
      with TNextGrid(DirContainer.SG_Directory) do
      begin
        AddRow;
        Cells[0, LastAddedRow] := Name1;
        Cells[1, LastAddedRow] := ID_NewRecord;
        if DirCode in [DIR_CODE_CITY] then
        begin
          Cells[2, LastAddedRow] := Name2; // NAME_ALT
          Cells[3, LastAddedRow] := GetNameByID(DIR_CODE_TO_TABLE[DIR_CODE_REGION], ID_Region); // REGION_TEXT
          Cells[4, LastAddedRow] := ID_REGION; // ID_REGION
        end;
        SelectLastRow;
      end;

    if Assigned(DirContainer.Edit_Editor) then
    begin
      if DirContainer.IsAdresEdits = True then
      begin
        EditControlName := TsComboBoxEx(DirContainer.Edit_Editor).Name;
        delete(EditControlName, Length(EditControlName), 1);
        for i := 1 to 10 do
          TsComboBoxEx(FormEditor.FindComponent(EditControlName + IntToStr(i))).AddItem(Name1, TObject(ID_NewRecord.ToInteger));
      end
      else
        with TsComboBoxEx(DirContainer.Edit_Editor) do
          AddItem(Name1, TObject(ID_NewRecord.ToInteger));
    end;

    if Assigned(DirContainer.Edit_DirectoryQuery) then
      with TsComboBoxEx(DirContainer.Edit_DirectoryQuery) do
        AddItem(Name1, TObject(ID_NewRecord.ToInteger));

    // update rubrikator_tree if needed
    if DirCode = DIR_CODE_RUBRIKA then
    begin
      New_Node := FormMain.TVRubrikator.Items.AddChildObject(nil, Name1, Pointer(ID_NewRecord.ToInteger));
      New_Node.ImageIndex := 0;
      New_Node.SelectedIndex := 0;
      FormMain.TVRubrikator.CustomSort(@Helpers.CustomSortProc, 0, True);
    end;
  finally
    DirContainer.Free;
  end;

  Result := StrToIntDef(ID_NewRecord, -1);
end;

// #TODO1: IMPORTANT: need to think of a way to update opened elements that using directory entries while they being updated.
// say I have add firm dialog opened with already entered city and then I decided to update that city from directory editor window...?
// also opened tabs currently doesn't get updated even by GLOBAL_RELOAD.
function TFormDirectory.Directory_EDIT(DirCode: integer; ID_Directory: string; Values: array of string): integer;
var
  Query: TIBCQuery;
  Name1, Name2, ID_Region, SQL_select, SQL_update, EditControlName: string;
  i: integer;
  DirContainer: TDirectoryContainer;
  EditControl: TsComboBoxEx;
begin
  Result := -1;

  if not DirCode in [0 .. DIR_CODE_TOTAL] then
    exit;

  if Trim(ID_Directory) = EmptyStr then
    exit;

  { DATABASE UPDATE part }

  { #TODO1: IMPORTANT : if during city editing region_id changes then we have to update each record from BASE that used this city
    <select ADRES from BASE where ADRES like '%#^id$%'> <while not Q.Eof update with new id_region> }

  Name1 := Values[0];
  Name2 := Values[1];
  ID_Region := Values[2];

  if Trim(Name1) = EmptyStr then
    exit;

  if DirCode in [DIR_CODE_CITY] then
  begin
    SQL_select := Format('select ID from %s where (lower(NAME) = :NAME and ID_REGION = :ID_REGION)', [DIR_CODE_TO_TABLE[DirCode]]);
    SQL_update := Format('update %s set NAME = :NAME, NAME_ALT = :NAME_ALT, ID_REGION = :ID_REGION where ID = :ID',
      [DIR_CODE_TO_TABLE[DirCode]]);
  end
  else
  begin
    SQL_select := Format('select ID from %s where lower(NAME) = :NAME', [DIR_CODE_TO_TABLE[DirCode]]);
    SQL_update := Format('update %s set NAME = :NAME where ID = :ID', [DIR_CODE_TO_TABLE[DirCode]]);
  end;

  Query := QueryCreate;
  try
    // Check if record already exist
    Query.SQL.Text := SQL_select;
    Query.ParamByName('NAME').AsString := AnsiLowerCase(Name1);
    if DirCode in [DIR_CODE_CITY] then
      Query.ParamByName('ID_REGION').AsString := ID_Region;
    Query.Open;
    if Query.RecordCount > 0 then
    begin
      MessageBox(handle, 'Запись с таким названием уже существует', 'Информация', MB_OK or MB_ICONINFORMATION);
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
      Query.ParamByName('ID_REGION').AsString := ID_Region;
    end;
    try
      Query.Execute;
      Query.Transaction.CommitRetaining;
    except
      on E: Exception do
      begin
        Query.Transaction.RollbackRetaining;
        WriteLog('TFormDirectory.Directory_EDIT' + #13 + 'Ошибка: ' + E.Message);
        MessageBox(handle, PChar('Ошибка при редактировании директории.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
        exit;
      end;
    end;
  finally
    Query.Free;
  end;

  { VISUAL DATA UPDATE part }

  DirContainer := TDirectoryContainer.Create(DirCode);
  try
    if Assigned(DirContainer.SG_Main) then
      with TNextGrid(DirContainer.SG_Main) do
      begin
        if FindText(1, ID_Directory, [soCaseInsensitive, soExactMatch]) then
          Cells[0, SelectedRow] := Name1;
      end;

    if Assigned(DirContainer.SG_Directory) then
      with TNextGrid(DirContainer.SG_Directory) do
      begin
        if FindText(1, ID_Directory, [soCaseInsensitive, soExactMatch]) then
        begin
          Cells[0, SelectedRow] := Name1;
          if DirCode in [DIR_CODE_CITY] then
          begin
            Cells[2, SelectedRow] := Name2; // NAME_ALT
            Cells[3, SelectedRow] := GetNameByID(DIR_CODE_TO_TABLE[DIR_CODE_REGION], ID_Region); // REGION_TEXT
            Cells[4, SelectedRow] := ID_Region; // ID_REGION
          end;
        end;
      end;

    if Assigned(DirContainer.Edit_Editor) then
    begin
      if DirContainer.IsAdresEdits = True then
      begin
        EditControlName := TsComboBoxEx(DirContainer.Edit_Editor).Name;
        delete(EditControlName, Length(EditControlName), 1);
        for i := 1 to 10 do
        begin
          EditControl := TsComboBoxEx(FormEditor.FindComponent(EditControlName + IntToStr(i)));
          EditControl.Items[EditControl.GetIndexOfObject(ID_Directory.ToInteger)] := Name1;
        end;
      end
      else
        with TsComboBoxEx(DirContainer.Edit_Editor) do
          Items[GetIndexOfObject(ID_Directory.ToInteger)] := Name1;
    end;

    if Assigned(DirContainer.Edit_DirectoryQuery) then
      with TsComboBoxEx(DirContainer.Edit_DirectoryQuery) do
        Items[GetIndexOfObject(ID_Directory.ToInteger)] := Name1;

    // Currently this code only updates SG_CITY.REGION_NAME when REGION is edited
    if DirCode = DIR_CODE_REGION then
    begin
      SGCity.BeginUpdate;
      for i := 0 to SGCity.RowCount - 1 do
        if SGCity.Cells[4, i] = ID_Directory then
          SGCity.Cells[3, i] := Name1;
      SGCity.EndUpdate;
    end;

    // update rubrikator_tree if needed
    if DirCode = DIR_CODE_RUBRIKA then
    begin
      with FormMain.SearchNode(FormMain.TVRubrikator, ID_Directory.ToInteger, 0) do
      begin
        Text := Name1;
        FormMain.TVRubrikator.CustomSort(@Helpers.CustomSortProc, 0, True);
      end;
    end;
  finally
    DirContainer.Free;
  end;

  Result := StrToIntDef(ID_Directory, -1);
end;

function TFormDirectory.Directory_DELETE(DirCode: integer; ID_Directory: string): integer;
var
  Query: TIBCQuery;
  SQL_select, EditControlName: string;
  DirContainer: TDirectoryContainer;
  i: integer;
  EditControl: TsComboBoxEx;
  IsDeleted: Boolean;
begin
  Result := -1;
  IsDeleted := False;

  if not DirCode in [0 .. DIR_CODE_TOTAL] then
    exit;

  if Trim(ID_Directory) = EmptyStr then
    exit;

  { DATABASE UPDATE part }

  case DirCode of
    DIR_CODE_CURATOR:
      SQL_select := 'select COUNT(*) as CNT from BASE where CURATOR like ' + QuotedStr('%#' + ID_Directory + '$%');
    DIR_CODE_RUBRIKA:
      SQL_select := 'select COUNT(*) as CNT from BASE where RUBR like ' + QuotedStr('%#' + ID_Directory + '$%');
    DIR_CODE_FIRMTYPE:
      SQL_select := 'select COUNT(*) as CNT from BASE where FIRMTYPE like ' + QuotedStr('%#' + ID_Directory + '$%');
    DIR_CODE_NAPRAVLENIE:
      SQL_select := 'select COUNT(*) as CNT from BASE where NAPRAVLENIE like ' + QuotedStr('%#' + ID_Directory + '$%');
    DIR_CODE_OFFICETYPE:
      SQL_select := 'select COUNT(*) as CNT from BASE where ADRES like ' + QuotedStr('%#@' + ID_Directory + '$%');
    DIR_CODE_COUNTRY:
      SQL_select := 'select COUNT(*) as CNT from BASE where ADRES like ' + QuotedStr('%#&' + ID_Directory + '$%');
    DIR_CODE_REGION:
      SQL_select := 'select SUM(c) as CNT from (select COUNT(*) c from BASE where ADRES like ' + QuotedStr('%#*' + ID_Directory + '$%') +
        ' UNION ALL select COUNT(*) from CITY where ID_REGION = ' + ID_Directory + ')';
    DIR_CODE_CITY:
      SQL_select := 'select COUNT(*) as CNT from BASE where ADRES like ' + QuotedStr('%#^' + ID_Directory + '$%');
    DIR_CODE_PHONETYPE:
      SQL_select := EmptyStr;
  end;

  // check if record is in use
  if SQL_select <> EmptyStr then
  begin
    Query := QueryCreate;
    try
      Query.SQL.Text := SQL_select;
      Query.Open;
      Query.FetchAll := True;
      if Query.FieldValues['CNT'] > 0 then
      begin
        MessageBox(handle, PChar('Удаление директории невозможно так как она используется ' + VarToStrDef(Query.FieldByName('CNT').AsString,
          '?') + ' фирмами.'), 'Уведомление', MB_OK or MB_ICONWARNING);
        exit;
      end
    finally
      Query.Free;
    end;
  end;

  // delete record
  if MessageBox(handle, 'Вы уверены что хотите удалить выбранную директорию?', 'Удалить', MB_YESNO or MB_ICONQUESTION) = MRYES then
  begin
    Query := QueryCreate;
    try
      Query.SQL.Text := 'delete from ' + DIR_CODE_TO_TABLE[DirCode] + ' where ID = :ID';
      Query.ParamByName('ID').AsString := ID_Directory;
      try
        Query.Execute;
        Query.Transaction.CommitRetaining;
        IsDeleted := True;
      except
        on E: Exception do
        begin
          Query.Transaction.RollbackRetaining;
          WriteLog('TFormDirectory.Directory_DELETE' + #13 + 'Ошибка: ' + E.Message);
          MessageBox(handle, PChar('Ошибка при удалении директории.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
          exit;
        end;
      end;
    finally
      Query.Free;
    end;
  end;

  { VISUAL DATA UPDATE part }

  if IsDeleted = True then
  begin

    DirContainer := TDirectoryContainer.Create(DirCode);
    try
      if Assigned(DirContainer.SG_Main) then
        with TNextGrid(DirContainer.SG_Main) do
        begin
          if FindText(1, ID_Directory, [soCaseInsensitive, soExactMatch]) then
            DeleteRow(SelectedRow);
        end;

      if Assigned(DirContainer.SG_Directory) then
        with TNextGrid(DirContainer.SG_Directory) do
        begin
          if FindText(1, ID_Directory, [soCaseInsensitive, soExactMatch]) then
            DeleteRow(SelectedRow);
        end;

      if Assigned(DirContainer.Edit_Editor) then
      begin
        if DirContainer.IsAdresEdits = True then
        begin
          EditControlName := TsComboBoxEx(DirContainer.Edit_Editor).Name;
          delete(EditControlName, Length(EditControlName), 1);
          for i := 1 to 10 do
          begin
            EditControl := TsComboBoxEx(FormEditor.FindComponent(EditControlName + IntToStr(i)));
            EditControl.Items.Delete(EditControl.GetIndexOfObject(ID_Directory.ToInteger));
          end;
        end
        else
          with TsComboBoxEx(DirContainer.Edit_Editor) do
            Items.Delete(GetIndexOfObject(ID_Directory.ToInteger));
      end;

      if Assigned(DirContainer.Edit_DirectoryQuery) then
        with TsComboBoxEx(DirContainer.Edit_DirectoryQuery) do
          Items.Delete(GetIndexOfObject(ID_Directory.ToInteger));

      // update rubrikator_tree if needed
      if DirCode = DIR_CODE_RUBRIKA then
        with FormMain.SearchNode(FormMain.TVRubrikator, ID_Directory.ToInteger, 0) do
          Delete;
    finally
      DirContainer.Free;
    end;

  end;

  Result := StrToIntDef(ID_Directory, -1);
end;

end.
