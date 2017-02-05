unit Directory;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, sPageControl, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, StdCtrls, sEdit, Buttons, sSpeedButton, ExtCtrls,
  sPanel, IBQuery, NxColumns, NxColumnClasses, sButton;

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
    procedure LoadDataDirectory;
    procedure FormShow(Sender: TObject);
    procedure btnCreateClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure editCuratorChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams); override;
    { Public declarations }
  end;

var
  FormDirectory: TFormDirectory;
  isDataEdited: Boolean = False;

implementation

uses Main, Editor, Report, Logo;

{$R *.dfm}
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
  Q_Dir: TIBQuery;
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
  Q_Dir.FetchAll;
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
  Q_Dir.FetchAll;
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
  Q_Dir.FetchAll;
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
  Q_Dir.FetchAll;
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
  s, req, max_id: string;
  Q_Dir: TIBQuery;
begin
  if InputQuery('Создать', 'Управление директориями', s) then
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
      Q_Dir := TIBQuery.Create(FormDirectory);
      Q_Dir.Database := FormMain.IBDatabase1;
      Q_Dir.Transaction := FormMain.IBTransaction1;
      Q_Dir.Close;
      Q_Dir.SQL.Text := req;
      Q_Dir.Params[0].AsString := AnsiLowerCase(s);
      Q_Dir.Open; // Q_Dir.FetchAll;
      if Q_Dir.RecordCount > 0 then
      begin
        MessageBox(handle, 'Запись с таким названием уже существует', 'Информация', MB_OK or MB_ICONINFORMATION);
        Q_Dir.Close;
        Q_Dir.Free;
        exit;
      end;
      case PageControl.ActivePageIndex of
        0:
          req := 'insert into CURATOR (NAME) values (:NAME)';
        1:
          req := 'insert into RUBRIKATOR (NAME) values (:NAME)';
        2:
          req := 'insert into TYPE (NAME) values (:NAME)';
        3:
          req := 'insert into NAPRAVLENIE (NAME) values (:NAME)';
        4:
          req := 'insert into OFFICETYPE (NAME) values (:NAME)';
        5:
          req := 'insert into COUNTRY (NAME) values (:NAME)';
        6:
          req := 'insert into GOROD (NAME) values (:NAME)';
        7:
          req := 'insert into PHONETYPE (NAME) values (:NAME)';
      end;
      Q_Dir.Close;
      Q_Dir.SQL.Text := req;
      s := FormEditor.UpperFirst(s);
      Q_Dir.Params[0].AsString := s;
      try
        Q_Dir.ExecSQL;
      except
        on E: Exception do
        begin
          WriteLog('TFormDirectory.btnCreateClick' + #13 + 'Ошибка: ' + E.Message);
          MessageBox(handle, PChar('Ошибка при создании директории.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
          Q_Dir.Close;
          Q_Dir.Free;
          exit;
        end;
      end;
      FormMain.IBTransaction1.CommitRetaining;
      isDataEdited := True;
      Q_Dir.Close;
      case PageControl.ActivePageIndex of
        0:
          Q_Dir.SQL.Text := 'select MAX(ID) from CURATOR';
        1:
          Q_Dir.SQL.Text := 'select MAX(ID) from RUBRIKATOR';
        2:
          Q_Dir.SQL.Text := 'select MAX(ID) from TYPE';
        3:
          Q_Dir.SQL.Text := 'select MAX(ID) from NAPRAVLENIE';
        4:
          Q_Dir.SQL.Text := 'select MAX(ID) from OFFICETYPE';
        5:
          Q_Dir.SQL.Text := 'select MAX(ID) from COUNTRY';
        6:
          Q_Dir.SQL.Text := 'select MAX(ID) from GOROD';
        7:
          Q_Dir.SQL.Text := 'select MAX(ID) from PHONETYPE';
      end;
      Q_Dir.Open;
      max_id := Q_Dir.Fields[0].Value;
      if PageControl.ActivePageIndex = 0 then
      begin
        SGCurator.AddRow; // КУРАТОРЫ
        SGCurator.Cells[0, SGCurator.LastAddedRow] := s;
        SGCurator.Cells[1, SGCurator.LastAddedRow] := max_id;
        SGCurator.SelectLastRow;
        SGCurator.SetFocus;
      end
      else if PageControl.ActivePageIndex = 1 then
      begin
        SGRubr.AddRow; // РУБРИКИ
        SGRubr.Cells[0, SGRubr.LastAddedRow] := s;
        SGRubr.Cells[1, SGRubr.LastAddedRow] := max_id;
        SGRubr.SelectLastRow;
        SGRubr.SetFocus;
      end
      else if PageControl.ActivePageIndex = 2 then
      begin
        SGFirmType.AddRow; // ТИП ФИРМЫ
        SGFirmType.Cells[0, SGFirmType.LastAddedRow] := s;
        SGFirmType.Cells[1, SGFirmType.LastAddedRow] := max_id;
        SGFirmType.SelectLastRow;
        SGFirmType.SetFocus;
      end
      else if PageControl.ActivePageIndex = 3 then
      begin
        SGNapr.AddRow; // ДЕЯТЕЛЬНОСТЬ
        SGNapr.Cells[0, SGNapr.LastAddedRow] := s;
        SGNapr.Cells[1, SGNapr.LastAddedRow] := max_id;
        SGNapr.SelectLastRow;
        SGNapr.SetFocus;
      end
      else if PageControl.ActivePageIndex = 4 then
      begin
        SGOfficeType.AddRow; // ТИП ОФФИСА
        SGOfficeType.Cells[0, SGOfficeType.LastAddedRow] := s;
        SGOfficeType.Cells[1, SGOfficeType.LastAddedRow] := max_id;
        SGOfficeType.SelectLastRow;
        SGOfficeType.SetFocus;
      end
      else if PageControl.ActivePageIndex = 5 then
      begin
        SGCountry.AddRow; // СТРАНЫ
        SGCountry.Cells[0, SGCountry.LastAddedRow] := s;
        SGCountry.Cells[1, SGCountry.LastAddedRow] := max_id;
        SGCountry.SelectLastRow;
        SGCountry.SetFocus;
      end
      else if PageControl.ActivePageIndex = 6 then
      begin
        SGCity.AddRow; // ГОРОДА
        SGCity.Cells[0, SGCity.LastAddedRow] := s;
        SGCity.Cells[1, SGCity.LastAddedRow] := max_id;
        SGCity.SelectLastRow;
        SGCity.SetFocus;
      end
      else if PageControl.ActivePageIndex = 7 then
      begin
        SGPhoneType.AddRow; // ТИПЫ ТЕЛЕФОНОВ
        SGPhoneType.Cells[0, SGPhoneType.LastAddedRow] := s;
        SGPhoneType.Cells[1, SGPhoneType.LastAddedRow] := max_id;
        SGPhoneType.SelectLastRow;
        SGPhoneType.SetFocus;
      end;
      Q_Dir.Close;
      Q_Dir.Free;
    end;
end;

procedure TFormDirectory.btnEditClick(Sender: TObject);
var
  s, req, id: string;
  Q_Dir: TIBQuery;
  grid: TNextGrid;
begin
  grid := nil;
  if PageControl.ActivePageIndex = 0 then
  begin
    if SGCurator.SelectedCount = 0 then
      exit;
    s := SGCurator.Cells[0, SGCurator.SelectedRow];
    id := SGCurator.Cells[1, SGCurator.SelectedRow];
    grid := SGCurator;
  end
  else if PageControl.ActivePageIndex = 1 then
  begin
    if SGRubr.SelectedCount = 0 then
      exit;
    s := SGRubr.Cells[0, SGRubr.SelectedRow];
    id := SGRubr.Cells[1, SGRubr.SelectedRow];
    grid := SGRubr;
  end
  else if PageControl.ActivePageIndex = 2 then
  begin
    if SGFirmType.SelectedCount = 0 then
      exit;
    s := SGFirmType.Cells[0, SGFirmType.SelectedRow];
    id := SGFirmType.Cells[1, SGFirmType.SelectedRow];
    grid := SGFirmType;
  end
  else if PageControl.ActivePageIndex = 3 then
  begin
    if SGNapr.SelectedCount = 0 then
      exit;
    s := SGNapr.Cells[0, SGNapr.SelectedRow];
    id := SGNapr.Cells[1, SGNapr.SelectedRow];
    grid := SGNapr;
  end
  else if PageControl.ActivePageIndex = 4 then
  begin
    if SGOfficeType.SelectedCount = 0 then
      exit;
    s := SGOfficeType.Cells[0, SGOfficeType.SelectedRow];
    id := SGOfficeType.Cells[1, SGOfficeType.SelectedRow];
    grid := SGOfficeType;
  end
  else if PageControl.ActivePageIndex = 5 then
  begin
    if SGCountry.SelectedCount = 0 then
      exit;
    s := SGCountry.Cells[0, SGCountry.SelectedRow];
    id := SGCountry.Cells[1, SGCountry.SelectedRow];
    grid := SGCountry;
  end
  else if PageControl.ActivePageIndex = 6 then
  begin
    if SGCity.SelectedCount = 0 then
      exit;
    s := SGCity.Cells[0, SGCity.SelectedRow];
    id := SGCity.Cells[1, SGCity.SelectedRow];
    grid := SGCity;
  end
  else if PageControl.ActivePageIndex = 7 then
  begin
    if SGPhoneType.SelectedCount = 0 then
      exit;
    s := SGPhoneType.Cells[0, SGPhoneType.SelectedRow];
    id := SGPhoneType.Cells[1, SGPhoneType.SelectedRow];
    grid := SGPhoneType;
  end;
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
      Q_Dir := TIBQuery.Create(FormDirectory);
      Q_Dir.Database := FormMain.IBDatabase1;
      Q_Dir.Transaction := FormMain.IBTransaction1;
      Q_Dir.Close;
      Q_Dir.SQL.Text := req;
      Q_Dir.Params[0].AsString := AnsiLowerCase(s);
      Q_Dir.Open; // Q_Dir.FetchAll;
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
      s := FormEditor.UpperFirst(s);
      Q_Dir.ParamByName('NAME').AsString := s;
      Q_Dir.ParamByName('ID').AsString := id;
      try
        Q_Dir.ExecSQL;
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
    end;
end;

procedure TFormDirectory.btnDeleteClick(Sender: TObject);

  function CheckInUse(field, match: string; var count: integer): Boolean;
  var
    Q: TIBQuery;
  begin
    result := False;
    count := 0;
    if (Trim(field) = '') or (Trim(match) = '') then
      exit;
    Q := TIBQuery.Create(FormDirectory);
    Q.Database := FormMain.IBDatabase1;
    Q.Transaction := FormMain.IBTransaction1;
    Q.Close;
    Q.SQL.Text := 'select ID from BASE where lower(' + field + ') like :STR';
    Q.ParamByName('STR').AsString := match;
    Q.Open;
    Q.FetchAll;
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
  Q_Dir: TIBQuery;
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
    Q_Dir := TIBQuery.Create(FormDirectory);
    Q_Dir.Database := FormMain.IBDatabase1;
    Q_Dir.Transaction := FormMain.IBTransaction1;
    Q_Dir.Close;
    Q_Dir.SQL.Text := req;
    Q_Dir.Params[0].AsString := id;
    try
      Q_Dir.ExecSQL;
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

end.
