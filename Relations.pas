unit Relations;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, sPanel, StdCtrls, sButton, Buttons, sSpeedButton,
  sLabel, NxColumns, NxColumnClasses, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, sGroupBox, sComboBox;

type
  TFormRelations = class(TForm)
    panelRelations: TsPanel;
    editRubrRelations: TsComboBox;
    sGroupBox: TsGroupBox;
    SGNaprAviable: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    SGNaprSelected: TNextGrid;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    lblAviableNapr: TsLabel;
    lblSelectedNapr: TsLabel;
    btnAdd: TsSpeedButton;
    btnRemove: TsSpeedButton;
    btnRemoveAll: TsSpeedButton;
    btnAddAll: TsSpeedButton;
    BtnOK: TsButton;
    BtnCancel: TsButton;
    lblID: TsLabel;
    procedure BtnCancelClick(Sender: TObject);
    procedure BtnOKClick(Sender: TObject);
    procedure ClearEdits;
    procedure PrepareRecord(ID: string);
    function BuildRelationsString(RubrID: string): string;
    procedure btnAddClick(Sender: TObject);
    procedure btnRemoveClick(Sender: TObject);
    procedure btnAddAllClick(Sender: TObject);
    procedure btnRemoveAllClick(Sender: TObject);
    procedure editRubrRelationsChange(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams); override;
    { Public declarations }
  end;

var
  FormRelations: TFormRelations;
  listRubrRelations: TStrings;
  FullNaprList: string;

implementation

uses Main, Editor;

{$R *.dfm}

procedure TFormRelations.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.ExStyle := Params.ExStyle or WS_Ex_AppWindow;
end;

procedure TFormRelations.BtnCancelClick(Sender: TObject);
begin
  FormRelations.Close;
end;

procedure TFormRelations.BtnOKClick(Sender: TObject);
var
  i: integer;
  s: string;
begin
  try
    if listRubrRelations.Count > 0 then
    begin
      for i := 0 to listRubrRelations.Count - 1 do
        s := s + listRubrRelations[i];
      FormMain.IBQuery1.Close;
      FormMain.IBQuery1.SQL.Text := 'update BASE set RELATIONS = :RELATIONS where ID = :ID';
      FormMain.IBQuery1.ParamByName('RELATIONS').AsString := s;
      FormMain.IBQuery1.ParamByName('ID').AsString := lblID.Caption;
      FormMain.IBQuery1.ExecSQL;
      FormMain.IBTransaction1.CommitRetaining;
      FormMain.IBQuery1.Close;
      WriteLog('TFormRelations.BtnOKClick: RELATIONS успешно обновленны ' + lblID.Caption);
    end;
  except
    on E: Exception do
    begin
      WriteLog('TFormRelations.BtnOKClick: RELATIONS ошибка обновления' + #13 + 'Ошибка: ' + E.Message + #13 +
        'listRubrRelations.Text =' + #13 + listRubrRelations.Text);
      MessageBox(Handle, PChar('Ошибка при сохранении данных фирмы (RELATIONS).' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
    end;
  end;
  FormRelations.Close;
end;

procedure TFormRelations.ClearEdits;
begin
  editRubrRelations.Clear;
  SGNaprAviable.ClearRows;
  SGNaprSelected.ClearRows;
  if not Assigned(listRubrRelations) then
    listRubrRelations := TStringList.Create;
  listRubrRelations.Clear;
end;

procedure TFormRelations.PrepareRecord(ID: string);
var
  RUBRstr, NAPRstr, RELATIONSstr, str1: string;
begin
  lblID.Caption := ID;
  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text := 'select * from BASE where ID = :ID';
  FormMain.IBQuery1.ParamByName('ID').AsString := ID;
  FormMain.IBQuery1.Open;
  FormMain.IBQuery1.FetchAll;
  if FormMain.IBQuery1.FieldValues['RUBR'] <> null then
    RUBRstr := FormMain.IBQuery1.FieldValues['RUBR'];
  if FormMain.IBQuery1.FieldValues['NAPRAVLENIE'] <> null then
  begin
    FullNaprList := FormMain.IBQuery1.FieldValues['NAPRAVLENIE'];
    NAPRstr := FormMain.IBQuery1.FieldValues['NAPRAVLENIE'];
  end;
  if FormMain.IBQuery1.FieldValues['RELATIONS'] <> null then
    RELATIONSstr := FormMain.IBQuery1.FieldValues['RELATIONS'];
  FormMain.IBQuery1.Close;
  while Pos('$', RUBRstr) > 0 do
  begin
    str1 := copy(RUBRstr, 0, Pos('$', RUBRstr));
    delete(RUBRstr, 1, length(str1));
    delete(str1, 1, 1);
    delete(str1, length(str1), 1);
    editRubrRelations.AddItem(FormEditor.GetNameByID('RUBRIKATOR', str1), Pointer(StrToInt(str1)));
  end;
  if editRubrRelations.Items.Count > 0 then
    editRubrRelations.ItemIndex := 0;
  while Pos('$', RELATIONSstr) > 0 do
  begin
    str1 := copy(RELATIONSstr, 0, Pos('$', RELATIONSstr));
    delete(RELATIONSstr, 1, length(str1));
    listRubrRelations.Add(str1);
  end;
  editRubrRelationsChange(editRubrRelations);
  // FormMain.IBDatabase1.Connected := False;
end;

procedure TFormRelations.btnAddClick(Sender: TObject);
begin
  if SGNaprAviable.SelectedCount = 0 then
    exit;
  SGNaprSelected.AddRow;
  SGNaprSelected.Cells[0, SGNaprSelected.LastAddedRow] := SGNaprAviable.Cells[0, SGNaprAviable.SelectedRow];
  SGNaprSelected.Cells[1, SGNaprSelected.LastAddedRow] := SGNaprAviable.Cells[1, SGNaprAviable.SelectedRow];
  SGNaprAviable.DeleteRow(SGNaprAviable.SelectedRow);
  BuildRelationsString(IntToStr(integer(editRubrRelations.Items.Objects[editRubrRelations.ItemIndex])));
end;

procedure TFormRelations.btnRemoveClick(Sender: TObject);
begin
  if SGNaprSelected.SelectedCount = 0 then
    exit;
  SGNaprAviable.AddRow;
  SGNaprAviable.Cells[0, SGNaprAviable.LastAddedRow] := SGNaprSelected.Cells[0, SGNaprSelected.SelectedRow];
  SGNaprAviable.Cells[1, SGNaprAviable.LastAddedRow] := SGNaprSelected.Cells[1, SGNaprSelected.SelectedRow];
  SGNaprSelected.DeleteRow(SGNaprSelected.SelectedRow);
  BuildRelationsString(IntToStr(integer(editRubrRelations.Items.Objects[editRubrRelations.ItemIndex])));
end;

procedure TFormRelations.btnAddAllClick(Sender: TObject);
var
  i: integer;
begin
  if SGNaprAviable.RowCount = 0 then
    exit;
  SGNaprAviable.BeginUpdate;
  SGNaprSelected.BeginUpdate;
  for i := 0 to SGNaprAviable.RowCount - 1 do
  begin
    SGNaprSelected.AddRow;
    SGNaprSelected.Cells[0, SGNaprSelected.LastAddedRow] := SGNaprAviable.Cells[0, i];
    SGNaprSelected.Cells[1, SGNaprSelected.LastAddedRow] := SGNaprAviable.Cells[1, i];
  end;
  SGNaprAviable.ClearRows;
  SGNaprAviable.EndUpdate;
  SGNaprSelected.EndUpdate;
  BuildRelationsString(IntToStr(integer(editRubrRelations.Items.Objects[editRubrRelations.ItemIndex])));
end;

procedure TFormRelations.btnRemoveAllClick(Sender: TObject);
var
  i: integer;
begin
  if SGNaprSelected.RowCount = 0 then
    exit;
  SGNaprAviable.BeginUpdate;
  SGNaprSelected.BeginUpdate;
  for i := 0 to SGNaprSelected.RowCount - 1 do
  begin
    SGNaprAviable.AddRow;
    SGNaprAviable.Cells[0, SGNaprAviable.LastAddedRow] := SGNaprSelected.Cells[0, i];
    SGNaprAviable.Cells[1, SGNaprAviable.LastAddedRow] := SGNaprSelected.Cells[1, i];
  end;
  SGNaprSelected.ClearRows;
  SGNaprAviable.EndUpdate;
  SGNaprSelected.EndUpdate;
  BuildRelationsString(IntToStr(integer(editRubrRelations.Items.Objects[editRubrRelations.ItemIndex])));
end;

function TFormRelations.BuildRelationsString(RubrID: string): string;
var
  i, x: integer;
  finalStr: string;
begin
  Result := '';
  if listRubrRelations.Count = 0 then
    exit;
  for i := 0 to listRubrRelations.Count - 1 do
  begin
    if Pos('#' + RubrID + '=', listRubrRelations[i]) > 0 then
    begin
      finalStr := '#' + RubrID + '=';
      for x := 0 to SGNaprSelected.RowCount - 1 do
      begin
        finalStr := finalStr + '(' + SGNaprSelected.Cells[1, x] + ')';
      end;
      listRubrRelations[i] := finalStr + '$';
      Break;
    end;
  end;
  Result := finalStr + '$';
end;

procedure TFormRelations.editRubrRelationsChange(Sender: TObject);
var
  i: integer;
  RubrID, NAPRstr, str1, str2: string;
begin
  if editRubrRelations.ItemIndex = -1 then
    exit;
  SGNaprAviable.BeginUpdate;
  SGNaprSelected.BeginUpdate;
  RubrID := IntToStr(integer(editRubrRelations.Items.Objects[editRubrRelations.ItemIndex]));
  NAPRstr := FullNaprList;
  SGNaprAviable.ClearRows;
  SGNaprSelected.ClearRows;
  // очистили списки, добавили лист доступных направлений
  while Pos('$', NAPRstr) > 0 do
  begin
    str1 := copy(NAPRstr, 0, Pos('$', NAPRstr));
    delete(NAPRstr, 1, length(str1));
    delete(str1, 1, 1);
    delete(str1, length(str1), 1);
    SGNaprAviable.AddRow;
    SGNaprAviable.Cells[0, SGNaprAviable.LastAddedRow] := FormEditor.GetNameByID('NAPRAVLENIE', str1);
    SGNaprAviable.Cells[1, SGNaprAviable.LastAddedRow] := str1;
  end;

  for i := 0 to listRubrRelations.Count - 1 do
  begin
    if Pos('#' + RubrID + '=', listRubrRelations[i]) > 0 then
    begin
      str1 := listRubrRelations[i];
      Break;
      // получили строку с которой грузим выбранные направления
    end;
  end;
  delete(str1, 1, Pos('=', str1));
  while Pos(')', str1) > 0 do
  begin
    str2 := copy(str1, 0, Pos(')', str1));
    delete(str1, 1, length(str2));
    delete(str2, 1, 1);
    delete(str2, length(str2), 1);
    SGNaprSelected.AddRow;
    SGNaprSelected.Cells[0, SGNaprSelected.LastAddedRow] := FormEditor.GetNameByID('NAPRAVLENIE', str2);
    SGNaprSelected.Cells[1, SGNaprSelected.LastAddedRow] := str2;
    if SGNaprAviable.FindText(1, str2, [soCaseInsensitive, soExactMatch]) then
      SGNaprAviable.DeleteRow(SGNaprAviable.SelectedRow);
  end;
  SGNaprAviable.EndUpdate;
  SGNaprSelected.EndUpdate;
end;

end.
