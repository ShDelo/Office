{ #DONE: UI: Implement Region>City edits interaction. e.g when region not selected > allow full city list.
  when region is selected > limit city list by that id_region.
  I had to drop the idea of forming city list by region_id due to problems with updating edit controls in OnSelect event
  and implemented more simple way of this interaction:
  selecting region will clear current city.text
  selecting city will automatically load region for that city }

{ #TODO1: UI : edit controls that doesn't allowing to pass unlisted values (e.g: country, region, city) should auto-clear
  when losing focus, and need to be checked before let code proceed }

{ #TODO1: UI : think how we gonna display adres data now, that we have region field added.
  should it be displayed in Tabs, report simple, report complex? what about DELO? anywhere else? }

{ #TODO1: SAFECODE : check values length (adres, phones) before allow them to pass into insert statement and check for valid structure
  i.e 1-10 adreses and phones maybe something else }

unit Editor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, sButton, ExtCtrls, sPanel, sComboBoxes, sEdit,
  sCheckBox, sLabel, sGroupBox, ComCtrls, sPageControl, acPNG, IBC,
  NxColumns, NxColumnClasses, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, sSpeedButton, Buttons, sMemo, StrUtils;

type
  TFormEditor = class(TForm)
    PanelAdv: TsPanel;
    PanelMain: TsPanel;
    EditName: TsEdit;
    EditFIO: TsEdit;
    EditCurator: TsComboBoxEx;
    EditRubr: TsComboBoxEx;
    EditFirmType: TsComboBoxEx;
    EditNapravlenie: TsComboBoxEx;
    sPageControl1: TsPageControl;
    sTabSheet1: TsTabSheet;
    sGroupBox1: TsGroupBox;
    ImagePhone1: TImage;
    EditNO1: TsEdit;
    EditOfficeType1: TsComboBoxEx;
    EditZIP1: TsEdit;
    EditStreet1: TsEdit;
    EditCountry1: TsComboBoxEx;
    EditRegion1: TsComboBoxEx;
    EditCity1: TsComboBoxEx;
    sTabSheet2: TsTabSheet;
    sGroupBox2: TsGroupBox;
    EditNO2: TsEdit;
    EditOfficeType2: TsComboBoxEx;
    EditZIP2: TsEdit;
    EditStreet2: TsEdit;
    EditCountry2: TsComboBoxEx;
    EditRegion2: TsComboBoxEx;
    EditCity2: TsComboBoxEx;
    sTabSheet3: TsTabSheet;
    sGroupBox3: TsGroupBox;
    EditNO3: TsEdit;
    EditOfficeType3: TsComboBoxEx;
    EditZIP3: TsEdit;
    EditStreet3: TsEdit;
    EditCountry3: TsComboBoxEx;
    EditRegion3: TsComboBoxEx;
    EditCity3: TsComboBoxEx;
    sTabSheet4: TsTabSheet;
    sGroupBox4: TsGroupBox;
    EditNO4: TsEdit;
    EditOfficeType4: TsComboBoxEx;
    EditZIP4: TsEdit;
    EditStreet4: TsEdit;
    EditCountry4: TsComboBoxEx;
    EditRegion4: TsComboBoxEx;
    EditCity4: TsComboBoxEx;
    sTabSheet5: TsTabSheet;
    sGroupBox5: TsGroupBox;
    EditNO5: TsEdit;
    EditOfficeType5: TsComboBoxEx;
    EditZIP5: TsEdit;
    EditStreet5: TsEdit;
    EditCountry5: TsComboBoxEx;
    EditRegion5: TsComboBoxEx;
    EditCity5: TsComboBoxEx;
    sTabSheet6: TsTabSheet;
    sGroupBox6: TsGroupBox;
    EditNO6: TsEdit;
    EditOfficeType6: TsComboBoxEx;
    EditZIP6: TsEdit;
    EditStreet6: TsEdit;
    EditCountry6: TsComboBoxEx;
    EditRegion6: TsComboBoxEx;
    EditCity6: TsComboBoxEx;
    sTabSheet7: TsTabSheet;
    sGroupBox7: TsGroupBox;
    EditNO7: TsEdit;
    EditOfficeType7: TsComboBoxEx;
    EditZIP7: TsEdit;
    EditStreet7: TsEdit;
    EditCountry7: TsComboBoxEx;
    EditRegion7: TsComboBoxEx;
    EditCity7: TsComboBoxEx;
    sTabSheet8: TsTabSheet;
    sGroupBox8: TsGroupBox;
    EditNO8: TsEdit;
    EditOfficeType8: TsComboBoxEx;
    EditZIP8: TsEdit;
    EditStreet8: TsEdit;
    EditCountry8: TsComboBoxEx;
    EditRegion8: TsComboBoxEx;
    EditCity8: TsComboBoxEx;
    sTabSheet9: TsTabSheet;
    sGroupBox9: TsGroupBox;
    EditNO9: TsEdit;
    EditOfficeType9: TsComboBoxEx;
    EditZIP9: TsEdit;
    EditStreet9: TsEdit;
    EditCountry9: TsComboBoxEx;
    EditRegion9: TsComboBoxEx;
    EditCity9: TsComboBoxEx;
    sTabSheet10: TsTabSheet;
    sGroupBox10: TsGroupBox;
    EditNO10: TsEdit;
    EditOfficeType10: TsComboBoxEx;
    EditZIP10: TsEdit;
    EditStreet10: TsEdit;
    EditCountry10: TsComboBoxEx;
    EditRegion10: TsComboBoxEx;
    EditCity10: TsComboBoxEx;
    CBAdres1: TsCheckBox;
    CBAdres2: TsCheckBox;
    CBAdres3: TsCheckBox;
    CBAdres4: TsCheckBox;
    CBAdres5: TsCheckBox;
    CBAdres6: TsCheckBox;
    CBAdres7: TsCheckBox;
    CBAdres8: TsCheckBox;
    CBAdres9: TsCheckBox;
    CBAdres10: TsCheckBox;
    BtnAddRubrToList: TsSpeedButton;
    BtnDeleteRubrFromList: TsSpeedButton;
    BtnAddNaprToList: TsSpeedButton;
    BtnDeleteNaprFromList: TsSpeedButton;
    ImagePhone2: TImage;
    ImagePhone3: TImage;
    ImagePhone4: TImage;
    ImagePhone5: TImage;
    ImagePhone6: TImage;
    ImagePhone7: TImage;
    ImagePhone8: TImage;
    ImagePhone9: TImage;
    ImagePhone10: TImage;
    lblID: TsLabel;
    CBActivity: TsCheckBox;
    BtnOK: TsButton;
    BtnCancel: TsButton;
    BtnAddCuratorToList: TsSpeedButton;
    BtnDeleteCuratorFromList: TsSpeedButton;
    BtnAddFirmTypeToList: TsSpeedButton;
    BtnDeleteFirmTypeFromList: TsSpeedButton;
    EditPhoneType1: TsComboBoxEx;
    SGPhone1: TNextGrid;
    NxTextColumn9: TNxTextColumn;
    MemoPhone1: TsMemo;
    EditPhone1: TsEdit;
    BtnAddPhoneToList1: TsSpeedButton;
    BtnDeletePhoneFromList1: TsSpeedButton;
    EditPhoneType2: TsComboBoxEx;
    SGPhone2: TNextGrid;
    NxTextColumn10: TNxTextColumn;
    MemoPhone2: TsMemo;
    EditPhone2: TsEdit;
    BtnAddPhoneToList2: TsSpeedButton;
    BtnDeletePhoneFromList2: TsSpeedButton;
    EditPhoneType3: TsComboBoxEx;
    SGPhone3: TNextGrid;
    NxTextColumn11: TNxTextColumn;
    MemoPhone3: TsMemo;
    EditPhone3: TsEdit;
    BtnAddPhoneToList3: TsSpeedButton;
    BtnDeletePhoneFromList3: TsSpeedButton;
    EditPhoneType4: TsComboBoxEx;
    SGPhone4: TNextGrid;
    NxTextColumn12: TNxTextColumn;
    MemoPhone4: TsMemo;
    EditPhone4: TsEdit;
    BtnAddPhoneToList4: TsSpeedButton;
    BtnDeletePhoneFromList4: TsSpeedButton;
    EditPhoneType5: TsComboBoxEx;
    SGPhone5: TNextGrid;
    NxTextColumn13: TNxTextColumn;
    MemoPhone5: TsMemo;
    EditPhone5: TsEdit;
    BtnAddPhoneToList5: TsSpeedButton;
    BtnDeletePhoneFromList5: TsSpeedButton;
    EditPhoneType6: TsComboBoxEx;
    SGPhone6: TNextGrid;
    NxTextColumn14: TNxTextColumn;
    MemoPhone6: TsMemo;
    EditPhone6: TsEdit;
    BtnAddPhoneToList6: TsSpeedButton;
    BtnDeletePhoneFromList6: TsSpeedButton;
    EditPhoneType7: TsComboBoxEx;
    SGPhone7: TNextGrid;
    NxTextColumn15: TNxTextColumn;
    MemoPhone7: TsMemo;
    EditPhone7: TsEdit;
    BtnAddPhoneToList7: TsSpeedButton;
    BtnDeletePhoneFromList7: TsSpeedButton;
    EditPhoneType8: TsComboBoxEx;
    SGPhone8: TNextGrid;
    NxTextColumn16: TNxTextColumn;
    MemoPhone8: TsMemo;
    EditPhone8: TsEdit;
    BtnAddPhoneToList8: TsSpeedButton;
    BtnDeletePhoneFromList8: TsSpeedButton;
    EditPhoneType9: TsComboBoxEx;
    SGPhone9: TNextGrid;
    NxTextColumn17: TNxTextColumn;
    MemoPhone9: TsMemo;
    EditPhone9: TsEdit;
    BtnAddPhoneToList9: TsSpeedButton;
    BtnDeletePhoneFromList9: TsSpeedButton;
    EditPhoneType10: TsComboBoxEx;
    SGPhone10: TNextGrid;
    NxTextColumn18: TNxTextColumn;
    MemoPhone10: TsMemo;
    EditPhone10: TsEdit;
    BtnAddPhoneToList10: TsSpeedButton;
    BtnDeletePhoneFromList10: TsSpeedButton;
    btnDeleteAdres1: TsSpeedButton;
    btnDeleteAdres2: TsSpeedButton;
    btnDeleteAdres3: TsSpeedButton;
    btnDeleteAdres4: TsSpeedButton;
    btnDeleteAdres5: TsSpeedButton;
    btnDeleteAdres6: TsSpeedButton;
    btnDeleteAdres7: TsSpeedButton;
    btnDeleteAdres8: TsSpeedButton;
    btnDeleteAdres9: TsSpeedButton;
    btnDeleteAdres10: TsSpeedButton;
    BtnAddWebToList: TsSpeedButton;
    BtnDeleteWebFromList: TsSpeedButton;
    BtnAddEMailToList: TsSpeedButton;
    BtnDeleteEMailFromList: TsSpeedButton;
    EditWEB: TsEdit;
    EditEMAIL: TsEdit;
    cbDataRelevance: TsCheckBox;
    SGEMail: TNextGrid;
    NxTextColumn21: TNxTextColumn;
    SGWeb: TNextGrid;
    NxTextColumn19: TNxTextColumn;
    SGFirmType: TNextGrid;
    NxTextColumn7: TNxTextColumn;
    NxTextColumn8: TNxTextColumn;
    SGCurator: TNextGrid;
    NxTextColumn5: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    SGRubr: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    SGNapravlenie: TNextGrid;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    procedure LoadDataEditor;
    procedure BtnCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    function GetIDString(component: TNextGrid): string;
    function GetWebEmailString(component: TNextGrid): string;
    function ParseAdresFieldToEntriesList(Field_ADRES_LineByIndex: string): TStringList;
    function ParseAdresFieldToCityIDList(Field_ADRES: string; IncludeSyntax: boolean = false): TStringList;
    function ParseAdresFieldAndUpdateValues(Field_ADRES, Condition: string; UpdateTo: array of TVarRec): string;
    function UpdateAdresFieldByCondition(Condition: string; Values: array of TVarRec): boolean;
    function ParseIDString(IDString: string; IncludeSyntax: boolean = false): TStringList;
    procedure DoGarbageCollection(BASE_ID, DIR_Table, BASE_Field, IDString_OLD, IDString_NEW: string; SGTemp, SGDir: TNextGrid;
      EditOwner: TsComboBoxEx; Method: string);
    procedure DoNaprActivityValidation(BASE_ID, IDString_OLD, IDString_NEW: string; IsActive: boolean; Method: string);
    procedure DoCityActivityValidation(BASE_ID, FieldAdres_OLD, FieldAdres_NEW: string; IsActive: boolean; Method: string);
    function IsNewRecordFound_Notify(var listNewRecords: TStringList): boolean;
    procedure IsNewRecordCheck;
    procedure IsRecordDublicate;
    procedure ClearEdits;
    procedure AddRecord(Sender: TObject);
    procedure PrepareEditRecord(id: string);
    procedure EditRecord(Sender: TObject);
    procedure DeleteRecord(id: string);
    procedure BtnAddCuratorToListClick(Sender: TObject);
    procedure BtnDeleteCuratorFromListClick(Sender: TObject);
    procedure EditCuratorKeyPress(Sender: TObject; var Key: Char);
    procedure SGCuratorKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure EditCuratorExit(Sender: TObject);
    procedure BtnAddPhoneToListClick(Sender: TObject);
    procedure EditPhoneKeyPress(Sender: TObject; var Key: Char);
    procedure btnDeleteAdresClick(Sender: TObject);
    procedure BtnAddWebToListClick(Sender: TObject);
    procedure SGCuratorDblClick(Sender: TObject);
    procedure SGCuratorEnter(Sender: TObject);
    procedure EditRegionOrCityOnExit(Sender: TObject);
    procedure EditRegionOrCityOnKeyPress(Sender: TObject; var Key: Char);
    procedure EditRegionSelect(Sender: TObject);
    procedure EditCitySelect(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  end;

procedure ClearEdit(Edit: TsComboBoxEx; DoItemsClear: boolean = false);
procedure ClearAdresEdits(PageIndex: integer);

var
  FormEditor: TFormEditor;
  IsDublicate: Boolean = False;
  list_GC_IDs: TStringList; // used for DoGarbageCollection and DoNaprActivityValidation

implementation

uses Main, Logo, Directory, Report, Dublicate, Relations, MailSend, Helpers;

{$R *.dfm}

procedure ClearEdit(Edit: TsComboBoxEx; DoItemsClear: boolean = false);
begin
  Edit.Text := '';
  Edit.ItemIndex := -1;
  Edit.Tag := -1;
  if DoItemsClear then
    Edit.Clear;
end;

procedure ClearAdresEdits(PageIndex: integer);
begin
  if (PageIndex < 0) or (PageIndex > 10) then
    exit;

  TsCheckBox(FormEditor.FindComponent('CBAdres' + PageIndex.ToString)).Checked := False;
  ClearEdit(TsComboBoxEx(FormEditor.FindComponent('EditOfficeType' + PageIndex.ToString)));
  TsEdit(FormEditor.FindComponent('EditZIP' + PageIndex.ToString)).Text := '';
  ClearEdit(TsComboBoxEx(FormEditor.FindComponent('EditCountry' + PageIndex.ToString)));
  ClearEdit(TsComboBoxEx(FormEditor.FindComponent('EditRegion' + PageIndex.ToString)));
  ClearEdit(TsComboBoxEx(FormEditor.FindComponent('EditCity' + PageIndex.ToString)));
  TsEdit(FormEditor.FindComponent('EditStreet' + PageIndex.ToString)).Text := '';
  ClearEdit(TsComboBoxEx(FormEditor.FindComponent('EditPhoneType' + PageIndex.ToString)));
  TsEdit(FormEditor.FindComponent('EditPhone' + PageIndex.ToString)).Text := '';
  TNextGrid(FormEditor.FindComponent('SGPhone' + PageIndex.ToString)).ClearRows;
  TsMemo(FormEditor.FindComponent('MemoPhone' + PageIndex.ToString)).Clear;
end;

procedure TFormEditor.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.ExStyle := Params.ExStyle or WS_Ex_AppWindow;
end;

function TFormEditor.ParseAdresFieldToEntriesList(Field_ADRES_LineByIndex: string): TStringList;
var
  Entry: string;

  function RemoveSyntax(str: string): string;
  begin
    if Length(Trim(str)) > 0 then
      if str[1] in ['@', '&', '*', '^'] then
        delete(str, 1, 1);
    result := str;
  end;

begin
  // ResultList[0] = CBAdres; ResultList[1] = NO; ResultList[2] = OfficeType; ResultList[3] = ZIP;
  // ResultList[4] = Street; ResultList[5] = Country; ResultList[6] = Region; ResultList[7] = City;
  Result := TStringList.Create;

  while pos('$', Field_ADRES_LineByIndex) > 0 do
  begin
    Entry := copy(Field_ADRES_LineByIndex, 0, pos('$', Field_ADRES_LineByIndex));
    delete(Field_ADRES_LineByIndex, 1, length(Entry));
    delete(Entry, 1, 1);
    delete(Entry, length(Entry), 1);
    Result.Add(Trim(Entry));
  end;

  if Result[0] = '1' then
  begin
    Result[2] := RemoveSyntax(Result[2]);
    Result[5] := RemoveSyntax(Result[5]);
    Result[6] := RemoveSyntax(Result[6]);
    Result[7] := RemoveSyntax(Result[7]);
  end;
end;

function TFormEditor.ParseAdresFieldToCityIDList(Field_ADRES: string; IncludeSyntax: boolean): TStringList;
var
  tmp, AdresString: string;
  AdresList: TStringList;
  i: integer;
begin
  Result := TStringList.Create;
  AdresList := TStringList.Create;

  AdresList.Text := Field_ADRES;
  for i := 0 to AdresList.Count - 1 do
  begin
    AdresString := AdresList[i];
    if pos('#^', AdresString) > 0 then
    begin
      tmp := copy(AdresString, pos('#^', AdresString) + 2, Length(AdresString));
      delete(tmp, Pos('$', tmp), Length(tmp));
      if IncludeSyntax = True then
        tmp := '#' + tmp + '$';
      if Result.IndexOf(tmp) = -1 then
        Result.Add(tmp);
    end;
  end;

  AdresList.Free;
end;

function TFormEditor.ParseAdresFieldAndUpdateValues(Field_ADRES, Condition: string; UpdateTo: array of TVarRec): string;
var
  FieldAdres_List, FieldAdres_Entries: TStringList;
  i: Integer;
begin
  // UpdateTo[0] = ID_OfficeType; UpdateTo[1] = ID_Country; UpdateTo[2] = ID_Region; UpdateTo[3] = ID_City;
  FieldAdres_List := TStringList.Create;
  try
    FieldAdres_List.Text := Field_ADRES;
    for i := 0 to FieldAdres_List.Count - 1 do
    begin
      FieldAdres_Entries := ParseAdresFieldToEntriesList(FieldAdres_List[i]);
      // list2[0] = CBAdres; list2[1] = NO; list2[2] = OfficeType; list2[3] = ZIP;
      // list2[4] = Street; list2[5] = Country; list2[6] = Region; list2[7] = City;
      try
        if (ContainsText(FieldAdres_List[i], Condition)) and (FieldAdres_Entries[0] = '1') then
        begin
          // OfficeType
          if UpdateTo[0].VInteger > -1 then
            FieldAdres_Entries[2] := IntToStr(UpdateTo[0].VInteger);
          if FieldAdres_Entries[2] <> EmptyStr then
            FieldAdres_Entries[2] := '@' + FieldAdres_Entries[2];

          // Country
          if UpdateTo[1].VInteger > -1 then
            FieldAdres_Entries[5] := IntToStr(UpdateTo[1].VInteger);
          if FieldAdres_Entries[5] <> EmptyStr then
            FieldAdres_Entries[5] := '&' + FieldAdres_Entries[5];

          // Region
          if UpdateTo[2].VInteger > -1 then
            FieldAdres_Entries[6] := IntToStr(UpdateTo[2].VInteger);
          if FieldAdres_Entries[6] <> EmptyStr then
            FieldAdres_Entries[6] := '*' + FieldAdres_Entries[6];

          // City
          if UpdateTo[3].VInteger > -1 then
            FieldAdres_Entries[7] := IntToStr(UpdateTo[3].VInteger);
          if FieldAdres_Entries[7] <> EmptyStr then
            FieldAdres_Entries[7] := '^' + FieldAdres_Entries[7];

          FieldAdres_List[i] := Format('#%s$#%s$#%s$#%s$#%s$#%s$#%s$#%s$', [FieldAdres_Entries[0], FieldAdres_Entries[1],
            FieldAdres_Entries[2], FieldAdres_Entries[3], FieldAdres_Entries[4], FieldAdres_Entries[5], FieldAdres_Entries[6],
            FieldAdres_Entries[7]]);
        end;
      finally
        FieldAdres_Entries.Free;
      end;
    end;
    Result := FieldAdres_List.Text;

  finally
    FieldAdres_List.Free;
  end;
end;

function TFormEditor.UpdateAdresFieldByCondition(Condition: string; Values: array of TVarRec): boolean;
var
  QuerySelect, QueryUpdate: TIBCQuery;
  ParsedAdresStr: string;
begin
  { Condition format: #@ID$; #&ID$; #*ID$; #^ID$ }
  Result := False;
  QuerySelect := QueryCreate;
  try
    QuerySelect.SQL.Text := 'select ID, ADRES from BASE where ADRES like :ADRES';
    QuerySelect.ParamByName('ADRES').AsString := '%' + Condition + '%';
    QuerySelect.Open;
    QuerySelect.FetchAll := True;
    if QuerySelect.RecordCount > 0 then
    begin
      QueryUpdate := QueryCreate;
      try
        while not QuerySelect.Eof do
        begin
          ParsedAdresStr := ParseAdresFieldAndUpdateValues(QuerySelect.FieldByName('ADRES').AsString, Condition, Values);
          QueryUpdate.SQL.Text := 'update BASE set ADRES = :ADRES where ID = :ID';
          QueryUpdate.ParamByName('ADRES').AsString := ParsedAdresStr;
          QueryUpdate.ParamByName('ID').AsString := QuerySelect.FieldByName('ID').AsString;
          QueryUpdate.Prepare;
          try
            QueryUpdate.Execute;
          except
            on E: Exception do
            begin
              QueryUpdate.Transaction.Rollback;
              WriteLog('TFormEditor.UpdateAdresFieldByCondition' + #13 + 'Ошибка: ' + E.Message);
              MessageBox(handle, PChar('Ошибка при обновлении данных адресов фирм.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
              exit;
            end;
          end;
          QuerySelect.Next;
        end;
        QueryUpdate.Transaction.Commit;
        Result := True;

      finally
        QueryUpdate.Free;
      end;
    end;

  finally
    QuerySelect.Free;
  end;
end;

function TFormEditor.ParseIDString(IDString: string; IncludeSyntax: boolean = false): TStringList;
var
  tmp: string;
  StringList: TStringList;
begin
  StringList := TStringList.Create;

  while pos('$', IDString) > 0 do
  begin
    tmp := copy(IDString, 0, pos('$', IDString));
    delete(IDString, 1, length(tmp));
    if not IncludeSyntax then
    begin
      delete(tmp, 1, 1);
      delete(tmp, length(tmp), 1);
    end;
    StringList.Add(tmp);
  end;

  result := StringList;
end;

procedure TFormEditor.FormCreate(Sender: TObject);
begin
  LoadDataEditor;
end;

procedure TFormEditor.BtnCancelClick(Sender: TObject);
begin
  WriteLog('TFormEditor.BtnCancelClick: отмена');
  FormEditor.Close;
end;

procedure TFormEditor.BtnAddCuratorToListClick(Sender: TObject);

  procedure Adding(sg: TNextGrid; edit: TsComboBoxEx);
  var
    i: Integer;
  begin
    if Trim(edit.Text) = '' then
      exit;
    for i := 0 to sg.RowCount - 1 do
    begin
      if AnsiLowerCase(sg.Cells[0, i]) = AnsiLowerCase(Trim(edit.Text)) then
      begin
        MessageBox(handle, 'Запись с таким именем уже есть в списке.', 'Информация', MB_OK or MB_ICONINFORMATION);
        exit;
      end;
    end;
    sg.AddRow;
    sg.Cells[0, sg.LastAddedRow] := UpperFirst(edit.Text);
    sg.Cells[1, sg.LastAddedRow] := edit.GetID.ToString;
    sg.Resort;
    edit.Text := '';
  end;

begin
  if TsSpeedButton(Sender).Name = 'BtnAddCuratorToList' then
    Adding(SGCurator, EditCurator);
  if TsSpeedButton(Sender).Name = 'BtnAddRubrToList' then
    Adding(SGRubr, EditRubr);
  if TsSpeedButton(Sender).Name = 'BtnAddFirmTypeToList' then
    Adding(SGFirmType, EditFirmType);
  if TsSpeedButton(Sender).Name = 'BtnAddNaprToList' then
    Adding(SGNapravlenie, EditNapravlenie);
end;

procedure TFormEditor.BtnDeleteCuratorFromListClick(Sender: TObject);
var
  sg: TNextGrid;
begin
  sg := nil;
  if TsSpeedButton(Sender).Name = 'BtnDeleteCuratorFromList' then
    sg := SGCurator;
  if TsSpeedButton(Sender).Name = 'BtnDeleteRubrFromList' then
    sg := SGRubr;
  if TsSpeedButton(Sender).Name = 'BtnDeleteFirmTypeFromList' then
    sg := SGFirmType;
  if TsSpeedButton(Sender).Name = 'BtnDeleteNaprFromList' then
    sg := SGNapravlenie;
  if TsSpeedButton(Sender).Name = 'BtnDeleteWebFromList' then
    sg := SGWeb;
  if TsSpeedButton(Sender).Name = 'BtnDeleteEMailFromList' then
    sg := SGEMail;
  if TsSpeedButton(Sender).Name = 'BtnDeletePhoneFromList1' then
    sg := SGPhone1;
  if TsSpeedButton(Sender).Name = 'BtnDeletePhoneFromList2' then
    sg := SGPhone2;
  if TsSpeedButton(Sender).Name = 'BtnDeletePhoneFromList3' then
    sg := SGPhone3;
  if TsSpeedButton(Sender).Name = 'BtnDeletePhoneFromList4' then
    sg := SGPhone4;
  if TsSpeedButton(Sender).Name = 'BtnDeletePhoneFromList5' then
    sg := SGPhone5;
  if TsSpeedButton(Sender).Name = 'BtnDeletePhoneFromList6' then
    sg := SGPhone6;
  if TsSpeedButton(Sender).Name = 'BtnDeletePhoneFromList7' then
    sg := SGPhone7;
  if TsSpeedButton(Sender).Name = 'BtnDeletePhoneFromList8' then
    sg := SGPhone8;
  if TsSpeedButton(Sender).Name = 'BtnDeletePhoneFromList9' then
    sg := SGPhone9;
  if TsSpeedButton(Sender).Name = 'BtnDeletePhoneFromList10' then
    sg := SGPhone10;
  if Assigned(sg) then
  begin
    if sg.SelectedCount = 0 then
      exit;
    sg.DeleteRow(sg.SelectedRow);
    if sg.RowCount > 0 then
      sg.SelectFirstRow;
  end;
end;

procedure TFormEditor.BtnAddWebToListClick(Sender: TObject);
var
  i: Integer;
  edit: TsEdit;
  sg: TNextGrid;
begin
  edit := nil;
  sg := nil;
  if TsSpeedButton(Sender).Name = 'BtnAddWebToList' then
  begin
    edit := EditWEB;
    sg := SGWeb;
  end;
  if TsSpeedButton(Sender).Name = 'BtnAddEMailToList' then
  begin
    edit := EditEMAIL;
    sg := SGEMail;
  end;
  if Trim(edit.Text) = '' then
    exit;
  for i := 0 to sg.RowCount - 1 do
    if AnsiLowerCase(sg.Cells[0, i]) = AnsiLowerCase(Trim(edit.Text)) then
    begin
      MessageBox(handle, 'Запись с таким именем уже есть в списке.', 'Информация', MB_OK or MB_ICONINFORMATION);
      exit;
    end;
  if TsSpeedButton(Sender).Name = 'BtnAddWebToList' then
    if not FormMailSender.IsValidWeb(edit.Text) then
    begin
      // if (Pos('www.',AnsiLowerCase(edit.Text)) = 0) OR (Pos(',',edit.Text) > 0) OR (Pos(' ',edit.Text) > 0) then begin
      MessageBox(handle, 'Неверно указан электронный адрес.', 'Предупреждение', MB_OK or MB_ICONWARNING);
      exit;
    end;
  if TsSpeedButton(Sender).Name = 'BtnAddEMailToList' then
    if not FormMailSender.IsValidEmail(edit.Text) then
    begin
      MessageBox(handle, 'Неверно указан адрес электронной почты.', 'Предупреждение', MB_OK or MB_ICONWARNING);
      exit;
    end;
  sg.AddRow;
  sg.Cells[0, sg.LastAddedRow] := Trim(edit.Text);
  sg.Resort;
  edit.Text := '';
end;

procedure TFormEditor.EditCuratorKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  begin
    Key := #0;
    if TsComboBoxEx(Sender).Name = 'EditCurator' then
      BtnAddCuratorToListClick(BtnAddCuratorToList);
    if TsComboBoxEx(Sender).Name = 'EditRubr' then
      BtnAddCuratorToListClick(BtnAddRubrToList);
    if TsComboBoxEx(Sender).Name = 'EditFirmType' then
      BtnAddCuratorToListClick(BtnAddFirmTypeToList);
    if TsComboBoxEx(Sender).Name = 'EditNapravlenie' then
    begin
      BtnAddCuratorToListClick(BtnAddNaprToList);
    end;
    if TsEdit(Sender).Name = 'EditWEB' then
      BtnAddWebToListClick(BtnAddWebToList);
    if TsEdit(Sender).Name = 'EditEMAIL' then
      BtnAddWebToListClick(BtnAddEMailToList);
    if TsEdit(Sender).Name = 'EditPhone1' then
      BtnAddPhoneToListClick(BtnAddPhoneToList1);
    if TsEdit(Sender).Name = 'EditPhone2' then
      BtnAddPhoneToListClick(BtnAddPhoneToList2);
    if TsEdit(Sender).Name = 'EditPhone3' then
      BtnAddPhoneToListClick(BtnAddPhoneToList3);
    if TsEdit(Sender).Name = 'EditPhone4' then
      BtnAddPhoneToListClick(BtnAddPhoneToList4);
    if TsEdit(Sender).Name = 'EditPhone5' then
      BtnAddPhoneToListClick(BtnAddPhoneToList5);
    if TsEdit(Sender).Name = 'EditPhone6' then
      BtnAddPhoneToListClick(BtnAddPhoneToList6);
    if TsEdit(Sender).Name = 'EditPhone7' then
      BtnAddPhoneToListClick(BtnAddPhoneToList7);
    if TsEdit(Sender).Name = 'EditPhone8' then
      BtnAddPhoneToListClick(BtnAddPhoneToList8);
    if TsEdit(Sender).Name = 'EditPhone9' then
      BtnAddPhoneToListClick(BtnAddPhoneToList9);
    if TsEdit(Sender).Name = 'EditPhone10' then
      BtnAddPhoneToListClick(BtnAddPhoneToList10);
  end;
end;

procedure TFormEditor.EditCuratorExit(Sender: TObject);
begin
  if TsComboBoxEx(Sender).Name = 'EditCurator' then
    EditCurator.Text := ''
  else if TsComboBoxEx(Sender).Name = 'EditRubr' then
    EditRubr.Text := ''
  else if TsComboBoxEx(Sender).Name = 'EditFirmType' then
    EditFirmType.Text := ''
  else if TsComboBoxEx(Sender).Name = 'EditNapravlenie' then
    EditNapravlenie.Text := ''
  else if TsEdit(Sender).Name = 'EditWEB' then
    EditWEB.Text := ''
  else if TsEdit(Sender).Name = 'EditEMAIL' then
    EditEMAIL.Text := '';
end;

procedure TFormEditor.SGCuratorKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = 46 then
  begin
    Key := 0;
    if TNextGrid(Sender).Name = 'SGCurator' then
      BtnDeleteCuratorFromListClick(BtnDeleteCuratorFromList);
    if TNextGrid(Sender).Name = 'SGRubr' then
      BtnDeleteCuratorFromListClick(BtnDeleteRubrFromList);
    if TNextGrid(Sender).Name = 'SGFirmType' then
      BtnDeleteCuratorFromListClick(BtnDeleteFirmTypeFromList);
    if TNextGrid(Sender).Name = 'SGNapravlenie' then
      BtnDeleteCuratorFromListClick(BtnDeleteNaprFromList);
    if TNextGrid(Sender).Name = 'SGWeb' then
      BtnDeleteCuratorFromListClick(BtnDeleteWebFromList);
    if TNextGrid(Sender).Name = 'SGEMail' then
      BtnDeleteCuratorFromListClick(BtnDeleteEMailFromList);
    if TNextGrid(Sender).Name = 'SGPhone1' then
      BtnDeleteCuratorFromListClick(BtnDeletePhoneFromList1);
    if TNextGrid(Sender).Name = 'SGPhone2' then
      BtnDeleteCuratorFromListClick(BtnDeletePhoneFromList2);
    if TNextGrid(Sender).Name = 'SGPhone3' then
      BtnDeleteCuratorFromListClick(BtnDeletePhoneFromList3);
    if TNextGrid(Sender).Name = 'SGPhone4' then
      BtnDeleteCuratorFromListClick(BtnDeletePhoneFromList4);
    if TNextGrid(Sender).Name = 'SGPhone5' then
      BtnDeleteCuratorFromListClick(BtnDeletePhoneFromList5);
    if TNextGrid(Sender).Name = 'SGPhone6' then
      BtnDeleteCuratorFromListClick(BtnDeletePhoneFromList6);
    if TNextGrid(Sender).Name = 'SGPhone7' then
      BtnDeleteCuratorFromListClick(BtnDeletePhoneFromList7);
    if TNextGrid(Sender).Name = 'SGPhone8' then
      BtnDeleteCuratorFromListClick(BtnDeletePhoneFromList8);
    if TNextGrid(Sender).Name = 'SGPhone9' then
      BtnDeleteCuratorFromListClick(BtnDeletePhoneFromList9);
    if TNextGrid(Sender).Name = 'SGPhone10' then
      BtnDeleteCuratorFromListClick(BtnDeletePhoneFromList10);
  end;
  if Key = 13 then
  begin
    Key := 0;
    if TNextGrid(Sender).Name = 'SGCurator' then
      SGCuratorDblClick(SGCurator);
    if TNextGrid(Sender).Name = 'SGRubr' then
      SGCuratorDblClick(SGRubr);
    if TNextGrid(Sender).Name = 'SGFirmType' then
      SGCuratorDblClick(SGFirmType);
    if TNextGrid(Sender).Name = 'SGNapravlenie' then
      SGCuratorDblClick(SGNapravlenie);
    if TNextGrid(Sender).Name = 'SGWeb' then
      SGCuratorDblClick(SGWeb);
    if TNextGrid(Sender).Name = 'SGEMail' then
      SGCuratorDblClick(SGEMail);
  end;
end;

procedure TFormEditor.SGCuratorDblClick(Sender: TObject);

  procedure Editing(sg: TNextGrid; edit: TObject);
  begin
    if sg.SelectedCount = 0 then
      exit;
    if Assigned(edit) then
    begin
      if edit.ClassName = 'TsEdit' then
      begin
        TsEdit(edit).Text := sg.Cells[0, sg.SelectedRow];
        TsEdit(edit).SetFocus;
      end;
      if edit.ClassName = 'TsComboBoxEx' then
      begin
        TsComboBoxEx(edit).Text := sg.Cells[0, sg.SelectedRow];
        TsComboBoxEx(edit).SetFocus;
      end;
    end;
    sg.DeleteRow(sg.SelectedRow);
  end;

begin
  if TNextGrid(Sender).Name = 'SGCurator' then
    Editing(SGCurator, EditCurator);
  if TNextGrid(Sender).Name = 'SGRubr' then
    Editing(SGRubr, EditRubr);
  if TNextGrid(Sender).Name = 'SGFirmType' then
    Editing(SGFirmType, EditFirmType);
  if TNextGrid(Sender).Name = 'SGNapravlenie' then
    Editing(SGNapravlenie, EditNapravlenie);
  if TNextGrid(Sender).Name = 'SGWeb' then
    Editing(SGWeb, EditWEB);
  if TNextGrid(Sender).Name = 'SGEMail' then
    Editing(SGEMail, EditEMAIL);
end;

procedure TFormEditor.SGCuratorEnter(Sender: TObject);
begin
  if TNextGrid(Sender).RowCount > 0 then
    if TNextGrid(Sender).SelectedCount = 0 then
      TNextGrid(Sender).SelectFirstRow;
end;

procedure TFormEditor.BtnAddPhoneToListClick(Sender: TObject);

  procedure Adding(editPhoneType: TsComboBoxEx; editPhone: TsEdit; sg: TNextGrid);
  var
    tp, ph: string;
  begin
    tp := AnsiLowerCase(editPhoneType.Text);
    ph := AnsiLowerCase(editPhone.Text);
    if Trim(editPhoneType.Text) = '' then
    begin
      MessageBox(handle, 'Не указан тип телефона', 'Предупрежение', MB_OK or MB_ICONWARNING);
      exit;
    end;
    if Trim(editPhone.Text) = '' then
    begin
      MessageBox(handle, 'Не указан телефон', 'Предупрежение', MB_OK or MB_ICONWARNING);
      exit;
    end;
    if sg.FindText(0, '(' + tp + ') ' + ph, [soCaseInsensitive, soExactMatch]) then
    begin
      MessageBox(handle, 'Указанный телефон уже есть в списке', 'Предупрежение', MB_OK or MB_ICONWARNING);
      exit;
    end;
    sg.AddRow;
    sg.Cells[0, sg.LastAddedRow] := '(' + tp + ') ' + ph;
    editPhoneType.Text := '';
    editPhone.Text := '';
    editPhoneType.SetFocus;
  end;

begin
  if TsSpeedButton(Sender).Name = 'BtnAddPhoneToList1' then
    Adding(EditPhoneType1, EditPhone1, SGPhone1);
  if TsSpeedButton(Sender).Name = 'BtnAddPhoneToList2' then
    Adding(EditPhoneType2, EditPhone2, SGPhone2);
  if TsSpeedButton(Sender).Name = 'BtnAddPhoneToList3' then
    Adding(EditPhoneType3, EditPhone3, SGPhone3);
  if TsSpeedButton(Sender).Name = 'BtnAddPhoneToList4' then
    Adding(EditPhoneType4, EditPhone4, SGPhone4);
  if TsSpeedButton(Sender).Name = 'BtnAddPhoneToList5' then
    Adding(EditPhoneType5, EditPhone5, SGPhone5);
  if TsSpeedButton(Sender).Name = 'BtnAddPhoneToList6' then
    Adding(EditPhoneType6, EditPhone6, SGPhone6);
  if TsSpeedButton(Sender).Name = 'BtnAddPhoneToList7' then
    Adding(EditPhoneType7, EditPhone7, SGPhone7);
  if TsSpeedButton(Sender).Name = 'BtnAddPhoneToList8' then
    Adding(EditPhoneType8, EditPhone8, SGPhone8);
  if TsSpeedButton(Sender).Name = 'BtnAddPhoneToList9' then
    Adding(EditPhoneType9, EditPhone9, SGPhone9);
  if TsSpeedButton(Sender).Name = 'BtnAddPhoneToList10' then
    Adding(EditPhoneType10, EditPhone10, SGPhone10);
end;

procedure TFormEditor.EditPhoneKeyPress(Sender: TObject; var Key: Char);
begin
  EditCuratorKeyPress(Sender, Key);
  if not(Key in ['0' .. '9', #8, #32, '+', '-']) then
    Key := #0;
end;

procedure TFormEditor.btnDeleteAdresClick(Sender: TObject);
var
  PageIndex: string;
begin
  PageIndex := StringReplace(TsSpeedButton(Sender).Name, 'btnDeleteAdres', '', []);
  ClearAdresEdits(StrToIntDef(PageIndex, -1));
end;

procedure TFormEditor.LoadDataEditor;
var
  i: Integer;
  Q: TIBCQuery;
  Buffer: TStringList;
begin
  FormLogo.sLabel1.Caption := 'Подключаемые модули ...';
  Q := QueryCreate;
  Buffer := TStringList.Create;

  try
    for i := 0 to Main.sgCurator_tmp.RowCount - 1 do
    begin
      Buffer.AddObject(Main.sgCurator_tmp.Cells[0, i], TObject(StrToInt(Main.sgCurator_tmp.Cells[1, i])));
    end;
    EditCurator.Items := Buffer; // КУРАТОР
    FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
    Application.ProcessMessages;

    Buffer.Clear;
    for i := 0 to Main.sgRubr_tmp.RowCount - 1 do
    begin
      Buffer.AddObject(Main.sgRubr_tmp.Cells[0, i], TObject(StrToInt(Main.sgRubr_tmp.Cells[1, i])));
    end;
    EditRubr.Items := Buffer; // РУБРИКА
    FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
    Application.ProcessMessages;

    Buffer.Clear; // ТИП
    for i := 0 to Main.sgFirmType_tmp.RowCount - 1 do
    begin
      Buffer.AddObject(Main.sgFirmType_tmp.Cells[0, i], TObject(StrToInt(Main.sgFirmType_tmp.Cells[1, i])));
    end;
    EditFirmType.Items := Buffer;
    FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
    Application.ProcessMessages;

    Buffer.Clear; // НАПРАВЛЕНИЕ
    for i := 0 to Main.sgNapr_tmp.RowCount - 1 do
    begin
      Buffer.AddObject(Main.sgNapr_tmp.Cells[0, i], TObject(StrToInt(Main.sgNapr_tmp.Cells[1, i])));
    end;
    EditNapravlenie.Items := Buffer;
    FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
    Application.ProcessMessages;

    Q.Close; // ТИП ОФИСА
    Q.SQL.Text := 'select * from OFFICETYPE order by lower(NAME)';
    Q.Open;
    Q.FetchAll := True;
    Buffer.Clear;
    while not Q.Eof do
    begin
      Buffer.AddObject(Q.FieldValues['NAME'], TObject(Integer(Q.FieldValues['ID'])));
      Q.Next;
    end;
    for i := 1 to 10 do
      TsComboBoxEx(FindComponent('EditOfficeType' + IntToStr(i))).Items := Buffer;
    FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
    Application.ProcessMessages;

    Q.Close; // СТРАНЫ
    Q.SQL.Text := 'select * from COUNTRY order by lower(NAME)';
    Q.Open;
    Q.FetchAll := True;
    Buffer.Clear;
    while not Q.Eof do
    begin
      Buffer.AddObject(Q.FieldValues['NAME'], TObject(Integer(Q.FieldValues['ID'])));
      Q.Next;
    end;
    for i := 1 to 10 do
      TsComboBoxEx(FindComponent('EditCountry' + IntToStr(i))).Items := Buffer;
    FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
    Application.ProcessMessages;

    Q.Close; // ОБЛАСТИ
    Q.SQL.Text := 'select * from REGION order by lower(NAME)';
    Q.Open;
    Q.FetchAll := True;
    Buffer.Clear;
    while not Q.Eof do
    begin
      Buffer.AddObject(Q.FieldValues['NAME'], TObject(Integer(Q.FieldValues['ID'])));
      Q.Next;
    end;
    for i := 1 to 10 do
      TsComboBoxEx(FindComponent('EditRegion' + IntToStr(i))).Items := Buffer;
    FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
    Application.ProcessMessages;

    Q.Close; // ГОРОДА
    Q.SQL.Text := 'select * from CITY order by lower(NAME)';
    Q.Open;
    Q.FetchAll := True;
    Buffer.Clear;
    while not Q.Eof do
    begin
      Buffer.AddObject(Q.FieldValues['NAME'], TObject(Integer(Q.FieldValues['ID'])));
      Q.Next;
    end;
    for i := 1 to 10 do
      TsComboBoxEx(FindComponent('EditCity' + IntToStr(i))).Items := Buffer;
    FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
    Application.ProcessMessages;

    Q.Close; // ТИПЫ ТЕЛЕФОНОВ
    Q.SQL.Text := 'select * from PHONETYPE order by lower(NAME)';
    Q.Open;
    Q.FetchAll := True;
    Buffer.Clear;
    while not Q.Eof do
    begin
      Buffer.AddObject(Q.FieldValues['NAME'], TObject(Integer(Q.FieldValues['ID'])));
      Q.Next;
    end;
    for i := 1 to 10 do
      TsComboBoxEx(FindComponent('EditPhoneType' + IntToStr(i))).Items := Buffer;
    FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
    Application.ProcessMessages;

  finally
    Q.Free; // FormMain.IBDatabase1.Close;
    Buffer.Free;
  end;
end;

function TFormEditor.GetIDString(component: TNextGrid): string; { #1$#2$#3$ }
var
  i: Integer;
  tmp, id: string;
begin
  Result := '';
  if component.RowCount = 0 then
    exit;
  for i := 0 to component.RowCount - 1 do
  begin
    id := component.Cells[1, i];
    if component.Name = 'SGCurator' then
      if Main.sgCurator_tmp.FindText(1, id, [soCaseInsensitive, soExactMatch]) then
        tmp := tmp + '#' + Main.sgCurator_tmp.Cells[1, Main.sgCurator_tmp.SelectedRow] + '$';
    if component.Name = 'SGRubr' then
      if Main.sgRubr_tmp.FindText(1, id, [soCaseInsensitive, soExactMatch]) then
        tmp := tmp + '#' + Main.sgRubr_tmp.Cells[1, Main.sgRubr_tmp.SelectedRow] + '$';
    if component.Name = 'SGFirmType' then
      if Main.sgFirmType_tmp.FindText(1, id, [soCaseInsensitive, soExactMatch]) then
        tmp := tmp + '#' + Main.sgFirmType_tmp.Cells[1, Main.sgFirmType_tmp.SelectedRow] + '$';
    if component.Name = 'SGNapravlenie' then
      if Main.sgNapr_tmp.FindText(1, id, [soCaseInsensitive, soExactMatch]) then
        tmp := tmp + '#' + Main.sgNapr_tmp.Cells[1, Main.sgNapr_tmp.SelectedRow] + '$';
    if component.Name = 'SGWeb' then
      tmp := tmp + SGWeb.Cells[0, i] + ', ';
    if component.Name = 'SGEMail' then
      tmp := tmp + SGEMail.Cells[0, i] + ', ';
  end;
  Result := tmp;
end;

function TFormEditor.GetWebEmailString(component: TNextGrid): string;
var
  i: Integer;
  tmp: string;
begin
  Result := '';
  if component.RowCount = 0 then
    exit;
  for i := 0 to component.RowCount - 1 do
  begin
    if component.Name = 'SGWeb' then
      tmp := tmp + SGWeb.Cells[0, i] + ', ';
    if component.Name = 'SGEMail' then
      tmp := tmp + SGEMail.Cells[0, i] + ', ';
  end;
  tmp := Trim(tmp);
  if length(tmp) > 0 then
    if tmp[length(tmp)] = ',' then
      delete(tmp, length(tmp), 1);
  Result := tmp;
end;

procedure TFormEditor.ClearEdits;
var
  i: Integer;
begin
  lblID.Caption := '';
  SGCurator.ClearRows;
  SGRubr.ClearRows;
  SGFirmType.ClearRows;
  SGNapravlenie.ClearRows;
  SGWeb.ClearRows;
  SGEMail.ClearRows;
  CBActivity.Checked := True;
  cbDataRelevance.Checked := True;
  sPageControl1.ActivePageIndex := 0;
  EditName.Text := '';
  EditFIO.Text := '';
  ClearEdit(EditCurator);
  ClearEdit(EditRubr);
  ClearEdit(EditFirmType);
  ClearEdit(EditNapravlenie);
  EditWEB.Text := '';
  EditEMAIL.Text := '';

  for i := 1 to 10 do
    ClearAdresEdits(i);

  if not Assigned(list_GC_IDs) then
    list_GC_IDs := TStringList.Create;
  list_GC_IDs.Clear;
end;

function GetValidRegionOfCity(ID_City, ID_Region: integer): integer;
var
  Query: TIBCQuery;
  ID_Region_Valid: integer;
begin
  Result := ID_Region;

  if ID_City > ID_UNKNOWN then
  begin
    Query := QueryCreate;
    try
      Query.SQL.Text := 'select ID_REGION from CITY where ID = :ID';
      Query.ParamByName('ID').AsInteger := ID_City;
      Query.Open;
      if Query.RecordCount > 0 then
        ID_Region_Valid := Query.FieldByName('ID_REGION').AsInteger
      else
        ID_Region_Valid := ID_UNKNOWN;

      if ID_Region_Valid <> ID_Region then
      begin
        Result := ID_Region_Valid;
        debug('REGION IS INVALID. Passed = %s | Expected = %s', [GetNameByID('REGION', ID_Region.ToString),
          GetNameByID('REGION', ID_Region_Valid.ToString)], 2);
      end;
    finally
      Query.Free;
    end;
  end;
end;

{ #TODO1: IMPORTANT : Still need to somehow let user know that entered values to country,region,city controls are not valid }
procedure TFormEditor.AddRecord(Sender: TObject);
var
  str, phones: TStrings;

  procedure AdresProcs(CBAdres: TsCheckBox; No: TsEdit; OfficeType: TsComboBoxEx; ZIP: TsEdit; Street: TsEdit;
    Country, Region, City: TsComboBoxEx; MemoPhone: TsMemo; SGPhone: TNextGrid);
  var
    i, ID_Region_Valid: Integer;
    offtype_id, country_id, region_id, city_id: string;
  begin
    if CBAdres.Checked then
    begin
      if OfficeType.GetID = ID_UNKNOWN then
        offtype_id := EmptyStr
      else
        offtype_id := '@' + OfficeType.GetID.ToString;

      if Country.GetID = ID_UNKNOWN then
        country_id := EmptyStr
      else
        country_id := '&' + Country.GetID.ToString;

      ID_Region_Valid := GetValidRegionOfCity(City.GetID, Region.GetID);
      if ID_Region_Valid = ID_UNKNOWN then
        region_id := EmptyStr
      else
        region_id := '*' + ID_Region_Valid.ToString;

      if City.GetID = ID_UNKNOWN then
        city_id := EmptyStr
      else
        city_id := '^' + City.GetID.ToString;

      str.Add(Format('#%s$#%s$#%s$#%s$#%s$#%s$#%s$#%s$', ['1', No.Text, offtype_id, Trim(ZIP.Text), Trim(Street.Text), country_id,
        region_id, city_id]));
      if SGPhone.RowCount > 0 then
        for i := 0 to SGPhone.RowCount - 1 do
          MemoPhone.Lines.Add(SGPhone.Cells[0, i]);
      phones.Add('#' + No.Text + '$#' + Trim(MemoPhone.Text) + '$')
    end
    else
    begin
      str.Add('#0$#' + No.Text + '$#$#$#$#$#$#$');
      phones.Add('#' + No.Text + '$#$');
    end;
  end;

{ #TODO1: REVISIT : see if IsAdresFilled logic is up to date after region field has been added. Since fields like country,region,city
  are no longer auto-created I need a way to tell user that entered value is not valid. Mard edit with "red color" or icon stating that
  might be solution }
  procedure IsAdresFilled(CBAdres: TsCheckBox; OfficeType: TsComboBoxEx; ZIP: TsEdit; Street: TsEdit; Country, Region, City: TsComboBoxEx;
    SGPhone: TNextGrid);
  begin
    CBAdres.Checked := False;
    if Trim(OfficeType.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if Trim(ZIP.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if Trim(Street.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if Trim(Country.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if Trim(Region.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if Trim(City.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if SGPhone.RowCount > 0 then
    begin
      CBAdres.Checked := True;
      exit;
    end;
  end;

  function BuildRelations: string;
  var
    x, z: Integer;
    idRUBR, idNAPR, finalSTR: string;
  begin
    Result := '';
    if (SGRubr.RowCount = 0) or (SGNapravlenie.RowCount = 0) then
      exit;
    for x := 0 to SGRubr.RowCount - 1 do
    begin
      idRUBR := SGRubr.Cells[1, x];
      finalSTR := finalSTR + '#' + idRUBR + '=';
      for z := 0 to SGNapravlenie.RowCount - 1 do
      begin
        idNAPR := SGNapravlenie.Cells[1, z];
        finalSTR := finalSTR + '(' + idNAPR + ')';
      end;
      finalSTR := finalSTR + '$';
    end;
    Result := finalSTR;
    // result = #13=(534)(654)(823)$#27=(43)(534)(546)$
  end;

var
  id, NewRubr, tmp, tmp2, rubrID, FieldAdresText: string;
  nd, nd2: TTreeNode;
  listNewRecords: TStringList;
begin
  if Trim(EditName.Text) = '' then
  begin
    MessageBox(handle, 'Необходимо указать название.', 'Информация', MB_OK or MB_ICONWARNING);
    exit;
  end;
  if SGRubr.RowCount = 0 then
  begin
    MessageBox(handle, 'Необходимо указать рубрику.', 'Информация', MB_OK or MB_ICONWARNING);
    exit;
  end;

  if not(IsDublicate) then
  begin

    // shit code forced me to put this notify here
    // if run this above IsDublicate = false check then it will run twice
    // if run this below IsDublicate = false check then it pop after dublicate window and potentialy lose all data
    if IsNewRecordFound_Notify(listNewRecords) = true then
    begin
      if MessageBox(handle, PChar('Новые директории будут созданы. Продолжить?' + #13#10 + #13#10 + listNewRecords.Text), 'Подтверждение',
        MB_YESNO or MB_ICONQUESTION) = MRNO then
      begin
        listNewRecords.Free;
        exit;
      end;
    end;
    listNewRecords.Free;

    IsRecordDublicate;
    exit;
    // AddRecord запустится снова из IsRecordDublicate или после нажатия
    // кнопки "Продолжить" на форме FormDublicate, но уже с IsDublicate = True
  end;
  if FormDublicate.Visible then
    FormDublicate.Close;

  IsNewRecordCheck;

  EditName.Text := UpperFirst(EditName.Text);
  EditFIO.Text := UpperFirst(EditFIO.Text);
  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text :=
    'insert into BASE (ACTIVITY,RELEVANCE,NAME,FIO,CURATOR,RUBR,FIRMTYPE,NAPRAVLENIE,WEB,EMAIL,ADRES,PHONES,DATE_ADDED,DATE_EDITED,RELATIONS) values '
    + '(:ACTIVITY,:RELEVANCE,:NAME,:FIO,:CURATOR,:RUBR,:FIRMTYPE,:NAPRAVLENIE,:WEB,:EMAIL,:ADRES,:PHONES,:DATE_ADDED,:DATE_EDITED,:RELATIONS) returning ID';
  FormMain.IBQuery1.ParamByName('ACTIVITY').AsBoolean := CBActivity.Checked;
  FormMain.IBQuery1.ParamByName('RELEVANCE').AsBoolean := cbDataRelevance.Checked;
  FormMain.IBQuery1.ParamByName('NAME').AsString := EditName.Text;
  FormMain.IBQuery1.ParamByName('FIO').AsString := EditFIO.Text;
  FormMain.IBQuery1.ParamByName('CURATOR').AsString := GetIDString(SGCurator);
  FormMain.IBQuery1.ParamByName('RUBR').AsString := GetIDString(SGRubr);
  FormMain.IBQuery1.ParamByName('FIRMTYPE').AsString := GetIDString(SGFirmType);
  FormMain.IBQuery1.ParamByName('NAPRAVLENIE').AsString := GetIDString(SGNapravlenie);
  FormMain.IBQuery1.ParamByName('WEB').AsString := GetWebEmailString(SGWeb);
  FormMain.IBQuery1.ParamByName('EMAIL').AsString := GetWebEmailString(SGEMail);

  str := TStringList.Create;
  phones := TStringList.Create;
  try
    IsAdresFilled(CBAdres1, EditOfficeType1, EditZIP1, EditStreet1, EditCountry1, EditRegion1, EditCity1, SGPhone1);
    IsAdresFilled(CBAdres2, EditOfficeType2, EditZIP2, EditStreet2, EditCountry2, EditRegion2, EditCity2, SGPhone2);
    IsAdresFilled(CBAdres3, EditOfficeType3, EditZIP3, EditStreet3, EditCountry3, EditRegion3, EditCity3, SGPhone3);
    IsAdresFilled(CBAdres4, EditOfficeType4, EditZIP4, EditStreet4, EditCountry4, EditRegion4, EditCity4, SGPhone4);
    IsAdresFilled(CBAdres5, EditOfficeType5, EditZIP5, EditStreet5, EditCountry5, EditRegion5, EditCity5, SGPhone5);
    IsAdresFilled(CBAdres6, EditOfficeType6, EditZIP6, EditStreet6, EditCountry6, EditRegion6, EditCity6, SGPhone6);
    IsAdresFilled(CBAdres7, EditOfficeType7, EditZIP7, EditStreet7, EditCountry7, EditRegion7, EditCity7, SGPhone7);
    IsAdresFilled(CBAdres8, EditOfficeType8, EditZIP8, EditStreet8, EditCountry8, EditRegion8, EditCity8, SGPhone8);
    IsAdresFilled(CBAdres9, EditOfficeType9, EditZIP9, EditStreet9, EditCountry9, EditRegion9, EditCity9, SGPhone9);
    IsAdresFilled(CBAdres10, EditOfficeType10, EditZIP10, EditStreet10, EditCountry10, EditRegion10, EditCity10, SGPhone10);
    AdresProcs(CBAdres1, EditNO1, EditOfficeType1, EditZIP1, EditStreet1, EditCountry1, EditRegion1, EditCity1, MemoPhone1, SGPhone1);
    AdresProcs(CBAdres2, EditNO2, EditOfficeType2, EditZIP2, EditStreet2, EditCountry2, EditRegion2, EditCity2, MemoPhone2, SGPhone2);
    AdresProcs(CBAdres3, EditNO3, EditOfficeType3, EditZIP3, EditStreet3, EditCountry3, EditRegion3, EditCity3, MemoPhone3, SGPhone3);
    AdresProcs(CBAdres4, EditNO4, EditOfficeType4, EditZIP4, EditStreet4, EditCountry4, EditRegion4, EditCity4, MemoPhone4, SGPhone4);
    AdresProcs(CBAdres5, EditNO5, EditOfficeType5, EditZIP5, EditStreet5, EditCountry5, EditRegion5, EditCity5, MemoPhone5, SGPhone5);
    AdresProcs(CBAdres6, EditNO6, EditOfficeType6, EditZIP6, EditStreet6, EditCountry6, EditRegion6, EditCity6, MemoPhone6, SGPhone6);
    AdresProcs(CBAdres7, EditNO7, EditOfficeType7, EditZIP7, EditStreet7, EditCountry7, EditRegion7, EditCity7, MemoPhone7, SGPhone7);
    AdresProcs(CBAdres8, EditNO8, EditOfficeType8, EditZIP8, EditStreet8, EditCountry8, EditRegion8, EditCity8, MemoPhone8, SGPhone8);
    AdresProcs(CBAdres9, EditNO9, EditOfficeType9, EditZIP9, EditStreet9, EditCountry9, EditRegion9, EditCity9, MemoPhone9, SGPhone9);
    AdresProcs(CBAdres10, EditNO10, EditOfficeType10, EditZIP10, EditStreet10, EditCountry10, EditRegion10, EditCity10, MemoPhone10,
      SGPhone10);

    FieldAdresText := str.Text;
    FormMain.IBQuery1.ParamByName('ADRES').AsString := FieldAdresText;
    FormMain.IBQuery1.ParamByName('PHONES').AsString := phones.Text;
    FormMain.IBQuery1.ParamByName('DATE_ADDED').AsString := DateToStr(Now);
    FormMain.IBQuery1.ParamByName('DATE_EDITED').AsString := DateToStr(Now);
    FormMain.IBQuery1.ParamByName('RELATIONS').AsString := BuildRelations;
  finally
    str.Free;
    phones.Free;
  end;

  try
    FormMain.IBQuery1.Execute;
    id := FormMain.IBQuery1.ParamByName('RET_ID').AsString;

    DoNaprActivityValidation(EmptyStr, GetIDString(SGNapravlenie), EmptyStr, CBActivity.Checked, 'add');

    DoCityActivityValidation(EmptyStr, FieldAdresText, EmptyStr, CBActivity.Checked, 'add');

    WriteLog('TFormEditor.AddRecord: запись добавлена ' + id);
  except
    on E: Exception do
    begin
      WriteLog('TFormEditor.AddRecord' + #13 + 'Ошибка: ' + E.Message);
      MessageBox(handle, PChar('Ошибка при создании фирмы.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      exit;
    end;
  end;
  FormMain.IBTransaction1.CommitRetaining;
  FormMain.sStatusBar1.Panels[1].Text := 'Фирм в базе: ' + GetFirmCount;

  NewRubr := GetIDString(SGRubr);
  tmp2 := NewRubr;
  while pos('$', NewRubr) > 0 do // создали новые
  begin
    tmp := copy(NewRubr, 0, pos('$', NewRubr));
    delete(NewRubr, 1, length(tmp));
    delete(tmp, 1, 1);
    delete(tmp, length(tmp), 1);
    nd := FormMain.SearchNode(FormMain.TVRubrikator, StrToInt(tmp), 0);
    if nd <> nil then
    begin
      nd2 := FormMain.TVRubrikator.Items.AddChildObject(nd, EditName.Text, Pointer(StrToInt(id)));
      if CBActivity.Checked then
      begin
        nd2.ImageIndex := 1;
        nd2.SelectedIndex := 1;
      end
      else
      begin
        nd2.ImageIndex := 2;
        nd2.SelectedIndex := 2;
      end;
    end;
  end;
  if FormMain.TVRubrikator.Selected <> nil then
  begin { обновляем SGGeneral если показана рубрика в которую идет новый итем }
    rubrID := IntToStr(Integer(Pointer(FormMain.TVRubrikator.Selected.Data)));
    if pos('#' + rubrID + '$', tmp2) > 0 then
      FormMain.TVRubrikatorChange(FormMain.TVRubrikator, FormMain.TVRubrikator.Selected);
  end;
  FormMain.TVRubrikator.CustomSort(@Helpers.CustomSortProc, 0, True);
  FormMain.IBQuery1.Close;

  // might already be closed from @IsRecordDublicate if any dublicates were found
  if FormEditor.Visible = true then
    FormEditor.Close;

  FormRelations.ClearEdits;
  FormRelations.PrepareRecord(id);
  FormRelations.Show;
  // FormMain.IBDatabase1.Connected := False;
end;

procedure TFormEditor.PrepareEditRecord(id: string);

  procedure AdresProcs(AdresList: TStrings; Num: Integer; CBAdres: TsCheckBox; OfficeType: TsComboBoxEx; ZIP, Street: TsEdit;
    Country, Region, City: TsComboBoxEx);
  var
    ID_OfficeType, ID_Country, ID_Region, ID_City: integer;
    list: TStrings;
  begin
    if length(AdresList.Text) = 0 then
      exit;
    // такая же процедура в Main.OpenTabByID и Editor.PrepareEdit и Report.GenerateReport и MailSend.SendRegInfoCheck
    // list2[0] = CBAdres; list2[1] = NO; list2[2] = OfficeType; list2[3] = ZIP;
    // list2[4] = Street; list2[5] = Country; list2[6] = Region; list2[7] = City;
    list := ParseAdresFieldToEntriesList(AdresList[Num - 1]);
    if list[0] = '0' then
    begin
      CBAdres.Checked := False;
      list.Free;
      exit;
    end;
    if list[0] = '1' then
      CBAdres.Checked := True;
    ZIP.Text := list[3];
    Street.Text := list[4];
    ID_OfficeType := StrToIntDef(list[2], -1);
    ID_Country := StrToIntDef(list[5], -1);
    ID_Region := StrToIntDef(list[6], -1);
    ID_City := StrToIntDef(list[7], -1);
    OfficeType.SetIndexOfObject(ID_OfficeType);
    Country.SetIndexOfObject(ID_Country);
    Region.SetIndexOfObject(ID_Region);
    City.SetIndexOfObject(ID_City);
    list.Free;
  end;

  procedure PhonesProcs(PhonesList: TStrings);

    procedure Loading(MemoPhone: TsMemo; SGPhone: TNextGrid; PhoneText: string);
    var
      i: Integer;
    begin
      MemoPhone.Text := PhoneText;
      for i := 0 to MemoPhone.Lines.Count - 1 do
      begin
        SGPhone.AddRow;
        SGPhone.Cells[0, SGPhone.LastAddedRow] := MemoPhone.Lines[i];
      end;
      MemoPhone.Clear;
    end;

  var
    ph_list, ph_tmp, ph_tmp2: WideString;
  begin
    ph_list := PhonesList.Text;
    while pos('$', ph_list) > 0 do
    begin
      // ph_tmp = NUM; ph_tmp2 = PHONES;
      ph_tmp := copy(ph_list, 0, pos('$', ph_list));
      delete(ph_list, 1, length(ph_tmp));
      ph_tmp := AnsiReplaceText(ph_tmp, Chr(13), '');
      ph_tmp := AnsiReplaceText(ph_tmp, Chr(10), '');
      delete(ph_tmp, 1, 1);
      delete(ph_tmp, length(ph_tmp), 1);
      ph_tmp2 := copy(ph_list, 0, pos('$', ph_list));
      delete(ph_list, 1, length(ph_tmp2));
      delete(ph_tmp2, 1, 1);
      delete(ph_tmp2, length(ph_tmp2), 1);
      if (ph_tmp = '1') and (CBAdres1.Checked) then
        Loading(MemoPhone1, SGPhone1, ph_tmp2);
      if (ph_tmp = '2') and (CBAdres2.Checked) then
        Loading(MemoPhone2, SGPhone2, ph_tmp2);
      if (ph_tmp = '3') and (CBAdres3.Checked) then
        Loading(MemoPhone3, SGPhone3, ph_tmp2);
      if (ph_tmp = '4') and (CBAdres4.Checked) then
        Loading(MemoPhone4, SGPhone4, ph_tmp2);
      if (ph_tmp = '5') and (CBAdres5.Checked) then
        Loading(MemoPhone5, SGPhone5, ph_tmp2);
      if (ph_tmp = '6') and (CBAdres6.Checked) then
        Loading(MemoPhone6, SGPhone6, ph_tmp2);
      if (ph_tmp = '7') and (CBAdres7.Checked) then
        Loading(MemoPhone7, SGPhone7, ph_tmp2);
      if (ph_tmp = '8') and (CBAdres8.Checked) then
        Loading(MemoPhone8, SGPhone8, ph_tmp2);
      if (ph_tmp = '9') and (CBAdres9.Checked) then
        Loading(MemoPhone9, SGPhone9, ph_tmp2);
      if (ph_tmp = '10') and (CBAdres10.Checked) then
        Loading(MemoPhone10, SGPhone10, ph_tmp2);
    end;
  end;

  procedure LoadDataToGrids(component: TNextGrid; str: string);
  var
    tmp, Name, id: string;
  begin
    if length(Trim(str)) = 0 then
      exit;
    component.BeginUpdate;
    while pos('$', str) > 0 do
    begin
      tmp := copy(str, 0, pos('$', str));
      delete(str, 1, length(tmp));
      delete(tmp, 1, 1);
      delete(tmp, length(tmp), 1);
      if component.Name = 'SGCurator' then
        if Main.sgCurator_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
        begin
          name := Main.sgCurator_tmp.Cells[0, Main.sgCurator_tmp.SelectedRow];
          id := Main.sgCurator_tmp.Cells[1, Main.sgCurator_tmp.SelectedRow];
        end;
      if component.Name = 'SGRubr' then
        if Main.sgRubr_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
        begin
          name := Main.sgRubr_tmp.Cells[0, Main.sgRubr_tmp.SelectedRow];
          id := Main.sgRubr_tmp.Cells[1, Main.sgRubr_tmp.SelectedRow];
        end;
      if component.Name = 'SGFirmType' then
        if Main.sgFirmType_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
        begin
          name := Main.sgFirmType_tmp.Cells[0, Main.sgFirmType_tmp.SelectedRow];
          id := Main.sgFirmType_tmp.Cells[1, Main.sgFirmType_tmp.SelectedRow];
        end;
      if component.Name = 'SGNapravlenie' then
        if Main.sgNapr_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
        begin
          name := Main.sgNapr_tmp.Cells[0, Main.sgNapr_tmp.SelectedRow];
          id := Main.sgNapr_tmp.Cells[1, Main.sgNapr_tmp.SelectedRow];
        end;
      if (Trim(name) <> '') and (Trim(id) <> '') then
      begin
        component.AddRow;
        component.Cells[0, component.LastAddedRow] := name;
        component.Cells[1, component.LastAddedRow] := id;
      end;
    end;
    component.Resort;
    component.EndUpdate;
  end;

  procedure LoadWebEmailToGrids(component: TNextGrid; str: string);
  var
    tmp: string;
  begin
    str := Trim(str);
    if length(str) = 0 then
      exit;
    if str[length(str)] <> ',' then
      str := str + ',';
    while pos(',', str) > 0 do
    begin
      tmp := copy(str, 0, pos(',', str));
      delete(str, 1, length(tmp));
      delete(tmp, length(tmp), 1);
      tmp := Trim(tmp);
      if tmp <> '' then
      begin
        component.AddRow;
        component.Cells[0, component.LastAddedRow] := tmp;
      end;
    end;
    component.Resort;
  end;

var
  str, phones: TStrings;
begin
  lblID.Caption := id;
  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text := 'select * from BASE where ID = :ID';
  FormMain.IBQuery1.ParamByName('ID').AsString := id;
  FormMain.IBQuery1.Open;
  FormMain.IBQuery1.FetchAll := True;
  if FormMain.IBQuery1.RecordCount = 0 then
    exit;

  // Build list_GC_IDs that is used for DoGarbageCollection
  if FormMain.IBQuery1.FieldValues['CURATOR'] <> null then
    list_GC_IDs.Add('CURATOR=' + FormMain.IBQuery1.FieldValues['CURATOR']);
  if FormMain.IBQuery1.FieldValues['RUBR'] <> null then
    list_GC_IDs.Add('RUBRIKATOR=' + FormMain.IBQuery1.FieldValues['RUBR']);
  if FormMain.IBQuery1.FieldValues['FIRMTYPE'] <> null then
    list_GC_IDs.Add('FIRMTYPE=' + FormMain.IBQuery1.FieldValues['FIRMTYPE']);
  if FormMain.IBQuery1.FieldValues['NAPRAVLENIE'] <> null then
    list_GC_IDs.Add('NAPRAVLENIE=' + FormMain.IBQuery1.FieldValues['NAPRAVLENIE']);
  if FormMain.IBQuery1.FieldValues['ADRES'] <> null then
    list_GC_IDs.Add('ADRES=' + FormMain.IBQuery1.FieldValues['ADRES']);

  if FormMain.IBQuery1.FieldValues['ACTIVITY'] <> null then
    CBActivity.Checked := FormMain.IBQuery1.FieldValues['ACTIVITY']
  else
    CBActivity.Checked := False;
  if FormMain.IBQuery1.FieldValues['RELEVANCE'] <> null then
    cbDataRelevance.Checked := FormMain.IBQuery1.FieldValues['RELEVANCE']
  else
    cbDataRelevance.Checked := False;
  if FormMain.IBQuery1.FieldValues['NAME'] <> null then
    EditName.Text := FormMain.IBQuery1.FieldValues['NAME'];
  if FormMain.IBQuery1.FieldValues['FIO'] <> null then
    EditFIO.Text := FormMain.IBQuery1.FieldValues['FIO'];
  if FormMain.IBQuery1.FieldValues['CURATOR'] <> null then
    LoadDataToGrids(SGCurator, FormMain.IBQuery1.FieldValues['CURATOR']);
  if FormMain.IBQuery1.FieldValues['RUBR'] <> null then
    LoadDataToGrids(SGRubr, FormMain.IBQuery1.FieldValues['RUBR']);
  if FormMain.IBQuery1.FieldValues['FIRMTYPE'] <> null then
    LoadDataToGrids(SGFirmType, FormMain.IBQuery1.FieldValues['FIRMTYPE']);
  if FormMain.IBQuery1.FieldValues['NAPRAVLENIE'] <> null then
    LoadDataToGrids(SGNapravlenie, FormMain.IBQuery1.FieldValues['NAPRAVLENIE']);
  if FormMain.IBQuery1.FieldValues['WEB'] <> null then
    LoadWebEmailToGrids(SGWeb, FormMain.IBQuery1.FieldValues['WEB']);
  if FormMain.IBQuery1.FieldValues['EMAIL'] <> null then
    LoadWebEmailToGrids(SGEMail, FormMain.IBQuery1.FieldValues['EMAIL']);
  if FormMain.IBQuery1.FieldValues['ADRES'] <> null then
  begin
    str := TStringList.Create;
    phones := TStringList.Create;
    str.Text := FormMain.IBQuery1.FieldByName('ADRES').AsVariant;
    phones.Text := FormMain.IBQuery1.FieldByName('PHONES').AsVariant;
    AdresProcs(str, 1, CBAdres1, EditOfficeType1, EditZIP1, EditStreet1, EditCountry1, EditRegion1, EditCity1);
    AdresProcs(str, 2, CBAdres2, EditOfficeType2, EditZIP2, EditStreet2, EditCountry2, EditRegion2, EditCity2);
    AdresProcs(str, 3, CBAdres3, EditOfficeType3, EditZIP3, EditStreet3, EditCountry3, EditRegion3, EditCity3);
    AdresProcs(str, 4, CBAdres4, EditOfficeType4, EditZIP4, EditStreet4, EditCountry4, EditRegion4, EditCity4);
    AdresProcs(str, 5, CBAdres5, EditOfficeType5, EditZIP5, EditStreet5, EditCountry5, EditRegion5, EditCity5);
    AdresProcs(str, 6, CBAdres6, EditOfficeType6, EditZIP6, EditStreet6, EditCountry6, EditRegion6, EditCity6);
    AdresProcs(str, 7, CBAdres7, EditOfficeType7, EditZIP7, EditStreet7, EditCountry7, EditRegion7, EditCity7);
    AdresProcs(str, 8, CBAdres8, EditOfficeType8, EditZIP8, EditStreet8, EditCountry8, EditRegion8, EditCity8);
    AdresProcs(str, 9, CBAdres9, EditOfficeType9, EditZIP9, EditStreet9, EditCountry9, EditRegion9, EditCity9);
    AdresProcs(str, 10, CBAdres10, EditOfficeType10, EditZIP10, EditStreet10, EditCountry10, EditRegion10, EditCity10);
    PhonesProcs(phones);
    str.Free;
    phones.Free;
  end;
  FormMain.IBQuery1.Close;
  // FormMain.IBDatabase1.Connected := False;
end;

procedure TFormEditor.EditRecord(Sender: TObject);
var
  str, phones: TStrings;
  NewRubr, rubrID, tmp, FieldAdresText: string;
  nd, nd2: TTreeNode;
  i: Integer;
  listNewRecords: TStringList;

  procedure AdresProcs(CBAdres: TsCheckBox; No: TsEdit; OfficeType: TsComboBoxEx; ZIP: TsEdit; Street: TsEdit;
    Country, Region, City: TsComboBoxEx; MemoPhone: TsMemo; SGPhone: TNextGrid);
  var
    i, ID_Region_Valid: Integer;
    offtype_id, country_id, region_id, city_id: string;
  begin
    if CBAdres.Checked then
    begin
      if OfficeType.GetID = ID_UNKNOWN then
        offtype_id := EmptyStr
      else
        offtype_id := '@' + OfficeType.GetID.ToString;

      if Country.GetID = ID_UNKNOWN then
        country_id := EmptyStr
      else
        country_id := '&' + Country.GetID.ToString;

      ID_Region_Valid := GetValidRegionOfCity(City.GetID, Region.GetID);
      if ID_Region_Valid = ID_UNKNOWN then
        region_id := EmptyStr
      else
        region_id := '*' + ID_Region_Valid.ToString;

      if City.GetID = ID_UNKNOWN then
        city_id := EmptyStr
      else
        city_id := '^' + City.GetID.ToString;

      str.Add(Format('#%s$#%s$#%s$#%s$#%s$#%s$#%s$#%s$', ['1', No.Text, offtype_id, Trim(ZIP.Text), Trim(Street.Text), country_id,
        region_id, city_id]));
      if SGPhone.RowCount > 0 then
        for i := 0 to SGPhone.RowCount - 1 do
          MemoPhone.Lines.Add(SGPhone.Cells[0, i]);
      phones.Add('#' + No.Text + '$#' + Trim(MemoPhone.Text) + '$')
    end
    else
    begin
      str.Add('#0$#' + No.Text + '$#$#$#$#$#$#$');
      phones.Add('#' + No.Text + '$#$');
    end;
  end;

  procedure IsAdresFilled(CBAdres: TsCheckBox; OfficeType: TsComboBoxEx; ZIP: TsEdit; Street: TsEdit; Country, Region, City: TsComboBoxEx;
    SGPhone: TNextGrid);
  begin
    CBAdres.Checked := False;
    if Trim(OfficeType.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if Trim(ZIP.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if Trim(Street.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if Trim(Country.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if Trim(Region.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if Trim(City.Text) <> '' then
    begin
      CBAdres.Checked := True;
      exit;
    end;
    if SGPhone.RowCount > 0 then
    begin
      CBAdres.Checked := True;
      exit;
    end;
  end;

  function BuildRelations: string;
  var
    x, z: Integer;
    idRUBR, idNAPR, finalSTR: string;
  begin
    Result := '';
    if (SGRubr.RowCount = 0) or (SGNapravlenie.RowCount = 0) then
      exit;
    for x := 0 to SGRubr.RowCount - 1 do
    begin
      idRUBR := SGRubr.Cells[1, x];
      finalSTR := finalSTR + '#' + idRUBR + '=';
      for z := 0 to SGNapravlenie.RowCount - 1 do
      begin
        idNAPR := SGNapravlenie.Cells[1, z];
        finalSTR := finalSTR + '(' + idNAPR + ')';
      end;
      finalSTR := finalSTR + '$';
    end;
    Result := finalSTR;
    // result = #13=(534)(654)(823)$#27=(43)(534)(546)$
  end;

begin
  if Trim(EditName.Text) = '' then
  begin
    MessageBox(handle, 'Необходимо указать название.', 'Информация', MB_OK or MB_ICONWARNING);
    exit;
  end;
  if SGRubr.RowCount = 0 then
  begin
    MessageBox(handle, 'Необходимо указать рубрику.', 'Информация', MB_OK or MB_ICONWARNING);
    exit;
  end;

  if IsNewRecordFound_Notify(listNewRecords) = true then
  begin
    if MessageBox(handle, PChar('Новые директории будут созданы. Продолжить?' + #13#10 + #13#10 + listNewRecords.Text), 'Подтверждение',
      MB_YESNO or MB_ICONQUESTION) = MRNO then
    begin
      listNewRecords.Free;
      exit;
    end;
  end;
  listNewRecords.Free;

  IsNewRecordCheck;

  EditName.Text := UpperFirst(EditName.Text);
  EditFIO.Text := UpperFirst(EditFIO.Text);
  FormMain.IBQuery1.Close;
  NewRubr := GetIDString(SGRubr);
  FormMain.IBQuery1.SQL.Text := 'update BASE set ACTIVITY=:ACTIVITY,RELEVANCE=:RELEVANCE,NAME=:NAME,FIO=:FIO,CURATOR=:CURATOR,' +
    'RUBR=:RUBR,FIRMTYPE=:FIRMTYPE,NAPRAVLENIE=:NAPRAVLENIE,WEB=:WEB,EMAIL=:EMAIL,ADRES=:ADRES,PHONES=:PHONES,DATE_EDITED=:DATE_EDITED,RELATIONS=:RELATIONS where ID=:ID';
  FormMain.IBQuery1.ParamByName('ACTIVITY').AsBoolean := CBActivity.Checked;
  FormMain.IBQuery1.ParamByName('RELEVANCE').AsBoolean := cbDataRelevance.Checked;
  FormMain.IBQuery1.ParamByName('NAME').AsString := EditName.Text;
  FormMain.IBQuery1.ParamByName('FIO').AsString := EditFIO.Text;
  FormMain.IBQuery1.ParamByName('CURATOR').AsString := GetIDString(SGCurator);
  FormMain.IBQuery1.ParamByName('RUBR').AsString := NewRubr;
  FormMain.IBQuery1.ParamByName('FIRMTYPE').AsString := GetIDString(SGFirmType);
  FormMain.IBQuery1.ParamByName('NAPRAVLENIE').AsString := GetIDString(SGNapravlenie);
  FormMain.IBQuery1.ParamByName('WEB').AsString := GetWebEmailString(SGWeb);
  FormMain.IBQuery1.ParamByName('EMAIL').AsString := GetWebEmailString(SGEMail);

  str := TStringList.Create;
  phones := TStringList.Create;
  try
    { #TODO2: REVISIT : AdresProcs code is bad. See how can I rework this to be less bad }
    IsAdresFilled(CBAdres1, EditOfficeType1, EditZIP1, EditStreet1, EditCountry1, EditRegion1, EditCity1, SGPhone1);
    IsAdresFilled(CBAdres2, EditOfficeType2, EditZIP2, EditStreet2, EditCountry2, EditRegion2, EditCity2, SGPhone2);
    IsAdresFilled(CBAdres3, EditOfficeType3, EditZIP3, EditStreet3, EditCountry3, EditRegion3, EditCity3, SGPhone3);
    IsAdresFilled(CBAdres4, EditOfficeType4, EditZIP4, EditStreet4, EditCountry4, EditRegion4, EditCity4, SGPhone4);
    IsAdresFilled(CBAdres5, EditOfficeType5, EditZIP5, EditStreet5, EditCountry5, EditRegion5, EditCity5, SGPhone5);
    IsAdresFilled(CBAdres6, EditOfficeType6, EditZIP6, EditStreet6, EditCountry6, EditRegion6, EditCity6, SGPhone6);
    IsAdresFilled(CBAdres7, EditOfficeType7, EditZIP7, EditStreet7, EditCountry7, EditRegion7, EditCity7, SGPhone7);
    IsAdresFilled(CBAdres8, EditOfficeType8, EditZIP8, EditStreet8, EditCountry8, EditRegion8, EditCity8, SGPhone8);
    IsAdresFilled(CBAdres9, EditOfficeType9, EditZIP9, EditStreet9, EditCountry9, EditRegion9, EditCity9, SGPhone9);
    IsAdresFilled(CBAdres10, EditOfficeType10, EditZIP10, EditStreet10, EditCountry10, EditRegion10, EditCity10, SGPhone10);
    AdresProcs(CBAdres1, EditNO1, EditOfficeType1, EditZIP1, EditStreet1, EditCountry1, EditRegion1, EditCity1, MemoPhone1, SGPhone1);
    AdresProcs(CBAdres2, EditNO2, EditOfficeType2, EditZIP2, EditStreet2, EditCountry2, EditRegion2, EditCity2, MemoPhone2, SGPhone2);
    AdresProcs(CBAdres3, EditNO3, EditOfficeType3, EditZIP3, EditStreet3, EditCountry3, EditRegion3, EditCity3, MemoPhone3, SGPhone3);
    AdresProcs(CBAdres4, EditNO4, EditOfficeType4, EditZIP4, EditStreet4, EditCountry4, EditRegion4, EditCity4, MemoPhone4, SGPhone4);
    AdresProcs(CBAdres5, EditNO5, EditOfficeType5, EditZIP5, EditStreet5, EditCountry5, EditRegion5, EditCity5, MemoPhone5, SGPhone5);
    AdresProcs(CBAdres6, EditNO6, EditOfficeType6, EditZIP6, EditStreet6, EditCountry6, EditRegion6, EditCity6, MemoPhone6, SGPhone6);
    AdresProcs(CBAdres7, EditNO7, EditOfficeType7, EditZIP7, EditStreet7, EditCountry7, EditRegion7, EditCity7, MemoPhone7, SGPhone7);
    AdresProcs(CBAdres8, EditNO8, EditOfficeType8, EditZIP8, EditStreet8, EditCountry8, EditRegion8, EditCity8, MemoPhone8, SGPhone8);
    AdresProcs(CBAdres9, EditNO9, EditOfficeType9, EditZIP9, EditStreet9, EditCountry9, EditRegion9, EditCity9, MemoPhone9, SGPhone9);
    AdresProcs(CBAdres10, EditNO10, EditOfficeType10, EditZIP10, EditStreet10, EditCountry10, EditRegion10, EditCity10, MemoPhone10,
      SGPhone10);

    FieldAdresText := str.Text;
    FormMain.IBQuery1.ParamByName('ADRES').AsString := FieldAdresText;
    FormMain.IBQuery1.ParamByName('PHONES').AsString := phones.Text;
    FormMain.IBQuery1.ParamByName('DATE_EDITED').AsString := DateToStr(Now);
    FormMain.IBQuery1.ParamByName('RELATIONS').AsString := BuildRelations;
    FormMain.IBQuery1.ParamByName('ID').AsString := lblID.Caption;
  finally
    str.Free;
    phones.Free;
  end;

  try
    FormMain.IBQuery1.Execute;

    DoGarbageCollection(lblID.Caption, 'CURATOR', 'CURATOR', list_GC_IDs.Values['CURATOR'], GetIDString(sgCurator), Main.sgCurator_tmp,
      FormDirectory.SGCurator, EditCurator, 'edit');
    DoGarbageCollection(lblID.Caption, 'FIRMTYPE', 'FIRMTYPE', list_GC_IDs.Values['FIRMTYPE'], GetIDString(SGFirmType), Main.sgFirmType_tmp,
      FormDirectory.SGFirmType, EditFirmType, 'edit');
    DoGarbageCollection(lblID.Caption, 'NAPRAVLENIE', 'NAPRAVLENIE', list_GC_IDs.Values['NAPRAVLENIE'], GetIDString(SGNapravlenie),
      Main.sgNapr_tmp, FormDirectory.SGNapr, EditNapravlenie, 'edit');

    DoNaprActivityValidation(lblID.Caption, list_GC_IDs.Values['NAPRAVLENIE'], GetIDString(SGNapravlenie), CBActivity.Checked, 'edit');

    DoCityActivityValidation(lblID.Caption, list_GC_IDs.Values['ADRES'], FieldAdresText, CBActivity.Checked, 'edit');

    WriteLog('TFormEditor.EditRecord: запись отредактирована ' + lblID.Caption);
  except
    on E: Exception do
    begin
      WriteLog('TFormEditor.EditRecord' + #13 + 'Ошибка: ' + E.Message);
      MessageBox(handle, PChar('Ошибка при редактировании фирмы.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      exit;
    end;
  end;
  FormMain.IBTransaction1.CommitRetaining;

  // если вкладка была открыта, заркрываем старую, открываем новую
  if FormMain.CloseTabByID(lblID.Caption) then
    FormMain.OpenTabByID(lblID.Caption);

  // удалили все старые ( до редактирования ) TVRubrikator
  FormMain.TVRubrikator.Items.BeginUpdate;
  for i := FormMain.TVRubrikator.Items.Count - 1 downto 0 do
    if (IntToStr(Integer(Pointer(FormMain.TVRubrikator.Items[i].Data))) = lblID.Caption) and (FormMain.TVRubrikator.Items[i].Level <> 0)
    then
      FormMain.TVRubrikator.Items[i].delete;
  rubrID := NewRubr;
  while pos('$', rubrID) > 0 do // создали новые ( после редактирования )
  begin
    tmp := copy(rubrID, 0, pos('$', rubrID));
    delete(rubrID, 1, length(tmp));
    delete(tmp, 1, 1);
    delete(tmp, length(tmp), 1);
    nd := FormMain.SearchNode(FormMain.TVRubrikator, StrToInt(tmp), 0);
    if nd <> nil then
    begin
      nd2 := FormMain.TVRubrikator.Items.AddChildObject(nd, EditName.Text, Pointer(StrToInt(lblID.Caption)));
      if CBActivity.Checked then
      begin
        nd2.ImageIndex := 1;
        nd2.SelectedIndex := 1;
      end
      else
      begin
        nd2.ImageIndex := 2;
        nd2.SelectedIndex := 2;
      end;
    end;
  end;

  if FormMain.TVRubrikator.Selected <> nil then
  begin { обновляем SGGeneral если показана рубрика в которой изменился итем }
    rubrID := IntToStr(Integer(Pointer(FormMain.TVRubrikator.Selected.Data)));
    if pos('#' + rubrID + '$', NewRubr) > 0 then
      FormMain.TVRubrikatorChange(FormMain.TVRubrikator, FormMain.TVRubrikator.Selected);
  end;

  FormMain.TVRubrikator.CustomSort(@Helpers.CustomSortProc, 0, True);
  FormMain.TVRubrikator.Items.EndUpdate;
  FormMain.IBQuery1.Close;
  FormEditor.Close;

  FormRelations.ClearEdits;
  FormRelations.PrepareRecord(lblID.Caption);
  FormRelations.Show;
  // FormMain.IBDatabase1.Connected := False;
end;

{ #TODO1: DELO : Implement same OnExit event in DELO }
procedure TFormEditor.EditRegionOrCityOnExit(Sender: TObject);
var
  Key: Char;
begin
  { The logic in this interaction is quite messy. Point of this is to allow selecting value OnExit without having to press Enter
    Basically this event will either set new value if it valid or decide what to do with invalid value.
    OnExit is also called from OnSelect event when edit trying to change it ItemIndex to -1 }
  if TsComboBoxEx(Sender).GetID <> -1 then
  begin
    // valid value entered, call for OnSelect...
    Key := #13;
    EditRegionOrCityOnKeyPress(Sender, Key);
  end
  else
  begin
    // invalid value entered, check for what to do...
    if TsComboBoxEx(Sender).Tag = -1 then // nothing was set in this edit -> clear invalid value
      TsComboBoxEx(Sender).Text := EmptyStr
    else if Trim(TsComboBoxEx(Sender).Text) = EmptyStr then // edit had value set but user cleared text -> set selection to none
      TsComboBoxEx(Sender).Tag := -1
    else if (TsComboBoxEx(Sender).ItemIndex = -1) and (TsComboBoxEx(Sender).Tag <> -1) then
      // entered value is not valid, reset back to previous valid value
      TsComboBoxEx(Sender).ItemIndex := TsComboBoxEx(Sender).Tag;
  end;
end;

procedure TFormEditor.EditRegionOrCityOnKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  begin
    if ContainsText(TsComboBoxEx(Sender).Name, 'EditRegion') then
      EditRegionSelect(Sender)
    else if ContainsText(TsComboBoxEx(Sender).Name, 'EditCity') then
      EditCitySelect(Sender)
  end;
end;

procedure TFormEditor.EditRegionSelect(Sender: TObject);
var
  EditRegion, EditCity: TsComboBoxEx;
  ID_AdresPage: integer;
begin
  EditRegion := TsComboBoxEx(Sender);
  if EditRegion.Tag <> EditRegion.ItemIndex then
  begin
    if EditRegion.ItemIndex <> -1 then
    begin
      EditRegion.Tag := EditRegion.ItemIndex;
      ID_AdresPage := StrToInt(TsTabSheet(EditRegion.Parent.Parent).Caption);
      EditCity := TsComboBoxEx(FormEditor.FindComponent('EditCity' + IntToStr(ID_AdresPage)));
      if Assigned(EditCity) then
        ClearEdit(EditCity);
    end
    else
      EditRegionOrCityOnExit(Sender);
  end;
end;

procedure TFormEditor.EditCitySelect(Sender: TObject);
var
  EditRegion, EditCity: TsComboBoxEx;
  ID_City, ID_Region, ID_AdresPage: integer;
  Query: TIBCQuery;
begin
  EditCity := TsComboBoxEx(Sender);
  ID_AdresPage := StrToInt(TsTabSheet(EditCity.Parent.Parent).Caption);
  EditRegion := TsComboBoxEx(FormEditor.FindComponent('EditRegion' + IntToStr(ID_AdresPage)));

  if EditCity.Tag <> EditCity.ItemIndex then
  begin
    if EditCity.ItemIndex <> -1 then
    begin
      EditCity.Tag := EditCity.ItemIndex;
      Query := QueryCreate;
      try
        ClearEdit(EditRegion);
        ID_City := EditCity.GetID;
        Query.SQL.Text := 'select ID_REGION from CITY where ID = :ID';
        Query.ParamByName('ID').AsInteger := ID_City;
        Query.Open;
        if Query.RecordCount > 0 then
        begin
          ID_Region := VarToStrDef(Query.FieldValues['ID_REGION'], '-1').ToInteger;
          EditRegion.SetIndexOfObject(ID_Region);
        end
      finally
        Query.Free;
      end;
    end
    else
      EditRegionOrCityOnExit(Sender);
  end;
end;

procedure TFormEditor.DeleteRecord(id: string);
var
  i: Integer;
  Q: TIBCQuery;
  capt: string;
begin
  Q := QueryCreate;
  Q.SQL.Text := 'select * from BASE where ID = :ID';
  Q.ParamByName('ID').AsString := id;
  Q.Open;
  capt := Q.FieldByName('NAME').AsString;
  if MessageBox(handle, 'Запись будет удалена. Продолжить?', PChar(capt), MB_YESNO or MB_ICONQUESTION) = MRNO then
  begin
    Q.Close;
    Q.Free;
    exit;
  end;

  // Build list_GC_IDs that is used for DoGarbageCollection
  if not Assigned(list_GC_IDs) then
    list_GC_IDs := TStringList.Create;
  list_GC_IDs.Clear;
  if Q.FieldValues['CURATOR'] <> null then
    list_GC_IDs.Add('CURATOR=' + Q.FieldValues['CURATOR']);
  if Q.FieldValues['RUBR'] <> null then
    list_GC_IDs.Add('RUBRIKATOR=' + Q.FieldValues['RUBR']);
  if Q.FieldValues['FIRMTYPE'] <> null then
    list_GC_IDs.Add('FIRMTYPE=' + Q.FieldValues['FIRMTYPE']);
  if Q.FieldValues['NAPRAVLENIE'] <> null then
    list_GC_IDs.Add('NAPRAVLENIE=' + Q.FieldValues['NAPRAVLENIE']);
  if Q.FieldValues['ADRES'] <> null then
    list_GC_IDs.Add('ADRES=' + Q.FieldValues['ADRES']);

  Q.Close;
  Q.SQL.Text := 'delete from BASE where ID = :ID';
  Q.ParamByName('ID').AsString := id;
  try
    Q.Execute;

    // DO Garbage Collection
    DoGarbageCollection(id, 'CURATOR', 'CURATOR', list_GC_IDs.Values['CURATOR'], EmptyStr, Main.sgCurator_tmp, FormDirectory.SGCurator,
      EditCurator, 'delete');
    DoGarbageCollection(id, 'FIRMTYPE', 'FIRMTYPE', list_GC_IDs.Values['FIRMTYPE'], EmptyStr, Main.sgFirmType_tmp, FormDirectory.SGFirmType,
      EditFirmType, 'delete');
    DoGarbageCollection(id, 'NAPRAVLENIE', 'NAPRAVLENIE', list_GC_IDs.Values['NAPRAVLENIE'], EmptyStr, Main.sgNapr_tmp,
      FormDirectory.SGNapr, EditNapravlenie, 'delete');

    DoNaprActivityValidation(id, list_GC_IDs.Values['NAPRAVLENIE'], EmptyStr, False, 'delete');

    DoCityActivityValidation(id, list_GC_IDs.Values['ADRES'], EmptyStr, False, 'delete');

    WriteLog('TFormEditor.DeleteRecord: запись удалена ' + id);
  except
    on E: Exception do
    begin
      WriteLog('TFormEditor.DeleteRecord' + #13 + 'Ошибка: ' + E.Message);
      MessageBox(handle, PChar('Ошибка при удалении фирмы ' + id + '.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      exit;
    end;
  end;
  FormMain.IBTransaction1.CommitRetaining;
  FormMain.sStatusBar1.Panels[1].Text := 'Фирм в базе: ' + GetFirmCount;
  Q.Close;
  Q.Free;
  // FormMain.IBDatabase1.Close;
  for i := FormMain.TVRubrikator.Items.Count - 1 downto 0 do
    if (IntToStr(Integer(Pointer(FormMain.TVRubrikator.Items[i].Data))) = id) and (FormMain.TVRubrikator.Items[i].Level <> 0) then
      FormMain.TVRubrikator.Items[i].delete;
  for i := FormMain.SGGeneral.RowCount - 1 downto 0 do
    if FormMain.SGGeneral.Cells[0, i] = id then
    begin
      FormMain.SGGeneral.DeleteRow(i);
      break;
    end;
  FormMain.CloseTabByID(id);
end;

function TFormEditor.IsNewRecordFound_Notify(var listNewRecords: TStringList): boolean;
var
  i, x: integer;
  edit: TsComboBoxEx;
begin
  result := false;
  if not Assigned(listNewRecords) then
    listNewRecords := TStringList.Create;

  for i := 0 to SGCurator.RowCount - 1 do
    if SGCurator.Cells[1, i] = '-1' then
    begin
      result := true;
      listNewRecords.Add('Куратор: ' + SGCurator.Cells[0, i]);
    end;
  for i := 0 to SGRubr.RowCount - 1 do
    if SGRubr.Cells[1, i] = '-1' then
    begin
      result := true;
      listNewRecords.Add('Рубрика: ' + SGRubr.Cells[0, i]);
    end;
  for i := 0 to SGFirmType.RowCount - 1 do
    if SGFirmType.Cells[1, i] = '-1' then
    begin
      result := true;
      listNewRecords.Add('Тип фирмы: ' + SGFirmType.Cells[0, i]);
    end;
  { for i := 0 to SGNapravlenie.RowCount - 1 do
    if SGNapravlenie.Cells[1, i] = '-1' then
    begin
    result := true;
    listNewRecords.Add('Профиль: ' + SGNapravlenie.Cells[0, i]);
    end; }
  for x := 1 to 10 do
  begin
    edit := TsComboBoxEx(FindComponent('EditOfficeType' + IntToStr(x)));
    if (edit.GetID = -1) and (Trim(edit.Text) <> '') then
    begin
      result := true;
      if listNewRecords.IndexOf('Тип адреса: ' + Trim(edit.Text)) = -1 then
        listNewRecords.Add('Тип адреса: ' + Trim(edit.Text));
    end;
    edit := TsComboBoxEx(FindComponent('EditCountry' + IntToStr(x)));
    if (edit.GetID = -1) and (Trim(edit.Text) <> '') then
    begin
      result := true;
      if listNewRecords.IndexOf('Страна: ' + Trim(edit.Text)) = -1 then
        listNewRecords.Add('Страна: ' + Trim(edit.Text));
    end;
    // #TODO2: NEW_RECORD_NOTIFY left overs
    { edit := TsComboBoxEx(FindComponent('EditRegion' + IntToStr(x)));
      if (edit.Items.IndexOf(Trim(edit.Text)) = -1) and (Trim(edit.Text) <> '') then
      begin
      result := true;
      listNewRecords.Add('Область: ' + Trim(edit.Text));
      end;
      edit := TsComboBoxEx(FindComponent('EditCity' + IntToStr(x)));
      if (edit.Items.IndexOf(Trim(edit.Text)) = -1) and (Trim(edit.Text) <> '') then
      begin
      result := true;
      listNewRecords.Add('Город: ' + Trim(edit.Text));
      end; }
  end;
end;

{ #TODO1: REVISIT : Idea is to instead of automatically create new directory, which is what this produre currently does,
  I want to show confirmation dialog on Enter/add button press asking user if he want to create new directory. }
procedure TFormEditor.IsNewRecordCheck;
var
  i, x: Integer;
  isNewRubr: Boolean;
  edit: TsComboBoxEx;

  procedure AddingNewRec(table, Value: string; sg_Main, sg_Directory: TNextGrid);
  var
    Q: TIBCQuery;
    id: string;
    node: TTreeNode;
    z: Integer;
  begin
    Q := QueryCreate;
    Q.Close;
    Q.SQL.Text := 'insert into ' + table + ' (NAME) values (:NAME) returning ID';
    Q.ParamByName('NAME').AsString := Value;
    try
      Q.Execute;
      Q.Transaction.CommitRetaining;
    except
      on E: Exception do
      begin
        Q.Transaction.RollbackRetaining;
        WriteLog('TFormEditor.IsNewRecordCheck' + #13 + 'Ошибка: ' + E.Message);
        MessageBox(handle, PChar('Ошибка при проверке существующих директорий.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
        Q.Free;
        exit;
      end;
    end;
    id := Q.ParamByName('RET_ID').AsString;
    Q.Free;

    if AnsiLowerCase(table) = 'curator' then
    begin
      SGCurator.Cells[1, i] := id;
      EditCurator.AddItem(Value, Pointer(StrToInt(id)));
    end;
    if AnsiLowerCase(table) = 'rubrikator' then
    begin
      node := FormMain.TVRubrikator.Items.AddChildObject(nil, Value, Pointer(StrToInt(id)));
      node.ImageIndex := 0;
      node.SelectedIndex := 0;
      SGRubr.Cells[1, i] := id; { i = счетчик из основной процедуры }
      EditRubr.AddItem(Value, Pointer(StrToInt(id)));
    end;
    if AnsiLowerCase(table) = 'firmtype' then
    begin
      SGFirmType.Cells[1, i] := id;
      EditFirmType.AddItem(Value, Pointer(StrToInt(id)));
    end;
    if AnsiLowerCase(table) = 'napravlenie' then
    begin
      SGNapravlenie.Cells[1, i] := id;
      EditNapravlenie.AddItem(Value, Pointer(StrToInt(id)));
    end;
    if AnsiLowerCase(table) = 'officetype' then
      for z := 1 to 10 do
        TsComboBoxEx(FindComponent('EditOfficeType' + IntToStr(z))).AddItem(Value, Pointer(StrToInt(id)));
    if AnsiLowerCase(table) = 'country' then
      for z := 1 to 10 do
        TsComboBoxEx(FindComponent('EditCountry' + IntToStr(z))).AddItem(Value, Pointer(StrToInt(id)));
    // #TODO2: NEW_RECORD_CHECK left overs
    { if AnsiLowerCase(table) = 'region' then
      for z := 1 to 10 do
      TsComboBoxEx(FindComponent('EditRegion' + IntToStr(z))).AddItem(Value, Pointer(StrToInt(id)));
      if AnsiLowerCase(table) = 'city' then
      for z := 1 to 10 do
      TsComboBoxEx(FindComponent('EditCity' + IntToStr(z))).AddItem(Value, Pointer(StrToInt(id))); }
    if sg_Main <> nil then
    begin
      sg_Main.AddRow;
      sg_Main.Cells[0, sg_Main.LastAddedRow] := Value;
      sg_Main.Cells[1, sg_Main.LastAddedRow] := id;
    end;
    if sg_Directory <> nil then
    begin
      sg_Directory.AddRow;
      sg_Directory.Cells[0, sg_Directory.LastAddedRow] := Value;
      sg_Directory.Cells[1, sg_Directory.LastAddedRow] := id;
    end;
  end;

begin
  isNewRubr := False;
  for i := 0 to SGCurator.RowCount - 1 do
    if SGCurator.Cells[1, i] = '-1' then
    begin
      AddingNewRec('CURATOR', SGCurator.Cells[0, i], Main.sgCurator_tmp, FormDirectory.SGCurator);
    end;
  for i := 0 to SGRubr.RowCount - 1 do
    if SGRubr.Cells[1, i] = '-1' then
    begin
      AddingNewRec('RUBRIKATOR', SGRubr.Cells[0, i], Main.sgRubr_tmp, FormDirectory.SGRubr);
      isNewRubr := True;
    end;
  for i := 0 to SGFirmType.RowCount - 1 do
    if SGFirmType.Cells[1, i] = '-1' then
    begin
      AddingNewRec('FIRMTYPE', SGFirmType.Cells[0, i], Main.sgFirmType_tmp, FormDirectory.SGFirmType);
    end;
  for i := 0 to SGNapravlenie.RowCount - 1 do
    if SGNapravlenie.Cells[1, i] = '-1' then
    begin
      AddingNewRec('NAPRAVLENIE', SGNapravlenie.Cells[0, i], Main.sgNapr_tmp, FormDirectory.SGNapr);
    end;
  for x := 1 to 10 do
  begin
    edit := TsComboBoxEx(FindComponent('EditOfficeType' + IntToStr(x)));
    if (edit.GetID = -1) and (Trim(edit.Text) <> '') then
      AddingNewRec('OFFICETYPE', UpperFirst(edit.Text), nil, FormDirectory.SGOfficeType);

    edit := TsComboBoxEx(FindComponent('EditCountry' + IntToStr(x)));
    if (edit.GetID = -1) and (Trim(edit.Text) <> '') then
      AddingNewRec('COUNTRY', UpperFirst(edit.Text), nil, FormDirectory.SGCountry);
    // #TODO2: NEW_RECORD_CHECK left overs
    { edit := TsComboBoxEx(FindComponent('EditRegion' + IntToStr(x)));
      if (edit.Items.IndexOf(Trim(edit.Text)) = -1) and (Trim(edit.Text) <> '') then
      AddingNewRec('REGION', UpperFirst(edit.Text), nil, FormDirectory.SGRegion);
      edit := TsComboBoxEx(FindComponent('EditCity' + IntToStr(x)));
      if (edit.Items.IndexOf(Trim(edit.Text)) = -1) and (Trim(edit.Text) <> '') then
      AddingNewRec('CITY', UpperFirst(edit.Text), nil, FormDirectory.SGCity); }
  end;
  if isNewRubr then
    FormMain.TVRubrikator.CustomSort(@Helpers.CustomSortProc, 0, True);
end;

procedure TFormEditor.IsRecordDublicate;
var
  Name, WEBstr, WEBtmp, EMAILstr, EMAILtmp, PHONEstr: string;
  WEBlist, EMAILlist, PHONElist: TStrings;
  REQ, REQ1, REQ2, REQ3, REQ4: string;
  x: Integer;
  Q: TIBCQuery;

  procedure GetPhoneList(SGPhone: TNextGrid);
  var
    z: Integer;
  begin
    if SGPhone.RowCount > 0 then
    begin
      for z := 0 to SGPhone.RowCount - 1 do
      begin
        PHONEstr := SGPhone.Cells[0, z];
        delete(PHONEstr, 1, pos(')', PHONEstr) + 1);
        PHONEstr := AnsiLowerCase(Trim(PHONEstr));
        if PHONElist.IndexOf(PHONEstr) = -1 then
          PHONElist.Add(PHONEstr);
      end;
    end;
  end;

begin
  IsDublicate := True;
  name := AnsiLowerCase(Trim(EditName.Text));
  WEBstr := GetWebEmailString(SGWeb);
  WEBlist := TStringList.Create;
  if length(WEBstr) > 0 then
  begin
    if WEBstr[length(WEBstr)] <> ',' then
      WEBstr := WEBstr + ',';
    while pos(',', WEBstr) > 0 do
    begin
      WEBtmp := copy(WEBstr, 0, pos(',', WEBstr));
      delete(WEBstr, 1, length(WEBtmp));
      delete(WEBtmp, length(WEBtmp), 1);
      WEBtmp := AnsiLowerCase(Trim(WEBtmp));
      WEBlist.Add(WEBtmp);
    end;
  end;
  EMAILstr := GetWebEmailString(SGEMail);
  EMAILlist := TStringList.Create;
  if length(EMAILstr) > 0 then
  begin
    if EMAILstr[length(EMAILstr)] <> ',' then
      EMAILstr := EMAILstr + ',';
    while pos(',', EMAILstr) > 0 do
    begin
      EMAILtmp := copy(EMAILstr, 0, pos(',', EMAILstr));
      delete(EMAILstr, 1, length(EMAILtmp));
      delete(EMAILtmp, length(EMAILtmp), 1);
      EMAILtmp := AnsiLowerCase(Trim(EMAILtmp));
      EMAILlist.Add(EMAILtmp);
    end;
  end;

  PHONElist := TStringList.Create;
  GetPhoneList(SGPhone1);
  GetPhoneList(SGPhone2);
  GetPhoneList(SGPhone3);
  GetPhoneList(SGPhone4);
  GetPhoneList(SGPhone5);
  GetPhoneList(SGPhone6);
  GetPhoneList(SGPhone7);
  GetPhoneList(SGPhone8);
  GetPhoneList(SGPhone9);
  GetPhoneList(SGPhone10);

  { #TODO2: REVISIT : IsDublicate logic is kinda bad. Probably need to get rid of phones comparing and maybe other stuff }

  // REQ1 := ' lower(NAME) = ''' + Name + ''' or';
  REQ1 := ' lower(NAME) = :NAME or';
  if WEBlist.Count > 0 then
    for x := 0 to WEBlist.Count - 1 do
      REQ2 := REQ2 + ' lower(WEB) like ''%' + WEBlist[x] + '%'' or';
  if EMAILlist.Count > 0 then
    for x := 0 to EMAILlist.Count - 1 do
      REQ3 := REQ3 + ' lower(EMAIL) like ''%' + EMAILlist[x] + '%'' or';
  if PHONElist.Count > 0 then
    for x := 0 to PHONElist.Count - 1 do
      REQ4 := REQ4 + ' lower(PHONES) like ''%' + PHONElist[x] + '%'' or';
  REQ := 'select * from BASE where' + REQ1 + REQ2 + REQ3 + REQ4;
  delete(REQ, length(REQ) - 2, 3);
  WEBlist.Free;
  EMAILlist.Free;
  PHONElist.Free;

  Q := QueryCreate;
  Q.Close;
  Q.SQL.Text := REQ;
  Q.Params[0].AsString := name;
  try
    Q.Open;
  except
    on E: Exception do
    begin
      WriteLog('TFormEditor.IsRecordDublicate' + #13 + 'Ошибка: ' + E.Message);
      MessageBox(handle, PChar('Произошла ошибка при проверке фирм дубликатов' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      exit;
    end;
  end;
  Q.FetchAll := True;
  if Q.RecordCount > 0 then
  begin
    FormDublicate.SGDublicate.ClearRows;
    FormDublicate.SGDublicate.BeginUpdate;
    for x := 1 to Q.RecordCount do
    begin
      FormDublicate.SGDublicate.AddRow;
      if Q.FieldValues['ACTIVITY'] = '1' then
        FormDublicate.SGDublicate.Cells[0, FormDublicate.SGDublicate.LastAddedRow] := '1'
      else
        FormDublicate.SGDublicate.Cells[0, FormDublicate.SGDublicate.LastAddedRow] := '2';
      FormDublicate.SGDublicate.Cells[1, FormDublicate.SGDublicate.LastAddedRow] := Q.FieldValues['NAME'];
      FormDublicate.SGDublicate.Cells[2, FormDublicate.SGDublicate.LastAddedRow] := Q.FieldValues['ID'];
      Q.Next;
    end;
    FormDublicate.SGDublicate.EndUpdate;
    FormDublicate.BtnOK.OnClick := AddRecord;
    FormEditor.Close;
    FormDublicate.Show;
    // Нашли дубликаты, показали окно, ждем выбора пользователя
  end
  else
  begin
    AddRecord(BtnOK);
    // Дубликатов не найденно, добовляем фирму
  end;
  Q.Close;
  Q.Free;
  FormMain.IBDatabase1.Close;
end;

procedure TFormEditor.DoGarbageCollection(BASE_ID, DIR_Table, BASE_Field, IDString_OLD, IDString_NEW: string; SGTemp, SGDir: TNextGrid;
  EditOwner: TsComboBoxEx; Method: string);
var
  ID_OLD: string;
  i, index: integer;
  listID_OLD: TStringList;
  Query: TIBCQuery;

  procedure GC_RUN;
  begin
    Query.SQL.Text := 'select ID from BASE where ID <> :BASE_ID and ' + BASE_Field + ' like :ID_OLD rows 1';
    Query.ParamByName('BASE_ID').AsString := BASE_ID;
    Query.ParamByName('ID_OLD').AsString := '%#' + ID_OLD + '$%';
    Query.Open;
    if Query.RecordCount = 0 then
    begin
      Query.Close;
      Query.SQL.Text := 'delete from ' + DIR_Table + ' where ID = :ID';
      Query.ParamByName('ID').AsString := ID_OLD;

      try
        WriteLog('TFormEditor.DoGarbageCollection_SQL_LOG (Method: ' + Method + '): delete from ' + DIR_Table + ' where ID = ' + ID_OLD);
        debug('DELETED: DIR = %s | ID = %s', [DIR_Table, ID_OLD]);

        Query.Execute;
        Query.Transaction.Commit;

        { #TODO1: REFACTOR : Update this code to DirectoryContrainer class }
        if SGTemp.FindText(1, ID_OLD, [soCaseInsensitive, soExactMatch]) then
        begin
          debug('SGTEMP row found. deleted data: name = %s | ID = %s', [SGTemp.Cells[0, SGTemp.SelectedRow],
            SGTemp.Cells[1, SGTemp.SelectedRow]]);

          SGTemp.DeleteRow(SGTemp.SelectedRow);
        end
        else
          debug('WARNING! SGTEMP row NOT FOUND. expected data: ID = %s', [ID_OLD], 2);

        if SGDir.FindText(1, ID_OLD, [soCaseInsensitive, soExactMatch]) then
        begin
          debug('SGDIR row found. deleted data: name = %s | ID = %s', [SGDir.Cells[0, SGDir.SelectedRow],
            SGDir.Cells[1, SGDir.SelectedRow]]);

          SGDir.DeleteRow(SGDir.SelectedRow);
        end
        else
          debug('WARNING! SGDir row NOT FOUND. expected data: ID = %s', [ID_OLD], 2);

        index := EditOwner.GetIndexOfObject(ID_OLD.ToInteger);
        if index <> -1 then
        begin
          debug('EditOwner item found. deleted data: name = %s | ID = %s',
            [EditOwner.Items[index], IntToStr(Integer(EditOwner.Items.Objects[index]))]);

          EditOwner.Items.Delete(index);
        end
        else
          debug('WARNING! EditOwner item NOT FOUND. expected data: ID = %s', [ID_OLD], 2);

      except
        on E: Exception do
        begin
          Query.Transaction.Rollback;
          WriteLog('TFormEditor.DoGarbageCollection' + #13 + 'Ошибка: ' + E.Message);
          MessageBox(handle, PChar('Ошибка при очистке директорий.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
        end;
      end;
    end
    else
    begin
      debug('IN_USE: DIR = %s | ID = %s', [DIR_Table, ID_OLD]);
    end;
  end;

begin
  Query := QueryCreate;
  listID_OLD := ParseIDString(IDString_OLD);
  debug('*** GarbageCollection | METHOD = %s | DIR = %s ... ***', [Method, DIR_Table], 1);

  try
    for i := 0 to listID_OLD.Count - 1 do
    begin

      ID_OLD := listID_OLD[i];

      if Method = 'edit' then
      begin
        if not AnsiContainsStr(IDString_NEW, '#' + ID_OLD + '$') then
          GC_RUN
        else
          debug('SKIPPED: DIR = %s | ID = %s', [DIR_Table, ID_OLD]);
      end
      else if Method = 'delete' then
        GC_RUN;

    end;
  finally
    listID_OLD.Free;
    Query.Free;
  end;
end;

procedure TFormEditor.DoNaprActivityValidation(BASE_ID, IDString_OLD, IDString_NEW: string; IsActive: boolean; Method: string);
var
  Query: TIBCQuery;
  i: integer;
  listID_OLD, listID_NEW, listID_ALL, listID_DEL: TStringList;
  IDString_ALL, IDString_DEL, ID_OLD, ID_NEW, ID_ALL, ID_DEL: string;

  function IsNaprShouldBeActive(BASE_ID, NAPR_ID: string): boolean;
  begin
    if BASE_ID = EmptyStr then
      BASE_ID := '-1';
    Query.Close;
    Query.SQL.Text := 'select COUNT(*) from BASE where (ACTIVITY = 1 and ID <> :BASE_ID and NAPRAVLENIE like :NAPR_ID)';
    Query.ParamByName('BASE_ID').AsString := BASE_ID;
    Query.ParamByName('NAPR_ID').AsString := '%#' + NAPR_ID + '$%';
    Query.Open;
    result := Query.FieldValues['COUNT'] > 0;
  end;

  procedure UpdateNaprAcitivity(NAPR_ID: string; SetToActive: boolean = false);
  var
    QueryUpdate: TIBCQuery;
    nActive: integer;
  begin
    QueryUpdate := QueryCreate;

    if SetToActive = true then
      nActive := 1
    else
      nActive := 0;

    QueryUpdate.SQL.Text := 'update NAPRAVLENIE set ACTIVITY = :ACTIVITY where ID = :ID';
    QueryUpdate.ParamByName('ACTIVITY').AsInteger := nActive;
    QueryUpdate.ParamByName('ID').AsString := NAPR_ID;
    try
      QueryUpdate.Execute;
    except
      on E: Exception do
      begin
        WriteLog('TFTFormEditor.DoNaprActivityValidation' + #13 + 'Ошибка: ' + E.Message);
        MessageBox(handle, PChar('Ошибка при проверке активности видов деятельности.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      end;
    end;
    FormMain.IBTransaction1.CommitRetaining;
    QueryUpdate.Close;
    QueryUpdate.Free;
  end;

begin
  debug('*** DoNaprActivityValidation ... ***', [], 1);
  Query := QueryCreate;
  Method := AnsiLowerCase(Method);

  // build IDString_ALL with unique oldID's + newID's and IDString_DEL with deleted(edited) ID's
  IDString_ALL := IDString_NEW;
  listID_OLD := ParseIDString(IDString_OLD);
  for i := 0 to listID_OLD.Count - 1 do
  begin
    ID_OLD := listID_OLD[i];
    if not AnsiContainsStr(IDString_NEW, '#' + ID_OLD + '$') then
    begin
      IDString_DEL := IDString_DEL + '#' + ID_OLD + '$';
      IDString_ALL := IDString_ALL + '#' + ID_OLD + '$';
    end;
  end;

  listID_DEL := ParseIDString(IDString_DEL);
  listID_NEW := ParseIDString(IDString_NEW);
  listID_ALL := ParseIDString(IDString_ALL);

  // if Method = EDIT first we want to check new value of ACTIVITY for edited firm
  if Method = 'edit' then
  begin
    // if ACTIVE = TRUE then we have to do two seperate iteration
    // 1: go through list of DEL(deleted\edited) NAPR's and set their ACTIVITY values to FALSE except if current NAPR is in use by active firm
    // 2: go through list of NEW(made after edit) NAPR's and set their ACTIVITY values to TRUE
    if IsActive = true then
    begin
      debug('METHOD = %s | IsActive = %s | listID_DEL.COUNT = %s', [Method, 'TRUE', IntToStr(listID_DEL.Count)]);
      for i := 0 to listID_DEL.Count - 1 do
      begin
        ID_DEL := listID_DEL[i];
        UpdateNaprAcitivity(ID_DEL, IsNaprShouldBeActive(BASE_ID, ID_DEL));
        debug('listID_DEL item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_DEL, GetNameByID('NAPRAVLENIE', ID_DEL),
          BoolToStr(IsNaprShouldBeActive(BASE_ID, ID_DEL), True)]);
      end;
      debug('METHOD = %s | IsActive = %s | listID_NEW.COUNT = %s', [Method, 'TRUE', IntToStr(listID_NEW.Count)]);
      for i := 0 to listID_NEW.Count - 1 do
      begin
        ID_NEW := listID_NEW[i];
        UpdateNaprAcitivity(ID_NEW, true);
        debug('listID_NEW item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_NEW, GetNameByID('NAPRAVLENIE', ID_NEW),
          'TRUE']);
      end;
    end
    // if ACTIVE = FALSE then we have to do single iteration
    // 1: go through list of ALL(old+new) NAPR's and set their ACTIVITY values to FALSE except if current NAPR is in use by active firm
    else
    begin
      debug('METHOD = %s | IsActive = %s | listID_ALL.COUNT = %s', [Method, 'FALSE', IntToStr(listID_ALL.Count)]);
      for i := 0 to listID_ALL.Count - 1 do
      begin
        ID_ALL := listID_ALL[i];
        UpdateNaprAcitivity(ID_ALL, IsNaprShouldBeActive(BASE_ID, ID_ALL));
        debug('listID_ALL item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_ALL, GetNameByID('NAPRAVLENIE', ID_ALL),
          BoolToStr(IsNaprShouldBeActive(BASE_ID, ID_ALL), True)]);
      end;
    end;
  end;

  // if Method = ADD first we want to check the value of ACTIVITY for added firm
  if Method = 'add' then
  begin
    debug('METHOD = %s | IsActive = %s | listID_ALL.COUNT = %s', [Method, BoolToStr(IsActive, True), IntToStr(listID_ALL.Count)]);
    for i := 0 to listID_ALL.Count - 1 do
    begin
      ID_ALL := listID_ALL[i];
      // if firm ACITIVITY is TRUE then we set ACTIVITY of the current NAPR to TRUE
      if IsActive = true then
      begin
        UpdateNaprAcitivity(ID_ALL, true);
        debug('listID_ALL item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_ALL, GetNameByID('NAPRAVLENIE', ID_ALL),
          'TRUE']);
      end
      // if firm ACITIVITY is FALSE then we attempt to update ACTIVITY of the current NAPR to FALSE except if current NAPR is in use by active firm
      else
      begin
        UpdateNaprAcitivity(ID_ALL, IsNaprShouldBeActive(BASE_ID, ID_ALL));
        debug('listID_ALL item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_ALL, GetNameByID('NAPRAVLENIE', ID_ALL),
          BoolToStr(IsNaprShouldBeActive(BASE_ID, ID_ALL), True)]);
      end;
    end;
  end;

  // if Method = DELETE we attempt to update ACTIVITY of every NAPR to FALSE except if current NAPR is in use by active firm
  if Method = 'delete' then
  begin
    debug('METHOD = %s | IsActive = %s | listID_ALL.COUNT = %s', [Method, 'FALSE', IntToStr(listID_ALL.Count)]);
    for i := 0 to listID_ALL.Count - 1 do
    begin
      ID_ALL := listID_ALL[i];
      UpdateNaprAcitivity(ID_ALL, IsNaprShouldBeActive(BASE_ID, ID_ALL));
      debug('listID_ALL item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_ALL, GetNameByID('NAPRAVLENIE', ID_ALL),
        BoolToStr(IsNaprShouldBeActive(BASE_ID, ID_ALL), True)]);
    end;
  end;

  Query.Close;
  Query.Free;
  listID_OLD.Free;
  listID_NEW.Free;
  listID_ALL.Free;
  listID_DEL.Free;
end;

procedure TFormEditor.DoCityActivityValidation(BASE_ID, FieldAdres_OLD, FieldAdres_NEW: string; IsActive: boolean; Method: string);
var
  Query: TIBCQuery;
  i: integer;
  listID_OLD, listID_NEW, listID_ALL, listID_DEL: TStringList;
  ID_OLD, ID_NEW, ID_ALL, ID_DEL: string;

  function IsCityShouldBeActive(BASE_ID, CITY_ID: string): boolean;
  begin
    if BASE_ID = EmptyStr then
      BASE_ID := '-1';
    Query.Close;
    Query.SQL.Text := 'select COUNT(*) from BASE where (ACTIVITY = 1 and ID <> :BASE_ID and ADRES like :CITY_ID)';
    Query.ParamByName('BASE_ID').AsString := BASE_ID;
    Query.ParamByName('CITY_ID').AsString := '%#^' + CITY_ID + '$%';
    Query.Open;
    result := Query.FieldValues['COUNT'] > 0;
  end;

  procedure UpdateCityAcitivity(CITY_ID: string; SetToActive: boolean = false);
  var
    QueryUpdate: TIBCQuery;
    nActive: integer;
  begin
    if CITY_ID = EmptyStr then
      exit;

    QueryUpdate := QueryCreate;
    try
      if SetToActive = true then
        nActive := 1
      else
        nActive := 0;

      QueryUpdate.SQL.Text := 'update CITY set ACTIVITY = :ACTIVITY where ID = :ID';
      QueryUpdate.ParamByName('ACTIVITY').AsInteger := nActive;
      QueryUpdate.ParamByName('ID').AsString := CITY_ID;
      try
        QueryUpdate.Execute;
        QueryUpdate.Transaction.CommitRetaining;
      except
        on E: Exception do
        begin
          QueryUpdate.Transaction.RollbackRetaining;
          WriteLog('TFTFormEditor.DoCityActivityValidation' + #13 + 'Ошибка: ' + E.Message);
          MessageBox(handle, PChar('Ошибка при проверке активности городов.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
        end;
      end;
    finally
      QueryUpdate.Free;
    end;
  end;

begin
  debug('*** DoCityActivityValidation ... ***', [], 1);
  Query := QueryCreate;
  Method := AnsiLowerCase(Method);

  listID_ALL := TStringList.Create;
  listID_DEL := TStringList.Create;
  listID_OLD := ParseAdresFieldToCityIDList(FieldAdres_OLD);
  listID_NEW := ParseAdresFieldToCityIDList(FieldAdres_NEW);

  // build listID_ALL with unique oldID's + newID's and listID_DEL with deleted(edited) ID's
  listID_ALL.Text := listID_NEW.Text;
  for i := 0 to listID_OLD.Count - 1 do
  begin
    ID_OLD := listID_OLD[i];
    if listID_NEW.IndexOf(ID_OLD) = -1 then
    begin
      listID_DEL.Add(ID_OLD);
      listID_ALL.Add(ID_OLD);
    end;
  end;

  // if Method = EDIT first we want to check new value of ACTIVITY for edited firm
  if Method = 'edit' then
  begin
    // if ACTIVE = TRUE then we have to do two seperate iteration
    // 1: go through list of DEL(deleted\edited) CITY's and set their ACTIVITY values to FALSE except if current CITY is in use by active firm
    // 2: go through list of NEW(made after edit) CITY's and set their ACTIVITY values to TRUE
    if IsActive = true then
    begin
      debug('METHOD = %s | IsActive = %s | listID_DEL.COUNT = %s', [Method, 'TRUE', IntToStr(listID_DEL.Count)]);
      for i := 0 to listID_DEL.Count - 1 do
      begin
        ID_DEL := listID_DEL[i];
        UpdateCityAcitivity(ID_DEL, IsCityShouldBeActive(BASE_ID, ID_DEL));
        debug('listID_DEL item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_DEL, GetNameByID('CITY', ID_DEL),
          BoolToStr(IsCityShouldBeActive(BASE_ID, ID_DEL), True)]);
      end;
      debug('METHOD = %s | IsActive = %s | listID_NEW.COUNT = %s', [Method, 'TRUE', IntToStr(listID_NEW.Count)]);
      for i := 0 to listID_NEW.Count - 1 do
      begin
        ID_NEW := listID_NEW[i];
        UpdateCityAcitivity(ID_NEW, true);
        debug('listID_NEW item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_NEW, GetNameByID('CITY', ID_NEW), 'TRUE']);
      end;
    end
    // if ACTIVE = FALSE then we have to do single iteration
    // 1: go through list of ALL(old+new) CITY's and set their ACTIVITY values to FALSE except if current CITY is in use by active firm
    else
    begin
      debug('METHOD = %s | IsActive = %s | listID_ALL.COUNT = %s', [Method, 'FALSE', IntToStr(listID_ALL.Count)]);
      for i := 0 to listID_ALL.Count - 1 do
      begin
        ID_ALL := listID_ALL[i];
        UpdateCityAcitivity(ID_ALL, IsCityShouldBeActive(BASE_ID, ID_ALL));
        debug('listID_ALL item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_ALL, GetNameByID('CITY', ID_ALL),
          BoolToStr(IsCityShouldBeActive(BASE_ID, ID_ALL), True)]);
      end;
    end;
  end;

  // if Method = ADD first we want to check the value of ACTIVITY for added firm
  if Method = 'add' then
  begin
    debug('METHOD = %s | IsActive = %s | listID_ALL.COUNT = %s', [Method, BoolToStr(IsActive, True), IntToStr(listID_ALL.Count)]);
    for i := 0 to listID_ALL.Count - 1 do
    begin
      ID_ALL := listID_ALL[i];
      // if firm ACITIVITY is TRUE then we set ACTIVITY of the current CITY to TRUE
      if IsActive = true then
      begin
        UpdateCityAcitivity(ID_ALL, true);
        debug('listID_ALL item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_ALL, GetNameByID('CITY', ID_ALL), 'TRUE']);
      end
      // if firm ACITIVITY is FALSE then we attempt to update ACTIVITY of the current CITY to FALSE except if current CITY is in use by active firm
      else
      begin
        UpdateCityAcitivity(ID_ALL, IsCityShouldBeActive(BASE_ID, ID_ALL));
        debug('listID_ALL item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_ALL, GetNameByID('CITY', ID_ALL),
          BoolToStr(IsCityShouldBeActive(BASE_ID, ID_ALL), True)]);
      end;
    end;
  end;

  // if Method = DELETE we attempt to update ACTIVITY of every CITY to FALSE except if current CITY is in use by active firm
  if Method = 'delete' then
  begin
    debug('METHOD = %s | IsActive = %s | listID_ALL.COUNT = %s', [Method, 'FALSE', IntToStr(listID_ALL.Count)]);
    for i := 0 to listID_ALL.Count - 1 do
    begin
      ID_ALL := listID_ALL[i];
      UpdateCityAcitivity(ID_ALL, IsCityShouldBeActive(BASE_ID, ID_ALL));
      debug('listID_ALL item: = %s | ID = %s | NAME = %s | SetTo = %s', [IntToStr(i), ID_ALL, GetNameByID('CITY', ID_ALL),
        BoolToStr(IsCityShouldBeActive(BASE_ID, ID_ALL), True)]);
    end;
  end;

  Query.Close;
  Query.Free;
  listID_OLD.Free;
  listID_NEW.Free;
  listID_ALL.Free;
  listID_DEL.Free;
end;

end.
