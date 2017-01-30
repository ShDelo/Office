unit Editor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, sButton, ExtCtrls, sPanel, sComboBox, sEdit,
  sCheckBox, sLabel, sGroupBox, ComCtrls, sPageControl, acPNG, IBQuery,
  NxColumns, NxColumnClasses, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, sSpeedButton, Buttons, sMemo, StrUtils;

type
  TFormEditor = class(TForm)
    PanelAdv: TsPanel;
    PanelMain: TsPanel;
    EditName: TsEdit;
    EditFIO: TsEdit;
    EditCurator: TsComboBox;
    EditRubr: TsComboBox;
    EditType: TsComboBox;
    EditNapravlenie: TsComboBox;
    sPageControl1: TsPageControl;
    sTabSheet1: TsTabSheet;
    sGroupBox1: TsGroupBox;
    ImagePhone1: TImage;
    EditNO1: TsEdit;
    EditOfficeType1: TsComboBox;
    EditZIP1: TsEdit;
    EditStreet1: TsEdit;
    EditCountry1: TsComboBox;
    EditCity1: TsComboBox;
    sTabSheet2: TsTabSheet;
    sGroupBox2: TsGroupBox;
    EditNO2: TsEdit;
    EditOfficeType2: TsComboBox;
    EditZIP2: TsEdit;
    EditStreet2: TsEdit;
    EditCountry2: TsComboBox;
    EditCity2: TsComboBox;
    sTabSheet3: TsTabSheet;
    sGroupBox3: TsGroupBox;
    EditNO3: TsEdit;
    EditOfficeType3: TsComboBox;
    EditZIP3: TsEdit;
    EditStreet3: TsEdit;
    EditCountry3: TsComboBox;
    EditCity3: TsComboBox;
    sTabSheet4: TsTabSheet;
    sGroupBox4: TsGroupBox;
    EditNO4: TsEdit;
    EditOfficeType4: TsComboBox;
    EditZIP4: TsEdit;
    EditStreet4: TsEdit;
    EditCountry4: TsComboBox;
    EditCity4: TsComboBox;
    sTabSheet5: TsTabSheet;
    sGroupBox5: TsGroupBox;
    EditNO5: TsEdit;
    EditOfficeType5: TsComboBox;
    EditZIP5: TsEdit;
    EditStreet5: TsEdit;
    EditCountry5: TsComboBox;
    EditCity5: TsComboBox;
    sTabSheet6: TsTabSheet;
    sGroupBox6: TsGroupBox;
    EditNO6: TsEdit;
    EditOfficeType6: TsComboBox;
    EditZIP6: TsEdit;
    EditStreet6: TsEdit;
    EditCountry6: TsComboBox;
    EditCity6: TsComboBox;
    sTabSheet7: TsTabSheet;
    sGroupBox7: TsGroupBox;
    EditNO7: TsEdit;
    EditOfficeType7: TsComboBox;
    EditZIP7: TsEdit;
    EditStreet7: TsEdit;
    EditCountry7: TsComboBox;
    EditCity7: TsComboBox;
    sTabSheet8: TsTabSheet;
    sGroupBox8: TsGroupBox;
    EditNO8: TsEdit;
    EditOfficeType8: TsComboBox;
    EditZIP8: TsEdit;
    EditStreet8: TsEdit;
    EditCountry8: TsComboBox;
    EditCity8: TsComboBox;
    sTabSheet9: TsTabSheet;
    sGroupBox9: TsGroupBox;
    EditNO9: TsEdit;
    EditOfficeType9: TsComboBox;
    EditZIP9: TsEdit;
    EditStreet9: TsEdit;
    EditCountry9: TsComboBox;
    EditCity9: TsComboBox;
    sTabSheet10: TsTabSheet;
    sGroupBox10: TsGroupBox;
    EditNO10: TsEdit;
    EditOfficeType10: TsComboBox;
    EditZIP10: TsEdit;
    EditStreet10: TsEdit;
    EditCountry10: TsComboBox;
    EditCity10: TsComboBox;
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
    SGRubr: TNextGrid;
    NxTextColumn2: TNxTextColumn;
    BtnAddRubrToList: TsSpeedButton;
    BtnDeleteRubrFromList: TsSpeedButton;
    NxTextColumn1: TNxTextColumn;
    SGNapravlenie: TNextGrid;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
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
    SGCurator: TNextGrid;
    NxTextColumn5: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    SGType: TNextGrid;
    NxTextColumn7: TNxTextColumn;
    NxTextColumn8: TNxTextColumn;
    lblID: TsLabel;
    CBActivity: TsCheckBox;
    BtnOK: TsButton;
    BtnCancel: TsButton;
    BtnAddCuratorToList: TsSpeedButton;
    BtnDeleteCuratorFromList: TsSpeedButton;
    BtnAddTypeToList: TsSpeedButton;
    BtnDeleteTypeFromList: TsSpeedButton;
    EditPhoneType1: TsComboBox;
    SGPhone1: TNextGrid;
    NxTextColumn9: TNxTextColumn;
    MemoPhone1: TsMemo;
    EditPhone1: TsEdit;
    BtnAddPhoneToList1: TsSpeedButton;
    BtnDeletePhoneFromList1: TsSpeedButton;
    EditPhoneType2: TsComboBox;
    SGPhone2: TNextGrid;
    NxTextColumn10: TNxTextColumn;
    MemoPhone2: TsMemo;
    EditPhone2: TsEdit;
    BtnAddPhoneToList2: TsSpeedButton;
    BtnDeletePhoneFromList2: TsSpeedButton;
    EditPhoneType3: TsComboBox;
    SGPhone3: TNextGrid;
    NxTextColumn11: TNxTextColumn;
    MemoPhone3: TsMemo;
    EditPhone3: TsEdit;
    BtnAddPhoneToList3: TsSpeedButton;
    BtnDeletePhoneFromList3: TsSpeedButton;
    EditPhoneType4: TsComboBox;
    SGPhone4: TNextGrid;
    NxTextColumn12: TNxTextColumn;
    MemoPhone4: TsMemo;
    EditPhone4: TsEdit;
    BtnAddPhoneToList4: TsSpeedButton;
    BtnDeletePhoneFromList4: TsSpeedButton;
    EditPhoneType5: TsComboBox;
    SGPhone5: TNextGrid;
    NxTextColumn13: TNxTextColumn;
    MemoPhone5: TsMemo;
    EditPhone5: TsEdit;
    BtnAddPhoneToList5: TsSpeedButton;
    BtnDeletePhoneFromList5: TsSpeedButton;
    EditPhoneType6: TsComboBox;
    SGPhone6: TNextGrid;
    NxTextColumn14: TNxTextColumn;
    MemoPhone6: TsMemo;
    EditPhone6: TsEdit;
    BtnAddPhoneToList6: TsSpeedButton;
    BtnDeletePhoneFromList6: TsSpeedButton;
    EditPhoneType7: TsComboBox;
    SGPhone7: TNextGrid;
    NxTextColumn15: TNxTextColumn;
    MemoPhone7: TsMemo;
    EditPhone7: TsEdit;
    BtnAddPhoneToList7: TsSpeedButton;
    BtnDeletePhoneFromList7: TsSpeedButton;
    EditPhoneType8: TsComboBox;
    SGPhone8: TNextGrid;
    NxTextColumn16: TNxTextColumn;
    MemoPhone8: TsMemo;
    EditPhone8: TsEdit;
    BtnAddPhoneToList8: TsSpeedButton;
    BtnDeletePhoneFromList8: TsSpeedButton;
    EditPhoneType9: TsComboBox;
    SGPhone9: TNextGrid;
    NxTextColumn17: TNxTextColumn;
    MemoPhone9: TsMemo;
    EditPhone9: TsEdit;
    BtnAddPhoneToList9: TsSpeedButton;
    BtnDeletePhoneFromList9: TsSpeedButton;
    EditPhoneType10: TsComboBox;
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
    SGWeb: TNextGrid;
    NxTextColumn19: TNxTextColumn;
    SGEMail: TNextGrid;
    NxTextColumn21: TNxTextColumn;
    BtnAddWebToList: TsSpeedButton;
    BtnDeleteWebFromList: TsSpeedButton;
    BtnAddEMailToList: TsSpeedButton;
    BtnDeleteEMailFromList: TsSpeedButton;
    EditWEB: TsComboBox;
    EditEMAIL: TsComboBox;
    cbDataRelevance: TsCheckBox;
    procedure LoadDataEditor;
    procedure BtnCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    function UpperFirst(s: string): string;
    function GetIDByName(component: TsComboBox): string;
    function GetNameByID(table, id: string): string;
    function GetIDString(component: TNextGrid): string;
    function GetWebEmailString(component: TNextGrid): string;
    function GetFirmCount: string;
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
    procedure BtnAddPhoneToList1Click(Sender: TObject);
    procedure EditPhone1KeyPress(Sender: TObject; var Key: Char);
    procedure btnDeleteAdresClick(Sender: TObject);
    procedure BtnAddWebToListClick(Sender: TObject);
    procedure SGCuratorDblClick(Sender: TObject);
    procedure SGCuratorEnter(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  end;

var
  FormEditor: TFormEditor;
  IsDublicate: Boolean = False;

implementation

uses Main, Logo, Directory, Report, Dublicate, Relations, MailSend;

{$R *.dfm}

procedure TFormEditor.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.ExStyle := Params.ExStyle or WS_Ex_AppWindow;
end;

function CustomSortProc(Node1, Node2: TTreeNode; iUpToThisLevel: Integer): Integer; stdcall;
begin
  Result := AnsiStrIComp(PChar(Node1.Text), PChar(Node2.Text));
end;

function TFormEditor.UpperFirst(s: string): string;
var
  t: string;
begin
  Result := '';
  if length(Trim(s)) = 0 then
    exit;
  s := Trim(s);
  t := s[1];
  delete(s, 1, 1);
  t := AnsiUpperCase(t);
  Result := t + s;
end;

procedure TFormEditor.FormCreate(Sender: TObject);
begin
  LoadDataEditor;
end;

procedure TFormEditor.BtnCancelClick(Sender: TObject);
begin
  FormMain.WriteLog('TFormEditor.BtnCancelClick: отмена');
  FormEditor.Close;
end;

procedure TFormEditor.BtnAddCuratorToListClick(Sender: TObject);

  procedure Adding(sg: TNextGrid; edit: TsComboBox);
  var
    i: Integer;
  begin
    if Trim(edit.Text) = '' then
      exit;
    for i := 0 to sg.RowCount - 1 do
      if AnsiLowerCase(sg.Cells[0, i]) = AnsiLowerCase(Trim(edit.Text)) then
      begin
        MessageBox(handle, 'Запись с таким именем уже есть в списке.', 'Информация', MB_OK or MB_ICONINFORMATION);
        exit;
      end;
    sg.AddRow;
    sg.Cells[0, sg.LastAddedRow] := UpperFirst(edit.Text);
    sg.Cells[1, sg.LastAddedRow] := GetIDByName(edit);
    sg.Resort;
    edit.Text := '';
  end;

begin
  if TsSpeedButton(Sender).Name = 'BtnAddCuratorToList' then
    Adding(SGCurator, EditCurator);
  if TsSpeedButton(Sender).Name = 'BtnAddRubrToList' then
    Adding(SGRubr, EditRubr);
  if TsSpeedButton(Sender).Name = 'BtnAddTypeToList' then
    Adding(SGType, EditType);
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
  if TsSpeedButton(Sender).Name = 'BtnDeleteTypeFromList' then
    sg := SGType;
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
  edit: TsComboBox;
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
  // if TsComboBox(Sender).Name = 'EditNapravlenie' then
  // EditNapravlenie.DroppedDown := True; {Был баг: стерало всю строку едита при выпадании списка}
  if Key = #13 then
  begin
    Key := #0;
    if TsComboBox(Sender).Name = 'EditCurator' then
      BtnAddCuratorToListClick(BtnAddCuratorToList);
    if TsComboBox(Sender).Name = 'EditRubr' then
      BtnAddCuratorToListClick(BtnAddRubrToList);
    if TsComboBox(Sender).Name = 'EditType' then
      BtnAddCuratorToListClick(BtnAddTypeToList);
    if TsComboBox(Sender).Name = 'EditNapravlenie' then
    begin
      BtnAddCuratorToListClick(BtnAddNaprToList); // EditNapravlenie.DroppedDown := False;
    end;
    if TsComboBox(Sender).Name = 'EditWEB' then
      BtnAddWebToListClick(BtnAddWebToList);
    if TsComboBox(Sender).Name = 'EditEMAIL' then
      BtnAddWebToListClick(BtnAddEMailToList);
    if (TsComboBox(Sender).Name = 'EditPhoneType1') or (TsEdit(Sender).Name = 'EditPhone1') then
      BtnAddPhoneToList1Click(BtnAddPhoneToList1);
    if (TsComboBox(Sender).Name = 'EditPhoneType2') or (TsEdit(Sender).Name = 'EditPhone2') then
      BtnAddPhoneToList1Click(BtnAddPhoneToList2);
    if (TsComboBox(Sender).Name = 'EditPhoneType3') or (TsEdit(Sender).Name = 'EditPhone3') then
      BtnAddPhoneToList1Click(BtnAddPhoneToList3);
    if (TsComboBox(Sender).Name = 'EditPhoneType4') or (TsEdit(Sender).Name = 'EditPhone4') then
      BtnAddPhoneToList1Click(BtnAddPhoneToList4);
    if (TsComboBox(Sender).Name = 'EditPhoneType5') or (TsEdit(Sender).Name = 'EditPhone5') then
      BtnAddPhoneToList1Click(BtnAddPhoneToList5);
    if (TsComboBox(Sender).Name = 'EditPhoneType6') or (TsEdit(Sender).Name = 'EditPhone6') then
      BtnAddPhoneToList1Click(BtnAddPhoneToList6);
    if (TsComboBox(Sender).Name = 'EditPhoneType7') or (TsEdit(Sender).Name = 'EditPhone7') then
      BtnAddPhoneToList1Click(BtnAddPhoneToList7);
    if (TsComboBox(Sender).Name = 'EditPhoneType8') or (TsEdit(Sender).Name = 'EditPhone8') then
      BtnAddPhoneToList1Click(BtnAddPhoneToList8);
    if (TsComboBox(Sender).Name = 'EditPhoneType9') or (TsEdit(Sender).Name = 'EditPhone9') then
      BtnAddPhoneToList1Click(BtnAddPhoneToList9);
    if (TsComboBox(Sender).Name = 'EditPhoneType10') or (TsEdit(Sender).Name = 'EditPhone10') then
      BtnAddPhoneToList1Click(BtnAddPhoneToList10);
  end;
end;

procedure TFormEditor.EditCuratorExit(Sender: TObject);
begin
  if TsComboBox(Sender).Name = 'EditCurator' then
    EditCurator.Text := ''
  else if TsComboBox(Sender).Name = 'EditRubr' then
    EditRubr.Text := ''
  else if TsComboBox(Sender).Name = 'EditType' then
    EditType.Text := ''
  else if TsComboBox(Sender).Name = 'EditNapravlenie' then
    EditNapravlenie.Text := ''
  else if TsComboBox(Sender).Name = 'EditWEB' then
    EditWEB.Text := ''
  else if TsComboBox(Sender).Name = 'EditEMAIL' then
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
    if TNextGrid(Sender).Name = 'SGType' then
      BtnDeleteCuratorFromListClick(BtnDeleteTypeFromList);
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
end;

procedure TFormEditor.SGCuratorDblClick(Sender: TObject);

  procedure Editing(sg: TNextGrid; edit: TsComboBox);
  begin
    if sg.SelectedCount = 0 then
      exit;
    if Assigned(edit) then
    begin
      edit.Text := sg.Cells[0, sg.SelectedRow];
      edit.SetFocus;
    end;
    sg.DeleteRow(sg.SelectedRow);
  end;

begin
  if TNextGrid(Sender).Name = 'SGCurator' then
    Editing(SGCurator, EditCurator);
  if TNextGrid(Sender).Name = 'SGRubr' then
    Editing(SGRubr, EditRubr);
  if TNextGrid(Sender).Name = 'SGType' then
    Editing(SGType, EditType);
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

procedure TFormEditor.BtnAddPhoneToList1Click(Sender: TObject);

  procedure Adding(editPhoneType: TsComboBox; editPhone: TsEdit; sg: TNextGrid);
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

procedure TFormEditor.EditPhone1KeyPress(Sender: TObject; var Key: Char);
begin
  EditCuratorKeyPress(Sender, Key);
  if not (Key in ['0'..'9', #8, #32, '+', '-']) then
    Key := #0;
end;

procedure TFormEditor.btnDeleteAdresClick(Sender: TObject);

  procedure ClearingEdits(OfficeType, Country, City: TsComboBox; ZIP, Street: TsEdit; SGPhone: TNextGrid);
  begin
    OfficeType.Text := '';
    Country.Text := '';
    City.Text := '';
    ZIP.Text := '';
    Street.Text := '';
    SGPhone.ClearRows;
  end;

begin
  if TsSpeedButton(Sender).Name = 'btnDeleteAdres1' then
    ClearingEdits(EditOfficeType1, EditCountry1, EditCity1, EditZIP1, EditStreet1, SGPhone1);
  if TsSpeedButton(Sender).Name = 'btnDeleteAdres2' then
    ClearingEdits(EditOfficeType2, EditCountry2, EditCity2, EditZIP2, EditStreet2, SGPhone2);
  if TsSpeedButton(Sender).Name = 'btnDeleteAdres3' then
    ClearingEdits(EditOfficeType3, EditCountry3, EditCity3, EditZIP3, EditStreet3, SGPhone3);
  if TsSpeedButton(Sender).Name = 'btnDeleteAdres4' then
    ClearingEdits(EditOfficeType4, EditCountry4, EditCity4, EditZIP4, EditStreet4, SGPhone4);
  if TsSpeedButton(Sender).Name = 'btnDeleteAdres5' then
    ClearingEdits(EditOfficeType5, EditCountry5, EditCity5, EditZIP5, EditStreet5, SGPhone5);
  if TsSpeedButton(Sender).Name = 'btnDeleteAdres6' then
    ClearingEdits(EditOfficeType6, EditCountry6, EditCity6, EditZIP6, EditStreet6, SGPhone6);
  if TsSpeedButton(Sender).Name = 'btnDeleteAdres7' then
    ClearingEdits(EditOfficeType7, EditCountry7, EditCity7, EditZIP7, EditStreet7, SGPhone7);
  if TsSpeedButton(Sender).Name = 'btnDeleteAdres8' then
    ClearingEdits(EditOfficeType8, EditCountry8, EditCity8, EditZIP8, EditStreet8, SGPhone8);
  if TsSpeedButton(Sender).Name = 'btnDeleteAdres9' then
    ClearingEdits(EditOfficeType9, EditCountry9, EditCity9, EditZIP9, EditStreet9, SGPhone9);
  if TsSpeedButton(Sender).Name = 'btnDeleteAdres10' then
    ClearingEdits(EditOfficeType10, EditCountry10, EditCity10, EditZIP10, EditStreet10, SGPhone10);
end;

procedure TFormEditor.LoadDataEditor;
var
  i, z: Integer;
  Q: TIBQuery;
begin
  FormLogo.sLabel1.Caption := 'Подключаемые модули ...';
  EditCurator.Clear; // КУРАТОРЫ
  for i := 0 to Main.sgCurator_tmp.RowCount - 1 do
  begin
    EditCurator.AddItem(Main.sgCurator_tmp.Cells[0, i], Pointer(StrToInt(Main.sgCurator_tmp.Cells[1, i])));
  end;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  EditRubr.Clear; // РУБРИКИ
  for i := 0 to Main.sgRubr_tmp.RowCount - 1 do
  begin
    EditRubr.AddItem(Main.sgRubr_tmp.Cells[0, i], Pointer(StrToInt(Main.sgRubr_tmp.Cells[1, i])));
  end;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  EditType.Clear; // ТИП
  for i := 0 to Main.sgType_tmp.RowCount - 1 do
  begin
    EditType.AddItem(Main.sgType_tmp.Cells[0, i], Pointer(StrToInt(Main.sgType_tmp.Cells[1, i])));
  end;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  EditNapravlenie.Clear; // НАПРАВЛЕНИЕ
  for i := 0 to Main.sgNapr_tmp.RowCount - 1 do
  begin
    EditNapravlenie.AddItem(Main.sgNapr_tmp.Cells[0, i], Pointer(StrToInt(Main.sgNapr_tmp.Cells[1, i])));
  end;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  for z := 1 to 10 do
    TsComboBox(FindComponent('EditOfficeType' + IntToStr(z))).Clear;
  Q := TIBQuery.Create(FormEditor);
  Q.Database := FormMain.IBDatabase1;
  Q.Transaction := FormMain.IBTransaction1;
  Q.Close; // ТИП ОФИСА
  Q.SQL.Text := 'select * from OFFICETYPE order by lower(NAME)';
  Q.Open;
  Q.FetchAll;
  for i := 1 to Q.RecordCount do
  begin
    for z := 1 to 10 do
      TsComboBox(FindComponent('EditOfficeType' + IntToStr(z))).AddItem(Q.FieldValues['NAME'], Pointer(Integer(Q.FieldValues['ID'])));
    Q.Next;
  end;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  for z := 1 to 10 do
    TsComboBox(FindComponent('EditCountry' + IntToStr(z))).Clear;
  Q.Close; // СТРАНЫ
  Q.SQL.Text := 'select * from COUNTRY order by lower(NAME)';
  Q.Open;
  Q.FetchAll;
  for i := 1 to Q.RecordCount do
  begin
    for z := 1 to 10 do
      TsComboBox(FindComponent('EditCountry' + IntToStr(z))).AddItem(Q.FieldValues['NAME'], Pointer(Integer(Q.FieldValues['ID'])));
    Q.Next;
  end;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  for z := 1 to 10 do
    TsComboBox(FindComponent('EditCity' + IntToStr(z))).Clear;
  Q.Close; // ГОРОДА
  Q.SQL.Text := 'select * from GOROD order by lower(NAME)';
  Q.Open;
  Q.FetchAll;
  for i := 1 to Q.RecordCount do
  begin
    for z := 1 to 10 do
      TsComboBox(FindComponent('EditCity' + IntToStr(z))).AddItem(Q.FieldValues['NAME'], Pointer(Integer(Q.FieldValues['ID'])));
    Q.Next;
  end;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;
  for z := 1 to 10 do
    TsComboBox(FindComponent('EditPhoneType' + IntToStr(z))).Clear;
  Q.Close; // ТИПЫ ТЕЛЕФОНОВ
  Q.SQL.Text := 'select * from PHONETYPE order by lower(NAME)';
  Q.Open;
  Q.FetchAll;
  for i := 1 to Q.RecordCount do
  begin
    for z := 1 to 10 do
      TsComboBox(FindComponent('EditPhoneType' + IntToStr(z))).AddItem(Q.FieldValues['NAME'], Pointer(Integer(Q.FieldValues['ID'])));
    Q.Next;
  end;
  FormLogo.sGauge1.Progress := FormLogo.sGauge1.Progress + 1;
  Application.ProcessMessages;

  Q.Close;
  Q.Free; // FormMain.IBDatabase1.Close;
end;

function TFormEditor.GetFirmCount: string;
begin
  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text := 'select COUNT(*) from BASE';
  FormMain.IBQuery1.Open;
  Result := FormMain.IBQuery1.FieldByName('COUNT').AsString;
  FormMain.IBQuery1.Close;
  FormMain.IBDatabase1.Close;
end;

function TFormEditor.GetNameByID(table, id: string): string;
var
  Q: TIBQuery;
begin
  Result := '';
  if (Trim(table) = '') or (Trim(id) = '') then
    exit;
  Q := TIBQuery.Create(FormMain);
  Q.Database := FormMain.IBDatabase1;
  Q.Transaction := FormMain.IBTransaction1;
  Q.Close;
  Q.SQL.Text := 'select * from ' + table + ' where id = ' + id;
  Q.Open;
  Q.FetchAll;
  if Q.RecordCount > 0 then
    Result := Q.FieldValues['NAME'];
  Q.Close;
  Q.Free;
end;

function TFormEditor.GetIDByName(component: TsComboBox): string;
var
  index: Integer;
begin
  index := component.Items.IndexOf(Trim(component.Text));
  if index = -1 then
    Result := ''
  else
    Result := IntToStr(Integer(component.Items.Objects[index]));
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
    if component.Name = 'SGType' then
      if Main.sgType_tmp.FindText(1, id, [soCaseInsensitive, soExactMatch]) then
        tmp := tmp + '#' + Main.sgType_tmp.Cells[1, Main.sgType_tmp.SelectedRow] + '$';
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
  SGType.ClearRows;
  SGNapravlenie.ClearRows;
  SGWeb.ClearRows;
  SGEMail.ClearRows;
  CBActivity.Checked := True;
  cbDataRelevance.Checked := True;
  sPageControl1.ActivePageIndex := 0;
  EditName.Text := '';
  EditFIO.Text := '';
  EditCurator.Text := '';
  EditRubr.Text := '';
  EditType.Text := '';
  EditNapravlenie.Text := '';
  EditWEB.Text := '';
  EditEMAIL.Text := '';
  for i := 1 to 10 do
  begin
    TsCheckBox(FindComponent('CBAdres' + IntToStr(i))).Checked := False;
    TsComboBox(FindComponent('EditOfficeType' + IntToStr(i))).Text := '';
    TsComboBox(FindComponent('EditZIP' + IntToStr(i))).Text := '';
    TsComboBox(FindComponent('EditStreet' + IntToStr(i))).Text := '';
    TsComboBox(FindComponent('EditCountry' + IntToStr(i))).Text := '';
    TsComboBox(FindComponent('EditCity' + IntToStr(i))).Text := '';
    TsComboBox(FindComponent('EditPhoneType' + IntToStr(i))).Text := '';
    TsComboBox(FindComponent('EditPhone' + IntToStr(i))).Text := '';
    TNextGrid(FindComponent('SGPhone' + IntToStr(i))).ClearRows;
    TsMemo(FindComponent('MemoPhone' + IntToStr(i))).Clear;
  end;
end;

procedure TFormEditor.AddRecord(Sender: TObject);
var
  str, phones: TStrings;

  procedure AdresProcs(CBAdres: TsCheckBox; No: TsEdit; OfficeType: TsComboBox; ZIP: TsEdit; Street: TsEdit; Country: TsComboBox;
    City: TsComboBox; MemoPhone: TsMemo; SGPhone: TNextGrid);
  var
    i: Integer;
    offtype_id, country_id, city_id: string;
  begin
    if CBAdres.Checked then
    begin
      offtype_id := GetIDByName(OfficeType);
      country_id := GetIDByName(Country);
      city_id := GetIDByName(City);
      str.Add('#1$#' + No.Text + '$#@' + offtype_id + '$#' + Trim(ZIP.Text) + '$#' + Trim(Street.Text) + '$#&' + country_id + '$#^' +
        city_id + '$');
      if SGPhone.RowCount > 0 then
        for i := 0 to SGPhone.RowCount - 1 do
          MemoPhone.Lines.Add(SGPhone.Cells[0, i]);
      phones.Add('#' + No.Text + '$#' + Trim(MemoPhone.Text) + '$')
    end
    else
    begin
      str.Add('#0$#' + No.Text + '$#$#$#$#$#$');
      phones.Add('#' + No.Text + '$#$');
    end;
  end;

  procedure IsAdresFilled(CBAdres: TsCheckBox; OfficeType: TsComboBox; ZIP: TsEdit; Street: TsEdit; Country: TsComboBox; City: TsComboBox;
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
  id, NewRubr, tmp, tmp2, rubrID: string;
  nd, nd2: TTreeNode;
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
  if not (IsDublicate) then
  begin
    FormEditor.Close;
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
    'insert into BASE (ACTIVITY,RELEVANCE,NAME,FIO,CURATOR,RUBR,TYPE,NAPRAVLENIE,WEB,EMAIL,ADRES,PHONES,DATE_ADDED,DATE_EDITED,RELATIONS) values '
    + '(:ACTIVITY,:RELEVANCE,:NAME,:FIO,:CURATOR,:RUBR,:TYPE,:NAPRAVLENIE,:WEB,:EMAIL,:ADRES,:PHONES,:DATE_ADDED,:DATE_EDITED,:RELATIONS)';
  FormMain.IBQuery1.ParamByName('ACTIVITY').AsBoolean := CBActivity.Checked;
  FormMain.IBQuery1.ParamByName('RELEVANCE').AsBoolean := cbDataRelevance.Checked;
  FormMain.IBQuery1.ParamByName('NAME').AsString := EditName.Text;
  FormMain.IBQuery1.ParamByName('FIO').AsString := EditFIO.Text;
  FormMain.IBQuery1.ParamByName('CURATOR').AsString := GetIDString(SGCurator);
  FormMain.IBQuery1.ParamByName('RUBR').AsString := GetIDString(SGRubr);
  FormMain.IBQuery1.ParamByName('TYPE').AsString := GetIDString(SGType);
  FormMain.IBQuery1.ParamByName('NAPRAVLENIE').AsString := GetIDString(SGNapravlenie);
  FormMain.IBQuery1.ParamByName('WEB').AsString := GetWebEmailString(SGWeb);
  FormMain.IBQuery1.ParamByName('EMAIL').AsString := GetWebEmailString(SGEMail);

  str := TStringList.Create;
  phones := TStringList.Create;
  IsAdresFilled(CBAdres1, EditOfficeType1, EditZIP1, EditStreet1, EditCountry1, EditCity1, SGPhone1);
  IsAdresFilled(CBAdres2, EditOfficeType2, EditZIP2, EditStreet2, EditCountry2, EditCity2, SGPhone2);
  IsAdresFilled(CBAdres3, EditOfficeType3, EditZIP3, EditStreet3, EditCountry3, EditCity3, SGPhone3);
  IsAdresFilled(CBAdres4, EditOfficeType4, EditZIP4, EditStreet4, EditCountry4, EditCity4, SGPhone4);
  IsAdresFilled(CBAdres5, EditOfficeType5, EditZIP5, EditStreet5, EditCountry5, EditCity5, SGPhone5);
  IsAdresFilled(CBAdres6, EditOfficeType6, EditZIP6, EditStreet6, EditCountry6, EditCity6, SGPhone6);
  IsAdresFilled(CBAdres7, EditOfficeType7, EditZIP7, EditStreet7, EditCountry7, EditCity7, SGPhone7);
  IsAdresFilled(CBAdres8, EditOfficeType8, EditZIP8, EditStreet8, EditCountry8, EditCity8, SGPhone8);
  IsAdresFilled(CBAdres9, EditOfficeType9, EditZIP9, EditStreet9, EditCountry9, EditCity9, SGPhone9);
  IsAdresFilled(CBAdres10, EditOfficeType10, EditZIP10, EditStreet10, EditCountry10, EditCity10, SGPhone10);
  AdresProcs(CBAdres1, EditNO1, EditOfficeType1, EditZIP1, EditStreet1, EditCountry1, EditCity1, MemoPhone1, SGPhone1);
  AdresProcs(CBAdres2, EditNO2, EditOfficeType2, EditZIP2, EditStreet2, EditCountry2, EditCity2, MemoPhone2, SGPhone2);
  AdresProcs(CBAdres3, EditNO3, EditOfficeType3, EditZIP3, EditStreet3, EditCountry3, EditCity3, MemoPhone3, SGPhone3);
  AdresProcs(CBAdres4, EditNO4, EditOfficeType4, EditZIP4, EditStreet4, EditCountry4, EditCity4, MemoPhone4, SGPhone4);
  AdresProcs(CBAdres5, EditNO5, EditOfficeType5, EditZIP5, EditStreet5, EditCountry5, EditCity5, MemoPhone5, SGPhone5);
  AdresProcs(CBAdres6, EditNO6, EditOfficeType6, EditZIP6, EditStreet6, EditCountry6, EditCity6, MemoPhone6, SGPhone6);
  AdresProcs(CBAdres7, EditNO7, EditOfficeType7, EditZIP7, EditStreet7, EditCountry7, EditCity7, MemoPhone7, SGPhone7);
  AdresProcs(CBAdres8, EditNO8, EditOfficeType8, EditZIP8, EditStreet8, EditCountry8, EditCity8, MemoPhone8, SGPhone8);
  AdresProcs(CBAdres9, EditNO9, EditOfficeType9, EditZIP9, EditStreet9, EditCountry9, EditCity9, MemoPhone9, SGPhone9);
  AdresProcs(CBAdres10, EditNO10, EditOfficeType10, EditZIP10, EditStreet10, EditCountry10, EditCity10, MemoPhone10, SGPhone10);
  FormMain.IBQuery1.ParamByName('ADRES').AsString := str.Text;
  FormMain.IBQuery1.ParamByName('PHONES').AsString := phones.Text;
  str.Free;
  phones.Free;
  FormMain.IBQuery1.ParamByName('DATE_ADDED').AsString := DateToStr(Now);
  FormMain.IBQuery1.ParamByName('DATE_EDITED').AsString := DateToStr(Now);
  FormMain.IBQuery1.ParamByName('RELATIONS').AsString := BuildRelations;
  try
    FormMain.IBQuery1.ExecSQL;
  except
    on E: Exception do
    begin
      FormMain.WriteLog('TFormEditor.AddRecord' + #13 + 'Ошибка: ' + E.Message);
      MessageBox(handle, PChar('Ошибка при создании фирмы.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      exit;
    end;
  end;
  FormMain.IBTransaction1.CommitRetaining;
  FormMain.sStatusBar1.Panels[1].Text := 'Фирм в базе: ' + GetFirmCount;

  FormMain.IBQuery1.Close;
  FormMain.IBQuery1.SQL.Text := 'select MAX(ID) from BASE';
  FormMain.IBQuery1.Open;
  id := FormMain.IBQuery1.Fields[0].Value;
  FormMain.WriteLog('TFormEditor.AddRecord: запись добавлена ' + id);
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
  FormMain.TVRubrikator.CustomSort(@CustomSortProc, 0, True);
  FormMain.IBQuery1.Close;

  FormRelations.ClearEdits;
  FormRelations.PrepareRecord(id);
  FormRelations.Show;
  // FormEditor.Close; // закрывается в  if NOT(IsDublicate) then begin ^^^
  // FormMain.IBDatabase1.Connected := False;
end;

procedure TFormEditor.PrepareEditRecord(id: string);

  procedure AdresProcs(AdresList: TStrings; Num: Integer; CBAdres: TsCheckBox; OfficeType: TsComboBox; ZIP: TsEdit; Street: TsEdit;
    Country: TsComboBox; City: TsComboBox);
  var
    tmp, city_str, country_str, ofType, s: string;
    list: TStrings;
  begin
    if length(AdresList.Text) = 0 then
      exit; // Грузятся адреса {begin}
    tmp := AdresList[Num - 1];
    list := TStringList.Create;
    while pos('$', tmp) > 0 do
    begin
      s := copy(tmp, 0, pos('$', tmp));
      delete(tmp, 1, length(s));
      delete(s, 1, 1);
      delete(s, length(s), 1);
      list.Add(s);
    end;
    if list[0] = '0' then
    begin
      CBAdres.Checked := False;
      list.Free;
      exit;
    end;
    // такая же процедура в Main.OpenTabByID и Report.GenerateReport и ReportSimple.GenerateReport
    if list[0] = '1' then
      CBAdres.Checked := True;
    city_str := list[6];
    if city_str[1] = '^' then
      delete(city_str, 1, 1);
    country_str := list[5];
    if country_str[1] = '&' then
      delete(country_str, 1, 1);
    ofType := list[2];
    if ofType[1] = '@' then
      delete(ofType, 1, 1);
    OfficeType.Text := GetNameByID('OFFICETYPE', ofType);
    ZIP.Text := list[3];
    Street.Text := list[4];
    Country.Text := GetNameByID('COUNTRY', country_str);
    City.Text := GetNameByID('GOROD', city_str);
    list.Free; // Грузятся адреса {end}
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
      if component.Name = 'SGType' then
        if Main.sgType_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
        begin
          name := Main.sgType_tmp.Cells[0, Main.sgType_tmp.SelectedRow];
          id := Main.sgType_tmp.Cells[1, Main.sgType_tmp.SelectedRow];
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
  FormMain.IBQuery1.FetchAll;
  if FormMain.IBQuery1.RecordCount = 0 then
    exit;
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
  if FormMain.IBQuery1.FieldValues['TYPE'] <> null then
    LoadDataToGrids(SGType, FormMain.IBQuery1.FieldValues['TYPE']);
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
    AdresProcs(str, 1, CBAdres1, EditOfficeType1, EditZIP1, EditStreet1, EditCountry1, EditCity1);
    AdresProcs(str, 2, CBAdres2, EditOfficeType2, EditZIP2, EditStreet2, EditCountry2, EditCity2);
    AdresProcs(str, 3, CBAdres3, EditOfficeType3, EditZIP3, EditStreet3, EditCountry3, EditCity3);
    AdresProcs(str, 4, CBAdres4, EditOfficeType4, EditZIP4, EditStreet4, EditCountry4, EditCity4);
    AdresProcs(str, 5, CBAdres5, EditOfficeType5, EditZIP5, EditStreet5, EditCountry5, EditCity5);
    AdresProcs(str, 6, CBAdres6, EditOfficeType6, EditZIP6, EditStreet6, EditCountry6, EditCity6);
    AdresProcs(str, 7, CBAdres7, EditOfficeType7, EditZIP7, EditStreet7, EditCountry7, EditCity7);
    AdresProcs(str, 8, CBAdres8, EditOfficeType8, EditZIP8, EditStreet8, EditCountry8, EditCity8);
    AdresProcs(str, 9, CBAdres9, EditOfficeType9, EditZIP9, EditStreet9, EditCountry9, EditCity9);
    AdresProcs(str, 10, CBAdres10, EditOfficeType10, EditZIP10, EditStreet10, EditCountry10, EditCity10);
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
  NewRubr, rubrID, tmp: string;
  nd, nd2: TTreeNode;
  i: Integer;

  procedure AdresProcs(CBAdres: TsCheckBox; No: TsEdit; OfficeType: TsComboBox; ZIP: TsEdit; Street: TsEdit; Country: TsComboBox;
    City: TsComboBox; MemoPhone: TsMemo; SGPhone: TNextGrid);
  var
    i: Integer;
    offtype_id, country_id, city_id: string;
  begin
    if CBAdres.Checked then
    begin
      offtype_id := GetIDByName(OfficeType);
      country_id := GetIDByName(Country);
      city_id := GetIDByName(City);
      str.Add('#1$#' + No.Text + '$#@' + offtype_id + '$#' + Trim(ZIP.Text) + '$#' + Trim(Street.Text) + '$#&' + country_id + '$#^' +
        city_id + '$');
      if SGPhone.RowCount > 0 then
        for i := 0 to SGPhone.RowCount - 1 do
          MemoPhone.Lines.Add(SGPhone.Cells[0, i]);
      phones.Add('#' + No.Text + '$#' + Trim(MemoPhone.Text) + '$')
    end
    else
    begin
      str.Add('#0$#' + No.Text + '$#$#$#$#$#$');
      phones.Add('#' + No.Text + '$#$');
    end;
  end;

  procedure IsAdresFilled(CBAdres: TsCheckBox; OfficeType: TsComboBox; ZIP: TsEdit; Street: TsEdit; Country: TsComboBox; City: TsComboBox;
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
  IsNewRecordCheck;
  EditName.Text := UpperFirst(EditName.Text);
  EditFIO.Text := UpperFirst(EditFIO.Text);
  FormMain.IBQuery1.Close;
  NewRubr := GetIDString(SGRubr);
  FormMain.IBQuery1.SQL.Text := 'update BASE set ACTIVITY=:ACTIVITY,RELEVANCE=:RELEVANCE,NAME=:NAME,FIO=:FIO,CURATOR=:CURATOR,' +
    'RUBR=:RUBR,TYPE=:TYPE,NAPRAVLENIE=:NAPRAVLENIE,WEB=:WEB,EMAIL=:EMAIL,ADRES=:ADRES,PHONES=:PHONES,DATE_EDITED=:DATE_EDITED,RELATIONS=:RELATIONS where ID=:ID';
  FormMain.IBQuery1.ParamByName('ACTIVITY').AsBoolean := CBActivity.Checked;
  FormMain.IBQuery1.ParamByName('RELEVANCE').AsBoolean := cbDataRelevance.Checked;
  FormMain.IBQuery1.ParamByName('NAME').AsString := EditName.Text;
  FormMain.IBQuery1.ParamByName('FIO').AsString := EditFIO.Text;
  FormMain.IBQuery1.ParamByName('CURATOR').AsString := GetIDString(SGCurator);
  FormMain.IBQuery1.ParamByName('RUBR').AsString := NewRubr;
  FormMain.IBQuery1.ParamByName('TYPE').AsString := GetIDString(SGType);
  FormMain.IBQuery1.ParamByName('NAPRAVLENIE').AsString := GetIDString(SGNapravlenie);
  FormMain.IBQuery1.ParamByName('WEB').AsString := GetWebEmailString(SGWeb);
  FormMain.IBQuery1.ParamByName('EMAIL').AsString := GetWebEmailString(SGEMail);

  str := TStringList.Create;
  phones := TStringList.Create;
  IsAdresFilled(CBAdres1, EditOfficeType1, EditZIP1, EditStreet1, EditCountry1, EditCity1, SGPhone1);
  IsAdresFilled(CBAdres2, EditOfficeType2, EditZIP2, EditStreet2, EditCountry2, EditCity2, SGPhone2);
  IsAdresFilled(CBAdres3, EditOfficeType3, EditZIP3, EditStreet3, EditCountry3, EditCity3, SGPhone3);
  IsAdresFilled(CBAdres4, EditOfficeType4, EditZIP4, EditStreet4, EditCountry4, EditCity4, SGPhone4);
  IsAdresFilled(CBAdres5, EditOfficeType5, EditZIP5, EditStreet5, EditCountry5, EditCity5, SGPhone5);
  IsAdresFilled(CBAdres6, EditOfficeType6, EditZIP6, EditStreet6, EditCountry6, EditCity6, SGPhone6);
  IsAdresFilled(CBAdres7, EditOfficeType7, EditZIP7, EditStreet7, EditCountry7, EditCity7, SGPhone7);
  IsAdresFilled(CBAdres8, EditOfficeType8, EditZIP8, EditStreet8, EditCountry8, EditCity8, SGPhone8);
  IsAdresFilled(CBAdres9, EditOfficeType9, EditZIP9, EditStreet9, EditCountry9, EditCity9, SGPhone9);
  IsAdresFilled(CBAdres10, EditOfficeType10, EditZIP10, EditStreet10, EditCountry10, EditCity10, SGPhone10);
  AdresProcs(CBAdres1, EditNO1, EditOfficeType1, EditZIP1, EditStreet1, EditCountry1, EditCity1, MemoPhone1, SGPhone1);
  AdresProcs(CBAdres2, EditNO2, EditOfficeType2, EditZIP2, EditStreet2, EditCountry2, EditCity2, MemoPhone2, SGPhone2);
  AdresProcs(CBAdres3, EditNO3, EditOfficeType3, EditZIP3, EditStreet3, EditCountry3, EditCity3, MemoPhone3, SGPhone3);
  AdresProcs(CBAdres4, EditNO4, EditOfficeType4, EditZIP4, EditStreet4, EditCountry4, EditCity4, MemoPhone4, SGPhone4);
  AdresProcs(CBAdres5, EditNO5, EditOfficeType5, EditZIP5, EditStreet5, EditCountry5, EditCity5, MemoPhone5, SGPhone5);
  AdresProcs(CBAdres6, EditNO6, EditOfficeType6, EditZIP6, EditStreet6, EditCountry6, EditCity6, MemoPhone6, SGPhone6);
  AdresProcs(CBAdres7, EditNO7, EditOfficeType7, EditZIP7, EditStreet7, EditCountry7, EditCity7, MemoPhone7, SGPhone7);
  AdresProcs(CBAdres8, EditNO8, EditOfficeType8, EditZIP8, EditStreet8, EditCountry8, EditCity8, MemoPhone8, SGPhone8);
  AdresProcs(CBAdres9, EditNO9, EditOfficeType9, EditZIP9, EditStreet9, EditCountry9, EditCity9, MemoPhone9, SGPhone9);
  AdresProcs(CBAdres10, EditNO10, EditOfficeType10, EditZIP10, EditStreet10, EditCountry10, EditCity10, MemoPhone10, SGPhone10);

  FormMain.IBQuery1.ParamByName('ADRES').AsString := str.Text;
  FormMain.IBQuery1.ParamByName('PHONES').AsString := phones.Text;
  FormMain.IBQuery1.ParamByName('DATE_EDITED').AsString := DateToStr(Now);
  FormMain.IBQuery1.ParamByName('RELATIONS').AsString := BuildRelations;
  FormMain.IBQuery1.ParamByName('ID').AsString := lblID.Caption;
  str.Free;
  phones.Free;

  try
    FormMain.IBQuery1.ExecSQL;
    FormMain.WriteLog('TFormEditor.EditRecord: запись отредактирована ' + lblID.Caption);
  except
    on E: Exception do
    begin
      FormMain.WriteLog('TFormEditor.EditRecord' + #13 + 'Ошибка: ' + E.Message);
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

  FormMain.TVRubrikator.CustomSort(@CustomSortProc, 0, True);
  FormMain.TVRubrikator.Items.EndUpdate;
  FormMain.IBQuery1.Close;
  FormEditor.Close;

  FormRelations.ClearEdits;
  FormRelations.PrepareRecord(lblID.Caption);
  FormRelations.Show;
  // FormMain.IBDatabase1.Connected := False;
end;

procedure TFormEditor.DeleteRecord(id: string);
var
  i: Integer;
  Q: TIBQuery;
  capt: string;
begin
  Q := TIBQuery.Create(FormEditor);
  Q.Database := FormMain.IBDatabase1;
  Q.Transaction := FormMain.IBTransaction1;
  Q.Close;
  Q.SQL.Text := 'select NAME from BASE where ID = :ID';
  Q.ParamByName('ID').AsString := id;
  Q.Open;
  capt := Q.Fields[0].Value;
  if MessageBox(handle, 'Запись будет удалена. Продолжить?', PChar(capt), MB_YESNO or MB_ICONQUESTION) = MRNO then
  begin
    Q.Close;
    Q.Free;
    exit;
  end;
  Q.Close;
  Q.SQL.Text := 'delete from BASE where ID = :ID';
  Q.ParamByName('ID').AsString := id;
  try
    Q.ExecSQL;
    FormMain.WriteLog('TFormEditor.DeleteRecord: запись удалена ' + id);
  except
    on E: Exception do
    begin
      FormMain.WriteLog('TFormEditor.DeleteRecord' + #13 + 'Ошибка: ' + E.Message);
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

procedure TFormEditor.IsNewRecordCheck;
var
  i, x: Integer;
  isNewRubr: Boolean;
  edit: TsComboBox;

  procedure AddingNewRec(table, Value: string; sg_Main, sg_Directory: TNextGrid);
  var
    Q: TIBQuery;
    id: string;
    node: TTreeNode;
    z: Integer;
  begin
    Q := TIBQuery.Create(FormEditor);
    Q.Database := FormMain.IBDatabase1;
    Q.Transaction := FormMain.IBTransaction1;
    Q.Close;
    Q.SQL.Text := 'insert into ' + table + ' (NAME) values (:NAME)';
    Q.ParamByName('NAME').AsString := Value;
    try
      Q.ExecSQL;
    except
      on E: Exception do
      begin
        FormMain.WriteLog('TFormEditor.IsNewRecordCheck' + #13 + 'Ошибка: ' + E.Message);
        MessageBox(handle, PChar('Ошибка при проверке существующих директорий.' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
        Q.Close;
        Q.Free;
        exit;
      end;
    end;
    FormMain.IBTransaction1.CommitRetaining;
    Q.Close;
    Q.SQL.Text := 'select MAX(ID) from ' + table;
    Q.Open;
    Q.FetchAll;
    id := Q.Fields[0].Value;
    Q.Close;
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
    if AnsiLowerCase(table) = 'type' then
    begin
      SGType.Cells[1, i] := id;
      EditType.AddItem(Value, Pointer(StrToInt(id)));
    end;
    if AnsiLowerCase(table) = 'napravlenie' then
    begin
      SGNapravlenie.Cells[1, i] := id;
      EditNapravlenie.AddItem(Value, Pointer(StrToInt(id)));
    end;
    if AnsiLowerCase(table) = 'officetype' then
      for z := 1 to 10 do
        TsComboBox(FindComponent('EditOfficeType' + IntToStr(z))).AddItem(Value, Pointer(StrToInt(id)));
    if AnsiLowerCase(table) = 'country' then
      for z := 1 to 10 do
        TsComboBox(FindComponent('EditCountry' + IntToStr(z))).AddItem(Value, Pointer(StrToInt(id)));
    if AnsiLowerCase(table) = 'gorod' then
      for z := 1 to 10 do
        TsComboBox(FindComponent('EditCity' + IntToStr(z))).AddItem(Value, Pointer(StrToInt(id)));
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
    if SGCurator.Cells[1, i] = '' then
    begin
      AddingNewRec('CURATOR', SGCurator.Cells[0, i], Main.sgCurator_tmp, FormDirectory.SGCurator);
    end;
  for i := 0 to SGRubr.RowCount - 1 do
    if SGRubr.Cells[1, i] = '' then
    begin
      AddingNewRec('RUBRIKATOR', SGRubr.Cells[0, i], Main.sgRubr_tmp, FormDirectory.SGRubr);
      isNewRubr := True;
    end;
  for i := 0 to SGType.RowCount - 1 do
    if SGType.Cells[1, i] = '' then
    begin
      AddingNewRec('TYPE', SGType.Cells[0, i], Main.sgType_tmp, FormDirectory.SGFirmType);
    end;
  for i := 0 to SGNapravlenie.RowCount - 1 do
    if SGNapravlenie.Cells[1, i] = '' then
    begin
      AddingNewRec('NAPRAVLENIE', SGNapravlenie.Cells[0, i], Main.sgNapr_tmp, FormDirectory.SGNapr);
    end;
  for x := 1 to 10 do
  begin
    edit := TsComboBox(FindComponent('EditOfficeType' + IntToStr(x)));
    if (edit.Items.IndexOf(Trim(edit.Text)) = -1) and (Trim(edit.Text) <> '') then
      AddingNewRec('OFFICETYPE', UpperFirst(edit.Text), nil, FormDirectory.SGOfficeType);
    edit := TsComboBox(FindComponent('EditCountry' + IntToStr(x)));
    if (edit.Items.IndexOf(Trim(edit.Text)) = -1) and (Trim(edit.Text) <> '') then
      AddingNewRec('COUNTRY', UpperFirst(edit.Text), nil, FormDirectory.SGCountry);
    edit := TsComboBox(FindComponent('EditCity' + IntToStr(x)));
    if (edit.Items.IndexOf(Trim(edit.Text)) = -1) and (Trim(edit.Text) <> '') then
      AddingNewRec('GOROD', UpperFirst(edit.Text), nil, FormDirectory.SGCity);
  end;
  if isNewRubr then
    FormMain.TVRubrikator.CustomSort(@CustomSortProc, 0, True);
end;

procedure TFormEditor.IsRecordDublicate;
var
  Name, WEBstr, WEBtmp, EMAILstr, EMAILtmp, PHONEstr: string;
  WEBlist, EMAILlist, PHONElist: TStrings;
  REQ, REQ1, REQ2, REQ3, REQ4: string;
  x: Integer;
  Q: TIBQuery;

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
  Name := AnsiLowerCase(Trim(EditName.Text));
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

  Q := TIBQuery.Create(FormEditor);
  Q.Database := FormMain.IBDatabase1;
  Q.Transaction := FormMain.IBTransaction1;
  Q.Close;
  Q.SQL.Text := REQ;
  Q.Params[0].AsString := NAME;
  try
    Q.Open;
  except
    on E: Exception do
    begin
      FormMain.WriteLog('TFormEditor.IsRecordDublicate' + #13 + 'Ошибка: ' + E.Message);
      MessageBox(handle, PChar('Произошла ошибка при проверке фирм дубликатов' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      exit;
    end;
  end;
  Q.FetchAll;
  if Q.RecordCount > 0 then
  begin
    FormDublicate.SGDublicate.ClearRows;
    FormDublicate.SGDublicate.BeginUpdate;
    for x := 1 to Q.RecordCount do
    begin
      FormDublicate.SGDublicate.AddRow;
      if Q.FieldValues['ACTIVITY'] = '-1' then
        FormDublicate.SGDublicate.Cells[0, FormDublicate.SGDublicate.LastAddedRow] := '1'
      else
        FormDublicate.SGDublicate.Cells[0, FormDublicate.SGDublicate.LastAddedRow] := '2';
      FormDublicate.SGDublicate.Cells[1, FormDublicate.SGDublicate.LastAddedRow] := Q.FieldValues['NAME'];
      FormDublicate.SGDublicate.Cells[2, FormDublicate.SGDublicate.LastAddedRow] := Q.FieldValues['ID'];
      Q.Next;
    end;
    FormDublicate.SGDublicate.EndUpdate;
    FormDublicate.BtnOK.OnClick := AddRecord;
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

end.
