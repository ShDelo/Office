unit DirectoryQuery;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, sButton, Vcl.ComCtrls, sComboBoxes, Vcl.ExtCtrls,
  sPanel, NxScrollControl, NxCustomGridControl, NxCustomGrid, NxGrid, IBC, Directory;

type
  TFormDirectoryQuery = class(TForm)
    PanelGeneral: TsPanel;
    Edit3: TsComboBoxEx;
    BtnCancel: TsButton;
    BtnOK: TsButton;
    Edit1: TsComboBoxEx;
    Edit2: TsComboBoxEx;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LoadDataDirectoryQuery;
    function ShowDirectoryQuery(DirCode: integer; Method: string; Data: TDirectoryData; var EnteredValues: array of string): boolean;
    procedure BtnOKClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormDirectoryQuery: TFormDirectoryQuery;

implementation

uses Main;

{$R *.dfm}
{ TFormDirectoryQuery }

procedure TFormDirectoryQuery.FormCreate(Sender: TObject);
begin
  { #TODO2: DESIGN : ComboBoxed 1 and 2 does not using drop-down capabilty. They are currently empty }
  LoadDataDirectoryQuery;
end;

procedure TFormDirectoryQuery.FormShow(Sender: TObject);
begin
  Edit1.SetFocus;
end;

procedure TFormDirectoryQuery.LoadDataDirectoryQuery;
var
  Query: TIBCQuery;
begin
  Query := QueryCreate;
  try
    Query.SQL.Text := 'select * from REGION order by lower(NAME)';
    Query.Open;
    Query.FetchAll := True;
    while not Query.Eof do
    begin
      Edit3.AddItem(Query.FieldByName('NAME').AsString, TObject(Query.FieldByName('ID').AsInteger));
      Query.Next;
    end;
  finally
    Query.Free;
  end;
end;

procedure TFormDirectoryQuery.BtnOKClick(Sender: TObject);
begin
  if Trim(Edit1.Text) = EmptyStr then
  begin
    MessageBox(handle, 'Необходимо указать название директории.', 'Информация', MB_OK or MB_ICONINFORMATION);
    Edit1.SetFocus;
    exit;
  end;

  if BtnOK.Tag in [DIR_CODE_CITY] then
  begin
    if Edit3.ItemIndex = -1 then
    begin
      { #TODO1: DESIGN : allow new region entry creation from here?! if olbast.text <> EmptySTR and itemIndex = -1
        i.e region doesn't exist but we entered text, we can create region from here? }
      MessageBox(handle, 'Необходимо указать область для города.', 'Информация', MB_OK or MB_ICONINFORMATION);
      Edit3.SetFocus;
      exit;
    end;
  end;

  FormDirectoryQuery.ModalResult := mrOk;
end;

function TFormDirectoryQuery.ShowDirectoryQuery(DirCode: integer; Method: string; Data: TDirectoryData;
  var EnteredValues: array of string): boolean;
begin
  Result := False;

  if not DirCode in [0 .. DIR_CODE_TOTAL] then
  begin
    MessageBox(Handle, PChar('Неизвестный DirCode(' + IntToStr(DirCode) + ') передаётся в FormDirectoryQuery.'), 'Ошибка',
      MB_OK or MB_ICONERROR);
    exit;
  end;

  BtnOK.Tag := -1;
  Edit1.Items.Clear;
  Edit1.Text := EmptyStr;
  Edit2.Items.Clear;
  Edit2.Text := EmptyStr;
  Edit3.Text := EmptyStr;
  Edit3.ItemIndex := -1;
  Edit1.BoundLabel.Caption := EmptyStr;
  Edit2.BoundLabel.Caption := EmptyStr;
  Edit3.BoundLabel.Caption := EmptyStr;

  if AnsiLowerCase(Method) = 'add' then
  begin
    FormDirectoryQuery.Caption := 'Создать директорию';
    BtnOK.Caption := 'Создать';
  end
  else if AnsiLowerCase(Method) = 'edit' then
  begin
    FormDirectoryQuery.Caption := 'Редактировать директорию';
    BtnOK.Caption := 'Редактировать';
  end;

  Edit1.MaxLength := 255;
  Edit2.MaxLength := 255;
  Edit3.MaxLength := 255;

  case DirCode of
    DIR_CODE_CURATOR:
      Edit1.BoundLabel.Caption := 'Куратор:';
    DIR_CODE_RUBRIKA:
      Edit1.BoundLabel.Caption := 'Рубрика:';
    DIR_CODE_FIRMTYPE:
      Edit1.BoundLabel.Caption := 'Тип фирмы:';
    DIR_CODE_NAPRAVLENIE:
      begin
        Edit1.MaxLength := 5000;
        Edit1.BoundLabel.Caption := 'Вид деятельности:';
      end;
    DIR_CODE_OFFICETYPE:
      Edit1.BoundLabel.Caption := 'Тип адреса:';
    DIR_CODE_COUNTRY:
      Edit1.BoundLabel.Caption := 'Страна:';
    DIR_CODE_REGION:
      Edit1.BoundLabel.Caption := 'Область:';
    DIR_CODE_CITY:
      begin
        Edit1.BoundLabel.Caption := 'Город (рус):';
        Edit2.BoundLabel.Caption := 'Город (укр):';
        Edit3.BoundLabel.Caption := 'Область:';
      end;
    DIR_CODE_PHONETYPE:
      Edit1.BoundLabel.Caption := 'Тип телефона:';
  end;

  if DirCode in [DIR_CODE_CITY] then
  begin
    // show additional edits for City directory
    Edit2.Visible := True;
    Edit3.Visible := True;
    FormDirectoryQuery.Height := 218;

    // load data to edits
    if AnsiLowerCase(Method) = 'edit' then
    begin
      Edit1.Text := Data.Name1;
      Edit2.Text := Data.Name2;
      Edit3.ItemIndex := Edit3.Items.IndexOfObject(TObject(Data.ID_REGION));
    end;
  end
  else
  begin
    Edit2.Visible := False;
    Edit3.Visible := False;
    FormDirectoryQuery.Height := 143;

    // load data to edits
    if AnsiLowerCase(Method) = 'edit' then
    begin
      Edit1.Text := Data.Name1;
    end
  end;

  BtnOK.Tag := DirCode;

  if FormDirectoryQuery.ShowModal = mrOk then
  begin
    // return entered values
    if DirCode = DIR_CODE_PHONETYPE then
      EnteredValues[0] := Trim(AnsiLowerCase(Edit1.Text))
    else
      EnteredValues[0] := Trim(UpperFirst(Edit1.Text));

    EnteredValues[1] := Trim(UpperFirst(Edit2.Text));

    if Edit3.ItemIndex = -1 then
      EnteredValues[2] := '-1'
    else
      EnteredValues[2] := IntToStr(Integer(Edit3.Items.Objects[Edit3.ItemIndex]));

    Result := True;
  end;
end;

end.
