unit Dublicate;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, sButton, NxColumns, NxColumnClasses, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxGrid;

type
  TFormDublicate = class(TForm)
    SGDublicate: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    BtnOK: TsButton;
    BtnCancel: TsButton;
    NxImageColumn1: TNxImageColumn;
    procedure BtnCancelClick(Sender: TObject);
    procedure SGDublicateDblClick(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams); override;
    { Public declarations }
  end;

var
  FormDublicate: TFormDublicate;

implementation

uses Main;

{$R *.dfm}

procedure TFormDublicate.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.ExStyle := Params.ExStyle or WS_Ex_AppWindow;
end;

procedure TFormDublicate.BtnCancelClick(Sender: TObject);
begin
  FormDublicate.Close;
end;

procedure TFormDublicate.SGDublicateDblClick(Sender: TObject);
var
  ID: string;
  i: integer;
begin
  if SGDublicate.SelectedCount = 0 then
    exit;
  ID := SGDublicate.Cells[2, SGDublicate.SelectedRow];
  for i := 0 to FormMain.sPageControl1.PageCount - 1 do
  begin
    if FormMain.sPageControl1.Pages[i].Name = 'Tab' + ID then
    begin
      FormMain.sPageControl1.Pages[i].Show;
      exit;
    end;
  end;
  FormMain.OpenTabByID(ID);
end;

end.
