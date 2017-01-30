unit Logo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, sGauge, sLabel, ComCtrls, acProgressBar, acPNG;

type
  TFormLogo = class(TForm)
    Image1: TImage;
    sGauge1: TsGauge;
    sLabel1: TsLabel;
    procedure Loading(msg: string);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormLogo: TFormLogo;
  LoadingProgress: integer = 0;

implementation

uses Main;

{$R *.dfm}

procedure TFormLogo.Loading(msg: string);
begin
  INC(LoadingProgress);
  sLabel1.Caption := msg;
  sGauge1.Progress := LoadingProgress;
  Application.ProcessMessages;
end;

procedure TFormLogo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

end.
