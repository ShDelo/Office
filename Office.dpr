program Office;

uses
  Forms,
  Windows,
  Main in 'Main.pas' {FormMain},
  Editor in 'Editor.pas' {FormEditor},
  Logo in 'Logo.pas' {FormLogo},
  MailSend in 'MailSend.pas' {FormMailSender},
  Report in 'Report.pas' {FormReport},
  Directory in 'Directory.pas' {FormDirectory},
  ReportSimple in 'ReportSimple.pas' {FormReportSimple},
  Dublicate in 'Dublicate.pas' {FormDublicate},
  Relations in 'Relations.pas' {FormRelations},
  DirectoryQuery in 'DirectoryQuery.pas' {FormDirectoryQuery},
  Helpers in 'Helpers.pas';

{$R *.res}

var
  k : Cardinal;

begin

  CreateMutex(nil,false,'{36954184-2EAF-46D8-B490-AD45A71DDF8B}');
  k := GetLastError();
  if (k=ERROR_ALREADY_EXISTS)or(k=ERROR_ACCESS_DENIED) then
  begin
   Application.Terminate;
   Exit;
  end;

  Application.Initialize;
  Application.Title := 'Швейное Дело';
  FormLogo := TFormLogo.Create(Application);
  FormLogo.Show;
//  FormLogo.Update;
  Application.CreateForm(TFormMain, FormMain);
  Application.CreateForm(TFormEditor, FormEditor);
  Application.CreateForm(TFormMailSender, FormMailSender);
  Application.CreateForm(TFormReport, FormReport);
  Application.CreateForm(TFormDirectory, FormDirectory);
  Application.CreateForm(TFormReportSimple, FormReportSimple);
  Application.CreateForm(TFormDublicate, FormDublicate);
  Application.CreateForm(TFormRelations, FormRelations);
  Application.CreateForm(TFormDirectoryQuery, FormDirectoryQuery);
  FormLogo.Hide;
  FormLogo.Free;  
  Application.Run;
end.
