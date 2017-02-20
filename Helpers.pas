unit Helpers;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, ComCtrls, Forms, sComboBoxes,
  NxGrid, IBC;

procedure debug(Text: string; Params: array of TVarRec);
procedure WriteLog(Text: string);
function CustomSortProc(Node1, Node2: TTreeNode; iUpToThisLevel: integer): integer; stdcall;
function QueryCreate: TIBCQuery;
function AppPath: string;
function UpperFirst(s: string): string;

function GetFirmCount: string;
function GetNameByID(table, id: string): string;
function GetIDByName(component: TsComboBoxEx): string;
function GetIndexOfText(component: TsComboBoxEx; TextToIndex: string = ''): integer;
function GetIndexOfObject(component: TsComboBoxEx; IDToIndex: integer = -1): integer;
function GetObjectOfIndex(component: TsComboBoxEx): integer;

const
  bDebug: Boolean = True;

implementation

uses
  Main;

procedure debug(Text: string; Params: array of TVarRec);
begin
  if bDebug then
  begin
    FormMain.memoDebug.Lines.Add(DateTimeToStr(now) + ': ' + Format(Text, Params));
  end;
end;

procedure WriteLog(Text: string);
var
  LogFile, d: string;
  f: TextFile;
begin
  d := DateToStr(now);
  delete(d, 1, 3);
  LogFile := AppPath + 'log_' + d + '.txt';
  AssignFile(f, LogFile);
  try
    if not FileExists(LogFile) then
      Rewrite(f)
    else
      Append(f);
    Writeln(f, DateTimeToStr(now) + ': ' + Text);
  finally
    CloseFile(f);
  end;
end;

function CustomSortProc(Node1, Node2: TTreeNode; iUpToThisLevel: integer): integer; stdcall;
begin
  Result := AnsiStrIComp(PChar(Node1.Text), PChar(Node2.Text));
end;

function QueryCreate: TIBCQuery;
begin
  Result := TIBCQuery.Create(nil);
  Result.Connection := FormMain.IBDatabase1;
  Result.Transaction := FormMain.IBTransaction1;
  Result.AutoCommit := False;
  Result.FetchRows := 1;
end;

function AppPath: string;
begin
  Result := ExtractFilePath(Application.ExeName);
end;

function UpperFirst(s: string): string;
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

function GetFirmCount: string;
var
  Query: TIBCQuery;
begin
  Query := QueryCreate;
  try
    Query.SQL.Text := 'select COUNT(*) from BASE';
    Query.Open;
    Result := Query.FieldByName('COUNT').AsString;
  finally
    Query.Free;
  end;
end;

function GetNameByID(table, id: string): string;
var
  Q: TIBCQuery;
begin
  Result := EmptyStr;
  if (Trim(table) = EmptyStr) or (Trim(id) = EmptyStr) then
    exit;
  Q := QueryCreate;
  try
    Q.SQL.Text := 'select NAME from ' + table + ' where id = ' + id;
    Q.Open;
    Q.FetchAll := True;
    if Q.RecordCount > 0 then
      Result := VarToStr(Q.FieldValues['NAME']);
  finally
    Q.Free;
  end;
end;

function GetIDByName(component: TsComboBoxEx): string;
var
  index: Integer;
begin
  index := GetIndexOfText(component);
  if index = -1 then
    Result := ''
  else
    Result := IntToStr(Integer(component.Items.Objects[index]));
end;

function GetIndexOfText(component: TsComboBoxEx; TextToIndex: string = ''): integer;
begin
  if TextToIndex <> EmptyStr then
    Result := component.Items.IndexOf(Trim(TextToIndex))
  else
    Result := component.Items.IndexOf(Trim(component.Text));
end;

function GetIndexOfObject(component: TsComboBoxEx; IDToIndex: integer = -1): integer;
begin
  if IDToIndex = -1 then
    Result := -1
  else
    Result := component.Items.IndexOfObject(TObject(IDToIndex));
end;

function GetObjectOfIndex(component: TsComboBoxEx): integer;
begin
  Result := -1;

  if component.ItemIndex <> -1 then
    if component.Items.Objects[component.ItemIndex] <> nil then
      Result := Integer(component.Items.Objects[component.ItemIndex])
end;

end.
