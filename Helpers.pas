unit Helpers;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, ComCtrls, Forms, sComboBoxes,
  NxGrid, IBC;

type
  TsComboBoxEx_Helper = class helper for TsComboBoxEx
    function GetIndexOfObject(ObjectToIndex: integer): integer;
    function SetIndexOfObject(ObjectToIndex: integer): integer;
    function GetObjectOfIndex(AIndex: integer = -1): integer;
    function GetID: integer;
    function GetIndexOfText(TextToIndex: string = ''; DoTrimText: boolean = true): integer;
  end;

procedure debug(Text: string; Params: array of TVarRec);
procedure WriteLog(Text: string);
function CustomSortProc(Node1, Node2: TTreeNode; iUpToThisLevel: integer): integer; stdcall;
function QueryCreate: TIBCQuery;
function AppPath: string;
function UpperFirst(s: string): string;
function GetFirmCount: string;
function GetNameByID(table, id: string): string;

const
  bDebug: Boolean = True;

implementation

uses
  Main;

{ TsComboBoxEx_Helper }

function TsComboBoxEx_Helper.GetIndexOfObject(ObjectToIndex: integer): integer;
begin
  if ObjectToIndex >= 0 then
    Result := self.Items.IndexOfObject(TObject(ObjectToIndex))
  else
    Result := -1
end;

function TsComboBoxEx_Helper.SetIndexOfObject(ObjectToIndex: integer): integer;
var
  IndexOfObject: integer;
begin
  if ObjectToIndex >= 0 then
  begin
    IndexOfObject := self.GetIndexOfObject(ObjectToIndex);
    Result := IndexOfObject;
    if self.ItemIndex <> IndexOfObject then
    begin
      self.ItemIndex := IndexOfObject;
      self.Tag := IndexOfObject;
    end;
  end
  else
  begin
    Result := -1;
    if self.ItemIndex <> -1 then
    begin
      self.ItemIndex := -1;
      self.Tag := -1;
      self.Text := '';
    end;
  end;

end;

function TsComboBoxEx_Helper.GetID: integer;
var
  IndexOfText: integer;
begin
  Result := -1;

  if self.ItemIndex <> -1 then
  begin
    if self.Items.Objects[self.ItemIndex] <> nil then
      Result := Integer(self.Items.Objects[self.ItemIndex])
  end
  else
  begin
    IndexOfText := self.GetIndexOfText;
    if IndexOfText <> -1 then
      if self.Items.Objects[IndexOfText] <> nil then
        Result := Integer(self.Items.Objects[IndexOfText]);
  end;
end;

function TsComboBoxEx_Helper.GetObjectOfIndex(AIndex: integer): integer;
begin
  Result := -1;

  if AIndex = -1 then
  begin
    if self.ItemIndex <> -1 then
      if self.Items.Objects[self.ItemIndex] <> nil then
        Result := Integer(self.Items.Objects[self.ItemIndex])
  end
  else
  begin
    if self.Items.Objects[AIndex] <> nil then
      Result := Integer(self.Items.Objects[AIndex])
  end;
end;

function TsComboBoxEx_Helper.GetIndexOfText(TextToIndex: string = ''; DoTrimText: boolean = true): integer;
begin
  if TextToIndex <> EmptyStr then
  begin
    if DoTrimText then
      TextToIndex := Trim(TextToIndex);
    Result := self.Items.IndexOf(TextToIndex);
  end
  else
  begin
    if DoTrimText then
      Result := self.Items.IndexOf(Trim(self.Text))
    else
      Result := self.Items.IndexOf(self.Text);
  end;
end;

{ Helpers }

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

end.
