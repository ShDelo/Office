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

  TDirectoryData = packed record
    Name1: string[255];
    Name2: string[255];
    ID_Region: integer;
    constructor Create(Name1, Name2: string; ID_Region: integer);
  end;

  TDirectoryContainer = class
    SG_Main: Pointer; // temp grid e.g main.sgCurator_tmp
    SG_Directory: Pointer; // SG control on FormDirectory
    Edit_Editor: Pointer; // edit on FormEditor
    SG_Editor: Pointer; // SG on FormEditor
    Edit_DirectoryQuery: Pointer; // edit on DirectoryQuery
    IsAdresEdits: Boolean;
    constructor Create(DirCode: integer);
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
  ID_UNKNOWN = -1;

  DIR_CODE_TOTAL = 8; // 0-based total number of codes
  DIR_CODE_TO_TABLE: array [0 .. DIR_CODE_TOTAL] of string = ('curator', 'rubrikator', 'firmtype', 'napravlenie', 'officetype', 'country',
    'region', 'city', 'phonetype');
  DIR_CODE_CURATOR = 0;
  DIR_CODE_RUBRIKA = 1;
  DIR_CODE_FIRMTYPE = 2;
  DIR_CODE_NAPRAVLENIE = 3;
  DIR_CODE_OFFICETYPE = 4;
  DIR_CODE_COUNTRY = 5;
  DIR_CODE_REGION = 6;
  DIR_CODE_CITY = 7;
  DIR_CODE_PHONETYPE = 8;

  bDebug: Boolean = True;

implementation

uses
  Main, Editor, Directory, DirectoryQuery;

{ TsComboBoxEx_Helper }

function TsComboBoxEx_Helper.GetIndexOfObject(ObjectToIndex: integer): integer;
begin
  if ObjectToIndex >= 0 then
    Result := self.Items.IndexOfObject(TObject(ObjectToIndex))
  else
    Result := ID_UNKNOWN;
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
    Result := ID_UNKNOWN;
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
  Result := ID_UNKNOWN;;

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

function TsComboBoxEx_Helper.GetObjectOfIndex(AIndex: integer = -1): integer;
begin
  Result := ID_UNKNOWN;;

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

{ TDirectoryData }

constructor TDirectoryData.Create(Name1, Name2: string; ID_Region: integer);
begin
  self.Name1 := Name1;
  self.Name2 := Name2;
  self.ID_Region := ID_Region;
end;

{ TDirectoryContainer }

constructor TDirectoryContainer.Create(DirCode: integer);
begin
  if not(DirCode in [0 .. DIR_CODE_TOTAL]) then
  begin
    self.SG_Main := nil;
    self.SG_Directory := nil;
    self.Edit_Editor := nil;
    self.SG_Editor := nil;
    self.Edit_DirectoryQuery := nil;
    self.IsAdresEdits := False;
  end;

  case DirCode of
    DIR_CODE_CURATOR:
      begin
        self.SG_Main := main.sgCurator_tmp;
        self.SG_Directory := FormDirectory.SGCurator;
        self.Edit_Editor := FormEditor.EditCurator;
        self.SG_Editor := FormEditor.SGCurator;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := False;
      end;
    DIR_CODE_RUBRIKA:
      begin
        self.SG_Main := main.sgRubr_tmp;
        self.SG_Directory := FormDirectory.SGRubr;
        self.Edit_Editor := FormEditor.EditRubr;
        self.SG_Editor := FormEditor.SGRubr;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := False;
      end;
    DIR_CODE_FIRMTYPE:
      begin
        self.SG_Main := main.sgFirmType_tmp;
        self.SG_Directory := FormDirectory.SGFirmType;
        self.Edit_Editor := FormEditor.EditFirmType;
        self.SG_Editor := FormEditor.SGFirmType;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := False;
      end;
    DIR_CODE_NAPRAVLENIE:
      begin
        self.SG_Main := main.sgNapr_tmp;
        self.SG_Directory := FormDirectory.SGNapr;
        self.Edit_Editor := FormEditor.EditNapravlenie;
        self.SG_Editor := FormEditor.SGNapravlenie;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := False;
      end;
    DIR_CODE_OFFICETYPE:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGOfficeType;
        self.Edit_Editor := FormEditor.EditOfficeType1;
        self.SG_Editor := nil;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := True;
      end;
    DIR_CODE_COUNTRY:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGCountry;
        self.Edit_Editor := FormEditor.EditCountry1;
        self.SG_Editor := nil;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := True;
      end;
    DIR_CODE_REGION:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGRegion;
        self.Edit_Editor := FormEditor.EditRegion1;
        self.SG_Editor := nil;
        self.Edit_DirectoryQuery := FormDirectoryQuery.Edit3;
        self.IsAdresEdits := True;
      end;
    DIR_CODE_CITY:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGCity;
        self.Edit_Editor := FormEditor.EditCity1;
        self.SG_Editor := nil;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := True;
      end;
    DIR_CODE_PHONETYPE:
      begin
        self.SG_Main := nil;
        self.SG_Directory := FormDirectory.SGPhoneType;
        self.Edit_Editor := FormEditor.EditPhoneType1;
        self.SG_Editor := nil;
        self.Edit_DirectoryQuery := nil;
        self.IsAdresEdits := True;
      end;
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
