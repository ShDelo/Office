unit Report;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, sComboBox, ExtCtrls, sPanel, sButton, sCheckListBox,
  sGroupBox, sListBox, DB, IBCustomDataSet, IBQuery, sRichEdit, sGauge, ComObj,
  sLabel, DateUtils, StrUtils, Mask, sMaskEdit, sCustomComboEdit, sTooledit,
  sCheckBox, NxScrollControl, NxCustomGridControl, NxCustomGrid, NxGrid,
  IniFiles, sMemo, Registry, ComCtrls;

type
  TFormReport = class(TForm)
    panelSelectData: TsPanel;
    groupSelectData: TsGroupBox;
    editSelect9: TsComboBox;
    editSelect7: TsComboBox;
    editSelect5: TsComboBox;
    editSelect3: TsComboBox;
    editSelect1: TsComboBox;
    groupFormatData: TsGroupBox;
    editFormatDoc: TsCheckListBox;
    btnOK: TsButton;
    btnCancel: TsButton;
    ProgressBar: TsGauge;
    lblStatus: TsLabel;
    editDateAdded1: TsDateEdit;
    editDateAdded2: TsDateEdit;
    editDateEdited1: TsDateEdit;
    editDateEdited2: TsDateEdit;
    cbDateAdded: TsCheckBox;
    cbDateEdited: TsCheckBox;
    editSelect2: TsComboBox;
    editSelect4: TsComboBox;
    editSelect6: TsComboBox;
    editSelect8: TsComboBox;
    editSelect10: TsComboBox;
    groupLocation: TsGroupBox;
    cbLocGeneral: TsCheckBox;
    cbLocWord: TsCheckBox;
    cbLocExcel: TsCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure IniLoadReport;
    function IsWordInstalled: Boolean;    
    procedure IsWordAviable(cbLocGeneral,cbLocWord : TsCheckBox);
    procedure ClearEdits;
    procedure GenerateReport;
    procedure FormatReport(FilesCount: integer; RE: TsRichEdit; PB: TsGauge; lbl: TsLabel);
    procedure DeleteTempReports;
    procedure btnCancelClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure editSelect1Change(Sender: TObject);
    procedure cbLocGeneralClick(Sender: TObject);
    procedure cbLocWordClick(Sender: TObject);
    procedure cbLocExcelClick(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams); override;
    { Public declarations }
  end;

var
  FormReport: TFormReport;
  Q_GEN_BREAK : Boolean = False;

implementation

uses Main, Editor;

{$R *.dfm}

procedure TFormReport.CreateParams(var Params: TCreateParams);
begin
 inherited CreateParams(Params);
 Params.ExStyle := Params.ExStyle or WS_Ex_AppWindow;
end;

procedure TFormReport.FormCreate(Sender: TObject);
begin
 IniLoadReport;
end;

procedure TFormReport.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Q_GEN_BREAK := True;
end;

procedure TFormReport.IniLoadReport;
begin
 IniMain := TIniFile.Create(AppPath+'office.ini');
 cbLocGeneral.Checked := IniMain.ReadBool('report','cbLocGeneral',false);
 cbLocWord.Checked := IniMain.ReadBool('report','cbLocWord',false);
 cbLocExcel.Checked := IniMain.ReadBool('report','cbLocExcel',false);
 if (NOT cbLocGeneral.Checked) and (NOT cbLocWord.Checked) then
  cbLocWord.Checked := True;
 editFormatDoc.Checked[0] := IniMain.ReadBool('report','formatDoc0',false);
 editFormatDoc.Checked[1] := IniMain.ReadBool('report','formatDoc1',false);
 editFormatDoc.Checked[2] := IniMain.ReadBool('report','formatDoc2',false);
 editFormatDoc.Checked[3] := IniMain.ReadBool('report','formatDoc3',false);
 editFormatDoc.Checked[4] := IniMain.ReadBool('report','formatDoc4',false);
 editFormatDoc.Checked[5] := IniMain.ReadBool('report','formatDoc5',false);
 editFormatDoc.Checked[6] := IniMain.ReadBool('report','formatDoc6',false);
 editFormatDoc.Checked[7] := IniMain.ReadBool('report','formatDoc7',false);
 editFormatDoc.Checked[8] := IniMain.ReadBool('report','formatDoc8',false);
 editFormatDoc.Checked[9] := IniMain.ReadBool('report','formatDoc9',false); 
 IniMain.Free;
end;

procedure TFormReport.ClearEdits;
begin
 editSelect1.ItemIndex := 0; editSelect2.Clear;
 editSelect3.ItemIndex := 0; editSelect4.Clear;
 editSelect5.ItemIndex := 0; editSelect6.Clear;
 editSelect7.ItemIndex := 0; editSelect8.Clear;
 editSelect9.ItemIndex := 0; editSelect10.Clear;
 cbDateAdded.Checked := False;
 cbDateEdited.Checked := False;
 editDateAdded1.Date := Now; editDateAdded2.Date := Now;
 editDateEdited1.Date := Now; editDateEdited2.Date := Now; 
end;

function TFormReport.IsWordInstalled: Boolean;
var
 Reg: TRegistry;
begin
 Reg := TRegistry.Create;
 try Reg.RootKey := HKEY_CLASSES_ROOT;
     Result := Reg.KeyExists('Word.Application');
 finally Reg.Free;
 end;
end;

procedure TFormReport.IsWordAviable(cbLocGeneral, cbLocWord: TsCheckBox);
begin
 if NOT IsWordInstalled then begin
  MessageBox(Handle,'В операционной системе не установлена программа для просмотра файлов Microsoft Office.'+#13+'Установите программу и повторите попытку.','Предупреждение',MB_OK or MB_ICONWARNING);
  cbLocGeneral.Checked := True; cbLocGeneral.Enabled := False;
  cbLocWord.Checked := False; cbLocWord.Enabled := False;
 end else begin
  cbLocGeneral.Enabled := True; cbLocWord.Enabled := True;
 end;
end;

procedure TFormReport.cbLocGeneralClick(Sender: TObject);
begin
 if NOT cbLocGeneral.Checked and NOT cbLocWord.Checked and NOT cbLocExcel.Checked then cbLocWord.Checked := True;
end;

procedure TFormReport.cbLocWordClick(Sender: TObject);
begin
 if NOT cbLocGeneral.Checked and NOT cbLocWord.Checked and NOT cbLocExcel.Checked then begin
  cbLocGeneral.Checked := True
 end else if cbLocWord.Checked and cbLocExcel.Checked then begin
  cbLocExcel.Checked := False
 end
end;

procedure TFormReport.cbLocExcelClick(Sender: TObject);
begin
 if NOT cbLocGeneral.Checked and NOT cbLocWord.Checked and NOT cbLocExcel.Checked then begin
  cbLocGeneral.Checked := True
 end else if cbLocWord.Checked and cbLocExcel.Checked then begin
  cbLocWord.Checked := False
 end
end;

procedure TFormReport.btnCancelClick(Sender: TObject);
begin
 Q_GEN_BREAK := True;
 FormReport.Close;
end;

procedure TFormReport.btnOKClick(Sender: TObject);
begin
 GenerateReport;
end;

procedure TFormReport.editSelect1Change(Sender: TObject);

 procedure editSelectChange(select1,select2 : TsComboBox);
 var
  Q : TIBQuery;
  i : integer;
 begin
  select2.Clear;
  if (select1.ItemIndex = 1) or (select1.ItemIndex = 2) then begin
   select2.Items.Add('Да'); select2.Items.Add('Нет'); select2.ItemIndex := 0;
  end else if select1.ItemIndex = 3 then begin {СТРАНА}
   Q := TIBQuery.Create(FormReport);
   Q.Database := FormMain.IBDatabase1;
   Q.Transaction := FormMain.IBTransaction1;
   Q.Close; Q.SQL.Text := 'select * from COUNTRY order by lower(NAME)'; Q.Open; Q.FetchAll;
   for i := 1 to Q.RecordCount do
   begin
    select2.AddItem(Q.FieldValues['NAME'],Pointer(Integer(Q.FieldValues['ID'])));
    Q.Next;
   end;
   Q.Close; Q.Free;
   if select2.Items.Count > 0 then select2.ItemIndex := 0;
  end else if select1.ItemIndex = 4 then begin {ГОРОД}
   Q := TIBQuery.Create(FormReport);
   Q.Database := FormMain.IBDatabase1;
   Q.Transaction := FormMain.IBTransaction1;
   Q.Close; Q.SQL.Text := 'select * from GOROD order by lower(NAME)'; Q.Open; Q.FetchAll;
   for i := 1 to Q.RecordCount do
   begin
    select2.AddItem(Q.FieldValues['NAME'],Pointer(Integer(Q.FieldValues['ID'])));
    Q.Next;
   end;
   Q.Close; Q.Free;
   if select2.Items.Count > 0 then select2.ItemIndex := 0;
  end else if select1.ItemIndex = 5 then begin {КУРАТОР}
   for i := 0 to Main.sgCurator_tmp.RowCount - 1 do begin
    select2.AddItem(Main.sgCurator_tmp.Cells[0,i],Pointer(StrToInt(Main.sgCurator_tmp.Cells[1,i])));
   end;
   if select2.Items.Count > 0 then select2.ItemIndex := 0;
  end else if select1.ItemIndex = 6 then begin {РУБРИКА}
   for i := 0 to Main.sgRubr_tmp.RowCount - 1 do begin
    select2.AddItem(Main.sgRubr_tmp.Cells[0,i],Pointer(StrToInt(Main.sgRubr_tmp.Cells[1,i])));
   end;
   if select2.Items.Count > 0 then select2.ItemIndex := 0;
  end else if select1.ItemIndex = 7 then begin {ТИП}
   for i := 0 to Main.sgType_tmp.RowCount - 1 do begin
    select2.AddItem(Main.sgType_tmp.Cells[0,i],Pointer(StrToInt(Main.sgType_tmp.Cells[1,i])));
   end;
   if select2.Items.Count > 0 then select2.ItemIndex := 0;
  end;
 end;

begin
 if TsComboBox(Sender).Name = 'editSelect1' then
  editSelectChange(editSelect1,editSelect2);
 if TsComboBox(Sender).Name = 'editSelect3' then
  editSelectChange(editSelect3,editSelect4);
 if TsComboBox(Sender).Name = 'editSelect5' then
  editSelectChange(editSelect5,editSelect6);
 if TsComboBox(Sender).Name = 'editSelect7' then
  editSelectChange(editSelect7,editSelect8);
 if TsComboBox(Sender).Name = 'editSelect9' then
  editSelectChange(editSelect9,editSelect10);
end;

procedure TFormReport.GenerateReport;
var
 i, n, row, RE_TextLength : integer;
 Q_GEN : TIBQuery;
 REQ : string;
 d1,d2,d3,d4 : string;
 req1,req2,req3,req4,req5 : string;
 param1,param2,param3,param4,param5 : string;
 RE : TsRichEdit;
 WordApp : OleVariant;
 ExcelApp : OleVariant;
 WorkSheet : Variant;
 listEmail : TStrings; 
 t1, t2 : TDateTime;
 TmpFilesCount : integer;

 procedure AddLine(AText: string; AColor: TColor; AFontSize: integer;
  AFontName: TFontName; AFontStyle: TFontStyles);
 begin
  RE_TextLength := RE_TextLength + Length(AText) + 2;
  RE.SelStart := RE_TextLength - Length(AText) - 2;
//  RE.SelStart := Length(RE.Text);
  RE.SelAttributes.Color := AColor;
  RE.SelAttributes.Size := AFontSize; RE.SelAttributes.Name := AFontName;
  RE.SelAttributes.Style := AFontStyle; RE.Lines.Add(AText);
 end;

 procedure FormatAdres(AAdres,APhones : string);
 var
  list,list2 : TStrings;
  Rubr,phones,tmp,adres,city_str,country_str,ofType : string;
  x : integer;
 begin
  list := TStringList.Create;
  list2 := TStringList.Create;
  list.Text := AAdres;
  phones := APhones;
  for x := 0 to list.Count - 1 do
  begin
   Rubr := list[x];
   { Rubr = #CBAdres$#NUM$#OfficeType$#ZIP$#Street$#Country$#City$ }
   list2.Clear;
   while pos('$',Rubr) > 0 do
   begin
    tmp := copy(Rubr ,0 , pos('$',Rubr));
    delete(Rubr, 1 , length(tmp));
    delete(tmp,1,1); delete(tmp,length(tmp),1);
    list2.Add(tmp);
 // list2[0] = CBAdres; list2[1] = NO; list2[2] = OfficeType; list2[3] = ZIP;
 // list2[4] = Street; list2[5] = Country; list2[6] = City;
   end;

    tmp := copy(phones ,0 , pos('$',phones));
    delete(phones , 1 , length(tmp));
    tmp := copy(phones ,0 , pos('$',phones));
    delete(phones , 1 , length(tmp));
    delete(tmp,1,1); delete(tmp,length(tmp),1);
    tmp := AnsiReplaceText(tmp,Chr(13),', ');
    tmp := AnsiReplaceText(tmp,Chr(10),'');
    tmp := '    Телефон: ' + tmp;
    // в TMP сейчас хранятся все телефоны для адреса list[x]

   if list2[0] = '1' then begin
    // такая же процедура в Main.OpenTabByID и Editor.PrepareEdit и ReportSimple.GenerateReport
    city_str := list2[6]; if city_str[1] = '^' then delete(city_str,1,1);
    country_str := list2[5]; if country_str[1] = '&' then delete(country_str,1,1);
    ofType := list2[2]; if ofType[1] = '@' then delete(ofType,1,1); // НЕ ИСПОЛЬЗУЕТСЯ ТУТ
    adres := Format('    Адрес: %s - %s, %s, %s,',
     [FormEditor.GetNameByID('COUNTRY',country_str),list2[3],FormEditor.GetNameByID('GOROD',city_str),list2[4]]);
    AddLine(adres,clWindowText,10,'Times New Roman',[],);
    AddLine(tmp,clWindowText,10,'Times New Roman',[]);
   end;
  end; // for x := 0 to list.Count - 1 do
  list.Free; list2.Free;
 end;

 function FormatRubrNapr(Field,Value : string) : string;
 var
  str,tmp : string;
 begin
  Result := '';
  if Length(Value) = 0 then exit;
  str := Value;
  while pos('$',str) > 0 do
  begin
   tmp := copy(str ,0 , pos('$',str));
   delete(str, 1 , length(tmp));
   delete(tmp,1,1); delete(tmp,length(tmp),1);
   if AnsiLowerCase(Field) = 'curator' then
    if Main.sgCurator_tmp.FindText(1,tmp,[soCaseInsensitive,soExactMatch]) then
     Result := Result + ', ' + Main.sgCurator_tmp.Cells[0,Main.sgCurator_tmp.SelectedRow];
   if AnsiLowerCase(Field) = 'rubr' then
    if Main.sgRubr_tmp.FindText(1,tmp,[soCaseInsensitive,soExactMatch]) then
     Result := Result + ', ' + Main.sgRubr_tmp.Cells[0,Main.sgRubr_tmp.SelectedRow];
   if AnsiLowerCase(Field) = 'type' then
    if Main.sgType_tmp.FindText(1,tmp,[soCaseInsensitive,soExactMatch]) then
     Result := Result + ', ' + Main.sgType_tmp.Cells[0,Main.sgType_tmp.SelectedRow];
   if AnsiLowerCase(Field) = 'napr' then
    if Main.sgNapr_tmp.FindText(1,tmp,[soCaseInsensitive,soExactMatch]) then
     Result := Result + ', ' + Main.sgNapr_tmp.Cells[0,Main.sgNapr_tmp.SelectedRow];
   if Length(Result) > 0 then if Result[1] = ',' then delete(Result,1,2);
  end;
 end;

 function GetEMailList(fieldEMail : string) : TStrings;
 var
  list : TStrings;
  tmp : string;
 begin
  list := TStringList.Create;
  fieldEMail := StringReplace(fieldEMail, ' ', '', [rfReplaceAll]);
  if pos(',', fieldEMail) > 0 then begin
    fieldEMail := fieldEMail + ',';
    while pos(',',fieldEMail) > 0 do
    begin
     tmp := copy(fieldEMail ,0 , pos(',',fieldEMail));
     delete(fieldEMail, 1 , length(tmp));
     delete(tmp,length(tmp),1);
     list.Add(tmp);
    end;  
   end
  else if Length(fieldEMail) > 0 then begin
    list.Add(fieldEMail);
  end;
  Result := list;
 end;

 procedure GetRequest(select1,select2 : TsComboBox; paramNO : string;
  var Req,param : string);
 var
  ID : String;
 begin
  Req := ''; param := '';
  if select1.ItemIndex = 1 then begin // АКТИВНОСТЬ
   Req := ' (ACTIVITY = :' + paramNO + ') and';
   if select2.ItemIndex = 0 then param := '-1' else param := '0';
  end else if select1.ItemIndex = 2 then begin // АКТУАЛЬНОСТЬ
   Req := ' (RELEVANCE = :' + paramNO + ') and';
   if select2.ItemIndex = 0 then param := '-1' else param := '0';
  end else if select1.ItemIndex = 3 then begin // СТРАНЫ
   if select2.Items.Count = 0 then exit;
   Req := ' (ADRES like :' + paramNO + ') and';
   ID := IntToStr(Integer(select2.Items.Objects[select2.ItemIndex]));
   param := '%#&'+ID+'$%';
  end else if select1.ItemIndex = 4 then begin // ГОРОДА
   if select2.Items.Count = 0 then exit;
   Req := ' (ADRES like :' + paramNO + ') and';
   ID := IntToStr(Integer(select2.Items.Objects[select2.ItemIndex]));
   param := '%#^'+ID+'$%';
  end else if select1.ItemIndex = 5 then begin // КУРАТОРЫ
   if select2.Items.Count = 0 then exit;
   Req := ' (CURATOR like :' + paramNO + ') and';
   ID := IntToStr(Integer(select2.Items.Objects[select2.ItemIndex]));
   param := '%#'+ID+'$%';
  end else if select1.ItemIndex = 6 then begin // РУБРИКИ
   if select2.Items.Count = 0 then exit;
   Req := ' (RUBR like :' + paramNO + ') and';
   ID := IntToStr(Integer(select2.Items.Objects[select2.ItemIndex]));
   param := '%#'+ID+'$%';
  end else if select1.ItemIndex = 7 then begin // ТИП
   if select2.Items.Count = 0 then exit;
   Req := ' (TYPE like :' + paramNO + ') and';
   ID := IntToStr(Integer(select2.Items.Objects[select2.ItemIndex]));
   param := '%#'+ID+'$%';
  end;
 end;

begin
 if (editSelect1.ItemIndex = 0) and (editSelect3.ItemIndex = 0) and
 (editSelect5.ItemIndex = 0) and (editSelect7.ItemIndex = 0) and
 (editSelect9.ItemIndex = 0) and (NOT cbDateAdded.Checked) and
 (NOT cbDateEdited.Checked) then begin
  MessageBox(handle,'Укажите данные для генерации отчета','Предупреждение',MB_OK or MB_ICONWARNING);
  exit;
 end;
 Q_GEN := TIBQuery.Create(FormReport);
 Q_GEN.Database := FormMain.IBDatabase1;
 Q_GEN.Transaction := FormMain.IBTransaction1;
 RE := TsRichEdit.Create(FormReport);
 RE.Visible := False;
 RE.Parent := FormReport;
 RE.Clear;
 try
 Q_GEN.Close; Q_GEN.SQL.Text := ':param1,:param2,:param3,:param4,:param5)'; // tmp
 REQ := 'select * from BASE where (';
 GetRequest(editSelect1,editSelect2,'param1',req1,param1);
 GetRequest(editSelect3,editSelect4,'param2',req2,param2);
 GetRequest(editSelect5,editSelect6,'param3',req3,param3);
 GetRequest(editSelect7,editSelect8,'param4',req4,param4);
 GetRequest(editSelect9,editSelect10,'param5',req5,param5);
 if Length(req1) > 0 then begin
  REQ := REQ + req1; Q_GEN.ParamByName('param1').AsString := param1; end;
 if Length(req2) > 0 then begin
  REQ := REQ + req2; Q_GEN.ParamByName('param2').AsString := param2; end;
 if Length(req3) > 0 then begin
  REQ := REQ + req3; Q_GEN.ParamByName('param3').AsString := param3; end;
 if Length(req4) > 0 then begin
  REQ := REQ + req4; Q_GEN.ParamByName('param4').AsString := param4; end;
 if Length(req5) > 0 then begin
  REQ := REQ + req5; Q_GEN.ParamByName('param5').AsString := param5; end;
 d1 := ''''+editDateAdded1.Text+''''; d2 := ''''+editDateAdded2.Text+'''';
 d3 := ''''+editDateEdited1.Text+''''; d4 := ''''+editDateEdited2.Text+'''';
 if cbDateAdded.Checked then REQ := REQ + ' (DATE_ADDED BETWEEN DATE '+d1+' AND DATE '+d2+') and';
 if cbDateEdited.Checked then REQ := REQ + ' (DATE_EDITED BETWEEN DATE '+d3+' AND DATE '+d4+') and';
 if REQ = 'select * from BASE where (' then begin
  MessageBox(handle,'По Вашему запросу не было найдено ни одной записи','Информация',MB_OK or MB_ICONINFORMATION);
  RE.Free; Q_GEN.Close; Q_GEN.Free; exit;
 end;
 delete(REQ,length(REQ)-3,length(REQ));
 REQ := REQ + ' ) order by lower(NAME)';
 Q_GEN.SQL.Text := REQ; Q_GEN.Open; Q_GEN.FetchAll;
 if Q_GEN.RecordCount = 0 then
 begin
  MessageBox(handle,'По Вашему запросу не было найдено ни одной записи','Информация',MB_OK or MB_ICONINFORMATION);
  RE.Free; Q_GEN.Close; Q_GEN.Free; exit;
 end;
 btnOK.Enabled := False; editFormatDoc.Enabled := False;
 cbLocGeneral.Enabled := False; cbLocWord.Enabled := False;
 FormMain.DisableAllForms('FormReport');
 RE.Lines.BeginUpdate;
 FormMain.SGGeneral.BeginUpdate;
 if cbLocGeneral.Checked then FormMain.SGGeneral.ClearRows;
 lblStatus.Caption := 'Построение отчета...';
 lblStatus.Visible := True;
 ProgressBar.MinValue := 0;
 ProgressBar.MaxValue := Q_GEN.RecordCount;
 ProgressBar.Visible := True;
 t1 := now;
 if cbLocWord.Checked then begin
  AddLine('Найдено записей: '+IntToStr(Q_GEN.RecordCount),clWindowText,10,'Times New Roman',[]);
  AddLine('Условия выбора:',clWindowText,10,'Times New Roman',[]);
  if editSelect1.ItemIndex > 0 then
   AddLine(editSelect1.Text+' = '+editSelect2.Text,clWindowText,10,'Times New Roman',[]);
  if editSelect3.ItemIndex > 0 then
   AddLine(editSelect3.Text+' = '+editSelect4.Text,clWindowText,10,'Times New Roman',[]);
  if editSelect5.ItemIndex > 0 then
   AddLine(editSelect5.Text+' = '+editSelect6.Text,clWindowText,10,'Times New Roman',[]);
  if editSelect7.ItemIndex > 0 then
   AddLine(editSelect7.Text+' = '+editSelect8.Text,clWindowText,10,'Times New Roman',[]);
  if editSelect9.ItemIndex > 0 then
   AddLine(editSelect9.Text+' = '+editSelect10.Text,clWindowText,10,'Times New Roman',[]);
  if cbDateAdded.Checked then
   AddLine(cbDateAdded.Caption+' = '+editDateAdded1.Text+' - '+editDateAdded2.Text,clWindowText,10,'Times New Roman',[]);
  if cbDateEdited.Checked then
   AddLine(cbDateEdited.Caption+' = '+editDateEdited1.Text+' - '+editDateEdited2.Text,clWindowText,10,'Times New Roman',[]);
  AddLine('',clWindowText,10,'Times New Roman',[]);
 end;
 if cbLocExcel.Checked then begin // create Excel doc
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Workbooks.Add();
  WorkSheet := ExcelApp.Workbooks[1].WorkSheets[1];
  row := 1;
 end;
 TmpFilesCount := 0;
 for i := 1 to Q_GEN.RecordCount do
 begin
  if Q_GEN_BREAK then Break;
  ProgressBar.Progress := i;
  
  if RE.Lines.Count >= 10000 then begin
   Inc(TmpFilesCount,1);
   RE.Lines.SaveToFile(AppPath+'\report_'+IntToStr(TmpFilesCount)+'.tmp');
   RE.Clear;
  end;

  if cbLocGeneral.Checked then begin
   FormMain.SGAddRow(FormMain.SGGeneral,Q_GEN.FieldByName('ACTIVITY').AsInteger,Q_GEN.FieldByName('RELEVANCE').AsInteger,
   Q_GEN.FieldValues['NAME'],Q_GEN.FieldValues['CURATOR'],Q_GEN.FieldValues['DATE_ADDED'],Q_GEN.FieldValues['DATE_EDITED'],
   Q_GEN.FieldValues['WEB'],Q_GEN.FieldValues['EMAIL'],Q_GEN.FieldValues['TYPE'],Q_GEN.FieldValues['ID'],
   Q_GEN.FieldValues['FIO'],Q_GEN.FieldValues['RUBR']);
  end;

  if cbLocWord.Checked then begin
   AddLine('Фирма: '+Q_GEN.FieldValues['NAME'],clWindowText,12,'Times New Roman',[fsBold]);
   if editFormatDoc.Checked[0] then
    FormatAdres(Q_GEN.FieldValues['ADRES'],Q_GEN.FieldValues['PHONES']);
   if editFormatDoc.Checked[1] then begin
    if Q_GEN.FieldByName('ACTIVITY').AsInteger = -1 then
     AddLine('    Активность: да',clWindowText,10,'Times New Roman',[])
    else AddLine('    Активность: нет',clWindowText,10,'Times New Roman',[]);
   end;
   if editFormatDoc.Checked[2] then begin
    if Q_GEN.FieldByName('RELEVANCE').AsInteger = -1 then
     AddLine('    Актуальность: да',clWindowText,10,'Times New Roman',[])
    else AddLine('    Актуальность: нет',clWindowText,10,'Times New Roman',[]);
   end;
   if editFormatDoc.Checked[3] then
    AddLine('    Куратор: '+FormatRubrNapr('CURATOR',Q_GEN.FieldValues['CURATOR']),clWindowText,10,'Times New Roman',[]);
   if editFormatDoc.Checked[4] then
    AddLine('    Рубрика: '+FormatRubrNapr('RUBR',Q_GEN.FieldValues['RUBR']),clWindowText,10,'Times New Roman',[]);
   if editFormatDoc.Checked[5] then
    AddLine('    Тип: '+FormatRubrNapr('TYPE',Q_GEN.FieldValues['TYPE']),clWindowText,10,'Times New Roman',[]);
   if editFormatDoc.Checked[6] then
    AddLine('    Деятельность: '+FormatRubrNapr('NAPR',Q_GEN.FieldValues['NAPRAVLENIE']),clWindowText,10,'Times New Roman',[]);
   if editFormatDoc.Checked[7] then
    AddLine('    Веб: '+Q_GEN.FieldValues['WEB'],clWindowText,10,'Times New Roman',[]);
   if editFormatDoc.Checked[8] then
    AddLine('    Мейл: '+Q_GEN.FieldValues['EMAIL'],clWindowText,10,'Times New Roman',[]);
   if editFormatDoc.Checked[9] then
    AddLine('    Представитель: '+Q_GEN.FieldValues['FIO'],clWindowText,10,'Times New Roman',[]);
   AddLine('',clWindowText,10,'Times New Roman',[]);
  end;

  if cbLocExcel.Checked then begin
//    listEmail := TStringList.Create;
    listEmail := GetEMailList(Q_GEN.FieldValues['EMAIL']);
    if Length(listEmail.GetText) > 0 then begin
     for n := 0 to listEmail.Count - 1 do
     begin
      WorkSheet.Cells[row,1] := Q_GEN.FieldValues['NAME'];
      WorkSheet.Cells[row,2] := listEmail[n];
      Inc(row);
     end;
    end;
  end;

  Q_GEN.Next;
  Application.ProcessMessages;
 end; // for i := 1 to Q_GEN.RecordCount
 t2 := now;
 FormMain.sStatusBar1.Panels[2].Text := 'Отчет создан за '+IntToStr(SecondsBetween(t1, t2))+' сек';
 RE.Lines.EndUpdate;
 FormMain.SGGeneral.EndUpdate;
 if (Length(RE.Text) > 0) AND (NOT Q_GEN_BREAK) then
 begin
  try
   WordApp := GetActiveOleObject('Word.Application');
   WordApp.Documents.Close(0);
   WordApp.Quit;
  except end;
  if TmpFilesCount > 0 then begin
   // сохраняем посл часть отчета потому-что было разделение и форматируем
   Inc(TmpFilesCount,1);
   RE.Lines.SaveToFile(AppPath+'\report_'+IntToStr(TmpFilesCount)+'.tmp');
   FormatReport(TmpFilesCount,RE,ProgressBar,lblStatus);
  end;
  RE.Lines.SaveToFile(AppPath + '\tmp.rtf');
  WordApp := CreateOLEObject('Word.Application');
  WordApp.Documents.Add(AppPath + '\tmp.rtf');
  WordApp.Visible := True;
  {ShellExecute(handle,'open',PChar(AppPath + '\tmp.rtf'),nil,nil,SW_SHOWNORMAL);}
 end;
 if cbLocExcel.Checked then begin // show Excel doc
  ExcelApp.Visible := True;
  VarClear(ExcelApp);
 end; 
 if cbLocGeneral.Checked then FormMain.sPageControl1.ActivePageIndex := 0;
 except on E:Exception do begin
  FormMain.WriteLog('TFormReport.GenerateReport'+#13+'Произошел сбой при генерации отчета'+#13+E.Message);
  MessageBox(Handle,PChar('Произошел сбой при генерации отчета'+#13+E.Message),
   'Ошибка',MB_OK or MB_ICONERROR); end;
 end;
 DeleteTempReports;
 lblStatus.Visible := False;
 ProgressBar.Visible := False;
 btnOK.Enabled := True; editFormatDoc.Enabled := True;
 cbLocGeneral.Enabled := True; cbLocWord.Enabled := True;
 FormMain.EnableAllForms(''); 
 RE.Free; Q_GEN.Close; Q_GEN.Free;
 FormMain.IBDatabase1.Close;
 FormReport.Close;
end;

procedure TFormReport.FormatReport(FilesCount: integer; RE: TsRichEdit;
 PB: TsGauge; lbl: TsLabel);
var
 CurFileNO: integer;
 FileName, UnitedText, Tmp: string;
 TextStream: TStringStream;
begin
 TextStream := TStringStream.Create('');
 UnitedText := '';
 RE.MaxLength := $7FFFFFF0;
 lbl.Caption := 'Форматирование отчета...';
 PB.Progress := 0;
 PB.MaxValue := FilesCount;
 Application.ProcessMessages;
 try
  for CurFileNO := 1 to FilesCount do
  begin
   PB.Progress := CurFileNO;
   FileName := 'report_'+IntTOStr(CurFileNO)+'.tmp';
   RE.Lines.LoadFromFile(FileName);
   RE.Lines.SaveToStream(TextStream);
   Tmp := TextStream.DataString;
   TextStream.Position := 0;   
   if CurFileNO = 1 then Tmp := copy(Tmp, 0, length(Tmp) - 5)
   else if CurFileNO = FilesCount then Tmp := copy(Tmp, 2, length(Tmp))
   else Tmp := copy(Tmp, 2, length(Tmp) - 5);
   UnitedText := UnitedText + Tmp;
   Application.ProcessMessages;
  end;
  lbl.Caption := 'Загрузка отчета...';
  Application.ProcessMessages;
  TextStream.WriteString(UnitedText);
  TextStream.Position := 0;
  RE.Lines.LoadFromStream(TextStream);
//  RE.Text := UnitedText; // как вариант вместо TextStream.WriteString(UnitedText)
 finally
  TextStream.Free;
 end;
end;

procedure TFormReport.DeleteTempReports;

 function FilesInDir(Dir: String): Integer;
 var
  sr: TSearchRec;
 begin
  Result := 0;
  if FindFirst(Dir + '\*', faAnyFile - faDirectory, sr) <> 0 then
  begin FindClose(sr); exit; end;
  repeat inc(Result); until (FindNext(sr) <> 0);
  FindClose(sr);
 end;

var
 i : integer;
 FileName : string;
begin
 try for i := 1 to FilesInDir(AppPath) do begin
  FileName := AppPath+'\report_'+IntToStr(i)+'.tmp';
  if FileExists(FileName) then DeleteFile(FileName);
 end;
 except on E:Exception do begin
  FormMain.WriteLog('TFormReport.GenerateReport'+#13+'Ошибка удаления временных файлов'+#13+E.Message);
  MessageBox(Handle,PChar('Ошибка удаления временных файлов'+#13+E.Message),
   'Ошибка',MB_OK or MB_ICONERROR); end;
 end;
end;

end.
