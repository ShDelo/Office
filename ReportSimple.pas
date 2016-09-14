unit ReportSimple;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, sButton, sCheckBox, sGroupBox, NxColumns,
  NxColumnClasses, NxScrollControl, NxCustomGridControl, NxCustomGrid,
  NxGrid, sEdit, sComboBox, ExtCtrls, sPanel, IniFiles, IBQuery, sLabel,
  Mask, sMaskEdit, sCustomComboEdit, sTooledit, sRichEdit, StrUtils, ComObj,
  sGauge, DateUtils;

type
  TFormReportSimple = class(TForm)
    panelSelectData: TsPanel;
    groupLocation: TsGroupBox;
    cbLocGeneral: TsCheckBox;
    cbLocWord: TsCheckBox;
    groupSelectData: TsGroupBox;
    editFilter: TsComboBox;
    btnOK: TsButton;
    btnCancel: TsButton;
    panelDates: TsPanel;
    editDate1: TsDateEdit;
    editDate2: TsDateEdit;
    lblDate: TsLabel;
    panelFilters: TsPanel;
    editFilterData: TsComboBox;
    ProgressBar: TsGauge;
    lblStatus: TsLabel;
    procedure FormCreate(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure ClearEdits;
    procedure IniLoadReportSimple;
    procedure cbLocGeneralClick(Sender: TObject);
    procedure GenerateReport;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure editFilterSelect(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams); override;
    { Public declarations }
  end;

var
  FormReportSimple: TFormReportSimple;
  Q_GEN_SIMPLE_BREAK : Boolean = False;

implementation

uses Main, Editor, Report;

{$R *.dfm}

procedure TFormReportSimple.CreateParams(var Params: TCreateParams);
begin
 inherited CreateParams(Params);
 Params.ExStyle := Params.ExStyle or WS_Ex_AppWindow;
end;

procedure TFormReportSimple.FormCreate(Sender: TObject);
begin
 IniLoadReportSimple;
end;

procedure TFormReportSimple.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
 Q_GEN_SIMPLE_BREAK := True;
end;

procedure TFormReportSimple.btnOKClick(Sender: TObject);
begin
 GenerateReport;
end;

procedure TFormReportSimple.btnCancelClick(Sender: TObject);
begin
 Q_GEN_SIMPLE_BREAK := True;
 FormReportSimple.Close;
end;

procedure TFormReportSimple.ClearEdits;
begin
 editFilter.ItemIndex := -1;
 editFilterData.Clear;
 panelFilters.Visible := True; panelDates.Visible := False;
end;

procedure TFormReportSimple.IniLoadReportSimple;
begin
 IniMain := TIniFile.Create(AppPath+'office.ini');
 cbLocGeneral.Checked := IniMain.ReadBool('reportsimple','cbLocGeneral',false);
 cbLocWord.Checked := IniMain.ReadBool('reportsimple','cbLocWord',false);
 if (NOT cbLocGeneral.Checked) and (NOT cbLocWord.Checked) then
  cbLocWord.Checked := True;
 IniMain.Free;
end;

procedure TFormReportSimple.cbLocGeneralClick(Sender: TObject);
begin
 if NOT cbLocGeneral.Checked and NOT cbLocWord.Checked then cbLocGeneral.Checked := True;
end;

procedure TFormReportSimple.editFilterSelect(Sender: TObject);
var
 Q : TIBQuery;
 i,n : integer;
 memo : TStrings;
 Phones,Email,Web,tmp : string;
begin
 editFilterData.Items.BeginUpdate;
 editFilterData.Clear;
 editDate1.Date := Now; editDate2.Date := Now;
 case editFilter.ItemIndex of
  0..9 : begin panelFilters.Visible := True; panelDates.Visible := False; end;
  10..11 : begin panelFilters.Visible := False; panelDates.Visible := True; end;
 end;
 Q := TIBQuery.Create(FormReportSimple);
 Q.Database := FormMain.IBDatabase1;
 Q.Transaction := FormMain.IBTransaction1;
 if editFilter.ItemIndex = 0 then begin // РУБРИКИ
  for i := 0 to Main.sgRubr_tmp.RowCount - 1 do begin
   editFilterData.AddItem(Main.sgRubr_tmp.Cells[0,i],Pointer(StrToInt(Main.sgRubr_tmp.Cells[1,i])));
  end;
 end else if editFilter.ItemIndex = 1 then begin // ГОРОДА
  Q.Close;
  Q.SQL.Text := 'select * from GOROD';
  Q.Open; Q.FetchAll;
  for i := 1 to Q.RecordCount do begin
   editFilterData.AddItem(Q.FieldValues['NAME'],Pointer(Integer(Q.FieldValues['ID'])));
   Q.Next;
  end;
 end else if editFilter.ItemIndex = 2 then begin // СТРАНЫ
  Q.Close;
  Q.SQL.Text := 'select * from COUNTRY';
  Q.Open; Q.FetchAll;
  for i := 1 to Q.RecordCount do begin
   editFilterData.AddItem(Q.FieldValues['NAME'],Pointer(Integer(Q.FieldValues['ID'])));
   Q.Next;
  end;
 end else if editFilter.ItemIndex = 3 then begin // КУРАТОРЫ
  for i := 0 to Main.sgCurator_tmp.RowCount - 1 do begin
   editFilterData.AddItem(Main.sgCurator_tmp.Cells[0,i],Pointer(StrToInt(Main.sgCurator_tmp.Cells[1,i])));
  end;
 end else if editFilter.ItemIndex = 4 then begin // ТИПЫ
  for i := 0 to Main.sgType_tmp.RowCount - 1 do begin
   editFilterData.AddItem(Main.sgType_tmp.Cells[0,i],Pointer(StrToInt(Main.sgType_tmp.Cells[1,i])));
  end;
 end else if editFilter.ItemIndex = 5 then begin // АКТИВНОСТЬ
  editFilterData.AddItem('Нет',Pointer(0));
  editFilterData.AddItem('Да',Pointer(1));
 end else if editFilter.ItemIndex = 6 then begin // АКТУАЛЬНОСТЬ
  editFilterData.AddItem('Нет',Pointer(0));
  editFilterData.AddItem('Да',Pointer(1));
 end else if editFilter.ItemIndex = 7 then begin // ТЕЛЕФОНЫ
  Q.Close; Q.SQL.Text := 'select PHONES from BASE'; Q.Open; Q.FetchAll;
  memo := TStringList.Create;
  for i := 1 to Q.RecordCount do begin
   PHONES := Q.Fields[0].AsString;
   while pos('$',PHONES) > 0 do begin
    memo.Clear;
    tmp := copy(PHONES ,0 , pos('$',PHONES));
    delete(PHONES , 1 , length(tmp));
    tmp := copy(PHONES ,0 , pos('$',PHONES));
    delete(PHONES , 1 , length(tmp));
    delete(tmp,1,1); delete(tmp,length(tmp),1);
    if trim(tmp) = '' then Continue;
    memo.Text := tmp;
    for n := 0 to memo.Count - 1 do begin
     tmp := memo[n];
     if pos(')',tmp) > 0 then delete(tmp,1,pos(')',tmp)); tmp := Trim(tmp);
      editFilterData.AddItem(tmp,Pointer(0));
    end;
   end;
   Q.Next;
  end;
  memo.Free;
 end else if editFilter.ItemIndex = 8 then begin // МЕЙЛ
  Q.Close; Q.SQL.Text := 'select EMAIL from BASE'; Q.Open; Q.FetchAll;
  for i := 1 to Q.RecordCount do begin
   EMAIL := Q.Fields[0].AsString;
   if Length(Trim(EMAIL)) > 0 then begin
    if EMAIL[Length(EMAIL)] <> ',' then EMAIL := EMAIL + ',';
    while pos(',',EMAIL) > 0 do begin
     tmp := copy(EMAIL ,0 , pos(',',EMAIL));
     delete(EMAIL, 1 , length(tmp));
     tmp := Trim(tmp); delete(tmp,length(tmp),1);
      editFilterData.AddItem(tmp,Pointer(0));
    end;
   end;
   Q.Next;
  end;
 end else if editFilter.ItemIndex = 9 then begin // САЙТ
  Q.Close; Q.SQL.Text := 'select WEB from BASE'; Q.Open; Q.FetchAll;
  for i := 1 to Q.RecordCount do begin
   WEB := Q.Fields[0].AsString;
   if Length(Trim(WEB)) > 0 then begin
    if WEB[Length(WEB)] <> ',' then WEB := WEB + ',';
    while pos(',',WEB) > 0 do begin
     tmp := copy(WEB ,0 , pos(',',WEB));
     delete(WEB, 1 , length(tmp));
     tmp := Trim(tmp); delete(tmp,length(tmp),1);
      editFilterData.AddItem(tmp,Pointer(0));
    end;
   end;
   Q.Next;
  end;
 end else if editFilter.ItemIndex = 10 then begin
  lblDate.Caption := 'Дата добавления (период):'
 end else if editFilter.ItemIndex = 11 then begin
  lblDate.Caption := 'Дата изменения (период):'
 end;
 Q.Close; Q.Free; FormMain.IBDatabase1.Close;
 editFilterData.Items.EndUpdate;
 if editFilterData.Items.Count > 0 then editFilterData.ItemIndex := 0;
 if editFilter.ItemIndex in [0..9] then begin
  editFilterData.SetFocus;
  editFilterData.DroppedDown := True;  
 end;
end;

procedure TFormReportSimple.GenerateReport;
var
 i,RE_TextLength,index : integer;
 Q_GEN : TIBQuery;
 REQ,param,d1,d2,ID_tmp,str1,str2,finalStr : string;
 RE : TsRichEdit;
 WordApp : OleVariant;
 t1, t2 : TDateTime;
 TmpFilesCount : integer;

 procedure AddLine(AText: string; AColor: TColor; AFontSize: integer;
  AFontName: TFontName; AFontStyle: TFontStyles);
 begin
  RE_TextLength := RE_TextLength + Length(AText) + 2;
  RE.SelStart := RE_TextLength - Length(AText) - 2;
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
    tmp := StringReplace(tmp,'(','',[rfReplaceAll]);
    tmp := StringReplace(tmp,')','',[rfReplaceAll]);
    // в TMP сейчас хранятся все телефоны для адреса list[x] а формате Мемо

   if list2[0] = '1' then begin
    // такая же процедура в Main.OpenTabByID и Editor.PrepareEdit и Report.GenerateReport
    city_str := list2[6]; if city_str[1] = '^' then delete(city_str,1,1);
    country_str := list2[5]; if country_str[1] = '&' then delete(country_str,1,1);
    ofType := list2[2]; if ofType[1] = '@' then delete(ofType,1,1); // НЕ ИСПОЛЬЗУЕТСЯ ТУТ
    adres := '';
    if Trim(list2[3]) <> '' then adres := list2[3] + ', '; // ZIP
    if Trim(city_str) <> '' then adres := adres + FormEditor.GetNameByID('GOROD',city_str) + ','; // CITY
    if Trim(adres) <> '' then AddLine(adres,clWindowText,10,'Times New Roman',[],);
    if Trim(list2[4]) <> '' then AddLine(list2[4],clWindowText,10,'Times New Roman',[],);
{    adres := adres + Format('%s, %s, %s,',
     [FormEditor.GetNameByID('COUNTRY',country_str),FormEditor.GetNameByID('GOROD',city_str),list2[4]]);
    if Trim(adres) <> ', , ,' then AddLine(adres,clWindowText,10,'Times New Roman',[],);}
    if Trim(tmp) <> '' then AddLine(tmp,clWindowText,10,'Times New Roman',[fsItalic]);
//    AddLine('',clWindowText,10,'Times New Roman',[]);
   end;
  end; // for x := 0 to list.Count - 1 do
  list.Free; list2.Free;
 end;

 function FormatNapr(Field,Value : string) : string;
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
   if AnsiLowerCase(Field) = 'napr' then
    if Main.sgNapr_tmp.FindText(1,tmp,[soCaseInsensitive,soExactMatch]) then
     Result := Result + #13 + Main.sgNapr_tmp.Cells[0,Main.sgNapr_tmp.SelectedRow];
   if Length(Result) > 0 then if Result[1] = #13 then delete(Result,1,1);
  end;
 end;

 function FormatEMailWeb(Value : string) : string;
 var
  str,tmp : string;
 begin
  Result := '';
  if Length(Value) = 0 then exit;
  str := Value;
  if str[Length(str)] <> ',' then str := str + ',';
  while pos(',',str) > 0 do
  begin
   tmp := copy(str ,0 , pos(',',str));
   delete(str, 1 , length(tmp));
   tmp := Trim(tmp);
   delete(tmp,length(tmp),1);
   Result := Result + #13 + tmp;
   if Length(Result) > 0 then if Result[1] = #13 then delete(Result,1,1);
  end;
 end;

 procedure GetRequest(var Req,param : string);
 var
  ID : String;
  index : integer;
 begin
  Req := ''; param := '';
  index := editFilterData.Items.IndexOf(Trim(editFilterData.Text));
  if index <> -1 then ID := IntToStr(Integer(editFilterData.Items.Objects[index]))
   else ID := '-1';
  if editFilter.ItemIndex = 0 then begin // РУБРИКА
   Req := 'select * from BASE where RUBR like :param and ACTIVITY = -1 order by lower(NAME)';
   param := '%#'+ID+'$%';
  end else if editFilter.ItemIndex = 1 then begin // ГОРОД
   Req := 'select * from BASE where ADRES like :param order by lower(NAME)';
   param := '%#^'+ID+'$%';
  end else if editFilter.ItemIndex = 2 then begin // СТРАНА
   Req := 'select * from BASE where ADRES like :param order by lower(NAME)';
   param := '%#&'+ID+'$%';
  end else if editFilter.ItemIndex = 3 then begin // КУРАТОР
   Req := 'select * from BASE where CURATOR like :param order by lower(NAME)';
   param := '%#'+ID+'$%';
  end else if editFilter.ItemIndex = 4 then begin // ТИП
   Req := 'select * from BASE where TYPE like :param order by lower(NAME)';
   param := '%#'+ID+'$%';
  end else if editFilter.ItemIndex = 5 then begin // АКТИВНОСТЬ
   Req := 'select * from BASE where ACTIVITY like :param order by lower(NAME)';
   if AnsiLowerCase(Trim(editFilterData.Text)) = 'да' then ID := '-1' else
   if AnsiLowerCase(Trim(editFilterData.Text)) = 'нет' then ID := '0' else
   ID := '-2';
   param := ID;
  end else if editFilter.ItemIndex = 6 then begin // АКТУАЛЬНОСТЬ
   Req := 'select * from BASE where RELEVANCE like :param order by lower(NAME)';
   if AnsiLowerCase(Trim(editFilterData.Text)) = 'да' then ID := '-1' else
   if AnsiLowerCase(Trim(editFilterData.Text)) = 'нет' then ID := '0' else
   ID := '-2';
   param := ID;
  end else if editFilter.ItemIndex = 7 then begin // ТЕЛЕФОН
   Req := 'select * from BASE where PHONES like :param order by lower(NAME)';
   param := '%'+ AnsiLowerCase(Trim(editFilterData.Text)) +'%';
  end else if editFilter.ItemIndex = 8 then begin // МЕЙЛ
   Req := 'select * from BASE where lower(EMAIL) like :param order by lower(NAME)';
   param := '%'+ AnsiLowerCase(Trim(editFilterData.Text)) +'%';
  end else if editFilter.ItemIndex = 9 then begin // САЙТ
   Req := 'select * from BASE where lower(WEB) like :param order by lower(NAME)';
   param := '%'+ AnsiLowerCase(Trim(editFilterData.Text)) +'%';
  end else if editFilter.ItemIndex = 10 then begin // ДАТА ДОБАВЛЕНИЯ
   d1 := ''''+editDate1.Text+''''; d2 := ''''+editDate2.Text+'''';
   Req := 'select * from BASE where (DATE_ADDED BETWEEN DATE '+d1+' AND DATE '+d2+') order by lower(NAME)';
  end else if editFilter.ItemIndex = 11 then begin // ДАТА ИЗМЕНЕНИЯ
   d1 := ''''+editDate1.Text+''''; d2 := ''''+editDate2.Text+'''';
   Req := 'select * from BASE where (DATE_EDITED BETWEEN DATE '+d1+' AND DATE '+d2+') order by lower(NAME)';
  end;
 end;

begin
 if (editFilter.ItemIndex = -1) or
 ((editFilter.ItemIndex in [0..9]) and (Trim(editFilterData.Text) = '')) then begin
  MessageBox(handle,'Укажите данные для генерации отчета','Предупреждение',MB_OK or MB_ICONWARNING);
  exit;
 end;
 index := editFilterData.Items.IndexOf(Trim(editFilterData.Text));
 Q_GEN := TIBQuery.Create(FormReportSimple);
 Q_GEN.Database := FormMain.IBDatabase1;
 Q_GEN.Transaction := FormMain.IBTransaction1;
 RE := TsRichEdit.Create(FormReportSimple);
 RE.Visible := False;
 RE.Parent := FormReportSimple;
 RE.Clear;
 try
 Q_GEN.Close; Q_GEN.SQL.Text := ':param'; // tmp
 GetRequest(REQ,param);
 if Length(REQ) > 0 then begin
  Q_GEN.ParamByName('param').AsString := param; end;
 if REQ = '' then begin
  MessageBox(handle,'По Вашему запросу не было найдено ни одной записи','Информация',MB_OK or MB_ICONINFORMATION);
  RE.Free; Q_GEN.Close; Q_GEN.Free; exit;
 end;
 Q_GEN.SQL.Text := REQ; Q_GEN.Open; Q_GEN.FetchAll;
 if Q_GEN.RecordCount = 0 then
 begin
  MessageBox(handle,'По Вашему запросу не было найдено ни одной записи','Информация',MB_OK or MB_ICONINFORMATION);
  RE.Free; Q_GEN.Close; Q_GEN.Free;  exit;
 end;
 btnOK.Enabled := False; editFilter.Enabled := False;
 panelFilters.Enabled := False; panelDates.Enabled := False;
 cbLocGeneral.Enabled := False; cbLocWord.Enabled := False;
 FormMain.DisableAllForms('FormReportSimple');
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
  case editFilter.ItemIndex of
   0..9 : AddLine(editFilter.Text+' = '+Trim(editFilterData.Text),clWindowText,10,'Times New Roman',[]);
   10..11 : AddLine(editFilter.Text+' = '+editDate1.Text+' - '+editDate2.Text,clWindowText,10,'Times New Roman',[]);
  end;
  AddLine('',clWindowText,10,'Times New Roman',[]);
 end;
 TmpFilesCount := 0;
 for i := 1 to Q_GEN.RecordCount do
 begin
  if Q_GEN_SIMPLE_BREAK then Break;
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
   AddLine(Q_GEN.FieldValues['NAME'],clWindowText,12,'Times New Roman',[fsBold]);
   FormatAdres(Q_GEN.FieldValues['ADRES'],Q_GEN.FieldValues['PHONES']);
   if Trim(Q_GEN.FieldValues['EMAIL']) <> '' then begin
    AddLine(FormatEMailWeb(Q_GEN.FieldValues['EMAIL']),clWindowText,10,'Times New Roman',[fsUnderline]);
   end;
   if Trim(Q_GEN.FieldValues['WEB']) <> '' then begin
    AddLine(FormatEMailWeb(Q_GEN.FieldValues['WEB']),clWindowText,10,'Times New Roman',[]);
   end;
   if editFilter.ItemIndex = 0 then begin
   // Отображаем направления только закрепленные за этой рубрикой (RELATIONS)
    if index <> -1 then
     ID_tmp := IntToStr(Integer(editFilterData.Items.Objects[index]))
    else ID_tmp := '-1';
    str1 := Q_GEN.FieldValues['RELATIONS'];
    if Pos('#'+ID_tmp+'=',str1) > 0 then begin
     delete(str1,1,Pos('#'+ID_tmp+'=',str1));
     delete(str1,1,Pos('=',str1));
     delete(str1,Pos('$',str1),length(str1));
     finalStr := '';
     while Pos(')',str1) > 0 do begin
      str2 := copy(str1 ,0 , pos(')',str1));
      delete(str1, 1 , length(str2));
      delete(str2, 1 , 1); delete(str2,length(str2),1);
      finalStr := finalStr + '#' + str2 + '$';
     end;
     AddLine(FormatNapr('NAPR',finalStr),clWindowText,10,'Times New Roman',[fsItalic]);
    end;
   end else // Отображаем все направления
   if Trim(Q_GEN.FieldValues['NAPRAVLENIE']) <> '' then begin
    AddLine(FormatNapr('NAPR',Q_GEN.FieldValues['NAPRAVLENIE']),clWindowText,10,'Times New Roman',[fsItalic]);
   end;
   AddLine('',clWindowText,10,'Times New Roman',[]);
  end;

  Q_GEN.Next;
  Application.ProcessMessages;
 end; // for i := 1 to Q_GEN.RecordCount
 t2 := now;
 FormMain.sStatusBar1.Panels[2].Text := 'Отчет создан за '+IntToStr(SecondsBetween(t1, t2))+' сек';
 RE.Lines.EndUpdate;
 FormMain.SGGeneral.EndUpdate;
 if (Length(RE.Text) > 0) AND (NOT Q_GEN_SIMPLE_BREAK) then
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
   FormReport.FormatReport(TmpFilesCount,RE,ProgressBar,lblStatus);
  end;

  RE.Lines.SaveToFile(AppPath + '\tmp.rtf');
  WordApp := CreateOLEObject('Word.Application');
  WordApp.Documents.Add(AppPath + '\tmp.rtf');
  WordApp.Visible := True;
 end;
 if cbLocGeneral.Checked then FormMain.sPageControl1.ActivePageIndex := 0;
 except on E:Exception do begin
  FormMain.WriteLog('TFormReportSimple.GenerateReport'+#13+'Произошел сбой при генерации отчета'+#13+E.Message);
  MessageBox(Handle,PChar('Произошел сбой при генерации отчета'+#13+E.Message),
   'Ошибка',MB_OK or MB_ICONERROR); end;
 end;
 FormReport.DeleteTempReports;
 lblStatus.Visible := False;
 ProgressBar.Visible := False;
 btnOK.Enabled := True; editFilter.Enabled := True;
 panelFilters.Enabled := True; panelDates.Enabled := True;
 cbLocGeneral.Enabled := True; cbLocWord.Enabled := True;
 FormMain.EnableAllForms('');
 RE.Free; Q_GEN.Close; Q_GEN.Free;
 FormMain.IBDatabase1.Close;
 FormReportSimple.Close;
end;

end.
