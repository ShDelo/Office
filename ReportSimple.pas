unit ReportSimple;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, sButton, sCheckBox, sGroupBox, NxColumns,
  NxColumnClasses, NxScrollControl, NxCustomGridControl, NxCustomGrid,
  NxGrid, sEdit, sComboBox, ExtCtrls, sPanel, IniFiles, IBC, sLabel,
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
  Q_GEN_SIMPLE_BREAK: Boolean = False;

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

procedure TFormReportSimple.FormClose(Sender: TObject; var Action: TCloseAction);
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
  panelFilters.Visible := True;
  panelDates.Visible := False;
end;

procedure TFormReportSimple.IniLoadReportSimple;
begin
  IniMain := TIniFile.Create(AppPath + 'office.ini');
  cbLocGeneral.Checked := IniMain.ReadBool('reportsimple', 'cbLocGeneral', False);
  cbLocWord.Checked := IniMain.ReadBool('reportsimple', 'cbLocWord', False);
  if (not cbLocGeneral.Checked) and (not cbLocWord.Checked) then
    cbLocWord.Checked := True;
  IniMain.Free;
end;

procedure TFormReportSimple.cbLocGeneralClick(Sender: TObject);
begin
  if not cbLocGeneral.Checked and not cbLocWord.Checked then
    cbLocGeneral.Checked := True;
end;

procedure TFormReportSimple.editFilterSelect(Sender: TObject);
var
  Q: TIBCQuery;
  i, n: integer;
  memo: TStrings;
  Phones, Email, Web, tmp: string;
begin
  editFilterData.Items.BeginUpdate;
  editFilterData.Clear;
  editDate1.Date := Now;
  editDate2.Date := Now;
  case editFilter.ItemIndex of
    0 .. 9:
      begin
        panelFilters.Visible := True;
        panelDates.Visible := False;
      end;
    10 .. 11:
      begin
        panelFilters.Visible := False;
        panelDates.Visible := True;
      end;
  end;
  Q := QueryCreate;
  if editFilter.ItemIndex = 0 then
  begin // РУБРИКИ
    for i := 0 to Main.sgRubr_tmp.RowCount - 1 do
    begin
      editFilterData.AddItem(Main.sgRubr_tmp.Cells[0, i], Pointer(StrToInt(Main.sgRubr_tmp.Cells[1, i])));
    end;
  end
  else if editFilter.ItemIndex = 1 then
  begin // ГОРОДА
    Q.Close;
    Q.SQL.Text := 'select * from CITY';
    Q.Open;
    Q.FetchAll := True;
    for i := 1 to Q.RecordCount do
    begin
      editFilterData.AddItem(Q.FieldValues['NAME'], Pointer(integer(Q.FieldValues['ID'])));
      Q.Next;
    end;
  end
  else if editFilter.ItemIndex = 2 then
  begin // СТРАНЫ
    Q.Close;
    Q.SQL.Text := 'select * from COUNTRY';
    Q.Open;
    Q.FetchAll := True;
    for i := 1 to Q.RecordCount do
    begin
      editFilterData.AddItem(Q.FieldValues['NAME'], Pointer(integer(Q.FieldValues['ID'])));
      Q.Next;
    end;
  end
  else if editFilter.ItemIndex = 3 then
  begin // КУРАТОРЫ
    for i := 0 to Main.sgCurator_tmp.RowCount - 1 do
    begin
      editFilterData.AddItem(Main.sgCurator_tmp.Cells[0, i], Pointer(StrToInt(Main.sgCurator_tmp.Cells[1, i])));
    end;
  end
  else if editFilter.ItemIndex = 4 then
  begin // ТИПЫ
    for i := 0 to Main.sgType_tmp.RowCount - 1 do
    begin
      editFilterData.AddItem(Main.sgType_tmp.Cells[0, i], Pointer(StrToInt(Main.sgType_tmp.Cells[1, i])));
    end;
  end
  else if editFilter.ItemIndex = 5 then
  begin // АКТИВНОСТЬ
    editFilterData.AddItem('Нет', Pointer(0));
    editFilterData.AddItem('Да', Pointer(1));
  end
  else if editFilter.ItemIndex = 6 then
  begin // АКТУАЛЬНОСТЬ
    editFilterData.AddItem('Нет', Pointer(0));
    editFilterData.AddItem('Да', Pointer(1));
  end
  else if editFilter.ItemIndex = 7 then
  begin // ТЕЛЕФОНЫ
    Q.Close;
    Q.SQL.Text := 'select PHONES from BASE';
    Q.Open;
    Q.FetchAll := True;
    memo := TStringList.Create;
    for i := 1 to Q.RecordCount do
    begin
      Phones := Q.Fields[0].AsString;
      while pos('$', Phones) > 0 do
      begin
        memo.Clear;
        tmp := copy(Phones, 0, pos('$', Phones));
        delete(Phones, 1, length(tmp));
        tmp := copy(Phones, 0, pos('$', Phones));
        delete(Phones, 1, length(tmp));
        delete(tmp, 1, 1);
        delete(tmp, length(tmp), 1);
        if trim(tmp) = '' then
          Continue;
        memo.Text := tmp;
        for n := 0 to memo.Count - 1 do
        begin
          tmp := memo[n];
          if pos(')', tmp) > 0 then
            delete(tmp, 1, pos(')', tmp));
          tmp := trim(tmp);
          editFilterData.AddItem(tmp, Pointer(0));
        end;
      end;
      Q.Next;
    end;
    memo.Free;
  end
  else if editFilter.ItemIndex = 8 then
  begin // МЕЙЛ
    Q.Close;
    Q.SQL.Text := 'select EMAIL from BASE';
    Q.Open;
    Q.FetchAll := True;
    for i := 1 to Q.RecordCount do
    begin
      Email := Q.Fields[0].AsString;
      if length(trim(Email)) > 0 then
      begin
        if Email[length(Email)] <> ',' then
          Email := Email + ',';
        while pos(',', Email) > 0 do
        begin
          tmp := copy(Email, 0, pos(',', Email));
          delete(Email, 1, length(tmp));
          tmp := trim(tmp);
          delete(tmp, length(tmp), 1);
          editFilterData.AddItem(tmp, Pointer(0));
        end;
      end;
      Q.Next;
    end;
  end
  else if editFilter.ItemIndex = 9 then
  begin // САЙТ
    Q.Close;
    Q.SQL.Text := 'select WEB from BASE';
    Q.Open;
    Q.FetchAll := True;
    for i := 1 to Q.RecordCount do
    begin
      Web := Q.Fields[0].AsString;
      if length(trim(Web)) > 0 then
      begin
        if Web[length(Web)] <> ',' then
          Web := Web + ',';
        while pos(',', Web) > 0 do
        begin
          tmp := copy(Web, 0, pos(',', Web));
          delete(Web, 1, length(tmp));
          tmp := trim(tmp);
          delete(tmp, length(tmp), 1);
          editFilterData.AddItem(tmp, Pointer(0));
        end;
      end;
      Q.Next;
    end;
  end
  else if editFilter.ItemIndex = 10 then
  begin
    lblDate.Caption := 'Дата добавления (период):'
  end
  else if editFilter.ItemIndex = 11 then
  begin
    lblDate.Caption := 'Дата изменения (период):'
  end;
  Q.Close;
  Q.Free;
  FormMain.IBDatabase1.Close;
  editFilterData.Items.EndUpdate;
  if editFilterData.Items.Count > 0 then
    editFilterData.ItemIndex := 0;
  if editFilter.ItemIndex in [0 .. 9] then
  begin
    editFilterData.SetFocus;
    editFilterData.DroppedDown := True;
  end;
end;

procedure TFormReportSimple.GenerateReport;
var
  i, RE_TextLength, index: integer;
  Q_GEN: TIBCQuery;
  REQ, param, d1, d2, ID_tmp, str1, str2, finalStr: string;
  RE: TsRichEdit;
  WordApp: OleVariant;
  t1, t2: TDateTime;
  TmpFilesCount: integer;

  procedure AddLine(AText: string; AColor: TColor; AFontSize: integer; AFontName: TFontName; AFontStyle: TFontStyles);
  begin
    RE_TextLength := RE_TextLength + length(AText) + 2;
    RE.SelStart := RE_TextLength - length(AText) - 2;
    RE.SelAttributes.Color := AColor;
    RE.SelAttributes.Size := AFontSize;
    RE.SelAttributes.Name := AFontName;
    RE.SelAttributes.Style := AFontStyle;
    RE.Lines.Add(AText);
  end;

  procedure FormatAdres(AAdres, APhones: string);
  var
    list, list2: TStrings;
    Phones, tmp, adres, country_str, region_str, city_str, ofType: string;
    x: integer;
  begin
    list := TStringList.Create;
    list.Text := AAdres;
    Phones := APhones;
    for x := 0 to list.Count - 1 do
    begin
      // такая же процедура в Main.OpenTabByID и Editor.PrepareEdit и Report.GenerateReport и MailSend.SendRegInfoCheck
      // list2[0] = CBAdres; list2[1] = NO; list2[2] = OfficeType; list2[3] = ZIP;
      // list2[4] = Street; list2[5] = Country; list2[6] = Region; list2[7] = City;
      list2 := FormEditor.ParseAdresFieldToEntriesList(list[x]);

      tmp := copy(Phones, 0, pos('$', Phones));
      delete(Phones, 1, length(tmp));
      tmp := copy(Phones, 0, pos('$', Phones));
      delete(Phones, 1, length(tmp));
      delete(tmp, 1, 1);
      delete(tmp, length(tmp), 1);
      tmp := StringReplace(tmp, '(', '', [rfReplaceAll]);
      tmp := StringReplace(tmp, ')', '', [rfReplaceAll]);
      // в TMP сейчас хранятся все телефоны для адреса list[x] в формате Мемо

      if list2[0] = '1' then
      begin
        ofType := list2[2];
        country_str := list2[5];
        region_str := list2[6];
        city_str := list2[7];
        adres := '';
        if trim(list2[3]) <> '' then
          adres := list2[3] + ', '; // ZIP
        if trim(city_str) <> '' then
          adres := adres + FormEditor.GetNameByID('CITY', city_str) + ','; // CITY
        if trim(adres) <> '' then
          AddLine(adres, clWindowText, 10, 'Times New Roman', []);
        if trim(list2[4]) <> '' then
          AddLine(list2[4], clWindowText, 10, 'Times New Roman', []);
        { adres := adres + Format('%s, %s, %s,',
          [FormEditor.GetNameByID('COUNTRY',country_str),FormEditor.GetNameByID('CITY',city_str),list2[4]]);
          if Trim(adres) <> ', , ,' then AddLine(adres,clWindowText,10,'Times New Roman',[],); }
        if trim(tmp) <> '' then
          AddLine(tmp, clWindowText, 10, 'Times New Roman', [fsItalic]);
        // AddLine('',clWindowText,10,'Times New Roman',[]);
      end;

      list2.Free;

    end; // for x := 0 to list.Count - 1 do
    list.Free;
  end;

  function FormatNapr(Field, Value: string): string;
  var
    str, tmp: string;
  begin
    Result := '';
    if length(Value) = 0 then
      exit;
    str := Value;
    while pos('$', str) > 0 do
    begin
      tmp := copy(str, 0, pos('$', str));
      delete(str, 1, length(tmp));
      delete(tmp, 1, 1);
      delete(tmp, length(tmp), 1);
      if AnsiLowerCase(Field) = 'napr' then
        if Main.sgNapr_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
          Result := Result + #13 + Main.sgNapr_tmp.Cells[0, Main.sgNapr_tmp.SelectedRow];
      if length(Result) > 0 then
        if Result[1] = #13 then
          delete(Result, 1, 1);
    end;
  end;

  function FormatEMailWeb(Value: string): string;
  var
    str, tmp: string;
  begin
    Result := '';
    if length(Value) = 0 then
      exit;
    str := Value;
    if str[length(str)] <> ',' then
      str := str + ',';
    while pos(',', str) > 0 do
    begin
      tmp := copy(str, 0, pos(',', str));
      delete(str, 1, length(tmp));
      tmp := trim(tmp);
      delete(tmp, length(tmp), 1);
      Result := Result + #13 + tmp;
      if length(Result) > 0 then
        if Result[1] = #13 then
          delete(Result, 1, 1);
    end;
  end;

  procedure GetRequest(var REQ, param: string);
  var
    ID: string;
    index: integer;
  begin
    REQ := '';
    param := '';
    index := editFilterData.Items.IndexOf(trim(editFilterData.Text));
    if index <> -1 then
      ID := IntToStr(integer(editFilterData.Items.Objects[index]))
    else
      ID := '-1';
    if editFilter.ItemIndex = 0 then
    begin // РУБРИКА
      REQ := 'select * from BASE where RUBR like :param and ACTIVITY = 1 order by lower(NAME)';
      param := '%#' + ID + '$%';
    end
    else if editFilter.ItemIndex = 1 then
    begin // ГОРОД
      REQ := 'select * from BASE where ADRES like :param order by lower(NAME)';
      param := '%#^' + ID + '$%';
    end
    else if editFilter.ItemIndex = 2 then
    begin // СТРАНА
      REQ := 'select * from BASE where ADRES like :param order by lower(NAME)';
      param := '%#&' + ID + '$%';
    end
    else if editFilter.ItemIndex = 3 then
    begin // КУРАТОР
      REQ := 'select * from BASE where CURATOR like :param order by lower(NAME)';
      param := '%#' + ID + '$%';
    end
    else if editFilter.ItemIndex = 4 then
    begin // ТИП
      REQ := 'select * from BASE where TYPE like :param order by lower(NAME)';
      param := '%#' + ID + '$%';
    end
    else if editFilter.ItemIndex = 5 then
    begin // АКТИВНОСТЬ
      REQ := 'select * from BASE where ACTIVITY like :param order by lower(NAME)';
      if AnsiLowerCase(trim(editFilterData.Text)) = 'да' then
        ID := '1'
      else if AnsiLowerCase(trim(editFilterData.Text)) = 'нет' then
        ID := '0'
      else
        ID := '-2';
      param := ID;
    end
    else if editFilter.ItemIndex = 6 then
    begin // АКТУАЛЬНОСТЬ
      REQ := 'select * from BASE where RELEVANCE like :param order by lower(NAME)';
      if AnsiLowerCase(trim(editFilterData.Text)) = 'да' then
        ID := '1'
      else if AnsiLowerCase(trim(editFilterData.Text)) = 'нет' then
        ID := '0'
      else
        ID := '-2';
      param := ID;
    end
    else if editFilter.ItemIndex = 7 then
    begin // ТЕЛЕФОН
      REQ := 'select * from BASE where PHONES like :param order by lower(NAME)';
      param := '%' + AnsiLowerCase(trim(editFilterData.Text)) + '%';
    end
    else if editFilter.ItemIndex = 8 then
    begin // МЕЙЛ
      REQ := 'select * from BASE where lower(EMAIL) like :param order by lower(NAME)';
      param := '%' + AnsiLowerCase(trim(editFilterData.Text)) + '%';
    end
    else if editFilter.ItemIndex = 9 then
    begin // САЙТ
      REQ := 'select * from BASE where lower(WEB) like :param order by lower(NAME)';
      param := '%' + AnsiLowerCase(trim(editFilterData.Text)) + '%';
    end
    else if editFilter.ItemIndex = 10 then
    begin // ДАТА ДОБАВЛЕНИЯ
      d1 := '''' + editDate1.Text + '''';
      d2 := '''' + editDate2.Text + '''';
      REQ := 'select * from BASE where (DATE_ADDED BETWEEN DATE ' + d1 + ' AND DATE ' + d2 + ') order by lower(NAME)';
    end
    else if editFilter.ItemIndex = 11 then
    begin // ДАТА ИЗМЕНЕНИЯ
      d1 := '''' + editDate1.Text + '''';
      d2 := '''' + editDate2.Text + '''';
      REQ := 'select * from BASE where (DATE_EDITED BETWEEN DATE ' + d1 + ' AND DATE ' + d2 + ') order by lower(NAME)';
    end;
  end;

begin
  if (editFilter.ItemIndex = -1) or ((editFilter.ItemIndex in [0 .. 9]) and (trim(editFilterData.Text) = '')) then
  begin
    MessageBox(handle, 'Укажите данные для генерации отчета', 'Предупреждение', MB_OK or MB_ICONWARNING);
    exit;
  end;
  index := editFilterData.Items.IndexOf(trim(editFilterData.Text));
  Q_GEN := QueryCreate;
  RE := TsRichEdit.Create(FormReportSimple);
  RE.Visible := False;
  RE.Parent := FormReportSimple;
  RE.Clear;
  try
    Q_GEN.Close;
    Q_GEN.SQL.Text := ':param'; // tmp
    GetRequest(REQ, param);
    if length(REQ) > 0 then
    begin
      Q_GEN.ParamByName('param').AsString := param;
    end;
    if REQ = '' then
    begin
      MessageBox(handle, 'По Вашему запросу не было найдено ни одной записи', 'Информация', MB_OK or MB_ICONINFORMATION);
      RE.Free;
      Q_GEN.Close;
      Q_GEN.Free;
      exit;
    end;
    WriteLog('TFormReportSimple.GenerateReport: создание простого отчета');
    Q_GEN.SQL.Text := REQ;
    Q_GEN.Open;
    Q_GEN.FetchAll := True;
    if Q_GEN.RecordCount = 0 then
    begin
      MessageBox(handle, 'По Вашему запросу не было найдено ни одной записи', 'Информация', MB_OK or MB_ICONINFORMATION);
      RE.Free;
      Q_GEN.Close;
      Q_GEN.Free;
      exit;
    end;
    btnOK.Enabled := False;
    editFilter.Enabled := False;
    panelFilters.Enabled := False;
    panelDates.Enabled := False;
    cbLocGeneral.Enabled := False;
    cbLocWord.Enabled := False;
    FormMain.DisableAllForms('FormReportSimple');
    RE.Lines.BeginUpdate;
    FormMain.SGGeneral.BeginUpdate;
    if cbLocGeneral.Checked then
      FormMain.SGGeneral.ClearRows;
    lblStatus.Caption := 'Построение отчета...';
    lblStatus.Visible := True;
    ProgressBar.MinValue := 0;
    ProgressBar.MaxValue := Q_GEN.RecordCount;
    ProgressBar.Visible := True;
    t1 := Now;
    if cbLocWord.Checked then
    begin
      AddLine('Найдено записей: ' + IntToStr(Q_GEN.RecordCount), clWindowText, 10, 'Times New Roman', []);
      AddLine('Условия выбора:', clWindowText, 10, 'Times New Roman', []);
      case editFilter.ItemIndex of
        0 .. 9:
          AddLine(editFilter.Text + ' = ' + trim(editFilterData.Text), clWindowText, 10, 'Times New Roman', []);
        10 .. 11:
          AddLine(editFilter.Text + ' = ' + editDate1.Text + ' - ' + editDate2.Text, clWindowText, 10, 'Times New Roman', []);
      end;
      AddLine('', clWindowText, 10, 'Times New Roman', []);
    end;
    TmpFilesCount := 0;
    for i := 1 to Q_GEN.RecordCount do
    begin
      if Q_GEN_SIMPLE_BREAK then
        Break;
      ProgressBar.Progress := i;

      if RE.Lines.Count >= 10000 then
      begin
        Inc(TmpFilesCount, 1);
        RE.Lines.SaveToFile(AppPath + '\report_' + IntToStr(TmpFilesCount) + '.tmp');
        RE.Clear;
      end;

      if cbLocGeneral.Checked then
      begin
        FormMain.SGAddRow(FormMain.SGGeneral, Q_GEN.FieldByName('ACTIVITY').AsInteger, Q_GEN.FieldByName('RELEVANCE').AsInteger,
          Q_GEN.FieldValues['NAME'], Q_GEN.FieldValues['CURATOR'], Q_GEN.FieldValues['DATE_ADDED'], Q_GEN.FieldValues['DATE_EDITED'],
          Q_GEN.FieldValues['WEB'], Q_GEN.FieldValues['EMAIL'], Q_GEN.FieldValues['TYPE'], Q_GEN.FieldValues['ID'],
          Q_GEN.FieldValues['FIO'], Q_GEN.FieldValues['RUBR']);
      end;

      if cbLocWord.Checked then
      begin
        AddLine(Q_GEN.FieldValues['NAME'], clWindowText, 12, 'Times New Roman', [fsBold]);
        FormatAdres(Q_GEN.FieldValues['ADRES'], Q_GEN.FieldValues['PHONES']);
        if trim(Q_GEN.FieldValues['EMAIL']) <> '' then
        begin
          AddLine(FormatEMailWeb(Q_GEN.FieldValues['EMAIL']), clWindowText, 10, 'Times New Roman', [fsUnderline]);
        end;
        if trim(Q_GEN.FieldValues['WEB']) <> '' then
        begin
          AddLine(FormatEMailWeb(Q_GEN.FieldValues['WEB']), clWindowText, 10, 'Times New Roman', []);
        end;
        if editFilter.ItemIndex = 0 then
        begin
          // Отображаем направления только закрепленные за этой рубрикой (RELATIONS)
          if index <> -1 then
            ID_tmp := IntToStr(integer(editFilterData.Items.Objects[index]))
          else
            ID_tmp := '-1';
          str1 := Q_GEN.FieldValues['RELATIONS'];
          if pos('#' + ID_tmp + '=', str1) > 0 then
          begin
            delete(str1, 1, pos('#' + ID_tmp + '=', str1));
            delete(str1, 1, pos('=', str1));
            delete(str1, pos('$', str1), length(str1));
            finalStr := '';
            while pos(')', str1) > 0 do
            begin
              str2 := copy(str1, 0, pos(')', str1));
              delete(str1, 1, length(str2));
              delete(str2, 1, 1);
              delete(str2, length(str2), 1);
              finalStr := finalStr + '#' + str2 + '$';
            end;
            AddLine(FormatNapr('NAPR', finalStr), clWindowText, 10, 'Times New Roman', [fsItalic]);
          end;
        end
        else { // Отображаем все направления }
          if trim(Q_GEN.FieldValues['NAPRAVLENIE']) <> '' then
          begin
            AddLine(FormatNapr('NAPR', Q_GEN.FieldValues['NAPRAVLENIE']), clWindowText, 10, 'Times New Roman', [fsItalic]);
          end;
        AddLine('', clWindowText, 10, 'Times New Roman', []);
      end;

      Q_GEN.Next;
      Application.ProcessMessages;
    end; // for i := 1 to Q_GEN.RecordCount
    t2 := Now;
    FormMain.sStatusBar1.Panels[2].Text := 'Отчет создан за ' + IntToStr(SecondsBetween(t1, t2)) + ' сек';
    RE.Lines.EndUpdate;
    FormMain.SGGeneral.EndUpdate;
    if (length(RE.Text) > 0) and (not Q_GEN_SIMPLE_BREAK) then
    begin
      try
        WordApp := GetActiveOleObject('Word.Application');
        WordApp.Documents.Close(0);
        WordApp.Quit;
      except
      end;
      if TmpFilesCount > 0 then
      begin
        // сохраняем посл часть отчета потому-что было разделение и форматируем
        Inc(TmpFilesCount, 1);
        RE.Lines.SaveToFile(AppPath + '\report_' + IntToStr(TmpFilesCount) + '.tmp');
        FormReport.FormatReport(TmpFilesCount, RE, ProgressBar, lblStatus);
      end;

      RE.Lines.SaveToFile(AppPath + '\tmp.rtf');
      WordApp := CreateOLEObject('Word.Application');
      WordApp.Documents.Add(AppPath + '\tmp.rtf');
      WordApp.Visible := True;
    end;
    if cbLocGeneral.Checked then
      FormMain.sPageControl1.ActivePageIndex := 0;
  except
    on E: Exception do
    begin
      WriteLog('TFormReportSimple.GenerateReport' + #13 + 'Произошел сбой при генерации отчета' + #13 + E.Message);
      MessageBox(handle, PChar('Произошел сбой при генерации отчета' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
    end;
  end;
  FormReport.DeleteTempReports;
  lblStatus.Visible := False;
  ProgressBar.Visible := False;
  btnOK.Enabled := True;
  editFilter.Enabled := True;
  panelFilters.Enabled := True;
  panelDates.Enabled := True;
  cbLocGeneral.Enabled := True;
  cbLocWord.Enabled := True;
  FormMain.EnableAllForms('');
  RE.Free;
  Q_GEN.Close;
  Q_GEN.Free;
  FormMain.IBDatabase1.Close;
  FormReportSimple.Close;
  WriteLog('TFormReportSimple.GenerateReport: отчет создан успешно');
end;

end.
