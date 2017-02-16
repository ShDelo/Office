unit Report;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, sComboBox, ExtCtrls, sPanel, sButton, sCheckListBox,
  sGroupBox, sListBox, DB, IBC, sRichEdit, sGauge, ComObj,
  sLabel, DateUtils, StrUtils, Mask, sMaskEdit, sCustomComboEdit, sTooledit,
  sCheckBox, NxScrollControl, NxCustomGridControl, NxCustomGrid, NxGrid,
  IniFiles, sMemo, Registry, ComCtrls, ShellAPI, XLSFile, XLSWorkbook, XLSFormat;

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
    cbLocExcel_List: TsCheckBox;
    cbLocExcel_Report: TsCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure IniLoadReport;
    function IsWordInstalled: Boolean;
    procedure IsWordAviable(cbLocGeneral, cbLocWord: TsCheckBox);
    procedure ClearEdits;
    procedure GenerateReport;
    procedure FormatReport(FilesCount: integer; RE: TsRichEdit; PB: TsGauge; lbl: TsLabel);
    procedure DeleteTempReports;
    procedure btnCancelClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure editSelect1Change(Sender: TObject);
    procedure cbLocGeneralClick(Sender: TObject);
    procedure cbLocWordClick(Sender: TObject);
    procedure cbLocExcel_ListClick(Sender: TObject);
    procedure cbLocExcel_ReportClick(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams); override;
    { Public declarations }
  end;

var
  FormReport: TFormReport;
  Q_GEN_BREAK: Boolean = False;

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
  IniMain := TIniFile.Create(AppPath + 'office.ini');
  cbLocGeneral.Checked := IniMain.ReadBool('report', 'cbLocGeneral', False);
  cbLocWord.Checked := IniMain.ReadBool('report', 'cbLocWord', False);
  cbLocExcel_List.Checked := IniMain.ReadBool('report', 'cbLocExcel_List', False);
  cbLocExcel_Report.Checked := IniMain.ReadBool('report', 'cbLocExcel_Report', False);
  if (not cbLocGeneral.Checked) and (not cbLocWord.Checked) and (not cbLocExcel_List.Checked) and (not cbLocExcel_Report.Checked) then
    cbLocWord.Checked := True;
  editFormatDoc.Checked[0] := IniMain.ReadBool('report', 'formatDoc0', False);
  editFormatDoc.Checked[1] := IniMain.ReadBool('report', 'formatDoc1', False);
  editFormatDoc.Checked[2] := IniMain.ReadBool('report', 'formatDoc2', False);
  editFormatDoc.Checked[3] := IniMain.ReadBool('report', 'formatDoc3', False);
  editFormatDoc.Checked[4] := IniMain.ReadBool('report', 'formatDoc4', False);
  editFormatDoc.Checked[5] := IniMain.ReadBool('report', 'formatDoc5', False);
  editFormatDoc.Checked[6] := IniMain.ReadBool('report', 'formatDoc6', False);
  editFormatDoc.Checked[7] := IniMain.ReadBool('report', 'formatDoc7', False);
  editFormatDoc.Checked[8] := IniMain.ReadBool('report', 'formatDoc8', False);
  editFormatDoc.Checked[9] := IniMain.ReadBool('report', 'formatDoc9', False);
  IniMain.Free;
end;

procedure TFormReport.ClearEdits;
begin
  editSelect1.ItemIndex := 0;
  editSelect2.Clear;
  editSelect3.ItemIndex := 0;
  editSelect4.Clear;
  editSelect5.ItemIndex := 0;
  editSelect6.Clear;
  editSelect7.ItemIndex := 0;
  editSelect8.Clear;
  editSelect9.ItemIndex := 0;
  editSelect10.Clear;
  cbDateAdded.Checked := False;
  cbDateEdited.Checked := False;
  editDateAdded1.Date := Now;
  editDateAdded2.Date := Now;
  editDateEdited1.Date := Now;
  editDateEdited2.Date := Now;
end;

function TFormReport.IsWordInstalled: Boolean;
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;
    Result := Reg.KeyExists('Word.Application');
  finally
    Reg.Free;
  end;
end;

procedure TFormReport.IsWordAviable(cbLocGeneral, cbLocWord: TsCheckBox);
begin
  if not IsWordInstalled then
  begin
    MessageBox(Handle, 'В операционной системе не установлена программа для просмотра файлов Microsoft Office.' + #13 +
      'Установите программу и повторите попытку.', 'Предупреждение', MB_OK or MB_ICONWARNING);
    cbLocGeneral.Checked := True;
    cbLocGeneral.Enabled := False;
    cbLocWord.Checked := False;
    cbLocWord.Enabled := False;
  end
  else
  begin
    cbLocGeneral.Enabled := True;
    cbLocWord.Enabled := True;
  end;
end;

procedure TFormReport.cbLocGeneralClick(Sender: TObject);
begin
  if not cbLocGeneral.Checked and not cbLocWord.Checked and not cbLocExcel_List.Checked and not cbLocExcel_Report.Checked then
    cbLocGeneral.Checked := True;
end;

procedure TFormReport.cbLocWordClick(Sender: TObject);
begin
  if not cbLocGeneral.Checked and not cbLocWord.Checked and not cbLocExcel_List.Checked and not cbLocExcel_Report.Checked then
  begin
    cbLocGeneral.Checked := True
  end
  else if cbLocWord.Checked and (cbLocExcel_List.Checked or cbLocExcel_Report.Checked) then
  begin
    cbLocExcel_List.Checked := False;
    cbLocExcel_Report.Checked := False;
  end
end;

procedure TFormReport.cbLocExcel_ListClick(Sender: TObject);
begin
  if not cbLocGeneral.Checked and not cbLocWord.Checked and not cbLocExcel_List.Checked and not cbLocExcel_Report.Checked then
  begin
    cbLocGeneral.Checked := True
  end
  else if cbLocExcel_List.Checked and (cbLocWord.Checked or cbLocExcel_Report.Checked) then
  begin
    cbLocWord.Checked := False;
    cbLocExcel_Report.Checked := False;
  end
end;

procedure TFormReport.cbLocExcel_ReportClick(Sender: TObject);
begin
  if not cbLocGeneral.Checked and not cbLocWord.Checked and not cbLocExcel_List.Checked and not cbLocExcel_Report.Checked then
  begin
    cbLocGeneral.Checked := True
  end
  else if cbLocExcel_Report.Checked and (cbLocWord.Checked or cbLocExcel_List.Checked) then
  begin
    cbLocWord.Checked := False;
    cbLocExcel_List.Checked := False;
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

  procedure editSelectChange(select1, select2: TsComboBox);
  var
    Q: TIBCQuery;
    i: integer;
  begin
    select2.Clear;
    if (select1.ItemIndex = 1) or (select1.ItemIndex = 2) then
    begin
      select2.Items.Add('Да');
      select2.Items.Add('Нет');
      select2.ItemIndex := 0;
    end
    else if select1.ItemIndex = 3 then
    begin { СТРАНА }
      Q := QueryCreate;
      Q.Close;
      Q.SQL.Text := 'select * from COUNTRY order by lower(NAME)';
      Q.Open;
      Q.FetchAll := True;
      for i := 1 to Q.RecordCount do
      begin
        select2.AddItem(Q.FieldValues['NAME'], Pointer(integer(Q.FieldValues['ID'])));
        Q.Next;
      end;
      Q.Close;
      Q.Free;
      if select2.Items.Count > 0 then
        select2.ItemIndex := 0;
    end
    else if select1.ItemIndex = 4 then
    begin { ГОРОД }
      Q := QueryCreate;
      Q.Close;
      Q.SQL.Text := 'select * from GOROD order by lower(NAME)';
      Q.Open;
      Q.FetchAll := True;
      for i := 1 to Q.RecordCount do
      begin
        select2.AddItem(Q.FieldValues['NAME'], Pointer(integer(Q.FieldValues['ID'])));
        Q.Next;
      end;
      Q.Close;
      Q.Free;
      if select2.Items.Count > 0 then
        select2.ItemIndex := 0;
    end
    else if select1.ItemIndex = 5 then
    begin { КУРАТОР }
      for i := 0 to Main.sgCurator_tmp.RowCount - 1 do
      begin
        select2.AddItem(Main.sgCurator_tmp.Cells[0, i], Pointer(StrToInt(Main.sgCurator_tmp.Cells[1, i])));
      end;
      if select2.Items.Count > 0 then
        select2.ItemIndex := 0;
    end
    else if select1.ItemIndex = 6 then
    begin { РУБРИКА }
      for i := 0 to Main.sgRubr_tmp.RowCount - 1 do
      begin
        select2.AddItem(Main.sgRubr_tmp.Cells[0, i], Pointer(StrToInt(Main.sgRubr_tmp.Cells[1, i])));
      end;
      if select2.Items.Count > 0 then
        select2.ItemIndex := 0;
    end
    else if select1.ItemIndex = 7 then
    begin { ТИП }
      for i := 0 to Main.sgType_tmp.RowCount - 1 do
      begin
        select2.AddItem(Main.sgType_tmp.Cells[0, i], Pointer(StrToInt(Main.sgType_tmp.Cells[1, i])));
      end;
      if select2.Items.Count > 0 then
        select2.ItemIndex := 0;
    end;
  end;

begin
  if TsComboBox(Sender).Name = 'editSelect1' then
    editSelectChange(editSelect1, editSelect2);
  if TsComboBox(Sender).Name = 'editSelect3' then
    editSelectChange(editSelect3, editSelect4);
  if TsComboBox(Sender).Name = 'editSelect5' then
    editSelectChange(editSelect5, editSelect6);
  if TsComboBox(Sender).Name = 'editSelect7' then
    editSelectChange(editSelect7, editSelect8);
  if TsComboBox(Sender).Name = 'editSelect9' then
    editSelectChange(editSelect9, editSelect10);
end;

procedure TFormReport.GenerateReport;
var
  i, n, row, RE_TextLength: integer;
  Q_GEN: TIBCQuery;
  REQ, strError, strReportFileName: string;
  d1, d2, d3, d4: string;
  req1, req2, req3, req4, req5: string;
  param1, param2, param3, param4, param5: string;
  RE: TsRichEdit;
  WordApp: OleVariant;
  ExcelApp: OleVariant;
  XLSDoc: TXLSFile;
  bShowReport: Boolean;
  listEmail: TStrings;
  t1, t2: TDateTime;
  TmpFilesCount: integer;

  procedure XLS_SetStyle(nRow, nCol: integer; bHeader: Boolean = False);
  begin
    if (nCol < 0) or (nRow < 0) then
      Exit;

    with XLSDoc.Workbook.Sheets[0] do
    begin
      if bHeader then
      begin
        Cells[nRow, nCol].HAlign := xlHAlignCenter;
        Cells[nRow, nCol].VAlign := xlVAlignCenter;
        // Cells[nRow, nCol].FontBold:= True;
        Rows[nRow].HeightPx := 30;
      end
      else
      begin
        Cells[nRow, nCol].VAlign := xlVAlignTop;
        Cells[nRow, nCol].Wrap := True;
        Rows[nRow].AutoFit;
      end;
      Cells[nRow, nCol].FontName := 'Tahoma';
      Cells[nRow, nCol].FontHeight := 10;
      Cells[nRow, nCol].BorderStyle[xlBorderAll] := bsThin;
      Cells[nRow, nCol].BorderColorRGB[xlBorderAll] := RGB(0, 0, 0);
    end;
  end;

  procedure AddLine(AText: string; AColor: TColor; AFontSize: integer; AFontName: TFontName; AFontStyle: TFontStyles);
  begin
    RE_TextLength := RE_TextLength + Length(AText) + 2;
    RE.SelStart := RE_TextLength - Length(AText) - 2;
    // RE.SelStart := Length(RE.Text);
    RE.SelAttributes.Color := AColor;
    RE.SelAttributes.Size := AFontSize;
    RE.SelAttributes.Name := AFontName;
    RE.SelAttributes.Style := AFontStyle;
    RE.Lines.Add(AText);
  end;

  procedure FormatAdres(AAdres, APhones: string);
  var
    list, list2: TStrings;
    phones, tmp, adres, country_str, oblast_str, city_str, ofType: string;
    x: integer;
  begin
    list := TStringList.Create;
    list.Text := AAdres;
    phones := APhones;
    for x := 0 to list.Count - 1 do
    begin
      // list2[0] = CBAdres; list2[1] = NO; list2[2] = OfficeType; list2[3] = ZIP;
      // list2[4] = Street; list2[5] = Country; list2[6] = Oblast; list2[7] = City;
      list2 := FormEditor.ParseAdresFieldToEntriesList(list[x]);

      tmp := copy(phones, 0, pos('$', phones));
      delete(phones, 1, Length(tmp));
      tmp := copy(phones, 0, pos('$', phones));
      delete(phones, 1, Length(tmp));
      delete(tmp, 1, 1);
      delete(tmp, Length(tmp), 1);
      tmp := AnsiReplaceText(tmp, Chr(13), ', ');
      tmp := AnsiReplaceText(tmp, Chr(10), '');
      tmp := '    Телефон: ' + tmp;
      // в TMP сейчас хранятся все телефоны для адреса list[x]

      if list2[0] = '1' then
      begin
        // такая же процедура в Main.OpenTabByID и Editor.PrepareEdit и Report.GenerateReport и MailSend.SendRegInfoCheck
        ofType := list2[2];
        country_str := list2[5];
        oblast_str := list2[6];
        city_str := list2[7];
        adres := Format('    Адрес: %s - %s, %s, %s,', [FormEditor.GetNameByID('COUNTRY', country_str), list2[3],
          FormEditor.GetNameByID('GOROD', city_str), list2[4]]);
        AddLine(adres, clWindowText, 10, 'Times New Roman', []);
        AddLine(tmp, clWindowText, 10, 'Times New Roman', []);
      end;

      list2.Free;
    end; // for x := 0 to list.Count - 1 do
    list.Free;
  end;

  function GetPhones(APhones: string): string;
  var
    phones, tmp: string;
  begin
    phones := APhones;
    tmp := copy(phones, 0, pos('$', phones));
    delete(phones, 1, Length(tmp));
    tmp := copy(phones, 0, pos('$', phones));
    delete(phones, 1, Length(tmp));
    delete(tmp, 1, 1);
    delete(tmp, Length(tmp), 1);
    // в TMP сейчас хранятся все телефоны для адреса list[x]
    Result := tmp;
  end;

  function GetAdres(AAdres: string): string;
  var
    list, list2, ResultList: TStrings;
    adres, country_str, oblast_str, city_str, ofType: string;
    x: integer;
  begin
    list := TStringList.Create;
    ResultList := TStringList.Create;
    list.Text := AAdres;
    for x := 0 to list.Count - 1 do
    begin
      list2 := FormEditor.ParseAdresFieldToEntriesList(list[x]);
      // list2[0] = CBAdres; list2[1] = NO; list2[2] = OfficeType; list2[3] = ZIP;
      // list2[4] = Street; list2[5] = Country; list2[6] = Oblast; list2[7] = City;

      if list2[0] = '1' then
      begin
        // такая же процедура в Main.OpenTabByID и Editor.PrepareEdit и Report.GenerateReport и MailSend.SendRegInfoCheck
        ofType := list2[2];
        country_str := list2[5];
        oblast_str := list2[6];
        city_str := list2[7];
        if Length(list2[3]) > 0 then
          list2[3] := ' - ' + list2[3];
        if Length(list2[4]) > 0 then
          list2[4] := ', ' + list2[4];
        adres := Format('%s%s, %s%s', [FormEditor.GetNameByID('COUNTRY', country_str), list2[3], FormEditor.GetNameByID('GOROD', city_str),
          list2[4]]);
        adres := Trim(adres);
        if adres[Length(adres)] = ',' then
          delete(adres, Length(adres), 1);
        ResultList.Add(adres);
      end;

      list2.Free;
    end; // for x := 0 to list.Count - 1 do
    Result := Trim(ResultList.Text);
    list.Free;
    ResultList.Free;
  end;

  function FormatRubrNapr(Field, Value: string; bList: Boolean = False): string;
  var
    str, tmp: string;
    list: TStringList;
  begin
    Result := '';
    if Length(Value) = 0 then
      Exit;
    str := Value;
    list := TStringList.Create;
    list.Clear;
    while pos('$', str) > 0 do
    begin
      tmp := copy(str, 0, pos('$', str));
      delete(str, 1, Length(tmp));
      delete(tmp, 1, 1);
      delete(tmp, Length(tmp), 1);
      if AnsiLowerCase(Field) = 'curator' then
        if Main.sgCurator_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
        begin
          Result := Result + ', ' + Main.sgCurator_tmp.Cells[0, Main.sgCurator_tmp.SelectedRow];
          list.Add(Main.sgCurator_tmp.Cells[0, Main.sgCurator_tmp.SelectedRow]);
        end;
      if AnsiLowerCase(Field) = 'rubr' then
        if Main.sgRubr_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
        begin
          Result := Result + ', ' + Main.sgRubr_tmp.Cells[0, Main.sgRubr_tmp.SelectedRow];
          list.Add(Main.sgRubr_tmp.Cells[0, Main.sgRubr_tmp.SelectedRow]);
        end;
      if AnsiLowerCase(Field) = 'type' then
        if Main.sgType_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
        begin
          Result := Result + ', ' + Main.sgType_tmp.Cells[0, Main.sgType_tmp.SelectedRow];
          list.Add(Main.sgType_tmp.Cells[0, Main.sgType_tmp.SelectedRow]);
        end;
      if AnsiLowerCase(Field) = 'napr' then
        if Main.sgNapr_tmp.FindText(1, tmp, [soCaseInsensitive, soExactMatch]) then
        begin
          Result := Result + ', ' + Main.sgNapr_tmp.Cells[0, Main.sgNapr_tmp.SelectedRow];
          list.Add(Main.sgNapr_tmp.Cells[0, Main.sgNapr_tmp.SelectedRow]);
        end;
      if Length(Result) > 0 then
        if Result[1] = ',' then
          delete(Result, 1, 2);
    end;
    if bList then
      Result := Trim(list.Text);
    list.Free;
  end;

  function GetEMailList(fieldEMail: string): TStrings;
  var
    list: TStrings;
    tmp: string;
  begin
    list := TStringList.Create;
    fieldEMail := StringReplace(fieldEMail, ' ', '', [rfReplaceAll]);
    if pos(',', fieldEMail) > 0 then
    begin
      fieldEMail := fieldEMail + ',';
      while pos(',', fieldEMail) > 0 do
      begin
        tmp := copy(fieldEMail, 0, pos(',', fieldEMail));
        delete(fieldEMail, 1, Length(tmp));
        delete(tmp, Length(tmp), 1);
        list.Add(tmp);
      end;
    end
    else if Length(fieldEMail) > 0 then
    begin
      list.Add(fieldEMail);
    end;
    Result := list;
  end;

  procedure GetRequest(select1, select2: TsComboBox; paramNO: string; var REQ, param: string);
  var
    ID: string;
  begin
    REQ := '';
    param := '';
    if select1.ItemIndex = 1 then
    begin // АКТИВНОСТЬ
      REQ := ' (ACTIVITY = :' + paramNO + ') and';
      if select2.ItemIndex = 0 then
        param := '1'
      else
        param := '0';
    end
    else if select1.ItemIndex = 2 then
    begin // АКТУАЛЬНОСТЬ
      REQ := ' (RELEVANCE = :' + paramNO + ') and';
      if select2.ItemIndex = 0 then
        param := '1'
      else
        param := '0';
    end
    else if select1.ItemIndex = 3 then
    begin // СТРАНЫ
      if select2.Items.Count = 0 then
        Exit;
      REQ := ' (ADRES like :' + paramNO + ') and';
      ID := IntToStr(integer(select2.Items.Objects[select2.ItemIndex]));
      param := '%#&' + ID + '$%';
    end
    else if select1.ItemIndex = 4 then
    begin // ГОРОДА
      if select2.Items.Count = 0 then
        Exit;
      REQ := ' (ADRES like :' + paramNO + ') and';
      ID := IntToStr(integer(select2.Items.Objects[select2.ItemIndex]));
      param := '%#^' + ID + '$%';
    end
    else if select1.ItemIndex = 5 then
    begin // КУРАТОРЫ
      if select2.Items.Count = 0 then
        Exit;
      REQ := ' (CURATOR like :' + paramNO + ') and';
      ID := IntToStr(integer(select2.Items.Objects[select2.ItemIndex]));
      param := '%#' + ID + '$%';
    end
    else if select1.ItemIndex = 6 then
    begin // РУБРИКИ
      if select2.Items.Count = 0 then
        Exit;
      REQ := ' (RUBR like :' + paramNO + ') and';
      ID := IntToStr(integer(select2.Items.Objects[select2.ItemIndex]));
      param := '%#' + ID + '$%';
    end
    else if select1.ItemIndex = 7 then
    begin // ТИП
      if select2.Items.Count = 0 then
        Exit;
      REQ := ' (TYPE like :' + paramNO + ') and';
      ID := IntToStr(integer(select2.Items.Objects[select2.ItemIndex]));
      param := '%#' + ID + '$%';
    end;
  end;

begin
  if (editSelect1.ItemIndex = 0) and (editSelect3.ItemIndex = 0) and (editSelect5.ItemIndex = 0) and (editSelect7.ItemIndex = 0) and
    (editSelect9.ItemIndex = 0) and (not cbDateAdded.Checked) and (not cbDateEdited.Checked) then
  begin
    MessageBox(Handle, 'Укажите данные для генерации отчета', 'Предупреждение', MB_OK or MB_ICONWARNING);
    Exit;
  end;
  Q_GEN := QueryCreate;
  RE := TsRichEdit.Create(FormReport);
  RE.Visible := False;
  RE.Parent := FormReport;
  RE.Clear;
  try
    Q_GEN.Close;
    Q_GEN.SQL.Text := ':param1,:param2,:param3,:param4,:param5)'; // tmp
    REQ := 'select * from BASE where (';
    GetRequest(editSelect1, editSelect2, 'param1', req1, param1);
    GetRequest(editSelect3, editSelect4, 'param2', req2, param2);
    GetRequest(editSelect5, editSelect6, 'param3', req3, param3);
    GetRequest(editSelect7, editSelect8, 'param4', req4, param4);
    GetRequest(editSelect9, editSelect10, 'param5', req5, param5);
    if Length(req1) > 0 then
    begin
      REQ := REQ + req1;
      Q_GEN.ParamByName('param1').AsString := param1;
    end;
    if Length(req2) > 0 then
    begin
      REQ := REQ + req2;
      Q_GEN.ParamByName('param2').AsString := param2;
    end;
    if Length(req3) > 0 then
    begin
      REQ := REQ + req3;
      Q_GEN.ParamByName('param3').AsString := param3;
    end;
    if Length(req4) > 0 then
    begin
      REQ := REQ + req4;
      Q_GEN.ParamByName('param4').AsString := param4;
    end;
    if Length(req5) > 0 then
    begin
      REQ := REQ + req5;
      Q_GEN.ParamByName('param5').AsString := param5;
    end;
    d1 := '''' + editDateAdded1.Text + '''';
    d2 := '''' + editDateAdded2.Text + '''';
    d3 := '''' + editDateEdited1.Text + '''';
    d4 := '''' + editDateEdited2.Text + '''';
    if cbDateAdded.Checked then
      REQ := REQ + ' (DATE_ADDED BETWEEN DATE ' + d1 + ' AND DATE ' + d2 + ') and';
    if cbDateEdited.Checked then
      REQ := REQ + ' (DATE_EDITED BETWEEN DATE ' + d3 + ' AND DATE ' + d4 + ') and';
    if REQ = 'select * from BASE where (' then
    begin
      MessageBox(Handle, 'По Вашему запросу не было найдено ни одной записи', 'Информация', MB_OK or MB_ICONINFORMATION);
      RE.Free;
      Q_GEN.Close;
      Q_GEN.Free;
      Exit;
    end;
    delete(REQ, Length(REQ) - 3, Length(REQ));
    REQ := REQ + ' ) order by lower(NAME)';
    WriteLog('TFormReport.GenerateReport: создание расширенного отчета');
    Q_GEN.SQL.Text := REQ;
    Q_GEN.Open;
    Q_GEN.FetchAll := True;
    if Q_GEN.RecordCount = 0 then
    begin
      MessageBox(Handle, 'По Вашему запросу не было найдено ни одной записи', 'Информация', MB_OK or MB_ICONINFORMATION);
      RE.Free;
      Q_GEN.Close;
      Q_GEN.Free;
      Exit;
    end;
    btnOK.Enabled := False;
    editFormatDoc.Enabled := False;
    cbLocGeneral.Enabled := False;
    cbLocWord.Enabled := False;
    FormMain.DisableAllForms('FormReport');
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
      if editSelect1.ItemIndex > 0 then
        AddLine(editSelect1.Text + ' = ' + editSelect2.Text, clWindowText, 10, 'Times New Roman', []);
      if editSelect3.ItemIndex > 0 then
        AddLine(editSelect3.Text + ' = ' + editSelect4.Text, clWindowText, 10, 'Times New Roman', []);
      if editSelect5.ItemIndex > 0 then
        AddLine(editSelect5.Text + ' = ' + editSelect6.Text, clWindowText, 10, 'Times New Roman', []);
      if editSelect7.ItemIndex > 0 then
        AddLine(editSelect7.Text + ' = ' + editSelect8.Text, clWindowText, 10, 'Times New Roman', []);
      if editSelect9.ItemIndex > 0 then
        AddLine(editSelect9.Text + ' = ' + editSelect10.Text, clWindowText, 10, 'Times New Roman', []);
      if cbDateAdded.Checked then
        AddLine(cbDateAdded.Caption + ' = ' + editDateAdded1.Text + ' - ' + editDateAdded2.Text, clWindowText, 10, 'Times New Roman', []);
      if cbDateEdited.Checked then
        AddLine(cbDateEdited.Caption + ' = ' + editDateEdited1.Text + ' - ' + editDateEdited2.Text, clWindowText, 10,
          'Times New Roman', []);
      AddLine('', clWindowText, 10, 'Times New Roman', []);
    end;

    // Init Excel Doc
    XLSDoc := TXLSFile.Create;
    XLSDoc.Workbook.Sheets[0].Name := 'Отчет ' + DateToStr(Now());
    row := 0;

    // Format Excel Doc
    if cbLocExcel_Report.Checked then
    begin
      with XLSDoc.Workbook.Sheets[0] do
      begin
        row := 1;
        // Header
        Cells[0, 0].Value := 'Актив-сть';
        XLS_SetStyle(0, 0, True);
        Cells[0, 1].Value := 'Актуал-сть';
        XLS_SetStyle(0, 1, True);
        Cells[0, 2].Value := 'Куратор';
        XLS_SetStyle(0, 2, True);
        Cells[0, 3].Value := 'Предприятие';
        XLS_SetStyle(0, 3, True);
        Cells[0, 4].Value := 'Деятельность';
        XLS_SetStyle(0, 4, True);
        Cells[0, 5].Value := 'Тип';
        XLS_SetStyle(0, 5, True);
        Cells[0, 6].Value := 'Рубрика';
        XLS_SetStyle(0, 6, True);
        Cells[0, 7].Value := 'Телефон';
        XLS_SetStyle(0, 7, True);
        Cells[0, 8].Value := 'Конт. лицо';
        XLS_SetStyle(0, 8, True);
        Cells[0, 9].Value := 'Адрес';
        XLS_SetStyle(0, 9, True);
        Cells[0, 10].Value := 'E-mail';
        XLS_SetStyle(0, 10, True);
        Cells[0, 11].Value := 'Примечание';
        XLS_SetStyle(0, 11, True);

        // Cols width and visibility
        Columns[0].WidthPx := 70;
        Columns[0].Hidden := not editFormatDoc.Checked[1];

        Columns[1].WidthPx := 70;
        Columns[1].Hidden := not editFormatDoc.Checked[2];

        Columns[2].WidthPx := 70;
        Columns[2].Hidden := not editFormatDoc.Checked[3];

        Columns[3].WidthPx := 150;

        Columns[4].WidthPx := 250;
        Columns[4].Hidden := not editFormatDoc.Checked[6];

        Columns[5].WidthPx := 100;
        Columns[5].Hidden := not editFormatDoc.Checked[5];

        Columns[6].WidthPx := 175;
        Columns[6].Hidden := not editFormatDoc.Checked[4];

        Columns[7].WidthPx := 160;
        Columns[7].Hidden := not editFormatDoc.Checked[0];

        Columns[8].WidthPx := 100;
        Columns[8].Hidden := not editFormatDoc.Checked[9];

        Columns[9].WidthPx := 250;
        Columns[9].Hidden := not editFormatDoc.Checked[0];

        Columns[10].WidthPx := 170;
        Columns[10].Hidden := not editFormatDoc.Checked[8];

        Columns[11].WidthPx := 100;

        // Print settings
        PageSetup.CenterHorizontally := False;
        PageSetup.CenterVertically := False;
        PageSetup.HeaderMargin := 0;
        PageSetup.FooterMargin := 0;
        PageSetup.TopMargin := 0.4;
        PageSetup.BottomMargin := 0.4;
        PageSetup.LeftMargin := 0.5;
        PageSetup.RightMargin := 0.4;
        PageSetup.PrintRowsOnEachPageFrom := 0;
        PageSetup.PrintRowsOnEachPageTo := 0;
        PageSetup.Orientation := xlLandscape;
        PageSetup.FitPagesWidth := 1;
        PageSetup.FitPagesHeight := 0;
        PageSetup.Zoom := False;
      end;
    end;

    TmpFilesCount := 0;
    for i := 1 to Q_GEN.RecordCount do
    begin
      if Q_GEN_BREAK then
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
        AddLine('Фирма: ' + Q_GEN.FieldValues['NAME'], clWindowText, 12, 'Times New Roman', [fsBold]);
        if editFormatDoc.Checked[0] then
          FormatAdres(Q_GEN.FieldValues['ADRES'], Q_GEN.FieldValues['PHONES']);
        if editFormatDoc.Checked[1] then
        begin
          if Q_GEN.FieldByName('ACTIVITY').AsInteger = 1 then
            AddLine('    Активность: да', clWindowText, 10, 'Times New Roman', [])
          else
            AddLine('    Активность: нет', clWindowText, 10, 'Times New Roman', []);
        end;
        if editFormatDoc.Checked[2] then
        begin
          if Q_GEN.FieldByName('RELEVANCE').AsInteger = 1 then
            AddLine('    Актуальность: да', clWindowText, 10, 'Times New Roman', [])
          else
            AddLine('    Актуальность: нет', clWindowText, 10, 'Times New Roman', []);
        end;
        if editFormatDoc.Checked[3] then
          AddLine('    Куратор: ' + FormatRubrNapr('CURATOR', Q_GEN.FieldValues['CURATOR']), clWindowText, 10, 'Times New Roman', []);
        if editFormatDoc.Checked[4] then
          AddLine('    Рубрика: ' + FormatRubrNapr('RUBR', Q_GEN.FieldValues['RUBR']), clWindowText, 10, 'Times New Roman', []);
        if editFormatDoc.Checked[5] then
          AddLine('    Тип: ' + FormatRubrNapr('TYPE', Q_GEN.FieldValues['TYPE']), clWindowText, 10, 'Times New Roman', []);
        if editFormatDoc.Checked[6] then
          AddLine('    Деятельность: ' + FormatRubrNapr('NAPR', Q_GEN.FieldValues['NAPRAVLENIE']), clWindowText, 10, 'Times New Roman', []);
        if editFormatDoc.Checked[7] then
          AddLine('    Веб: ' + Q_GEN.FieldValues['WEB'], clWindowText, 10, 'Times New Roman', []);
        if editFormatDoc.Checked[8] then
          AddLine('    Мейл: ' + Q_GEN.FieldValues['EMAIL'], clWindowText, 10, 'Times New Roman', []);
        if editFormatDoc.Checked[9] then
          AddLine('    Представитель: ' + Q_GEN.FieldValues['FIO'], clWindowText, 10, 'Times New Roman', []);
        AddLine('', clWindowText, 10, 'Times New Roman', []);
      end;

      // build Excel List
      if cbLocExcel_List.Checked then
      begin
        listEmail := GetEMailList(Q_GEN.FieldValues['EMAIL']);
        if Length(listEmail.GetText) > 0 then
        begin
          for n := 0 to listEmail.Count - 1 do
          begin
            XLSDoc.Workbook.Sheets[0].Cells[row, 0].Value := Q_GEN.FieldValues['NAME'];
            XLSDoc.Workbook.Sheets[0].Cells[row, 1].Value := listEmail[n];
            Inc(row);
          end;
        end;
      end
      else if cbLocExcel_Report.Checked then
      begin
        // build Excel Report
        with XLSDoc.Workbook.Sheets[0] do
        begin

          if Q_GEN.FieldByName('ACTIVITY').AsInteger = 1 then
            Cells[row, 0].Value := 'да'
          else
            Cells[row, 0].Value := 'нет';
          XLS_SetStyle(row, 0);
          XLS_SetStyle(row, 0);

          if Q_GEN.FieldByName('RELEVANCE').AsInteger = 1 then
            Cells[row, 1].Value := 'да'
          else
            Cells[row, 1].Value := 'нет';
          XLS_SetStyle(row, 1);

          Cells[row, 2].Value := FormatRubrNapr('CURATOR', Q_GEN.FieldValues['CURATOR'], True);
          XLS_SetStyle(row, 2);
          Cells[row, 3].Value := Q_GEN.FieldValues['NAME'];
          XLS_SetStyle(row, 3);
          Cells[row, 4].Value := FormatRubrNapr('NAPR', Q_GEN.FieldValues['NAPRAVLENIE'], True);
          XLS_SetStyle(row, 4);
          Cells[row, 5].Value := FormatRubrNapr('TYPE', Q_GEN.FieldValues['TYPE'], True);
          XLS_SetStyle(row, 5);
          Cells[row, 6].Value := FormatRubrNapr('RUBR', Q_GEN.FieldValues['RUBR'], True);
          XLS_SetStyle(row, 6);
          Cells[row, 7].Value := GetPhones(Q_GEN.FieldValues['PHONES']);
          XLS_SetStyle(row, 7);
          Cells[row, 8].Value := Q_GEN.FieldValues['FIO'];
          XLS_SetStyle(row, 8);
          Cells[row, 9].Value := GetAdres(Q_GEN.FieldValues['ADRES']);
          XLS_SetStyle(row, 9);
          Cells[row, 10].Value := Trim(GetEMailList(Q_GEN.FieldValues['EMAIL']).Text);
          XLS_SetStyle(row, 10);
          Cells[row, 11].Value := '';
          XLS_SetStyle(row, 11);
        end;

        Inc(row);
      end;

      Q_GEN.Next;
      Application.ProcessMessages;
    end; // for i := 1 to Q_GEN.RecordCount
    t2 := Now;
    FormMain.sStatusBar1.Panels[2].Text := 'Отчет создан за ' + IntToStr(SecondsBetween(t1, t2)) + ' сек';
    RE.Lines.EndUpdate;
    FormMain.SGGeneral.EndUpdate;
    if (Length(RE.Text) > 0) and (not Q_GEN_BREAK) then
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
        FormatReport(TmpFilesCount, RE, ProgressBar, lblStatus);
      end;
      RE.Lines.SaveToFile(AppPath + '\tmp.rtf');
      WordApp := CreateOLEObject('Word.Application');
      WordApp.Documents.Add(AppPath + '\tmp.rtf');
      WordApp.Visible := True;
      { ShellExecute(handle,'open',PChar(AppPath + '\tmp.rtf'),nil,nil,SW_SHOWNORMAL); }
    end;

    // close Excel, save report file and show it
    bShowReport := True;
    strReportFileName := AppPath + 'report.xls';
    strError := 'ExcelApp: no errors';
    if cbLocExcel_List.Checked or cbLocExcel_Report.Checked then
    begin

      if cbLocExcel_List.Checked then
        XLSDoc.Workbook.Sheets[0].Columns.AutoFit(0, 1); // AutoFit for List
      if cbLocExcel_Report.Checked then
        XLSDoc.Workbook.Sheets[0].Freeze(1, 0); // Freeze header row for Report

      try // Closing
        ExcelApp := GetActiveOleObject('Excel.Application');
        ExcelApp.Visible := True;
        for i := ExcelApp.Workbooks.Count downto 1 do
          ExcelApp.Workbooks[i].Save;
        ExcelApp.Workbooks.Close;
        ExcelApp.Quit;
        VarClear(ExcelApp);
      except
        on E: Exception do
          strError := 'ExcelApp: ' + E.Message;
      end;

      try // Saving
        XLSDoc.SaveAs(strReportFileName);
      except
        on E: Exception do
        begin
          WriteLog('TFormReport.GenerateReport' + #13 + strError + #13 + 'Ошибка при сохранении файла отчета:' + #13 + E.Message);
          MessageBox(Handle, PChar('Ошибка при сохранении файла отчета:' + #13 + strError + #13 + E.Message), 'Ошибка',
            MB_OK or MB_ICONERROR);
          bShowReport := False;
        end;
      end;

      if bShowReport then
      begin
        try // Showing
          ExcelApp := CreateOLEObject('Excel.Application');
          ExcelApp.Visible := True;
          ExcelApp.Workbooks.Open(strReportFileName);
          // ShellExecute(0, 'open', PChar(strReportFileName'), nil, nil, SW_SHOW);
        except
          on E: Exception do
          begin
            WriteLog('TFormReport.GenerateReport' + #13 + 'Ошибка при отображении файла отчета:' + #13 + E.Message);
            MessageBox(Handle, PChar('Ошибка при отображении файла отчета:' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
          end;
        end;
      end;

      XLSDoc.Destroy;
    end;

    if cbLocGeneral.Checked then
      FormMain.sPageControl1.ActivePageIndex := 0;
  except
    on E: Exception do
    begin
      WriteLog('TFormReport.GenerateReport' + #13 + 'Произошел сбой при генерации отчета' + #13 + E.Message);
      MessageBox(Handle, PChar('Произошел сбой при генерации отчета' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
    end;
  end;
  DeleteTempReports;
  lblStatus.Visible := False;
  ProgressBar.Visible := False;
  btnOK.Enabled := True;
  editFormatDoc.Enabled := True;
  cbLocGeneral.Enabled := True;
  cbLocWord.Enabled := True;
  FormMain.EnableAllForms('');
  RE.Free;
  Q_GEN.Close;
  Q_GEN.Free;
  FormMain.IBDatabase1.Close;
  FormReport.Close;
  WriteLog('TFormReport.GenerateReport: отчет создан успешно');
end;

procedure TFormReport.FormatReport(FilesCount: integer; RE: TsRichEdit; PB: TsGauge; lbl: TsLabel);
var
  CurFileNO: integer;
  FileName, UnitedText, tmp: string;
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
      FileName := 'report_' + IntToStr(CurFileNO) + '.tmp';
      RE.Lines.LoadFromFile(FileName);
      RE.Lines.SaveToStream(TextStream);
      tmp := TextStream.DataString;
      TextStream.Position := 0;
      if CurFileNO = 1 then
        tmp := copy(tmp, 0, Length(tmp) - 5)
      else if CurFileNO = FilesCount then
        tmp := copy(tmp, 2, Length(tmp))
      else
        tmp := copy(tmp, 2, Length(tmp) - 5);
      UnitedText := UnitedText + tmp;
      Application.ProcessMessages;
    end;
    lbl.Caption := 'Загрузка отчета...';
    Application.ProcessMessages;
    TextStream.WriteString(UnitedText);
    TextStream.Position := 0;
    RE.Lines.LoadFromStream(TextStream);
    // RE.Text := UnitedText; // как вариант вместо TextStream.WriteString(UnitedText)
  finally
    TextStream.Free;
  end;
end;

procedure TFormReport.DeleteTempReports;

  function FilesInDir(Dir: string): integer;
  var
    sr: TSearchRec;
  begin
    Result := 0;
    if FindFirst(Dir + '\*', faAnyFile - faDirectory, sr) <> 0 then
    begin
      FindClose(sr);
      Exit;
    end;
    repeat
      Inc(Result);
    until (FindNext(sr) <> 0);
    FindClose(sr);
  end;

var
  i: integer;
  FileName: string;
begin
  try
    for i := 1 to FilesInDir(AppPath) do
    begin
      FileName := AppPath + '\report_' + IntToStr(i) + '.tmp';
      if FileExists(FileName) then
        DeleteFile(FileName);
    end;
  except
    on E: Exception do
    begin
      WriteLog('TFormReport.GenerateReport' + #13 + 'Ошибка удаления временных файлов' + #13 + E.Message);
      MessageBox(Handle, PChar('Ошибка удаления временных файлов' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
    end;
  end;
end;

end.
