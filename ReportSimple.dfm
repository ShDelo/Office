object FormReportSimple: TFormReportSimple
  Left = 335
  Top = 143
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1043#1077#1085#1077#1088#1072#1094#1080#1103' '#1086#1090#1095#1077#1090#1072
  ClientHeight = 250
  ClientWidth = 658
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object panelSelectData: TsPanel
    Left = 0
    Top = 0
    Width = 658
    Height = 250
    Align = alClient
    TabOrder = 0
    SkinData.SkinSection = 'PANEL'
    object groupLocation: TsGroupBox
      Left = 463
      Top = 8
      Width = 186
      Height = 205
      Caption = #1042#1099#1074#1077#1089#1090#1080' '#1076#1072#1085#1085#1099#1077' '#1085#1072
      TabOrder = 1
      SkinData.SkinSection = 'GROUPBOX'
      object ProgressBar: TsGauge
        Left = 4
        Top = 179
        Width = 176
        Height = 22
        Visible = False
        SkinData.SkinSection = 'GAUGE'
        ForeColor = clBlack
        MaxValue = 100
        Progress = 50
        Suffix = '%'
      end
      object lblStatus: TsLabel
        Left = 4
        Top = 164
        Width = 40
        Height = 13
        Caption = 'lblStatus'
        ParentFont = False
        Visible = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
      end
      object cbLocGeneral: TsCheckBox
        Left = 12
        Top = 24
        Width = 151
        Height = 17
        Caption = #1043#1083#1072#1074#1085#1086#1077' '#1086#1082#1085#1086' '#1087#1088#1086#1075#1088#1072#1084#1099
        TabOrder = 0
        OnClick = cbLocGeneralClick
        SkinData.SkinSection = 'CHECKBOX'
        ImgChecked = 0
        ImgUnchecked = 0
        ShowFocus = False
      end
      object cbLocWord: TsCheckBox
        Left = 12
        Top = 48
        Width = 100
        Height = 17
        Caption = 'Microsoft Word'
        TabOrder = 1
        OnClick = cbLocGeneralClick
        SkinData.SkinSection = 'CHECKBOX'
        ImgChecked = 0
        ImgUnchecked = 0
        ShowFocus = False
      end
    end
    object groupSelectData: TsGroupBox
      Left = 8
      Top = 8
      Width = 449
      Height = 233
      Caption = #1042#1099#1073#1088#1072#1090#1100' '#1076#1072#1085#1085#1099#1077' '#1087#1086
      TabOrder = 0
      SkinData.SkinSection = 'GROUPBOX'
      object editFilter: TsComboBox
        Left = 16
        Top = 24
        Width = 417
        Height = 21
        Alignment = taLeftJustify
        SkinData.SkinSection = 'COMBOBOX'
        VerticalAlignment = taAlignTop
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = -1
        ParentFont = False
        TabOrder = 0
        OnSelect = editFilterSelect
        Items.Strings = (
          #1056#1091#1073#1088#1080#1082#1072
          #1043#1086#1088#1086#1076
          #1057#1090#1088#1072#1085#1072
          #1050#1091#1088#1072#1090#1086#1088
          #1058#1080#1087
          #1040#1082#1090#1080#1074#1085#1086#1089#1090#1100
          #1040#1082#1090#1091#1072#1083#1100#1085#1086#1089#1090#1100
          #1058#1077#1083#1077#1092#1086#1085
          #1052#1077#1081#1083
          #1057#1072#1081#1090
          #1044#1072#1090#1072' '#1076#1086#1073#1072#1074#1083#1077#1085#1080#1103
          #1044#1072#1090#1072' '#1080#1079#1084#1077#1085#1077#1085#1080#1103)
      end
      object panelDates: TsPanel
        Left = 16
        Top = 52
        Width = 420
        Height = 177
        BevelOuter = bvNone
        TabOrder = 2
        Visible = False
        SkinData.SkinSection = 'CHECKBOX'
        object lblDate: TsLabel
          Left = 0
          Top = 8
          Width = 134
          Height = 13
          Caption = #1044#1072#1090#1072' '#1076#1086#1073#1072#1074#1083#1077#1085#1080#1103' ('#1087#1077#1088#1080#1086#1076')'
          ParentFont = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
        end
        object editDate1: TsDateEdit
          Left = 159
          Top = 4
          Width = 125
          Height = 21
          AutoSize = False
          Color = clWhite
          EditMask = '!99/99/9999;1; '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 0
          BoundLabel.Layout = sclTopLeft
          SkinData.SkinSection = 'EDIT'
          GlyphMode.Blend = 0
          GlyphMode.Grayed = False
          DirectInput = False
          MaxDate = 73415.000000000000000000
          DefaultToday = True
          DialogTitle = #1059#1082#1072#1078#1080#1090#1077' '#1076#1072#1090#1091
          StartOfWeek = dowMonday
          Weekends = [dowSaturday, dowSunday]
        end
        object editDate2: TsDateEdit
          Left = 292
          Top = 4
          Width = 125
          Height = 21
          AutoSize = False
          Color = clWhite
          EditMask = '!99/99/9999;1; '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 1
          SkinData.SkinSection = 'EDIT'
          GlyphMode.Blend = 0
          GlyphMode.Grayed = False
          DirectInput = False
          MaxDate = 73415.000000000000000000
          DefaultToday = True
          DialogTitle = #1059#1082#1072#1078#1080#1090#1077' '#1076#1072#1090#1091
          StartOfWeek = dowMonday
          Weekends = [dowSaturday, dowSunday]
        end
      end
      object panelFilters: TsPanel
        Left = 16
        Top = 52
        Width = 420
        Height = 177
        BevelOuter = bvNone
        TabOrder = 1
        SkinData.SkinSection = 'CHECKBOX'
        object editFilterData: TsComboBox
          Left = 0
          Top = 4
          Width = 417
          Height = 21
          Alignment = taLeftJustify
          DropDownCount = 9
          SkinData.SkinSection = 'COMBOBOX'
          VerticalAlignment = taAlignTop
          Color = clWhite
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 15
          ItemIndex = -1
          ParentFont = False
          Sorted = True
          TabOrder = 0
          AutoDropDown = True
        end
      end
    end
    object btnOK: TsButton
      Left = 463
      Top = 217
      Width = 90
      Height = 25
      Caption = 'OK'
      TabOrder = 2
      OnClick = btnOKClick
      SkinData.SkinSection = 'BUTTON'
    end
    object btnCancel: TsButton
      Left = 559
      Top = 217
      Width = 90
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 3
      OnClick = btnCancelClick
      SkinData.SkinSection = 'BUTTON'
    end
  end
end
