object FormReport: TFormReport
  Left = 340
  Top = 137
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1043#1077#1085#1077#1088#1072#1094#1080#1103' '#1086#1090#1095#1077#1090#1072
  ClientHeight = 403
  ClientWidth = 702
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
    Width = 702
    Height = 403
    Align = alClient
    TabOrder = 0
    SkinData.SkinSection = 'PANEL'
    object ProgressBar: TsGauge
      Left = 20
      Top = 366
      Width = 301
      Height = 30
      Visible = False
      SkinData.SkinSection = 'GAUGE'
      ForeColor = clBlack
      Suffix = '%'
    end
    object lblStatus: TsLabel
      Left = 324
      Top = 375
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
    object groupSelectData: TsGroupBox
      Left = 8
      Top = 12
      Width = 453
      Height = 345
      Caption = #1042#1099#1073#1088#1072#1090#1100' '#1076#1072#1085#1085#1099#1077' '#1087#1086
      TabOrder = 0
      SkinData.SkinSection = 'GROUPBOX'
      object editSelect9: TsComboBox
        Left = 20
        Top = 216
        Width = 140
        Height = 21
        Alignment = taLeftJustify
        BoundLabel.Indent = 5
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'COMBOBOX'
        Style = csDropDownList
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = 0
        ParentFont = False
        TabOrder = 8
        Text = '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
        OnChange = editSelect1Change
        Items.Strings = (
          '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
          #1040#1082#1090#1080#1074#1085#1086#1089#1090#1100
          #1040#1082#1090#1091#1072#1083#1100#1085#1086#1089#1090#1100
          #1057#1090#1088#1072#1085#1072
          #1043#1086#1088#1086#1076
          #1050#1091#1088#1072#1090#1086#1088
          #1056#1091#1073#1088#1080#1082#1072
          #1058#1080#1087)
      end
      object editSelect7: TsComboBox
        Left = 20
        Top = 168
        Width = 140
        Height = 21
        Alignment = taLeftJustify
        BoundLabel.Indent = 5
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'COMBOBOX'
        Style = csDropDownList
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = 0
        ParentFont = False
        TabOrder = 6
        Text = '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
        OnChange = editSelect1Change
        Items.Strings = (
          '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
          #1040#1082#1090#1080#1074#1085#1086#1089#1090#1100
          #1040#1082#1090#1091#1072#1083#1100#1085#1086#1089#1090#1100
          #1057#1090#1088#1072#1085#1072
          #1043#1086#1088#1086#1076
          #1050#1091#1088#1072#1090#1086#1088
          #1056#1091#1073#1088#1080#1082#1072
          #1058#1080#1087)
      end
      object editSelect5: TsComboBox
        Left = 20
        Top = 120
        Width = 140
        Height = 21
        Alignment = taLeftJustify
        BoundLabel.Indent = 5
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'COMBOBOX'
        Style = csDropDownList
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = 0
        ParentFont = False
        TabOrder = 4
        Text = '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
        OnChange = editSelect1Change
        Items.Strings = (
          '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
          #1040#1082#1090#1080#1074#1085#1086#1089#1090#1100
          #1040#1082#1090#1091#1072#1083#1100#1085#1086#1089#1090#1100
          #1057#1090#1088#1072#1085#1072
          #1043#1086#1088#1086#1076
          #1050#1091#1088#1072#1090#1086#1088
          #1056#1091#1073#1088#1080#1082#1072
          #1058#1080#1087)
      end
      object editSelect3: TsComboBox
        Left = 20
        Top = 72
        Width = 140
        Height = 21
        Alignment = taLeftJustify
        BoundLabel.Indent = 5
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'COMBOBOX'
        Style = csDropDownList
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = 0
        ParentFont = False
        TabOrder = 2
        Text = '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
        OnChange = editSelect1Change
        Items.Strings = (
          '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
          #1040#1082#1090#1080#1074#1085#1086#1089#1090#1100
          #1040#1082#1090#1091#1072#1083#1100#1085#1086#1089#1090#1100
          #1057#1090#1088#1072#1085#1072
          #1043#1086#1088#1086#1076
          #1050#1091#1088#1072#1090#1086#1088
          #1056#1091#1073#1088#1080#1082#1072
          #1058#1080#1087)
      end
      object editSelect1: TsComboBox
        Left = 20
        Top = 24
        Width = 140
        Height = 21
        Alignment = taLeftJustify
        BoundLabel.Indent = 5
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'COMBOBOX'
        Style = csDropDownList
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = 0
        ParentFont = False
        TabOrder = 0
        Text = '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
        OnChange = editSelect1Change
        Items.Strings = (
          '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
          #1040#1082#1090#1080#1074#1085#1086#1089#1090#1100
          #1040#1082#1090#1091#1072#1083#1100#1085#1086#1089#1090#1100
          #1057#1090#1088#1072#1085#1072
          #1043#1086#1088#1086#1076
          #1050#1091#1088#1072#1090#1086#1088
          #1056#1091#1073#1088#1080#1082#1072
          #1058#1080#1087)
      end
      object editDateAdded1: TsDateEdit
        Left = 170
        Top = 264
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
        TabOrder = 11
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
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
      object editDateAdded2: TsDateEdit
        Left = 305
        Top = 264
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
        TabOrder = 12
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
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
      object editDateEdited1: TsDateEdit
        Left = 170
        Top = 304
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
        TabOrder = 14
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
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
      object editDateEdited2: TsDateEdit
        Left = 305
        Top = 304
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
        TabOrder = 15
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
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
      object cbDateAdded: TsCheckBox
        Left = 18
        Top = 267
        Width = 117
        Height = 19
        TabStop = False
        Caption = #1044#1072#1090#1072' '#1076#1086#1073#1072#1074#1083#1077#1085#1080#1103
        TabOrder = 10
        SkinData.SkinSection = 'CHECKBOX'
        ImgChecked = 0
        ImgUnchecked = 0
        ShowFocus = False
      end
      object cbDateEdited: TsCheckBox
        Left = 18
        Top = 307
        Width = 140
        Height = 19
        TabStop = False
        Caption = #1044#1072#1090#1072' '#1088#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1085#1080#1103
        TabOrder = 13
        SkinData.SkinSection = 'CHECKBOX'
        ImgChecked = 0
        ImgUnchecked = 0
        ShowFocus = False
      end
      object editSelect2: TsComboBox
        Left = 170
        Top = 24
        Width = 260
        Height = 21
        Alignment = taLeftJustify
        BoundLabel.Indent = 5
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'COMBOBOX'
        Style = csDropDownList
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = 0
        ParentFont = False
        TabOrder = 1
        Text = '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
        Items.Strings = (
          '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
          #1044#1072
          #1053#1077#1090)
      end
      object editSelect4: TsComboBox
        Left = 170
        Top = 72
        Width = 260
        Height = 21
        Alignment = taLeftJustify
        BoundLabel.Indent = 5
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'COMBOBOX'
        Style = csDropDownList
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = 0
        ParentFont = False
        TabOrder = 3
        Text = '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
        Items.Strings = (
          '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>')
      end
      object editSelect6: TsComboBox
        Left = 170
        Top = 120
        Width = 260
        Height = 21
        Alignment = taLeftJustify
        BoundLabel.Indent = 5
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'COMBOBOX'
        Style = csDropDownList
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = 0
        ParentFont = False
        TabOrder = 5
        Text = '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
        Items.Strings = (
          '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>')
      end
      object editSelect8: TsComboBox
        Left = 170
        Top = 168
        Width = 260
        Height = 21
        Alignment = taLeftJustify
        BoundLabel.Indent = 5
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'COMBOBOX'
        Style = csDropDownList
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = 0
        ParentFont = False
        TabOrder = 7
        Text = '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
        Items.Strings = (
          '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>')
      end
      object editSelect10: TsComboBox
        Left = 170
        Top = 216
        Width = 260
        Height = 21
        Alignment = taLeftJustify
        BoundLabel.Indent = 5
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclTopLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'COMBOBOX'
        Style = csDropDownList
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 15
        ItemIndex = 0
        ParentFont = False
        TabOrder = 9
        Text = '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>'
        Items.Strings = (
          '<'#1085#1077#1090' '#1074#1099#1076#1077#1083#1077#1085#1080#1103'>')
      end
    end
    object groupFormatData: TsGroupBox
      Left = 466
      Top = 142
      Width = 225
      Height = 215
      Caption = #1060#1086#1088#1084#1072#1090#1080#1088#1086#1074#1072#1085#1080#1077' '#1076#1086#1082#1091#1084#1077#1085#1090#1072
      TabOrder = 2
      SkinData.SkinSection = 'GROUPBOX'
      DesignSize = (
        225
        215)
      object editFormatDoc: TsCheckListBox
        Left = 12
        Top = 21
        Width = 200
        Height = 184
        AutoComplete = False
        Anchors = [akLeft, akTop, akRight, akBottom]
        BorderStyle = bsSingle
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 18
        Items.Strings = (
          #1040#1076#1088#1077#1089
          #1040#1082#1090#1080#1074#1085#1086#1089#1090#1100
          #1040#1082#1090#1091#1072#1083#1100#1085#1086#1089#1090#1100
          #1050#1091#1088#1072#1090#1086#1088
          #1056#1091#1073#1088#1080#1082#1072
          #1058#1080#1087
          #1044#1077#1103#1090#1077#1083#1100#1085#1086#1089#1090#1100
          #1057#1072#1081#1090
          #1052#1077#1081#1083
          #1050#1086#1085#1090#1072#1082#1090#1085#1086#1077' '#1083#1080#1094#1086)
        ParentFont = False
        TabOrder = 0
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
        SkinData.SkinSection = 'EDIT'
      end
    end
    object btnOK: TsButton
      Left = 503
      Top = 368
      Width = 90
      Height = 25
      Caption = 'OK'
      TabOrder = 4
      OnClick = btnOKClick
      SkinData.SkinSection = 'BUTTON'
    end
    object btnCancel: TsButton
      Left = 599
      Top = 368
      Width = 90
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 3
      OnClick = btnCancelClick
      SkinData.SkinSection = 'BUTTON'
    end
    object groupLocation: TsGroupBox
      Left = 466
      Top = 12
      Width = 225
      Height = 123
      Caption = #1042#1099#1074#1077#1089#1090#1080' '#1076#1072#1085#1085#1099#1077' '#1085#1072
      TabOrder = 1
      SkinData.SkinSection = 'GROUPBOX'
      object cbLocGeneral: TsCheckBox
        Left = 12
        Top = 24
        Width = 151
        Height = 19
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
        Height = 19
        Caption = 'Microsoft Word'
        TabOrder = 1
        OnClick = cbLocWordClick
        SkinData.SkinSection = 'CHECKBOX'
        ImgChecked = 0
        ImgUnchecked = 0
        ShowFocus = False
      end
      object cbLocExcel_List: TsCheckBox
        Left = 12
        Top = 72
        Width = 159
        Height = 19
        Caption = 'Microsoft Excel ('#1088#1072#1089#1089#1099#1083#1082#1072')'
        TabOrder = 2
        OnClick = cbLocExcel_ListClick
        SkinData.SkinSection = 'CHECKBOX'
        ImgChecked = 0
        ImgUnchecked = 0
        ShowFocus = False
      end
      object cbLocExcel_Report: TsCheckBox
        Left = 12
        Top = 96
        Width = 136
        Height = 19
        Caption = 'Microsoft Excel ('#1086#1090#1095#1077#1090')'
        TabOrder = 3
        OnClick = cbLocExcel_ReportClick
        SkinData.SkinSection = 'CHECKBOX'
        ImgChecked = 0
        ImgUnchecked = 0
        ShowFocus = False
      end
    end
  end
end
