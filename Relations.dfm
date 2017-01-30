object FormRelations: TFormRelations
  Left = 366
  Top = 184
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1047#1072#1082#1088#1077#1087#1080#1090#1100' '#1085#1072#1087#1088#1072#1074#1083#1077#1085#1080#1103' '#1076#1077#1103#1090#1077#1083#1100#1085#1086#1089#1090#1080' '#1079#1072' '#1088#1091#1073#1088#1080#1082#1072#1084#1080
  ClientHeight = 384
  ClientWidth = 541
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 13
  object panelRelations: TsPanel
    Left = 0
    Top = 0
    Width = 541
    Height = 384
    Align = alClient
    TabOrder = 0
    SkinData.SkinSection = 'PANEL'
    object lblID: TsLabel
      Left = 11
      Top = 352
      Width = 21
      Height = 13
      Caption = 'lblID'
      ParentFont = False
      Visible = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
    end
    object editRubrRelations: TsComboBox
      Left = 8
      Top = 12
      Width = 525
      Height = 21
      Alignment = taLeftJustify
      BoundLabel.Indent = 0
      BoundLabel.Font.Charset = DEFAULT_CHARSET
      BoundLabel.Font.Color = clWindowText
      BoundLabel.Font.Height = -11
      BoundLabel.Font.Name = 'MS Sans Serif'
      BoundLabel.Font.Style = []
      BoundLabel.Layout = sclLeft
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
      ItemIndex = -1
      ParentFont = False
      TabOrder = 0
      OnChange = editRubrRelationsChange
    end
    object sGroupBox: TsGroupBox
      Left = 8
      Top = 40
      Width = 525
      Height = 297
      TabOrder = 1
      SkinData.SkinSection = 'GROUPBOX'
      object lblAviableNapr: TsLabel
        Left = 10
        Top = 20
        Width = 202
        Height = 13
        Caption = #1044#1086#1089#1090#1091#1087#1085#1099#1077' '#1085#1072#1087#1088#1072#1074#1083#1077#1085#1080#1103' '#1076#1077#1103#1090#1077#1083#1100#1085#1086#1089#1090#1080':'
        ParentFont = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
      end
      object lblSelectedNapr: TsLabel
        Left = 10
        Top = 172
        Width = 204
        Height = 13
        Caption = #1042#1099#1073#1088#1072#1085#1085#1099#1077' '#1085#1072#1087#1088#1072#1074#1083#1077#1085#1080#1103' '#1076#1077#1103#1090#1077#1083#1100#1085#1086#1089#1090#1080':'
        ParentFont = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
      end
      object btnAdd: TsSpeedButton
        Left = 204
        Top = 144
        Width = 23
        Height = 22
        Caption = '>'
        OnClick = btnAddClick
        SkinData.SkinSection = 'SPEEDBUTTON'
      end
      object btnRemove: TsSpeedButton
        Left = 232
        Top = 144
        Width = 23
        Height = 22
        Caption = '<'
        OnClick = btnRemoveClick
        SkinData.SkinSection = 'SPEEDBUTTON'
      end
      object btnRemoveAll: TsSpeedButton
        Left = 288
        Top = 144
        Width = 23
        Height = 22
        Caption = '<<'
        OnClick = btnRemoveAllClick
        SkinData.SkinSection = 'SPEEDBUTTON'
      end
      object btnAddAll: TsSpeedButton
        Left = 260
        Top = 144
        Width = 23
        Height = 22
        Caption = '>>'
        OnClick = btnAddAllClick
        SkinData.SkinSection = 'SPEEDBUTTON'
      end
      object SGNaprAviable: TNextGrid
        Left = 11
        Top = 36
        Width = 501
        Height = 100
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Options = [goSelectFullRow]
        TabOrder = 0
        TabStop = True
        object NxTextColumn1: TNxTextColumn
          DefaultWidth = 475
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 475
        end
        object NxTextColumn2: TNxTextColumn
          Header.Caption = 'ID'
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object SGNaprSelected: TNextGrid
        Left = 10
        Top = 188
        Width = 502
        Height = 100
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Options = [goSelectFullRow]
        TabOrder = 1
        TabStop = True
        object NxTextColumn3: TNxTextColumn
          DefaultWidth = 475
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 475
        end
        object NxTextColumn4: TNxTextColumn
          Header.Caption = 'ID'
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
    end
    object BtnOK: TsButton
      Left = 315
      Top = 346
      Width = 105
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      TabOrder = 2
      OnClick = BtnOKClick
      SkinData.SkinSection = 'BUTTON'
    end
    object BtnCancel: TsButton
      Left = 427
      Top = 346
      Width = 105
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 3
      OnClick = BtnCancelClick
      SkinData.SkinSection = 'BUTTON'
    end
  end
end