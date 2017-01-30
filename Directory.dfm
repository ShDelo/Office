object FormDirectory: TFormDirectory
  Left = 370
  Top = 159
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1059#1087#1088#1072#1074#1083#1077#1085#1080#1077' '#1076#1080#1088#1077#1082#1090#1086#1088#1080#1103#1084#1080
  ClientHeight = 365
  ClientWidth = 604
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl: TsPageControl
    Left = 0
    Top = 0
    Width = 604
    Height = 365
    ActivePage = tabCurator
    Align = alClient
    TabOrder = 0
    TabStop = False
    ShowFocus = False
    SkinData.SkinSection = 'PAGECONTROL'
    object tabCurator: TsTabSheet
      Caption = #1050#1091#1088#1072#1090#1086#1088#1099
      SkinData.CustomColor = False
      SkinData.CustomFont = False
      object panelCurator: TsPanel
        Left = 0
        Top = 0
        Width = 596
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnCuratorCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCuratorEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCuratorDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGCurator: TNextGrid
        Left = 0
        Top = 46
        Width = 596
        Height = 291
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn1: TNxTextColumn
          DefaultWidth = 575
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 575
        end
        object NxTextColumn2: TNxTextColumn
          Header.Caption = 'ID'
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editCurator: TsEdit
        Left = 0
        Top = 25
        Width = 596
        Height = 21
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        Align = alTop
        SkinData.SkinSection = 'EDIT'
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
      end
    end
    object tabRubr: TsTabSheet
      Caption = #1056#1091#1073#1088#1080#1082#1080
      SkinData.CustomColor = False
      SkinData.CustomFont = False
      object panelRubr: TsPanel
        Left = 0
        Top = 0
        Width = 596
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnRubrCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnRubrEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnRubrDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGRubr: TNextGrid
        Left = 0
        Top = 46
        Width = 596
        Height = 291
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn3: TNxTextColumn
          DefaultWidth = 575
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 575
        end
        object NxTextColumn4: TNxTextColumn
          Header.Caption = 'ID'
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editRubr: TsEdit
        Left = 0
        Top = 25
        Width = 596
        Height = 21
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        Align = alTop
        SkinData.SkinSection = 'EDIT'
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
      end
    end
    object tabFirmType: TsTabSheet
      Caption = #1058#1080#1087#1099' '#1092#1080#1088#1084
      SkinData.CustomColor = False
      SkinData.CustomFont = False
      object panelFirmType: TsPanel
        Left = 0
        Top = 0
        Width = 596
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnFirmTypeCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnFirmTypeEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnFirmTypeDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGFirmType: TNextGrid
        Left = 0
        Top = 46
        Width = 596
        Height = 291
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn5: TNxTextColumn
          DefaultWidth = 575
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 575
        end
        object NxTextColumn6: TNxTextColumn
          Header.Caption = 'ID'
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editFirmType: TsEdit
        Left = 0
        Top = 25
        Width = 596
        Height = 21
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        Align = alTop
        SkinData.SkinSection = 'EDIT'
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
      end
    end
    object tabNapr: TsTabSheet
      Caption = #1042#1080#1076#1099' '#1076#1077#1103#1090#1077#1083#1100#1085#1086#1089#1090#1080
      SkinData.CustomColor = False
      SkinData.CustomFont = False
      object panelNapr: TsPanel
        Left = 0
        Top = 0
        Width = 596
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnNaprCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnNaprEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnNaprDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGNapr: TNextGrid
        Left = 0
        Top = 46
        Width = 596
        Height = 291
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn7: TNxTextColumn
          DefaultWidth = 575
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 575
        end
        object NxTextColumn8: TNxTextColumn
          Header.Caption = 'ID'
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editNapr: TsEdit
        Left = 0
        Top = 25
        Width = 596
        Height = 21
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        Align = alTop
        SkinData.SkinSection = 'EDIT'
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
      end
    end
    object tabOfficeType: TsTabSheet
      Caption = #1058#1080#1087#1099' '#1072#1076#1088#1077#1089#1086#1074
      SkinData.CustomColor = False
      SkinData.CustomFont = False
      object panelOfficeType: TsPanel
        Left = 0
        Top = 0
        Width = 596
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnOfficeTypeCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnOfficeTypeEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnOfficeTypeDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGOfficeType: TNextGrid
        Left = 0
        Top = 46
        Width = 596
        Height = 291
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn9: TNxTextColumn
          DefaultWidth = 575
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 575
        end
        object NxTextColumn10: TNxTextColumn
          Header.Caption = 'ID'
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editOfficeType: TsEdit
        Left = 0
        Top = 25
        Width = 596
        Height = 21
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        Align = alTop
        SkinData.SkinSection = 'EDIT'
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
      end
    end
    object tabCountry: TsTabSheet
      Caption = #1057#1090#1088#1072#1085#1099
      SkinData.CustomColor = False
      SkinData.CustomFont = False
      object panelCountry: TsPanel
        Left = 0
        Top = 0
        Width = 596
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnCountryCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCountryEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCountryDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGCountry: TNextGrid
        Left = 0
        Top = 46
        Width = 596
        Height = 291
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn11: TNxTextColumn
          DefaultWidth = 575
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 575
        end
        object NxTextColumn12: TNxTextColumn
          Header.Caption = 'ID'
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editCountry: TsEdit
        Left = 0
        Top = 25
        Width = 596
        Height = 21
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        Align = alTop
        SkinData.SkinSection = 'EDIT'
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
      end
    end
    object tabCity: TsTabSheet
      Caption = #1043#1086#1088#1086#1076#1072
      SkinData.CustomColor = False
      SkinData.CustomFont = False
      object panelCity: TsPanel
        Left = 0
        Top = 0
        Width = 596
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnCityCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCityEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCityDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGCity: TNextGrid
        Left = 0
        Top = 46
        Width = 596
        Height = 291
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn13: TNxTextColumn
          DefaultWidth = 575
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 575
        end
        object NxTextColumn14: TNxTextColumn
          Header.Caption = 'ID'
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editCity: TsEdit
        Left = 0
        Top = 25
        Width = 596
        Height = 21
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        Align = alTop
        SkinData.SkinSection = 'EDIT'
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
      end
    end
    object sTabSheet1: TsTabSheet
      Caption = #1058#1080#1087#1099' '#1090#1077#1083#1077#1092#1086#1085#1086#1074
      SkinData.CustomColor = False
      SkinData.CustomFont = False
      object panelPhoneType: TsPanel
        Left = 0
        Top = 0
        Width = 596
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnPhoneTypeCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnPhoneTypeEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnPhoneTypeDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          Align = alLeft
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object editPhoneType: TsEdit
        Left = 0
        Top = 25
        Width = 596
        Height = 21
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        Align = alTop
        SkinData.SkinSection = 'EDIT'
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
      end
      object SGPhoneType: TNextGrid
        Left = 0
        Top = 46
        Width = 596
        Height = 291
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn15: TNxTextColumn
          DefaultWidth = 575
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 575
        end
        object NxTextColumn16: TNxTextColumn
          Header.Caption = 'ID'
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
    end
  end
end