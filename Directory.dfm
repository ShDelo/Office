object FormDirectory: TFormDirectory
  Left = 370
  Top = 159
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1059#1087#1088#1072#1074#1083#1077#1085#1080#1077' '#1076#1080#1088#1077#1082#1090#1086#1088#1080#1103#1084#1080
  ClientHeight = 365
  ClientWidth = 653
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl: TsPageControl
    Left = 0
    Top = 0
    Width = 653
    Height = 365
    ActivePage = tabCurator
    Align = alClient
    TabOrder = 0
    TabStop = False
    ShowFocus = False
    TabPadding = 3
    SkinData.SkinSection = 'PAGECONTROL'
    object tabCurator: TsTabSheet
      Caption = #1050#1091#1088#1072#1090#1086#1088#1099
      object panelCurator: TsPanel
        Left = 0
        Top = 0
        Width = 645
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnCuratorCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Align = alLeft
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCuratorEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Align = alLeft
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCuratorDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Align = alLeft
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGCurator: TNextGrid
        Left = 0
        Top = 46
        Width = 645
        Height = 291
        Touch.InteractiveGestures = [igPan, igPressAndTap]
        Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Caption = ''
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn1: TNxTextColumn
          DefaultWidth = 625
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 625
        end
        object NxTextColumn2: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = 'ID'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editCurator: TsEdit
        Left = 0
        Top = 25
        Width = 645
        Height = 21
        Align = alTop
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        SkinData.SkinSection = 'EDIT'
      end
    end
    object tabRubr: TsTabSheet
      Caption = #1056#1091#1073#1088#1080#1082#1080
      object panelRubr: TsPanel
        Left = 0
        Top = 0
        Width = 645
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnRubrCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Align = alLeft
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnRubrEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Align = alLeft
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnRubrDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Align = alLeft
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGRubr: TNextGrid
        Left = 0
        Top = 46
        Width = 645
        Height = 291
        Touch.InteractiveGestures = [igPan, igPressAndTap]
        Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Caption = ''
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn3: TNxTextColumn
          DefaultWidth = 625
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 625
        end
        object NxTextColumn4: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = 'ID'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editRubr: TsEdit
        Left = 0
        Top = 25
        Width = 645
        Height = 21
        Align = alTop
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        SkinData.SkinSection = 'EDIT'
      end
    end
    object tabFirmType: TsTabSheet
      Caption = #1058#1080#1087#1099' '#1092#1080#1088#1084
      object panelFirmType: TsPanel
        Left = 0
        Top = 0
        Width = 645
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnFirmTypeCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Align = alLeft
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnFirmTypeEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Align = alLeft
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnFirmTypeDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Align = alLeft
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGFirmType: TNextGrid
        Left = 0
        Top = 46
        Width = 645
        Height = 291
        Touch.InteractiveGestures = [igPan, igPressAndTap]
        Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Caption = ''
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn5: TNxTextColumn
          DefaultWidth = 625
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 625
        end
        object NxTextColumn6: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = 'ID'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editFirmType: TsEdit
        Left = 0
        Top = 25
        Width = 645
        Height = 21
        Align = alTop
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        SkinData.SkinSection = 'EDIT'
      end
    end
    object tabNapr: TsTabSheet
      Caption = #1042#1080#1076#1099' '#1076#1077#1103#1090#1077#1083#1100#1085#1086#1089#1090#1080
      object panelNapr: TsPanel
        Left = 0
        Top = 0
        Width = 645
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnNaprCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Align = alLeft
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnNaprEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Align = alLeft
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnNaprDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Align = alLeft
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGNapr: TNextGrid
        Left = 0
        Top = 46
        Width = 645
        Height = 291
        Touch.InteractiveGestures = [igPan, igPressAndTap]
        Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Caption = ''
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn7: TNxTextColumn
          DefaultWidth = 625
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 625
        end
        object NxTextColumn8: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = 'ID'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editNapr: TsEdit
        Left = 0
        Top = 25
        Width = 645
        Height = 21
        Align = alTop
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        SkinData.SkinSection = 'EDIT'
      end
    end
    object tabOfficeType: TsTabSheet
      Caption = #1058#1080#1087#1099' '#1072#1076#1088#1077#1089#1086#1074
      object panelOfficeType: TsPanel
        Left = 0
        Top = 0
        Width = 645
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnOfficeTypeCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Align = alLeft
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnOfficeTypeEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Align = alLeft
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnOfficeTypeDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Align = alLeft
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGOfficeType: TNextGrid
        Left = 0
        Top = 46
        Width = 645
        Height = 291
        Touch.InteractiveGestures = [igPan, igPressAndTap]
        Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Caption = ''
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn9: TNxTextColumn
          DefaultWidth = 625
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 625
        end
        object NxTextColumn10: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = 'ID'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editOfficeType: TsEdit
        Left = 0
        Top = 25
        Width = 645
        Height = 21
        Align = alTop
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        SkinData.SkinSection = 'EDIT'
      end
    end
    object tabCountry: TsTabSheet
      Caption = #1057#1090#1088#1072#1085#1099
      object panelCountry: TsPanel
        Left = 0
        Top = 0
        Width = 645
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnCountryCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Align = alLeft
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCountryEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Align = alLeft
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCountryDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Align = alLeft
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGCountry: TNextGrid
        Left = 0
        Top = 46
        Width = 645
        Height = 291
        Touch.InteractiveGestures = [igPan, igPressAndTap]
        Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Caption = ''
        Options = [goDisableColumnMoving, goHeader, goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn11: TNxTextColumn
          DefaultWidth = 310
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1057#1090#1088#1072#1085#1072' ('#1088#1091#1089')'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          Options = [coCanClick, coCanInput, coCanSort, coFixedSize, coPublicUsing, coShowTextFitHint]
          ParentFont = False
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 310
        end
        object NxTextColumn12: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = 'ID'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
        object NxTextColumn22: TNxTextColumn
          DefaultWidth = 310
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1057#1090#1088#1072#1085#1072' ('#1091#1082#1088')'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          Options = [coCanClick, coCanInput, coCanSort, coFixedSize, coPublicUsing, coShowTextFitHint]
          ParentFont = False
          Position = 2
          SortType = stAlphabetic
          Width = 310
        end
      end
      object editCountry: TsEdit
        Left = 0
        Top = 25
        Width = 645
        Height = 21
        Align = alTop
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        SkinData.SkinSection = 'EDIT'
      end
    end
    object tabRegion: TsTabSheet
      Caption = #1054#1073#1083#1072#1089#1090#1100
      object panelRegion: TsPanel
        Left = 0
        Top = 0
        Width = 645
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnRegionCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Align = alLeft
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnRegionEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Align = alLeft
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnRegionDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Align = alLeft
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object editRegion: TsEdit
        Left = 0
        Top = 25
        Width = 645
        Height = 21
        Align = alTop
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        SkinData.SkinSection = 'EDIT'
      end
      object SGRegion: TNextGrid
        Left = 0
        Top = 46
        Width = 645
        Height = 291
        Touch.InteractiveGestures = [igPan, igPressAndTap]
        Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Caption = ''
        Options = [goDisableColumnMoving, goHeader, goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn17: TNxTextColumn
          DefaultWidth = 310
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1054#1073#1083#1072#1089#1090#1100' ('#1088#1091#1089')'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          Options = [coCanClick, coCanInput, coCanSort, coFixedSize, coPublicUsing, coShowTextFitHint]
          ParentFont = False
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 310
        end
        object NxTextColumn18: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = 'ID'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
        object NxTextColumn23: TNxTextColumn
          DefaultWidth = 310
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1054#1073#1083#1072#1089#1090#1100' ('#1091#1082#1088')'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          Options = [coCanClick, coCanInput, coCanSort, coFixedSize, coPublicUsing, coShowTextFitHint]
          ParentFont = False
          Position = 2
          SortType = stAlphabetic
          Width = 310
        end
      end
    end
    object tabCity: TsTabSheet
      Caption = #1043#1086#1088#1086#1076#1072
      object panelCity: TsPanel
        Left = 0
        Top = 0
        Width = 645
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnCityCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Align = alLeft
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCityEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Align = alLeft
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnCityDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Align = alLeft
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object SGCity: TNextGrid
        Left = 0
        Top = 46
        Width = 645
        Height = 291
        Touch.InteractiveGestures = [igPan, igPressAndTap]
        Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Caption = ''
        Options = [goDisableColumnMoving, goHeader, goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn13: TNxTextColumn
          DefaultWidth = 200
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1043#1086#1088#1086#1076' ('#1088#1091#1089')'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          Options = [coCanClick, coCanInput, coCanSort, coFixedSize, coPublicUsing, coShowTextFitHint]
          ParentFont = False
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 200
        end
        object NxTextColumn14: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = 'ID'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
        object NxTextColumn19: TNxTextColumn
          DefaultWidth = 200
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1043#1086#1088#1086#1076' ('#1091#1082#1088')'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          Options = [coCanClick, coCanInput, coCanSort, coFixedSize, coPublicUsing, coShowTextFitHint]
          ParentFont = False
          Position = 2
          SortType = stAlphabetic
          Width = 200
        end
        object NxTextColumn20: TNxTextColumn
          DefaultWidth = 225
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1054#1073#1083#1072#1089#1090#1100
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          Options = [coCanClick, coCanInput, coCanSort, coFixedSize, coPublicUsing, coShowTextFitHint]
          ParentFont = False
          Position = 3
          SortType = stAlphabetic
          Width = 225
        end
        object NxTextColumn21: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = 'ID_OBLAST'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 4
          SortType = stAlphabetic
          Visible = False
        end
      end
      object editCity: TsEdit
        Left = 0
        Top = 25
        Width = 645
        Height = 21
        Align = alTop
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        SkinData.SkinSection = 'EDIT'
      end
    end
    object tabPhoneType: TsTabSheet
      Caption = #1058#1080#1087#1099' '#1090#1077#1083#1077#1092#1086#1085#1086#1074
      object panelPhoneType: TsPanel
        Left = 0
        Top = 0
        Width = 645
        Height = 25
        Align = alTop
        TabOrder = 0
        SkinData.SkinSection = 'PANEL'
        object btnPhoneTypeCreate: TsSpeedButton
          Left = 1
          Top = 1
          Width = 56
          Height = 23
          Align = alLeft
          Caption = #1057#1086#1079#1076#1072#1090#1100
          OnClick = btnCreateClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnPhoneTypeEdit: TsSpeedButton
          Left = 57
          Top = 1
          Width = 84
          Height = 23
          Align = alLeft
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          OnClick = btnEditClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
        object btnPhoneTypeDelete: TsSpeedButton
          Left = 141
          Top = 1
          Width = 58
          Height = 23
          Align = alLeft
          Caption = #1059#1076#1072#1083#1080#1090#1100
          OnClick = btnDeleteClick
          SkinData.SkinSection = 'SPEEDBUTTON_SMALL'
        end
      end
      object editPhoneType: TsEdit
        Left = 0
        Top = 25
        Width = 645
        Height = 21
        Align = alTop
        Color = clWhite
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = editCuratorChange
        SkinData.SkinSection = 'EDIT'
      end
      object SGPhoneType: TNextGrid
        Left = 0
        Top = 46
        Width = 645
        Height = 291
        Touch.InteractiveGestures = [igPan, igPressAndTap]
        Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
        Align = alClient
        AppearanceOptions = [aoHideFocus, aoHighlightSlideCells]
        Caption = ''
        Options = [goSelectFullRow]
        TabOrder = 2
        TabStop = True
        object NxTextColumn15: TNxTextColumn
          DefaultWidth = 625
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 0
          Sorted = True
          SortType = stAlphabetic
          Width = 625
        end
        object NxTextColumn16: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Header.Caption = 'ID'
          Header.Font.Charset = DEFAULT_CHARSET
          Header.Font.Color = clWindowText
          Header.Font.Height = -11
          Header.Font.Name = 'Tahoma'
          Header.Font.Style = []
          ParentFont = False
          Position = 1
          SortType = stAlphabetic
          Visible = False
        end
      end
    end
  end
end
