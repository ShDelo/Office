object FormDublicate: TFormDublicate
  Left = 336
  Top = 187
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1044#1091#1073#1083#1080#1082#1072#1090#1099' '#1092#1080#1088#1084
  ClientHeight = 335
  ClientWidth = 381
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
  object SGDublicate: TNextGrid
    Left = 4
    Top = 4
    Width = 373
    Height = 289
    Touch.InteractiveGestures = [igPan, igPressAndTap]
    Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
    AppearanceOptions = [aoHideFocus, aoHideSelection, aoHighlightSlideCells]
    Caption = ''
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goSelectFullRow]
    RowSize = 22
    ParentFont = False
    TabOrder = 0
    TabStop = True
    OnDblClick = SGDublicateDblClick
    object NxImageColumn1: TNxImageColumn
      DefaultValue = '0'
      DefaultWidth = 24
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Header.Font.Charset = DEFAULT_CHARSET
      Header.Font.Color = clWindowText
      Header.Font.Height = -11
      Header.Font.Name = 'Tahoma'
      Header.Font.Style = []
      Options = [coCanClick, coCanInput, coCanSort, coDontHighlight, coPublicUsing]
      ParentFont = False
      Position = 0
      SortType = stNumeric
      Width = 24
      Images = FormMain.ImageList1
    end
    object NxTextColumn1: TNxTextColumn
      DefaultWidth = 325
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Header.Caption = #1053#1072#1079#1074#1072#1085#1080#1077
      Header.Font.Charset = DEFAULT_CHARSET
      Header.Font.Color = clWindowText
      Header.Font.Height = -11
      Header.Font.Name = 'Tahoma'
      Header.Font.Style = []
      ParentFont = False
      Position = 1
      Sorted = True
      SortType = stAlphabetic
      Width = 325
    end
    object NxTextColumn2: TNxTextColumn
      DefaultWidth = 30
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Header.Caption = 'ID'
      Header.Font.Charset = DEFAULT_CHARSET
      Header.Font.Color = clWindowText
      Header.Font.Height = -11
      Header.Font.Name = 'Tahoma'
      Header.Font.Style = []
      InputCaption = #1042#1099#1073#1077#1088#1080#1090#1077' '#1088#1091#1073#1088#1080#1082#1091'...'
      ParentFont = False
      Position = 2
      SortType = stAlphabetic
      Visible = False
      Width = 30
    end
  end
  object BtnOK: TsButton
    Left = 159
    Top = 302
    Width = 105
    Height = 25
    Caption = #1055#1088#1086#1076#1086#1083#1078#1080#1090#1100
    TabOrder = 1
    SkinData.SkinSection = 'BUTTON'
  end
  object BtnCancel: TsButton
    Left = 271
    Top = 302
    Width = 105
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    TabOrder = 2
    OnClick = BtnCancelClick
    SkinData.SkinSection = 'BUTTON'
  end
end
