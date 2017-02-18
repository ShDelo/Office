object FormDirectoryQuery: TFormDirectoryQuery
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1059#1087#1088#1072#1074#1083#1077#1085#1080#1077' '#1076#1080#1088#1077#1082#1090#1086#1088#1080#1103#1084#1080
  ClientHeight = 190
  ClientWidth = 494
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PanelGeneral: TsPanel
    Left = 0
    Top = 0
    Width = 494
    Height = 190
    Align = alClient
    TabOrder = 0
    DesignSize = (
      494
      190)
    object Edit3: TsComboBoxEx
      Left = 15
      Top = 115
      Width = 461
      Height = 22
      BoundLabel.Active = True
      BoundLabel.Indent = 3
      BoundLabel.Caption = 'ComboBoxEx3'
      BoundLabel.Layout = sclTopLeft
      AutoCompleteOptions = [acoAutoSuggest, acoAutoAppend]
      ItemsEx = <>
      Anchors = [akLeft, akTop, akRight]
      Color = clWhite
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ItemIndex = -1
      ParentFont = False
      TabOrder = 0
    end
    object BtnCancel: TsButton
      Left = 376
      Top = 155
      Width = 100
      Height = 25
      Anchors = [akBottom]
      Caption = #1054#1090#1084#1077#1085#1072
      ModalResult = 2
      TabOrder = 1
    end
    object BtnOK: TsButton
      Left = 270
      Top = 155
      Width = 100
      Height = 25
      Anchors = [akBottom]
      Caption = 'OK'
      TabOrder = 2
      OnClick = BtnOKClick
    end
    object Edit1: TsComboBoxEx
      Left = 15
      Top = 25
      Width = 461
      Height = 22
      BoundLabel.Active = True
      BoundLabel.Indent = 3
      BoundLabel.Caption = 'ComboBoxEx1'
      BoundLabel.Layout = sclTopLeft
      AutoCompleteOptions = [acoAutoSuggest]
      ItemsEx = <>
      Anchors = [akLeft, akTop, akRight]
      Color = clWhite
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ItemIndex = -1
      ParentFont = False
      TabOrder = 3
    end
    object Edit2: TsComboBoxEx
      Left = 15
      Top = 70
      Width = 461
      Height = 22
      BoundLabel.Active = True
      BoundLabel.Indent = 3
      BoundLabel.Caption = 'ComboBoxEx2'
      BoundLabel.Layout = sclTopLeft
      AutoCompleteOptions = [acoAutoSuggest]
      ItemsEx = <>
      Anchors = [akLeft, akTop, akRight]
      Color = clWhite
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ItemIndex = -1
      ParentFont = False
      TabOrder = 4
    end
  end
end
