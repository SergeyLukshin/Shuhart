object fLogin: TfLogin
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1040#1074#1090#1086#1088#1080#1079#1072#1094#1080#1103
  ClientHeight = 93
  ClientWidth = 238
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poDesktopCenter
  OnKeyPress = FormKeyPress
  PixelsPerInch = 96
  TextHeight = 13
  object cxGroupBox1: TcxGroupBox
    Left = 0
    Top = 0
    Align = alClient
    Style.LookAndFeel.SkinName = 'Caramel'
    StyleDisabled.LookAndFeel.SkinName = 'Caramel'
    StyleFocused.LookAndFeel.SkinName = 'Caramel'
    StyleHot.LookAndFeel.SkinName = 'Caramel'
    TabOrder = 0
    ExplicitLeft = 8
    ExplicitTop = -1
    ExplicitWidth = 185
    ExplicitHeight = 105
    Height = 93
    Width = 238
    object Label12: TLabel
      Left = 13
      Top = 21
      Width = 86
      Height = 13
      Caption = #1042#1074#1077#1076#1080#1090#1077' '#1087#1072#1088#1086#1083#1100':'
    end
    object cxPassword: TEdit
      Left = 120
      Top = 18
      Width = 105
      Height = 21
      PasswordChar = '*'
      TabOrder = 0
    end
    object cxButton1: TcxButton
      Left = 13
      Top = 56
      Width = 109
      Height = 25
      Caption = #1042#1093#1086#1076' '#1074' '#1089#1080#1089#1090#1077#1084#1091
      LookAndFeel.SkinName = 'Caramel'
      TabOrder = 1
      OnClick = cxButton1Click
    end
    object cxButton2: TcxButton
      Left = 128
      Top = 56
      Width = 97
      Height = 25
      Caption = #1047#1072#1082#1088#1099#1090#1100
      LookAndFeel.SkinName = 'Caramel'
      ModalResult = 2
      TabOrder = 2
      OnClick = cxButton2Click
    end
  end
end
