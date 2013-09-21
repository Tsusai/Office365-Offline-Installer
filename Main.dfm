object Form1: TForm1
  Left = 212
  Top = 185
  Width = 342
  Height = 159
  Caption = 'Office 2013/365 Installer'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 224
    Top = 88
    Width = 97
    Height = 25
    Caption = 'Install'
    TabOrder = 0
    OnClick = Button1Click
  end
  object RadioGroup1: TRadioGroup
    Left = 8
    Top = 6
    Width = 209
    Height = 107
    Caption = 'Version'
    ItemIndex = 0
    Items.Strings = (
      'Office 2013 Home && Student'
      'Office 2013 Home && Buisness'
      'Office 365 Home Premium'
      'Office 365 Pro Plus'
      'Office 365 Small Buisness Premium')
    TabOrder = 1
  end
  object VistaAltFix1: TVistaAltFix
    Left = 224
    Top = 8
  end
end
