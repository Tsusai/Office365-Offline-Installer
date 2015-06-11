object Form1: TForm1
  Left = 212
  Top = 185
  Caption = 'Office 2013/365 Installer'
  ClientHeight = 157
  ClientWidth = 369
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Scaled = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 8
    Top = 127
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
    Height = 115
    Caption = 'Version'
    ItemIndex = 0
    Items.Strings = (
      'Office 2013 Home && Student'
      'Office 2013 Home && Buisness'
      'Office   365 Home Premium'
      'Office   365 Pro Plus'
      'Office   365 Buisness'
      'Office   365 Small Buisness Premium')
    TabOrder = 1
    OnClick = RadioGroup1Click
  end
  object CheckListBox1: TCheckListBox
    Left = 223
    Top = 8
    Width = 138
    Height = 144
    Color = clBtnFace
    Enabled = False
    ItemHeight = 13
    Items.Strings = (
      'Excel'
      'OneNote'
      'PowerPoint'
      'Word'
      'Outlook'
      'Publisher'
      'Access'
      'OneDrive'
      'InfoPath'
      'Lync')
    TabOrder = 2
  end
  object CustomCheck: TCheckBox
    Left = 111
    Top = 131
    Width = 97
    Height = 17
    Caption = 'Customize Install'
    TabOrder = 3
    OnClick = CustomCheckClick
  end
end
