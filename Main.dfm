object Mainform: TMainform
  Left = 212
  Top = 185
  Caption = 'Office Offline Installer'
  ClientHeight = 182
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
  object RunBtn: TButton
    Left = 8
    Top = 149
    Width = 89
    Height = 25
    Caption = 'Run'
    TabOrder = 0
    OnClick = RunBtnClick
  end
  object VersionList: TRadioGroup
    Left = 8
    Top = 6
    Width = 209
    Height = 115
    Caption = 'Version'
    ItemIndex = 0
    Items.Strings = (
      'Office Home && Student'
      'Office Home && Business'
      'Office 365 Home Premium'
      'Office 365 Pro Plus'
      'Office 365 Business'
      'Office 365 Small Business Premium')
    TabOrder = 1
    OnClick = VersionListClick
  end
  object SoftwareList: TCheckListBox
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
    Left = 103
    Top = 153
    Width = 97
    Height = 17
    Caption = 'Customize Install'
    TabOrder = 3
    OnClick = CustomCheckClick
  end
  object VerBox: TComboBox
    Left = 8
    Top = 126
    Width = 89
    Height = 21
    Style = csDropDownList
    ItemIndex = 0
    TabOrder = 4
    Text = 'Office 2013'
    Items.Strings = (
      'Office 2013'
      'Office 2016')
  end
  object InstrBtn: TButton
    Left = 296
    Top = 158
    Width = 65
    Height = 16
    Caption = 'Instructions'
    TabOrder = 5
    OnClick = InstrBtnClick
  end
  object TaskBox: TComboBox
    Left = 103
    Top = 126
    Width = 114
    Height = 21
    Style = csDropDownList
    ItemIndex = 0
    TabOrder = 6
    Text = 'Download'
    Items.Strings = (
      'Download'
      'Install')
  end
end
