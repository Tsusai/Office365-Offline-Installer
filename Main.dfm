object Mainform: TMainform
  Left = 212
  Top = 185
  BorderIcons = [biSystemMenu, biMinimize, biHelp]
  Caption = 'Office Offline Installer'
  ClientHeight = 177
  ClientWidth = 507
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  Scaled = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 303
    Top = 156
    Width = 213
    Height = 13
    Caption = #169' 2016 Tsusai (http://github.com/tsusai)     '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = [fsItalic]
    ParentFont = False
  end
  object VersionList: TRadioGroup
    Left = 8
    Top = 32
    Width = 209
    Height = 115
    Caption = 'Version'
    ItemIndex = 0
    Items.Strings = (
      'Home && Student'
      'Home && Business'
      '365 Home Premium'
      '365 Pro Plus'
      '365 Business (&& Premium)'
      '365 Business + Skype (Full Premium)')
    TabOrder = 0
    OnClick = VersionListClick
  end
  object SoftwareList: TCheckListBox
    Left = 223
    Top = 8
    Width = 138
    Height = 139
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
    TabOrder = 1
  end
  object CustomCheck: TCheckBox
    Left = 369
    Top = 39
    Width = 98
    Height = 17
    Caption = 'Customize Install'
    TabOrder = 2
    OnClick = CustomCheckClick
  end
  object VerBox: TComboBox
    Left = 128
    Top = 8
    Width = 89
    Height = 21
    Style = csDropDownList
    ItemIndex = 1
    TabOrder = 3
    Text = 'Office 2016'
    Items.Strings = (
      'Office 2013'
      'Office 2016')
  end
  object InstrBtn: TButton
    Left = 369
    Top = 124
    Width = 130
    Height = 25
    Caption = 'Instructions / About'
    TabOrder = 4
    OnClick = InstrBtnClick
  end
  object TaskBox: TComboBox
    Left = 8
    Top = 8
    Width = 114
    Height = 21
    Style = csDropDownList
    ItemIndex = 0
    TabOrder = 5
    Text = '(Select an Option)'
    Items.Strings = (
      '(Select an Option)'
      'Download'
      'Install')
  end
  object VerifyBtn: TButton
    Left = 369
    Top = 93
    Width = 130
    Height = 25
    Caption = 'Verify Office Files'
    TabOrder = 6
    OnClick = VerifyBtnClick
  end
  object UpdateBtn: TButton
    Left = 369
    Top = 62
    Width = 130
    Height = 25
    Caption = 'Check for Office Update'
    Enabled = False
    TabOrder = 7
  end
  object RunBtn: TButton
    Left = 369
    Top = 8
    Width = 130
    Height = 25
    Caption = 'Launch Configured Setup'
    TabOrder = 8
    OnClick = RunBtnClick
  end
end
