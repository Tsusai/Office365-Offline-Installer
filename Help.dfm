object HelpForm: THelpForm
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Instructions'
  ClientHeight = 213
  ClientWidth = 499
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnShow = FormShow
  DesignSize = (
    499
    213)
  PixelsPerInch = 96
  TextHeight = 13
  object Memo1: TMemo
    Left = 8
    Top = 8
    Width = 482
    Height = 197
    Anchors = [akLeft, akTop, akRight, akBottom]
    BorderStyle = bsNone
    Color = clBtnFace
    Lines.Strings = (
      '*Obtain the 2013 Deployment tool from Microsoft here:'
      '  http://www.microsoft.com/en-us/download/details.aspx?id=36778'
      ''
      '*Obtain the 2016 Deployment tool from Microsoft here:'
      '  http://www.microsoft.com/en-us/download/details.aspx?id=49117'
      ''
      
        '*Run the 2013 tool to extract. Rename setup.exe to setup2013 and' +
        ' place into the setup folder'
      ''
      
        '*Run the 2016 tool to extract. Rename setup.exe to setup2016 and' +
        ' place into the setup folder'
      ''
      
        '*Using the GUI, select 2013 & Download, and click Run.  Repeat f' +
        'or 2016.'
      '  Note: "Custom Install" has no effect on Download.'
      ''
      '*Congradulations, you have an Offline Installer for Office!')
    ReadOnly = True
    TabOrder = 0
  end
end
