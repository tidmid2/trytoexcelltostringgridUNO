object Form3: TForm3
  Left = 0
  Top = 0
  Caption = 'Form3'
  ClientHeight = 482
  ClientWidth = 857
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 857
    Height = 481
    TabOrder = 0
    object StringGrid1: TStringGrid
      Left = 240
      Top = 80
      Width = 320
      Height = 120
      TabOrder = 0
    end
    object Button1: TButton
      Left = 392
      Top = 232
      Width = 75
      Height = 25
      Caption = 'Button1'
      TabOrder = 1
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 640
    Top = 96
  end
end
