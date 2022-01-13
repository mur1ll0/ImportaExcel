object frmDividir: TfrmDividir
  Left = 0
  Top = 0
  Caption = 'Dividir'
  ClientHeight = 108
  ClientWidth = 236
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 136
    Height = 13
    Caption = 'Caracter para dividir coluna:'
  end
  object edtChar: TEdit
    Left = 8
    Top = 27
    Width = 217
    Height = 21
    TabOrder = 0
  end
  object btnDividir: TButton
    Left = 150
    Top = 70
    Width = 75
    Height = 25
    Caption = 'Dividir'
    TabOrder = 1
    OnClick = btnDividirClick
  end
end
