object Form4: TForm4
  Left = 0
  Top = 0
  Caption = 'Colunas'
  ClientHeight = 580
  ClientWidth = 296
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
  object LabelTipoImp: TLabel
    Left = 64
    Top = 8
    Width = 165
    Height = 23
    Caption = 'Tipo de Importa'#231#227'o'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clHotLight
    Font.Height = -19
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object ListColunas: TListBox
    Left = 16
    Top = 37
    Width = 257
    Height = 535
    ItemHeight = 13
    TabOrder = 0
    OnKeyDown = ListColunasKeyDown
  end
end
