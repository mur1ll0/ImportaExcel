object frmColunas: TfrmColunas
  Left = 0
  Top = 0
  Caption = 'Colunas'
  ClientHeight = 768
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
    Left = 0
    Top = 0
    Width = 296
    Height = 23
    Align = alTop
    Alignment = taCenter
    Caption = 'Tipo de Importa'#231#227'o'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clHotLight
    Font.Height = -19
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    ExplicitWidth = 165
  end
  object ListColunas: TListBox
    Left = 0
    Top = 23
    Width = 296
    Height = 745
    Align = alClient
    ItemHeight = 13
    ScrollWidth = 5
    TabOrder = 0
    OnKeyDown = ListColunasKeyDown
  end
end
