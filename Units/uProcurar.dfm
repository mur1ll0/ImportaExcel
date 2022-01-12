object frmProcurar: TfrmProcurar
  Left = 0
  Top = 0
  Caption = 'Procurar'
  ClientHeight = 111
  ClientWidth = 397
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
    Left = 160
    Top = 8
    Width = 88
    Height = 13
    Caption = 'Valor para Buscar:'
  end
  object rgSelecao: TRadioGroup
    Left = 8
    Top = 8
    Width = 121
    Height = 95
    Caption = 'Sele'#231#227'o'
    ItemIndex = 0
    Items.Strings = (
      'Coluna'
      'Tudo')
    TabOrder = 0
  end
  object edtSearch: TEdit
    Left = 160
    Top = 27
    Width = 217
    Height = 21
    TabOrder = 1
  end
  object btnProximo: TButton
    Left = 302
    Top = 78
    Width = 75
    Height = 25
    Caption = 'Pr'#243'ximo'
    TabOrder = 2
    OnClick = btnProximoClick
  end
  object btnAnterior: TButton
    Left = 208
    Top = 78
    Width = 75
    Height = 25
    Caption = 'Anterior'
    TabOrder = 3
    OnClick = btnAnteriorClick
  end
end
