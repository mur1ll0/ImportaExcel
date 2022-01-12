object frmSubstituir: TfrmSubstituir
  Left = 0
  Top = 0
  Caption = 'Substituir'
  ClientHeight = 163
  ClientWidth = 469
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 160
    Top = 8
    Width = 62
    Height = 13
    Caption = 'Valor Antigo:'
  end
  object Label2: TLabel
    Left = 160
    Top = 57
    Width = 56
    Height = 13
    Caption = 'Valor Novo:'
  end
  object rgSelecao: TRadioGroup
    Left = 8
    Top = 8
    Width = 121
    Height = 89
    Caption = 'Sele'#231#227'o'
    ItemIndex = 1
    Items.Strings = (
      'C'#233'lula'
      'Coluna'
      'Tudo')
    TabOrder = 0
  end
  object edtOldValue: TEdit
    Left = 160
    Top = 27
    Width = 217
    Height = 21
    TabOrder = 1
  end
  object edtNewValue: TEdit
    Left = 160
    Top = 76
    Width = 217
    Height = 21
    TabOrder = 2
  end
  object btnReplace: TButton
    Left = 386
    Top = 74
    Width = 75
    Height = 25
    Caption = 'Substituir'
    TabOrder = 3
    OnClick = btnReplaceClick
  end
  object gbExtras: TGroupBox
    Left = 8
    Top = 105
    Width = 453
    Height = 50
    Caption = 'Extras'
    TabOrder = 4
    object btnRemoveAcentos: TButton
      Left = 11
      Top = 16
      Width = 89
      Height = 25
      Caption = 'Remove Acentos'
      TabOrder = 0
      OnClick = btnRemoveAcentosClick
    end
    object btnMinuscula: TButton
      Left = 119
      Top = 16
      Width = 89
      Height = 25
      Caption = 'Min'#250'sculas'
      TabOrder = 1
      OnClick = btnMinusculaClick
    end
    object btnMaiusculas: TButton
      Left = 229
      Top = 16
      Width = 89
      Height = 25
      Caption = 'Mai'#250'sculas'
      TabOrder = 2
      OnClick = btnMaiusculasClick
    end
    object btnPrimeiraMaiuscula: TButton
      Left = 336
      Top = 16
      Width = 98
      Height = 25
      Caption = 'Primeira Mai'#250'scula'
      TabOrder = 3
      OnClick = btnPrimeiraMaiusculaClick
    end
  end
end
