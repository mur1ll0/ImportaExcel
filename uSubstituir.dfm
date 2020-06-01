object frmSubstituir: TfrmSubstituir
  Left = 0
  Top = 0
  Caption = 'Substituir'
  ClientHeight = 152
  ClientWidth = 409
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
    Left = 302
    Top = 119
    Width = 75
    Height = 25
    Caption = 'Substituir'
    TabOrder = 3
    OnClick = btnReplaceClick
  end
end
