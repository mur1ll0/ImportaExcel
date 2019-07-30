object Form1: TForm1
  Left = 0
  Top = 0
  Anchors = []
  BiDiMode = bdLeftToRight
  Caption = 'Importa Excel'
  ClientHeight = 600
  ClientWidth = 1004
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = Menu
  OldCreateOrder = False
  ParentBiDiMode = False
  Position = poDesigned
  DesignSize = (
    1004
    600)
  PixelsPerInch = 96
  TextHeight = 13
  object BtnLoad: TBitBtn
    AlignWithMargins = True
    Left = 8
    Top = 33
    Width = 75
    Height = 25
    Caption = 'Carregar'
    TabOrder = 0
    OnClick = BtnLoadClick
  end
  object StringGrid1: TStringGrid
    Left = 8
    Top = 64
    Width = 988
    Height = 528
    Anchors = [akLeft, akTop, akRight, akBottom]
    BiDiMode = bdLeftToRight
    ColCount = 3
    FixedColor = clBlack
    RowCount = 3
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goRowSizing, goColSizing, goRowMoving, goColMoving, goEditing]
    ParentBiDiMode = False
    TabOrder = 1
    OnClick = StringGrid1Click
    OnDblClick = StringGrid1DblClick
    OnKeyDown = StringGrid1KeyDown
    OnMouseDown = StringGrid1MouseDown
  end
  object FilePath: TEdit
    AlignWithMargins = True
    Left = 85
    Top = 2
    Width = 288
    Height = 21
    BiDiMode = bdLeftToRight
    ParentBiDiMode = False
    TabOrder = 2
    Text = 'Caminho da planilha - CSV, XLS, XLSX, etc'
  end
  object BtnAbrir: TBitBtn
    AlignWithMargins = True
    Left = 8
    Top = 2
    Width = 75
    Height = 25
    BiDiMode = bdLeftToRight
    Caption = 'Abrir...'
    ParentBiDiMode = False
    TabOrder = 3
    OnClick = BtnAbrirClick
  end
  object SelectImport: TComboBox
    Left = 851
    Top = 2
    Width = 145
    Height = 21
    Anchors = [akTop, akRight]
    TabOrder = 4
    Text = 'Tipo de Importa'#231#227'o'
    Items.Strings = (
      'Clie/Forn'
      'Grupos'
      'SubGrupos'
      'Marcas'
      'Produtos'
      'T'#237'tulos a Pagar'
      'T'#237'tulos a Receber')
  end
  object ButImport: TBitBtn
    AlignWithMargins = True
    Left = 921
    Top = 29
    Width = 75
    Height = 25
    Anchors = [akTop, akRight]
    Caption = 'Importar'
    TabOrder = 5
    OnClick = ButImportClick
  end
  object DBPath: TEdit
    Left = 460
    Top = 2
    Width = 284
    Height = 21
    BiDiMode = bdLeftToRight
    ParentBiDiMode = False
    TabOrder = 6
    Text = 'Caminho da base de dados - FDB'
  end
  object ButOpenDB: TBitBtn
    Left = 379
    Top = 2
    Width = 75
    Height = 25
    BiDiMode = bdLeftToRight
    Caption = 'Abrir...'
    ParentBiDiMode = False
    TabOrder = 7
    OnClick = BtnOpenDB
  end
  object ButSave: TBitBtn
    AlignWithMargins = True
    Left = 112
    Top = 33
    Width = 75
    Height = 25
    Caption = 'Salvar'
    TabOrder = 8
    OnClick = ButSaveClick
  end
  object OpenDialog1: TOpenDialog
    Options = [ofEnableSizing]
    Left = 80
    Top = 32
  end
  object OpenDialog2: TOpenDialog
    Options = [ofEnableSizing]
    Left = 688
  end
  object Connect: TSQLConnection
    DriverName = 'Firebird'
    Params.Strings = (
      'DriverUnit=Data.DBXFirebird'
      
        'DriverPackageLoader=TDBXDynalinkDriverLoader,DbxCommonDriver210.' +
        'bpl'
      
        'DriverAssemblyLoader=Borland.Data.TDBXDynalinkDriverLoader,Borla' +
        'nd.Data.DbxCommonDriver,Version=21.0.0.0,Culture=neutral,PublicK' +
        'eyToken=91d62ebb5b0d1b1b'
      
        'MetaDataPackageLoader=TDBXFirebirdMetaDataCommandFactory,DbxFire' +
        'birdDriver210.bpl'
      
        'MetaDataAssemblyLoader=Borland.Data.TDBXFirebirdMetaDataCommandF' +
        'actory,Borland.Data.DbxFirebirdDriver,Version=21.0.0.0,Culture=n' +
        'eutral,PublicKeyToken=91d62ebb5b0d1b1b'
      'GetDriverFunc=getSQLDriverINTERBASE'
      'LibraryName=dbxfb.dll'
      'LibraryNameOsx=libsqlfb.dylib'
      'VendorLib=fbclient.dll'
      'VendorLibWin64=fbclient.dll'
      'VendorLibOsx=/Library/Frameworks/Firebird.framework/Firebird'
      'Database=DBPath'
      'User_Name=sysdba'
      'Password=masterkey'
      'Role=RoleName'
      'MaxBlobSize=-1'
      'LocaleCode=0000'
      'IsolationLevel=ReadCommitted'
      'SQLDialect=3'
      'CommitRetain=False'
      'WaitOnLocks=True'
      'TrimChar=False'
      'BlobSize=-1'
      'ErrorResourceFile='
      'RoleName=RoleName'
      'ServerCharSet='
      'Trim Char=False')
    Left = 728
  end
  object SaveDialog1: TSaveDialog
    Left = 192
    Top = 32
  end
  object Menu: TMainMenu
    BiDiMode = bdLeftToRight
    ParentBiDiMode = False
    Left = 232
    Top = 32
    object t1: TMenuItem
      Caption = 'Carregar'
      object Cabealho1: TMenuItem
        Caption = 'Cabe'#231'alho'
        OnClick = Cabealho1Click
      end
      object DadosEmpr: TMenuItem
        Caption = 'Dados da Empresa'
        OnClick = DadosEmprClick
      end
    end
    object Editar1: TMenuItem
      Caption = 'Editar'
      object AdicionarColuna: TMenuItem
        Caption = 'Adicionar Coluna (F1)'
        OnClick = AdicionarColunaClick
      end
      object AdicionarLinha: TMenuItem
        Caption = 'Adicionar Linha (F3)'
        OnClick = AdicionarLinhaClick
      end
      object DeletarColuna: TMenuItem
        Caption = 'Deletar Coluna (Del)'
        OnClick = DeletarColunaClick
      end
      object DeletarLinha: TMenuItem
        Caption = 'Deletar Linha (Del)'
        OnClick = DeletarLinhaClick
      end
    end
    object Limpar: TMenuItem
      Caption = 'Limpar'
      object LimpaClieForn: TMenuItem
        Caption = 'Clie/Forn'#13#10
        OnClick = LimpaClieFornClick
      end
      object LimpaGrupos: TMenuItem
        Caption = 'Grupos'#13#10
        OnClick = LimpaGruposClick
      end
      object LimpaSubGrupos: TMenuItem
        Caption = 'SubGrupos'#13#10
        OnClick = LimpaSubGruposClick
      end
      object LimpaMarcas: TMenuItem
        Caption = 'Marcas'
        OnClick = LimpaMarcasClick
      end
      object LimpaProdutos: TMenuItem
        Caption = 'Produtos'
        OnClick = LimpaProdutosClick
      end
      object LimpaTituP: TMenuItem
        Caption = 'T'#237'tulos a Pagar'
        OnClick = LimpaTituPClick
      end
      object LimpaTituR: TMenuItem
        Caption = 'T'#237'tulos a Receber'
        OnClick = LimpaTituRClick
      end
    end
  end
end
