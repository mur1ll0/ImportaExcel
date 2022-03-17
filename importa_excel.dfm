object frmPrinc: TfrmPrinc
  Left = 0
  Top = 0
  VertScrollBar.ParentColor = False
  Anchors = []
  BiDiMode = bdLeftToRight
  Caption = 'Importa Excel'
  ClientHeight = 729
  ClientWidth = 1100
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
  PixelsPerInch = 96
  TextHeight = 13
  object StringGrid1: TStringGrid
    Left = 0
    Top = 89
    Width = 1100
    Height = 640
    Align = alClient
    BiDiMode = bdLeftToRight
    ColCount = 3
    FixedColor = clBlack
    RowCount = 3
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goRowSizing, goColSizing, goRowMoving, goColMoving, goEditing]
    ParentBiDiMode = False
    TabOrder = 0
    OnDblClick = StringGrid1DblClick
    OnDrawCell = StringGrid1DrawCell
    OnKeyDown = StringGrid1KeyDown
    OnMouseDown = StringGrid1MouseDown
  end
  object pnlTop: TPanel
    Left = 0
    Top = 0
    Width = 1100
    Height = 89
    Align = alTop
    Color = clWindow
    ParentBackground = False
    TabOrder = 1
    DesignSize = (
      1100
      89)
    object Label1: TLabel
      Left = 811
      Top = 41
      Width = 100
      Height = 16
      Anchors = [akTop, akRight]
      Caption = 'Iniciar da linha:'
      DragMode = dmAutomatic
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clCaptionText
      Font.Height = -13
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label2: TLabel
      Left = 8
      Top = 8
      Width = 94
      Height = 16
      Caption = 'Dados Origem:'
      DragMode = dmAutomatic
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clCaptionText
      Font.Height = -13
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label3: TLabel
      Left = 8
      Top = 41
      Width = 94
      Height = 16
      Caption = 'Dados Destino:'
      DragMode = dmAutomatic
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clCaptionText
      Font.Height = -13
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblColUpdate: TLabel
      Left = 8
      Top = 66
      Width = 93
      Height = 15
      Caption = 'Colunas Update:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      Visible = False
    end
    object lblTableName: TLabel
      Left = 635
      Top = 8
      Width = 145
      Height = 16
      Anchors = [akTop, akRight]
      Caption = 'Nome tabela CUSTOM:'
      DragMode = dmAutomatic
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clCaptionText
      Font.Height = -13
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ParentFont = False
      Visible = False
    end
    object btnLoadOrigem: TBitBtn
      AlignWithMargins = True
      Left = 479
      Top = 6
      Width = 24
      Height = 23
      Hint = 'Carregar Dados Origem'
      Glyph.Data = {
        76050000424D7605000000000000360000002800000015000000150000000100
        18000000000040050000C40E0000C40E00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FF9494945A5A5A20241F161E1420231E5C5C
        5C989898FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00FF00FFFF00FF
        FF00FFFF00FFFF00FF06210E17702D2AB3402BB03C2CA4312F9D25309B203599
        1537960F2256040A1800FF00FFFF00FFFF00FFFF00FFFF00FF00FF00FFFF00FF
        FF00FFA9A9A914191622B44F25BD5026AD412CA83135A41731A00933A00D3398
        1533921439940A378705181A16ADADADFF00FFFF00FFFF00FF00FF00FFFF00FF
        A0A0A018AB5C1FCD6A24B84C32AD2237AA1437AD163A931AF7F6F6DCE3D532AA
        1137AB1535A6163495103A9105317401A3A3A3FF00FFFF00FF00FF00FFFF00FF
        0B16111BD3741AC56D33AE2034AD1C31AF1E249F11DAE5D6FFFFFFFFFFFF2AAB
        1633AE2033AF2033AE20358A093B91050E120AFF00FFFF00FF00FF00FF063A25
        15DB871ACA702CB6362DB5322DB53283CF87FFFFFFFFFFFFFFFFFFDEEDDF69B5
        6F26A52F2EB5332DB4322EB230349313398E050F2500FF00FF00FF00FF0EA469
        14E18E23C1532CB6372AB83B2AB83B7FD58AFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFC0DEC721B4332BB83C2ABA3E2EAA2C38930A2A6804FF00FF0075757510EEA0
        0FDD9426BB4825BD4C25BD4C25BD4C1FBC472CC052FFFFFFE4F4EA5AC97AFAFD
        FBFFFFFFC9EAD522B94B25BD4C24C04F338F1137980F737373004040400EEDA2
        0EE09723BE5122C05421C0531DBF5023C05518BD4CEEFAF1FFFFFF70D68F63D3
        86F8FDF9FFFFFF6DD19122C05421C2573294173598143E3E3E001624200BEDA9
        0DE49D1EC4611DC5629DE2BFF2FAF81BC3611DC56216C35D21C66519C45F15C3
        5C43CF7CFFFFFFE2F7EB1BC4601EC6642E9D26319A1D171D140014251F09F0AE
        0BE7A31CC76718C666F0FCF6FFFFFF30CC7817C7661CC86919C7671CC86917C7
        6630CC77FFFFFFEBFAF218C6661BC96B2D9F2A309D221720170018191306F9BC
        07F0B219C97218CC76AFEDCFFFFFFFA6EBCF07C86F18CC7618CB7618CC7607C8
        6FA6EBCFFFFFFFAFEDCF18CB7517CE7A2BA22E2DA52E191B1500403E3E04FFC6
        05F4B911D88D15CE7E0CCE7CADEFD3FFFFFFFFFFFF58DEAD32D79958DEADFFFF
        FFFFFFFFADEFD30CCE7C13D18419C77127A93C2BAF39404040007373730DECA6
        02FCC409EBAB13D1850DD0821AD389FFFFFFFFFFFFFFFFFFF8FEFDFFFFFFFFFF
        FFFFFFFF1AD3890DD08213D18620B85726AF432CAA3675757500FF00FF255609
        0EE59F04F7BE05F5BA11D28B11D38D04D18735DA9FBEF3E0D6F8EBBEF3E035DA
        9F04D18710D48E0FD58F20BA5821B85428AD4124590BFF00FF00FF00FF101A00
        2A912F02FDC904F7BE0AE7A70FD6920FD48F08D38C08D38C12D59108D38C08D3
        8C0FD59010D48F17CA781EBE5E1FBD5C2E89200E1E02FF00FF00FF00FFCCCCCC
        0C110B21B05603F6C204F7BE05F4BA09ECAC0BE4A20DDA970ED8940FD89311D6
        8C14D38417CD7719CA731CC36728A33F0C110BCCCCCCFF00FF00FF00FFFF00FF
        A1A1A1266F1923AA4D01FFCB03FAC207F1B508EFB00BE8A50EE29B0FDF9512D9
        8B13D68616D27E17D07B26A544247624A0A0A0FF00FFFF00FF00FF00FFFF00FF
        FF00FFACACAC161A17219D471DBC660AE5A707EFB308EEAF0BE7A40CE5A00FDE
        9511D58B1CBE691FA450151916AAAAAAFF00FFFF00FFFF00FF00FF00FFFF00FF
        FF00FFFF00FFFF00FF061D0E126D3916CE8111DE9707F1B504FAC008F1B410E1
        9A14D38711703E061F10FF00FFFF00FFFF00FFFF00FFFF00FF00FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FF9797975C5C5C1D28251226221D28255B5B
        5B969696FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00}
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = btnLoadOrigemClick
    end
    object FilePath: TEdit
      AlignWithMargins = True
      Left = 147
      Top = 6
      Width = 324
      Height = 21
      BiDiMode = bdLeftToRight
      ParentBiDiMode = False
      TabOrder = 1
      Text = 'Caminho Planilha (CSV, XLS, XLSX, etc) ou DADOS (.FDB)'
    end
    object btnAbrirOrigem: TBitBtn
      AlignWithMargins = True
      Left = 108
      Top = 6
      Width = 33
      Height = 23
      Hint = 'Selecionar Dados Origem'
      BiDiMode = bdLeftToRight
      Glyph.Data = {
        360C0000424D360C000000000000360000002800000020000000200000000100
        180000000000000C0000C40E0000C40E00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFD8D8D7D2D2D1D2D2D1D2D2D1D0D0CFC9C8C7C8C7
        C6C8C8C6C8C8C6C8C8C6C8C8C6C8C8C6C8C8C6C8C7C6C9C8C7D0D0CFD2D2D1D2
        D2D1D3D2D1DAD9D8FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        C8C3C2C3BCB8AAB1B4A6ADB1A6ACAFA2A9ACA0A7ABA0A7ABA0A7ABA0A7AB9FA6
        A99DA4A89DA4A89DA4A89DA4A89DA4A89DA4A89FA6AAA0A7ABA0A7ABA0A7ABA0
        A7ABA2A9ADA6ACB0A6ADB0ABB1B5C3BDB9CAC5C3D7D6D5FF00FFFF00FFFF00FF
        FF00FF11AEE800ABEB00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00AB
        EA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00
        ABEA00ABEA00ABEA00ABEA00ABEB0BACE8FF00FFFF00FFFF00FFFF00FFFF00FF
        22B5EB0DABE50EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAA
        E40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40E
        AAE40EAAE40EAAE40EAAE40EAAE40DABE515B0E8FF00FFFF00FFFF00FFFF00FF
        13ADE913ACE613ACE613ACE613ACE613ACE613ACE613ACE613ACE613ACE613AC
        E613ACE613ACE613ACE613ACE613ACE613ACE613ACE613ACE613ACE613ACE613
        ACE613ACE613ACE613ACE613ACE613ACE613ACE6FF00FFFF00FFFF00FFFF00FF
        16AEE815ACE515ACE515ACE515ACE515ACE515ACE515ACE515ACE515ACE515AC
        E515ACE515ACE515ACE515ACE515ACE515ACE515ACE515ACE515ACE515ACE515
        ACE515ACE515ACE515ACE515ACE515ACE515ACE5FF00FFFF00FFFF00FFFF00FF
        1BB1E91BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAF
        E71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71B
        AFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE7FF00FFFF00FFFF00FFFF00FF
        1FB1EB1EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0
        E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81E
        B0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E8FF00FFFF00FFFF00FFFF00FF
        24B2EC24B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0
        E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924
        B0E924B0E924B0E924B0E924B0E924B0E924B0E9FF00FFFF00FFFF00FFFF00FF
        28B5EB28B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3
        E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828
        B3E828B3E828B3E828B3E828B3E828B3E828B3E8FF00FFFF00FFFF00FFFF00FF
        2CB5ED2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4
        EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2B
        B4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EAFF00FFFF00FFFF00FFFF00FF
        31B6EE31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4
        EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31
        B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EBFF00FFFF00FFFF00FFFF00FF
        38B9EF33B6EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7
        EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35
        B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7ECFF00FFFF00FFFF00FFFF00FF
        4DC3F140BAEB34B6EB37B6EB37B6EB37B6EB37B6EB37B6EB37B6EB37B6EB37B7
        EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38
        B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EBFF00FFFF00FFFF00FFFF00FF
        47BDEF86DCF85CC7F14DBFEE4EC0EE4EC0EE4EC0EE4EC0EE4EC0EE4FC0EE46BD
        ED36B5EB3BB7EC3EB9ED3EB9ED3EB9ED3EB9ED3EB9ED3EB9ED3EB9ED3EB9ED3E
        B9ED3EB9ED3EB9ED3EB9ED3EB9ED3EB9ED3EB9EDFF00FFFF00FFFF00FFFF00FF
        3192D63FB9ED6CD3F777D7F776D6F776D6F776D6F776D6F776D6F775D5F77FDA
        F891E2FA61CAF23BB6EC41B9EE41B9EE41B9EE41B9EE41B9EE41B9EE41B9EE41
        B9EE41B9EE41B9EE41B9EE41B9EE41B9EE41B9EEFF00FFFF00FFFF00FFFF00FF
        298AD40772C70C7DCE0C7CCD0C7CCD0C7CCD0C7CCD0C7CCD0C7CCD0C7CCC1284
        D02AA2E271D3F66CCEF342B9EC47BCED47BCED47BCED47BCED47BCED47BCED47
        BCED47BCED47BCED47BCED47BCED47BCED47BCEDFF00FFFF00FFFF00FFFF00FF
        2E92DB107ED1107DD10F7DD10F7DD10F7DD10F7DD10F7DD10F7DD10F7DD10F7C
        D00B78CE1688D772D3F556C1F049BBEE4BBDEF4BBDEF4BBDEF4BBDEF4BBDEF4B
        BDEF4BBDEF4BBDEF4BBDEF4BBDEF4BBDEF4BBDEFFF00FFFF00FFFF00FFFF00FF
        3298E21786DA1385D81386D91386D91386D91386D91386D91386D91386D91386
        D91386D90F81D62397E087DDF948BBEE4EBDF04EBEF04EBEF04EBEF04EBEF04E
        BEF04EBEF04EBEF04EBEF04EBEF04EBEF04DBDF0FF00FFFF00FFFF00FFFF00FF
        44A4EA439FE30F88DF1389DF1389DF1389DF1389DF1389DF1389DF1389DF1389
        DF1389DF1389DF0C84DD42B2EC80D6F64CBAEF53BEF154BEF154BEF154BEF154
        BEF154BEF154BEF154BEF154BEF152BDF05AC3F2FF00FFFF00FFFF00FFFF00FF
        FF00FF2898E858AAEB54A9E954A9E954A9E954A9E954A9E954A9E954A9E954A9
        E954A9E954A9E953A9E953A8E86ACCF488DBF852BDF053BEEF53BEEF53BEEF53
        BEEF53BEEF53BEEF53BEEF52BEEF5EC4F273CFF4FF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FF65C6F292E0FA92DFFA91DFFA91DFFA91
        DFFA91DFFA91DFFA91DFFA92E0FA8BDCF95EC1F1FF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF79CBF45CC0F25CC0F25CC0F25C
        C0F25CC0F25CC0F25CC0F25CC0F25DC0F2FF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      ParentBiDiMode = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = btnAbrirOrigemClick
    end
    object SelectImport: TComboBox
      Left = 920
      Top = 5
      Width = 171
      Height = 21
      Anchors = [akTop, akRight]
      TabOrder = 3
      Text = 'Tipo de Importa'#231#227'o'
      OnChange = SelectImportChange
      Items.Strings = (
        'Clie/Forn'
        'Grupos'
        'SubGrupos'
        'Marcas'
        'Produtos'
        'T'#237'tulos a Pagar'
        'T'#237'tulos a Receber'
        'Grades'
        'CUSTOM')
    end
    object ButImport: TBitBtn
      AlignWithMargins = True
      Left = 1016
      Top = 37
      Width = 75
      Height = 25
      Anchors = [akTop, akRight]
      Caption = 'Importar'
      TabOrder = 4
      OnClick = ButImportClick
    end
    object DBPath: TEdit
      Left = 147
      Top = 39
      Width = 324
      Height = 21
      BiDiMode = bdLeftToRight
      ParentBiDiMode = False
      TabOrder = 5
      Text = 'Caminho do destino (TXT, SQL) ou DADOS (.FDB)'
    end
    object ButOpenDB: TBitBtn
      Left = 108
      Top = 39
      Width = 33
      Height = 23
      Hint = 'Selecionar Dados Destino'
      BiDiMode = bdLeftToRight
      Glyph.Data = {
        360C0000424D360C000000000000360000002800000020000000200000000100
        180000000000000C0000C40E0000C40E00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFD8D8D7D2D2D1D2D2D1D2D2D1D0D0CFC9C8C7C8C7
        C6C8C8C6C8C8C6C8C8C6C8C8C6C8C8C6C8C8C6C8C7C6C9C8C7D0D0CFD2D2D1D2
        D2D1D3D2D1DAD9D8FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        C8C3C2C3BCB8AAB1B4A6ADB1A6ACAFA2A9ACA0A7ABA0A7ABA0A7ABA0A7AB9FA6
        A99DA4A89DA4A89DA4A89DA4A89DA4A89DA4A89FA6AAA0A7ABA0A7ABA0A7ABA0
        A7ABA2A9ADA6ACB0A6ADB0ABB1B5C3BDB9CAC5C3D7D6D5FF00FFFF00FFFF00FF
        FF00FF11AEE800ABEB00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00AB
        EA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00ABEA00
        ABEA00ABEA00ABEA00ABEA00ABEB0BACE8FF00FFFF00FFFF00FFFF00FFFF00FF
        22B5EB0DABE50EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAA
        E40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40EAAE40E
        AAE40EAAE40EAAE40EAAE40EAAE40DABE515B0E8FF00FFFF00FFFF00FFFF00FF
        13ADE913ACE613ACE613ACE613ACE613ACE613ACE613ACE613ACE613ACE613AC
        E613ACE613ACE613ACE613ACE613ACE613ACE613ACE613ACE613ACE613ACE613
        ACE613ACE613ACE613ACE613ACE613ACE613ACE6FF00FFFF00FFFF00FFFF00FF
        16AEE815ACE515ACE515ACE515ACE515ACE515ACE515ACE515ACE515ACE515AC
        E515ACE515ACE515ACE515ACE515ACE515ACE515ACE515ACE515ACE515ACE515
        ACE515ACE515ACE515ACE515ACE515ACE515ACE5FF00FFFF00FFFF00FFFF00FF
        1BB1E91BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAF
        E71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE71B
        AFE71BAFE71BAFE71BAFE71BAFE71BAFE71BAFE7FF00FFFF00FFFF00FFFF00FF
        1FB1EB1EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0
        E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E81E
        B0E81EB0E81EB0E81EB0E81EB0E81EB0E81EB0E8FF00FFFF00FFFF00FFFF00FF
        24B2EC24B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0
        E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924B0E924
        B0E924B0E924B0E924B0E924B0E924B0E924B0E9FF00FFFF00FFFF00FFFF00FF
        28B5EB28B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3
        E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828B3E828
        B3E828B3E828B3E828B3E828B3E828B3E828B3E8FF00FFFF00FFFF00FFFF00FF
        2CB5ED2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4
        EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2B
        B4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EA2BB4EAFF00FFFF00FFFF00FFFF00FF
        31B6EE31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4
        EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31
        B4EB31B4EB31B4EB31B4EB31B4EB31B4EB31B4EBFF00FFFF00FFFF00FFFF00FF
        38B9EF33B6EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7
        EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35
        B7EC35B7EC35B7EC35B7EC35B7EC35B7EC35B7ECFF00FFFF00FFFF00FFFF00FF
        4DC3F140BAEB34B6EB37B6EB37B6EB37B6EB37B6EB37B6EB37B6EB37B6EB37B7
        EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38
        B8EB38B8EB38B8EB38B8EB38B8EB38B8EB38B8EBFF00FFFF00FFFF00FFFF00FF
        47BDEF86DCF85CC7F14DBFEE4EC0EE4EC0EE4EC0EE4EC0EE4EC0EE4FC0EE46BD
        ED36B5EB3BB7EC3EB9ED3EB9ED3EB9ED3EB9ED3EB9ED3EB9ED3EB9ED3EB9ED3E
        B9ED3EB9ED3EB9ED3EB9ED3EB9ED3EB9ED3EB9EDFF00FFFF00FFFF00FFFF00FF
        3192D63FB9ED6CD3F777D7F776D6F776D6F776D6F776D6F776D6F775D5F77FDA
        F891E2FA61CAF23BB6EC41B9EE41B9EE41B9EE41B9EE41B9EE41B9EE41B9EE41
        B9EE41B9EE41B9EE41B9EE41B9EE41B9EE41B9EEFF00FFFF00FFFF00FFFF00FF
        298AD40772C70C7DCE0C7CCD0C7CCD0C7CCD0C7CCD0C7CCD0C7CCD0C7CCC1284
        D02AA2E271D3F66CCEF342B9EC47BCED47BCED47BCED47BCED47BCED47BCED47
        BCED47BCED47BCED47BCED47BCED47BCED47BCEDFF00FFFF00FFFF00FFFF00FF
        2E92DB107ED1107DD10F7DD10F7DD10F7DD10F7DD10F7DD10F7DD10F7DD10F7C
        D00B78CE1688D772D3F556C1F049BBEE4BBDEF4BBDEF4BBDEF4BBDEF4BBDEF4B
        BDEF4BBDEF4BBDEF4BBDEF4BBDEF4BBDEF4BBDEFFF00FFFF00FFFF00FFFF00FF
        3298E21786DA1385D81386D91386D91386D91386D91386D91386D91386D91386
        D91386D90F81D62397E087DDF948BBEE4EBDF04EBEF04EBEF04EBEF04EBEF04E
        BEF04EBEF04EBEF04EBEF04EBEF04EBEF04DBDF0FF00FFFF00FFFF00FFFF00FF
        44A4EA439FE30F88DF1389DF1389DF1389DF1389DF1389DF1389DF1389DF1389
        DF1389DF1389DF0C84DD42B2EC80D6F64CBAEF53BEF154BEF154BEF154BEF154
        BEF154BEF154BEF154BEF154BEF152BDF05AC3F2FF00FFFF00FFFF00FFFF00FF
        FF00FF2898E858AAEB54A9E954A9E954A9E954A9E954A9E954A9E954A9E954A9
        E954A9E954A9E953A9E953A8E86ACCF488DBF852BDF053BEEF53BEEF53BEEF53
        BEEF53BEEF53BEEF53BEEF52BEEF5EC4F273CFF4FF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FF65C6F292E0FA92DFFA91DFFA91DFFA91
        DFFA91DFFA91DFFA91DFFA92E0FA8BDCF95EC1F1FF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF79CBF45CC0F25CC0F25CC0F25C
        C0F25CC0F25CC0F25CC0F25CC0F25DC0F2FF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      ParentBiDiMode = False
      TabOrder = 6
      OnClick = BtnOpenDB
    end
    object StartLine: TEdit
      Left = 920
      Top = 39
      Width = 65
      Height = 21
      Anchors = [akTop, akRight]
      NumbersOnly = True
      TabOrder = 7
      Text = '1'
    end
    object btnTXT: TBitBtn
      AlignWithMargins = True
      Left = 479
      Top = 39
      Width = 23
      Height = 23
      Hint = 'Salvar Comandos SQL em arquivo de texto'
      Glyph.Data = {
        76050000424D7605000000000000360000002800000015000000150000000100
        18000000000040050000C40E0000C40E00000000000000000000FF00FFE7E7E7
        E7E7E7E7E7E7E7E7E7E7E7E7E6E6E6E7E7E7E7E7E7E6E6E6E6E6E6E6E6E6E6E6
        E6E6E6E6E6E6E6E6E6E6E6E6E6E6E6E6E6E6E6E8E8E8FF00FF004848484A4A4A
        239EFF2498FF2091FF1F8EFF1D8BFF1985FF1782FF167FFF1278FF1075FF0E72
        FF0B6CFF0969FF0766FF035FFF0B64FF404040474747AAAAAA004646464A4A4A
        FFFFFFF5F4F4241CED241CEDF3F2F2F3F2F2F2F2F1241CED241CED241CEDEFEF
        EF241CED241CED241CED241CEDECEBEB3E3E3E474747E1E1E100313131303030
        FFFFFF241CEDF4F4F4F4F4F4241CEDF2F2F2241CEDF1F1F1241CEDF0F0F0F3F3
        F3241CEDEFEFEFEEEEEEEDEDEDECECEC2A2A2A484848FF00FF00323232313131
        FFFFFFF6F6F6F6F6F6F5F5F5241CEDF4F4F4241CEDF2F2F2F2F2F2241CEDF0F0
        F0241CEDF0F0F0F0F0F0EFEFEFEDEDED2B2B2B474747FF00FF00323232313131
        FFFFFFF7F7F7241CED241CEDF5F5F5F4F4F4241CEDF3F3F3F2F2F2241CEDF3F3
        F3241CEDF1F1F1F0F0F0EFEFEFEEEEEE2B2B2B474747FF00FF00323232313131
        FFFFFF241CEDFAFAFAFAFAFAF9F9F9F8F8F8241CEDF7F7F7F6F6F6241CEDF3F3
        F3241CEDF3F3F3F4F4F4F3F3F3EEEEEE2B2B2B474747FF00FF00323232313131
        FFFFFF241CEDFAFAFAFAFAFA241CEDF8F8F8241CEDF7F7F7F6F6F6241CEDF3F3
        F3241CEDF4F4F4F4F4F4F3F3F3EEEEEE2B2B2B474747FF00FF00343434313131
        FFFFFFFAFAFA241CED241CEDF9F9F9F8F8F8F8F8F8241CED241CEDF3F3F3F3F3
        F3241CEDF4F4F4F4F4F4F3F3F3EEEEEE2B2B2B474747FF00FF00373737343434
        FFFFFFFBFBFBFAFAFAFAFAFAFAFAFAF9F9F9F8F8F8F8F8F8F7F7F7F6F6F6F6F6
        F6F5F5F5F4F4F4F4F4F4F3F3F3EEEEEE2B2B2B474747FF00FF00383838353535
        FFFFFFC4C4C4C4C4C4C4C4C4C4C4C4C3C3C3C3C3C3C3C3C3C3C3C3C3C3C3C3C3
        C3C2C2C2C2C2C2C2C2C2C0C0C0F0F0F02C2C2C494949FF00FF003A3A3A373737
        FFFFFFFAFAFAF9F9F9F9F9F9F8F8F8F7F7F7F7F7F7F6F6F6F5F5F5F5F5F5F4F4
        F4F3F3F3F3F3F3F3F3F3F2F2F2F0F0F02E2E2E4B4B4BFF00FF003D3D3D373737
        4444444444444444444444444444444545454646464747474949494A4A4A4B4B
        4B4D4D4D4E4E4E4F4F4F5151515151514A4A4A4D4D4DFF00FF003F3F3F3A3A3A
        3939393939393939393939393A3A3A3B3B3B3D3D3D3F3F3F4242424343434545
        454848484A4A4A4B4B4B4C4C4C4D4D4D4D4D4D4F4F4FFF00FF00424242313131
        303030303030313131F5F5F5EBEBEBCACACABBBBBBABABAB8D8D8DCACACA1E1E
        1E1E1E1EE1E1E1F4F4F44F4F4F5050505151515E5E5EFF00FF00434343333333
        323232323232333333F5F5F5EBEBEBCACACABBBBBBABABAB8D8D8DCBCBCB3232
        32323232E1E1E1F4F4F45252525353535353535F5F5FFF00FF00454545353535
        343434343434353535F5F5F5EBEBEBCACACABBBBBBABABAB8D8D8DCBCBCB3434
        34343434E1E1E1F4F4F4545454555555555555616161FF00FF004646463B3B3B
        3B3B3B3B3B3B3C3C3CF6F6F6EBEBEBCACACABBBBBBABABAB8D8D8DCBCBCB3B3B
        3B3B3B3BE1E1E1F4F4F45858586262625D5D5D636363FF00FF00464646404040
        404040404040414141F6F6F6EBEBEBCACACABBBBBBABABAB8D8D8DCBCBCB4040
        40404040E1E1E1F4F4F45959595D5D5D000000646464FF00FF00474747444444
        444444444444454545F7F7F7EBEBEBCACACABBBBBBABABAB8D8D8DCACACA4444
        44444444E1E1E1F4F4F4575757585858595959646464FF00FF004C4C4CF5F5F5
        ECECECECECECECECECF3F3F3F0F0F0EEEEEEECECECEBEBEBE9E9E9EAEAEAECEC
        ECEEEEEEF0F0F0F1F1F14A4A4A494949646464FF00FFFF00FF00}
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      OnClick = btnTXTClick
    end
    object edtTableName: TEdit
      Left = 786
      Top = 5
      Width = 121
      Height = 21
      Anchors = [akTop, akRight]
      TabOrder = 9
      Text = 'TABLE'
      Visible = False
    end
    object btnSalvar: TBitBtn
      AlignWithMargins = True
      Left = 509
      Top = 6
      Width = 23
      Height = 23
      Hint = 'Salvar Dados Origem'
      Glyph.Data = {
        76050000424D7605000000000000360000002800000015000000150000000100
        18000000000040050000C40E0000C40E00000000000000000000FF00FFE7E7E7
        E7E7E7E7E7E7E7E7E7E7E7E7E6E6E6E7E7E7E7E7E7E6E6E6E6E6E6E6E6E6E6E6
        E6E6E6E6E6E6E6E6E6E6E6E6E6E6E6E6E6E6E6E8E8E8FF00FF004848484A4A4A
        239EFF2498FF2091FF1F8EFF1D8BFF1985FF1782FF167FFF1278FF1075FF0E72
        FF0B6CFF0969FF0766FF035FFF0B64FF404040474747AAAAAA004646464A4A4A
        FFFFFFF6F5F5F5F4F4F5F4F4F4F3F3F3F2F2F3F2F2F2F2F1F1F1F0F1F0F0F0F0
        F0EFEFEFEFEEEEEFEEEEEEEDEDECEBEB3E3E3E474747E1E1E100313131303030
        FFFFFFF6F6F6F5F5F5F4F4F4F4F4F4F3F3F3F2F2F2F2F2F2F1F1F1F1F1F1F0F0
        F0EFEFEFEFEFEFEEEEEEEDEDEDECECEC2A2A2A484848FF00FF00323232313131
        FFFFFFF7F7F7F6F6F6F6F6F6F5F5F5F4F4F4F4F4F4F3F3F3F2F2F2F2F2F2F1F1
        F1F0F0F0F0F0F0F0F0F0EFEFEFEDEDED2B2B2B474747FF00FF00323232313131
        FFFFFFF8F8F8F7F7F7F6F6F6F6F6F6F5F5F5F4F4F4F4F4F4F3F3F3F2F2F2F2F2
        F2F1F1F1F1F1F1F0F0F0EFEFEFEEEEEE2B2B2B474747FF00FF00323232313131
        FFFFFFFBFBFBFAFAFAFAFAFAFAFAFAF9F9F9F8F8F8F8F8F8F7F7F7F6F6F6F6F6
        F6F5F5F5F4F4F4F4F4F4F3F3F3EEEEEE2B2B2B474747FF00FF00323232313131
        FFFFFFC4C4C4C4C4C4C4C4C4C4C4C4C3C3C3C3C3C3C3C3C3C3C3C3C3C3C3C3C3
        C3C2C2C2C2C2C2C2C2C2C0C0C0F0F0F02B2B2B474747FF00FF00343434313131
        FFFFFFFAFAFAF9F9F9F9F9F9F8F8F8F7F7F7F7F7F7F6F6F6F5F5F5F5F5F5F4F4
        F4F3F3F3F3F3F3F3F3F3F2F2F2F0F0F02B2B2B474747FF00FF00373737343434
        FFFFFFFBFBFBFAFAFAFAFAFAF9F9F9F8F8F8F8F8F8F8F8F8F7F7F7F6F6F6F6F6
        F6F5F5F5F4F4F4F4F4F4F3F3F3F1F1F12B2B2B474747FF00FF00383838353535
        FFFFFFFCFCFCFBFBFBFBFBFBFAFAFAF9F9F9F9F9F9F8F8F8F7F7F7F7F7F7F6F6
        F6F5F5F5F5F5F5F4F4F4F3F3F3F2F2F22C2C2C494949FF00FF003A3A3A373737
        FFFFFFFDFDFDFCFCFCFBFBFBFBFBFBFAFAFAF9F9F9F9F9F9F8F8F8F7F7F7F7F7
        F7F6F6F6F6F6F6F5F5F5F4F4F4F3F3F32E2E2E4B4B4BFF00FF003D3D3D373737
        4444444444444444444444444444444545454646464747474949494A4A4A4B4B
        4B4D4D4D4E4E4E4F4F4F5151515151514A4A4A4D4D4DFF00FF003F3F3F3A3A3A
        3939393939393939393939393A3A3A3B3B3B3D3D3D3F3F3F4242424343434545
        454848484A4A4A4B4B4B4C4C4C4D4D4D4D4D4D4F4F4FFF00FF00424242313131
        303030303030313131F5F5F5EBEBEBCACACABBBBBBABABAB8D8D8DCACACA1E1E
        1E1E1E1EE1E1E1F4F4F44F4F4F5050505151515E5E5EFF00FF00434343333333
        323232323232333333F5F5F5EBEBEBCACACABBBBBBABABAB8D8D8DCBCBCB3232
        32323232E1E1E1F4F4F45252525353535353535F5F5FFF00FF00454545353535
        343434343434353535F5F5F5EBEBEBCACACABBBBBBABABAB8D8D8DCBCBCB3434
        34343434E1E1E1F4F4F4545454555555555555616161FF00FF004646463B3B3B
        3B3B3B3B3B3B3C3C3CF6F6F6EBEBEBCACACABBBBBBABABAB8D8D8DCBCBCB3B3B
        3B3B3B3BE1E1E1F4F4F45858586262625D5D5D636363FF00FF00464646404040
        404040404040414141F6F6F6EBEBEBCACACABBBBBBABABAB8D8D8DCBCBCB4040
        40404040E1E1E1F4F4F45959595D5D5D000000646464FF00FF00474747444444
        444444444444454545F7F7F7EBEBEBCACACABBBBBBABABAB8D8D8DCACACA4444
        44444444E1E1E1F4F4F4575757585858595959646464FF00FF004C4C4CF5F5F5
        ECECECECECECECECECF3F3F3F0F0F0EEEEEEECECECEBEBEBE9E9E9EAEAEAECEC
        ECEEEEEEF0F0F0F1F1F14A4A4A494949646464FF00FFFF00FF00}
      ParentShowHint = False
      ShowHint = True
      TabOrder = 10
      OnClick = btnSalvarClick
    end
  end
  object opnDadosOrigem: TOpenDialog
    Options = [ofEnableSizing]
    Left = 542
  end
  object opnDadosDestino: TOpenDialog
    Options = [ofEnableSizing]
    Left = 536
    Top = 41
  end
  object conDestino: TSQLConnection
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
    Left = 574
    Top = 41
  end
  object SaveDialog1: TSaveDialog
    Left = 574
  end
  object Menu: TMainMenu
    BiDiMode = bdLeftToRight
    ParentBiDiMode = False
    Left = 608
    Top = 40
    object t1: TMenuItem
      Caption = 'Carregar'
      object mnuCabecalho: TMenuItem
        Caption = 'Cabe'#231'alho'
        OnClick = mnuCabecalhoClick
      end
      object mnuDadosEmpr: TMenuItem
        Caption = 'Dados da Empresa'
        OnClick = mnuDadosEmprClick
      end
      object mnuColunas: TMenuItem
        Caption = 'Colunas'
        OnClick = mnuColunasClick
      end
    end
    object Editar1: TMenuItem
      Caption = 'Editar'
      object mnuAdicionarColuna: TMenuItem
        Caption = 'Adicionar Coluna (F1)'
        OnClick = mnuAdicionarColunaClick
      end
      object mnuAdicionarLinha: TMenuItem
        Caption = 'Adicionar Linha (F3)'
        OnClick = mnuAdicionarLinhaClick
      end
      object mnuDeletarColuna: TMenuItem
        Caption = 'Deletar Coluna (Del)'
        OnClick = mnuDeletarColunaClick
      end
      object mnuDeletarLinha: TMenuItem
        Caption = 'Deletar Linha (Del)'
        OnClick = mnuDeletarLinhaClick
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object mnuProcurar: TMenuItem
        Caption = 'Procurar (CTRL + F)'
        OnClick = mnuProcurarClick
      end
      object mnuSubstituir: TMenuItem
        Caption = 'Substituir (CTRL + H)'
        OnClick = mnuSubstituirClick
      end
      object mnuDividir: TMenuItem
        Caption = 'Dividir (CTRL + D)'
        OnClick = mnuDividirClick
      end
    end
    object Limpar: TMenuItem
      Caption = 'Limpar'
      object mnuLimpaClieForn: TMenuItem
        Caption = 'Clie/Forn'#13#10
        OnClick = mnuLimpaClieFornClick
      end
      object mnuLimpaGrupos: TMenuItem
        Caption = 'Grupos'#13#10
        OnClick = mnuLimpaGruposClick
      end
      object mnuLimpaSubGrupos: TMenuItem
        Caption = 'SubGrupos'#13#10
        OnClick = mnuLimpaSubGruposClick
      end
      object mnuLimpaMarcas: TMenuItem
        Caption = 'Marcas'
        OnClick = mnuLimpaMarcasClick
      end
      object mnuLimpaProdutos: TMenuItem
        Caption = 'Produtos'
        OnClick = mnuLimpaProdutosClick
      end
      object mnuLimpaTituP: TMenuItem
        Caption = 'T'#237'tulos a Pagar'
        OnClick = mnuLimpaTituPClick
      end
      object mnuLimpaTituR: TMenuItem
        Caption = 'T'#237'tulos a Receber'
        OnClick = mnuLimpaTituRClick
      end
    end
  end
  object conOrigem: TSQLConnection
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
      'Database=FilePath'
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
    Left = 606
  end
end
