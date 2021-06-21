unit Colunas;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Clipbrd;

type
  TfrmColunas = class(TForm)
    LabelTipoImp: TLabel;
    ListColunas: TListBox;
    procedure MostrarColunas();
    procedure FormShow(Sender: TObject);
    procedure ListColunasKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmColunas: TfrmColunas;

implementation

{$R *.dfm}

procedure TfrmColunas.MostrarColunas();
begin
  ListColunas.Items.Clear;
  //Se for Clientes/Fornecedores
  if LabelTipoImp.Caption = 'Clie/Forn' then
  begin
    ListColunas.Items.add('CODI');
    ListColunas.Items.add('UF');
    ListColunas.Items.add('CIDA');
    ListColunas.Items.add('EMPRESA');
    ListColunas.Items.add('GRUPO');
    ListColunas.Items.add('NOME');
    ListColunas.Items.add('NOME_FANT');
    ListColunas.Items.add('DATA_NASC');
    ListColunas.Items.add('CPF');
    ListColunas.Items.add('CNPJ');
    ListColunas.Items.add('CPF_CNPJ');
    ListColunas.Items.add('RG');
    ListColunas.Items.add('INSC');
    ListColunas.Items.add('INSCR_PRODUTOR');
    ListColunas.Items.add('ENDE');
    ListColunas.Items.add('BAIR');
    ListColunas.Items.add('COMP');
    ListColunas.Items.add('CEP');
    ListColunas.Items.add('PROX');
    ListColunas.Items.add('FONE');
    ListColunas.Items.add('FONE2');
    ListColunas.Items.add('FONE_FIRM');
    ListColunas.Items.add('FAX');
    ListColunas.Items.add('FIRM');
    ListColunas.Items.add('TRABALHA_DESDE');
    ListColunas.Items.add('ENDE_FIRM');
    ListColunas.Items.add('CARG');
    ListColunas.Items.add('SALA');
    ListColunas.Items.add('BAIR_FIRM');
    ListColunas.Items.add('CIDA_FIRM');
    ListColunas.Items.add('UF_FIRM');
    ListColunas.Items.add('CEP_FIRM');
    ListColunas.Items.add('ESTA_CIVI');
    ListColunas.Items.add('NOME_PAI');
    ListColunas.Items.add('NOME_MAE');
    ListColunas.Items.add('CONJ');
    ListColunas.Items.add('CONJ_FIRM');
    ListColunas.Items.add('CONJ_SALA');
    ListColunas.Items.add('CONJ_CARG');
    ListColunas.Items.add('DATA_CONJ');
    ListColunas.Items.add('OBS');
    ListColunas.Items.add('REFE_COME');
    ListColunas.Items.add('MAIL');
    ListColunas.Items.add('SEXO');
    ListColunas.Items.add('TIPOCAD');
    ListColunas.Items.add('ATIVO');
  end
  else if LabelTipoImp.Caption = 'Produtos' then
  begin
    ListColunas.Items.add('EMPR');
    ListColunas.Items.add('CODI');
    ListColunas.Items.add('GRUP');
    ListColunas.Items.add('SUB_GRUP');
    ListColunas.Items.add('DEPARTAMENTO');
    ListColunas.Items.add('MARCA');
    ListColunas.Items.add('TIPO');
    ListColunas.Items.add('QTD');
    ListColunas.Items.add('EST');
    ListColunas.Items.add('MAX');
    ListColunas.Items.add('FORN');
    ListColunas.Items.add('COLECAO');
    ListColunas.Items.add('DESCR');
    ListColunas.Items.add('DESCR2');
    ListColunas.Items.add('REFE');
    ListColunas.Items.add('REFE_ORIGINAL');
    ListColunas.Items.add('LOCALIZACAO');
    ListColunas.Items.add('CODI_BARRA');
    ListColunas.Items.add('CODI_BARRA_COM');
    ListColunas.Items.add('OBS');
    ListColunas.Items.add('NCM');
    ListColunas.Items.add('CEST');
    ListColunas.Items.add('UNID');
    ListColunas.Items.add('PESL');
    ListColunas.Items.add('PESB');
    ListColunas.Items.add('FATOR_CONV');
    ListColunas.Items.add('QTD_MINIMA');
    ListColunas.Items.add('CUSTO');
    ListColunas.Items.add('CUSTO_MEDIO');
    ListColunas.Items.add('IPI');
    ListColunas.Items.add('PIS');
    ListColunas.Items.add('COFINS');
    ListColunas.Items.add('ICMS');
    ListColunas.Items.add('FRETE');
    ListColunas.Items.add('CUSTO_REAL');
    ListColunas.Items.add('PRECO_PRAZO');
    ListColunas.Items.add('PRECO_VISTA');
    ListColunas.Items.add('MARGEM');
    ListColunas.Items.add('CSOSN');
    ListColunas.Items.add('CSOSN_ESTA');
    ListColunas.Items.add('CSOSN_INTER');
    ListColunas.Items.add('CST');
    ListColunas.Items.add('CST_ESTA');
    ListColunas.Items.add('CST_INTER');
    ListColunas.Items.add('ALIQ_ICMS');
    ListColunas.Items.add('REDU_ESTA');
    ListColunas.Items.add('REDU_INTER');
    ListColunas.Items.add('CST_IPI');
    ListColunas.Items.add('ALIQ_IPI');
    ListColunas.Items.add('CST_PIS');
    ListColunas.Items.add('ALIQ_PIS');
    ListColunas.Items.add('CST_COFINS');
    ListColunas.Items.add('ALIQ_COFINS');
    ListColunas.Items.add('ATIVO');
  end
  else if LabelTipoImp.Caption = 'Grupos' then
  begin
    ListColunas.Items.add('CODI');
    ListColunas.Items.add('EMPR');
    ListColunas.Items.add('DESCR');
  end
  else if LabelTipoImp.Caption = 'SubGrupos' then
  begin
    ListColunas.Items.add('CODI');
    ListColunas.Items.add('EMPR');
    ListColunas.Items.add('DESCR');
  end
  else if LabelTipoImp.Caption = 'Marcas' then
  begin
    ListColunas.Items.add('CODI');
    ListColunas.Items.add('DESCR');
  end
  else if LabelTipoImp.Caption = 'T�tulos a Pagar' then
  begin
    ListColunas.Items.add('CODI');
    ListColunas.Items.add('EMPR');
    ListColunas.Items.add('FORN');
    ListColunas.Items.add('LOCA_COBR');
    ListColunas.Items.add('CART');
    ListColunas.Items.add('OPER');
    ListColunas.Items.add('C_FUNC');
    ListColunas.Items.add('DATA');
    ListColunas.Items.add('VENC');
    ListColunas.Items.add('VALO');
    ListColunas.Items.add('SALD');
    ListColunas.Items.add('HIST');
    ListColunas.Items.add('DATA_BAIXA');
  end
  else if LabelTipoImp.Caption = 'T�tulos a Receber' then
  begin
    ListColunas.Items.add('CODI');
    ListColunas.Items.add('EMPR');
    ListColunas.Items.add('CLIE');
    ListColunas.Items.add('LOCA_COBR');
    ListColunas.Items.add('CART');
    ListColunas.Items.add('OPER');
    ListColunas.Items.add('C_FUNC');
    ListColunas.Items.add('DATA');
    ListColunas.Items.add('VENC');
    ListColunas.Items.add('VALO');
    ListColunas.Items.add('SALD');
    ListColunas.Items.add('HIST');
    ListColunas.Items.add('DATA_BAIXA');
  end
  else if LabelTipoImp.Caption = 'Marcas' then
  begin
    ListColunas.Items.add('CODI');
    ListColunas.Items.add('EMPR');
    ListColunas.Items.add('CLIE');
    ListColunas.Items.add('LOCA_COBR');
    ListColunas.Items.add('CART');
    ListColunas.Items.add('OPER');
    ListColunas.Items.add('C_FUNC');
    ListColunas.Items.add('DATA');
    ListColunas.Items.add('VENC');
    ListColunas.Items.add('VALOR');
    ListColunas.Items.add('SALD');
    ListColunas.Items.add('HIST');
  end
  ;
end;

procedure TfrmColunas.FormShow(Sender: TObject);
begin
  MostrarColunas;
end;


procedure TfrmColunas.ListColunasKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   //RECONHECER CTRL+C
   if ((Shift = [ssCtrl]) and (Key = 67)) then
   begin
     ClipBoard.AsText := ListColunas.Items[ListColunas.ItemIndex];
   end;
end;

end.
