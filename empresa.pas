unit empresa;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Data.SqlExpr;

type
  TForm3 = class(TForm)
    LabelRazao: TLabel;
    razaoSocial: TEdit;
    LabelFantasia: TLabel;
    nomeFantasia: TEdit;
    Label1: TLabel;
    cnpj: TEdit;
    Label2: TLabel;
    empr: TEdit;
    Label3: TLabel;
    inscr: TEdit;
    simples: TRadioButton;
    normal: TRadioButton;
    Label4: TLabel;
    ende: TEdit;
    Label5: TLabel;
    bair: TEdit;
    Label6: TLabel;
    cep: TEdit;
    Label7: TLabel;
    cida: TComboBox;
    Label8: TLabel;
    fone: TEdit;
    mail: TEdit;
    Label9: TLabel;
    save: TBitBtn;
    cancel: TBitBtn;
    uf: TEdit;
    Label10: TLabel;
    procedure cancelClick(Sender: TObject);
    procedure saveClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation


{$R *.dfm}

uses importa_excel;


//Bot�o Cancelar (Sair)
procedure TForm3.cancelClick(Sender: TObject);
begin
  Form3.Close;
end;


//Fun��o para executar SQL e retornar algo
function query(comando: string): string;
var
  queryTemp: TSQLQuery;
begin
  try
    Form1.Connect.Open;
    queryTemp := TSQLQuery.Create(nil);
    queryTemp.SQLConnection := Form1.Connect;
    queryTemp.SQL.Clear;
    queryTemp.SQL.Add(comando);
    queryTemp.Open;

    Result := queryTemp.Fields[0].AsString;
  finally
    queryTemp.Free;
    Form1.Connect.Close;
  end;
end;


//Ao criar formulario, carregar dados
procedure TForm3.FormActivate(Sender: TObject);
var
  crt,cidade: string;
begin
  //Carregar valores do banco
  razaoSocial.Text := query('select d.razao from dados_empre d where d.codi='+empr.Text+';');
  nomeFantasia.Text := query('select d.nome_fantasia from dados_empre d where d.codi='+empr.Text+';');
  ende.Text := query('select d.endereco from dados_empre d where d.codi='+empr.Text+';');
  inscr.Text := query('select d.ie from dados_empre d where d.codi='+empr.Text+';');
  fone.Text := query('select d.fone1 from dados_empre d where d.codi='+empr.Text+';');
  mail.Text := query('select d.mail from dados_empre d where d.codi='+empr.Text+';');
  cnpj.Text := query('select d.cgc from dados_empre d where d.codi='+empr.Text+';');
  cep.Text := query('select d.cep from dados_empre d where d.codi='+empr.Text+';');
  bair.Text := query('select d.bairro from dados_empre d where d.codi='+empr.Text+';');
  crt := query('select d.crt from dados_empre d where d.codi='+empr.Text+';');

  cidade := query('select d.codi_cida from dados_empre d where d.codi='+empr.Text+';');
  cida.Text := query('select c.cid_desc from cidade c where c.cid_codi='+cidade+';');
  uf.Text := query('select c.cid_uf from cidade c where c.cid_codi='+cidade+';');

  if (crt='1') then
  begin
    simples.Checked := True;
  end
  else begin
    normal.Checked := True;
  end;

end;


//Bot�o para salvar no banco
procedure TForm3.saveClick(Sender: TObject);
var
  SQL: TSQLDataSet;
  crt: Integer;
  cgc,razao,fant: string;
begin

  //Simples ou Normal
  if simples.Checked then
  begin
    crt:=1;
  end
  else begin
    crt:=3;
  end;

  //Formatar CNPJ
  cgc := Trim(cnpj.Text);
  cgc := stringreplace(cgc, '-', '',[rfReplaceAll, rfIgnoreCase]);
  cgc := stringreplace(cgc, '/', '',[rfReplaceAll, rfIgnoreCase]);
  cgc := stringreplace(cgc, '.', '',[rfReplaceAll, rfIgnoreCase]);
  cgc := (Copy(cgc,1,2))+ '.' + (Copy(cgc,3,3)) + '.' + (Copy(cgc,6,3)) + '/' + (Copy(cgc,9,4)) + '-' + (Copy(cgc,13,2));

  //Tratar ' na razao social e nome fantasia
  razao := stringreplace(UpperCase(razaoSocial.Text), '''', ' ',[rfReplaceAll, rfIgnoreCase]);
  fant := stringreplace(UpperCase(nomeFantasia.Text), '''', ' ',[rfReplaceAll, rfIgnoreCase]);

  try
    try
      //Abrir conexoes
      Form1.Connect.Open;
      SQL := TSQLDataSet.Create(Nil);
      SQL.SQLConnection := Form1.Connect;

      //Executar COMANDO
      SQL.CommandText := 'UPDATE DADOS_EMPRE SET '+
                          'RAZAO = '+''''+razao+''''+','+
                          'ENDERECO = '+''''+UpperCase(ende.Text)+''''+','+
                          'CIDADE = '+''''+UpperCase(cida.Text)+''''+','+
                          'UF = '+''''+UpperCase(uf.Text)+''''+','+
                          'IE = '+''''+UpperCase(inscr.Text)+''''+','+
                          'FONE1 = '+''''+UpperCase(fone.Text)+''''+','+
                          'MAIL = '+''''+LowerCase(mail.Text)+''''+','+
                          'CGC = '+''''+cgc+''''+','+
                          'CEP = '+''''+UpperCase(cep.Text)+''''+','+
                          'NOME_FANTASIA = '+''''+fant+''''+','+
                          'CODI_CIDA = '+''''+IntToStr(TForm1.buscaCidade(UpperCase(cida.Text), UpperCase(uf.Text)))+''''+','+
                          'BAIRRO = '+''''+UpperCase(bair.Text)+''''+','+
                          'CRT = '+IntToStr(crt)+' '+
                      'WHERE (CODI = '+''''+empr.Text+''''+');';
      SQL.ExecSQL;

    except
      on e: exception do
      begin
        ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
      end;
    end;

  finally
    SQL.Free;
    Form1.Connect.Close;
    ShowMessage('Dados da empresa salvos.');
    Form3.Visible := False;
  end;
end;

end.
