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


//Botão Cancelar (Sair) sem matar a tela, para que as informações permanecam ali
procedure TForm3.cancelClick(Sender: TObject);
begin
  Form3.Visible := False;
end;


//Botão para salvar no banco
procedure TForm3.saveClick(Sender: TObject);
var
  SQL: TSQLDataSet;
  crt: Integer;
  crc: string;
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
  crc := Trim(cnpj.Text);
  crc := (Copy(crc,1,2))+ '.' + (Copy(crc,3,3)) + '.' + (Copy(crc,6,3)) + '/' + (Copy(crc,9,4)) + '-' + (Copy(crc,13,2));

  try
    try
      //Abrir conexoes
      Form1.Connect.Open;
      SQL := TSQLDataSet.Create(Nil);
      SQL.SQLConnection := Form1.Connect;

      //Executar COMANDO
      SQL.CommandText := 'UPDATE DADOS_EMPRE SET '+
                          'RAZAO = '+''''+UpperCase(razaoSocial.Text)+''''+','+
                          'ENDERECO = '+''''+UpperCase(ende.Text)+''''+','+
                          'CIDADE = '+''''+UpperCase(cida.Text)+''''+','+
                          'IE = '+''''+UpperCase(inscr.Text)+''''+','+
                          'FONE1 = '+''''+UpperCase(fone.Text)+''''+','+
                          'MAIL = '+''''+UpperCase(mail.Text)+''''+','+
                          'CGC = '+''''+crc+''''+','+
                          'CEP = '+''''+UpperCase(cep.Text)+''''+','+
                          'NOME_FANTASIA = '+''''+UpperCase(nomeFantasia.Text)+''''+','+
                          'CODI_CIDA = '+''''+IntToStr(TForm1.buscaCidade(UpperCase(cida.Text), UpperCase(uf.Text)))+''''+','+
                          'BAIRRO = '+''''+UpperCase(bair.Text)+''''+','+
                          'CRT = '+IntToStr(crt)+','+
                      'WHERE (CODI = '+''''+empr.Text+''''+',);';
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
  end;
end;

end.
