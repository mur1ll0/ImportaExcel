unit importando;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons;


type
  TForm2 = class(TForm)
    BitBtn1: TBitBtn;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    procedure fim(status: integer);
    procedure atualizaItens(min: integer; max: integer);
    procedure atualizaStatus(status: string);
    procedure BitBtn1Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

procedure TForm2.atualizaItens(min: integer; max: integer);
begin
  Form2.Label2.Caption := IntToStr(min) + ' / ' + IntToStr(max);
  Form2.Label2.Left := (Width - Form2.Label2.Width ) div 2;
  //Form2.Label2.Top := (Height - Form2.Label2.Height) div 2;
  Application.ProcessMessages; //Isso aqui faz a magica da atualizacao na tela acontecer
end;

procedure TForm2.atualizaStatus(status: string);
begin
  Form2.Label3.Caption := status;
  Form2.Label3.Left := (Width - Form2.Label3.Width ) div 2;
  //Form2.Label3.Top := (Height - Form2.Label3.Height) div 2;
  Application.ProcessMessages; //Isso aqui faz a magica da atualizacao na tela acontecer
end;

procedure TForm2.BitBtn1Click(Sender: TObject);
begin
  Form2.Close;
end;

procedure TForm2.fim(status: integer);
begin
  if status = 1 then
  begin
    Form2.Label2.Caption := Form2.Label2.Caption + ' - Concluído!';
    Form2.Label2.Font.Color := clGreen;
    Form2.BitBtn1.Caption := 'Fechar';
  end
  else begin
    Form2.Label2.Caption := Form2.Label2.Caption + ' - Erro.';
    Form2.Label2.Font.Color := clRed;
    Form2.BitBtn1.Caption := 'Fechar';
  end;
  Form2.Label2.Left := (Width - Form2.Label2.Width ) div 2;
  //Form2.Label2.Top := (Height - Form2.Label2.Height) div 2;
end;

end.
