unit empresa;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

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

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}



end.
