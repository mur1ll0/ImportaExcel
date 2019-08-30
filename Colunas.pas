unit Colunas;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Clipbrd;

type
  TForm4 = class(TForm)
    LabelTipoImp: TLabel;
    ListColunas: TListBox;
    procedure FormShow(Sender: TObject);
    procedure ListColunasKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;

implementation

{$R *.dfm}

procedure TForm4.FormShow(Sender: TObject);
begin
  //Se for Clientes/Fornecedores
  if LabelTipoImp.Caption = 'Clie/Forn' then
  begin
    ListColunas.Items.Clear;
    ListColunas.Items.add('CODI');
    ListColunas.Items.add('UF');
    ListColunas.Items.add('CIDA');
  end
  else if LabelTipoImp.Caption = 'Grupos' then
  begin

  end

  ;


  {

SubGrupos
Marcas
Produtos
T�tulos a Pagar
T�tulos a Receber
}

end;


procedure TForm4.ListColunasKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   //RECONHECER CTRL+C
   if ((Shift = [ssCtrl]) and (Key = 67)) then
   begin
     ClipBoard.AsText := ListColunas.Items[ListColunas.ItemIndex];
   end;
end;

end.