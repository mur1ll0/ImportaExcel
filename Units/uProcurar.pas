unit uProcurar;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TfrmProcurar = class(TForm)
    rgSelecao: TRadioGroup;
    Label1: TLabel;
    edtSearch: TEdit;
    btnProximo: TButton;
    btnAnterior: TButton;
    procedure FormShow(Sender: TObject);
    procedure btnProximoClick(Sender: TObject);
    procedure btnAnteriorClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    currentRow: Integer;

    function Buscar(Avancar: Boolean = True): Boolean;
  end;

var
  frmProcurar: TfrmProcurar;

implementation

uses
  importa_excel, uUtil;

{$R *.dfm}

procedure TfrmProcurar.btnAnteriorClick(Sender: TObject);
var
  encontrou: Boolean;
  but: Integer;
begin
  //Salvar tabela em memoria para permitir CTRL+Z
  frmPrinc.StringGridToArray(frmPrinc.StringGrid1);

  //Receber n�mero da linha atual
  currentRow := frmPrinc.StringGrid1.Row-1;

  //Buscar
  encontrou := Buscar(False);

  //Se n�o encontrou pede pra reiniciar
  if not encontrou then
  begin
    but := Mensagem('Termo n�o encontrado. Deseja recome�ar a busca do fim?', mtCustom, [mbYes, mbNo],['Sim','N�o'], 'N�o encontrado.');
    //Sim
    if (but = 6) then
    begin
      //Recome�ar do in�cio
      currentRow := frmPrinc.StringGrid1.RowCount-1;
      encontrou := Buscar(False);
      if not encontrou then
      begin
        ShowMessage('Termo n�o encontrado na tabela.');
      end;
    end;
  end;
end;

procedure TfrmProcurar.btnProximoClick(Sender: TObject);
var
  encontrou: Boolean;
  but: Integer;
begin
  //Salvar tabela em memoria para permitir CTRL+Z
  frmPrinc.StringGridToArray(frmPrinc.StringGrid1);

  //Receber n�mero da linha atual
  currentRow := frmPrinc.StringGrid1.Row+1;

  //Buscar
  encontrou := Buscar;

  //Se n�o encontrou pede pra reiniciar
  if not encontrou then
  begin
    but := Mensagem('Termo n�o encontrado. Deseja recome�ar a busca do inc�o?', mtCustom, [mbYes, mbNo],['Sim','N�o'], 'N�o encontrado.');
    //Sim
    if (but = 6) then
    begin
      //Recome�ar do in�cio
      currentRow := 0;
      encontrou := Buscar;
      if not encontrou then
      begin
        ShowMessage('Termo n�o encontrado na tabela.');
      end;
    end;
  end;
end;

procedure TfrmProcurar.FormShow(Sender: TObject);
begin
  edtSearch.Text := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row];
  edtSearch.SetFocus;
end;

function TfrmProcurar.Buscar(Avancar: Boolean = True): Boolean;
var
  i, j: Integer;
  encontrou: Boolean;
begin
  encontrou := False;

  //Buscar avan�ando
  if Avancar then
  begin
    case rgSelecao.ItemIndex of
      //Coluna
      0: begin
        for i := currentRow to frmPrinc.StringGrid1.RowCount-1 do
        begin
          if (frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i].IndexOf(edtSearch.Text) > 0) or
             (frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i] = edtSearch.Text) then
          begin
            //Setar foco na linha e coluna encontrada
            frmPrinc.StringGrid1.Row := i;
            frmPrinc.StringGrid1.SetFocus;
            encontrou := True;
            Break;
          end;
        end;
      end;
      //Tudo
      1: begin
        for i := currentRow to frmPrinc.StringGrid1.RowCount-1 do
        begin
          for j := 0 to frmPrinc.StringGrid1.ColCount-1 do
          begin
            if (frmPrinc.StringGrid1.Cells[j, i].IndexOf(edtSearch.Text) > 0) or
               (frmPrinc.StringGrid1.Cells[j, i] = edtSearch.Text) then
            begin
              //Setar foco na linha e coluna encontrada
              frmPrinc.StringGrid1.Row := i;
              frmPrinc.StringGrid1.Col := j;
              frmPrinc.StringGrid1.SetFocus;
              encontrou := True;
              Break;
            end;
          end;
          if encontrou then
            Break;
        end;
      end;
    end;
  end
  //Buscar regredindo
  else
  begin
    case rgSelecao.ItemIndex of
      //Coluna
      0: begin
        for i := currentRow downto 0 do
        begin
          if (frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i].IndexOf(edtSearch.Text) > 0) or
             (frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i] = edtSearch.Text) then
          begin
            //Setar foco na linha e coluna encontrada
            frmPrinc.StringGrid1.Row := i;
            frmPrinc.StringGrid1.SetFocus;
            encontrou := True;
            Break;
          end;
        end;
      end;
      //Tudo
      1: begin
        for i := currentRow downto 0 do
        begin
          for j := 0 to frmPrinc.StringGrid1.ColCount-1 do
          begin
            if (frmPrinc.StringGrid1.Cells[j, i].IndexOf(edtSearch.Text) > 0) or
               (frmPrinc.StringGrid1.Cells[j, i] = edtSearch.Text) then
            begin
              //Setar foco na linha e coluna encontrada
              frmPrinc.StringGrid1.Row := i;
              frmPrinc.StringGrid1.Col := j;
              frmPrinc.StringGrid1.SetFocus;
              encontrou := True;
              Break;
            end;
          end;
          if encontrou then
            Break;
        end;
      end;
    end;
  end;

  Result := encontrou;
end;

end.
