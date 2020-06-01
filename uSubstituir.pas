unit uSubstituir;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TfrmSubstituir = class(TForm)
    rgSelecao: TRadioGroup;
    Label1: TLabel;
    edtOldValue: TEdit;
    Label2: TLabel;
    edtNewValue: TEdit;
    btnReplace: TButton;
    procedure btnReplaceClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSubstituir: TfrmSubstituir;

implementation

uses importa_excel;

{$R *.dfm}

procedure TfrmSubstituir.btnReplaceClick(Sender: TObject);
var
  temp: string;
  i, j: Integer;
begin
  case rgSelecao.ItemIndex of
    //Célula
    0: begin
      temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row];
      temp := stringreplace(temp, edtOldValue.Text, edtNewValue.Text,[rfReplaceAll, rfIgnoreCase]);
      frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row] := temp;
    end;
    //Coluna
    1: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        temp := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i];
        temp := stringreplace(temp, edtOldValue.Text, edtNewValue.Text,[rfReplaceAll, rfIgnoreCase]);
        frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col, i] := temp;
      end;
    end;
    //Tudo
    2: begin
      for i := 1 to frmPrinc.StringGrid1.RowCount-1 do begin
        for j := 0 to frmPrinc.StringGrid1.ColCount-1 do begin
          temp := frmPrinc.StringGrid1.Cells[j, i];
          temp := stringreplace(temp, edtOldValue.Text, edtNewValue.Text,[rfReplaceAll, rfIgnoreCase]);
          frmPrinc.StringGrid1.Cells[j, i] := temp;
        end;
      end;
    end;
  end;
end;

procedure TfrmSubstituir.FormShow(Sender: TObject);
begin
  edtOldValue.Text := frmPrinc.StringGrid1.Cells[frmPrinc.StringGrid1.Col,frmPrinc.StringGrid1.Row];

  if edtOldValue.Text = '' then
    edtOldValue.SetFocus
  else edtNewValue.SetFocus;
end;

end.
