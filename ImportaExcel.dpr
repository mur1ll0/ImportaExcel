program ImportaExcel;

uses
  Vcl.Forms,
  importa_excel in 'importa_excel.pas' {frmPrinc},
  importando in 'importando.pas' {frmImportando},
  empresa in 'empresa.pas' {frmEmpr},
  Colunas in 'Colunas.pas' {frmColunas},
  uSubstituir in 'uSubstituir.pas' {frmSubstituir},
  uUtil in 'Units\uUtil.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmPrinc, frmPrinc);
  Application.CreateForm(TfrmImportando, frmImportando);
  Application.CreateForm(TfrmEmpr, frmEmpr);
  Application.CreateForm(TfrmColunas, frmColunas);
  Application.CreateForm(TfrmSubstituir, frmSubstituir);
  Application.Run;
end.
