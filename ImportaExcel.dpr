program ImportaExcel;

uses
  Vcl.Forms,
  importa_excel in 'importa_excel.pas' {Form1},
  importando in 'importando.pas' {Form2},
  empresa in 'empresa.pas' {Form3};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TForm3, Form3);
  Application.Run;
end.
