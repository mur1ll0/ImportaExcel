unit uImportaGrupo;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaGrupo = class
    private
      colGrupo, dadosGrupo, condUpdateGrupo, dadosUpdateGrupo: String;

    public
      k: Integer;
      StringGrid1: TStringGrid;
      constructor ImportaRegistro(numReg: Integer; Grid: TStringGrid);
      procedure CarregaColunas;
      procedure Gravar;
  end;

implementation

uses
  importa_excel, importando;

constructor cImportaGrupo.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

procedure cImportaGrupo.CarregaColunas;
var
  i: Integer;
  temp: string;
begin
  frmImportando.atualizaStatus('Grupo '+IntToStr(k));

  colGrupo := '';
  dadosGrupo := '';

  //Carregar informa��es para importar
  //-------------------------------------------------------

  //Codigo � obrigat�rio, se n�o tiver preenche com o generator
  //CODI (CODIGO)
  i:=BuscaColuna(StringGrid1,'codi');
  if (i<>-1) then
  begin
    colGrupo := colGrupo + 'codi';
    StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll, rfIgnoreCase]);
    dadosGrupo := dadosGrupo + '''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colGrupo := colGrupo + 'codi';
    dadosGrupo := dadosGrupo + 'gen_id(gen_grup_prod_id,1)';
  end;

  //Empresa � obrigat�rio, se n�o tiver preenche com 1
  //EMPR (EMPRESA)
  i:=BuscaColuna(StringGrid1,'empr');
  if (i<>-1) then
  begin
    colGrupo := colGrupo + ',empr';
    dadosGrupo := dadosGrupo + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colGrupo := colGrupo + ',empr';
    dadosGrupo := dadosGrupo + ',''' + '1' + '''';
  end;


  for i := 0 to StringGrid1.ColCount-1 do
  begin
    //DESCR (DESCRICAO)
    if (LowerCase(StringGrid1.Cells[i,0])='descr') then
    begin
      colGrupo := colGrupo + ',descr';
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,30));
      dadosGrupo := dadosGrupo + ',''' + temp + '''';
    end
    ;

  //Fim do For das colunas
  end;
end;

procedure cImportaGrupo.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //----------------------------------------
  //Gravar no banco de GRUP_PROD
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL := TSQLDataSet.Create(Nil);
        SQL.SQLConnection := frmPrinc.conDestino;

        //Executar INSERT
        frmImportando.atualizaStatus('Inserindo dados na tabela GRUP_PROD.');
        SQL.CommandText := 'insert into grup_prod ('+ colGrupo +') values ' + '(' + dadosGrupo + ');';
        SQL.ExecSQL;

      except
        on e: exception do
        begin
          ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
          status := 0;
          Exit;
        end;
      end;

    finally
      SQL.Free;
      frmPrinc.conDestino.Close;
    end;
  end
  //Gravar comandos em TXT
  else if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
          (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then begin
    try
      try
        //Carregar arquivo TXT
        AssignFile(fileTXT, frmPrinc.DBPath.Text);
        if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
        else append(fileTXT);

        frmImportando.atualizaStatus('Comandos da GRUP_PROD.');
        WriteLn(fileTXT, '----------Comandos da GRUP_PROD----------');

        WriteLn(fileTXT, 'insert into grup_prod ('+ colGrupo +') values ' + '(' + dadosGrupo + ');');
        WriteLn(fileTXT, 'COMMIT WORK;');
      except
        on e: exception do
        begin
          ShowMessage('Erro TXT: '+e.message);
          status := 0;
          CloseFile(fileTXT);
          Exit;
        end;
      end;
    finally
      CloseFile(fileTXT);
    end;
  end;
end;

end.
