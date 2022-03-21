unit uImportaGrade;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaGrades = class
    private
      colGrades, dadosGrades, condUpdateGrades, dadosUpdateGrades: string;
      colLanc, dadosLanc: string;
      colLancItems, dadosLancItems: string;

      function VerificaLinha(linhaDesc: string; gradeId: string): string;
      function VerificaColuna(colunaDesc: string; gradeId: string): string;
      procedure VerificaGradeEstoque(grade: string; prod: string; linha: string; coluna: string; empr: string);
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

constructor cImportaGrades.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

function cImportaGrades.VerificaColuna(colunaDesc, gradeId: string): string;
var
  queryTemp: TSQLQuery;
  str: string;
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //Salvar em arquivo
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then
  begin
    //Comando para tentar inserir valor da COLUNA no banco de dados
    //Se ja existir, vai dar erro no insert e não vai inserir
    str := 'insert into grade_coluna (grco_id,grad_id,grco_descricao) '+
      'select '+
      '    case '+
      '        when (select gc1.grco_id from grade_coluna gc1 where gc1.grad_id='+gradeId+' and gc1.grco_descricao='+QuotedStr(colunaDesc)+') is null '+
      '            then gen_id(GEN_GRADE_COLUNA_ID,1) '+
      '        else '+
      '            (select gc1.grco_id from grade_coluna gc1 where gc1.grad_id='+gradeId+' and gc1.grco_descricao='+QuotedStr(colunaDesc)+') '+
      '    end CODI_COLUNA, '+
      '    '+gradeId+' CODI_GRADE, '+
      '    '+QuotedStr(colunaDesc)+' DESCR '+
      'from RDB$DATABASE; ';
    try
      try
        //Carregar arquivo TXT
        AssignFile(fileTXT, frmPrinc.DBPath.Text);
        if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
        else append(fileTXT);

        frmImportando.atualizaStatus('Inserindo coluna '+colunaDesc);
        WriteLn(fileTXT, '----------Inserindo coluna '+colunaDesc+'----------');

        WriteLn(fileTXT, str);
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
    //Retornar valor do código da coluna (no caso o SQL que chega até ele)
    str := 'select '+
      '    case '+
      '        when (select gc1.grco_id from grade_coluna gc1 where gc1.grad_id='+gradeId+' and gc1.grco_descricao='+QuotedStr(colunaDesc)+') is null '+
      '            then gen_id(GEN_GRADE_COLUNA_ID,1) '+
      '        else '+
      '            (select gc1.grco_id from grade_coluna gc1 where gc1.grad_id='+gradeId+' and gc1.grco_descricao='+QuotedStr(colunaDesc)+') '+
      '    end CODI_COLUNA '+
      'from RDB$DATABASE ';
    Result := '('+str+')';
  end
  //Executar no banco de dados
  else
  begin
    try
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('select gc1.grco_id from grade_coluna gc1 where gc1.grad_id= :PGRADE and gc1.grco_descricao= :PDESC');
      queryTemp.ParamByName('PGRADE').AsString := gradeId;
      queryTemp.ParamByName('PDESC').AsString := colunaDesc;
      queryTemp.Open;
    finally
      if queryTemp.IsEmpty then
      begin
        //Inserir coluna e carregar código
        try
          try
            //Abrir conexoes
            frmPrinc.conDestino.Open;
            SQL := TSQLDataSet.Create(Nil);
            SQL.SQLConnection := frmPrinc.conDestino;

            SQL.CommandText := 'insert into grade_coluna (grco_id,grad_id,grco_descricao) values (gen_id(GEN_GRADE_COLUNA_ID,1),'+gradeId+','+QuotedStr(colunaDesc)+');';
            SQL.ExecSQL;

          except
            on e: exception do
            begin
              ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
              status := 0;
            end;
          end;

        finally
          SQL.Free;
          frmPrinc.conDestino.Close;
        end;

        //Para ter o resultado, executar denovo
        queryTemp.Close;
        queryTemp.Open;
        Result := queryTemp.FieldByName('grco_id').AsString;
      end
      else begin
        //Carregar código
        Result := queryTemp.FieldByName('grco_id').AsString;
      end;
      queryTemp.Close;
      frmPrinc.conDestino.Close;
    end;
  end;
end;

procedure cImportaGrades.VerificaGradeEstoque(grade, prod, linha,
  coluna: string; empr: string);
var
  queryTemp: TSQLQuery;
  str: string;
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //Salvar em arquivo
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then
  begin
    //Comando para tentar inserir na GRADE_ESTOQUE
    //Se já existir o registro, vai dar erro e não executar o comando apenas.
    str := 'insert into grade_estoque (gres_id, gres_qtd, gres_qtd_max, grad_id, prod_codi, prod_empr, grli_id, grco_id, empr) '+
      'select '+
      '    case '+
      '        when (select ge.gres_id from grade_estoque ge where ge.grad_id='+grade+' and ge.prod_codi='+prod+' and ge.empr='+empr+' and ge.grli_id='+linha+' and ge.grco_id='+coluna+') is null '+
      '            then gen_id(GEN_GRADE_ESTOQUE_ID,1) '+
      '        else '+
      '            (select ge.gres_id from grade_estoque ge where ge.grad_id='+grade+' and ge.prod_codi='+prod+' and ge.empr='+empr+' and ge.grli_id='+linha+' and ge.grco_id='+coluna+') '+
      '    end ID, '+
      '    0.0 QTD, '+
      '    0.0 QTD_MAX, '+
      '    '+grade+' GRADE_ID, '+
      '    '+prod+' PROD_CODI, '+
      '    '+empr+' PROD_EMPR, '+
      '    '+linha+' LINHA, '+
      '    '+coluna+' COLUNA, '+
      '    '+empr+' EMPR '+
      'from RDB$DATABASE; ';
    try
      try
        //Carregar arquivo TXT
        AssignFile(fileTXT, frmPrinc.DBPath.Text);
        if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
        else append(fileTXT);

        frmImportando.atualizaStatus('Inserindo registro na GRADE_ESTOQUE');
        WriteLn(fileTXT, '----------Inserindo registro na GRADE_ESTOQUE----------');

        WriteLn(fileTXT, str);
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
  end
  //Executar no banco de dados
  else
  begin
    try
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('select ge.gres_id from grade_estoque ge where ge.grad_id='+grade+' and ge.prod_codi='+prod+' and ge.empr='+empr+' and ge.grli_id='+linha+' and ge.grco_id='+coluna);
      queryTemp.Open;
    finally
      if queryTemp.IsEmpty then
      begin
        //Se não existir, inserir registro
        try
          try
            //Abrir conexoes
            frmPrinc.conDestino.Open;
            SQL := TSQLDataSet.Create(Nil);
            SQL.SQLConnection := frmPrinc.conDestino;

            SQL.CommandText := 'insert into grade_estoque (gres_id, gres_qtd, gres_qtd_max, grad_id, prod_codi, prod_empr, grli_id, grco_id, empr) values '+
              ' (gen_id(GEN_GRADE_ESTOQUE_ID,1),0.0,0.0,'+grade+','+prod+','+empr+','+linha+','+coluna+','+empr+');';
            SQL.ExecSQL;

          except
            on e: exception do
            begin
              ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
              status := 0;
            end;
          end;

        finally
          SQL.Free;
          frmPrinc.conDestino.Close;
        end;
      end;
      queryTemp.Close;
      frmPrinc.conDestino.Close;
    end;
  end;
end;

function cImportaGrades.VerificaLinha(linhaDesc: string; gradeId: string): string;
var
  queryTemp: TSQLQuery;
  str: string;
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //Salvar em arquivo
  if (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT') or
     (UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.SQL')
  then
  begin
    //Comando para tentar inserir valor da LINHA no banco de dados
    //Se ja existir, vai dar erro no insert e não vai inserir
    str := 'insert into grade_linha (grli_id,grad_id,grli_descricao) '+
      'select '+
      '    case '+
      '        when (select gl1.grli_id from grade_linha gl1 where gl1.grad_id='+gradeId+' and gl1.grli_descricao='+QuotedStr(linhaDesc)+') is null '+
      '            then gen_id(GEN_GRADE_LINHA_ID,1) '+
      '        else '+
      '            (select gl1.grli_id from grade_linha gl1 where gl1.grad_id='+gradeId+' and gl1.grli_descricao='+QuotedStr(linhaDesc)+') '+
      '    end CODI_LINHA, '+
      '    '+gradeId+' CODI_GRADE, '+
      '    '+QuotedStr(linhaDesc)+' DESCR '+
      'from RDB$DATABASE; ';
    try
      try
        //Carregar arquivo TXT
        AssignFile(fileTXT, frmPrinc.DBPath.Text);
        if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
        else append(fileTXT);

        frmImportando.atualizaStatus('Inserindo linha '+linhaDesc);
        WriteLn(fileTXT, '----------Inserindo linha '+linhaDesc+'----------');

        WriteLn(fileTXT, str);
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
    //Retornar valor do código da linha (no caso o SQL que chega até ele)
    str := 'select '+
      '    case '+
      '        when (select gl1.grli_id from grade_linha gl1 where gl1.grad_id='+gradeId+' and gl1.grli_descricao='+QuotedStr(linhaDesc)+') is null '+
      '            then gen_id(GEN_GRADE_LINHA_ID,0) '+
      '        else '+
      '            (select gl1.grli_id from grade_linha gl1 where gl1.grad_id=20 and gl1.grli_descricao='+QuotedStr(linhaDesc)+') '+
      '    end CODI_LINHA '+
      'from RDB$DATABASE ';
    Result := '('+str+')';
  end
  //Executar no banco de dados
  else
  begin
    try
      frmPrinc.conDestino.Open;
      queryTemp := TSQLQuery.Create(nil);
      queryTemp.SQLConnection := frmPrinc.conDestino;
      queryTemp.SQL.Clear;
      queryTemp.SQL.Add('select gl1.grli_id from grade_linha gl1 where gl1.grad_id= :PGRADE and gl1.grli_descricao= :PDESC');
      queryTemp.ParamByName('PGRADE').AsString := gradeId;
      queryTemp.ParamByName('PDESC').AsString := linhaDesc;
      queryTemp.Open;
    finally
      if queryTemp.IsEmpty then
      begin
        //Inserir linha e carregar código
        try
          try
            //Abrir conexoes
            frmPrinc.conDestino.Open;
            SQL := TSQLDataSet.Create(Nil);
            SQL.SQLConnection := frmPrinc.conDestino;

            SQL.CommandText := 'insert into grade_linha (grli_id,grad_id,grli_descricao) values (gen_id(GEN_GRADE_LINHA_ID,1),'+gradeId+','+QuotedStr(linhaDesc)+');';
            SQL.ExecSQL;

          except
            on e: exception do
            begin
              ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
              status := 0;
            end;
          end;

        finally
          SQL.Free;
          frmPrinc.conDestino.Close;
        end;

        //Para ter o resultado, executar denovo
        queryTemp.Close;
        queryTemp.Open;
        Result := queryTemp.FieldByName('grli_id').AsString;
      end
      else begin
        //Carregar código
        Result := queryTemp.FieldByName('grli_id').AsString;
      end;
      queryTemp.Close;
      frmPrinc.conDestino.Close;
    end;
  end;
end;

procedure cImportaGrades.CarregaColunas;
var
  i: Integer;
  temp: string;
  gradeId: string;
  prodId: string;
  linhaId, colunaId: string;
  but: Integer;
  empr: string;
begin
  frmImportando.atualizaStatus('Grades '+IntToStr(k));

  colLanc := '';
  dadosLanc := '';
  colLancItems := '';
  dadosLancItems := '';

  //Carregar informações para importar
  //-------------------------------------------------------

  //GRADE (Código da Grade)
  i:=BuscaColuna(StringGrid1,'grade');
  if (i<>-1) then
  begin
    StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll, rfIgnoreCase]);
    if StringGrid1.Cells[i,k] <> '' then
      gradeId := StringGrid1.Cells[i,k];
  end;

  //PROD (Produto vinculado a grade)
  i:=BuscaColuna(StringGrid1,'prod');
  if (i<>-1) then
  begin
    StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll, rfIgnoreCase]);
    prodId := StringGrid1.Cells[i,k];

    //Se não tiver código do produto sair
    if prodId = '' then
    begin
      ShowMessage('Nenhum produto informado.');
      status := 0;
      Exit;
    end;

    //Se não tiver gradeId informado, carregado do produto
    if gradeId = '' then
      gradeId := '(select grade_grad_id from prod where codi='+prodId+')';
  end
  else
  begin
    ShowMessage('Nenhum produto informado.');
    status := 0;
    Exit;
  end;

  //Generators
  colLanc := colLanc + 'grla_id';
  dadosLanc := dadosLanc + 'gen_id(GEN_GRADE_LANCAMENTO_ID,1)';
  colLancItems := colLancItems + 'grit_id';
  dadosLancItems := dadosLancItems + 'gen_id(GEN_GRADE_LANCAMENTO_ITENS_ID,1)';
  colLancItems := colLancItems + ',grla_id';
  dadosLancItems := dadosLancItems + ',' + 'gen_id(GEN_GRADE_LANCAMENTO_ID,0)';

  //Se não tiver código da grade sair
  if gradeId = '' then
  begin
    ShowMessage('Nenhuma grade encontrada.');
    status := 0;
    Exit;
  end
  //Se tem grade, adiciona Grade e Produto aos comandos
  else
  begin
    //Grade
    colLanc := colLanc + ',grade_grad_id';
    dadosLanc := dadosLanc + ',' + gradeId;

    //Produto
    colLanc := colLanc + ',prod_codi';
    dadosLanc := dadosLanc + ',' + QuotedStr(prodId);
  end;

  //EMPR (EMPRESA)
  i:=BuscaColuna(StringGrid1,'empr');
  if (i<>-1) then
  begin
    temp := StringGrid1.Cells[i,k];
    if temp = '' then
      temp := '1';
    empr := temp;
    colLanc := colLanc + ',prod_empr';
    dadosLanc := dadosLanc + ',''' + empr + '''';
    colLanc := colLanc + ',grla_empr';
    dadosLanc := dadosLanc + ',''' + empr + '''';
  end
  else begin
    colLanc := colLanc + ',prod_empr';
    dadosLanc := dadosLanc + ',1';
    colLanc := colLanc + ',grla_empr';
    dadosLanc := dadosLanc + ',1';
  end;

  //LINHA (Informação da linha da grade)
  i:=BuscaColuna(StringGrid1,'linha');
  if (i<>-1) then
  begin
    temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
    temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
    temp := (Copy(temp,1,32));
    linhaId := VerificaLinha(temp, gradeId);

    if linhaId = '' then
    begin
      ShowMessage('Linha não encontrada.');
      status := 0;
      Exit;
    end
    else
    begin
      colLancItems := colLancItems + ',grli_id';
      dadosLancItems := dadosLancItems + ',' + linhaId;
    end;
  end;

  //COLUNA (Informação da coluna da grade)
  i:=BuscaColuna(StringGrid1,'coluna');
  if (i<>-1) then
  begin
    temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
    temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
    temp := (Copy(temp,1,32));
    colunaId := VerificaColuna(temp, gradeId);

    if colunaId = '' then
    begin
      ShowMessage('Coluna não encontrada.');
      status := 0;
      Exit;
    end
    else
    begin
      colLancItems := colLancItems + ',grco_id';
      dadosLancItems := dadosLancItems + ',' + colunaId;
    end;
  end;

  //Verificar se existe registro na tebela GRADE_ESTOQUE, sen criar
  VerificaGradeEstoque(gradeId,prodId,linhaId,colunaId,empr);

  //Quantidade é obrigatório, se não tiver põe 0
  //QTD
  i:=BuscaColuna(StringGrid1,'qtd');
  if (i<>-1) then
  begin
    if StringGrid1.Cells[i,k]='' then
    begin
      temp := '0';
    end
    else begin
      temp := StringGrid1.Cells[i,k];
    end;
    temp := corrigeFloat(temp);

    //Quantidade
    colLancItems := colLancItems + ',grit_qtd';
    dadosLancItems := dadosLancItems + ',' + temp;
    //Tipo
    colLancItems := colLancItems + ',grit_tipo';
    dadosLancItems := dadosLancItems + ',' + '6';
    //Saldo
    colLancItems := colLancItems + ',grit_qtd_saldo';
    dadosLancItems := dadosLancItems + ',' + temp;
  end;

  //Campo adicionais padrões
  //GRLA_ORIGEM_TABELA (Tabela de origem)
  colLanc := colLanc + ',GRLA_ORIGEM_TABELA';
  dadosLanc := dadosLanc + ',' + QuotedStr('GRADE_AJUS');
  //GRLA_ORIGEM_ID (Tabela de origem)
  colLanc := colLanc + ',GRLA_ORIGEM_ID';
  dadosLanc := dadosLanc + ',' + '(select gen_id(GEN_GRADE_AJUS_ID,0)+1 from RDB$DATABASE)';
  //GRLA_ORIGEM_TABELA_CABECALHO (Um nome que ta ai e n sei pq)
  colLanc := colLanc + ',GRLA_ORIGEM_TABELA_CABECALHO';
  dadosLanc := dadosLanc + ',' + QuotedStr('GRADE_ESTOQUE');
  //DATA (Data do lançamento, por padrão vou usar o dia atual)
  colLanc := colLanc + ',GRLA_DATA';
  dadosLanc := dadosLanc + ',' + 'CURRENT_DATE';
end;

procedure cImportaGrades.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //Sair se estiver em erro
  if status = 0 then
    Exit;

  //----------------------------------------
  //Gravar no banco as Grades
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL := TSQLDataSet.Create(Nil);
        SQL.SQLConnection := frmPrinc.conDestino;

        //Executar INSERT
        frmImportando.atualizaStatus('Inserindo dados na tabela GRADE_LANCAMENTO.');
        SQL.CommandText := 'insert into GRADE_LANCAMENTO ('+ colLanc +') values ' + '(' + dadosLanc + ');';
        SQL.ExecSQL;

        frmImportando.atualizaStatus('Inserindo dados na tabela GRADE_LANCAMENTO_ITENS.');
        SQL.CommandText := 'insert into GRADE_LANCAMENTO_ITENS ('+ colLancItems +') values ' + '(' + dadosLancItems + ');';
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

        frmImportando.atualizaStatus('Comandos da GRADE_LANCAMENTO.');
        WriteLn(fileTXT, '----------Comandos da GRADE_LANCAMENTO----------');

        WriteLn(fileTXT, 'insert into GRADE_LANCAMENTO ('+ colLanc +') values ' + '(' + dadosLanc + ');');
        WriteLn(fileTXT, 'COMMIT WORK;');

        frmImportando.atualizaStatus('Comandos da GRADE_LANCAMENTO_ITENS.');
        WriteLn(fileTXT, '----------Comandos da GRADE_LANCAMENTO_ITENS----------');

        WriteLn(fileTXT, 'insert into GRADE_LANCAMENTO_ITENS ('+ colLancItems +') values ' + '(' + dadosLancItems + ');');
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
