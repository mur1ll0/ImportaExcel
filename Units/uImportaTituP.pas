unit uImportaTituP;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaTituP = class
    private
      colTituP, dadosTituP, condUpdateTituP, dadosUpdateTituP: string;
      colBTitu, dadosBTitu, condUpdateBTitu, dadosUpdateBTitu: string;
      saldo: Double;

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

constructor cImportaTituP.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

procedure cImportaTituP.CarregaColunas;
var
  i, l, j, count: Integer;
  temp, temp2: string;
begin
  frmImportando.atualizaStatus('T�tulos a Pagar '+IntToStr(k));

  colTituP := '';
  dadosTituP := '';
  colBTitu := '';
  dadosBTitu := '';

  //Carregar informa��es para importar
  //-------------------------------------------------------

  //CODI (CODIGO)
  i:=BuscaColuna(StringGrid1,'codi');
  if (i<>-1) then
  begin
    StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll, rfIgnoreCase]);
    StringGrid1.Cells[i,k] := (Copy(StringGrid1.Cells[i,k],1,12));
    l := Length(StringGrid1.Cells[i,k]);
    //Testar se ja existir o c�digo do t�tulo e inserir uma barra.
    count := 0;
    while (frmPrinc.temCodTituloP(StringGrid1.Cells[i,k]) = True) do
    begin
      for j := 1 to count do
      begin
        StringGrid1.Cells[i,k] := stringreplace(StringGrid1.Cells[i,k], '/'+IntToStr(j), '',[rfReplaceAll, rfIgnoreCase]);
        StringGrid1.Cells[i,k] := (Copy(StringGrid1.Cells[i,k],1,l));
      end;

      count := count+1;
      StringGrid1.Cells[i,k] := StringGrid1.Cells[i,k] + '/' + IntToStr(count);
    end;

    colTituP := colTituP + 'codi';
    dadosTituP := dadosTituP + '''' + StringGrid1.Cells[i,k] + '''';
    colBTitu := colBTitu + 'codi';
    dadosBTitu := dadosBTitu + '''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituP := colTituP + 'codi';
    dadosTituP := dadosTituP + IntToStr(k);
    colBTitu := colBTitu + 'codi';
    dadosBTitu := dadosBTitu + IntToStr(k);
  end;

  //EMPR (EMPRESA)
  i:=BuscaColuna(StringGrid1,'empr');
  if (i<>-1) then
  begin
    colTituP := colTituP + ',empr';
    dadosTituP := dadosTituP + ',''' + StringGrid1.Cells[i,k] + '''';
    colBTitu := colBTitu + ',empr';
    dadosBTitu := dadosBTitu + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituP := colTituP + ',empr';
    dadosTituP := dadosTituP + ',1';
    colBTitu := colBTitu + ',empr';
    dadosBTitu := dadosBTitu + ',1';
  end;

  //FORN (FORNECEDOR)
  i:=BuscaColuna(StringGrid1,'forn');
  if (i<>-1) then
  begin
    temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
    temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
    temp := (Copy(temp,1,60));
    //Se for letras, buscar c�digo.
    if not (IsNumeric(temp)) then
    begin
      temp2 := IntToStr(frmPrinc.getCodiClieForn(''''+temp+''''));
      //Se n�o encontrar a string, cadastrar fornecedor
      if temp2='0' then begin
        temp2 := IntToStr(frmPrinc.cadastraClieForn('nome',''''+temp+''''));
      end;
      colTituP := colTituP + ',forn';
      dadosTituP := dadosTituP + ',''' + temp2 + '''';
      colBTitu := colBTitu + ',forn';
      dadosBTitu := dadosBTitu + ',''' + temp2 + '''';
    end
    else begin
      //Se for n�meros, considera como c�digo
      temp2 := IntToStr(frmPrinc.getCodiClieForn('(select c.nome from clieforn c where c.codi = '+temp+')'));
      //Antes buscamos se existe o c�digo cadastrado, se n�o encontrar colocamos o generator mesmo
      if temp2='0' then begin
        colTituP := colTituP + ',forn';
        dadosTituP := dadosTituP + ',' + 'gen_id(gen_clieforn_id,0)';
        colBTitu := colBTitu + ',forn';
        dadosBTitu := dadosBTitu + ',' + 'gen_id(gen_clieforn_id,0)';
      end
      else begin
        //Se achar o c�digo, usamos o c�digo
        colTituP := colTituP + ',forn';
        dadosTituP := dadosTituP + ',''' + temp + '''';
        colBTitu := colBTitu + ',forn';
        dadosBTitu := dadosBTitu + ',''' + temp + '''';
      end;
    end;
  end
  else begin
    //Se n�o tiver fornecedor, colocar o generator.
    colTituP := colTituP + ',forn';
    dadosTituP := dadosTituP + ',' + 'gen_id(gen_clieforn_id,0)';
    colBTitu := colBTitu + ',forn';
    dadosBTitu := dadosBTitu + ',' + 'gen_id(gen_clieforn_id,0)';
  end;

  //LOCA_COBR (LOCAL DE COBRAN�A)
  i:=BuscaColuna(StringGrid1,'loca_cobr');
  if (i<>-1) then
  begin
    colTituP := colTituP + ',loca_cobr';
    dadosTituP := dadosTituP + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituP := colTituP + ',loca_cobr';
    dadosTituP := dadosTituP + ',''' + '1' + '''';
  end;

  //CART (TIPO DE CARTEIRA)
  i:=BuscaColuna(StringGrid1,'cart');
  if (i<>-1) then
  begin
    colTituP := colTituP + ',cart';
    dadosTituP := dadosTituP + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituP := colTituP + ',cart';
    dadosTituP := dadosTituP + ',''' + '1' + '''';
  end;

  //OPER (OPERA��O DO PLANO DE CONTAS)
  i:=BuscaColuna(StringGrid1,'oper');
  if (i<>-1) then
  begin
    colTituP := colTituP + ',oper';
    dadosTituP := dadosTituP + ',''' + StringGrid1.Cells[i,k] + '''';
    colBTitu := colBTitu + ',oper';
    dadosBTitu := dadosBTitu + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituP := colTituP + ',oper';
    dadosTituP := dadosTituP + ',''' + '101' + '''';
    colBTitu := colBTitu + ',oper';
    dadosBTitu := dadosBTitu + ',''' + '101' + '''';
  end;

  //C_FUNC (FUNCION�RIO)
  i:=BuscaColuna(StringGrid1,'c_func');
  if (i<>-1) then
  begin
    colTituP := colTituP + ',c_func';
    dadosTituP := dadosTituP + ',''' + StringGrid1.Cells[i,k] + '''';
    colBTitu := colBTitu + ',c_func';
    dadosBTitu := dadosBTitu + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colTituP := colTituP + ',c_func';
    dadosTituP := dadosTituP + ',''' + '1' + '''';
    colBTitu := colBTitu + ',c_func';
    dadosBTitu := dadosBTitu + ',''' + '1' + '''';
  end;


  for i := 0 to StringGrid1.ColCount-1 do
  begin
    //DATA (Data de cria��o do t�tulo)
    if (LowerCase(StringGrid1.Cells[i,0])='data') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if temp.Length >= 8 then
      begin
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
        colTituP := colTituP + ',data';
        dadosTituP := dadosTituP + ',''' + temp + '''';
      end
      else if temp.Length = 6 then begin
        temp2 := (Copy(DateToStr(Date()),9,2));
        //Testa os dois ultimos caracteres da data atual com data do titulo
        //Se os caracteres da data do titulo forem maiores, significa que � um s�culo antes
        if StrToInt(temp2)<StrToInt(Copy(temp,5,2)) then temp2 := IntToStr(StrToInt(temp2)-1);
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ temp2 + (Copy(temp,5,2));
        colTituP := colTituP + ',data';
        dadosTituP := dadosTituP + ',''' + temp + '''';
      end;
    end
    //VENC (Data de vencimento do t�tulo)
    else if (LowerCase(StringGrid1.Cells[i,0])='venc') then
    begin
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if temp.Length >= 8 then
      begin
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
        colTituP := colTituP + ',venc';
        dadosTituP := dadosTituP + ',''' + temp + '''';
      end
      else if temp.Length = 6 then begin
        temp2 := (Copy(DateToStr(Date()),9,2));
        //Testa os dois ultimos caracteres da data atual com data do titulo
        //Se os caracteres da data do titulo forem maiores, significa que � um s�culo antes
        if StrToInt(temp2)<StrToInt(Copy(temp,5,2)) then temp2 := IntToStr(StrToInt(temp2)-1);
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ temp2 + (Copy(temp,5,2));
        colTituP := colTituP + ',venc';
        dadosTituP := dadosTituP + ',''' + temp + '''';
      end;
    end
    //VALO (Valor do t�tulo)
    else if (LowerCase(StringGrid1.Cells[i,0])='valo') then
    begin
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0.0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      colTituP := colTituP + ',valo';
      dadosTituP := dadosTituP + ',' + temp;
      colBTitu := colBTitu + ',valo';
      dadosBTitu := dadosBTitu + ',' + temp;
      colBTitu := colBTitu + ',tota';
      dadosBTitu := dadosBTitu + ',' + temp;

      //Testar se n�o existe coluna saldo, se n�o existir joga o valor da VALO
      count:=BuscaColuna(StringGrid1,'sald');
      if (count=-1) then
      begin
        colTituP := colTituP + ',sald';
        if StringGrid1.Cells[i,k]='' then
        begin
          temp2 := '0.0';
        end
        else begin
          temp2 := StringGrid1.Cells[i,k];
        end;
        temp2 := corrigeFloat(temp2);
        dadosTituP := dadosTituP + ',' + temp2;
        temp2 := stringreplace(temp2, '.', ',',[rfReplaceAll, rfIgnoreCase]);
        saldo := StrToFloat(temp2);
      end;
    end
    //SALDO (Saldo do t�tulo)
    else if (LowerCase(StringGrid1.Cells[i,0])='sald') then
    begin
      colTituP := colTituP + ',sald';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0.0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      dadosTituP := dadosTituP + ',' + temp;
      temp := stringreplace(temp, '.', ',',[rfReplaceAll, rfIgnoreCase]);
      saldo := StrToFloat(temp);

      //Testar se n�o existe coluna valo, se n�o existir joga o valor da SALD
      count:=BuscaColuna(StringGrid1,'valo');
      if (count=-1) then
      begin
        if StringGrid1.Cells[i,k]='' then
        begin
          temp2 := '0.0';
        end
        else begin
          temp2 := StringGrid1.Cells[i,k];
        end;
        temp := corrigeFloat(temp);
        colTituP := colTituP + ',valo';
        dadosTituP := dadosTituP + ',' + temp2;
        colBTitu := colBTitu + ',valo';
        dadosBTitu := dadosBTitu + ',' + temp2;
        colBTitu := colBTitu + ',tota';
        dadosBTitu := dadosBTitu + ',' + temp2;
      end;
    end
    //HIST (HISTORICO)
    else if (LowerCase(StringGrid1.Cells[i,0])='hist') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,150));
      colTituP := colTituP + ',hist';
      dadosTituP := dadosTituP + ',''' + temp + '''';
      colBTitu := colBTitu + ',hist';
      dadosBTitu := dadosBTitu + ',''' + temp + '''';
    end

    //Colunas extras para BTITUP
    //DATA_BAIXA (DATA DA BAIXA)
    else if (LowerCase(StringGrid1.Cells[i,0])='data_baixa') then
    begin
      colBTitu := colBTitu + ',data';
      temp := Trim(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      if temp.Length >= 8 then
      begin
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
        dadosBTitu := dadosBTitu + ',''' + temp + '''';
      end
      else if temp.Length = 6 then begin
        temp2 := (Copy(DateToStr(Date()),9,2));
        //Testa os dois ultimos caracteres da data atual com data do titulo
        //Se os caracteres da data do titulo forem maiores, significa que � um s�culo antes
        if StrToInt(temp2)<StrToInt(Copy(temp,5,2)) then temp2 := IntToStr(StrToInt(temp2)-1);
        temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ temp2 + (Copy(temp,5,2));
        dadosBTitu := dadosBTitu + ',''' + temp + '''';
      end;
    end
    ;

  //Fim do For das colunas
  end;

  //ID (ID DA BTITUP)
  colBTitu := colBTitu + ',id';
  dadosBTitu := dadosBTitu + ',' + 'gen_id(gen_btitup_id,1)';
  //JURO (JUROS)
  colBTitu := colBTitu + ',juro';
  dadosBTitu := dadosBTitu + ',''' + '0.0' + '''';
  //MULT (MULTA)
  colBTitu := colBTitu + ',mult';
  dadosBTitu := dadosBTitu + ',''' + '0.0' + '''';
  //DESCO (DESCONTO)
  colBTitu := colBTitu + ',desco';
  dadosBTitu := dadosBTitu + ',''' + '0.0' + '''';
  //CONT (CONTA)
  colBTitu := colBTitu + ',cont';
  dadosBTitu := dadosBTitu + ',''' + '1' + '''';
  //EMPR_BAIX (EMPRESA ONDE O T�TULO FOI BAIXADO)
  colBTitu := colBTitu + ',empr_baix';
  dadosBTitu := dadosBTitu + ',''' + '1' + '''';
end;

procedure cImportaTituP.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
begin
  //----------------------------------------
  //Gravar no banco de T�tulos a Pagar
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL := TSQLDataSet.Create(Nil);
        SQL.SQLConnection := frmPrinc.conDestino;

        //Executar INSERT
        frmImportando.atualizaStatus('Inserindo dados na tabela TITUP.');
        SQL.CommandText := 'insert into titup ('+ colTituP +') values ' + '(' + dadosTituP + ');';
        SQL.ExecSQL;

        if saldo <= 0.0 then begin
          //Inserir na BTITUP
          frmImportando.atualizaStatus('Inserindo dados na tabela BTITUP.');
          SQL.CommandText := 'insert into btitup ('+ colBTitu +') values ' + '(' + dadosBTitu +');';
          SQL.ExecSQL;
        end;

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

        frmImportando.atualizaStatus('Comandos da TITUP.');
        WriteLn(fileTXT, '----------Comandos da TITUP----------');

        WriteLn(fileTXT, 'insert into TITUP ('+ colTituP +') values ' + '(' + dadosTituP + ');');
        WriteLn(fileTXT, 'COMMIT WORK;');

        if saldo <= 0.0 then begin
          //Inserir na BTITUP
          frmImportando.atualizaStatus('Comandos BTITUP.');
          WriteLn(fileTXT, '----------Comandos da BTITUP----------');
          WriteLn(fileTXT, 'insert into btitup ('+ colBTitu +') values ' + '(' + dadosBTitu +');');
          WriteLn(fileTXT, 'COMMIT WORK;');
        end;
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
