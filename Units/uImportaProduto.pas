unit uImportaProduto;

interface

uses
  System.SysUtils, Vcl.Grids, Vcl.Dialogs, Data.SqlExpr,
  uUtil;

type
  cImportaProduto = class
    private
      colProd, dadosProd, condUpdateProd, dadosUpdateProd: String;
      colProdTrib, dadosProdTrib, condUpdateProdTrib, dadosUpdateProdTrib, colRegistroProdTrib, dadosRegistroProdTrib: String;
      colProdAdic, dadosProdAdic, condUpdateProdAdic, dadosUpdateProdAdic, colRegistroProdAdic, dadosRegistroProdAdic: String;
      colProdCust, dadosProdCust, condUpdateProdCust, dadosUpdateProdCust, colRegistroProdCust, dadosRegistroProdCust: String;
      colItens, dadosItens, condUpdateItens, dadosUpdateItens: String;
      colMVA, dadosMVA, colRegistroMVA, dadosRegistroMVA: String;
      colProdForn, dadosProdForn: String;
      prodCod, prodEmpr: string;

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

constructor cImportaProduto.ImportaRegistro(numReg: Integer; Grid: TStringGrid);
begin
  k := numReg;
  StringGrid1 := Grid;

  CarregaColunas;
  Gravar;
end;

procedure cImportaProduto.CarregaColunas;
var
  i, count: Integer;
  temp, temp2: string;
  fileTXT: TextFile;
  but: Integer;
begin
  frmImportando.atualizaStatus('Produto '+IntToStr(k));

  colProd := '';
  dadosProd := '';
  colProdTrib := '';
  dadosProdTrib := '';
  colProdAdic := '';
  dadosProdAdic := '';
  colProdCust := '';
  dadosProdCust := '';
  colMVA := '';
  dadosMVA := '';
  colItens := '';
  dadosItens := '';
  colProdForn := '';
  dadosProdForn := '';

  colRegistroProdTrib := '';
  colRegistroProdAdic := '';
  colRegistroProdCust := '';
  colRegistroMVA := '';
  dadosRegistroProdTrib := '';
  dadosRegistroProdAdic := '';
  dadosRegistroProdCust := '';
  dadosRegistroMVA := '';

  dadosUpdateProd := '';
  dadosUpdateProdTrib := '';
  dadosUpdateProdAdic := '';
  dadosUpdateProdCust := '';
  dadosUpdateItens := '';
  condUpdateProd := '';
  condUpdateProdTrib := '';
  condUpdateProdAdic := '';
  condUpdateProdCust := '';

  //Será a condição da PROD_ESTO
  condUpdateItens := '';

  prodCod := '';
  prodEmpr := '';

  //Carregar informações para importar
  //-------------------------------------------------------

  //Empresa é obrigatório, se não tiver preenche com 1
  //EMPR (EMPRESA)
  i:=BuscaColuna(StringGrid1,'empr');
  if (i<>-1) then
  begin
    prodEmpr := StringGrid1.Cells[i,k];
    colProd := colProd + 'empr';
    dadosProd := dadosProd + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colProdTrib := colProdTrib + 'trib_prod_empr';
    dadosProdTrib := dadosProdTrib + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colProdAdic := colProdAdic + 'adic_prod_empr';
    dadosProdAdic := dadosProdAdic + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colProdCust := colProdCust + 'cust_prod_empr';
    dadosProdCust := dadosProdCust + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colProdTrib := colProdTrib + ',trib_empr';
    dadosProdTrib := dadosProdTrib + ',''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colProdAdic := colProdAdic + ',adic_empr';
    dadosProdAdic := dadosProdAdic + ',''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colProdCust := colProdCust + ',cust_empr';
    dadosProdCust := dadosProdCust + ',''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colMVA := colMVA + 'empr';
    dadosMVA := dadosMVA + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colMVA := colMVA + ',mva_empr';
    dadosMVA := dadosMVA + ',''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colProdForn := colProdForn + 'empr';
    dadosProdForn := dadosProdForn + '''' + prodEmpr + '''';
    colItens := colItens + 'empr';
    dadosItens := dadosItens + '''' + prodEmpr + '''';

    //Registros para outras empresas
    colRegistroProdTrib := colRegistroProdTrib + 'trib_prod_empr';
    dadosRegistroProdTrib := dadosRegistroProdTrib + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colRegistroProdAdic := colRegistroProdAdic + 'adic_prod_empr';
    dadosRegistroProdAdic := dadosRegistroProdAdic + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colRegistroProdCust := colRegistroProdCust + 'cust_prod_empr';
    dadosRegistroProdCust := dadosRegistroProdCust + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    colRegistroMVA := colRegistroMVA + 'empr';
    dadosRegistroMVA := dadosRegistroMVA + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
//            colRegistroMVA := colRegistroMVA + ',mva_empr';
//            dadosRegistroMVA := dadosRegistroMVA + ',''' + UpperCase(RemoveAcento(prodEmpr)) + '''';

    //Testa se é Update
    if VerificaUpdate('empr') = 1 then begin
      if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
      condUpdateProd := condUpdateProd + 'empr=' + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
      if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
      condUpdateProdTrib := condUpdateProdTrib + 'trib_empr=' + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
      if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
      condUpdateProdAdic := condUpdateProdAdic + 'adic_empr=' + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
      if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
      condUpdateProdCust := condUpdateProdCust + 'cust_empr=' + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
      if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
      condUpdateItens := condUpdateItens + 'cod_empr=' + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    end
    else begin
      if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
      dadosUpdateProd := dadosUpdateProd + 'empr=' + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
      if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
      dadosUpdateProdTrib := dadosUpdateProdTrib + 'trib_empr=' + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
      if dadosUpdateProdAdic <> '' then dadosUpdateProdAdic := dadosUpdateProdAdic + ', ';
      dadosUpdateProdAdic := dadosUpdateProdAdic + 'adic_empr' + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
      if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
      dadosUpdateProdCust := dadosUpdateProdCust + 'cust_empr=' + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
      if dadosUpdateItens <> '' then dadosUpdateItens := dadosUpdateItens + ', ';
      dadosUpdateItens := dadosUpdateItens + 'empr=' + '''' + UpperCase(RemoveAcento(prodEmpr)) + '''';
    end;
  end
  else begin
    prodEmpr := '1';
    colProd := colProd + 'empr';
    dadosProd := dadosProd + '''' + '1' + '''';
    colProdTrib := colProdTrib + 'trib_prod_empr';
    dadosProdTrib := dadosProdTrib + '''' + '1' + '''';
    colProdAdic := colProdAdic + 'adic_prod_empr';
    dadosProdAdic := dadosProdAdic + '''' + '1' + '''';
    colProdCust := colProdCust + 'cust_prod_empr';
    dadosProdCust := dadosProdCust + '''' + '1' + '''';
    colProdTrib := colProdTrib + ',trib_empr';
    dadosProdTrib := dadosProdTrib + ',''' + '1' + '''';
    colProdAdic := colProdAdic + ',adic_empr';
    dadosProdAdic := dadosProdAdic + ',''' + '1' + '''';
    colProdCust := colProdCust + ',cust_empr';
    dadosProdCust := dadosProdCust + ',''' + '1' + '''';
    colMVA := colMVA + 'empr';
    dadosMVA := dadosMVA + '''' + '1' + '''';
    colMVA := colMVA + ',mva_empr';
    dadosMVA := dadosMVA + ',''' + '1' + '''';
    colProdForn := colProdForn + 'empr';
    dadosProdForn := dadosProdForn + '''' + '1' + '''';
    colItens := colItens + 'empr';
    dadosItens := dadosItens + '''' + '1' + '''';

    //Registros para outras empresas
    colRegistroProdTrib := colRegistroProdTrib + 'trib_prod_empr';
    dadosRegistroProdTrib := dadosRegistroProdTrib + '''' + '1' + '''';
    colRegistroProdAdic := colRegistroProdAdic + 'adic_prod_empr';
    dadosRegistroProdAdic := dadosRegistroProdAdic + '''' + '1' + '''';
    colRegistroProdCust := colRegistroProdCust + 'cust_prod_empr';
    dadosRegistroProdCust := dadosRegistroProdCust + '''' + '1' + '''';
    colRegistroMVA := colRegistroMVA + 'empr';
    dadosRegistroMVA := dadosRegistroMVA + '''' + '1' + '''';
//            colRegistroMVA := colRegistroMVA + ',mva_empr';
//            dadosRegistroMVA := dadosRegistroMVA + ',''' + '1' + '''';
  end;

  //Codigo é obrigatório, se não tiver preenche com o generator
  //CODI (CODIGO)
  i:=BuscaColuna(StringGrid1,'Codi');
  if (i<>-1) then
  begin
    prodCod := stringreplace(StringGrid1.Cells[i,k], '.', '',[rfReplaceAll, rfIgnoreCase]);
    StringGrid1.Cells[i,k] := prodCod;
    colProd := colProd + ',codi';
    dadosProd := dadosProd + ',''' + prodCod + '''';
    colProd := colProd + ',cod_mestre';
    dadosProd := dadosProd + ',''' + prodCod + '''';
    colProdTrib := colProdTrib + ',trib_id';
    dadosProdTrib := dadosProdTrib + ',' + 'gen_id(gen_prod_tributos_id,1)';
    colProdTrib := colProdTrib + ',trib_prod_codi';
    dadosProdTrib := dadosProdTrib + ',''' + prodCod + '''';
    colProdAdic := colProdAdic + ',adic_id';
    dadosProdAdic := dadosProdAdic + ',' + 'gen_id(gen_prod_adicionais_id,1)';
    colProdAdic := colProdAdic + ',adic_prod_codi';
    dadosProdAdic := dadosProdAdic + ',''' + prodCod + '''';
    colProdCust := colProdCust + ',cust_id';
    dadosProdCust := dadosProdCust + ',' + 'gen_id(gen_prod_custos_id,1)';
    colProdCust := colProdCust + ',cust_prod_codi';
    dadosProdCust := dadosProdCust + ',''' + prodCod + '''';
    colMVA := colMVA + ',id';
    dadosMVA := dadosMVA + ',' + 'gen_id(gen_mva_id,1)';
    colMVA := colMVA + ',codi_prod';
    dadosMVA := dadosMVA + ',''' + prodCod + '''';
    colItens := colItens + ',codi';
    dadosItens := dadosItens + ',gen_id(gen_itens_id,1)';
    colItens := colItens + ',prodcod';
    dadosItens := dadosItens + ',''' + prodCod + '''';
    colProdForn := colProdForn + ',prod';
    dadosProdForn := dadosProdForn + ',''' + prodCod + '''';
    colProdForn := colProdForn + ',id';
    dadosProdForn := dadosProdForn + ',' + 'gen_id(gen_prod_forn_id,1)';

    //Registros para outras empresas
    colRegistroProdTrib := colRegistroProdTrib + ',trib_id';
    dadosRegistroProdTrib := dadosRegistroProdTrib + ',' + 'gen_id(gen_prod_tributos_id,1)';
    colRegistroProdTrib := colRegistroProdTrib + ',trib_prod_codi';
    dadosRegistroProdTrib := dadosRegistroProdTrib + ',''' + prodCod + '''';
    colRegistroProdAdic := colRegistroProdAdic + ',adic_id';
    dadosRegistroProdAdic := dadosRegistroProdAdic + ',' + 'gen_id(gen_prod_adicionais_id,1)';
    colRegistroProdAdic := colRegistroProdAdic + ',adic_prod_codi';
    dadosRegistroProdAdic := dadosRegistroProdAdic + ',''' + prodCod + '''';
    colRegistroProdCust := colRegistroProdCust + ',cust_id';
    dadosRegistroProdCust := dadosRegistroProdCust + ',' + 'gen_id(gen_prod_custos_id,1)';
    colRegistroProdCust := colRegistroProdCust + ',cust_prod_codi';
    dadosRegistroProdCust := dadosRegistroProdCust + ',''' + prodCod + '''';
    colRegistroMVA := colRegistroMVA + ',id';
    dadosRegistroMVA := dadosRegistroMVA + ',' + 'gen_id(gen_mva_id,1)';
    colRegistroMVA := colRegistroMVA + ',codi_prod';
    dadosRegistroMVA := dadosRegistroMVA + ',''' + prodCod + '''';

    //Testa se é Update
    if VerificaUpdate('codi') = 1 then begin
      if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
      condUpdateProd := condUpdateProd + 'codi=' + '''' + UpperCase(RemoveAcento(prodCod)) + '''';
      if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
      condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi=' + '''' + UpperCase(RemoveAcento(prodCod)) + '''';
      if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
      condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi=' + '''' + UpperCase(RemoveAcento(prodCod)) + '''';
      if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
      condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi=' + '''' + UpperCase(RemoveAcento(prodCod)) + '''';
      if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
      condUpdateItens := condUpdateItens + 'cod_prod=' + '''' + UpperCase(RemoveAcento(prodCod)) + '''';
    end
    else begin
      if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
      dadosUpdateProd := dadosUpdateProd + 'codi=' + '''' + UpperCase(RemoveAcento(prodCod)) + '''';
      if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
      dadosUpdateProdTrib := dadosUpdateProdTrib + 'trib_prod_codi=' + '''' + UpperCase(RemoveAcento(prodCod)) + '''';
      if dadosUpdateProdAdic <> '' then dadosUpdateProdAdic := dadosUpdateProdAdic + ', ';
      dadosUpdateProdAdic := dadosUpdateProdAdic + 'adic_prod_codi' + '''' + UpperCase(RemoveAcento(prodCod)) + '''';
      if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
      dadosUpdateProdCust := dadosUpdateProdCust + 'cust_prod_codi=' + '''' + UpperCase(RemoveAcento(prodCod)) + '''';
      if dadosUpdateItens <> '' then dadosUpdateItens := dadosUpdateItens + ', ';
      dadosUpdateItens := dadosUpdateItens + 'prodcod=' + '''' + UpperCase(RemoveAcento(prodCod)) + '''';
    end;
  end
  else begin
    colProd := colProd + ',codi';
    dadosProd := dadosProd + ',' + 'gen_id(gen_prod_id,1)';
    colProd := colProd + ',cod_mestre';
    dadosProd := dadosProd + ',' + 'gen_id(gen_prod_id,0)';
    colProdTrib := colProdTrib + ',trib_prod_codi';
    dadosProdTrib := dadosProdTrib + ',' + 'gen_id(gen_prod_id,0)';
    colProdTrib := colProdTrib + ',trib_id';
    dadosProdTrib := dadosProdTrib + ',' + 'gen_id(gen_prod_tributos_id,1)';
    colProdAdic := colProdAdic + ',adic_prod_codi';
    dadosProdAdic := dadosProdAdic + ',' + 'gen_id(gen_prod_id,0)';
    colProdAdic := colProdAdic + ',adic_id';
    dadosProdAdic := dadosProdAdic + ',' + 'gen_id(gen_prod_adicionais_id,1)';
    colProdCust := colProdCust + ',cust_prod_codi';
    dadosProdCust := dadosProdCust + ',' + 'gen_id(gen_prod_id,0)';
    colProdCust := colProdCust + ',cust_id';
    dadosProdCust := dadosProdCust + ',' + 'gen_id(gen_prod_custos_id,1)';
    colMVA := colMVA + ',codi_prod';
    dadosMVA := dadosMVA + ',' + 'gen_id(gen_prod_id,0)';
    colMVA := colMVA + ',id';
    dadosMVA := dadosMVA + ',' + 'gen_id(gen_mva_id,1)';
    colItens := colItens + ',codi';
    dadosItens := dadosItens + ',gen_id(gen_itens_id,1)';
    colItens := colItens + ',prodcod';
    dadosItens := dadosItens + ',' + 'gen_id(gen_prod_id,0)';
    colProdForn := colProdForn + ',prod';
    dadosProdForn := dadosProdForn + ',gen_id(gen_prod_id,0)';
    colProdForn := colProdForn + ',id';
    dadosProdForn := dadosProdForn + ',' + 'gen_id(gen_prod_forn_id,1)';

    //Registros para outras empresas
    colRegistroProdTrib := colRegistroProdTrib + ',trib_id';
    dadosRegistroProdTrib := dadosRegistroProdTrib + ',' + 'gen_id(gen_prod_tributos_id,1)';
    colRegistroProdTrib := colRegistroProdTrib + ',trib_prod_codi';
    dadosRegistroProdTrib := dadosRegistroProdTrib + ',' + 'gen_id(gen_prod_id,0)';
    colRegistroProdAdic := colRegistroProdAdic + ',adic_id';
    dadosRegistroProdAdic := dadosRegistroProdAdic + ',' + 'gen_id(gen_prod_adicionais_id,1)';
    colRegistroProdAdic := colRegistroProdAdic + ',adic_prod_codi';
    dadosRegistroProdAdic := dadosRegistroProdAdic + ',' + 'gen_id(gen_prod_id,0)';
    colRegistroProdCust := colRegistroProdCust + ',cust_id';
    dadosRegistroProdCust := dadosRegistroProdCust + ',' + 'gen_id(gen_prod_custos_id,1)';
    colRegistroProdCust := colRegistroProdCust + ',cust_prod_codi';
    dadosRegistroProdCust := dadosRegistroProdCust + ',' + 'gen_id(gen_prod_id,0)';
    colRegistroMVA := colRegistroMVA + ',id';
    dadosRegistroMVA := dadosRegistroMVA + ',' + 'gen_id(gen_mva_id,1)';
    colRegistroMVA := colRegistroMVA + ',codi_prod';
    dadosRegistroMVA := dadosRegistroMVA + ',' + 'gen_id(gen_prod_id,0)';

    prodCod := frmPrinc.getProdCodUpdate(k);
    if (prodCod = '') and (butContinue = 0) then begin
      but := Mensagem('Recomenda-se utilizar ao menos uma das seguintes colunas: CODI, REFE, REFE_ORIGINAL, CODI_BARRA, CODI_BARRA_COM',mtCustom,[mbYes, mbNo],['Continuar','Parar'],'Precisa mais informações');
      if (but = 6) then begin
        butContinue := 1;
      end
      else if (but = 7) then begin
        status := 0;
        Exit;
      end;
    end;
  end;

  //Grupo, subgrupo, departamento, marca e tipo são obrigatórios, se não tiver colocar padroes
  //GRUP
  ///-----------------------------------------------------------------}
  i:=BuscaColuna(StringGrid1,'grup');
  if (i<>-1) then
  begin
    temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
    temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
    temp := (Copy(temp,1,30));
    //Se for letras, buscar código.
    if not (IsNumeric(temp)) then
    begin
      temp2 := querySelect('select g.codi from grup_prod g where g.descr = '''+temp+'''');
      //Se não encontrar a string, cadastrar grupo
      if temp2='' then begin
        if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT' then begin
            //Carregar arquivo TXT
            AssignFile(fileTXT, frmPrinc.DBPath.Text);
            if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
            else append(fileTXT);
            WriteLn(fileTXT, 'insert into grup_prod (CODI,DESCR,EMPR) values (case when (select g.codi from grup_prod g where g.descr = '''+temp+''') is null then gen_id(gen_grup_prod_id,1 ) else (select g.codi from grup_prod g where g.descr = '''+temp+''') end,'''+temp+''',1);');
            WriteLn(fileTXT, 'COMMIT WORK;');
            CloseFile(fileTXT);
          end
        else begin
          queryInsert('insert into grup_prod (CODI,DESCR,EMPR) values (gen_id(gen_grup_prod_id,1),'''+temp+''',1);');
        end;
        colProd := colProd + ',grup';
        dadosProd := dadosProd + ',' + 'gen_id(gen_grup_prod_id,0)';
      end
      else begin
        colProd := colProd + ',grup';
        dadosProd := dadosProd + ',''' + temp2 + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('grup') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'grup=(select codi from grup_prod where descr='+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where grup = (select codi from grup_prod where descr='+''''+temp+'''))';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where grup = (select codi from grup_prod where descr='+''''+temp+'''))';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where grup = (select codi from grup_prod where descr='+''''+temp+'''))';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where grup = (select codi from grup_prod where descr='+''''+temp+'''))';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'grup=' + '''' + temp + '''';
      end;
    end
    else begin
      //Se for números, considera como código
      //Antes buscamos se existe o código cadastrado, se não encontrar colocamos o generator mesmo
      temp2 := querySelect('select g.codi from grup_prod g where g.codi = '''+temp+'''');
      //Se não encontrar o codigo, colocamos o generator
      if temp2='' then begin
        colProd := colProd + ',grup';
        dadosProd := dadosProd + ',' + 'gen_id(gen_grup_prod_id,0)';
      end
      //Se encontrar usa o código
      else begin
        colProd := colProd + ',grup';
        dadosProd := dadosProd + ',''' + temp2 + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('grup') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'grup='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where grup = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where grup = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where grup = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where grup = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'grup=' + '''' + temp + '''';
      end;
    end;
  end
  else begin
    colProd := colProd + ',grup';
    dadosProd := dadosProd + ',''' + '1' + '''';
  end;
  //SUB_GRUP
  i:=BuscaColuna(StringGrid1,'sub_grup');
  if (i<>-1) then
  begin
    temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
    temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
    temp := (Copy(temp,1,30));
    //Se for letras, buscar código.
    if not (IsNumeric(temp)) then
    begin
      temp2 := querySelect('select g.codi from sub_grup_prod g where g.descr = '''+temp+'''');
      //Se não encontrar a string, cadastrar sub grupo
      if temp2='' then begin
        if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT' then begin
            //Carregar arquivo TXT
            AssignFile(fileTXT, frmPrinc.DBPath.Text);
            if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
            else append(fileTXT);
            WriteLn(fileTXT, 'insert into sub_grup_prod (CODI,DESCR,EMPR) values (case when (select g.codi from sub_grup_prod g where g.descr = '''+temp+''') is null then gen_id(gen_sub_grup_prod_id,1 ) else (select g.codi from sub_grup_prod g where g.descr = '''+temp+''') end,'''+temp+''',1);');
            WriteLn(fileTXT, 'COMMIT WORK;');
            CloseFile(fileTXT);
          end
        else begin
          queryInsert('insert into sub_grup_prod (CODI,DESCR,EMPR) values (gen_id(gen_sub_grup_prod_id,1),'''+temp+''',1);');
        end;
        colProd := colProd + ',sub_grup';
        dadosProd := dadosProd + ',' + 'gen_id(gen_sub_grup_prod_id,0)';
      end
      else begin
        colProd := colProd + ',sub_grup';
        dadosProd := dadosProd + ',''' + temp2 + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('sub_grup') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'sub_grup=(select codi from sub_grup_prod where descr='+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where sub_grup = (select codi from sub_grup_prod where descr='+''''+temp+'''))';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where sub_grup = (select codi from sub_grup_prod where descr='+''''+temp+'''))';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where sub_grup = (select codi from sub_grup_prod where descr='+''''+temp+'''))';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where sub_grup = (select codi from sub_grup_prod where descr='+''''+temp+'''))';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'sub_grup=' + '''' + temp + '''';
      end;
    end
    else begin
      //Se for números, considera como código
      //Antes buscamos se existe o código cadastrado, se não encontrar colocamos o generator mesmo
      temp2 := querySelect('select g.codi from sub_grup_prod g where g.codi = '''+temp+'''');
      //Se não encontrar o codigo, colocamos o generator
      if temp2='' then begin
        colProd := colProd + ',sub_grup';
        dadosProd := dadosProd + ',' + 'gen_id(gen_sub_grup_prod_id,0)';
      end
      //Se encontrar usa o código
      else begin
        colProd := colProd + ',sub_grup';
        dadosProd := dadosProd + ',''' + temp2 + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('sub_grup') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'sub_grup='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where sub_grup = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where sub_grup = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where sub_grup = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where sub_grup = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'sub_grup=' + '''' + temp + '''';
      end;
    end;
  end
  else begin
    colProd := colProd + ',sub_grup';
    dadosProd := dadosProd + ',''' + '1' + '''';
  end;
  //DEPARTAMENTO
  i:=BuscaColuna(StringGrid1,'departamento');
  if (i<>-1) then
  begin
    temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
    temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
    temp := (Copy(temp,1,30));
    //Se for letras, buscar código.
    if not (IsNumeric(temp)) then
    begin
      temp2 := querySelect('select g.codi from departamento g where g.descr = '''+temp+'''');
      //Se não encontrar a string, cadastrar departamento
      if temp2='' then begin
        if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT' then begin
            //Carregar arquivo TXT
            AssignFile(fileTXT, frmPrinc.DBPath.Text);
            if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
            else append(fileTXT);
            WriteLn(fileTXT, 'insert into departamento (CODI,DESCR) values (case when (select g.codi from departamento g where g.descr = '''+temp+''') is null then (select max(g.codi) from departamento g) else (select g.codi from departamento g where g.descr = '''+temp+''') end,'''+temp+''');');
            WriteLn(fileTXT, 'COMMIT WORK;');
            CloseFile(fileTXT);
          end
        else begin
          temp2 := querySelect('select max(g.codi) from departamento g');
          temp2 := IntToStr(StrToInt(temp2)+1);
          queryInsert('insert into departamento (CODI,DESCR) values ('+temp2+','''+temp+''');');
        end;
        colProd := colProd + ',codi_departamento';
        dadosProd := dadosProd + ',' + temp2;
      end
      else begin
        colProd := colProd + ',codi_departamento';
        dadosProd := dadosProd + ',''' + temp2 + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('departamento') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi_departamento=(select codi from departamento where descr='+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where codi_departamento = (select codi from departamento where descr='+''''+temp+'''))';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where codi_departamento = (select codi from departamento where descr='+''''+temp+'''))';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where codi_departamento = (select codi from departamento where descr='+''''+temp+'''))';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where codi_departamento = (select codi from departamento where descr='+''''+temp+'''))';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'codi_departamento=' + '''' + temp + '''';
      end;
    end
    else begin
      //Se for números, considera como código
      //Antes buscamos se existe o código cadastrado, se não encontrar colocamos o generator mesmo
      temp2 := querySelect('select g.codi from departamento g where g.codi = '''+temp+'''');
      //Se não encontrar o codigo, colocamos o generator
      if temp2='' then begin
        temp2 := querySelect('select max(g.codi) from departamento g');
        colProd := colProd + ',codi_departamento';
        dadosProd := dadosProd + ',' + temp2;
      end
      //Se encontrar usa o código
      else begin
        colProd := colProd + ',codi_departamento';
        dadosProd := dadosProd + ',''' + temp2 + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('departamento') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi_departamento='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where departamento = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where departamento = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where departamento = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where departamento = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'codi_departamento=' + '''' + temp + '''';
      end;
    end;
  end
  else begin
    colProd := colProd + ',codi_departamento';
    dadosProd := dadosProd + ',''' + '0' + '''';
  end;
  //MARCA
  i:=BuscaColuna(StringGrid1,'marca');
  if (i<>-1) then
  begin
    temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
    temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
    temp := (Copy(temp,1,50));
    //Se for letras, buscar código.
    if not (IsNumeric(temp)) then
    begin
      temp2 := querySelect('select g.codi from marca g where g.descr = '''+temp+'''');
      //Se não encontrar a string, cadastrar marca
      if temp2='' then begin
        if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT' then begin
            //Carregar arquivo TXT
            AssignFile(fileTXT, frmPrinc.DBPath.Text);
            if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
            else append(fileTXT);
            WriteLn(fileTXT, 'insert into marca (CODI,DESCR) values (case when (select g.codi from marca g where g.descr = '''+temp+''') is null then gen_id(gen_marca_id,1) else (select g.codi from marca g where g.descr = '''+temp+''') end,'''+temp+''');');
            WriteLn(fileTXT, 'COMMIT WORK;');
            CloseFile(fileTXT);
          end
        else begin
          queryInsert('insert into marca (CODI,DESCR) values (gen_id(gen_marca_id,1),'''+temp+''');');
        end;
        colProd := colProd + ',marca';
        dadosProd := dadosProd + ',' + 'gen_id(gen_marca_id,0)';
      end
      else begin
        colProd := colProd + ',marca';
        dadosProd := dadosProd + ',''' + temp2 + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('marca') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'marca=(select codi from marca where descr='+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where marca = (select codi from marca where descr='+''''+temp+'''))';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where marca = (select codi from marca where descr='+''''+temp+'''))';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where marca = (select codi from marca where descr='+''''+temp+'''))';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where marca = (select codi from marca where descr='+''''+temp+'''))';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'marca=' + '''' + temp + '''';
      end;
    end
    else begin
      //Se for números, considera como código
      //Antes buscamos se existe o código cadastrado, se não encontrar colocamos o generator mesmo
      temp2 := querySelect('select g.codi from marca g where g.codi = '''+temp+'''');
      //Se não encontrar o codigo, colocamos o generator
      if temp2='' then begin
        colProd := colProd + ',marca';
        dadosProd := dadosProd + ',' + 'gen_id(gen_marca_id,0)';
      end
      //Se encontrar usa o código
      else begin
        colProd := colProd + ',marca';
        dadosProd := dadosProd + ',''' + temp2 + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('marca') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'marca='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where marca = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where marca = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where marca = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where marca = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'marca=' + '''' + temp + '''';
      end;
    end;
  end;
  //TIPO
  i:=BuscaColuna(StringGrid1,'tipo');
  if (i<>-1) then
  begin
    colProd := colProd + ',codi_tipo';
    if StringGrid1.Cells[i,k] <> '' then begin
      dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
      //Testa se é Update
      if VerificaUpdate('tipo') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi_tipo='+''''+UpperCase(RemoveAcento(StringGrid1.Cells[i,k]))+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where codi_tipo = '+''''+UpperCase(RemoveAcento(StringGrid1.Cells[i,k]))+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where codi_tipo = '+''''+UpperCase(RemoveAcento(StringGrid1.Cells[i,k]))+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where codi_tipo = '+''''+UpperCase(RemoveAcento(StringGrid1.Cells[i,k]))+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where codi_tipo = '+''''+UpperCase(RemoveAcento(StringGrid1.Cells[i,k]))+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'codi_tipo=' + '''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
      end;
    end else begin
      dadosProd := dadosProd + ',''' + '0' + '''';
    end;
  end
  else begin
    colProd := colProd + ',codi_tipo';
    dadosProd := dadosProd + ',''' + '0' + '''';
  end;

  //PS (Produto ou serviço) Padrão deixar 'P' pois sempre importamos produtos
  colProd := colProd + ',ps';
  dadosProd := dadosProd + ',''' + 'P' + '''';

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
    dadosItens := dadosItens + ',' + temp;
    colItens := colItens + ',qtd';
    //Testa se é Update
    if VerificaUpdate('qtd') = 1 then begin
      if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
      condUpdateProd := condUpdateProd + 'codi in (select cod_prod from prod_esto where qtd = '+temp+')';
      if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
      condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cod_prod from prod_esto where qtd = '+temp+')';
      if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
      condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cod_prod from prod_esto where qtd = '+temp+')';
      if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
      condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select cod_prod from prod_esto where qtd = '+temp+')';
      if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
      condUpdateItens := condUpdateItens + 'cod_prod in (select cod_prod from prod_esto where qtd = '+temp+')';
    end
    else begin
      {
       Comando para inserir na ITENS quando a diferenca de estoque for diferente de 0
      }
      //Comando Insert na itens - Ainda nao tem condição no WHERE, pois irá usar condUpdateItens concatenado
      dadosUpdateItens := dadosUpdateItens + ' insert into ITENS (CODI,PRODCOD,NUME,TIPO,EPV,QTD,EMPR) '+
          '    select '+
          '        gen_id(gen_itens_id,1) CODI, '+
          '        '+prodCod+' PRODCOD, '+
          '        gen_id(gen_prod_ajus_id,0)+1 NUME, '+
          '        case '+
          '            when ('+temp+'-pe.qtd) > 0 then 6 '+
          '            when ('+temp+'-pe.qtd) < 0 then 3 '+
          '        end TIPO, '+
          '        ''A'' EPV, '+
          '        ABS('+temp+'-pe.qtd) QTD, '+
          '        '+prodEmpr+' EMPR '+
          '    from prod_esto pe '+
          '    where ('+temp+'-pe.qtd) <> 0 ';
    end;
    //Setar tipo do item
    colItens := colItens + ',tipo';
    dadosItens := dadosItens + ',''' + '6' + '''';
  end;

  //EST
  i:=BuscaColuna(StringGrid1,'est');
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
    dadosItens := dadosItens + ',' + temp;
    colItens := colItens + ',qtd';
    //Testa se é Update
    if VerificaUpdate('est') = 1 then begin
      if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
      condUpdateProd := condUpdateProd + 'codi in (select cod_prod from prod_esto where qtd = '+temp+')';
      if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
      condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cod_prod from prod_esto where qtd = '+temp+')';
      if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
      condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cod_prod from prod_esto where qtd = '+temp+')';
      if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
      condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select cod_prod from prod_esto where qtd = '+temp+')';
      if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
      condUpdateItens := condUpdateItens + 'cod_prod in (select cod_prod from prod_esto where qtd = '+temp+')';
    end
    else begin
      {
       Comando para inserir na ITENS quando a diferenca de estoque for diferente de 0
      }
      //Comando Insert na itens - Ainda nao tem condição no WHERE, pois irá usar condUpdateItens concatenado
      dadosUpdateItens := dadosUpdateItens + ' insert into ITENS (CODI,PRODCOD,NUME,TIPO,EPV,QTD,EMPR) '+
          '    select '+
          '        gen_id(gen_itens_id,1) CODI, '+
          '        '+prodCod+' PRODCOD, '+
          '        gen_id(gen_prod_ajus_id,0)+1 NUME, '+
          '        case '+
          '            when ('+temp+'-pe.qtd) > 0 then 5 '+
          '            when ('+temp+'-pe.qtd) < 0 then 2 '+
          '        end TIPO, '+
          '        ''A'' EPV, '+
          '        ABS('+temp+'-pe.qtd) QTD, '+
          '        '+prodEmpr+' EMPR '+
          '    from prod_esto pe '+
          '    where ('+temp+'-pe.qtd) <> 0 ';
    end;
    //Setar tipo do item
    colItens := colItens + ',tipo';
    dadosItens := dadosItens + ',''' + '5' + '''';
  end;

  //MAX
  i:=BuscaColuna(StringGrid1,'max');
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
    dadosItens := dadosItens + ',' + temp;
    colItens := colItens + ',qtd';
    //Testa se é Update
    if VerificaUpdate('max') = 1 then begin
      if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
      condUpdateProd := condUpdateProd + 'codi in (select cod_prod from prod_esto where qtd_max = '+temp+')';
      if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
      condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cod_prod from prod_esto where qtd_max = '+temp+')';
      if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
      condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cod_prod from prod_esto where qtd_max = '+temp+')';
      if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
      condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select cod_prod from prod_esto where qtd_max = '+temp+')';
      if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
      condUpdateItens := condUpdateItens + 'cod_prod in (select cod_prod from prod_esto where qtd_max = '+temp+')';
    end
    else begin
      {
       Comando para inserir na ITENS quando a diferenca de estoque for diferente de 0
      }
      //Comando Insert na itens - Ainda nao tem condição no WHERE, pois irá usar condUpdateItens concatenado
      dadosUpdateItens := dadosUpdateItens + ' insert into ITENS (CODI,PRODCOD,NUME,TIPO,EPV,QTD,EMPR) '+
          '    select '+
          '        gen_id(gen_itens_id,1) CODI, '+
          '        '+prodCod+' PRODCOD, '+
          '        gen_id(gen_prod_ajus_id,0)+1 NUME, '+
          '        case '+
          '            when ('+temp+'-pe.qtd_max) > 0 then 4 '+
          '            when ('+temp+'-pe.qtd_max) < 0 then 1 '+
          '        end TIPO, '+
          '        ''A'' EPV, '+
          '        ABS('+temp+'-pe.qtd_max) QTD, '+
          '        '+prodEmpr+' EMPR '+
          '    from prod_esto pe '+
          '    where ('+temp+'-pe.qtd_max) <> 0 ';
    end;
    //Setar tipo do item
    colItens := colItens + ',tipo';
    dadosItens := dadosItens + ',''' + '4' + '''';
  end;

  //Se não tiver informação de estoque, coloca padrões
  if (BuscaColuna(StringGrid1,'qtd')=-1) and (BuscaColuna(StringGrid1,'est')=-1) and (BuscaColuna(StringGrid1,'max')=-1) then begin
    colItens := colItens + ',qtd';
    dadosItens := dadosItens + ',''' + '0' + '''';
    //Setar tipo do item
    colItens := colItens + ',tipo';
    dadosItens := dadosItens + ',''' + '6' + '''';
  end;
  //Campos adicionais para a itens
  colItens := colItens + ',epv';
  dadosItens := dadosItens + ',''' + 'A' + '''';
  colItens := colItens + ',nume';
  dadosItens := dadosItens + ',' + 'gen_id(gen_prod_ajus_id,0)';

  //Fornecedor é obrigatório, se não tiver põe 1
  //FORN
  i:=BuscaColuna(StringGrid1,'forn');
  if (i<>-1) then
  begin
    colProdForn := colProdForn + ',forn';
    dadosProdForn := dadosProdForn + ',''' + StringGrid1.Cells[i,k] + '''';
  end
  else begin
    colProdForn := colProdForn + ',forn';
    dadosProdForn := dadosProdForn + ',''' + '1' + '''';
  end;



  for i := 0 to StringGrid1.ColCount-1 do
  begin
    //COLEÇÃO
    ///-----------------------------------------------------------------}
    if (LowerCase(StringGrid1.Cells[i,0])='colecao') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,30));
      if temp <> '' then
      begin
        temp2 := querySelect('select pc.codi from prod_colecao pc where pc.descri= '''+temp+'''');
        //Se não encontrar a string, cadastrar grupo
        if temp2='' then begin
          if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.TXT' then begin
              //Carregar arquivo TXT
              AssignFile(fileTXT, frmPrinc.DBPath.Text);
              if not FileExists(frmPrinc.DBPath.Text) then ReWrite(fileTXT)
              else append(fileTXT);
              WriteLn(fileTXT, 'insert into prod_colecao (CODI,DESCRI) values (case when (select pc.codi from prod_colecao pc where pc.descri= '''+temp+''') is null then gen_id(gen_prod_colecao_id,1 ) else (select pc.codi from prod_colecao pc where pc.descri= '''+temp+''') end,'''+temp+''');');
              WriteLn(fileTXT, 'COMMIT WORK;');
              CloseFile(fileTXT);
            end
          else begin
            queryInsert('insert into prod_colecao (CODI,DESCRI) values (gen_id(gen_prod_colecao_id,1),'''+temp+''');');
          end;
          colProd := colProd + ',codi_colecao';
          dadosProd := dadosProd + ',' + 'gen_id(gen_prod_colecao_id,0)';
        end
        else begin
          colProd := colProd + ',codi_colecao';
          dadosProd := dadosProd + ',''' + temp2 + '''';
        end;
        //Testa se é Update
        if VerificaUpdate('colecao') = 1 then begin
          if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
          condUpdateProd := condUpdateProd + 'codi_colecao=(select codi from prod_colecao where DESCRI='+''''+temp+''')';
          if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
          condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where codi_colecao = (select codi from prod_colecao where DESCRI='+''''+temp+'''))';
          if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
          condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where codi_colecao = (select codi from prod_colecao where DESCRI='+''''+temp+'''))';
          if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
          condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where codi_colecao = (select codi from prod_colecao where DESCRI='+''''+temp+'''))';
          if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
          condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where codi_colecao = (select codi from prod_colecao where DESCRI='+''''+temp+'''))';
        end
        else begin
          if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
          dadosUpdateProd := dadosUpdateProd + 'codi_colecao=' + '''' + temp + '''';
        end;
      end;
    end
    //DESCR
    else if (LowerCase(StringGrid1.Cells[i,0])='descr') then
    begin
      colProd := colProd + ',descr';
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', QuotedStr(''''),[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,120));
      dadosProd := dadosProd + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('descr') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'descr='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where descr = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where descr = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where descr = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where descr = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'descr=' + '''' + temp + '''';
      end;
    end
    //DESCR2 (DESCRIÇÃO COMPLEMENTAR)
    else if (LowerCase(StringGrid1.Cells[i,0])='descr2') then
    begin
      colProd := colProd + ',descr2';
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', QuotedStr(''''),[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,255));
      dadosProd := dadosProd + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('descr2') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'descr2='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where descr2 = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where descr2 = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where descr2 = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where descr2 = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'descr2=' + '''' + temp + '''';
      end;
    end
    //REFE (Referencia)
    else if (LowerCase(StringGrid1.Cells[i,0])='refe') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);
      colProd := colProd + ',refe';
      dadosProd := dadosProd + ',''' + temp + '''';

      //Testa se é Update
      if VerificaUpdate('refe') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'refe='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where refe = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where refe = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where refe = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where refe = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'refe=' + '''' + temp + '''';
      end;
    end
    //REFE_ORIGINAL (Referencia Original)
    else if (LowerCase(StringGrid1.Cells[i,0])='refe_original') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);
      colProd := colProd + ',refe_original';
      dadosProd := dadosProd + ',''' + temp + '''';

      //Testa se é Update
      if VerificaUpdate('refe_original') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'refe_original='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where refe_original = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where refe_original = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where refe_original = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where refe_original = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'refe_original=' + '''' + temp + '''';
      end;
    end
    //LOCALIZACAO (Localização)
    else if (LowerCase(StringGrid1.Cells[i,0])='localizacao') then
    begin
      temp := UpperCase(StringGrid1.Cells[i,k]);
      temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);
      colProdAdic := colProdAdic + ',adic_localizacao';
      dadosProdAdic := dadosProdAdic + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('localizacao') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'adic_localizacao='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where adic_localizacao = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where adic_localizacao = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where adic_localizacao = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where adic_localizacao = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'adic_localizacao=' + '''' + temp + '''';
      end;
    end
    //CODI_BARRA (Codigo de barras unitario)
    else if (LowerCase(StringGrid1.Cells[i,0])='codi_barra') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, ',', ' ',[rfReplaceAll, rfIgnoreCase]);
      colProd := colProd + ',codi_barra';
      dadosProd := dadosProd + ',''' + temp + '''';

      //Testa se é Update
      if VerificaUpdate('codi_barra') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi_barra='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where codi_barra = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where codi_barra = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where codi_barra = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where codi_barra = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'codi_barra=' + '''' + temp + '''';
      end;
    end
    //CODI_BARRA_COM (Codigo de barras embalagem)
    else if (LowerCase(StringGrid1.Cells[i,0])='codi_barra_com') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, ',', ' ',[rfReplaceAll, rfIgnoreCase]);
      colProd := colProd + ',codi_barra_com';
      dadosProd := dadosProd + ',''' + temp + '''';

      //Testa se é Update
      if VerificaUpdate('codi_barra_com') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi_barra_com='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where codi_barra_com = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where codi_barra_com = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where codi_barra_com = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where codi_barra_com = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'codi_barra_com=' + '''' + temp + '''';
      end;
    end
    //OBS (Observação)
    else if (LowerCase(StringGrid1.Cells[i,0])='obs') then
    begin
      colProd := colProd + ',obs';
      temp := StringGrid1.Cells[i,k];
      dadosProd := dadosProd + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('obs') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'obs='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where obs = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where obs = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where obs = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where obs = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'obs=' + '''' + temp + '''';
      end;
    end
    //NCM
    else if (LowerCase(StringGrid1.Cells[i,0])='ncm') then
    begin
      colProd := colProd + ',ncm';
      temp := stringreplace(StringGrid1.Cells[i,k], '''', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, ',', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,10));
      dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(temp)) + '''';
      //Testa se é Update
      if VerificaUpdate('ncm') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'ncm='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where ncm = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where ncm = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where ncm = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where ncm = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'ncm=' + '''' + temp + '''';
      end;
    end
    //CEST
    else if (LowerCase(StringGrid1.Cells[i,0])='cest') then
    begin
      colProd := colProd + ',cest';
      temp := stringreplace(StringGrid1.Cells[i,k], '''', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, ',', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,8));
      dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(temp)) + '''';
      //Testa se é Update
      if VerificaUpdate('cest') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'cest='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where cest = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where cest = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where cest = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where cest = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'cest=' + '''' + temp + '''';
      end;
    end
    //UNID (Unidade de medida)
    else if (LowerCase(StringGrid1.Cells[i,0])='unid') then
    begin
      temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
      temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
      temp := (Copy(temp,1,12));
      colProd := colProd + ',unid';
      dadosProd := dadosProd + ',''' + temp + '''';
      //Testa se é Update
      if VerificaUpdate('unid') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'unid='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where unid = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where unid = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where unid = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where unid = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'unid=' + '''' + temp + '''';
      end;
    end
    //PESL (Peso Líquido)
    else if (LowerCase(StringGrid1.Cells[i,0])='pesl') then
    begin
      colProd := colProd + ',pesl';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      dadosProd := dadosProd + ',' + temp;
      //Testa se é Update
      if VerificaUpdate('pesl') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'pesl='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where pesl = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where pesl = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where pesl = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where pesl = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'pesl=' + '''' + temp + '''';
      end;
    end
    //PESB (Peso Bruto)
    else if (LowerCase(StringGrid1.Cells[i,0])='pesb') then
    begin
      colProd := colProd + ',pesb';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      dadosProd := dadosProd + ',' + temp;
      //Testa se é Update
      if VerificaUpdate('pesb') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'pesb='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where pesb = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where pesb = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where pesb = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where pesb = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'pesb=' + '''' + temp + '''';
      end;
    end
    //FATOR_CONV (Fator de Conversão)
    else if (LowerCase(StringGrid1.Cells[i,0])='fator_conv') then
    begin
      colProd := colProd + ',fator_conv';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      dadosProd := dadosProd + ',' + temp;
    end
    //QTD_MINIMA (Quantidade Mínima)
    else if (LowerCase(StringGrid1.Cells[i,0])='qtd_minima') then
    begin
      colProdAdic := colProdAdic + ',adic_qtd_minima';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      dadosProdAdic := dadosProdAdic + ',' + temp;
    end
    //CUSTO (Custo)
    else if (LowerCase(StringGrid1.Cells[i,0])='custo') then
    begin
      colProdCust := colProdCust + ',cust_custo';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := stringreplace(temp, 'R', '',[rfReplaceAll, rfIgnoreCase]);
      temp := stringreplace(temp, '$', '',[rfReplaceAll, rfIgnoreCase]);
      temp := Trim(temp);
      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;
      //Testar se existe o custo_real, se não joga o custo mesmo
      if (BuscaColuna(StringGrid1,'custo_real')=-1) then
      begin
        colProdCust := colProdCust + ',cust_custo_real';
        dadosProdCust := dadosProdCust + ',' + temp;
      end;
      //Testa se é Update
      if VerificaUpdate('custo') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_custo='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_custo = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_custo = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_custo = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_custo = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_custo=' + '''' + temp + '''';
      end;
    end
    //CUSTO_MEDIO (Custo Medio)
    else if (LowerCase(StringGrid1.Cells[i,0])='custo_medio') then
    begin
      colProdCust := colProdCust + ',cust_custo_medio';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;
      //Testa se é Update
      if VerificaUpdate('custo_medio') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_custo_medio='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_custo_medio = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_custo_medio = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_custo_medio = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_custo_medio = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_custo_medio=' + '''' + temp + '''';
      end;
    end
    //IPI (IPI agregado ao custo)
    else if (LowerCase(StringGrid1.Cells[i,0])='ipi') then
    begin
      colProdCust := colProdCust + ',cust_ipi';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := stringreplace(temp, '%', '',[rfReplaceAll, rfIgnoreCase]);
      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;
      //Testa se é Update
      if VerificaUpdate('ipi') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_ipi='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_ipi = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_ipi = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_ipi = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_ipi = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_ipi=' + '''' + temp + '''';
      end;
    end
    //PIS (PIS agregado ao custo)
    else if (LowerCase(StringGrid1.Cells[i,0])='pis') then
    begin
      colProdCust := colProdCust + ',cust_pis';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := stringreplace(temp, '%', '',[rfReplaceAll, rfIgnoreCase]);
      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;
      //Testa se é Update
      if VerificaUpdate('pis') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_pis='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_pis = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_pis = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_pis = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_pis = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_pis=' + '''' + temp + '''';
      end;
    end
    //COFINS (COFINS agregado ao custo)
    else if (LowerCase(StringGrid1.Cells[i,0])='cofins') then
    begin
      colProdCust := colProdCust + ',cust_cofins';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := stringreplace(temp, '%', '',[rfReplaceAll, rfIgnoreCase]);
      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;
      //Testa se é Update
      if VerificaUpdate('cofins') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_cofins='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_cofins = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_cofins = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_cofins = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_cofins = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_cofins=' + '''' + temp + '''';
      end;
    end
    //ICMS (ICMS agregado ao custo)
    else if (LowerCase(StringGrid1.Cells[i,0])='icms') then
    begin
      colProdCust := colProdCust + ',cust_icms';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := stringreplace(temp, '%', '',[rfReplaceAll, rfIgnoreCase]);
      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;
      //Testa se é Update
      if VerificaUpdate('icms') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_icms='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_icms = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_icms = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_icms = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_icms = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_icms=' + '''' + temp + '''';
      end;
    end
    //FRETE
    else if (LowerCase(StringGrid1.Cells[i,0])='frete') then
    begin
      colProdCust := colProdCust + ',cust_frete';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := stringreplace(temp, '%', '',[rfReplaceAll, rfIgnoreCase]);
      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;
      //Testa se é Update
      if VerificaUpdate('frete') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_frete='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_frete = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_frete = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_frete = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_frete = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_frete=' + '''' + temp + '''';
      end;
    end
    //CUSTO_REAL (Custo Real)
    else if (LowerCase(StringGrid1.Cells[i,0])='custo_real') then
    begin
      colProdCust := colProdCust + ',cust_custo_real';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;
      //Testa se é Update
      if VerificaUpdate('custo_real') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_custo_real='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_custo_real = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_custo_real = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_custo_real = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_custo_real = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_custo_real=' + '''' + temp + '''';
      end;
    end
    //PRECO_PRAZO (Preço a Prazo)
    else if (LowerCase(StringGrid1.Cells[i,0])='preco_prazo') then
    begin
      colProdCust := colProdCust + ',cust_preco_prazo';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;
      //Testar se existe o preco a vista, se não joga o a prazo mesmo
      if (BuscaColuna(StringGrid1,'preco_vista')=-1) then
      begin
        colProdCust := colProdCust + ',cust_preco_vista';
        dadosProdCust := dadosProdCust + ',' + temp;
      end;
      //Testa se é Update
      if VerificaUpdate('preco_prazo') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_preco_prazo='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_preco_prazo = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_preco_prazo = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_preco_prazo = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_preco_prazo = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_preco_prazo=' + '''' + temp + '''';
      end;
    end
    //PRECO_VISTA (Preço a Vista)
    else if (LowerCase(StringGrid1.Cells[i,0])='preco_vista') then
    begin
      colProdCust := colProdCust + ',cust_preco_vista';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;
      //Testar se existe o preco a prazo, se não joga o a vista mesmo
      if (BuscaColuna(StringGrid1,'preco_prazo')=-1) then
      begin
        colProdCust := colProdCust + ',cust_preco_prazo';
        dadosProdCust := dadosProdCust + ',' + temp;
      end;
      //Testa se é Update
      if VerificaUpdate('preco_vista') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_preco_vista='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_preco_vista = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_preco_vista = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_preco_vista = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_preco_vista = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_preco_vista=' + '''' + temp + '''';
      end;
    end
    //MARGEM (Margem de Valor)
    else if (LowerCase(StringGrid1.Cells[i,0])='margem') then
    begin
      colProdCust := colProdCust + ',cust_margem1';
      if StringGrid1.Cells[i,k]='' then
      begin
        temp := '0';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      temp := stringreplace(temp, '%', '',[rfReplaceAll, rfIgnoreCase]);
      temp := corrigeFloat(temp);
      temp := stringreplace(temp, '.', ',',[rfReplaceAll, rfIgnoreCase]);

      //Testar se precisa multiplicar por 100
      if StrToFloat(temp) < 1 then
        temp := FloatToStr( StrToFloat(temp) * 100 );

      temp := corrigeFloat(temp);
      dadosProdCust := dadosProdCust + ',' + temp;

      //Testa se é Update
      if VerificaUpdate('margem') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_margem1='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select cust_prod_codi from prod_custos where cust_margem1 = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select cust_prod_codi from prod_custos where cust_margem1 = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select cust_prod_codi from prod_custos where cust_margem1 = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select cust_prod_codi from prod_custos where cust_margem1 = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdCust <> '' then dadosUpdateProdCust := dadosUpdateProdCust + ', ';
        dadosUpdateProdCust := dadosUpdateProdCust + 'cust_margem1=' + '''' + temp + '''';
      end;
    end
    //CSOSN
    else if (LowerCase(StringGrid1.Cells[i,0])='csosn') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        temp := '900';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_ESTADUAL';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_INTERESTADUAL';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_ESTA_CF';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_INTER_CF';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';

      //Testa se é Update
      if VerificaUpdate('csosn') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_ESTADUAL = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_SN_CSOSN_ESTADUAL='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_ESTADUAL = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_ESTADUAL = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_ESTADUAL = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_SN_CSOSN_ESTADUAL=' + '''' + temp + '''';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_SN_CSOSN_INTERESTADUAL=' + '''' + temp + '''';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_SN_CSOSN_ESTA_CF=' + '''' + temp + '''';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_SN_CSOSN_INTER_CF=' + '''' + temp + '''';
      end;
    end
    //CSOSN ESTADUAL
    else if (LowerCase(StringGrid1.Cells[i,0])='csosn_esta') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        temp := '900';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_ESTADUAL';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_ESTA_CF';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';

      //Testa se é Update
      if VerificaUpdate('csosn_esta') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_ESTADUAL = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_SN_CSOSN_ESTADUAL='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_ESTADUAL = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_ESTADUAL = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_ESTADUAL = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_SN_CSOSN_ESTADUAL=' + '''' + temp + '''';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_SN_CSOSN_ESTA_CF=' + '''' + temp + '''';
      end;
    end
    //CSOSN INTERESTADUAL
    else if (LowerCase(StringGrid1.Cells[i,0])='csosn_inter') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        temp := '900';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_INTERESTADUAL';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_INTER_CF';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';

      //Testa se é Update
      if VerificaUpdate('csosn_inter') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_INTERESTADUAL = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_SN_CSOSN_INTERESTADUAL='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_INTERESTADUAL = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_INTERESTADUAL = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_SN_CSOSN_INTERESTADUAL = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_SN_CSOSN_INTERESTADUAL=' + '''' + temp + '''';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_SN_CSOSN_INTER_CF=' + '''' + temp + '''';
      end;
    end
    //CST
    else if (LowerCase(StringGrid1.Cells[i,0])='cst') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        temp := '90';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      colProdTrib := colProdTrib + ',TRIB_CST_ICMS_ESTADUAL';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      colProdTrib := colProdTrib + ',TRIB_CST_ICMS_INTERESTADUAL';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      colProdTrib := colProdTrib + ',TRIB_CST_ICMS_ESTA_CF';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      colProdTrib := colProdTrib + ',TRIB_CST_ICMS_INTER_CF';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';

      //Testa se é Update
      if VerificaUpdate('cst') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_CST_ICMS_ESTADUAL='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_ESTADUAL = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_CST_ICMS_ESTADUAL=' + '''' + temp + '''';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_CST_ICMS_INTERESTADUAL=' + '''' + temp + '''';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_CST_ICMS_ESTA_CF=' + '''' + temp + '''';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_CST_ICMS_INTER_CF=' + '''' + temp + '''';
      end;
    end
    //CST ESTADUAL
    else if (LowerCase(StringGrid1.Cells[i,0])='cst_esta') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        temp := '90';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      colProdTrib := colProdTrib + ',TRIB_CST_ICMS_ESTADUAL';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      colProdTrib := colProdTrib + ',TRIB_CST_ICMS_ESTA_CF';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';

      //Testa se é Update
      if VerificaUpdate('cst_esta') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_CST_ICMS_ESTADUAL='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_ESTADUAL = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_CST_ICMS_ESTADUAL=' + '''' + temp + '''';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_CST_ICMS_ESTA_CF=' + '''' + temp + '''';
      end;
    end
    //CST INTERESTADUAL
    else if (LowerCase(StringGrid1.Cells[i,0])='cst_inter') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        temp := '90';
      end
      else begin
        temp := StringGrid1.Cells[i,k];
      end;
      colProdTrib := colProdTrib + ',TRIB_CST_ICMS_INTERESTADUAL';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      colProdTrib := colProdTrib + ',TRIB_CST_ICMS_INTER_CF';
      dadosProdTrib := dadosProdTrib + ',''' + temp + '''';

      //Testa se é Update
      if VerificaUpdate('cst_inter') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_INTERESTADUAL = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_CST_ICMS_INTERESTADUAL='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_INTERESTADUAL = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_INTERESTADUAL = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_CST_ICMS_INTERESTADUAL = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_CST_ICMS_INTERESTADUAL=' + '''' + temp + '''';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_CST_ICMS_INTER_CF=' + '''' + temp + '''';
      end;
    end
    //ALIQ_ICMS (Alíquota de ICMS)
    else if (LowerCase(StringGrid1.Cells[i,0])='aliq_icms') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        colProdTrib := colProdTrib + ',TRIB_ALIQ_ICMS_ESTADUAL';
        dadosProdTrib := dadosProdTrib + ',''' + '0' + '''';
        temp := '0';
      end
      else begin
        temp := stringreplace(StringGrid1.Cells[i,k], '%', '',[rfReplaceAll, rfIgnoreCase]);
        temp := corrigeFloat(temp);
        colProdTrib := colProdTrib + ',TRIB_ALIQ_ICMS_ESTADUAL';
        dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('aliq_icms') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_ALIQ_ICMS_ESTADUAL='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_ICMS_ESTADUAL = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_ALIQ_ICMS_ESTADUAL=' + '''' + temp + '''';
      end;
    end
    //REDU_ESTA (Redução da Alíquota de ICMS Estadual)
    else if (LowerCase(StringGrid1.Cells[i,0])='redu_esta') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        colProdTrib := colProdTrib + ',TRIB_REDU_ICMS_ESTADUAL';
        dadosProdTrib := dadosProdTrib + ',''' + '0' + '''';
        temp := '0';
      end
      else begin
        temp := stringreplace(StringGrid1.Cells[i,k], '%', '',[rfReplaceAll, rfIgnoreCase]);
        temp := corrigeFloat(temp);
        colProdTrib := colProdTrib + ',TRIB_REDU_ICMS_ESTADUAL';
        dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('redu_esta') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_REDU_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_REDU_ICMS_ESTADUAL='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_REDU_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_REDU_ICMS_ESTADUAL = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_REDU_ICMS_ESTADUAL = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_REDU_ICMS_ESTADUAL=' + '''' + temp + '''';
      end;
    end
    //REDU_INTER (Redução da Alíquota de ICMS Interestadual)
    else if (LowerCase(StringGrid1.Cells[i,0])='redu_inter') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        colProdTrib := colProdTrib + ',TRIB_REDU_ICMS_INTERESTADUAL';
        dadosProdTrib := dadosProdTrib + ',''' + '0' + '''';
        temp := '';
      end
      else begin
        temp := stringreplace(StringGrid1.Cells[i,k], '%', '',[rfReplaceAll, rfIgnoreCase]);
        temp := corrigeFloat(temp);
        colProdTrib := colProdTrib + ',TRIB_REDU_ICMS_INTERESTADUAL';
        dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('redu_inter') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_REDU_ICMS_INTERESTADUAL = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_REDU_ICMS_INTERESTADUAL='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_REDU_ICMS_INTERESTADUAL = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_REDU_ICMS_INTERESTADUAL = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_REDU_ICMS_INTERESTADUAL = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_REDU_ICMS_INTERESTADUAL=' + '''' + temp + '''';
      end;
    end
    //CST_IPI (Código de CST IPI)
    else if (LowerCase(StringGrid1.Cells[i,0])='cst_ipi') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        colProdTrib := colProdTrib + ',TRIB_CST_IPI';
        dadosProdTrib := dadosProdTrib + ',' + 'null';
      end
      else begin
        colProdTrib := colProdTrib + ',TRIB_CST_IPI';
        dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('cst_ipi') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_IPI = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_CST_IPI='+''''+StringGrid1.Cells[i,k]+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_IPI = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_CST_IPI = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_CST_IPI = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + ',TRIB_CST_IPI=' + '''' + StringGrid1.Cells[i,k] + '''';
      end;
    end
    //ALIQ_IPI (Alíquota de IPI)
    else if (LowerCase(StringGrid1.Cells[i,0])='aliq_ipi') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        colProdTrib := colProdTrib + ',TRIB_ALIQ_IPI';
        dadosProdTrib := dadosProdTrib + ',''' + '0' + '''';
      end
      else begin
        temp := stringreplace(StringGrid1.Cells[i,k], '%', '',[rfReplaceAll, rfIgnoreCase]);
        temp := corrigeFloat(temp);
        colProdTrib := colProdTrib + ',TRIB_ALIQ_IPI';
        dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('aliq_ipi') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_IPI = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_ALIQ_IPI='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_IPI = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_IPI = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_IPI = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_ALIQ_IPI=' + '''' + temp + '''';
      end;
    end
    //CST_PIS (Código de CST PIS)
    else if (LowerCase(StringGrid1.Cells[i,0])='cst_pis') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        colProdTrib := colProdTrib + ',TRIB_CST_PIS';
        dadosProdTrib := dadosProdTrib + ',''' + '0' + '''';
      end
      else begin
        colProdTrib := colProdTrib + ',TRIB_CST_PIS';
        dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('cst_pis') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_PIS = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_CST_PIS='+''''+StringGrid1.Cells[i,k]+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_PIS = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_CST_PIS = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_CST_PIS = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_CST_PIS=' + '''' + StringGrid1.Cells[i,k] + '''';
      end;
    end
    //ALIQ_PIS (Alíquota de PIS)
    else if (LowerCase(StringGrid1.Cells[i,0])='aliq_pis') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        colProdTrib := colProdTrib + ',TRIB_ALIQ_PIS';
        dadosProdTrib := dadosProdTrib + ',''' + '0' + '''';
        temp := '0';
      end
      else begin
        temp := stringreplace(StringGrid1.Cells[i,k], '%', '',[rfReplaceAll, rfIgnoreCase]);
        temp := corrigeFloat(temp);
        colProdTrib := colProdTrib + ',TRIB_ALIQ_PIS';
        dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('aliq_pis') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_PIS = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_ALIQ_PIS='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_PIS = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_PIS = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_PIS = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_ALIQ_PIS=' + '''' + temp + '''';
      end;
    end
    //CST_COFINS (Código de CST COFINS)
    else if (LowerCase(StringGrid1.Cells[i,0])='cst_cofins') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        colProdTrib := colProdTrib + ',TRIB_CST_COFINS';
        dadosProdTrib := dadosProdTrib + ',''' + '0' + '''';
      end
      else begin
        colProdTrib := colProdTrib + ',TRIB_CST_COFINS';
        dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('cst_cofins') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_COFINS = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_CST_COFINS='+''''+StringGrid1.Cells[i,k]+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_CST_COFINS = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_CST_COFINS = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_CST_COFINS = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_CST_COFINS=' + '''' + StringGrid1.Cells[i,k] + '''';
      end;
    end
    //ALIQ_COFINS (Alíquota de COFINS)
    else if (LowerCase(StringGrid1.Cells[i,0])='aliq_cofins') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        colProdTrib := colProdTrib + ',TRIB_ALIQ_COFINS';
        dadosProdTrib := dadosProdTrib + ',''' + '0' + '''';
        temp := '0';
      end
      else begin
        temp := stringreplace(StringGrid1.Cells[i,k], '%', '',[rfReplaceAll, rfIgnoreCase]);
        temp := corrigeFloat(temp);
        colProdTrib := colProdTrib + ',TRIB_ALIQ_COFINS';
        dadosProdTrib := dadosProdTrib + ',''' + temp + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('aliq_cofins') = 1 then begin
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_COFINS = '+''''+temp+''')';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'TRIB_ALIQ_COFINS='+''''+temp+'''';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_COFINS = '+''''+temp+''')';
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'codi in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_COFINS = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select trib_prod_codi from prod_tributos where TRIB_ALIQ_COFINS = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProdTrib <> '' then dadosUpdateProdTrib := dadosUpdateProdTrib + ', ';
        dadosUpdateProdTrib := dadosUpdateProdTrib + 'TRIB_ALIQ_COFINS=' + '''' + temp + '''';
      end;
    end
    //ATIVO
    else if (LowerCase(StringGrid1.Cells[i,0])='ativo') then
    begin
      if StringGrid1.Cells[i,k]='' then begin
        colProd := colProd + ',ATIVO';
        temp := 'S';
        dadosProd := dadosProd + ',''' + temp + '''';
      end
      else begin
        colProd := colProd + ',ATIVO';
        temp := UpperCase(StringGrid1.Cells[i,k]);
        if (temp='ATIVO') or
           (temp='ATIVAR') or
           (temp='1') or
           (temp='SIM') or
           (temp='S') or
           (temp='OK')
        then begin
          temp := 'S';
        end
        else if (temp='INATIVO') or
                (temp='INATIVAR') or
                (temp='0') or
                (temp='2') or
                (temp='NAO') or
                (temp='N') or
                (temp='NÃO')
        then begin
          temp := 'N';
        end
        else begin
          ShowMessage('Tratar valor da coluna ATIVO: '+temp);
          status := 0;
          Exit; //Quebra o for
        end;

        dadosProd := dadosProd + ',''' + temp + '''';
      end;
      //Testa se é Update
      if VerificaUpdate('ativo') = 1 then begin
        if condUpdateProd <> '' then condUpdateProd := condUpdateProd + ' and ';
        condUpdateProd := condUpdateProd + 'ATIVO='+''''+temp+'''';
        if condUpdateProdTrib <> '' then condUpdateProdTrib := condUpdateProdTrib + ' and ';
        condUpdateProdTrib := condUpdateProdTrib + 'trib_prod_codi in (select codi from prod where ATIVO = '+''''+temp+''')';
        if condUpdateProdAdic <> '' then condUpdateProdAdic := condUpdateProdAdic + ' and ';
        condUpdateProdAdic := condUpdateProdAdic + 'adic_prod_codi in (select codi from prod where ATIVO = '+''''+temp+''')';
        if condUpdateProdCust <> '' then condUpdateProdCust := condUpdateProdCust + ' and ';
        condUpdateProdCust := condUpdateProdCust + 'cust_prod_codi in (select codi from prod where ATIVO = '+''''+temp+''')';
        if condUpdateItens <> '' then condUpdateItens := condUpdateItens + ' and ';
        condUpdateItens := condUpdateItens + 'cod_prod in (select codi from prod where ATIVO = '+''''+temp+''')';
      end
      else begin
        if dadosUpdateProd <> '' then dadosUpdateProd := dadosUpdateProd + ', ';
        dadosUpdateProd := dadosUpdateProd + 'ATIVO=' + '''' + temp + '''';
      end;
    end
    ;

  //Fim do for
  end;
end;

procedure cImportaProduto.Gravar;
var
  fileTXT: TextFile;
  SQL: TSQLDataSet;
  i, j: Integer;
  temp: string;
begin
  //Sair se estiver em erro
  if status = 0 then
    Exit;

  //----------------------------------------
  //Gravar no banco de dados
  if UpperCase( ExtractFileExt(frmPrinc.DBPath.Text) ) = '.FDB' then begin
    try
      try
        //Abrir conexoes
        frmPrinc.conDestino.Open;
        SQL := TSQLDataSet.Create(Nil);
        SQL.SQLConnection := frmPrinc.conDestino;

        //Se for INSERT
        if colUpdateCount <= 0 then begin

          //Executar INSERTs
          frmImportando.atualizaStatus('Inserindo dados na tabela PROD.');
          SQL.CommandText := 'insert into prod ('+ colProd +') values ' + '(' + dadosProd + ');';
          SQL.ExecSQL;

          //Criar registros em todas as empresas
          {
            CUST_PROD_EMPR: Empresa onde foi criado
            CUST_EMPR: Empresa que irá aparecer, por isso tem esse FOR
          }
          j := BuscaColuna(StringGrid1,'empr');
          //Se nao achar seta padrão
          if j = -1 then temp := '1'
          //Se achar recebe valor no temp
          else temp := StringGrid1.Cells[j,k];

          //Criar registros em todas as empresas
          {
            CUST_PROD_EMPR: Empresa onde foi criado
            CUST_EMPR: Empresa que irá aparecer, por isso tem esse FOR
          }

          for i := 1 to qtdEmpr do begin
            //Testa se é a empresa onde o produto foi cadastrado
            if i = StrToInt(temp) then begin
              frmImportando.atualizaStatus('Inserindo dados na tabela PROD_TRIBUTOS.');
              SQL.CommandText := 'insert into prod_tributos ('+ colProdTrib +') values ' + '(' + dadosProdTrib + ');';
              SQL.ExecSQL;
              frmImportando.atualizaStatus('Inserindo dados na tabela PROD_ADICIONAIS.');
              SQL.CommandText := 'insert into prod_adicionais ('+ colProdAdic +') values ' + '(' + dadosProdAdic + ');';
              SQL.ExecSQL;
              frmImportando.atualizaStatus('Inserindo dados na tabela PROD_CUSTOS.');
              SQL.CommandText := 'insert into prod_custos ('+ colProdCust +') values ' + '(' + dadosProdCust + ');';
              SQL.ExecSQL;
              frmImportando.atualizaStatus('Inserindo dados na tabela MVA.');
              SQL.CommandText := 'insert into mva ('+ colMVA +') values ' + '(' + dadosMVA + ');';
              SQL.ExecSQL;
              frmImportando.atualizaStatus('Inserindo dados na tabela ITENS.');
              SQL.CommandText := 'insert into itens ('+ colItens +') values ' + '(' + dadosItens + ');';
              SQL.ExecSQL;
              frmImportando.atualizaStatus('Inserindo dados na tabela PROD_FORN.');
              SQL.CommandText := 'insert into prod_forn ('+ colProdForn +') values ' + '(' + dadosProdForn + ');';
              SQL.ExecSQL;
            end
            //Se não, cria registro em branco na outra empresa
            else begin
              frmImportando.atualizaStatus('Inserindo registro na tabela PROD_TRIBUTOS para Empresa '+IntToStr(i));
              SQL.CommandText := 'insert into prod_tributos ('+ colRegistroProdTrib + ',trib_empr) values ' + '(' + dadosRegistroProdTrib + ','+IntToStr(i)+');';
              SQL.ExecSQL;
              frmImportando.atualizaStatus('Inserindo registro na tabela PROD_ADICIONAIS para Empresa '+IntToStr(i));
              SQL.CommandText := 'insert into prod_adicionais ('+ colRegistroProdAdic +',adic_empr) values ' + '(' + dadosRegistroProdAdic + ','+IntToStr(i)+ ');';
              SQL.ExecSQL;
              frmImportando.atualizaStatus('Inserindo registro na tabela PROD_CUSTOS para Empresa '+IntToStr(i));
              SQL.CommandText := 'insert into prod_custos ('+ colRegistroProdCust +',cust_empr) values ' + '(' + dadosRegistroProdCust + ','+IntToStr(i)+ ');';
              SQL.ExecSQL;
              frmImportando.atualizaStatus('Inserindo registro na tabela MVA para Empresa '+IntToStr(i));
              SQL.CommandText := 'insert into mva ('+ colRegistroMVA +',mva_empr) values ' + '(' + dadosRegistroMVA + ','+IntToStr(i)+ ');';
              SQL.ExecSQL;
            end;
          end;

        end
        //Se for UPDATE
        else begin
          //Executar UPDATE
          if dadosUpdateProd <> '' then begin
            SQL.CommandText := 'update prod set '+ dadosUpdateProd +' where ' + condUpdateProd + ';';
            SQL.ExecSQL;
          end;
          if dadosUpdateProdTrib <> '' then begin
            SQL.CommandText := 'update prod_tributos set '+ dadosUpdateProdTrib +' where ' + condUpdateProdTrib + ';';
            SQL.ExecSQL;
          end;
          if dadosUpdateProdAdic <> '' then begin
            SQL.CommandText := 'update prod_adicionais set '+ dadosUpdateProdAdic +' where ' + condUpdateProdAdic + ';';
            SQL.ExecSQL;
          end;
          if dadosUpdateProdCust <> '' then begin
            SQL.CommandText := 'update prod_custos set '+ dadosUpdateProdCust +' where ' + condUpdateProdCust + ';';
            SQL.ExecSQL;
          end;
          if dadosUpdateItens <> '' then begin
            SQL.CommandText := dadosUpdateItens +' and ' + condUpdateItens + ';';
            SQL.ExecSQL;
          end;
        end;
      except
        on e: exception do
        begin
          Mensagem('Erro SQL: '+e.message+sLineBreak+SQL.CommandText,mtCustom,[],[],'Erro SQL Produtos');
          //ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
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

        frmImportando.atualizaStatus('Comandos da PROD.');
        WriteLn(fileTXT, '----------Comandos da PROD----------');

        //Se for INSERT
        if colUpdateCount <= 0 then begin

          frmImportando.atualizaStatus('Inserindo dados na tabela PROD.');
          WriteLn(fileTXT, 'insert into prod ('+ colProd +') values ' + '(' + dadosProd + ');');
          WriteLn(fileTXT, 'COMMIT WORK;');

          //Criar registros em todas as empresas
          {
            CUST_PROD_EMPR: Empresa onde foi criado
            CUST_EMPR: Empresa que irá aparecer, por isso tem esse FOR
          }
          j := BuscaColuna(StringGrid1,'empr');
          //Se nao achar seta padrão
          if j = -1 then temp := '1'
          //Se achar recebe valor no temp
          else temp := StringGrid1.Cells[j,k];

          //Criar registros em todas as empresas
          {
            CUST_PROD_EMPR: Empresa onde foi criado
            CUST_EMPR: Empresa que irá aparecer, por isso tem esse FOR
          }

          for i := 1 to qtdEmpr do begin
            //Testa se é a empresa onde o produto foi cadastrado
            if i = StrToInt(temp) then begin
              frmImportando.atualizaStatus('Inserindo dados na tabela PROD_TRIBUTOS.');
              WriteLn(fileTXT, 'insert into prod_tributos ('+ colProdTrib +') values ' + '(' + dadosProdTrib + ');');
              WriteLn(fileTXT, 'COMMIT WORK;');
              frmImportando.atualizaStatus('Inserindo dados na tabela PROD_ADICIONAIS.');
              WriteLn(fileTXT, 'insert into prod_adicionais ('+ colProdAdic +') values ' + '(' + dadosProdAdic + ');');
              WriteLn(fileTXT, 'COMMIT WORK;');
              frmImportando.atualizaStatus('Inserindo dados na tabela PROD_CUSTOS.');
              WriteLn(fileTXT, 'insert into prod_custos ('+ colProdCust +') values ' + '(' + dadosProdCust + ');');
              WriteLn(fileTXT, 'COMMIT WORK;');
              frmImportando.atualizaStatus('Inserindo dados na tabela MVA.');
              WriteLn(fileTXT, 'insert into mva ('+ colMVA +') values ' + '(' + dadosMVA + ');');
              WriteLn(fileTXT, 'COMMIT WORK;');
              frmImportando.atualizaStatus('Inserindo dados na tabela ITENS.');
              WriteLn(fileTXT, 'insert into itens ('+ colItens +') values ' + '(' + dadosItens + ');');
              WriteLn(fileTXT, 'COMMIT WORK;');
              frmImportando.atualizaStatus('Inserindo dados na tabela PROD_FORN.');
              WriteLn(fileTXT, 'insert into prod_forn ('+ colProdForn +') values ' + '(' + dadosProdForn + ');');
              WriteLn(fileTXT, 'COMMIT WORK;');
            end
            //Se não, cria registro em branco na outra empresa
            else begin
              frmImportando.atualizaStatus('Inserindo registro na tabela PROD_TRIBUTOS para Empresa '+IntToStr(i));
              WriteLn(fileTXT, 'insert into prod_tributos ('+ colRegistroProdTrib + ',trib_empr) values ' + '(' + dadosRegistroProdTrib + ','+IntToStr(i)+');');
              WriteLn(fileTXT, 'COMMIT WORK;');
              frmImportando.atualizaStatus('Inserindo registro na tabela PROD_ADICIONAIS para Empresa '+IntToStr(i));
              WriteLn(fileTXT, 'insert into prod_adicionais ('+ colRegistroProdAdic +',adic_empr) values ' + '(' + dadosRegistroProdAdic + ','+IntToStr(i)+ ');');
              WriteLn(fileTXT, 'COMMIT WORK;');
              frmImportando.atualizaStatus('Inserindo registro na tabela PROD_CUSTOS para Empresa '+IntToStr(i));
              WriteLn(fileTXT, 'insert into prod_custos ('+ colRegistroProdCust +',cust_empr) values ' + '(' + dadosRegistroProdCust + ','+IntToStr(i)+ ');');
              WriteLn(fileTXT, 'COMMIT WORK;');
              frmImportando.atualizaStatus('Inserindo registro na tabela MVA para Empresa '+IntToStr(i));
              WriteLn(fileTXT, 'insert into mva ('+ colRegistroMVA +',mva_empr) values ' + '(' + dadosRegistroMVA + ','+IntToStr(i)+ ');');
              WriteLn(fileTXT, 'COMMIT WORK;');
            end;
          end;
        end

        //Se for UPDATE
        else begin
          frmImportando.atualizaStatus('Atualizando dados na tabela PROD.');

          //Executar UPDATE
          if dadosUpdateProd <> '' then begin
            WriteLn(fileTXT, 'update prod set '+ dadosUpdateProd +' where ' + condUpdateProd + ';');
            WriteLn(fileTXT, 'COMMIT WORK;');
          end;
          if dadosUpdateProdTrib <> '' then begin
            WriteLn(fileTXT, 'update prod_tributos set '+ dadosUpdateProdTrib +' where ' + condUpdateProdTrib + ';');
            WriteLn(fileTXT, 'COMMIT WORK;');
          end;
          if dadosUpdateProdAdic <> '' then begin
            WriteLn(fileTXT, 'update prod_adicionais set '+ dadosUpdateProdAdic +' where ' + condUpdateProdAdic + ';');
            WriteLn(fileTXT, 'COMMIT WORK;');
          end;
          if dadosUpdateProdCust <> '' then begin
            WriteLn(fileTXT, 'update prod_custos set '+ dadosUpdateProdCust +' where ' + condUpdateProdCust + ';');
            WriteLn(fileTXT, 'COMMIT WORK;');
          end;
          if dadosUpdateItens <> '' then begin
            WriteLn(fileTXT, dadosUpdateItens +' and ' + condUpdateItens + ';');
            WriteLn(fileTXT, 'COMMIT WORK;');
          end;
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
