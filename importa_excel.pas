unit importa_excel;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.StdCtrls, Vcl.Buttons, ComObj, IniFiles,
  Vcl.FileCtrl, Data.DBXFirebird, Data.DB, Data.SqlExpr, importando, OleAuto,
  Vcl.Menus;

type
  TForm1 = class(TForm)
    BtnLoad: TBitBtn;
    OpenDialog1: TOpenDialog;
    StringGrid1: TStringGrid;
    FilePath: TEdit;
    BtnAbrir: TBitBtn;
    SelectImport: TComboBox;
    ButImport: TBitBtn;
    DBPath: TEdit;
    ButOpenDB: TBitBtn;
    OpenDialog2: TOpenDialog;
    Connect: TSQLConnection;
    ButSave: TBitBtn;
    SaveDialog1: TSaveDialog;
    Menu: TMainMenu;
    t1: TMenuItem;
    Editar1: TMenuItem;
    Cabealho1: TMenuItem;
    Limpar: TMenuItem;
    LimpaClieForn: TMenuItem;
    LimpaGrupos: TMenuItem;
    LimpaSubGrupos: TMenuItem;
    LimpaMarcas: TMenuItem;
    LimpaProdutos: TMenuItem;
    LimpaTituP: TMenuItem;
    LimpaTituR: TMenuItem;

    //function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
    procedure BtnAbrirClick(Sender: TObject);
    procedure AutoSizeCol(Grid: TStringGrid; Column: integer);
    procedure BtnLoadClick(Sender: TObject);
    procedure BtnOpenDB(Sender: TObject);
    procedure StringGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure StringGrid1Click(Sender: TObject);
    procedure StringGrid1DblClick(Sender: TObject);
    procedure StringGrid1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ButImportClick(Sender: TObject);
    procedure ButSaveClick(Sender: TObject);
    procedure Cabealho1Click(Sender: TObject);
    procedure LimpaClieFornClick(Sender: TObject);
    procedure LimpaGruposClick(Sender: TObject);
    procedure LimpaSubGruposClick(Sender: TObject);
    procedure LimpaMarcasClick(Sender: TObject);
    procedure LimpaProdutosClick(Sender: TObject);
    procedure LimpaTituPClick(Sender: TObject);
    procedure LimpaTituRClick(Sender: TObject);


  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  gridTemp: Array of Array of string;

implementation


//Fun��o para carregar apenas o cabe�alho de uma planilha na StringGrid
function XlsHeaderLoad(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
  xlCellTypeLastCell = $0000000B;
var
  XLApp, Sheet: OLEVariant;
  RangeMatrix: Variant;
  x, y, k, r: Integer;
begin
  Result := False;
  // Create Excel-OLE Object
  XLApp := CreateOleObject('Excel.Application');
  try
    // Hide Excel
    XLApp.Visible := False;

    // Open the Workbook
    XLApp.Workbooks.Open(AXLSFile);

    // Sheet := XLApp.Workbooks[1].WorkSheets[1];
    Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];

    // In order to know the dimension of the WorkSheet, i.e the number of rows
    // and the number of columns, we activate the last non-empty cell of it

    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    // Get the value of the last row
    x := XLApp.ActiveCell.Row;
    // Get the value of the last column
    y := XLApp.ActiveCell.Column;

    // Assign the Variant associated with the WorkSheet to the Delphi Variant
    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;


    //Verificar quantidade de colunas
    if AGrid.ColCount < y then AGrid.ColCount := y;
    //Testar se a primeira coluna n�o � um '0'
    if RangeMatrix[1, 1] <> '0' then
    begin
      //AGrid.ColCount := AGrid.ColCount +1;
      AGrid.Cells[0,0] := '0';
    end;

    //Iterar colunas
    for r := 1 to y do
    begin
      AGrid.Cells[r,0] := RangeMatrix[1, r];
    end;

    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;


  finally
    // Quit Excel
    if not VarIsEmpty(XLApp) then
    begin
      // XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLAPP := Unassigned;
      Sheet := Unassigned;
      Result := True;
    end;
  end;
end;


//Fun��o para carregar planilha na StringGrid
function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
  xlCellTypeLastCell = $0000000B;
var
  XLApp, Sheet: OLEVariant;
  RangeMatrix: Variant;
  x, y, k, r: Integer;
begin
  Result := False;
  // Create Excel-OLE Object
  XLApp := CreateOleObject('Excel.Application');
  try
    // Hide Excel
    XLApp.Visible := False;

    // Open the Workbook
    XLApp.Workbooks.Open(AXLSFile);

    // Sheet := XLApp.Workbooks[1].WorkSheets[1];
    Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];

    // In order to know the dimension of the WorkSheet, i.e the number of rows
    // and the number of columns, we activate the last non-empty cell of it

    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    // Get the value of the last row
    x := XLApp.ActiveCell.Row;
    // Get the value of the last column
    y := XLApp.ActiveCell.Column;

    // Set Stringgrid's row &col dimensions.

    AGrid.RowCount := x;
    AGrid.ColCount := y;

    // Assign the Variant associated with the WorkSheet to the Delphi Variant
    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;

    {
    //  Define the loop for filling in the TStringGrid
    k := 1;
    repeat
      for r := 1 to y do
        AGrid.Cells[(r - 1), (k - 1)] := RangeMatrix[K, R];
      Inc(k, 1);
    until k > x;  }


    AGrid.ColCount := AGrid.ColCount +1;
    //Iterar linhas
    for k := 0 to AGrid.RowCount-1 do
    begin
      AGrid.Cells[0,k] := IntToStr(k);
      //Iterar colunas
      for r := 1 to AGrid.ColCount-1 do
      begin
        AGrid.Cells[r,k] := RangeMatrix[(K+1), r];
      end;
    end;

    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;


  finally
    // Quit Excel
    if not VarIsEmpty(XLApp) then
    begin
      // XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLAPP := Unassigned;
      Sheet := Unassigned;
      Result := True;
    end;
  end;
end;

//Fun��o para criar caixa de di�logos
function Mensagem(CONST Msg: string; DlgTypt: TmsgDlgType; button: TMsgDlgButtons;
  Caption: ARRAY OF string; dlgcaption: string): Integer;
var
  aMsgdlg: TForm;
  i: Integer;
  Dlgbutton: Tbutton;
  Captionindex: Integer;
begin
  aMsgdlg := createMessageDialog(Msg, DlgTypt, button);
  aMsgdlg.Caption := dlgcaption;
  aMsgdlg.BiDiMode := bdRightToLeft;
  Captionindex := 0;
  for i := 0 to aMsgdlg.componentcount - 1 Do
  begin
    if (aMsgdlg.components[i] is Tbutton) then
    Begin
      Dlgbutton := Tbutton(aMsgdlg.components[i]);
      if Captionindex <= High(Caption) then
        Dlgbutton.Caption := Caption[Captionindex];
      inc(Captionindex);
    end;
  end;
  Result := aMsgdlg.Showmodal;
end;


//Bot�o para selecionar arquivo
procedure TForm1.BtnAbrirClick(Sender: TObject);
var
  arquivo : String;

begin
  if OpenDialog1.Execute then
  begin
    arquivo := ExtractFilePath(Application.ExeName);
    FilePath.Text := OpenDialog1.FileName;
  end;

end;


//Bot�o para selecionar arquivo
procedure TForm1.BtnOpenDB(Sender: TObject);
var
  arquivo : String;

begin
  if OpenDialog2.Execute then
  begin
    arquivo := ExtractFilePath(Application.ExeName);
    DBPath.Text := OpenDialog2.FileName;
    Connect.Params.Values['DataBase'] := DBPath.Text;
  end;

end;


//Fun��o para redimensionar coluna
procedure TForm1.AutoSizeCol(Grid: TStringGrid; Column: integer);
var
  i, W, WMax: integer;
begin
  WMax := 0;
  for i := 0 to (Grid.RowCount - 1) do begin
    W := Grid.Canvas.TextWidth(Grid.Cells[Column, i]);
    if W > WMax then
      WMax := W;
  end;
  Grid.ColWidths[Column] := WMax + 5;

  //Se for uma coluna nova
  if Grid.ColWidths[Column] = 5 then
  begin
    Grid.ColWidths[Column] := 40;
  end;

end;


//Bot�o para Carregar arquivo Excel na StringGrid
procedure TForm1.BtnLoadClick(Sender: TObject);
var
  i: integer;

begin
  //Carregar Excel na StringGrid
  Xls_To_StringGrid(StringGrid1, FilePath.Text);

  //Redimensionar colunas
  for i := 0 to StringGrid1.ColCount - 1 do
    AutoSizeCol(StringGrid1, i);

end;


//FUN��O DO JEFINHO PARA REMOVER ACENTOS
function RemoveAcento(Str: string): string;
const
  ComAcento = '����������������������������';
  SemAcento = 'aaeouaoaeioucuAAEOUAOAEIOUCU';
var
  x: Integer;
begin;
  for x := 1 to Length(Str) do
    if Pos(Str[x], ComAcento) <> 0 then
      Str[x] := SemAcento[Pos(Str[x], ComAcento)];
  Result := Str;
end;


//Fun��o para buscar a coluna no StringGrid e retornar o indice
function BuscaColuna(Grid: TStringGrid; colName: String) : Integer;
var
  i: integer;
begin
  colName := UpperCase(colName);

  for i := 0 to Grid.ColCount-1 do
  begin
    if UpperCase(Grid.Cells[i,0]) = colName then
      Break;
  end;
  if i = Grid.ColCount then
  begin
    Result:=-1;
  end
  else begin
    Result:=i;
  end;
end;


//Fun��o para buscar a cidade no banco.
function buscaCidade(Cidade, UF: string): Integer;
var
  queryTemp: TSQLQuery;

begin
  try
    Form1.Connect.Open;
    queryTemp := TSQLQuery.Create(nil);
    queryTemp.SQLConnection := Form1.Connect;
    queryTemp.SQL.Clear;
    queryTemp.SQL.Add('SELECT * FROM CIDADE WHERE CID_DESC = :PDESC AND CID_UF = :PUF');
    queryTemp.ParamByName('PDESC').AsString := Cidade;
    queryTemp.ParamByName('PUF').AsString := UF;
    queryTemp.Open;
  finally
    if queryTemp.IsEmpty then
    begin
      Result := -1;
    end
    else begin
      Result := queryTemp.FieldByName('CID_CODI').AsInteger;
    end;
    queryTemp.Close;
    Form1.Connect.Close;
  end;
end;


//FUN��O PARA RECONHECER SE JA EXISTE O CODIGO DO TITULO PAGAR OU NAO
function temCodTituloP(Codigo: String): Boolean;
var
  queryTemp: TSQLQuery;

begin
  try
    Form1.Connect.Open;
    queryTemp := TSQLQuery.Create(nil);
    queryTemp.SQLConnection := Form1.Connect;
    queryTemp.SQL.Clear;
    queryTemp.SQL.Add('select tp.codi from titup tp where tp.codi = :PCODI');
    queryTemp.ParamByName('PCODI').AsString := Codigo;
    queryTemp.Open;

    if queryTemp.IsEmpty = True then
      Result := False
    else
      Result := True;
  finally
    queryTemp.Free;
    Form1.Connect.Close;
  end;
end;


//FUN��O PARA RECONHECER SE JA EXISTE O CODIGO DO TITULO RECEBER OU NAO
function temCodTituloR(Codigo: String): Boolean;
var
  queryTemp: TSQLQuery;

begin
  try
    Form1.Connect.Open;
    queryTemp := TSQLQuery.Create(nil);
    queryTemp.SQLConnection := Form1.Connect;
    queryTemp.SQL.Clear;
    queryTemp.SQL.Add('select tr.codi from titur tr where tr.codi = :PCODI');
    queryTemp.ParamByName('PCODI').AsString := Codigo;
    queryTemp.Open;

    if queryTemp.IsEmpty = True then
      Result := False
    else
      Result := True;
  finally
    queryTemp.Free;
    Form1.Connect.Close;
  end;
end;


//IMPORTAR DADOS
procedure TForm1.ButImportClick(Sender: TObject);
var
  SQL: TSQLDataSet;
  temp, temp2: String;
  colClieForn, dadosClieForn: String;
  colProd, dadosProd: String;
  colProdTrib, dadosProdTrib: String;
  colProdAdic, dadosProdAdic: String;
  colProdCust, dadosProdCust: String;
  colItens, dadosItens: String;
  colMVA, dadosMVA: String;
  colProdForn, dadosProdForn: String;
  colGrupo, dadosGrupo: String;
  colSubGrupo, dadosSubGrupo: String;
  colMarca, dadosMarca: String;
  colTituP, dadosTituP: string;
  colTituR, dadosTituR: string;
  i,k,status,max,count: integer;

begin

  //Status se esta OK ou se tem erro, setado como OK
  status := 1;

  //Valor do codigo maximo para atualizar o generator
  max := 0;

  try
    try
      //Criar tela de loading
      Form2.Show;
      Form2.Label2.Font.Color := clBlack;

      for k := 1 to StringGrid1.RowCount-1 do
      begin

        //Se clicou em cancelar, quebra o la�o das linhas e finaliza importa��o.
        if Form2.Active=False then break;

        //Atualizar Status
        Form2.atualizaItens(k,StringGrid1.RowCount-1);

        //Importar Clientes e Fornecedores
        if SelectImport.Text = 'Clie/Forn' then
        begin
          //ShowMessage('Importar Clie/Forn');

          Form2.atualizaStatus('Clie/Forn '+ IntToStr(k));

          colClieForn := '';
          dadosClieForn := '';

          //Carregar informa��es para importar
          //-------------------------------------------------------

          //Codigo � obrigat�rio, se n�o tiver preenche com o generator
          //CODI (CODIGO)
          i:=BuscaColuna(StringGrid1,'codi');
          if (i<>-1) then
          begin
            if (StringGrid1.Cells[i,k]='') then
            begin
              ShowMessage('C�digo em branco na linha '+IntToStr(k));
            end
            else begin
              if StrToInt(StringGrid1.Cells[i,k])>max then max:=StrToInt(StringGrid1.Cells[i,k]);
              colClieForn := colClieForn + 'codi';
              dadosClieForn := dadosClieForn + '''' + StringGrid1.Cells[i,k] + '''';
            end;
          end
          else begin
            colClieForn := colClieForn + 'codi';
            dadosClieForn := dadosClieForn + 'gen_id(gen_clieforn_id,1)';
          end;

          //Importar cidade se tiver, precisa ter o UF antes de importar a cidade
          //UF
          temp := '';
          i:=BuscaColuna(StringGrid1,'uf');
          if (i<>-1) then
          begin
            if (StringGrid1.Cells[i,k]<>'') then
            begin
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              colClieForn := colClieForn + ',uf';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end;
          end;
          //CIDA (CIDADE)
          i:=BuscaColuna(StringGrid1,'cida');
          if ((i<>-1) and (temp<>'')) then
          begin
            temp2 := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
            temp := IntToStr(buscaCidade(temp2, temp));
            if StrToInt(temp) > 0 then
            begin
              colClieForn := colClieForn + ',cida,codi_cida';
              dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end;
          end;


          for i := 0 to StringGrid1.ColCount-1 do
          begin
            //NOME
            if ((LowerCase(StringGrid1.Cells[i,0])='nome') and (StringGrid1.Cells[i,0]<>'')) then
            begin
              colClieForn := colClieForn + ',nome';
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp2 := (Copy(temp,1,60));
              dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
            end
            //NOME_FANT (NOME FANTASIA)
            else if (LowerCase(StringGrid1.Cells[i,0])='nome_fant') then
            begin
              colClieForn := colClieForn + ',nome_fant';
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp2 := (Copy(temp,1,60));
              dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
            end
            //DATA_NASC
            else if (LowerCase(StringGrid1.Cells[i,0])='data_nasc') then
            begin
              temp := Trim(StringGrid1.Cells[i,k]);
              temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              if temp.Length = 8 then
              begin
                temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ (Copy(temp,5,4));
                colClieForn := colClieForn + ',data_nasc';
                dadosClieForn := dadosClieForn + ',''' + temp + '''';
              end
              else if temp.Length = 6 then begin
                temp2 := (Copy(DateToStr(Date()),9,2));
                //Testa os dois ultimos caracteres da data atual com nascimento do cliente
                //Se os caracteres da data de nascimento do cliente forem maiores, significa que � um s�culo antes
                if StrToInt(temp2)<StrToInt(Copy(temp,5,2)) then temp2 := IntToStr(StrToInt(temp2)-1);
                temp := (Copy(temp,1,2)) +'.'+ (Copy(temp,3,2)) +'.'+ temp2 + (Copy(temp,5,2));
                colClieForn := colClieForn + ',data_nasc';
                dadosClieForn := dadosClieForn + ',''' + temp + '''';
              end;
            end
            //CPF
            else if (LowerCase(StringGrid1.Cells[i,0])='cpf') then
            begin
              temp := Trim(StringGrid1.Cells[i,k]);
              temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              if temp.Length = 11 then
              begin
                temp := (Copy(temp,1,3))+ '.' + (Copy(temp,4,3)) + '.' + (Copy(temp,7,3)) + '-' + (Copy(temp,10,2));
                colClieForn := colClieForn + ',cpf';
                dadosClieForn := dadosClieForn + ',''' + temp + '''';
              end;
            end
            //CNPJ
            else if (LowerCase(StringGrid1.Cells[i,0])='cnpj') then
            begin
              temp := Trim(StringGrid1.Cells[i,k]);
              temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              if temp.Length = 14 then
              begin
                temp := (Copy(temp,1,2))+ '.' + (Copy(temp,3,3)) + '.' + (Copy(temp,6,3)) + '/' + (Copy(temp,9,4)) + '-' + (Copy(temp,13,2));
                colClieForn := colClieForn + ',cnpj';
                dadosClieForn := dadosClieForn + ',''' + temp + '''';
              end;
            end
            //CPF OU CNPJ NO MESMO CAMPO
            else if (LowerCase(StringGrid1.Cells[i,0])='cpf_cnpj') then
            begin
              temp := Trim(StringGrid1.Cells[i,k]);
              temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              if temp.Length = 11 then
              begin
                temp2 := (Copy(temp,1,3))+ '.' + (Copy(temp,4,3)) + '.' + (Copy(temp,7,3)) + '-' + (Copy(temp,10,2));
                colClieForn := colClieForn + ',cpf';
                dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
              end
              else if temp.Length = 14 then
              begin
                temp2 := (Copy(temp,1,2))+ '.' + (Copy(temp,3,3)) + '.' + (Copy(temp,6,3)) + '/' + (Copy(temp,9,4)) + '-' + (Copy(temp,13,2));
                colClieForn := colClieForn + ',cnpj';
                dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
              end;
            end
            //RG
            else if (LowerCase(StringGrid1.Cells[i,0])='rg') then
            begin
              temp := Trim(StringGrid1.Cells[i,k]);
              temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              if temp.Length > 1 then
              begin
                colClieForn := colClieForn + ',rg';
                dadosClieForn := dadosClieForn + ',''' + temp + '''';
              end
              else begin
                colClieForn := colClieForn + ',rg';
                dadosClieForn := dadosClieForn + ',''' + 'ISENTO' + '''';
              end;
            end
            //INSC (INSCRICAO ESTADUAL-IE)
            else if (LowerCase(StringGrid1.Cells[i,0])='insc') then
            begin
              temp := Trim(StringGrid1.Cells[i,k]);
              temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              if temp.Length > 1 then
              begin
                colClieForn := colClieForn + ',insc';
                dadosClieForn := dadosClieForn + ',''' + temp + '''';
              end
              else begin
                colClieForn := colClieForn + ',insc';
                dadosClieForn := dadosClieForn + ',''' + 'ISENTO' + '''';
              end;
            end
            //ENDE (ENDERECO)
            else if (LowerCase(StringGrid1.Cells[i,0])='ende') then
            begin
              colClieForn := colClieForn + ',ende';
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,60));
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //BAIR (BAIRRO)
            else if (LowerCase(StringGrid1.Cells[i,0])='bair') then
            begin
              colClieForn := colClieForn + ',bair';
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,30));
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //COMP (COMPLEMENTO)
            else if (LowerCase(StringGrid1.Cells[i,0])='comp') then
            begin
              colClieForn := colClieForn + ',comp';
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,30));
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //CEP
            else if (LowerCase(StringGrid1.Cells[i,0])='cep') then
            begin
              temp := Trim(StringGrid1.Cells[i,k]);
              temp := stringreplace(temp, '-', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '/', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              colClieForn := colClieForn + ',cep';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //FONE
            else if (LowerCase(StringGrid1.Cells[i,0])='fone') then
            begin
              temp := StringGrid1.Cells[i,k];
              colClieForn := colClieForn + ',fone';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //FONE2
            else if (LowerCase(StringGrid1.Cells[i,0])='fone2') then
            begin
              temp := StringGrid1.Cells[i,k];
              colClieForn := colClieForn + ',fone2';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //FONE_FIRM
            else if (LowerCase(StringGrid1.Cells[i,0])='fone_firm') then
            begin
              temp := StringGrid1.Cells[i,k];
              colClieForn := colClieForn + ',fone_firm';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //FAX
            else if (LowerCase(StringGrid1.Cells[i,0])='fax') then
            begin
              temp := StringGrid1.Cells[i,k];
              colClieForn := colClieForn + ',fax';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //FIRM (Firma ou Empresa que trabalha)
            else if (LowerCase(StringGrid1.Cells[i,0])='firm') then
            begin
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,60));
              colClieForn := colClieForn + ',firm';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //ESTA_CIVI (Estado Civil)
            else if (LowerCase(StringGrid1.Cells[i,0])='esta_civi') then
            begin
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              colClieForn := colClieForn + ',esta_civi';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //NOME_PAI
            else if (LowerCase(StringGrid1.Cells[i,0])='nome_pai') then
            begin
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,60));
              colClieForn := colClieForn + ',nome_pai';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //NOME_MAE
            else if (LowerCase(StringGrid1.Cells[i,0])='nome_mae') then
            begin
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,60));
              colClieForn := colClieForn + ',nome_mae';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //CONJ (Nome do Conjuge)
            else if (LowerCase(StringGrid1.Cells[i,0])='conj') then
            begin
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,60));
              colClieForn := colClieForn + ',conj';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //CONJ_FIRM (Trabalho do Conjuge)
            else if (LowerCase(StringGrid1.Cells[i,0])='conj_firm') then
            begin
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,60));
              colClieForn := colClieForn + ',conj_firm';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //OBS
            else if (LowerCase(StringGrid1.Cells[i,0])='obs') then
            begin
              colClieForn := colClieForn + ',obs';
              temp := StringGrid1.Cells[i,k];
              temp := (Copy(temp,1,80));
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //REFE_COME (Referencia Comercial)
            else if (LowerCase(StringGrid1.Cells[i,0])='refe_come') then
            begin
              temp := StringGrid1.Cells[i,k];
              if (temp.Length>220) then
              begin
                colClieForn := colClieForn + ',refe_come3';
                temp2 := (Copy(temp,221,110));
                dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
              end;
              if (temp.Length>110) then
              begin
                colClieForn := colClieForn + ',refe_come2';
                temp2 := (Copy(temp,111,110));
                dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
              end;
              colClieForn := colClieForn + ',refe_come1';
              temp2 := (Copy(temp,1,110));
              dadosClieForn := dadosClieForn + ',''' + temp2 + '''';
            end
            //MAIL (EMAIL)
            else if (LowerCase(StringGrid1.Cells[i,0])='mail') then
            begin
              temp := Trim(StringGrid1.Cells[i,k]);
              colClieForn := colClieForn + ',mail';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //TIPO (SEXO 'M' ou 'F')
            else if (LowerCase(StringGrid1.Cells[i,0])='tipo') then
            begin
              temp := UpperCase(Trim(StringGrid1.Cells[i,k]));
              if ((temp='MASCULINO') or (temp='MASC') or (temp='M')) then
              begin
                temp := 'M';
                colClieForn := colClieForn + ',tipo';
                dadosClieForn := dadosClieForn + ',''' + temp + '''';
              end
              else if ((temp='FEMININO') or (temp='FEM') or (temp='F')) then
              begin
                temp := 'F';
                colClieForn := colClieForn + ',tipo';
                dadosClieForn := dadosClieForn + ',''' + temp + '''';
              end;
            end
            //TIPO_CAD (A=AMBOS, C=CLIENTE, F=FORNECEDOR)
            else if (LowerCase(StringGrid1.Cells[i,0])='tipocad') then
            begin
              temp := StringGrid1.Cells[i,k];
              if ((temp='S') or (temp='1')) then temp:='C';
              if ((temp='N') or (temp='2')) then temp:='F';
              colClieForn := colClieForn + ',tipocad';
              dadosClieForn := dadosClieForn + ',''' + temp + '''';
            end
            //ATIVO
            else if (LowerCase(StringGrid1.Cells[i,0])='ativo') then
            begin
              temp := UpperCase(Trim(StringGrid1.Cells[i,k]));
              if ((temp='S') or (temp='1') or (temp='A')) then
              begin
                temp := 'S';
                colClieForn := colClieForn + ',ativo';
                dadosClieForn := dadosClieForn + ',''' + temp + '''';
              end
              else if ((temp='N') or (temp='0') or (temp='I')) then
              begin
                temp := 'N';
                colClieForn := colClieForn + ',ativo';
                dadosClieForn := dadosClieForn + ',''' + temp + '''';
              end;
            end
            ;

          //Fim do for das colunas
          end;

          //----------------------------------------
          //Gravar no banco de dados ClieForn
          try
            try
              //Abrir conexoes
              Connect.Open;
              SQL := TSQLDataSet.Create(Application);
              SQL.SQLConnection := Connect;

              Form2.atualizaStatus('Inserindo dados na tabela CLIEFORN.');

              //Desativar Trigger das cidades
              SQL.CommandText := 'ALTER TRIGGER clieforn_biu0 INACTIVE;';
              SQL.ExecSQL;
              //Executar INSERT
              SQL.CommandText := 'insert into clieforn ('+ colClieForn +') values ' + '(' + dadosClieForn + ');';
              SQL.ExecSQL;
              //Reativar Trigger das cidades
              SQL.CommandText := 'ALTER TRIGGER clieforn_biu0 ACTIVE;';
              SQL.ExecSQL;

            except
              on e: exception do
              begin
                ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
                status := 0;
                SQL.Free;
                Connect.Close;
                break; //Quebra o for
              end;
            end;

          finally
            SQL.Free;
            Connect.Close;
          end;

        end
        //----------------------------------------------------------------------------
        //Importar Produtos
        else if SelectImport.Text = 'Produtos' then
        begin

          Form2.atualizaStatus('Produto '+IntToStr(k));

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

          //Carregar informa��es para importar
          //-------------------------------------------------------

          //Empresa � obrigat�rio, se n�o tiver preenche com 1
          //EMPR (EMPRESA)
          i:=BuscaColuna(StringGrid1,'empr');
          if (i<>-1) then
          begin
            colProd := colProd + 'empr';
            dadosProd := dadosProd + '''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            colProdTrib := colProdTrib + 'trib_prod_empr';
            dadosProdTrib := dadosProdTrib + '''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            colProdTrib := colProdTrib + ',trib_empr';
            dadosProdTrib := dadosProdTrib + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            colProdAdic := colProdAdic + 'adic_prod_empr';
            dadosProdAdic := dadosProdAdic + '''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            colProdAdic := colProdAdic + ',adic_empr';
            dadosProdAdic := dadosProdAdic + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            colProdCust := colProdCust + 'cust_prod_empr';
            dadosProdCust := dadosProdCust + '''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            colProdCust := colProdCust + ',cust_empr';
            dadosProdCust := dadosProdCust + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            colMVA := colMVA + 'empr';
            dadosMVA := dadosMVA + '''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            colMVA := colMVA + ',mva_empr';
            dadosMVA := dadosMVA + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            colItens := colItens + 'empr';
            dadosItens := dadosItens + '''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            colProdForn := colProdForn + 'empr';
            dadosProdForn := dadosProdForn + '''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
          end
          else begin
            colProd := colProd + 'empr';
            dadosProd := dadosProd + '''' + '1' + '''';
            colProdTrib := colProdTrib + 'trib_prod_empr';
            dadosProdTrib := dadosProdTrib + '''' + '1' + '''';
            colProdTrib := colProdTrib + ',trib_empr';
            dadosProdTrib := dadosProdTrib + ',''' + '1' + '''';
            colProdAdic := colProdAdic + 'adic_prod_empr';
            dadosProdAdic := dadosProdAdic + '''' + '1' + '''';
            colProdAdic := colProdAdic + ',adic_empr';
            dadosProdAdic := dadosProdAdic + ',''' + '1' + '''';
            colProdCust := colProdCust + 'cust_prod_empr';
            dadosProdCust := dadosProdCust + '''' + '1' + '''';
            colProdCust := colProdCust + ',cust_empr';
            dadosProdCust := dadosProdCust + ',''' + '1' + '''';
            colMVA := colMVA + 'empr';
            dadosMVA := dadosMVA + '''' + '1' + '''';
            colMVA := colMVA + ',mva_empr';
            dadosMVA := dadosMVA + ',''' + '1' + '''';
            colItens := colItens + 'empr';
            dadosItens := dadosItens + '''' + '1' + '''';
            colProdForn := colProdForn + 'empr';
            dadosProdForn := dadosProdForn + '''' + '1' + '''';
          end;

          //Codigo � obrigat�rio, se n�o tiver preenche com o generator
          //CODI (CODIGO)
          i:=BuscaColuna(StringGrid1,'Codi');
          if (i<>-1) then
          begin
            if StrToInt(StringGrid1.Cells[i,k])>max then max:=StrToInt(StringGrid1.Cells[i,k]);
            colProd := colProd + ',codi';
            dadosProd := dadosProd + ',''' + StringGrid1.Cells[i,k] + '''';
            colProdTrib := colProdTrib + ',trib_id';
            dadosProdTrib := dadosProdTrib + ',' + 'gen_id(gen_prod_tributos_id,1)';
            colProdTrib := colProdTrib + ',trib_prod_codi';
            dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
            colProdAdic := colProdAdic + ',adic_id';
            dadosProdAdic := dadosProdAdic + ',' + 'gen_id(gen_prod_adicionais_id,1)';
            colProdAdic := colProdAdic + ',adic_prod_codi';
            dadosProdAdic := dadosProdAdic + ',''' + StringGrid1.Cells[i,k] + '''';
            colProdCust := colProdCust + ',cust_id';
            dadosProdCust := dadosProdCust + ',' + 'gen_id(gen_prod_custos_id,1)';
            colProdCust := colProdCust + ',cust_prod_codi';
            dadosProdCust := dadosProdCust + ',''' + StringGrid1.Cells[i,k] + '''';
            colMVA := colMVA + ',id';
            dadosMVA := dadosMVA + ',' + 'gen_id(gen_mva_id,1)';
            colMVA := colMVA + ',codi_prod';
            dadosMVA := dadosMVA + ',''' + StringGrid1.Cells[i,k] + '''';
            colItens := colItens + ',codi';
            dadosItens := dadosItens + ',' + 'gen_id(gen_itens_id,1)';
            colItens := colItens + ',prodcod';
            dadosItens := dadosItens + ',''' + StringGrid1.Cells[i,k] + '''';
            colProdForn := colProdForn + ',prod';
            dadosProdForn := dadosProdForn + ',''' + StringGrid1.Cells[i,k] + '''';
            colProdForn := colProdForn + ',id';
            dadosProdForn := dadosProdForn + ',' + 'gen_id(gen_prod_forn_id,1)';
          end
          else begin
            colProd := colProd + ',codi';
            dadosProd := dadosProd + ',' + 'gen_id(gen_prod_id,1)';
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
            dadosItens := dadosItens + ',' + 'gen_id(gen_itens_id,1)';
            colItens := colItens + ',prodcod';
            dadosItens := dadosItens + ',' + 'gen_id(gen_prod_id,0)';
            colProdForn := colProdForn + ',prod';
            dadosProdForn := dadosProdForn + ',' + 'gen_id(gen_prod_id,0)';
            colProdForn := colProdForn + ',id';
            dadosProdForn := dadosProdForn + ',' + 'gen_id(gen_prod_forn_id,1)';
          end;

          //Grupo, subgrupo, departamento, marca e tipo s�o obrigat�rios, se n�o tiver colocar padroes
          //GRUP
          i:=BuscaColuna(StringGrid1,'grup');
          if (i<>-1) then
          begin
            if (StringGrid1.Cells[i,k]<>'') then
            begin
              colProd := colProd + ',grup';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            end
            else begin
              colProd := colProd + ',grup';
              dadosProd := dadosProd + ',' + 'gen_id(gen_grup_prod_id,0)';
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
            if (StringGrid1.Cells[i,k]<>'') then
            begin
              colProd := colProd + ',sub_grup';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            end
            else begin
              colProd := colProd + ',sub_grup';
              dadosProd := dadosProd + ',' + 'gen_id(gen_sub_grup_prod_id,0)';
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
            if (StringGrid1.Cells[i,k]<>'')then
            begin
              colProd := colProd + ',codi_departamento';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
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
            if (StringGrid1.Cells[i,k]<>'') then
            begin
              colProd := colProd + ',marca';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            end
            else begin
              colProd := colProd + ',marca';
              dadosProd := dadosProd + ',' + 'gen_id(gen_marca_id,0)';
            end;
          end
          else begin
            colProd := colProd + ',marca';
            dadosProd := dadosProd + ',''' + '0' + '''';
          end;
          //TIPO
          i:=BuscaColuna(StringGrid1,'tipo');
          if (i<>-1) then
          begin
            colProd := colProd + ',codi_tipo';
            dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
          end
          else begin
            colProd := colProd + ',codi_tipo';
            dadosProd := dadosProd + ',''' + '0' + '''';
          end;

          //PS (Produto ou servi�o) Padr�o deixar 'P' pois sempre importamos produtos
          colProd := colProd + ',ps';
          dadosProd := dadosProd + ',''' + 'P' + '''';

          //Quantidade � obrigat�rio, se n�o tiver p�e 0
          //QTD
          i:=BuscaColuna(StringGrid1,'qtd');
          if (i<>-1) then
          begin
            colItens := colItens + ',qtd';
            dadosItens := dadosItens + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
          end
          else begin
            colItens := colItens + ',qtd';
            dadosItens := dadosItens + ',''' + '0' + '''';
          end;
          //Campos adicionais para a itens
          colItens := colItens + ',tipo';
          dadosItens := dadosItens + ',''' + '6' + '''';
          colItens := colItens + ',epv';
          dadosItens := dadosItens + ',''' + 'A' + '''';
          colItens := colItens + ',nume';
          dadosItens := dadosItens + ',' + 'gen_id(gen_prod_ajus_id,0)';

          //Fornecedor � obrigat�rio, se n�o tiver p�e 1
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
            //DESCR
            if (LowerCase(StringGrid1.Cells[i,0])='descr') then
            begin
              colProd := colProd + ',descr';
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,120));
              dadosProd := dadosProd + ',''' + temp + '''';
            end
            //DESCR2 (DESCRI��O COMPLEMENTAR)
            else if (LowerCase(StringGrid1.Cells[i,0])='descr2') then
            begin
              colProd := colProd + ',descr2';
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,255));
              dadosProd := dadosProd + ',''' + temp + '''';
            end
            //REFE (Referencia)
            else if (LowerCase(StringGrid1.Cells[i,0])='refe') then
            begin
              colProd := colProd + ',refe';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            end
            //REFE_ORIGINAL (Referencia Original)
            else if (LowerCase(StringGrid1.Cells[i,0])='refe_original') then
            begin
              colProd := colProd + ',refe_original';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            end
            //CODI_BARRA (Codigo de barras unitario)
            else if (LowerCase(StringGrid1.Cells[i,0])='codi_barra') then
            begin
              colProd := colProd + ',codi_barra';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            end
            //CODI_BARRA_COM (Codigo de barras embalagem)
            else if (LowerCase(StringGrid1.Cells[i,0])='codi_barra_com') then
            begin
              colProd := colProd + ',codi_barra_com';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            end
            //NCM
            else if (LowerCase(StringGrid1.Cells[i,0])='ncm') then
            begin
              colProd := colProd + ',ncm';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            end
            //CEST
            else if (LowerCase(StringGrid1.Cells[i,0])='cest') then
            begin
              colProd := colProd + ',cest';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
            end
            //UNID (Unidade de medida)
            else if (LowerCase(StringGrid1.Cells[i,0])='unid') then
            begin
              colProd := colProd + ',unid';
              dadosProd := dadosProd + ',''' + UpperCase(RemoveAcento(StringGrid1.Cells[i,k])) + '''';
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
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);
              dadosProdCust := dadosProdCust + ',' + temp;
              //Testar se existe o custo_real, se n�o joga o custo mesmo
              if (BuscaColuna(StringGrid1,'custo_real')=-1) then
              begin
                colProdCust := colProdCust + ',cust_custo_real';
                dadosProdCust := dadosProdCust + ',' + temp;
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
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);
              dadosProdCust := dadosProdCust + ',' + temp;
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
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);
              dadosProdCust := dadosProdCust + ',' + temp;
            end
            //PRECO_PRAZO (Pre�o a Prazo)
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
              temp := stringreplace(temp, '.', '',[rfReplaceAll, rfIgnoreCase]);
              temp := stringreplace(temp, ',', '.',[rfReplaceAll, rfIgnoreCase]);
              dadosProdCust := dadosProdCust + ',' + temp;
              //Testar se existe o preco a vista, se n�o joga o a prazo mesmo
              if (BuscaColuna(StringGrid1,'preco_vista')=-1) then
              begin
                colProdCust := colProdCust + ',cust_preco_vista';
                dadosProdCust := dadosProdCust + ',' + temp;
              end;
            end
            //CSOSN
            else if (LowerCase(StringGrid1.Cells[i,0])='csosn') then
            begin
              colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_ESTADUAL';
              dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
              colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_INTERESTADUAL';
              dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
              colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_ESTA_CF';
              dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
              colProdTrib := colProdTrib + ',TRIB_SN_CSOSN_INTER_CF';
              dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
            end
            //CST
            else if (LowerCase(StringGrid1.Cells[i,0])='cst') then
            begin
              colProdTrib := colProdTrib + ',TRIB_CST_ICMS_ESTADUAL';
              dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
              colProdTrib := colProdTrib + ',TRIB_CST_ICMS_INTERESTADUAL';
              dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
              colProdTrib := colProdTrib + ',TRIB_CST_ICMS_ESTA_CF';
              dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
              colProdTrib := colProdTrib + ',TRIB_CST_ICMS_INTER_CF';
              dadosProdTrib := dadosProdTrib + ',''' + StringGrid1.Cells[i,k] + '''';
            end
            ;

          //Fim do for
          end;

          //----------------------------------------
          //Gravar no banco de dados
          try
            try
              //Abrir conexoes
              Connect.Open;
              SQL := TSQLDataSet.Create(Application);
              SQL.SQLConnection := Connect;

              //Executar INSERTs
              Form2.atualizaStatus('Inserindo dados na tabela PROD.');
              SQL.CommandText := 'insert into prod ('+ colProd +') values ' + '(' + dadosProd + ');';
              SQL.ExecSQL;
              Form2.atualizaStatus('Inserindo dados na tabela PROD_TRIBUTOS.');
              SQL.CommandText := 'insert into prod_tributos ('+ colProdTrib +') values ' + '(' + dadosProdTrib + ');';
              SQL.ExecSQL;
              Form2.atualizaStatus('Inserindo dados na tabela PROD_ADICIONAIS.');
              SQL.CommandText := 'insert into prod_adicionais ('+ colProdAdic +') values ' + '(' + dadosProdAdic + ');';
              SQL.ExecSQL;
              Form2.atualizaStatus('Inserindo dados na tabela PROD_CUSTOS.');
              SQL.CommandText := 'insert into prod_custos ('+ colProdCust +') values ' + '(' + dadosProdCust + ');';
              SQL.ExecSQL;
              Form2.atualizaStatus('Inserindo dados na tabela MVA.');
              SQL.CommandText := 'insert into mva ('+ colMVA +') values ' + '(' + dadosMVA + ');';
              SQL.ExecSQL;
              Form2.atualizaStatus('Inserindo dados na tabela PROD_AJUS.');
              SQL.CommandText := 'insert into prod_ajus (codi,data) values (gen_id(gen_prod_ajus_id,1),CURRENT_DATE);';
              SQL.ExecSQL;
              Form2.atualizaStatus('Inserindo dados na tabela ITENS.');
              SQL.CommandText := 'insert into itens ('+ colItens +') values ' + '(' + dadosItens + ');';
              SQL.ExecSQL;
              Form2.atualizaStatus('Inserindo dados na tabela PROD_FORN.');
              SQL.CommandText := 'insert into prod_forn ('+ colProdForn +') values ' + '(' + dadosProdForn + ');';
              SQL.ExecSQL;

            except
              on e: exception do
              begin
                ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
                status := 0;
                SQL.Free;
                Connect.Close;
                break; //Quebra o for
              end;
            end;

          finally
            SQL.Free;
            Connect.Close;
          end;

        end

        //----------------------------------------------------------------------------
        //Importar Grupos
        else if SelectImport.Text = 'Grupos' then
        begin
          //ShowMessage('Importar Grupos');

          Form2.atualizaStatus('Grupo '+IntToStr(k));

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
            if StrToInt(StringGrid1.Cells[i,k])>max then max:=StrToInt(StringGrid1.Cells[i,k]);
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

          //----------------------------------------
          //Gravar no banco de GRUP_PROD
          try
            try
              //Abrir conexoes
              Connect.Open;
              SQL := TSQLDataSet.Create(Application);
              SQL.SQLConnection := Connect;

              //Executar INSERT
              Form2.atualizaStatus('Inserindo dados na tabela GRUP_PROD.');
              SQL.CommandText := 'insert into grup_prod ('+ colGrupo +') values ' + '(' + dadosGrupo + ');';
              SQL.ExecSQL;

            except
              on e: exception do
              begin
                ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
                status := 0;
                SQL.Free;
                Connect.Close;
                break; //Quebra o for
              end;
            end;

          finally
            SQL.Free;
            Connect.Close;
          end;

        end

        //----------------------------------------------------------------------------
        //Importar SubGrupos
        else if SelectImport.Text = 'SubGrupos' then
        begin
          //ShowMessage('Importar SubGrupos');

          Form2.atualizaStatus('SubGrupo '+IntToStr(k));

          colSubGrupo := '';
          dadosSubGrupo := '';

          //Carregar informa��es para importar
          //-------------------------------------------------------

          //Codigo � obrigat�rio, se n�o tiver preenche com o generator
          //CODI (CODIGO)
          i:=BuscaColuna(StringGrid1,'codi');
          if (i<>-1) then
          begin
            colSubGrupo := colSubGrupo + 'codi';
            if StrToInt(StringGrid1.Cells[i,k])>max then max:=StrToInt(StringGrid1.Cells[i,k]);
            dadosSubGrupo := dadosSubGrupo + '''' + StringGrid1.Cells[i,k] + '''';
          end
          else begin
            colSubGrupo := colSubGrupo + 'codi';
            dadosSubGrupo := dadosSubGrupo + 'gen_id(gen_sub_grup_prod_id,1)';
          end;

          //Empresa � obrigat�rio, se n�o tiver preenche com 1
          //EMPR (EMPRESA)
          i:=BuscaColuna(StringGrid1,'empr');
          if (i<>-1) then
          begin
            colSubGrupo := colSubGrupo + ',empr';
            dadosSubGrupo := dadosSubGrupo + ',''' + StringGrid1.Cells[i,k] + '''';
          end
          else begin
            colSubGrupo := colSubGrupo + ',empr';
            dadosSubGrupo := dadosSubGrupo + ',''' + '1' + '''';
          end;


          for i := 0 to StringGrid1.ColCount-1 do
          begin
            //DESCR (DESCRICAO)
            if (LowerCase(StringGrid1.Cells[i,0])='descr') then
            begin
              colSubGrupo := colSubGrupo + ',descr';
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,30));
              dadosSubGrupo := dadosSubGrupo + ',''' + temp + '''';
            end
            ;

          //Fim do For das colunas
          end;

          //----------------------------------------
          //Gravar no banco de SUB_GRUP_PROD
          try
            try
              //Abrir conexoes
              Connect.Open;
              SQL := TSQLDataSet.Create(Application);
              SQL.SQLConnection := Connect;

              //Executar INSERT
              Form2.atualizaStatus('Inserindo dados na tabela SUB_GRUP_PROD.');
              SQL.CommandText := 'insert into sub_grup_prod ('+ colSubGrupo +') values ' + '(' + dadosSubGrupo + ');';
              SQL.ExecSQL;

            except
              on e: exception do
              begin
                ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
                status := 0;
                SQL.Free;
                Connect.Close;
                break; //Quebra o for
              end;
            end;

          finally
            SQL.Free;
            Connect.Close;
          end;

        end

        //----------------------------------------------------------------------------
        //Importar Marcas
        else if SelectImport.Text = 'Marcas' then
        begin
          //ShowMessage('Importar Marcas');

          Form2.atualizaStatus('Marca '+IntToStr(k));

          colMarca := '';
          dadosMarca := '';

          //Carregar informa��es para importar
          //-------------------------------------------------------

          //Codigo � obrigat�rio, se n�o tiver preenche com o generator
          //CODI (CODIGO)
          i:=BuscaColuna(StringGrid1,'codi');
          if (i<>-1) then
          begin
            colMarca := colMarca + 'codi';
            if StrToInt(StringGrid1.Cells[i,k])>max then max:=StrToInt(StringGrid1.Cells[i,k]);
            dadosMarca := dadosMarca + '''' + StringGrid1.Cells[i,k] + '''';
          end
          else begin
            colMarca := colMarca + 'codi';
            dadosMarca := dadosMarca + 'gen_id(gen_marca_id,1)';
          end;


          for i := 0 to StringGrid1.ColCount-1 do
          begin
            //DESCR (DESCRICAO)
            if (LowerCase(StringGrid1.Cells[i,0])='descr') then
            begin
              colMarca := colMarca + ',descr';
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,30));
              dadosMarca := dadosMarca + ',''' + temp + '''';
            end
            ;

          //Fim do For das colunas
          end;

          //----------------------------------------
          //Gravar no banco de MARCA
          try
            try
              //Abrir conexoes
              Connect.Open;
              SQL := TSQLDataSet.Create(Application);
              SQL.SQLConnection := Connect;

              //Executar INSERT
              Form2.atualizaStatus('Inserindo dados na tabela MARCA.');
              SQL.CommandText := 'insert into marca ('+ colMarca +') values ' + '(' + dadosMarca + ');';
              SQL.ExecSQL;

            except
              on e: exception do
              begin
                ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
                status := 0;
                SQL.Free;
                Connect.Close;
                break; //Quebra o for
              end;
            end;

          finally
            SQL.Free;
            Connect.Close;
          end;

        end

        //----------------------------------------------------------------------------
        //Importar T�tulos a Pagar
        else if SelectImport.Text = 'T�tulos a Pagar' then
        begin
          //ShowMessage('Importar T�tulos a Pagar');

          Form2.atualizaStatus('T�tulos a Pagar '+IntToStr(k));

          colTituP := '';
          dadosTituP := '';

          //Carregar informa��es para importar
          //-------------------------------------------------------

          //CODI (CODIGO)
          i:=BuscaColuna(StringGrid1,'codi');
          if (i<>-1) then
          begin
            //Testar se ja existir o c�digo do t�tulo e inserir uma barra.
            count := 0;
            while (temCodTituloR(StringGrid1.Cells[i,k]) = True) do
            begin
              count := count+1;
              StringGrid1.Cells[i,k] := StringGrid1.Cells[i,k] + '/' + IntToStr(count);
            end;

            colTituP := colTituP + 'codi';
            dadosTituP := dadosTituP + '''' + StringGrid1.Cells[i,k] + '''';
          end
          else begin
            colTituP := colTituP + 'codi';
            dadosTituP := dadosTituP + IntToStr(k);
          end;


          for i := 0 to StringGrid1.ColCount-1 do
          begin
            //DESCR (DESCRICAO)
            if (LowerCase(StringGrid1.Cells[i,0])='descr') then
            begin
              colMarca := colMarca + ',descr';
              temp := UpperCase(RemoveAcento(StringGrid1.Cells[i,k]));
              temp := stringreplace(temp, '''', ' ',[rfReplaceAll, rfIgnoreCase]);
              temp := (Copy(temp,1,30));
              dadosMarca := dadosMarca + ',''' + temp + '''';
            end
            ;

          //Fim do For das colunas
          end;

          //----------------------------------------
          //Gravar no banco de T�tulos a Pagar
          try
            try
              //Abrir conexoes
              Connect.Open;
              SQL := TSQLDataSet.Create(Application);
              SQL.SQLConnection := Connect;

              //Executar INSERT
              Form2.atualizaStatus('Inserindo dados na tabela MARCA.');
              SQL.CommandText := 'insert into marca ('+ colMarca +') values ' + '(' + dadosMarca + ');';
              SQL.ExecSQL;

            except
              on e: exception do
              begin
                ShowMessage('Erro SQL: '+e.message+sLineBreak+SQL.CommandText);
                status := 0;
                SQL.Free;
                Connect.Close;
                break; //Quebra o for
              end;
            end;

          finally
            SQL.Free;
            Connect.Close;
          end;

        end

        //----------------------------------------------------------------------------
        //OUTRAS OP��ES DE IMPORTA��O COLOCAR AQUI

        ;
      //Fim do For das Linhas
      end;

      //------------------------------------------------------------------------------
      //COMANDOS P�S IMPORTA��O

      //Abrir conexoes
      Connect.Open;
      SQL := TSQLDataSet.Create(Application);
      SQL.SQLConnection := Connect;

      if SelectImport.Text='Clie/Forn' then
      begin
        //Arrumar Generator dos Clientes e Fornecedores
        if max > 0 then
        begin
          Form2.atualizaStatus('Alterando generator do Clie/Forn.');
          SQL.CommandText := 'ALTER SEQUENCE GEN_CLIEFORN_ID RESTART WITH ' + IntToStr(max) + ';';
          SQL.ExecSQL;
        end;
      end

      else if SelectImport.Text='Produtos' then
      begin
        //Arrumar Generator dos Produtos
        if max > 0 then
        begin
          Form2.atualizaStatus('Alterar Generator do Produto.');
          SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_ID RESTART WITH ' + IntToStr(max) + ';';
          SQL.ExecSQL;
        end;

        //Atualizar MARGEM1
        SQL.CommandText := 'update prod_custos pc set pc.cust_margem1= abs(pc.cust_preco_prazo - pc.cust_custo_real)/pc.cust_custo_real where pc.cust_custo_real>0;';
        SQL.ExecSQL;
        SQL.CommandText := 'update prod_custos pc set pc.cust_margem1 = pc.cust_margem1 * 100;';
        SQL.ExecSQL;
        //Atualizar MARGEM2
        SQL.CommandText := 'update prod_custos pc set pc.cust_margem2 = (cast(pc.cust_preco_vista as numeric (18,2))/cast(pc.cust_preco_prazo as numeric (18,2)) -1)*100 where cast(pc.cust_preco_prazo as numeric (18,2))>0;';
        SQL.ExecSQL;
      end

      else if SelectImport.Text='Grupos' then
      begin
        //Arrumar Generator dos Grupos
        if max > 0 then
        begin
          Form2.atualizaStatus('Alterar Generator dos Grupos.');
          SQL.CommandText := 'ALTER SEQUENCE GEN_GRUP_PROD_ID RESTART WITH ' + IntToStr(max) + ';';
          SQL.ExecSQL;
        end;
      end

      else if SelectImport.Text='SubGrupos' then
      begin
        //Arrumar Generator dos SubGrupos
        if max > 0 then
        begin
          Form2.atualizaStatus('Alterar Generator dos SubGrupos.');
          SQL.CommandText := 'ALTER SEQUENCE GEN_SUB_GRUP_PROD_ID RESTART WITH ' + IntToStr(max) + ';';
          SQL.ExecSQL;
        end;
      end

      else if SelectImport.Text='Marcas' then
      begin
        //Arrumar Generator das MARCAS
        if max > 0 then
        begin
          Form2.atualizaStatus('Alterar Generator das Marcas.');
          SQL.CommandText := 'ALTER SEQUENCE GEN_MARCA_ID RESTART WITH ' + IntToStr(max) + ';';
          SQL.ExecSQL;
        end;
      end

      ;

      //Fechar conexoes
      SQL.Free;
      Connect.Close;


    except
      status := 0;
    end;
  finally
    Form2.fim(status);

  end;

end;


//Fun��o para salvar StringGrid em um arquivo Excel
function SaveAsExcelFile(stringGrid: TstringGrid; FileName: string): Boolean;
const
  xlWBATWorksheet = -4167;
var
  Row, Col: Integer;
  GridPrevFile: string;
  XLApp, Sheet: OLEVariant;

begin
  Result := False;
  XLApp := CreateOleObject('Excel.Application');
  try
    XLApp.Visible := False;
    XLApp.Workbooks.Add(xlWBatWorkSheet);
    Sheet := XLApp.Workbooks[1].WorkSheets[1];

    Sheet.Name := 'String Grid';

    for col := 0 to stringGrid.ColCount - 1 do
      for row := 0 to stringGrid.RowCount - 1 do
        Sheet.Cells[row + 1, col + 1] := stringGrid.Cells[col, row];
    try
      XLApp.Workbooks[1].SaveAs(FileName);
      Result := True;
    except
    // Error ?
    end;
  finally
    if not VarIsEmpty(XLApp) then
    begin
      XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLAPP := Unassigned;
      Sheet := Unassigned;
    end;
  end;
end;


//Salvar StringGrid em planilha
procedure TForm1.ButSaveClick(Sender: TObject);
var
  arquivo: string;

begin
  SaveDialog1.Filter := 'EXCEL files (*.xlsx)|*.XLSX|CSV files (*.csv)|*.CSV';
  if SaveDialog1.Execute then
  begin
    arquivo := ExtractFilePath(Application.ExeName);
    if SaveAsExcelFile(StringGrid1, SaveDialog1.FileName) then
      ShowMessage(SaveDialog1.FileName+sLineBreak+'StringGrid salva com sucesso!');
  end;
end;


//Fun��o para salvar StringGrid em um array de string
function StringGridToArray(Grid: TStringGrid): Integer;
var
  i,j: integer;
begin
  //Redimensionar array
  SetLength(gridTemp,Grid.RowCount);
  for i := 0 to Grid.RowCount-1 do
  begin
    SetLength(gridTemp[i],Grid.ColCount);
  end;

  //Copiar da StringGrid para o array
  for i := 0 to Grid.RowCount-1 do
  begin
    for j := 0 to Grid.ColCount-1 do
    begin
      gridTemp[i,j] := Grid.Cells[j,i];
    end;
  end;

end;



//Carregar Cabe�alho de outra tabela nesta tabela
procedure TForm1.Cabealho1Click(Sender: TObject);
var
  arquivo: string;

begin
  if OpenDialog1.Execute then
  begin
    ExtractFilePath(Application.ExeName);
    arquivo :=  OpenDialog1.FileName;

    StringGridToArray(StringGrid1);
    XlsHeaderLoad(StringGrid1,arquivo);

  end;
end;


//Limpar dados de clientes e fornecedores do banco
procedure TForm1.LimpaClieFornClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  //Abrir conexoes
  Connect.Open;
  SQL := TSQLDataSet.Create(Application);
  SQL.SQLConnection := Connect;

  SQL.CommandText := 'delete from clieforn;';
  SQL.ExecSQL;

  SQL.CommandText := 'ALTER SEQUENCE GEN_CLIEFORN_ID RESTART WITH 0;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de clientes e fornecedores.');

  //Fechar conexoes
  SQL.Free;
  Connect.Close;
end;


//Limpar dados de grupos do banco
procedure TForm1.LimpaGruposClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  //Abrir conexoes
  Connect.Open;
  SQL := TSQLDataSet.Create(Application);
  SQL.SQLConnection := Connect;

  SQL.CommandText := 'delete from grup_prod;';
  SQL.ExecSQL;

  SQL.CommandText := 'ALTER SEQUENCE GEN_grup_prod_ID RESTART WITH 0;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de grupos.');

  //Fechar conexoes
  SQL.Free;
  Connect.Close;
end;


//Limpar dados de marcas do banco
procedure TForm1.LimpaMarcasClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  //Abrir conexoes
  Connect.Open;
  SQL := TSQLDataSet.Create(Application);
  SQL.SQLConnection := Connect;

  SQL.CommandText := 'delete from marca;';
  SQL.ExecSQL;

  SQL.CommandText := 'ALTER SEQUENCE GEN_MARCA_ID RESTART WITH 0;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de marcas.');

  //Fechar conexoes
  SQL.Free;
  Connect.Close;
end;


//Limpar dados de produtos do banco
procedure TForm1.LimpaProdutosClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  //Abrir conexoes
  Connect.Open;
  SQL := TSQLDataSet.Create(Application);
  SQL.SQLConnection := Connect;

  SQL.CommandText := 'delete from prod;';
  SQL.ExecSQL;
  SQL.CommandText := 'delete from prod_ajus;';
  SQL.ExecSQL;
  SQL.CommandText := 'delete from itens;';
  SQL.ExecSQL;

  SQL.CommandText := 'ALTER SEQUENCE GEN_ITENS_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_MVA_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_ADICIONAIS_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_AJUS_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_CUSTOS_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_FORN_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_ICMS_ST_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_ID RESTART WITH 0;';
  SQL.ExecSQL;
  SQL.CommandText := 'ALTER SEQUENCE GEN_PROD_TRIBUTOS_ID RESTART WITH 0;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de produtos.');

  //Fechar conexoes
  SQL.Free;
  Connect.Close;
end;


//Limpar dados de subgrupos do banco
procedure TForm1.LimpaSubGruposClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  //Abrir conexoes
  Connect.Open;
  SQL := TSQLDataSet.Create(Application);
  SQL.SQLConnection := Connect;

  SQL.CommandText := 'delete from sub_grup_prod;';
  SQL.ExecSQL;

  SQL.CommandText := 'ALTER SEQUENCE GEN_sub_grup_prod_ID RESTART WITH 0;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de subgrupos.');

  //Fechar conexoes
  SQL.Free;
  Connect.Close;
end;


//Limpar dados de Titulos a pagar do banco
procedure TForm1.LimpaTituPClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  //Abrir conexoes
  Connect.Open;
  SQL := TSQLDataSet.Create(Application);
  SQL.SQLConnection := Connect;

  SQL.CommandText := 'delete from titup;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de T�tulos a Pagar.');

  //Fechar conexoes
  SQL.Free;
  Connect.Close;
end;


//Limpar dados de Titulos a receber do banco
procedure TForm1.LimpaTituRClick(Sender: TObject);
var
  SQL: TSQLDataSet;
begin
  //Abrir conexoes
  Connect.Open;
  SQL := TSQLDataSet.Create(Application);
  SQL.SQLConnection := Connect;

  SQL.CommandText := 'delete from titur;';
  SQL.ExecSQL;

  ShowMessage('Limpado dados de T�tulos a Receber.');

  //Fechar conexoes
  SQL.Free;
  Connect.Close;
end;


//Deletar Linha da StringGrid
procedure DeleteRow(Grid: TStringGrid; ARow: Integer);
var
  i: Integer;
begin
  for i := ARow to Grid.RowCount - 2 do
    Grid.Rows[i].Assign(Grid.Rows[i + 1]);
  Grid.RowCount := Grid.RowCount - 1;
end;


//Deletar coluna na StringGrid
procedure DeleteCol(Grid: TStringGrid; ACol: Integer);
var
  i: Integer;
begin
  for i := ACol to Grid.ColCount - 2 do
    Grid.Cols[i].Assign(Grid.Cols[i + 1]);
  Grid.ColCount := Grid.ColCount - 1;
end;


//Evento ao apertar Bot�o Delete na StringGrid
procedure TForm1.StringGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  i,j,but: integer;
  temp: string;
begin

  //RECONHECER CTRL+Z
   if ((Shift = [ssCtrl]) and (Key = 90)) then
   begin
    if Length(gridTemp) > 1 then
    begin
      StringGrid1.RowCount := Length(gridTemp);
      StringGrid1.ColCount := Length(gridTemp[0]);

      for i := 0 to Length(gridTemp)-1 do
      begin
        for j := 0 to Length(gridTemp[0])-1 do
        begin
          StringGrid1.Cells[j,i] := gridTemp[i,j];
        end;
      end;
    end;
    //Redimensionar colunas
    for i := 0 to StringGrid1.ColCount - 1 do
      AutoSizeCol(StringGrid1, i);
   end;

  //Se apertar 'Del'
  if (Key = VK_DELETE) then
  begin
    but := Mensagem('Deletar linha ou coluna', mtCustom, [mbYes, mbNo],['Linha','Coluna'], 'Deletar');
    if (but = 6) then
    begin
      //ShowMessage('Deletar Linha');
      StringGridToArray(StringGrid1);
      DeleteRow(StringGrid1, StringGrid1.Row);
    end
    else if (but = 7) then
    begin
      //ShowMessage('Deletar Coluna');
      StringGridToArray(StringGrid1);
      DeleteCol(StringGrid1, StringGrid1.Col);
      //Redimensionar colunas
      for i := 0 to StringGrid1.ColCount - 1 do
        AutoSizeCol(StringGrid1, i);
    end;
  end;

  //Teclas para voltar as linhas fixas
  if ((Key=VK_TAB) or
   (Key=VK_RETURN) or
   (Key=VK_ESCAPE) or
   (Key=VK_UP) or
   (Key=VK_DOWN)) then
  begin
    StringGrid1.FixedRows:=1;
  end;

  //Teclas para voltar as colunas fixas
  if ((Key=VK_TAB) or
   (Key=VK_RETURN) or
   (Key=VK_ESCAPE) or
   (Key=VK_LEFT) or
   (Key=VK_RIGHT)) then
  begin
    StringGrid1.FixedCols:=1;
  end;

  //Inserir coluna
  if (Key=VK_F1) then
  begin
    StringGridToArray(StringGrid1);
    StringGrid1.ColCount := StringGrid1.ColCount + 1;
    i:= StringGrid1.ColCount;
    while i>StringGrid1.Col do
    begin
      for j := 0 to StringGrid1.RowCount do
      begin
        temp := StringGrid1.Cells[i,j];
        StringGrid1.Cells[i,j] := StringGrid1.Cells[i-1,j];
      end;
      i:= i-1;
    end;
    for j := 0 to StringGrid1.RowCount do
        StringGrid1.Cells[i,j] := '';

    //Redimensionar colunas
    for i := 0 to StringGrid1.ColCount - 1 do
      AutoSizeCol(StringGrid1, i);
  end;

  //Inserir Linha
  if (Key=VK_F3) then
  begin
    StringGridToArray(StringGrid1);
    StringGrid1.RowCount := StringGrid1.RowCount + 1;
    i:= StringGrid1.RowCount;
    while i>StringGrid1.Row do
    begin
      for j := 0 to StringGrid1.ColCount do
      begin
        temp := StringGrid1.Cells[j,i];
        StringGrid1.Cells[j,i] := StringGrid1.Cells[j,i-1];
      end;
      i:= i-1;
    end;
    for j := 0 to StringGrid1.ColCount do
        StringGrid1.Cells[j,i] := '';
  end;
end;


//Reconhecer Right Click na celula
procedure TForm1.StringGrid1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  PMouse: TPoint;
  i, j, Col, Row, but, but2: integer;
  valor,temp: string;
begin
  //Right Click
  if Button = mbRight then
  begin
    //Testar qual coluna clicou
    PMouse := Mouse.CursorPos;
    PMouse := StringGrid1.ScreenToClient(PMouse);
    StringGrid1.MouseToCell(PMouse.x, PMouse.y, Col, Row);

    //Se for uma coluna n�o fixa
    if (((Col<>0) and (Row<>0)) or ((Col=0) and (Row<>0))) then
    begin
      //Setar focus
      StringGrid1.Col := Col;
      StringGrid1.Row := Row;

      but := Mensagem('Preencher coluna', mtCustom, [mbYes, mbNo],['Copiar','Serie'], 'Preencher');
      if (but = 6) then
      begin
        //ShowMessage('Copiar valor');
        StringGridToArray(StringGrid1);
        valor := StringGrid1.Cells[Col,Row];
        for i := 1 to StringGrid1.RowCount do
        begin
          StringGrid1.Cells[Col,i] := valor;
        end;
      end
      else if (but = 7) then
      begin
        //ShowMessage('Preencher serie');
        StringGridToArray(StringGrid1);
        valor := StringGrid1.Cells[Col,Row];
        j:=1;
        for i := Row+1 to StringGrid1.RowCount do
        begin
          StringGrid1.Cells[Col,i] := IntToStr(StrToInt(valor)+j);
          j:=j+1;
        end;
      end;
    end
    //Se for uma fixa pergunta se deseja mesclar ou copiar coluna
    else if Row=0 then     
    begin
      but := Mensagem('Mesclar ou copiar coluna', mtCustom, [mbYes, mbNo],['Mesclar','Copiar'], 'Mesclar ou copiar coluna');
      if (but = 6) then
      begin
        //ShowMessage('Mesclar Coluna');
        but2 := Mensagem('Mesclar com a coluna', mtCustom, [mbYes, mbNo],['� Esquerda','� Direita'], 'Mesclar colunas');
        if (but2 = 6) then
        begin
          if Col=0 then
          begin
            ShowMessage('N�o existem mais colunas � esquerda.');
          end
          else begin
            StringGridToArray(StringGrid1);
            for i := 0 to StringGrid1.RowCount-1 do
            begin
              StringGrid1.Cells[Col-1,i] := StringGrid1.Cells[Col-1,i] + StringGrid1.Cells[Col,i];
            end;
            DeleteCol(StringGrid1, Col);
          end;
          //Redimensionar colunas
          for i := 0 to StringGrid1.ColCount - 1 do
            AutoSizeCol(StringGrid1, i);
        end
        else if (but2 = 7) then
        begin
          if Col=StringGrid1.ColCount then
          begin
            ShowMessage('N�o existem mais colunas � direita.');
          end
          else begin
            StringGridToArray(StringGrid1);
            for i := 0 to StringGrid1.RowCount-1 do
            begin
              StringGrid1.Cells[Col+1,i] := StringGrid1.Cells[Col,i] + StringGrid1.Cells[Col+1,i];
            end;
            DeleteCol(StringGrid1, Col);
          end;
        end;
        //Redimensionar colunas
        for i := 0 to StringGrid1.ColCount - 1 do
          AutoSizeCol(StringGrid1, i);
      end
      else if (but = 7) then
      begin
        //ShowMessage('Copiar Coluna');
        StringGridToArray(StringGrid1);
        StringGrid1.ColCount := StringGrid1.ColCount + 1;
        i:= StringGrid1.ColCount;
        while i>Col do
        begin
          for j := 0 to StringGrid1.RowCount do
          begin
            temp := StringGrid1.Cells[i,j];
            StringGrid1.Cells[i,j] := StringGrid1.Cells[i-1,j];
          end;
          i:= i-1;
        end;
      end;
      //Redimensionar colunas
      for i := 0 to StringGrid1.ColCount - 1 do
        AutoSizeCol(StringGrid1, i);
    end;
  end;
end;



procedure TForm1.StringGrid1Click(Sender: TObject);
  var
  PMouse: TPoint;
  Col, Row: integer;
begin

  //Voltar as celulas fixas ap�s clicar fora
  PMouse := Mouse.CursorPos;
  PMouse := StringGrid1.ScreenToClient(PMouse);

  StringGrid1.MouseToCell(PMouse.x, PMouse.y, Col, Row);

  if Row<>0 then
  begin
    StringGrid1.FixedRows:=1;
  end;
  if Col<>0 then
  begin
    StringGrid1.FixedCols:=1;
  end;

end;

procedure TForm1.StringGrid1DblClick(Sender: TObject);
var
  PMouse: TPoint;
  Col, Row: integer;
begin

  //Desabilitar celulas fixas ao dar dois cliques
  PMouse := Mouse.CursorPos;
  PMouse := StringGrid1.ScreenToClient(PMouse);

  StringGrid1.MouseToCell(PMouse.x, PMouse.y, Col, Row);

  if Row=0 then
  begin
    StringGrid1.FixedRows:=0;
    //Setar focus
    StringGrid1.Col := Col;
    StringGrid1.Row := Row;
  end;
  if Col=0 then
  begin
    StringGrid1.FixedCols:=0;
    //Setar focus
    StringGrid1.Col := Col;
    StringGrid1.Row := Row;
  end;

end;


{$R *.dfm}

initialization
SetLength(gridTemp,1);
SetLength(gridTemp[0],1);

end.
