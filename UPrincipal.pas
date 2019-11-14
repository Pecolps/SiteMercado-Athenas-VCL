unit UPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Imaging.pngimage, Vcl.ComCtrls, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.VCLUI.Wait,
  Data.DB, FireDAC.Comp.Client, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.Phys.FB,
  FireDAC.Phys.FBDef, IniFiles, System.Generics.Collections, Vcl.Menus;

type
  TfrmPrincipal = class(TForm)
    pnTop: TPanel;
    imgLogo: TImage;
    btnSair: TPanel;
    lblTitulo: TLabel;
    pnBot: TPanel;
    lbl_ver: TLabel;
    pgControl: TPageControl;
    tabPrincipal: TTabSheet;
    tabConfig: TTabSheet;
    memoPrincipal: TMemo;
    gbConexao: TGroupBox;
    gpCampos: TGroupBox;
    lblSobre: TLabel;
    gbConfig: TGroupBox;
    fdConn: TFDConnection;
    fdQuery: TFDQuery;
    edServer: TLabeledEdit;
    edCaminho: TLabeledEdit;
    btnTesteConexao: TButton;
    sbCampos: TScrollBox;
    edUser: TLabeledEdit;
    edPassword: TLabeledEdit;
    Label1: TLabel;
    btnSalvar: TButton;
    edIntervalo: TLabeledEdit;
    edCaminhoIntegracao: TLabeledEdit;
    tmrPrincipal: TTimer;
    Tray: TTrayIcon;
    popTray: TPopupMenu;
    ppbtnSair: TMenuItem;
    popbtnProcessarAgora: TMenuItem;
    fdQuery2: TFDQuery;
    procedure pnTopMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnSairClick(Sender: TObject);
    procedure btnSairMouseEnter(Sender: TObject);
    procedure btnSairMouseLeave(Sender: TObject);
    procedure btnSairMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnTesteConexaoClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TrayClick(Sender: TObject);
    procedure popbtnProcessarAgoraClick(Sender: TObject);
    procedure ppbtnSairClick(Sender: TObject);
    procedure btnSalvarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

const
  versaoAtual: Double = 0.1;                                                            //Constante de vers�o

var
  frmPrincipal: TfrmPrincipal;
  Campos: TList<String>;
  IniPath: string = ''; //Caminho do Ini

implementation

{$R *.dfm}


procedure Logg(Texto: string);
var
Arquivo: TStringList;
FileName: string;
begin
  //Grava o texto e exibe
  frmPrincipal.memoPrincipal.Lines.Add(Texto);
  try
    FileName:= Application.ExeName + '_log.txt';
    Arquivo:= TStringList.Create();
    if FileExists(FileName) then
      Arquivo.LoadFromFile(FileName);
    Arquivo.Add(FormatdateTime('[' + 'dd/mm/yyy hh:mm:ss',Now) + '] ' + Texto);
    Arquivo.SaveToFile(FileName);
  except
    on e:exception do
      begin
        //Caso n�o consiga gravar o arquivo.
        frmPrincipal.memoPrincipal.Lines.Add(e.Message);
      end;
  end;
end;

procedure TfrmPrincipal.btnSairClick(Sender: TObject);
begin
Application.Terminate;
end;

procedure TfrmPrincipal.btnSairMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
Btnsair.Color:= $00994C00;
end;

procedure TfrmPrincipal.btnSairMouseEnter(Sender: TObject);
begin
Btnsair.Color:= clHighlight;
end;

procedure TfrmPrincipal.btnSairMouseLeave(Sender: TObject);
begin
btnsair.Color:= clHotLight;
end;

procedure TfrmPrincipal.btnTesteConexaoClick(Sender: TObject);
begin
  //Testa a Conex�o
  fdConn.Connected:= false;
  fdconn.Params.Clear;
  fdConn.DriverName := 'FB';
  fdConn.Params.Add('User_name=' + edUser.Text);
  fdConn.Params.Add('Password=' + edPassword.Text);
  fdConn.Params.Add('Database=' + edcaminho.Text);
  fdConn.Params.Add('Server='+  edServer.Text);
  fdconn.Connected:= true;
  if fdconn.Connected then
    begin
      showmessage('Contado com sucesso!');
    end;
end;

procedure AssociaCampos;
var
  I: Integer;
begin
  with frmPrincipal do
  begin
   Campos:= TList<String>.Create(); //Cria a Lista na mem�ria
   for I := 0 to 22 do //Preenche com os campos do SiteMarcado
      begin
       Campos.Add('');
      end;
   Campos[0]:= '';                  //Id_Loja - Padr�o 001
   Campos[1]:= 'grupo';             //Departamento
   Campos[2]:= 'subgrupo';          //Categoria
   Campos[3]:= 'subgrupo';          //Subcategoria
   Campos[4]:= '';                  //Marca - N�o Implementado
   Campos[5]:= 'unidade';           //Unidade
   Campos[6]:= '';                  //Volume - N�o Implementado
   Campos[7]:= 'codbarra';          //Codigo_barra
   Campos[8]:= 'produto';           //Nome
   Campos[9]:= 'data_cadastro';     //dt_cadastro
   Campos[10]:= 'ultima_alteracao'; //dt_ultima_alteracao
   Campos[11]:= 'precovenda';       //vlr_produto
   Campos[12]:= '';                 //vlr_promocao - N�o Implementado
   Campos[13]:= 'estoque_atual';    //qtd_estoque_atual
   Campos[14]:= 'estoqueminimo';    //qtd_estoque_minimo
   Campos[15]:= 'produto';          //Descricao
   Campos[16]:= 'situacao';         //ativo
   Campos[17]:= 'codigo';           //plu
   Campos[18]:= 'precocusto';       //vlr_compra
   Campos[19]:= '';                 //validade_proxima - N�o Implementado
   Campos[20]:= '';                 //vlr_atacado - N�o Implementado
   Campos[21]:= '';                 //qtd_atacado - N�o Implementado
   Campos[22]:= '';                  //image_url - N�o Implementado
  end;
end;

procedure GravaConfiguracoes;
var
ArquivoIni: TIniFile;
begin
  //Salva as configura��es
  With frmPrincipal do //Entra no Escopo do form principal
    begin
      //Abre o ini
      ArquivoIni:= TiniFile.Create(IniPath);
      //Salva configura��es do sistema
      ArquivoIni.WriteString('Configuracoes','CaminhoSaida',  edCaminhoIntegracao.Text);    //Caminho de sa�da padr�o
      ArquivoIni.WriteInteger('Configuracoes','Intervalo',    strtoint(edIntervalo.Text));  //Inteiro com os segundos
      //Salva conex�o do banco
      ArquivoIni.WriteString('Configuracoes','username', eduser.Text);      //Padr�o Firebird
      ArquivoIni.WriteString('Configuracoes','password', edPassword.Text);  //Padr�o Firebird
      ArquivoIni.WriteString('Configuracoes','database', edCaminho.Text);   //Caminho do Banco
      ArquivoIni.WriteString('Configuracoes','server',   edServer.Text);    //Caminho do Servidor
      //Conex�o com o Banco
      Logg('Conectando com o banco...');
      fdConn.Connected:= false;
      fdconn.Params.Clear;
      fdConn.DriverName := 'FB';
      fdConn.Params.Add('User_name=' + edUser.Text);
      fdConn.Params.Add('Password=' + edPassword.Text);
      fdConn.Params.Add('Database=' + edcaminho.Text);
      fdConn.Params.Add('Server='+  edServer.Text);
      fdconn.Connected:= true;
      Logg('Conectado ao banco de dados.');
      //Associa campos
      Logg('Associando Campos...');
      AssociaCampos();
      Arquivoini.Free;
    end;
end;

procedure LerConfiguracoes;
var
ArquivoIni: TIniFile;
begin
  //Ler as configura��es
  With frmPrincipal do //Entra no Escopo do form principal
    begin
      //Abre o ini
      ArquivoIni:= TiniFile.Create(IniPath);
      //L� configura��es do sistema
      edCaminhoIntegracao.Text:=  ArquivoIni.ReadString('Configuracoes','CaminhoSaida', 'C:\Integracao\'); //Caminho de sa�da padr�o
      edIntervalo.Text:= inttostr(ArquivoIni.ReadInteger('Configuracoes','Intervalo', 60));                //Inteiro com os segundos
      //L� conex�o do banco
      eduser.Text:=     ArquivoIni.ReadString('Configuracoes','username', 'sysdba');     //Padr�o Firebird
      edPassword.Text:= ArquivoIni.ReadString('Configuracoes','password', 'masterkey');  //Padr�o Firebird
      edCaminho.Text:=  ArquivoIni.ReadString('Configuracoes','database', '');           //Caminho do Banco
      edServer.Text:=   ArquivoIni.ReadString('Configuracoes','server',   'localhost');  //Caminho do Servidor
      //Conex�o com o Banco
      Logg('Conectando com o banco...');
      fdConn.Connected:= false;
      fdconn.Params.Clear;
      fdConn.DriverName := 'FB';
      fdConn.Params.Add('User_name='  + edUser.Text);
      fdConn.Params.Add('Password='   + edPassword.Text);
      fdConn.Params.Add('Database='   + edcaminho.Text);
      fdConn.Params.Add('Server='     +  edServer.Text);
      fdconn.Connected:= true;
      Logg('Conectado ao banco de dados.');
      //Associa campos
      Logg('Associando Campos...');
      AssociaCampos();
      Arquivoini.Free;
    end;
end;

function OnlyNumber(N: String): String;
var
   I: Byte;
begin
     Result := EmptyStr;
     for I := 1 to Length(N) do
     begin
          if CharInSet(N[I], ['0'..'9']) then
             Result := Result + N[I];
     end;
end;

function inttoativo(i: integer): string;
begin
  case i of
  0: Result:= 'S';
  else Result:= 'N';
  end;
end;

function IsEmptyOrNull(const Value: Variant): Boolean;
begin
  Result := VarIsClear(Value) or VarIsEmpty(Value) or VarIsNull(Value) or (VarCompareValue(Value, Unassigned) = vrEqual);
  if (not Result) and VarIsStr(Value) then
    Result := Value = '';
end;

function FormatarCampo(Dados: Variant; campo: integer): string;
begin
  if not(IsEmptyOrNull(Dados)) then
  begin
    case campo of
      0:begin //Campo Id_Loja
      Result:= Copy(Vartostr(Dados),0,25);
      end;
      1..6:begin //Campos Departamento at� Volume
      Result:= Copy(Vartostr(Dados),0,100);
      end;
      7:begin //Campos NUM�RICO(15) C�digo de barras
      Result:= OnlyNumber(Copy(Vartostr(Dados),0,15));
      end;
      8,15,22:begin //Campos Alfanum�rico 150: Nome, Descricao, image_url
      Result:= Copy(Vartostr(Dados),0,150);
      end;
      9..10:begin //Campos Data: dt_cadastro e dt_ultima_alteracao
      Result:= formatDateTime('dd/mm/yyyy', VarToDateTime(Dados));
      end;
      11,12,18:begin //Campos Num�rico 12 com 2
      Result:= formatFloat('###,##0.00', Double(Dados));
      end;
      13,14,20:begin //Campos Num�rico 12 com 4
      Result:= formatFloat('###,##0.0000', Double(Dados));
      end;
      16, 19:begin //Strings tamanho 1
      Result:= Copy(Vartostr(Dados),0,1);
      end;
      17:begin //Alfanum�rico 15
      Result:= Copy(Vartostr(Dados),0,15);
      end;
      21:begin //Num�rico 2
      Result:= OnlyNumber(Copy(Vartostr(Dados),0,2));
      end;
    end;
  end
    else
      begin
        Result:= '';
      end;
end;

procedure IniciaProcesso;
var
i: integer;
StrTemp: string;
ArquivoFinal: TStringList;
begin
  with frmPrincipal do
    begin
      //Para o timer
      tmrPrincipal.Enabled:= false;
      Logg('Iniciando processo...');
      //Define novo tempo de espera pra ele
      tmrPrincipal.Interval:= strtoint(edIntervalo.Text);

      //Prepara Vari�veis
      ArquivoFinal:= TStringList.Create();
      StrTemp:= '';

      fdQuery.Close;
      fdquery.SQL.Clear;
      fdquery.SQL.Add('select * from c000025');
      fdquery.open;
      fdquery.First;

      Logg('Gerando arquivo...');
      while not fdquery.Eof do
        begin
        //Gera o arquivo
        StrTemp:= '001;'; //Ainda n�o diferencia lojas
          for i := 1 to Campos.Count -1 do
            begin
              if not(Campos[i] = '') then
              begin
               //Condi��o Estoque Atual
                 case i of
                  1://Departamento
                  begin
                    fdQuery2.Close;
                    fdquery2.SQL.Clear;
                    fdquery2.SQL.Add('select * from c000017 where codigo = :CODGRUPO');
                    fdquery2.ParamByName('CODGRUPO').Value:= fdquery.FieldByName('CODGRUPO').Value;
                    fdquery2.open;
                    fdquery2.First;
                    StrTemp:= StrTemp + FormatarCampo(fdquery2.FieldByName(Campos[i]).Value, I) + ';';
                  end;
                  2://Grupo
                  begin
                    fdQuery2.Close;
                    fdquery2.SQL.Clear;
                    fdquery2.SQL.Add('select * from c000018 where codigo = :CODSUBGRUPO');
                    fdquery2.ParamByName('CODSUBGRUPO').Value:= fdquery.FieldByName('CODSUBGRUPO').Value;
                    fdquery2.open;
                    fdquery2.First;
                    StrTemp:= StrTemp + FormatarCampo(fdquery2.FieldByName(Campos[i]).Value, I) + ';';
                  end;
                  3://Subgrupo
                  begin
                    fdQuery2.Close;
                    fdquery2.SQL.Clear;
                    fdquery2.SQL.Add('select * from c000018 where codigo = :CODSUBGRUPO');
                    fdquery2.ParamByName('CODSUBGRUPO').Value:= fdquery.FieldByName('CODSUBGRUPO').Value;
                    fdquery2.open;
                    fdquery2.First;
                    StrTemp:= StrTemp + FormatarCampo(fdquery2.FieldByName(Campos[i]).Value, I) + ';';
                  end;
                  13:
                  begin
                    fdQuery2.Close;
                    fdquery2.SQL.Clear;
                    fdquery2.SQL.Add('select * from c000100 where codproduto = :prod');
                    fdquery2.ParamByName('prod').Value:= fdquery.FieldByName('codigo').Value;
                    fdquery2.open;
                    fdquery2.First;
                    StrTemp:= StrTemp + FormatarCampo(fdquery2.FieldByName(Campos[i]).Value, I) + ';';
                  end;
                  16:begin
                    //Int para Ativo
                    StrTemp:= StrTemp + FormatarCampo(inttoativo(fdquery.FieldByName(Campos[i]).Value), I) + ';';
                  end
                    else
                      begin
                        StrTemp:= StrTemp + FormatarCampo(fdquery.FieldByName(Campos[i]).Value, I) + ';';
                      end;
                 end;
              end
                else
                  begin
                    //Campo sem implemennta��o
                    StrTemp:= StrTemp + ';'
                  end;
            end;
        ArquivoFinal.Add(strtemp);
        fdquery.Next;
        end;

      //Grava na pasta designada
      ArquivoFinal.SaveToFile(edCaminhoIntegracao.Text + '\produtos.csv');
      Logg('Arquivo gerado.');
      //Reinicia Timer
      tmrPrincipal.Enabled:= true;
    end;
end;

procedure TfrmPrincipal.FormCreate(Sender: TObject);
begin
  //Inicializa O projeto
  IniPath:= ExtractFilePath(Application.ExeName) + 'SiteMercadoIntegrador.ini';
  Logg('Inicializando projeto, vers�o ' + FloatToStr(versaoAtual));
  lbl_ver.Caption:= FloatToStr(versaoAtual);
  Logg('Buscando Configura��es...');
  if FileExists(IniPath) then
    begin
      //L� configura��o existente
      LerConfiguracoes();
      IniciaProcesso();
    end
      else
        begin
          //Primeira Configura��o, abre tela na aba de configura��es
          Logg('Arquivo de configura��o n�o encontrado, iniciando primeira vez...');
          pgControl.TabIndex:= 1; //Tab de Configura��es
          frmPrincipal.Show;
        end;
end;

procedure TfrmPrincipal.FormShow(Sender: TObject);
begin
  Left:=(Screen.Width-Width)  div 2;
  Top:=(Screen.Height-Height) div 2;
end;

procedure TfrmPrincipal.pnTopMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
const
  SC_DRAGMOVE = $F012;
begin
  if Button = mbLeft then
  begin
    ReleaseCapture;
    Perform(WM_SYSCOMMAND, SC_DRAGMOVE, 0);
  end;
end;

procedure TfrmPrincipal.popbtnProcessarAgoraClick(Sender: TObject);
begin
IniciaProcesso;
end;

procedure TfrmPrincipal.ppbtnSairClick(Sender: TObject);
begin
Application.Terminate;
end;

procedure TfrmPrincipal.TrayClick(Sender: TObject);
begin
pgControl.TabIndex:= 0;
frmprincipal.Show;
end;

procedure TfrmPrincipal.btnSalvarClick(Sender: TObject);
begin
  GravaConfiguracoes;
  IniciaProcesso;
  frmprincipal.Hide;
end;

end.
