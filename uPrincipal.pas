unit uPrincipal;

interface

uses
  Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.AppEvnts, Vcl.StdCtrls, IdHTTPWebBrokerBridge, Web.HTTPApp, Vcl.Buttons,
  ufrmCertificado, untGerenciadorNFCe, ufrmEmitente, uItem, rest.json,
  system.generics.collections, uDestinatario, uArquivoJson;

type
  TfrmPrincipal = class(TForm)
    ButtonStart: TButton;
    ButtonStop: TButton;
    EditPort: TEdit;
    Label1: TLabel;
    ApplicationEvents1: TApplicationEvents;
    ButtonOpenBrowser: TButton;
    SpeedButton1: TSpeedButton;
    btnEmitente: TSpeedButton;
    SpeedButton2: TSpeedButton;
    Label2: TLabel;
    Edit1: TEdit;
    Label3: TLabel;
    Edit2: TEdit;
    GroupBox1: TGroupBox;
    procedure FormCreate(Sender: TObject);
    procedure ApplicationEvents1Idle(Sender: TObject; var Done: Boolean);
    procedure ButtonStartClick(Sender: TObject);
    procedure ButtonStopClick(Sender: TObject);
    procedure ButtonOpenBrowserClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnEmitenteClick(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    FServer: TIdHTTPWebBrokerBridge;
    procedure StartServer;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmPrincipal: TfrmPrincipal;

implementation

{$R *.dfm}

uses
  WinApi.Windows, Winapi.ShellApi, Datasnap.DSSession, System.JSON;

procedure TfrmPrincipal.ApplicationEvents1Idle(Sender: TObject; var Done: Boolean);
begin
  ButtonStart.Enabled := not FServer.Active;
  ButtonStop.Enabled := FServer.Active;
  EditPort.Enabled := not FServer.Active;
end;

procedure TfrmPrincipal.btnEmitenteClick(Sender: TObject);
begin
  if frmEmitente = nil then
    frmEmitente := TfrmEmitente.Create(Application);

  frmEmitente.ShowModal;

  FreeAndNil(frmEmitente);
end;

procedure TfrmPrincipal.ButtonOpenBrowserClick(Sender: TObject);
var
  LURL: string;
begin
  StartServer;
  LURL := Format('http://localhost:%s', [EditPort.Text]);
  ShellExecute(0,
        nil,
        PChar(LURL), nil, nil, SW_SHOWNOACTIVATE);
end;

procedure TfrmPrincipal.ButtonStartClick(Sender: TObject);
begin
  StartServer;
  if frmGerenciadorNFCe = nil then
    frmGerenciadorNFCe := TfrmGerenciadorNFCe.Create(Application);
end;

procedure TerminateThreads;
begin
  if TDSSessionManager.Instance <> nil then
    TDSSessionManager.Instance.TerminateAllSessions;
end;

procedure TfrmPrincipal.ButtonStopClick(Sender: TObject);
begin
  TerminateThreads;
  FServer.Active := False;
  FServer.Bindings.Clear;
end;

procedure TfrmPrincipal.FormCreate(Sender: TObject);
begin
  FServer := TIdHTTPWebBrokerBridge.Create(Self);
end;

procedure TfrmPrincipal.SpeedButton1Click(Sender: TObject);
begin
  if frmCertificado = nil then
     frmCertificado := TfrmCertificado.Create(application);

  frmCertificado.showmodal;

  freeandnil(frmCertificado);

end;

procedure TfrmPrincipal.SpeedButton2Click(Sender: TObject);
var
  vJson : Tjsonobject;
  vStr : string;
  cITEM : Titem;
  vDest : TDestinatario;
  vArquivo : TArquivo;
  vArrayItens : TArrayItens;
begin
  vDest := TDestinatario.Create;
  vdest.CNPJ   := '80892477504';
  vdest.XNOME  := 'ROGERIO ALVES';
  vdest.XLGR   := 'RODOVIA GO 070';
  VDEST.NRO    := 'SN';
  VDEST.BAIRRO := 'FAZ SAO DOMINGOS';
  VDEST.CMUN   := '5208707';
  VDEST.XMUN   := 'GOIANIA';
  VDEST.UF     := 'GO';
  VDEST.CEP    := '74477600';

  vArquivo := TArquivo.Create;
  vArquivo.Destinatario := vDest;

  setlength(vArrayItens,2);

  cITEM := Titem.Create;
  cITEM.Codigo     := '12';
  cITEM.Nome       := 'coca cola';
  cITEM.Quantidade := 3;
  cITEM.Unitario   := 1.50;
  cITEM.Total      := 4.50;
  cITEM.NCM        := '22021000';
  cITEM.Origem     := '0';
  cITEM.CST        := '00';
  cITEM.CSOSN      := '500';
  cITEM.Aliquota   := 17;
  cITEM.CFOP       := '5405';
  cITEM.unidade    := 'UND';
  cITEM.vBC        := 4.50;
  cITEM.pICMS      := 17;
  cITEM.vICMS      := 0.76;
  cItem.CEST       := '0301100';


  vArrayItens[0] := citem;

  cITEM := titem.Create;
  cITEM.Codigo     := '15';
  cITEM.Nome       := 'guarana x';
  cITEM.Quantidade := 4;
  cITEM.Unitario   := 1.90;
  cITEM.Total      := 7.60;
  cITEM.NCM        := '22021000';
  cITEM.Origem     := '0';
  cITEM.CST        := '00';
  cITEM.CSOSN      := '500';
  cITEM.Aliquota   := 12;
  cITEM.CFOP       := '5405';
  cITEM.unidade    := 'UND';
  cITEM.vBC        := 7.60;
  cITEM.pICMS      := 17;
  cITEM.vICMS      := 1.29;
  cItem.CEST       := '0301100';

  vArrayItens[1] := citem;

  varquivo.Itens := vArrayItens;

  vJSON := TJson.ObjectToJsonObject(varquivo);
  frmGerenciadorNFCe.EnviarNFCe(vjson);
end;

procedure TfrmPrincipal.StartServer;
begin
  if not FServer.Active then
  begin
    FServer.Bindings.Clear;
    FServer.DefaultPort := StrToInt(EditPort.Text);
    FServer.Active := True;
  end;
end;

end.
