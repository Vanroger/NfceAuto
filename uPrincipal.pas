unit uPrincipal;

interface

uses
  Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.AppEvnts, Vcl.StdCtrls, IdHTTPWebBrokerBridge, Web.HTTPApp, Vcl.Buttons,
  ufrmCertificado, untGerenciadorNFCe, ufrmEmitente, uItem, rest.json,
  system.generics.collections;

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
  vITEM : TITEM;
  vJson : Tjsonobject;
  vStr : string;
  cITEM : TObjectList<Titem>;
begin
  cITEM := TObjectList<Titem>.Create;

  cITEM.Add(titem.Create);
  cITEM[0].Codigo     := '12';
  cITEM[0].Nome       := 'coca cola';
  cITEM[0].Quantidade := 3;
  cITEM[0].Unitario   := 1.50;
  cITEM[0].Total      := 0;
  cITEM[0].NCM        := '795678';
  cITEM[0].Origem     := '0';
  cITEM[0].CST        := '60';
  cITEM[0].CSOSN      := '100';
  cITEM[0].Aliquota   := 17;

  cITEM.Add(titem.Create);
  cITEM[1].Codigo     := '15';
  cITEM[1].Nome       := 'guarana x';
  cITEM[1].Quantidade := 4;
  cITEM[1].Unitario   := 1.90;
  cITEM[1].Total      := 0;
  cITEM[1].NCM        := '793568';
  cITEM[1].Origem     := '0';
  cITEM[1].CST        := '60';
  cITEM[1].CSOSN      := '500';
  cITEM[1].Aliquota   := 12;

//  vStr := '{"PRODUTOS":[{"fQuantidade":3,"fNome":"coca cola","fCST":"60","fOrigem":"0","fCSOSN":"100","fUnitario":1.5,"fCodigo":"12","fTotal":0,"fNCM":"795678","fAliquota":17},{"fQuantidade":1,"fNome":"GUARANA","fCST":"60","fOrigem":"0","fCSOSN":"100","fUnitario":1.9,"fCodigo":"17","fTotal":0,"fNCM":"795678","fAliquota":17}]}';

  vjson  := TJson.ObjectToJsonObject(cITEM);

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
