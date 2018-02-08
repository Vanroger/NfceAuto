program NFceAuto;
{$APPTYPE GUI}

uses
  Vcl.Forms,
  Web.WebReq,
  IdHTTPWebBrokerBridge,
  uPrincipal in 'uPrincipal.pas' {frmPrincipal},
  uSM in 'uSM.pas' {ServerMethods1: TDSServerModule},
  uSC in 'uSC.pas' {ServerContainer1: TDataModule},
  uWM in 'uWM.pas' {WebModule1: TWebModule},
  ufrmCertificado in 'ufrmCertificado.pas' {frmCertificado},
  untGerenciadorNFCe in 'untGerenciadorNFCe.pas' {frmGerenciadorNFCe},
  uCertificado in 'uCertificado.pas',
  uEmitente in 'uEmitente.pas',
  ufrmEmitente in 'ufrmEmitente.pas' {frmEmitente},
  untIdentificacao in 'untIdentificacao.pas',
  uItem in 'uItem.pas',
  uDestinatario in 'uDestinatario.pas',
  uArquivoJson in 'uArquivoJson.pas',
  uWebServices in 'uWebServices.pas',
  uIdentificacao in 'uIdentificacao.pas';

{$R *.res}

begin
  if WebRequestHandler <> nil then
    WebRequestHandler.WebModuleClass := WebModuleClass;
  Application.Initialize;
  Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.Run;
end.
