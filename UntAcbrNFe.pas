unit UntAcbrNFe;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ConexaoAcbrMonitorNfe, IniFiles,dxGDIPlusClasses, ExtCtrls,
  untConstante, Printers, smtpsend, mimemess, mimepart, RpDevice, IdMessage,
  IdSMTP, IdBaseComponent, IdComponent, IdIOHandler, IdIOHandlerSocket,
  IdSSLOpenSSL, unitdm, untManifestacaoDest, frxClass, frxPreview, QuickRpt,
  ACBrBase, ACBrDFe, ACBrNFe, pcnConversao, ACBrNFeDANFEClass,
  pcnConversaoNFe, ACBrDFeUtil, ACBrNFeDANFeRLClass, ACBrDANFCeFortesFr,
  ACBrNFeDANFEFR, ACBrNFeDANFeESCPOS, ACBrPosPrinter, RLReport, ACBrUtil,
  ACBrMail, unitDanfeNfeSimplificadoFortes, IdRawBase, IdRawClient, IdIcmpClient;

type

  TfrmAcbrNFe = class(TForm)
    ACBrNFe1: TACBrNFe;
    ACBrNFeDANFeRL1: TACBrNFeDANFeRL;
    ACBrNFeDANFEFR: TACBrNFeDANFEFR;
    ACBrNFeDANFCeFortes1: TACBrNFeDANFCeFortes;
    ACBrPosPrinter1: TACBrPosPrinter;
    ACBrNFeDANFeESCPOS1: TACBrNFeDANFeESCPOS;
    ACBrMail1: TACBrMail;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    function ConvertStrRecived(AStr: String): String;
    function SetCertificado(pNumeroSerie: AnsiString; pSenha: AnsiString): Boolean; overload;
    function SetWebServices(pAmbiente: Integer; pUF: string): Boolean;
//    procedure EnviarEmailIndy(const sSmtpHost, sSmtpPort, sSmtpUser,
//      sSmtpPasswd, sFrom, sTo, sAssunto, sAttachment, sAttachment2: String;
//      sMensagem: TStrings; SSL, TLS: Boolean; sCopias: String);
    function SubstituirVariaveis(const ATexto: String): String;
    function IntToStrZero(vr, qtd: Integer): string;
    procedure ConfiguraDANFeNfce(pNFE: Boolean = FALSE);
    function ExtractRes: Boolean;
    procedure GerarNFCe(pComando: AnsiString);
    procedure GerarIniNFe(AStr: String);
    function UFparaCodigo(const UF: string): integer;
    function SetCSC_Token(pCSC_Token: String): boolean;
    function SetIdCSC_IdToken(pIdCSC_IdToken: String): boolean;
    function SetModeloDF(pModelo: String): boolean;
  public
    { Public declarations }
    OK : Boolean;
    function pingIp(host: String): Boolean;
    function AssinarNFE(pEnderecoXML: string; var ret: TRetorno): Boolean;
    function SetFormaEmisao(pTpEmiss: integer; var ret: TRetorno): boolean;
    function StatusServico(var ret: TRetorno): boolean;
    function SavetoFile(pArquivo: String; pXML: AnsiString; var ret: TRetorno): boolean;
    function loadfromfile(pFile: AnsiString; var ret: TRetorno): boolean;
    function FileExiste(pFile: AnsiString; var ret: TRetorno): boolean;
    function ValidarNFe(pXML: AnsiString; var ret: TRetorno): boolean;
    function SetCertificado(pNumeroSerie: AnsiString; pSenha: AnsiString; var ret: TRetorno): Boolean; overload;

    function ValidadeCertificado: string;
    function DataVencimentoCertificado: TDateTime;
    function CancelarNfe(pChave, pJust: string; var ret: TRetorno; cOrgao: Integer): Boolean;
    function CartaCorrecao(idLote     : Integer;
                           cOrgao     : Integer;
                           CNPJ       : string;
                           chNFe      : string;
                           nSeqEvento : array of Integer;
                           xCorrecao  : array of AnsiString;
                           var ret    : TRetorno): Boolean;
    procedure EnviarEmail(const sSmtpHost, sSmtpPort, sSmtpUser, sSmtpPasswd,
      sFrom, sTo, sAssunto, sAttachment, sAttachment2: String;
      sMensagem: TStrings; SSL, TLS: Boolean; sCopias: String='');
    procedure ConfiguraDANFe;
    function ImprimirDanfe(var ret : TRetorno; pXML: AnsiString): boolean;
    function ImprimirDanfeNfce(pArquivo: AnsiString; pNFE: Boolean): Boolean;
    procedure ImprimirDanfeNfceCancelado(pArquivo: AnsiString);
    function ConsultaNfeDest(pIndicadorNFE, pIndicadorEmissor, pUltimoNSU: string; var ret : TRetorno): Boolean;
    function EnviarManifestacao(pListaNFe: Tlist; pTpEvento: string; pJust: string; cOrgao    : integer; var ret: TRetorno): Boolean;
    function LerXML(pArquivo: AnsiString): Boolean; //Lê o XML pelo Arquivo
    function LeXML(pXML: AnsiString): Boolean; //Lê o XML que está na tabela NFe_historicoxml
    function ImprimeRelatorio(pTexto: TStringList): boolean;

    //funçoes acbrlocal
    function CriarNFe(var ret : TRetorno; pComando: AnsiString; pModelo: String; pVersaoNFe : String = '3.10'): Boolean;
    function EnviarNFCe(var ret : TRetorno; pVersaoNFe : string = '3.10'): Boolean;
    function ConsultaNFeXML(var ret: TRetorno; pEndereco: String): Boolean;
    function ConsultaNFeChave(var ret : TRetorno; pChave  : AnsiString): Boolean;
    function Inutilizar(var ret: TRetorno; cCNPJ, cJustificativa: string; nAno,
                        nModelo, nSerie, nNumInicial, nNumFinal: integer): boolean;
    function ImprimirEvento(var ret : TRetorno; pXML: AnsiString): boolean;
    function ImprimirDanfePDF(var ret: TRetorno; pXML: AnsiString): Boolean;
    function EnviaEmail(pDestino: String; pArquivo: String; pArquivo2: String; var ret : TRetorno): Boolean;

    //funções para buscar dados da nfe lida pelo xml
    function GetChaveNFE(index: Integer = 0): AnsiString;
    function GetDestCNPJ(index: Integer = 0): AnsiString;
    function GetDestNOME(index: Integer = 0): AnsiString;
    function GetNF(index: Integer = 0): AnsiString;
    function GetNumeroNota(index: Integer = 0): AnsiString;
    function GetEmissao(index: Integer = 0): TDateTime;
    function GetModelo(index: Integer = 0): Integer;
    function GetSerie(index: Integer = 0): Integer;
    function GetFinalidade(index: Integer = 0): Integer;

    //emitente
    function GetEmitcMun(index: Integer = 0): String;
    function GetEmitCNPJ(index: Integer = 0): String;
    function GetEmitNome(index: Integer = 0): String;
    function GetEmitLgr(index: Integer = 0): String;
    function GetEmitBairro(index: Integer = 0): String;
    function GetEmitCEP(index: Integer = 0): String;
    function GetEmitMun(index: Integer = 0): String;
    function GetEmitUF(index: Integer = 0): String;
    function GetEmitFone(index: Integer = 0): String;
    function GetEmitIE(index: Integer = 0): String;

    //Transportador
    function GetTranspCNPJ(index: Integer = 0): String;
    function GetTranspMun(index: Integer = 0): String;
    function GetTranspNome(index: Integer = 0): String;
    function GetTranspIE(index: Integer = 0): String;
    function GetTranspEndereco(index: Integer = 0): String;
    function GetTranspUF(index: Integer = 0): String;
    function GetTranspVeicPlaca(index: Integer = 0): String;

    //Totais
    function GetNfeValorICMS(index: Integer = 0): Double;
    function GetNfeValorProdutos(index: Integer = 0): Double;
    function GetNfeValorNF(index: Integer = 0): Double;
    function GetNfeValorPIS(index: Integer = 0): Double;
    function GetNfeValorBC(index: Integer = 0): double;
    function GetNfeValorDesconto(index: Integer = 0): double;
    function GetNfeValorFRETE(index: Integer = 0): double;
    function GetNfeValorOUTRO(index: Integer = 0): double;
    function GetNfeValorSeguro(index: Integer = 0): double;

    //Fatura
    function GetFaturaCount(index: Integer = 0): integer;
    function GetFaturaDataVenc(index: Integer = 0): TDateTime;
    function GetFaturaValor(index: Integer = 0): Double;

    //Itens
    function GetDetCount: Integer;
    function GetValorProduto(index: Integer = 0): Double;
    function GetProdVrDesc(index: Integer = 0): Double;
    function GetProdEAN(index: Integer = 0): string;
    function GetProdCod(index: Integer = 0): string;
    function GetProdDescricao(index: Integer = 0): string;
    function GetProd_qCom(index: Integer = 0): Double;
    function GetProd_vDesc(index: Integer = 0): Double;
    function GetProd_vFrete(index: Integer = 0): Double;
    function GetProd_vOutro(index: Integer = 0): Double;
    function GetProd_vProd(index: Integer = 0): Double;
    function GetProd_vSeg(index: Integer = 0): Double;
    function GetProd_vUnCom(index: Integer = 0): Double;
    function GetProd_CFOP(Index: Integer = 0): string;

    //imposto
    function GetProdImpCST(index: Integer = 0): string;
    function GetProdIcmsAliq(index: Integer = 0): Double;
    function GetProdIcmsBC(index: Integer = 0): Double;
    function GetProdIcmsBCST(index: Integer = 0): Double;
    function GetProd_pICMSST(index: Integer = 0): Double;
    function GetProd_vICMSST(index: Integer = 0): Double;
    function GetProd_ICMSorig(index: Integer = 0): Double;
    function GetProd_vICMS(index: Integer = 0): Double;

    function GetNfeValorBCICMSSUBS(index: Integer = 0): double;
    function GetNfeValorICMSSUBS(index: Integer = 0): double;

    function GetProd_vBCSTRet(index: Integer = 0): Double;
    function GetProd_vICMSSTRet(index: Integer = 0): Double;

    function GetProd_vPIS(index: Integer = 0): Double;
    function GetProd_pPIS(index: Integer = 0): Double;
    function GetProd_pCOFINS(index: Integer = 0): Double;
    function GetProd_vCOFINS(index: Integer = 0): Double;

    function GetProd_IPITrib_CST(index: Integer = 0): string;
    function GetProd_vIPI(index: Integer = 0): Double;
    function GetProd_vIPIBC(index: Integer = 0): Double;
    function GetProd_AliqIPI(index: Integer = 0): Double;
  end;

var
  frmAcbrNFe: TfrmAcbrNFe;


implementation

uses UnitFuncao, untFuncoes;

{$R *.dfm}
{$RESOURCE recursoFR.RES}

{ TForm2 }

function TfrmAcbrNFe.pingIp( host: String): Boolean;
var
  IdICMPClient: TIdICMPClient;
begin
  try

    IdICMPClient := TIdICMPClient.Create( nil );

    IdICMPClient.Host := host;
    IdICMPClient.ReceiveTimeout := 500;
    IdICMPClient.Ping;

    result := ( IdICMPClient.ReplyStatus.BytesReceived > 0 );

  finally
    IdICMPClient.Free;
  end
end;

function TfrmAcbrNFe.ExtractRes: Boolean;
var
  Path  : String;
  Path2 : string;
  Res   : TResourceStream;
  Res2  : TResourceStream;
  rec   : string;
begin
  Path := ExtractFilePath(Application.ExeName)+'Report\DANFeNFCeA4.fr3';
  Path2:= ExtractFilePath(Application.ExeName)+'Report\EVENTOS.fr3';
  if (not FileExists(Path) )then begin
    try
      try
        rec := 'REC1';
        Res := TResourceStream.Create(Hinstance, rec, RT_RCDATA);

        if not DirectoryExists(ExtractFilePath(Application.ExeName)+'\Report') then
          ForceDirectories(ExtractFilePath(Application.ExeName)+'\Report');

        Res.SavetoFile(Path);
      finally
        Res.Free;
        result := True;
      end;
    except
      result := false;
    end;
  end;
  if (not FileExists(Path2) ) then begin
    try
      try
        rec := 'REC2';
        Res2 := TResourceStream.Create(Hinstance, rec, RT_RCDATA);

        if not DirectoryExists(ExtractFilePath(Application.ExeName)+'\Report') then
          ForceDirectories(ExtractFilePath(Application.ExeName)+'\Report');

        Res2.SaveToFile(path2);
      finally
        Res2.Free;
        result := True;
      end;
    except
      result := false;
    end;
  end;
end;

function TfrmAcbrNFe.ConsultaNFeXML(var ret : TRetorno;
                                    pEndereco  : String): Boolean;
begin
  try
    acbrnfe1.NotasFiscais.Clear;
    result := ACBrNFe1.NotasFiscais.LoadFromFile(pEndereco);
    ACBrNFe1.Consultar;

    ret.status.add(ACBrNFe1.WebServices.Consulta.Msg);
    ret.status.add('[CONSULTA]');
    ret.status.add('Versao='+ACBrNFe1.WebServices.Consulta.verAplic);
    ret.status.add('TpAmb='+TpAmbToStr(ACBrNFe1.WebServices.Consulta.TpAmb));
    ret.status.add('VerAplic='+ACBrNFe1.WebServices.Consulta.VerAplic);
    ret.status.add('CStat='+IntToStr(ACBrNFe1.WebServices.Consulta.CStat));
    ret.status.add('XMotivo='+ACBrNFe1.WebServices.Consulta.XMotivo);
    ret.status.add('CUF='+IntToStr(ACBrNFe1.WebServices.Consulta.CUF));
    ret.status.add('ChNFe='+ACBrNFe1.WebServices.Consulta.NFeChave);
    ret.status.add('DhRecbto='+DateTimeToStr(ACBrNFe1.WebServices.Consulta.DhRecbto));
    ret.status.add('NProt='+ACBrNFe1.WebServices.Consulta.Protocolo);
    ret.status.add('DigVal='+ACBrNFe1.WebServices.Consulta.protNFe.digVal);
  except
    on e: exception do begin
      result := true;
      ret.status.add('XMotivo='+e.Message);
    end;
  end;
end;

function TfrmAcbrNFe.ConsultaNFeChave(var ret : TRetorno;
                                      pChave  : AnsiString): Boolean;
var
 sRetWS : String;
 sRetornoWS : sTRING;
begin
  try
    ACBrNFe1.NotasFiscais.Clear;
    ACBrNFe1.WebServices.Consulta.NFeChave := pChave;
    ACBrNFe1.WebServices.Consulta.Executar;
    Result := (ACBrNFe1.WebServices.Consulta.CStat > 0);

    ret.status.add(ACBrNFe1.WebServices.Consulta.Msg);
    ret.status.add('[CONSULTA]');
    ret.status.add('Versao='+ACBrNFe1.WebServices.Consulta.verAplic);
    ret.status.add('TpAmb='+TpAmbToStr(ACBrNFe1.WebServices.Consulta.TpAmb));
    ret.status.add('VerAplic='+ACBrNFe1.WebServices.Consulta.VerAplic);
    ret.status.add('CStat='+IntToStr(ACBrNFe1.WebServices.Consulta.CStat));
    ret.status.add('XMotivo='+ACBrNFe1.WebServices.Consulta.XMotivo);
    ret.status.add('CUF='+IntToStr(ACBrNFe1.WebServices.Consulta.CUF));
    ret.status.add('ChNFe='+ACBrNFe1.WebServices.Consulta.NFeChave);
    ret.status.add('DhRecbto='+DateTimeToStr(ACBrNFe1.WebServices.Consulta.DhRecbto));
    ret.status.add('NProt='+ACBrNFe1.WebServices.Consulta.Protocolo);
    ret.status.add('DigVal='+ACBrNFe1.WebServices.Consulta.protNFe.digVal);

    sRetWS := ACBrNFe1.WebServices.Consulta.RetWS;
    sRetornoWS := ACBrNFe1.WebServices.Consulta.RetornoWS;
  except
    result := false;
  end;
//  LoadXML(ACBrNFe1.WebServices.Consulta.RetornoWS, WBResposta);
//  LoadConsulta201(ACBrNFe1.WebServices.Consulta.RetWS);
end;

function TfrmAcbrNFe.CancelarNfe(pChave, pJust: string; var ret: TRetorno; cOrgao: Integer): Boolean;
begin
  ret.status := TStringList.Create;

  ACBrNFe1.NotasFiscais.Clear;
  ACBrNFe1.WebServices.Consulta.NFeChave := pChave;

  if not ACBrNFe1.WebServices.Consulta.Executar then
    raise Exception.Create(ACBrNFe1.WebServices.Consulta.Msg);

  ACBrNFe1.EventoNFe.Evento.Clear;
  with ACBrNFe1.EventoNFe.Evento.Add do begin
    infEvento.CNPJ     := copy(ACBrNFe1.WebServices.Consulta.NFeChave,7,14);
    infEvento.cOrgao   := cOrgao;
    infEvento.dhEvento := now;
    infEvento.tpEvento := teCancelamento;
    infEvento.chNFe    := ACBrNFe1.WebServices.Consulta.NFeChave;
    infEvento.detEvento.nProt := ACBrNFe1.WebServices.Consulta.Protocolo;
    infEvento.detEvento.xJust := pJust;
  end;
  try
    ACBrNFe1.EnviarEvento(1);
//    ACBrNFe1.EnviarEventoNFe(1);

     ret.status.Add(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.xMotivo);
     ret.status.Add('[CANCELAMENTO]');
     ret.status.Add('Versao='+ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.verAplic);
     ret.status.Add('TpAmb='+TpAmbToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.TpAmb));
     ret.status.Add('VerAplic='+ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.VerAplic);
     ret.status.Add('CStat='+IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.cStat));
     ret.status.Add('XMotivo='+ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.XMotivo);
     ret.status.Add('CUF='+IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.cOrgao));
     ret.status.Add('ChNFe='+ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.chNFe);
     ret.status.Add('DhRecbto='+DateTimeToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.dhRegEvento));
     ret.status.Add('NProt='+ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.nProt);
     ret.status.Add('tpEvento='+TpEventoToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.tpEvento));
     ret.status.Add('xEvento='+ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.xEvento);
     ret.status.Add('nSeqEvento='+IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.nSeqEvento));
     ret.status.Add('CNPJDest='+ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.CNPJDest);
     ret.status.Add('emailDest='+ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.emailDest);
     ret.status.Add('XML='+ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.XML);
     Result := true;
  except
    on E: Exception do begin
      result := false;
      raise Exception.Create(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.xMotivo+sLineBreak+E.Message);
    end;
  end;
end;

function TfrmAcbrNFe.CartaCorrecao(idLote     : Integer;
                                   cOrgao     : Integer;
                                   CNPJ       : string;
                                   chNFe      : string;
                                   nSeqEvento : array of Integer;
                                   xCorrecao  : array of AnsiString;
                                   var ret    : TRetorno): Boolean;
var
  i      : integer;
begin
  try
    result := False;
    ACBrNFe1.EventoNFe.Evento.Clear;

    ACBrNFe1.EventoNFe.idLote := idLote;
    ACBrNFe1.EventoNFe.Evento.Clear;

    for I := 0 to High(nSeqEvento) do begin
      with ACBrNFe1.EventoNFe.Evento.Add do begin
        infEvento.cOrgao := cOrgao;
        infEvento.CNPJ   := CNPJ;
        infEvento.chNFe  := chNFe;
        infEvento.dhEvento := Now;
        infEvento.tpEvento := teCCe;
        infEvento.nSeqEvento   := nSeqEvento[i];
        infEvento.versaoEvento := '1.00';
        infEvento.detEvento.xCorrecao := xCorrecao[i];
        infEvento.detEvento.descEvento := 'Carta de Correção';
        infEvento.detEvento.xCondUso  := ''; //Texto fixo conforme NT 2011.003 -  http://www.nfe.fazenda.gov.br/portal/exibirArquivo.aspx?conteudo=tsiloeZ6vBw=
      end;
    end;
    ACBrNFe1.EnviarEvento(ACBrNFe1.EventoNFe.idLote);
//    ACBrNFe1.EnviarEventoNFe(ACBrNFe1.EventoNFe.idLote);

    ret.status.Add('idLote='   +IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.idLote));
    ret.status.Add('tpAmb='    +TpAmbToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.tpAmb));
    ret.status.Add('verAplic=' +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.verAplic);
    ret.status.Add('cOrgao='   +IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.cOrgao));
    ret.status.Add('cStat='    +IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.cStat));
    ret.status.Add('xMotivo='  +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.xMotivo);

    for I:= 0 to ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Count-1 do
    begin
      ret.status.Add('[EVENTO'+Trim(IntToStr(I+1))+']');
      ret.status.Add('id='        +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.Id);
      ret.status.Add('tpAmb='     +TpAmbToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.tpAmb));
      ret.status.Add('verAplic='  +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.verAplic);
      ret.status.Add('cOrgao='    +IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.cOrgao));
      ret.status.Add('cStat='     +IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.cStat));
      ret.status.Add('xMotivo='   +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.xMotivo);
      ret.status.Add('chNFe='     +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.chNFe);
      ret.status.Add('tpEvento='  +TpEventoToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.tpEvento));
      ret.status.Add('xEvento='   +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.xEvento);
      ret.status.Add('nSeqEvento='+IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.nSeqEvento));
      ret.status.Add('CNPJDest='  +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.CNPJDest);
      ret.status.Add('emailDest=' +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.emailDest);
      ret.status.Add('dhRegEvento='+DateTimeToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.dhRegEvento));
      ret.status.Add('nProt='     +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.nProt);
    end;
    result := True;
  except
    on e: Exception do begin
      ShowMessage(e.Message+ sLineBreak + 'Não foi possivel enviar a carta de correção tente novamente!');
    end;
  end;
end;

Function TfrmAcbrNFe.ConvertStrRecived( AStr: String ) : String ;
 Var P   : Integer ;
     Hex : String ;
     CharHex : Char ;
begin
  { Verificando por codigos em Hexa }
  Result := AStr ;

  P := pos('\x',Result) ;
  while P > 0 do
  begin
     Hex := copy(Result,P+2,2) ;

     try
        CharHex := Chr(StrToInt('$'+Hex)) ;
     except
        CharHex := ' ' ;
     end ;

     Result := StringReplace(Result,'\x'+Hex,CharHex,[rfReplaceAll]) ;
     P      := pos('\x',Result) ;
  end ;
end;

function TfrmAcbrNFe.CriarNFe(var ret : TRetorno; pComando: AnsiString; pModelo: String; pVersaoNFe : String = '3.10'): Boolean;
var
//  Salvar  : boolean;
  ArqNFe  : String;
  Alertas : String;
  SL      : TStringList;
  PATH    : String;
begin
  try
    ACBrNFe1.NotasFiscais.Clear;
    SetModeloDF(pMODELO);

    if pVersaoNFe = '2.00' then
      ACBrNFe1.Configuracoes.Geral.VersaoDF := ve200
    else if pVersaoNfe = '3.10' then
      ACBrNFe1.Configuracoes.Geral.VersaoDF := ve310
    else if pVersaoNfe = '4.00' then
      ACBrNFe1.Configuracoes.Geral.VersaoDF := ve400;

    ACBrNFe1.Configuracoes.Geral.IncluirQRCodeXMLNFCe := true;

    GerarIniNFe(pComando);

    PATH := ExtractFilePath(Application.ExeName) + 'Arquivos\' + iif(pModelo = '65','NFCE','NFE');

    if not DirectoryExists(PATH) then
      ForceDirectories(PATH);

    ACBrNFe1.Configuracoes.Arquivos.PathSalvar := path;

    // para gravar o retorno do nfe a autorização essa opção abaixo tem que ficar false
    // e a opção ACBrNFe1.Configuracoes.arquivos.salvar = true
    // o xml sempre será salvo na pasta informada acima.
    ACBrNFe1.Configuracoes.Geral.Salvar := false;

    ACBrNFe1.NotasFiscais.GerarNFe;
    Alertas := ACBrNFe1.NotasFiscais.Items[0].Alertas;
    ACBrNFe1.NotasFiscais.Assinar;
    ArqNFe := PathWithDelim(ACBrNFe1.Configuracoes.Arquivos.PathSalvar)+StringReplace(ACBrNFe1.NotasFiscais.Items[0].NFe.infNFe.ID, 'NFe', '', [rfIgnoreCase])+'-nfe.xml';
    ACBrNFe1.NotasFiscais.GravarXML(ArqNFe);
    ACBrNFe1.NotasFiscais.Validar;
    if not FileExists(ArqNFe) then
      raise Exception.Create('Não foi possível criar o arquivo '+ArqNFe);

    ret.status.add('OK: '+ArqNFe);

    if Alertas <> '' then
      ret.status.add('Alertas:'+Alertas);

    SL := TStringList.Create;

    try
      SL.LoadFromFile(ArqNFe);
      ret.status.add(SL.Text);
    finally
      SL.Free;
      Result := true;
    end;
  except
    on e: exception do begin
      result := false;
      ShowMessage(e.message);
    end;
  end;
end;

function TfrmAcbrNFe.Inutilizar(var ret : TRetorno;
                                cCNPJ, cJustificativa: string;
                                nAno,nModelo, nSerie, nNumInicial, nNumFinal: integer): boolean;
var
  sRetWS, sRetornoWS: string;
begin
  try
    ACBrNFe1.WebServices.Inutiliza(cCNPJ, cJustificativa, nAno, nModelo, nSerie, nNumInicial, nNumFinal);
  except
  end;

  try
    sRetWS := ACBrNFe1.WebServices.Inutilizacao.RetWS;
    sRetornoWS := ACBrNFe1.WebServices.Inutilizacao.RetornoWS;
    ret.status.Add('[INUTILIZACAO]');
    ret.status.Add('OK');
    ret.status.Add('Versao='+ACBrNFe1.WebServices.Inutilizacao.verAplic);
    ret.status.Add('TpAmb='+TpAmbToStr(ACBrNFe1.WebServices.Inutilizacao.TpAmb));
    ret.status.Add('VerAplic='+ACBrNFe1.WebServices.Inutilizacao.VerAplic);
    ret.status.Add('XMotivo='+ACBrNFe1.WebServices.Inutilizacao.XMotivo);
    ret.status.Add('CUF='+IntToStr(ACBrNFe1.WebServices.Inutilizacao.CUF));
    ret.status.Add('DhRecbto='+DateTimeToStr(ACBrNFe1.WebServices.Inutilizacao.DhRecbto));
    ret.status.Add('NProt='+ACBrNFe1.WebServices.Inutilizacao.Protocolo);
    ret.status.Add('Arquivo='+ACBrNFe1.WebServices.Inutilizacao.NomeArquivo);
    ret.status.Add('XML='+ACBrNFe1.WebServices.Inutilizacao.XML_ProcInutNFe);
    ret.status.Add('XML_ENVIADO='+ACBrNFe1.WebServices.Inutilizacao.DadosMsg);
    ret.status.Add('CStat='+IntToStr(ACBrNFe1.WebServices.Inutilizacao.CStat));

    Result := True;
  except
    ret.status.Add('ERRO');
    Result := false;
  end;
end;

function TfrmAcbrNFe.AssinarNFE(pEnderecoXML: string; var ret: TRetorno): Boolean;
begin
  try
    if FileExists(pEnderecoXML) then begin
      ACBrNFe1.NotasFiscais.LoadFromFile(pEnderecoXML);
      ACBrNFe1.NotasFiscais.Assinar;
      ret.status.add('OK');
      result := true;
    end else begin
      ret.status.add('ERRO: Arquivo ' + pEnderecoXML + ' não encontrado.');
      result := true;
    end;
  except
    ret.status.Add('ERRO');
    Result := false;
  end;
end;

function TfrmAcbrNFe.SetFormaEmisao(pTpEmiss: integer; var ret: TRetorno): boolean;
begin
  try
    ACBrNFe1.Configuracoes.Geral.FormaEmissao := StrToTpEmis(OK,IntToStr(pTpEmiss));
    ret.status.add('OK:');
    Result := true;
  except
    result := false;
  end;
end;

function TfrmAcbrNFe.StatusServico(var ret: TRetorno): boolean;
begin
  try
    ACBrNFe1.WebServices.StatusServico.Executar;
  except
    on e: exception do begin
      ret.status.add('ERRO');
      ret.status.add('Inativo ou Inoperante tente novamente!');
      ret.status.add(e.Message);
      ShowMessage(e.message);
      result := true;
      exit;
    end;
  end;
  try
    ret.status.add(ACBrNFe1.WebServices.StatusServico.Msg);
    ret.status.add('[STATUS]');
    ret.status.add('Versao='+ACBrNFe1.WebServices.StatusServico.verAplic);
    ret.status.add('TpAmb='+TpAmbToStr(ACBrNFe1.WebServices.StatusServico.TpAmb));
    ret.status.add('VerAplic='+ACBrNFe1.WebServices.StatusServico.VerAplic);
    ret.status.add('CStat='+IntToStr(ACBrNFe1.WebServices.StatusServico.CStat));
    ret.status.add('XMotivo='+ACBrNFe1.WebServices.StatusServico.XMotivo);
    ret.status.add('CUF='+IntToStr(ACBrNFe1.WebServices.StatusServico.CUF));
    ret.status.add('DhRecbto='+DateTimeToStr(ACBrNFe1.WebServices.StatusServico.DhRecbto));
    ret.status.add('TMed='+IntToStr(ACBrNFe1.WebServices.StatusServico.TMed));
    ret.status.add('DhRetorno='+DateTimeToStr(ACBrNFe1.WebServices.StatusServico.DhRetorno));
    ret.status.add('XObs='+ACBrNFe1.WebServices.StatusServico.XObs);
    result := true;
  except
    on e: exception do begin
      ret.status.add('ERRO');
      ret.status.add(e.Message);
      ShowMessage(e.message);
      result := true;
    end;
  end;
end;

function TfrmAcbrNFe.SavetoFile(pArquivo: String; pXML: AnsiString; var ret: TRetorno): boolean;
var
  Lines : TStringList;
begin
  Lines := TStringList.Create;
  try
    Lines.Clear;
    Lines.Text := ConvertStrRecived( pXML );
    Lines.SaveToFile(pArquivo);
    ret.status.Add('OK');
    Result := true;
    Lines.Free;
  except
    result := false;
  end;
end;

function TfrmAcbrNfe.loadfromfile(pFile: AnsiString; var ret: TRetorno): boolean;
var
  Arquivo : TStringList;
begin
  try
    Result := false;
    Arquivo := TStringList.create;
    if FileExists(pFile) then begin
      Arquivo.LoadFromFile(pFile);
      ret.status.Add('OK =' + Arquivo.Text);
      result := true;
    end;
  except
    on e: exception do begin
      result := false;
      ShowMessage(e.Message);
    end;
  end;
end;

function TFrmAcbrNFe.FileExiste(pFile: AnsiString; var ret: TRetorno): boolean;
begin
  try
    result := FileExists(pFile);
    ret.status.add('OK');
  except
    result := false;
  end;
end;

function TFrmAcbrNFE.ValidarNFe(pXML: AnsiString; var ret: TRetorno): boolean;
begin
  try
    if FilesExists(pXML) then begin
      ACBrNFe1.NotasFiscais.LoadFromFile(pXML,False);
      acbrnfe1.NotasFiscais.Validar;
      ret.status.add('OK');
      Result := true;
    end;
  except
    Result := false;
    ret.status.add('ERRO');
  end;
end;

function TfrmAcbrNFe.EnviarNFCe(var ret : TRetorno; pVersaoNFe : string = '3.10'): Boolean;
var
  Sincrono : boolean;
  i        : integer;
begin
  try
    Result := TRUE;
    if pVersaoNFe = '2.00' then
      ACBrNFe1.Configuracoes.Geral.VersaoDF := ve200
    else if pVersaoNfe = '3.10' then
      ACBrNFe1.Configuracoes.Geral.VersaoDF := ve310
    else if pVersaoNfe = '4.00' then
      ACBrNFe1.Configuracoes.Geral.VersaoDF := ve400;

    // sincrono - já tem a resposta no retorno se foi ou não autorizada
    //assincrono - tem que fazer uma consulta com o protocolo.
    Sincrono := true;// False;

//    ret.status.add('ERRO');
//    raise exception.create('Teste de erro');

    ACBrNFe1.WebServices.Enviar.Lote := '1';
    ACBrNFe1.WebServices.Enviar.Sincrono := Sincrono;
    ACBrNFe1.WebServices.Enviar.Executar;

    ret.status.add(ACBrNFe1.WebServices.Enviar.Msg);
    ret.status.add('[ENVIO]');
    ret.status.add('Versao='+ACBrNFe1.WebServices.Enviar.verAplic);
    ret.status.add('TpAmb='+TpAmbToStr(ACBrNFe1.WebServices.Enviar.TpAmb));
    ret.status.add('VerAplic='+ACBrNFe1.WebServices.Enviar.VerAplic);
    ret.status.add('XMotivo='+ACBrNFe1.WebServices.Enviar.XMotivo);
    ret.status.add('NRec='+ACBrNFe1.WebServices.Enviar.Recibo);
    ret.status.add('CUF='+IntToStr(ACBrNFe1.WebServices.Enviar.CUF));
    ret.status.add('DhRecbto='+DateTimeToStr( ACBrNFe1.WebServices.Enviar.dhRecbto));
    ret.status.add('TMed='+IntToStr( ACBrNFe1.WebServices.Enviar.tMed));
    ret.status.add('Recibo='+ACBrNFe1.WebServices.Enviar.Recibo);
    ret.status.add('CStat='+IntToStr(ACBrNFe1.WebServices.Enviar.CStat));

    if sincrono then begin
      ret.status.add('[RETORNO]');
      ret.status.add('TpAmb='+TpAmbToStr(ACBrNFe1.WebServices.Enviar.TpAmb));
      ret.status.add('VerAplic='+ACBrNFe1.WebServices.Enviar.verAplic);
      ret.status.add('CHNFE='+ acbrnfe1.NotasFiscais.Items[0].NFe.procNFe.chNFe);
      ret.status.add('DhRecbto='+DateTimeToStr(acbrnfe1.NotasFiscais.Items[0].NFe.procNFe.dhRecbto));
      ret.status.add('NPROT='+acbrnfe1.NotasFiscais.Items[0].NFe.procNFe.nProt);
      ret.status.add('DigVal='+acbrnfe1.NotasFiscais.Items[0].NFe.procNFe.digVal);
      ret.status.add('CStat='+IntToStr(ACBrNFe1.WebServices.Enviar.cstat));
      ret.status.add('XMotivo='+ACBrNFe1.WebServices.Enviar.xmotivo);
      ret.status.add('Versao='+ACBrNFe1.WebServices.Enviar.versao);
      ret.status.add('CUF='+IntToStr(ACBrNFe1.WebServices.Enviar.CUF));
      ret.status.add('NREC='+ACBrNFe1.WebServices.Enviar.Recibo);
    end
    else begin
      if (ACBrNFe1.WebServices.Enviar.Recibo <> '') then begin
        ACBrNFe1.WebServices.Retorno.Recibo := ACBrNFe1.WebServices.Enviar.Recibo;
        ACBrNFe1.WebServices.Retorno.Executar;

        ret.status.add('[RETORNO]');
        ret.status.add('Versao='+ACBrNFe1.WebServices.Retorno.verAplic);
        ret.status.add('TpAmb='+TpAmbToStr(ACBrNFe1.WebServices.Retorno.TpAmb));
        ret.status.add('VerAplic='+ACBrNFe1.WebServices.Retorno.VerAplic);
        ret.status.add('NRec='+ACBrNFe1.WebServices.Retorno.NFeRetorno.nRec);
        ret.status.add('CStat='+IntToStr(ACBrNFe1.WebServices.Retorno.CStat));
        ret.status.add('XMotivo='+ACBrNFe1.WebServices.Retorno.XMotivo);
        ret.status.add('CUF='+IntToStr(ACBrNFe1.WebServices.Retorno.CUF));
        ret.status.add('cMsg='+ IntToStr(ACBrNFe1.WebServices.Retorno.cMsg));
        ret.status.add('xMsg='+ ACBrNFe1.WebServices.Retorno.xMsg);
        ret.status.add('Recibo='+ ACBrNFe1.WebServices.Retorno.Recibo);
        ret.status.add('NPROT='+ ACBrNFe1.WebServices.Retorno.Protocolo);

        if Length(trim(ACBrNFe1.WebServices.Retorno.ChaveNFe)) = 44 then
          ret.status.add('CHNFE='+ ACBrNFe1.WebServices.Retorno.ChaveNFe);

        for I:= 0 to ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Count-1 do begin
          ret.status.add('[NFE'+Trim(IntToStr(StrToInt(copy(ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Items[i].chNFe,26,9))))+']');
          ret.status.add('Versao='+ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Items[i].verAplic);
          ret.status.add('TpAmb='+TpAmbToStr(ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Items[i].tpAmb));
          ret.status.add('VerAplic='+ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Items[i].verAplic);
          ret.status.add('CStat='+IntToStr(ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Items[i].cStat));
          ret.status.add('XMotivo='+ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Items[i].xMotivo);
          ret.status.add('CUF='+IntToStr(ACBrNFe1.WebServices.Recibo.NFeRetorno.cUF));
          ret.status.add('ChNFe='+ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Items[i].chNFe);
          ret.status.add('DhRecbto='+DateTimeToStr(ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Items[i].dhRecbto));
          ret.status.add('NProt='+ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Items[i].nProt);
          ret.status.add('DigVal='+ACBrNFe1.WebServices.Recibo.NFeRetorno.ProtNFe.Items[i].digVal);
        end;

      end;
    end;
    ACBrNFe1.NotasFiscais.Clear;
  except
    on e: exception do begin
      result := false;
      ret.status.add(e.message);
      ACBrNFe1.NotasFiscais.Clear;
    end;
  end;
end;

procedure TfrmAcbrNFe.GerarNFCe(pComando: AnsiString);
begin
//
end;

procedure TfrmAcbrNFe.FormCreate(Sender: TObject);
begin
  SetCertificado(sNumeroSerie_Cert_Digital,sSenha_Cert_Digital);
  SetWebServices(sAmbiente,sUF_WebService);
  if trim(LowerCase(sCSC_Token)) <> '' then
    SetCSC_Token(sCSC_Token);
  if trim(sIdCSC_IdToken) <> '' then
    SetIdCSC_IdToken(sIdCSC_IdToken);
end;

function TfrmAcbrNFe.ImprimirDanfeNfce(pArquivo: AnsiString; pNFE: Boolean): Boolean;
begin
  try
    ACBrNFe1.NotasFiscais.Clear;
    ACBrNFe1.NotasFiscais.LoadFromString(pArquivo);
    ConfiguraDANFeNfce(pNFE);
    ACBrNFe1.NotasFiscais.Imprimir;
    result := true;
  except
    on e: exception do begin
      result := false;
    end;
  end;
end;

procedure TfrmAcbrNFe.ImprimirDanfeNfceCancelado(pArquivo: AnsiString);
begin
  try
    ACBrNFe1.NotasFiscais.Clear;
    ACBrNFe1.NotasFiscais.LoadFromString(pArquivo);
    ConfiguraDANFeNfce;
    ACBrNFe1.NotasFiscais.ImprimirCancelado;
  finally

  end;
end;

function TfrmAcbrNFe.ImprimirEvento(var ret : TRetorno; pXML: AnsiString): boolean;
begin
  try
    ACBrNFe1.EventoNFe.Evento.Clear;
    ACBrNFe1.EventoNFe.LerXML(pXML);
    ACBrNFe1.ImprimirEvento;

    ret.status.Add('OK= Evento Impresso com sucesso!');
    ret.status.Add('VERSAO='+ACBrNFe1.EventoNFe.Versao);
    Result := True;
  except
    on e: exception do begin
      ret.status.Add('ERRO='+ e.message);
      Result := false;
    end;
  end;
end;

function TfrmAcbrNFe.ImprimirDanfe(var ret : TRetorno; pXML: AnsiString): boolean;
begin
  try
    result := false;
    ACBrNFe1.NotasFiscais.Clear;
    if ACBrNFe1.NotasFiscais.LoadFromString(pXML) then begin
      ConfiguraDANFe;
      ACBrNFe1.NotasFiscais.Imprimir;

      ret.status.Add('OK= Danfe Impresso com sucesso!');
      ret.status.Add('VERSAO=');
      result := true;
    end;
  except
    on e: exception do begin
      ret.status.Add('ERRO='+ e.message);
      result := false;
    end;
  end;
end;

function TfrmAcbrNFe.ImprimirDanfePDF(var ret : TRetorno; pXML: AnsiString): boolean;
VAR
  ArqPDF : sTRING;
begin
  try
    result := false;
    ACBrNFe1.NotasFiscais.Clear;
    if ACBrNFe1.NotasFiscais.LoadFromString(pXML) then begin
      ConfiguraDANFe;
      ACBrNFe1.NotasFiscais.ImprimirPDF;
      acbrnfe1.DANFE.PathPDF;

      ArqPDF := OnlyNumber(ACBrNFe1.NotasFiscais.Items[0].NFe.infNFe.ID)+'-nfe.pdf';

      ret.status.Add('OK= PDF Impresso com sucesso!');
      ret.status.Add('PDF='+PathWithDelim(ACBrNFe1.DANFE.PathPDF) + ArqPDF);
      result := true;
    end;
  except
    on e: exception do begin
      ret.status.Add('ERRO='+ e.message);
      result := false;
    end;
  end;
end;

function TfrmAcbrNFe.SetCertificado(pNumeroSerie: AnsiString; pSenha: AnsiString): Boolean;
begin
  try
    Result := false;

    if pNumeroSerie <> '' then begin
      ACBrNFe1.Configuracoes.Certificados.NumeroSerie := pNumeroSerie;
      Result := true;
    end
    else
      ShowMessage('Não foi informado o certificado, verifique!');

    if pSenha <> '' then
      ACBrNFe1.Configuracoes.Certificados.Senha := pSenha;
  except
    on e: Exception do begin
      Result := False;
      ShowMessage('Ocorreu um erro ao informar o numero do certificado!' + sLineBreak +
                  e.Message);
    end;
  end;
end;

function TfrmAcbrNFe.SetCertificado(pNumeroSerie: AnsiString; pSenha: AnsiString; var ret: TRetorno): Boolean;
begin
  try
    Result := false;

    if pNumeroSerie <> '' then begin
      ACBrNFe1.Configuracoes.Certificados.NumeroSerie := pNumeroSerie;
      Result := true;
      ret.status.add('OK');
    end
    else
      ret.status.add('Não foi informado o certificado, verifique!');

    if pSenha <> '' then
      ACBrNFe1.Configuracoes.Certificados.Senha := pSenha;
  except
    on e: Exception do begin
      Result := False;
    end;
  end;
end;

function TfrmAcbrNFe.SetWebServices(pAmbiente: Integer; pUF: string): Boolean;
begin
  try
    result := False;
    if ACBrNFe1.Configuracoes.Certificados.NumeroSerie <> '' then begin
      case pAmbiente of
        0: ACBrNFe1.Configuracoes.WebServices.Ambiente := taProducao;
        1: ACBrNFe1.Configuracoes.WebServices.Ambiente := taHomologacao;
      end;

      if (pUF <> '') and (Length(pUF) = 2) then
        ACBrNFe1.Configuracoes.WebServices.UF := Trim(pUF);

      result := True;

      acbrnfe1.Configuracoes.WebServices.AguardarConsultaRet := 5000;
      acbrnfe1.Configuracoes.WebServices.IntervaloTentativas := 2000;
    end;
  except
    on e: Exception do begin
      result := False;
      ShowMessage('Ocorreu um erro ao informar dados do certificado!' + sLineBreak +
                  e.Message);
    end;
  end;
end;

function TfrmAcbrNFe.SetCSC_Token(pCSC_Token: String): boolean;
begin
  try
    result := false;
    if trim(pCSC_Token) = '' then
      exit;

    ACBrNFe1.Configuracoes.Geral.CSC := pCSC_Token;
    result := true;
  except
    on e: Exception do begin
      result := False;
      ShowMessage('Ocorreu um erro ao informar dados do CSC/Token!' + sLineBreak +
                  e.Message);
    end;
  end;
end;

function TfrmAcbrNFe.SetIdCSC_IdToken(pIdCSC_IdToken: String): boolean;
begin
  try
    result := false;
    if trim(pIdCSC_IdToken) = '' then
      exit;

    ACBrNFe1.Configuracoes.Geral.IdCSC := pIdCSC_IdToken;
    result := true;
  except
    on e: Exception do begin
      result := False;
      ShowMessage('Ocorreu um erro ao informar dados do IdCSC/IdToken!' + sLineBreak +
                  e.Message);
    end;
  end;
end;

function TfrmAcbrNFe.SetModeloDF(pModelo: String): boolean;
begin
  if pModelo = '65' then begin
    ACBrNFe1.Configuracoes.Geral.ModeloDF := moNFCe;
  end
  else if pModelo = '55' then begin
    ACBrNFe1.Configuracoes.Geral.ModeloDF := moNFe;
  end;
end;

function TfrmAcbrNFe.ValidadeCertificado: string;
begin
  try
    result := IntToStr(TRUNC(ACBrNFe1.SSL.CertDataVenc - NOW));
//    result := IntToStr(TRUNC(ACBrNFe1.Configuracoes.Certificados.DataVenc - NOW));
  except
    //
  end;
end;

function TfrmAcbrNFe.DataVencimentoCertificado: TDateTime;
begin
  try
    result :=ACBrNFe1.SSL.CertDataVenc;
  except
    //
  end;
end;

function TfrmAcbrNFe.EnviaEmail(pDestino, pArquivo: String; pArquivo2: String; var ret: TRetorno): Boolean;
var
  vMens : TStringList;
  vArquivo2 : String;

begin
  try
    result := true; //o retorno tem sempre que ser true

    vMens := TStringList.Create;
    vMens.Add(sConfigMsgEmail);

    if trim(sConfigSenhaEmail) = '' then
      exit;

    if trim(sConfigEmail) = '' then
      exit;

    if trim(pDestino) = '' then
      exit;

    if trim(sConfigSMTP_EMAIL) = '' then
      exit;

    if sConfigPorta_SMTP <= 0 then
      exit;

    if trim(pArquivo) = '' then
      exit;

    if trim(pArquivo2) = '' then
      vArquivo2 := ''
    else
      vArquivo2 := pArquivo2;

    EnviarEmail(sConfigSMTP_EMAIL,
                IntToStr(sConfigPorta_SMTP),
                sConfigUsuario,
                sConfigSenhaEmail,
                sConfigEmail,  //DE
                pDestino,      //para
                sConfigAssuntoEmail,
                pArquivo,
                vArquivo2,
                vMens,
                sConfigUSA_SSL,
                sConfigUSA_TLS);
    ret.status.add('OK: Email enviado com sucesso!');
  except
    on e: exception do begin
//      Atencao(e.Message);
      ret.status.add('ERRO ' + e.Message);
    end;
  end;
end;

procedure TfrmAcbrNFe.EnviarEmail(const sSmtpHost,
                                        sSmtpPort,
                                        sSmtpUser,
                                        sSmtpPasswd,
                                        sFrom,
                                        sTo,
                                        sAssunto,
                                        sAttachment,
                                        sAttachment2: String;
                                        sMensagem : TStrings;
                                        SSL,
                                        TLS : Boolean;
                                        sCopias: String='');
var
  smtp: TSMTPSend;
  msg_lines: TStringList;
  m:TMimemess;
  p: TMimepart;
  CC: Tstrings;
  I : Integer;
  CorpoEmail: TStringList;
begin

  msg_lines := TStringList.Create;
  CorpoEmail := TStringList.Create;
  smtp := TSMTPSend.Create;
  m:=TMimemess.create;
  try
     p := m.AddPartMultipart('mixed', nil);
     if sMensagem <> nil then
     begin
        CorpoEmail.Text := SubstituirVariaveis(sMensagem.Text);
        m.AddPartText(CorpoEmail, p);
     end;

     if sAttachment <> '' then
       m.AddPartBinaryFromFile(sAttachment, p);
     if sAttachment2 <> '' then
       m.AddPartBinaryFromFile(sAttachment2, p);
     m.header.tolist.add(sTo);
     m.header.From := sFrom;
     m.header.subject:=SubstituirVariaveis(sAssunto);
     m.EncodeMessage;
     msg_lines.Add(m.Lines.Text);

     smtp.UserName := sSmtpUser;
     smtp.Password := sSmtpPasswd;

     smtp.TargetHost := sSmtpHost;
     smtp.TargetPort := sSmtpPort;

     smtp.FullSSL := SSL;
     smtp.AutoTLS := TLS;
     if not smtp.Login() then
       raise Exception.Create('SMTP ERROR: Login:' + smtp.EnhCodeString+sLineBreak+smtp.FullResult.Text);

     if not smtp.MailFrom(sFrom, Length(sFrom)) then
       raise Exception.Create('SMTP ERROR: MailFrom:' + smtp.EnhCodeString+sLineBreak+smtp.FullResult.Text);
     if not smtp.MailTo(sTo) then
       raise Exception.Create('SMTP ERROR: MailTo:' + smtp.EnhCodeString+sLineBreak+smtp.FullResult.Text);
     if (sCopias <> '') then
      begin
        CC:=TstringList.Create;
        CC.DelimitedText := sLineBreak;
        CC.Text := StringReplace(sCopias,';',sLineBreak,[rfReplaceAll]);
        for I := 0 to CC.Count - 1 do
        begin
           if not smtp.MailTo(CC.Strings[i]) then
              raise Exception.Create('SMTP ERROR: MailTo:' + smtp.EnhCodeString+sLineBreak+smtp.FullResult.Text);
        end;
      end;
     if not smtp.MailData(msg_lines) then
       raise Exception.Create('SMTP ERROR: MailData:' + smtp.EnhCodeString+sLineBreak+smtp.FullResult.Text);

     if not smtp.Logout() then
       raise Exception.Create('SMTP ERROR: Logout:' + smtp.EnhCodeString+sLineBreak+smtp.FullResult.Text);
  finally
     msg_lines.Free;
     CorpoEmail.Free;
     smtp.Free;
     m.free;
  end;
end;

//procedure TfrmAcbrNFe.EnviarEmailIndy(const sSmtpHost, sSmtpPort, sSmtpUser, sSmtpPasswd, sFrom, sTo, sAssunto, sAttachment, sAttachment2: String; sMensagem : TStrings; SSL, TLS : Boolean; sCopias: String='');
//var
//  IdSMTP    : TIdSMTP;
//  IdMessage : TIdMessage;
//  IdISSLOHANDLERSocket : TIdSSLIOHandlerSocket;
//begin
//  IdSMTP    := TIdSMTP.Create(Application);
//  IdMessage := TIdMessage.Create(Application);
//  IdISSLOHANDLERSocket := TIdSSLIOHandlerSocket.Create(Application);
//  try
//     IdSMTP.Host := sSmtpHost;
//     IdSMTP.Port := StrToIntDef(sSmtpPort,25);
//     IdSMTP.Username := sSmtpUser;
//     IdSMTP.Password := sSmtpPasswd;
//
//     if SSL or TLS then
//      begin
//        IdISSLOHANDLERSocket.SSLOptions.Method := sslvSSLv3;
//        if TLS and not SSL then
//           IdISSLOHANDLERSocket.SSLOptions.Method := sslvTLSv1;
//        IdISSLOHANDLERSocket.SSLOptions.Mode := sslmClient;
//        IdSMTP.AuthenticationType := atLogin;
//        IdSMTP.IOHandler := IdISSLOHANDLERSocket;
//      end
//     else
//        IdSMTP.AuthenticationType := atNone;
//
//     IdMessage.From.Address := sFrom;
//     IdMessage.Recipients.EMailAddresses := sTo;
//     if (sCopias <> '') then
//        IdMessage.CCList.EMailAddresses := sCopias;
//
//     IdMessage.Priority := mpNormal;
//     IdMessage.Subject := SubstituirVariaveis(sAssunto);
//     IdMessage.Body.Text := SubstituirVariaveis(sMensagem.Text);
//
//     if DFeUtil.NaoEstaVazio(sAttachment) then
//        TIdAttachment.create(IdMessage.MessageParts, sAttachment);
//
//     if DFeUtil.NaoEstaVazio(sAttachment2) then
//        TIdAttachment.create(IdMessage.MessageParts, sAttachment2);
//
//     try
//        IdSMTP.Connect;
//     except
//        IdSMTP.Connect;
//     end;
//
//     try
//        IdSMTP.Send(IdMessage);
//     finally
//        IdSMTP.Disconnect;
//     end;
//  finally
//    IdISSLOHANDLERSocket.Free;
//    IdMessage.Free;
//    IdSMTP.Free;
//  end;
//end;

function TfrmAcbrNFe.SubstituirVariaveis(const ATexto: String): String;
var
  TextoStr: String;
begin
  if Trim(ATexto) = '' then
    Result := ''
  else
  begin
    TextoStr := ATexto;

    if ACBrNFe1.NotasFiscais.Count > 0 then
    begin
      with ACBrNFe1.NotasFiscais.Items[0].NFe do
      begin
        TextoStr := StringReplace(TextoStr,'[EmitNome]',     Emit.xNome,   [rfReplaceAll, rfIgnoreCase]);
        TextoStr := StringReplace(TextoStr,'[EmitFantasia]', Emit.xFant,   [rfReplaceAll, rfIgnoreCase]);
        TextoStr := StringReplace(TextoStr,'[EmitCNPJCPF]',  Emit.CNPJCPF, [rfReplaceAll, rfIgnoreCase]);
        TextoStr := StringReplace(TextoStr,'[EmitIE]',       Emit.IE,      [rfReplaceAll, rfIgnoreCase]);

        TextoStr := StringReplace(TextoStr,'[DestNome]',     Dest.xNome,   [rfReplaceAll, rfIgnoreCase]);
        TextoStr := StringReplace(TextoStr,'[DestCNPJCPF]',  Dest.CNPJCPF, [rfReplaceAll, rfIgnoreCase]);
        TextoStr := StringReplace(TextoStr,'[DestIE]',       Dest.IE,      [rfReplaceAll, rfIgnoreCase]);

        TextoStr := StringReplace(TextoStr,'[ChaveNFe]',     procNFe.chNFe, [rfReplaceAll, rfIgnoreCase]);

        TextoStr := StringReplace(TextoStr,'[SerieNF]',      FormatFloat('000',           Ide.serie),         [rfReplaceAll, rfIgnoreCase]);
        TextoStr := StringReplace(TextoStr,'[NumeroNF]',     FormatFloat('000000000',     Ide.nNF),           [rfReplaceAll, rfIgnoreCase]);
        TextoStr := StringReplace(TextoStr,'[ValorNF]',      FormatFloat('0.00',          Total.ICMSTot.vNF), [rfReplaceAll, rfIgnoreCase]);
        TextoStr := StringReplace(TextoStr,'[dtEmissao]',    FormatDateTime('dd/mm/yyyy', Ide.dEmi),          [rfReplaceAll, rfIgnoreCase]);
        TextoStr := StringReplace(TextoStr,'[dtSaida]',      FormatDateTime('dd/mm/yyyy', Ide.dSaiEnt),       [rfReplaceAll, rfIgnoreCase]);
        TextoStr := StringReplace(TextoStr,'[hrSaida]',      FormatDateTime('hh:mm:ss',   Ide.hSaiEnt),       [rfReplaceAll, rfIgnoreCase]);
      end;
    end;
    Result := TextoStr;
  end;
end;

procedure TfrmAcbrNFe.ConfiguraDANFeNfce(pNFE: Boolean = FALSE);
var
  OK : Boolean;
  DanfeSimplificado : TNFeDanfeSimplificadoFortes;
begin

  if ACBrNFe1.DANFE <> ACBrNFeDANFCeFortes1 then begin
    ACBrNFe1.DANFE :=  ACBrNFeDANFCeFortes1;
    acbrnfe1.danfe.LarguraBobina        := 294;
    ACBrNFe1.DANFE.MargemInferior       := 0.8;
    ACBrNFe1.DANFE.MargemSuperior       := 0.8;
    ACBrNFe1.DANFE.MargemDireita        := 0.1;
    ACBrNFe1.DANFE.MargemEsquerda       := 0.1;
  end;

  if pNFE then begin
    DanfeSimplificado := TNFeDanfeSimplificadoFortes.Create(nil);
    ACBrNFe1.DANFE :=  DanfeSimplificado;
    ACBrNFe1.DANFE.MargemInferior       := 0.8;
    ACBrNFe1.DANFE.MargemSuperior       := 0.8;
    ACBrNFe1.DANFE.MargemDireita        := 0;
    ACBrNFe1.DANFE.MargemEsquerda       := 0;
  end;

  if ACBrNFe1.DANFE <> nil then begin
    Acbrnfe1.DANFE.imprimeNomefantasia  := true;
    ACBrNFe1.DANFE.TipoDANFE            := tiNFCe;
    ACBrNFe1.DANFE.ExibirResumoCanhoto  := True;
    ACBrNFe1.DANFE.ExpandirLogoMarca    := True;

//    ACBrNFe1.DANFE.Logo                 := 'D:\Sistemas\projetoPosto\Branches\KontPosto\Imagens\NFCE.jpg';//dm.EmpresasCAMINHO_LOGOTIPO.Text;
//    acbrnfe1.DANFE.LogoemCima           := True;
    ACBrNFe1.DANFE.Sistema              := 'Desenvolvido por: Kontrol Tecnologia Ltda. (62)3926-3040';
    ACBrNFe1.DANFE.Site                 := '';//www.kontrolsistemas.com.br';
    ACBrNFe1.DANFE.Email                := trim(dm.EmpresasEmail.Text);
    ACBrNFe1.DANFE.Fax                  := Trim(DM.EmpresasFax.Text);
    ACBrNFe1.DANFE.ImprimirDescPorc     := True;
    ACBrNFe1.DANFE.MostrarPreview       := False;
    ACBrNFe1.DANFE.Impressora           := (Printer.Printers[iImpressoraNFCe]);
    ACBrNFe1.DANFE.NumCopias            := 1;
    ACBrNFe1.DANFE.ProdutosPorPagina    := 0;
    ACBrNFe1.DANFE.PathPDF              := '';
    ACBrNFe1.DANFE.ExibirResumoCanhoto  := False;
    ACBrNFe1.DANFE.ImprimirTotalLiquido := False;
    ACBrNFe1.DANFE.FormularioContinuo   := False;
    ACBrNFe1.DANFE.MostrarStatus        := False;
    ACBrNFe1.DANFE.ExpandirLogoMarca    := False;
    ACBrNFe1.DANFE.TamanhoFonte_DemaisCampos := 10;
    ACBrNFe1.DANFE.CasasDecimais._vUnCom     := 3;
    ACBrNFe1.DANFE.CasasDecimais._qCom       := 3;
    ACBrNFe1.DANFE.ViaConsumidor             := true;
    ACBrNFe1.DANFE.ImprimirItens             := true;
    ACBrNFe1.DANFE.LarguraBobina             := 294;
  end;
end;

procedure TfrmAcbrNFe.ConfiguraDANFe;
var
  OK : Boolean;
begin
  if ACBrNFe1.DANFE <> ACBrNFeDANFeRL1 then begin
    ACBrNFe1.DANFE := ACBrNFeDANFeRL1;
  end;

//    ACBrNFe1.DANFE.LocalImpCanhoto     := 1;
  ACBrNFe1.DANFE.ExibirResumoCanhoto := True;
  ACBrNFe1.DANFE.ExpandirLogoMarca   := True;
  ACBrNFe1.DANFE.TipoDANFE           := tiRetrato;
  ACBrNFe1.DANFE.Logo                := dm.EmpresasCAMINHO_LOGOTIPO.Text;
  ACBrNFe1.DANFE.Sistema             := 'Desenvolvido por: Kontrol Tecnologia Ltda. (62)3926-3040';
  ACBrNFe1.DANFE.Site                := '';//www.kontrolsistemas.com.br';
  ACBrNFe1.DANFE.Email               := trim(dm.EmpresasEmail.Text);
  ACBrNFe1.DANFE.Fax                 := Trim(DM.EmpresasFax.Text);//(62)3926-3040';
  ACBrNFe1.DANFE.ImprimirDescPorc    := True;
  ACBrNFe1.DANFE.MostrarPreview      := True;
  ACBrNFe1.DANFE.Impressora          := (Printer.Printers[Printer.PrinterIndex]);
  ACBrNFe1.DANFE.NumCopias           := 1;
  ACBrNFe1.DANFE.ProdutosPorPagina   := 0;
  ACBrNFe1.DANFE.MargemInferior      := 0.8;
  ACBrNFe1.DANFE.MargemSuperior      := 0.8;
  ACBrNFe1.DANFE.MargemDireita       := 0.51;
  ACBrNFe1.DANFE.MargemEsquerda      := 0.6;
  ACBrNFe1.DANFE.PathPDF             := '';
  ACBrNFe1.DANFE.CasasDecimais._qCom   := 2;
  ACBrNFe1.DANFE.CasasDecimais._vUnCom := 2;
  ACBrNFe1.DANFE.ExibirResumoCanhoto   := False;
  ACBrNFe1.DANFE.ImprimirTotalLiquido  := False;
  ACBrNFe1.DANFE.FormularioContinuo    := False;
  ACBrNFe1.DANFE.MostrarStatus         := False;
  ACBrNFe1.DANFE.ExpandirLogoMarca     := False;
  ACBrNFe1.DANFE.TamanhoFonte_DemaisCampos := 10;
  ACBrNFe1.DANFE.CasasDecimais._vUnCom     := 3;
  ACBrNFe1.DANFE.CasasDecimais._qCom       := 3;
//    ACBrNFeDANFERave1.EspessuraBorda := 1;
//    ACBrNFeDANFERave1.RavFile := 'C:\ACBrNFeMonitor\Report\DANFE_Rave513.rav';

end;

function TfrmAcbrNFe.ConsultaNfeDest(pIndicadorNFE,
                                     pIndicadorEmissor,
                                     pUltimoNSU : string;
                                     var ret    : TRetorno): Boolean;
var
  j, i: integer;
  ok : Boolean;
begin
  try
    Result := ACBrNFe1.ConsultaNFeDest(sEmpresaCNPJ,
                                       StrToIndicadorNFe(ok,pIndicadorNFE),
                                       StrToIndicadorEmissor(ok,pIndicadorEmissor),
                                       pUltimoNSU);
    ret.status.Add('');
    ret.status.Add('versao='   +ACBrNFe1.WebServices.ConsNFeDest.retConsNFeDest.versao);
    ret.status.Add('tpAmb='    +TpAmbToStr(ACBrNFe1.WebServices.ConsNFeDest.retConsNFeDest.tpAmb));
    ret.status.Add('verAplic=' +ACBrNFe1.WebServices.ConsNFeDest.retConsNFeDest.verAplic);
    ret.status.Add('cStat='    +IntToStr(ACBrNFe1.WebServices.ConsNFeDest.retConsNFeDest.cStat));
    ret.status.Add('xMotivo='  +ACBrNFe1.WebServices.ConsNFeDest.retConsNFeDest.xMotivo);
    ret.status.Add('dhResp='   +DateTimeToStr(ACBrNFe1.WebServices.ConsNFeDest.retConsNFeDest.dhResp));
    ret.status.Add('indCont='  +IndicadorContinuacaoToStr(ACBrNFe1.WebServices.ConsNFeDest.retConsNFeDest.indCont));
    ret.status.Add('ultNSU='   +ACBrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ultNSU);

    J := 1;
    for I:= 0 to AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Count-1 do begin
      if Trim(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.chNFe) <> '' then begin
        ret.status.Add('');
        ret.status.Add('[RESNFE'+Trim(IntToStrZero(J,3))+']');
        ret.status.Add('NSU='     +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.NSU);
        ret.status.Add('chNFe='   +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.chNFe);
        ret.status.Add('CNPJ='    +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.CNPJCPF);
        ret.status.Add('xNome='   +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.xNome);
        ret.status.Add('IE='      +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.IE);
        ret.status.Add('dEmi='    +DateTimeToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.dEmi));
        ret.status.Add('tpNF='    +tpNFToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.tpNF));
        ret.status.Add('vNF='     +FloatToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.vNF));
        ret.status.Add('digVal='  +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.digVal);
        ret.status.Add('dhRecbto='+DateTimeToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.dhRecbto));
        ret.status.Add('cSitNFe=' +SituacaoDFeToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.cSitNFe));
        ret.status.Add('cSitConf='+SituacaoManifDestToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resNFe.cSitConf));
        J := J + 1;
      end;
    end;

    J := 1;
    for i := 0 to AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Count -1 do begin
      if Trim(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.chNFe) <> '' then begin
        ret.status.Add('');
        ret.status.Add('[RESCANC'+Trim(IntToStrZero(J,3))+']');
        ret.status.Add('NSU='     +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.NSU);
        ret.status.Add('chNFe='   +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.chNFe);
        ret.status.Add('CNPJ='    +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.CNPJCPF);
        ret.status.Add('xNome='   +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.xNome);
        ret.status.Add('IE='      +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.IE);
        ret.status.Add('dEmi='    +DateTimeToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.dEmi));
        ret.status.Add('tpNF='    +tpNFToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.tpNF));
        ret.status.Add('vNF='     +FloatToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.vNF));
        ret.status.Add('digVal='  +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.digVal);
        ret.status.Add('dhRecbto='+DateTimeToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.dhRecbto));
        ret.status.Add('cSitNFe=' +SituacaoDFeToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.cSitNFe));
        ret.status.Add('cSitConf='+SituacaoManifDestToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCanc.cSitConf));
        J := J + 1;
      end;
    end;

    J := 1;
    for i := 0 to AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Count -1 do begin
      if Trim(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCCe.chNFe) <> '' then begin
        ret.status.Add('');
        ret.status.Add('[RESCCE'+Trim(IntToStrZero(J,3))+']');
        ret.status.Add('NSU='       +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCCe.NSU);
        ret.status.Add('chNFe='     +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCCe.chNFe);
        ret.status.Add('dhEvento='  +DateTimeToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCCe.dhEvento));
        ret.status.Add('tpEvento='  +TpEventoToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCCe.tpEvento));
        ret.status.Add('nSeqEvento='+IntToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCCe.nSeqEvento));
        ret.status.Add('descEvento='+AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCCe.descEvento);
        ret.status.Add('xCorrecao=' +AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCCe.xCorrecao);
        ret.status.Add('tpNF='      +tpNFToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCCe.tpNF));
        ret.status.Add('dhRecbto='  +DateTimeToStr(AcbrNFe1.WebServices.ConsNFeDest.retConsNFeDest.ret.Items[i].resCCe.dhRecbto));
        J := J + 1;
      end;
    end;
  except
    on E: Exception do begin
      Result := false;
      if Pos('Rejeicao: Consumo Indevido (Deve ser aguardado 1 hora para efetuar nova solicitacao se retornado cStat=137 e indCont=0, tente apos 1 hora)', e.Message) > 0 then
        ret.status.Add('ERRO: Consumo Indevido (Deve ser aguardado 1 hora para efetuar nova solicitacao, tente apos 1 hora)')
      else
        ret.status.add('ERRO: Não foi possível fazer a consulta! Tente Novamente!');
    end;
  end;
end;

function TfrmAcbrNFe.IntToStrZero(vr,qtd:Integer): string;
var
  sVr   : string;
  valor : Double;
  xtr   : string;
  i     : integer;
begin

  valor := vr;
  xtr   := '';

  for I := 1 to qtd do
    xtr := xtr + '0';

  Result := FormatFloat(xtr,valor);
end;

function TfrmAcbrNFe.EnviarManifestacao(pListaNFe : Tlist;
                                        pTpEvento : string;
                                        pJust     : string;
                                        cOrgao    : integer;
                                        var ret   : TRetorno): Boolean;
var
  i : integer;
  tpEvento : TpcnTpEvento;
  descEvento : String;
  TipoEvento :  integer;
begin
  try
    result := False;
    ACBrNFe1.EventoNFe.Evento.Clear;

    ACBrNFe1.EventoNFe.idLote := 1;
    ACBrNFe1.EventoNFe.Evento.Clear;

    TipoEvento := StrToInt(Copy(pTpEvento,5,1));
    case TipoEvento of
      0 : begin tpEvento := teManifDestConfirmacao;      descEvento := 'Confirmacao da Operacao'; end;
      1 : begin tpEvento := teManifDestCiencia;          descEvento := 'Ciencia da Operacao'; end;
      2 : begin tpEvento := teManifDestDesconhecimento;  descEvento := 'Desconhecimento da Operacao'; end;
      4 : begin tpEvento := teManifDestOperNaoRealizada; descEvento := 'Operacao nao Realizada'; end;
    End;

//    for I := 0 to pListaNFe.Count -1 do begin
      with ACBrNFe1.EventoNFe.Evento.Add do begin
//        InfEvento.id     := 'ID' + tpEvento +  TManifestacao(pListaNFe.Items[i]).chNFe + 1;
        infEvento.cOrgao := 91;//cOrgao;
        infEvento.CNPJ   := sEmpresaCNPJ;// TManifestacao(pListaNFe.Items[i]).CNPJ;
        infEvento.chNFe  := TManifestacao(pListaNFe).chNFe;
        infEvento.dhEvento := now;//StrToDate(TManifestacao(pListaNFe.Items[i]).dEmi);
        infEvento.tpEvento := tpEvento;
        infEvento.nSeqEvento   := 1;
        infEvento.versaoEvento := '1.00';
        infEvento.detEvento.descEvento := descEvento;
        if tpEvento = teManifDestOperNaoRealizada then
          InfEvento.detEvento.xJust := Trim(pJust);
      end;
//    end;

    ACBrNFe1.EnviarEvento(ACBrNFe1.EventoNFe.idLote);
//    ACBrNFe1.EnviarEventoNFe(ACBrNFe1.EventoNFe.idLote);

    ret.status.Add('idLote='   +IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.idLote));
    ret.status.Add('tpAmb='    +TpAmbToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.tpAmb));
    ret.status.Add('verAplic=' +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.verAplic);
    ret.status.Add('cOrgao='   +IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.cOrgao));
    ret.status.Add('cStat='    +IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.cStat));
    ret.status.Add('xMotivo='  +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.xMotivo);

    for I:= 0 to ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Count-1 do
    begin
      ret.status.Add('[EVENTO'+Trim(IntToStr(I+1))+']');
      ret.status.Add('id='        +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.Id);
      ret.status.Add('tpAmb='     +TpAmbToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.tpAmb));
      ret.status.Add('verAplic='  +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.verAplic);
      ret.status.Add('cOrgao='    +IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.cOrgao));
      ret.status.Add('cStat='     +IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.cStat));
      ret.status.Add('xMotivo='   +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.xMotivo);
      ret.status.Add('chNFe='     +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.chNFe);
      ret.status.Add('tpEvento='  +TpEventoToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.tpEvento));
      ret.status.Add('xEvento='   +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.xEvento);
      ret.status.Add('nSeqEvento='+IntToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.nSeqEvento));
      ret.status.Add('CNPJDest='  +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.CNPJDest);
      ret.status.Add('emailDest=' +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.emailDest);
      ret.status.Add('dhRegEvento='+DateTimeToStr(ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.dhRegEvento));
      ret.status.Add('nProt='     +ACBrNFe1.WebServices.EnvEvento.EventoRetorno.retEvento.Items[I].RetInfEvento.nProt);
    end;
    result := True;
  except
    on e: Exception do begin
      ShowMessage(e.Message+ sLineBreak + 'Não foi possivel fazer a manifestação! Tente novamente!');
    end;
  end;
end;

function TfrmAcbrNFe.LerXML(pArquivo: AnsiString): Boolean;
begin
  try
    result := false;
    if FileExists(pArquivo) then begin
      ACBrNFe1.NotasFiscais.Clear;
      ACBrNFe1.NotasFiscais.LoadFromFile(pArquivo,False);
      ConfiguraDANFe;
      result := true;
    end;
  except
    on e: Exception do begin
      result := false;
      ShowMessage('Ocorreu um erro ao ler o XML!' + sLineBreak +
                  e.Message);
    end;
  end;
end;

//le pelo xml
function TfrmAcbrNFe.LeXML(pXML: AnsiString): Boolean;
var
  vXml : String;
begin
  try
    result := false;
    vXml := pXML;
    ACBrNFe1.NotasFiscais.Clear;
    acbrnfe1.notasfiscais.LoadFromString(vXML);
    ConfiguraDANFe;
    result := true;
  except
    on e: Exception do begin
      result := false;
      ShowMessage('Ocorreu um erro ao ler o XML!' + sLineBreak +
                  e.Message);
    end;
  end;
end;

function TfrmAcbrNFe.GetChaveNFE(index: Integer = 0): AnsiString;
begin
  try
    result := somentenumeros(ACBrNFe1.NotasFiscais.Items[index].NFe.infNFe.ID);
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetDestCNPJ(index: Integer = 0): AnsiString;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Dest.CNPJCPF;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetDestNOME(index: Integer = 0): AnsiString;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Dest.xNome;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetNF(index: Integer = 0): AnsiString;
begin
  try
    result := IntToStr(ACBrNFe1.NotasFiscais.Items[index].NFe.Ide.cNF);
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetNumeroNota(index: Integer = 0): AnsiString;
begin
  try
    result := Copy(somentenumeros(ACBrNFe1.NotasFiscais.Items[index].NFe.infNFe.ID),26,9);
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetEmissao(index: Integer = 0): TDateTime;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Ide.dEmi;
  except
    Result := Now;
  end;
end;

function TfrmAcbrNFe.GetModelo(index: Integer = 0): Integer;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Ide.modelo;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetSerie(index: Integer = 0): Integer;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Ide.serie;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetEmitcMun(index: Integer = 0): String;
begin
  try
    result := IntToStr(ACBrNFe1.NotasFiscais.Items[index].NFe.Emit.EnderEmit.cMun);
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetEmitCNPJ(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Emit.CNPJCPF;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetEmitNome(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Emit.xNome;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetEmitLgr(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Emit.EnderEmit.xLgr;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetEmitBairro(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Emit.EnderEmit.xBairro;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetEmitMun(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Emit.EnderEmit.xMun;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetEmitUF(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Emit.EnderEmit.UF;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetEmitCEP(index: Integer = 0): String;
begin
  try
    result := IntToStr(ACBrNFe1.NotasFiscais.Items[index].NFe.Emit.EnderEmit.CEP);
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetEmitFone(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Emit.EnderEmit.fone;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetEmitIE(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Emit.IE;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetTranspCNPJ(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Transp.Transporta.CNPJCPF;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetTranspMun(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Transp.Transporta.xMun;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetTranspNome(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Transp.Transporta.xNome;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetTranspIE(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Transp.Transporta.IE;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetTranspEndereco(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Transp.Transporta.xEnder;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetTranspUF(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Transp.Transporta.UF;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetTranspVeicPlaca(index: Integer = 0): String;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Transp.veicTransp.placa
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetNfeValorICMS(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vICMS;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetNfeValorProdutos(index: Integer = 0): double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vProd;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetNfeValorNF(index: Integer = 0): double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vNF;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetNfeValorPIS(index: Integer = 0): double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vPIS;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetNfeValorOUTRO(index: Integer = 0): double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vOutro;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetNfeValorFRETE(index: Integer = 0): double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vFrete;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetNfeValorSeguro(index: Integer = 0): double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vSeg;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetNfeValorDesconto(index: Integer = 0): double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vDesc;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetNfeValorBC(index: Integer = 0): double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vBC;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetNfeValorICMSSUBS(index: Integer = 0): double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vST;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetNfeValorBCICMSSUBS(index: Integer = 0): double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Total.ICMSTot.vBCST;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetFaturaCount(index: Integer = 0): integer;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[index].NFe.Cobr.Dup.Count;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetFaturaDataVenc(index: Integer = 0): TDateTime;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Cobr.Dup.Items[index].dVenc;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetFaturaValor(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Cobr.Dup.Items[index].vDup;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetFinalidade(index: Integer): Integer;
VAR
  vFinalidade : TpcnFinalidadeNFe;
begin
  try
    vFinalidade := ACBrNFe1.NotasFiscais.Items[0].NFe.Ide.finNFe;
    case vFinalidade of
      fnNormal       : result := 1;
      fnComplementar : result := 2;
      fnAjuste       : result := 3;
      fnDevolucao    : result := 4;
    end;
  except
    Result := 1;
  end;
end;

function TfrmAcbrNFe.GetDetCount: Integer;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.count;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetValorProduto(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.vProd;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProdVrDesc(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.vDesc;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProdEAN(index: Integer = 0): string;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.cEAN;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetProdCod(index: Integer = 0): string;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.cProd;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetProdDescricao(index: Integer = 0): string;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.xProd;
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetProdImpCST(index: Integer = 0): string;
var
  vCST : TpcnCSTIcms;
begin
  try
    vCST := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.ICMS.CST;
    Result := EnumeradoToStr(vCST,[ '00',  '10',  '20',  '30',  '40',  '41',  '50',  '51',  '60',  '70',  '80',  '81',  '90',      '10',      '90',     '41',           '90', 'SN'],
                                  [cst00, cst10, cst20, cst30, cst40, cst41, cst50, cst51, cst60, cst70, cst80, cst81, cst90, cstPart10, cstPart90, cstRep41, cstICMSOutraUF, cstICMSSN]);

  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetProdIcmsAliq(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.ICMS.pICMS;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProdIcmsBC(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.ICMS.vBC;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProdIcmsBCST(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.ICMS.vBCST;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vICMSST(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.ICMS.vICMSST;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vICMSSTRet(index: Integer): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.ICMS.vICMSSTRet;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_pICMSST(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.ICMS.pICMSST;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_pPIS(index: Integer): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.PIS.pPIS;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vPIS(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.PIS.vPIS;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_pCOFINS(index: Integer): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.COFINS.pCOFINS;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vCOFINS(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.COFINS.vCOFINS;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vICMS(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.ICMS.vICMS;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_ICMSorig(index: Integer = 0): Double;
var
  vOrig : TpcnOrigemMercadoria;
begin
  try
    vOrig  := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.ICMS.orig;
    result := EnumeradoToStr(vOrig, ['0','1','2','3','4','5','6','7','8'],
                                    [oeNacional, oeEstrangeiraImportacaoDireta, oeEstrangeiraAdquiridaBrasil,
                                     oeNacionalConteudoImportacaoSuperior40, oeNacionalProcessosBasicos,
                                     oeNacionalConteudoImportacaoInferiorIgual40,
                                     oeEstrangeiraImportacaoDiretaSemSimilar, oeEstrangeiraAdquiridaBrasilSemSimilar,
                                     oeNacionalConteudoImportacaoSuperior70]);

  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_qCom(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.qCom;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vProd(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.vProd;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vUnCom(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.vUnCom;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vFrete(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.vFrete;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vSeg(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.vSeg;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vBCSTRet(index: Integer): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.ICMS.vBCSTRet;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vDesc(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.vDesc;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vOutro(index: Integer = 0): Double;
begin
  try
    result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.vOutro;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_IPITrib_CST(index: Integer = 0): string;
var
  vCST_IPI : TpcnCstIpi;
begin
  try
    vCST_IPI := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.IPI.CST;
    Result := EnumeradoToStr(vCST_IPI,['00' , '49' , '50' , '99' , '01' , '02' , '03' , '04' , '05' , '51' , '52' , '53' , '54' , '55'],
                                     [ipi00, ipi49, ipi50, ipi99, ipi01, ipi02, ipi03, ipi04, ipi05, ipi51, ipi52, ipi53, ipi54, ipi55]);
  except
    Result := '';
  end;
end;

function TfrmAcbrNFe.GetProd_vIPI(index: Integer = 0): Double;
begin
  try
    Result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.IPI.vIPI;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_vIPIBC(index: Integer = 0): Double;
begin
  try
    Result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.IPI.vBC;
  except
    Result := 0;
  end;
end;

function TfrmAcbrNFe.GetProd_AliqIPI(index: Integer = 0): Double;
begin
  try
    Result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Imposto.IPI.pIPI;
  except
    Result := 0;
  end;
end;



function TfrmAcbrNFe.GetProd_CFOP(Index: Integer): string;
begin
  try
    Result := ACBrNFe1.NotasFiscais.Items[0].NFe.Det.Items[index].Prod.CFOP;
  except
    Result := '';
  end;
end;


function TfrmAcbrNFe.ImprimeRelatorio(pTexto: TStringList): boolean;
begin
  try
//    ACBrNFe1.Configuracoes.Geral.IdCSC := '1';
//    ACBrNFe1.Configuracoes.Geral.CSC   := '6f6e3f83c4ff3554';

    ACBrPosPrinter1.Modelo        := ppEscBematech;
    ACBrPosPrinter1.Device.Porta  := 'COM5';
    ACBrPosPrinter1.Device.Baud   := 9600;
    ACBrPosPrinter1.IgnorarTags   := True;
    ACBrPosPrinter1.ControlePorta := True; // True faz com que o componente abra e feche a porta conforme a necessidade automaticamente

    ACBrNFeDANFeESCPOS1.ImprimeEmUmaLinha     := false;
    ACBrNFeDANFeESCPOS1.ImprimeDescAcrescItem := false;

    ACBrNFeDANFeESCPOS1.ImprimirRelatorio(pTexto, 1,True,false);
    result := true;
  except
    on e: exception do begin
      result := false;
      ShowMessage(e.message);
    end;
  end;

end;

procedure TfrmAcbrNFe.GerarIniNFe(AStr: String);
var
  I, J, K : Integer;
  versao, sSecao, sFim, sCodPro, sNumeroDI, sNumeroADI, sQtdVol,
  sNumDup, sCampoAdic, sTipo, sDia, sDeduc, sNVE : String;
  INIRec : TMemIniFile ;
  SL     : TStringList;
  OK     : boolean;
//  vPath  : String;
begin
  INIRec := TMemIniFile.create( 'nfe.ini' ) ;
  SL := TStringList.Create;

  if FileExists(Astr) then
    SL.LoadFromFile(AStr)
  else
    Sl.Text := ConvertStrRecived( Astr );

//  vPath := ExtractFilePath(Application.ExeName) + 'arquivos\INI\' +
//               FormatDateTime('YYYY',Now) + '\' +
//               FormatDateTime('MM',Now) + '\' +
//               FormatDateTime('DD',Now) + '\';
//  if not DirectoryExists(vPath) then
//    ForceDirectories(vPath);
//  vPath := vPath + 'ini.txt';
//  sl.savetofile(vPath);


  INIRec.SetStrings( SL );
  SL.Free ;

 with frmAcbrNFe do begin
   try
      ACBrNFe1.NotasFiscais.Clear;
      with ACBrNFe1.NotasFiscais.Add.NFe do begin
         versao        :=                   INIRec.ReadString('infNFe','versao', VersaoDFToStr(ACBrNFe1.Configuracoes.Geral.VersaoDF));
         infNFe.versao := StringToFloatDef( INIRec.ReadString('infNFe','versao', VersaoDFToStr(ACBrNFe1.Configuracoes.Geral.VersaoDF)),0) ;

         versao := infNFe.VersaoStr;
         versao := StringReplace(versao,'versao="','',[rfReplaceAll,rfIgnoreCase]);
         versao := StringReplace(versao,'"','',[rfReplaceAll,rfIgnoreCase]);

         Ide.cNF        := INIRec.ReadInteger( 'Identificacao','Codigo' ,INIRec.ReadInteger( 'Identificacao','cNF' ,0));
         Ide.natOp      := INIRec.ReadString(  'Identificacao','NaturezaOperacao' ,INIRec.ReadString(  'Identificacao','natOp' ,''));
         Ide.indPag     := StrToIndpag(OK,INIRec.ReadString( 'Identificacao','FormaPag',INIRec.ReadString( 'Identificacao','indPag','0')));
         Ide.modelo     := INIRec.ReadInteger( 'Identificacao','Modelo' ,INIRec.ReadInteger( 'Identificacao','mod' ,55));
         ACBrNFe1.Configuracoes.Geral.ModeloDF := StrToModeloDF(OK,IntToStr(Ide.modelo));
         ACBrNFe1.Configuracoes.Geral.VersaoDF := StrToVersaoDF(OK,versao);
         Ide.serie      := INIRec.ReadInteger( 'Identificacao','Serie'  ,1);
         Ide.nNF        := INIRec.ReadInteger( 'Identificacao','Numero' ,INIRec.ReadInteger( 'Identificacao','nNF' ,0));
         Ide.dEmi       := StringToDateTime(INIRec.ReadString( 'Identificacao','Emissao',INIRec.ReadString( 'Identificacao','dEmi',INIRec.ReadString( 'Identificacao','dhEmi',FormatDateTimeBr(Now)))));
         Ide.dSaiEnt    := StringToDateTime(INIRec.ReadString( 'Identificacao','Saida'  ,INIRec.ReadString( 'Identificacao','dSaiEnt'  ,INIRec.ReadString( 'Identificacao','dhSaiEnt','0'))));
         Ide.hSaiEnt    := StringToDateTime(INIRec.ReadString( 'Identificacao','hSaiEnt','0'));  //NFe2
         Ide.tpNF       := StrToTpNF(OK,INIRec.ReadString( 'Identificacao','Tipo',INIRec.ReadString( 'Identificacao','tpNF','1')));

         Ide.idDest     := StrToDestinoOperacao(OK,INIRec.ReadString( 'Identificacao','idDest','1'));

         Ide.tpImp      := StrToTpImp(  OK, INIRec.ReadString( 'Identificacao','tpImp',TpImpToStr(ACBrNFe1.DANFE.TipoDANFE)));  //NFe2
         Ide.tpEmis     := StrToTpEmis( OK,INIRec.ReadString( 'Identificacao','tpEmis',IntToStr(ACBrNFe1.Configuracoes.Geral.FormaEmissaoCodigo)));
//         Ide.cDV
//         Ide.tpAmb
         Ide.finNFe     := StrToFinNFe( OK,INIRec.ReadString( 'Identificacao','Finalidade',INIRec.ReadString( 'Identificacao','finNFe','0')));
         Ide.indFinal   := StrToConsumidorFinal(OK,INIRec.ReadString( 'Identificacao','indFinal','0'));
         Ide.indPres    := StrToPresencaComprador(OK,INIRec.ReadString( 'Identificacao','indPres','0'));

         Ide.procEmi    := StrToProcEmi(OK,INIRec.ReadString( 'Identificacao','procEmi','0')); //NFe2
         Ide.verProc    := INIRec.ReadString(  'Identificacao','verProc' ,'Kontrol Sistemas' );
         Ide.dhCont     := StringToDateTime(INIRec.ReadString( 'Identificacao','dhCont'  ,'0')); //NFe2
         Ide.xJust      := INIRec.ReadString(  'Identificacao','xJust' ,'' ); //NFe2

         I := 1 ;
         while true do begin
            sSecao := 'NFRef'+IntToStrZero(I,3) ;
            sFim   := INIRec.ReadString(  sSecao,'Tipo'  ,'FIM');
            sTipo := UpperCase(INIRec.ReadString(  sSecao,'Tipo'  ,'NFe')); //NFe2 NF NFe NFP CTe ECF)
            if (sFim = 'FIM') or (Length(sFim) <= 0) then begin
              if INIRec.ReadString(sSecao,'refNFe','') <> '' then
                sTipo := 'NFE';
              break ;
            end;

            with Ide.NFref.Add do begin
              if sTipo = 'NFE' then
                refNFe :=  INIRec.ReadString(sSecao,'refNFe','');
            end;
            Inc(I);
         end;

         Emit.CNPJCPF           := INIRec.ReadString(  'Emitente','CNPJ'    ,INIRec.ReadString(  'Emitente','CNPJCPF'    ,''));
         Emit.xNome             := INIRec.ReadString(  'Emitente','Razao'   ,INIRec.ReadString(  'Emitente','xNome'   ,''));
         Emit.xFant             := INIRec.ReadString(  'Emitente','Fantasia',INIRec.ReadString(  'Emitente','xFant',''));
         Emit.IE                := INIRec.ReadString(  'Emitente','IE'      ,'');
         Emit.IEST              := INIRec.ReadString(  'Emitente','IEST','');
         Emit.IM                := INIRec.ReadString(  'Emitente','IM'  ,'');
         Emit.CNAE              := INIRec.ReadString(  'Emitente','CNAE','');
         Emit.CRT               := StrToCRT(ok, INIRec.ReadString(  'Emitente','CRT','3')); //NFe2

         Emit.EnderEmit.xLgr    := INIRec.ReadString(  'Emitente','Logradouro' ,INIRec.ReadString(  'Emitente','xLgr' ,''));
         if (INIRec.ReadString(  'Emitente','Numero'     ,'') <> '') or (INIRec.ReadString(  'Emitente','nro'     ,'') <> '') then
            Emit.EnderEmit.nro     := INIRec.ReadString(  'Emitente','Numero'     ,INIRec.ReadString(  'Emitente','nro'     ,''));
         if (INIRec.ReadString(  'Emitente','Complemento'     ,'') <> '') or (INIRec.ReadString(  'Emitente','xCpl'     ,'') <> '') then
            Emit.EnderEmit.xCpl    := INIRec.ReadString(  'Emitente','Complemento',INIRec.ReadString(  'Emitente','xCpl'     ,''));
         Emit.EnderEmit.xBairro := INIRec.ReadString(  'Emitente','Bairro'     ,INIRec.ReadString(  'Emitente','xBairro',''));
         Emit.EnderEmit.cMun    := INIRec.ReadInteger( 'Emitente','CidadeCod'  ,INIRec.ReadInteger( 'Emitente','cMun'   ,0));
         Emit.EnderEmit.xMun    := INIRec.ReadString(  'Emitente','Cidade'     ,INIRec.ReadString(  'Emitente','xMun'   ,''));
         Emit.EnderEmit.UF      := INIRec.ReadString(  'Emitente','UF'         ,'');
         Emit.EnderEmit.CEP     := INIRec.ReadInteger( 'Emitente','CEP'  ,0);
//         if Emit.EnderEmit.cMun <= 0 then
//            Emit.EnderEmit.cMun := ObterCodigoMunicipio(Emit.EnderEmit.xMun,Emit.EnderEmit.UF);
         Emit.EnderEmit.cPais   := INIRec.ReadInteger( 'Emitente','PaisCod'    ,INIRec.ReadInteger( 'Emitente','cPais'    ,1058));
         Emit.EnderEmit.xPais   := INIRec.ReadString(  'Emitente','Pais'       ,INIRec.ReadString(  'Emitente','xPais'    ,'BRASIL'));
         Emit.EnderEmit.fone    := INIRec.ReadString(  'Emitente','Fone' ,'');

         Ide.cUF       := INIRec.ReadInteger( 'Identificacao','cUF'       ,UFparaCodigo(Emit.EnderEmit.UF));
         Ide.cMunFG    := INIRec.ReadInteger( 'Identificacao','CidadeCod' ,INIRec.ReadInteger( 'Identificacao','cMunFG' ,Emit.EnderEmit.cMun));

         if INIRec.ReadString(  'Avulsa','CNPJ','') <> '' then
          begin
            Avulsa.CNPJ    := INIRec.ReadString(  'Avulsa','CNPJ','');
            Avulsa.xOrgao  := INIRec.ReadString(  'Avulsa','xOrgao','');
            Avulsa.matr    := INIRec.ReadString(  'Avulsa','matr','');
            Avulsa.xAgente := INIRec.ReadString(  'Avulsa','xAgente','');
            Avulsa.fone    := INIRec.ReadString(  'Avulsa','fone','');
            Avulsa.UF      := INIRec.ReadString(  'Avulsa','UF','');
            Avulsa.nDAR    := INIRec.ReadString(  'Avulsa','nDAR','');
            Avulsa.dEmi    := StringToDateTime(INIRec.ReadString(  'Avulsa','dEmi','0'));
            Avulsa.vDAR    := StringToFloatDef(INIRec.ReadString(  'Avulsa','vDAR',''),0);
            Avulsa.repEmi  := INIRec.ReadString(  'Avulsa','repEmi','');
            Avulsa.dPag    := StringToDateTime(INIRec.ReadString(  'Avulsa','dPag','0'));
          end;

         Dest.idEstrangeiro     := INIRec.ReadString(  'Destinatario','idEstrangeiro','');
         Dest.CNPJCPF           := INIRec.ReadString(  'Destinatario','CNPJ'       ,INIRec.ReadString(  'Destinatario','CNPJCPF',INIRec.ReadString(  'Destinatario','CPF','')));
         Dest.xNome             := INIRec.ReadString(  'Destinatario','NomeRazao'  ,INIRec.ReadString(  'Destinatario','xNome'  ,''));
         Dest.indIEDest         := StrToindIEDest(OK,INIRec.ReadString( 'Destinatario','indIEDest','1'));
         Dest.IE                := INIRec.ReadString(  'Destinatario','IE'         ,'');
         Dest.ISUF              := INIRec.ReadString(  'Destinatario','ISUF'       ,'');
         Dest.Email             := INIRec.ReadString(  'Destinatario','Email'      ,'');  //NFe2

         Dest.EnderDest.xLgr    := INIRec.ReadString(  'Destinatario','Logradouro' ,INIRec.ReadString(  'Destinatario','xLgr' ,''));
         if (INIRec.ReadString('Destinatario','Numero','') <> '') or (INIRec.ReadString('Destinatario','nro','') <> '') then
            Dest.EnderDest.nro     := INIRec.ReadString(  'Destinatario','Numero'     ,INIRec.ReadString('Destinatario','nro',''));
         if (INIRec.ReadString('Destinatario','Complemento','') <> '') or (INIRec.ReadString('Destinatario','xCpl','') <> '') then
            Dest.EnderDest.xCpl    := INIRec.ReadString(  'Destinatario','Complemento',INIRec.ReadString('Destinatario','xCpl',''));
         Dest.EnderDest.xBairro := INIRec.ReadString(  'Destinatario','Bairro'     ,INIRec.ReadString(  'Destinatario','xBairro',''));
         Dest.EnderDest.cMun    := INIRec.ReadInteger( 'Destinatario','CidadeCod'  ,INIRec.ReadInteger( 'Destinatario','cMun'   ,0));
         Dest.EnderDest.xMun    := INIRec.ReadString(  'Destinatario','Cidade'     ,INIRec.ReadString(  'Destinatario','xMun'   ,''));
         Dest.EnderDest.UF      := INIRec.ReadString(  'Destinatario','UF'         ,'');
         Dest.EnderDest.CEP     := INIRec.ReadInteger( 'Destinatario','CEP'       ,0);
//         if Dest.EnderDest.cMun <= 0 then
//            Dest.EnderDest.cMun := ObterCodigoMunicipio(Dest.EnderDest.xMun,Dest.EnderDest.UF);
         Dest.EnderDest.cPais   := INIRec.ReadInteger( 'Destinatario','PaisCod'    ,INIRec.ReadInteger('Destinatario','cPais',1058));
         Dest.EnderDest.xPais   := INIRec.ReadString(  'Destinatario','Pais'       ,INIRec.ReadString( 'Destinatario','xPais','BRASIL'));
         Dest.EnderDest.Fone    := INIRec.ReadString(  'Destinatario','Fone'       ,'');

         I := 1 ;

         while true do begin
            sSecao    := 'Produto'+IntToStrZero(I,3) ;
            sCodPro   := INIRec.ReadString(sSecao,'Codigo',INIRec.ReadString( sSecao,'cProd','FIM')) ;
            if sCodPro = 'FIM' then
               break ;

            with Det.Add do begin
               Prod.nItem := I;
               infAdProd      := INIRec.ReadString(sSecao,'infAdProd','');

               Prod.cProd    := INIRec.ReadString( sSecao,'Codigo'   ,INIRec.ReadString( sSecao,'cProd'   ,''));
               if (Length(INIRec.ReadString( sSecao,'EAN','')) > 0) or (Length(INIRec.ReadString( sSecao,'cEAN','')) > 0)  then
                  Prod.cEAN      := INIRec.ReadString( sSecao,'EAN'      ,INIRec.ReadString( sSecao,'cEAN'      ,''));
               Prod.xProd    := INIRec.ReadString( sSecao,'Descricao',INIRec.ReadString( sSecao,'xProd',''));
               Prod.NCM      := INIRec.ReadString( sSecao,'NCM'      ,'');
               Prod.CEST     := INIRec.ReadString( sSecao,'CEST'      ,'');
               Prod.EXTIPI   := INIRec.ReadString( sSecao,'EXTIPI'      ,'');
               Prod.CFOP     := INIRec.ReadString( sSecao,'CFOP'     ,'');
               Prod.uCom     := INIRec.ReadString( sSecao,'Unidade'  ,INIRec.ReadString( sSecao,'uCom'  ,''));
               Prod.qCom     := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'   ,INIRec.ReadString(sSecao,'qCom'  ,'')) ,0) ;
               Prod.vUnCom   := StringToFloatDef( INIRec.ReadString(sSecao,'ValorUnitario',INIRec.ReadString(sSecao,'vUnCom','')) ,0) ;
               Prod.vProd    := StringToFloatDef( INIRec.ReadString(sSecao,'ValorTotal'   ,INIRec.ReadString(sSecao,'vProd' ,'')) ,0) ;

               if Length(INIRec.ReadString( sSecao,'cEANTrib','')) > 0 then
                  Prod.cEANTrib      := INIRec.ReadString( sSecao,'cEANTrib'      ,'');
               Prod.uTrib     := INIRec.ReadString( sSecao,'uTrib'  , Prod.uCom);
               Prod.qTrib     := StringToFloatDef( INIRec.ReadString(sSecao,'qTrib'  ,''), Prod.qCom);
               Prod.vUnTrib   := StringToFloatDef( INIRec.ReadString(sSecao,'vUnTrib','') ,Prod.vUnCom) ;

               Prod.vFrete    := StringToFloatDef( INIRec.ReadString(sSecao,'vFrete','') ,0) ;
               Prod.vSeg      := StringToFloatDef( INIRec.ReadString(sSecao,'vSeg','') ,0) ;
               Prod.vDesc     := StringToFloatDef( INIRec.ReadString(sSecao,'ValorDesconto',INIRec.ReadString(sSecao,'vDesc','')) ,0) ;
               Prod.vOutro    := StringToFloatDef( INIRec.ReadString(sSecao,'vOutro','') ,0) ; //NFe2
               Prod.IndTot    := StrToindTot(OK,INIRec.ReadString(sSecao,'indTot','1'));       //NFe2

               Prod.xPed      := INIRec.ReadString( sSecao,'xPed'    , '');  //NFe2
               Prod.nItemPed  := INIRec.ReadString( sSecao,'nItemPed', '');  //NFe2

               Prod.nFCI      := INIRec.ReadString( sSecao,'nFCI','');  //NFe3
               Prod.nRECOPI   := INIRec.ReadString( sSecao,'nRECOPI','');  //NFe3

               pDevol := StringToFloatDef( INIRec.ReadString(sSecao,'pDevol','') ,0);
               vIPIDevol := StringToFloatDef( INIRec.ReadString(sSecao,'vIPIDevol','') ,0);

               Imposto.vTotTrib := StringToFloatDef( INIRec.ReadString(sSecao,'vTotTrib','') ,0) ; //NFe2

               J := 1 ;
               while true do
                begin
                  sSecao  := 'NVE'+IntToStrZero(I,3)+IntToStrZero(J,3) ;
                  sNVE    := INIRec.ReadString(sSecao,'NVE','') ;
                  if (sNVE <> '') then
                     Prod.NVE.Add.NVE := sNVE
                  else
                     Break;
                  Inc(J);
                end;

               J := 1 ;
               while true do
                begin
                  sSecao      := 'DI'+IntToStrZero(I,3)+IntToStrZero(J,3) ;
                  sNumeroDI := INIRec.ReadString(sSecao,'NumeroDI',INIRec.ReadString(sSecao,'nDi','')) ;

                  if sNumeroDI <> '' then
                   begin
                     with Prod.DI.Add do
                      begin
                        nDi         := sNumeroDI;
                        dDi         := StringToDateTime(INIRec.ReadString(sSecao,'DataRegistroDI'  ,INIRec.ReadString(sSecao,'dDi'  ,'0')));
                        xLocDesemb  := INIRec.ReadString(sSecao,'LocalDesembaraco',INIRec.ReadString(sSecao,'xLocDesemb',''));
                        UFDesemb    := INIRec.ReadString(sSecao,'UFDesembaraco'   ,INIRec.ReadString(sSecao,'UFDesemb'   ,''));
                        dDesemb     := StringToDateTime(INIRec.ReadString(sSecao,'DataDesembaraco',INIRec.ReadString(sSecao,'dDesemb','0')));

                        tpViaTransp  := StrToTipoViaTransp(OK,INIRec.ReadString(sSecao,'tpViaTransp',''));
                        vAFRMM       := StringToFloatDef( INIRec.ReadString(sSecao,'vAFRMM','') ,0) ;
                        tpIntermedio := StrToTipoIntermedio(OK,INIRec.ReadString(sSecao,'tpIntermedio',''));
                        CNPJ         := INIRec.ReadString(sSecao,'CNPJ','');
                        UFTerceiro   := INIRec.ReadString(sSecao,'UFTerceiro','');

                        cExportador := INIRec.ReadString(sSecao,'CodigoExportador',INIRec.ReadString(sSecao,'cExportador',''));

                        K := 1 ;
                        while true do
                         begin
                           sSecao      := 'LADI'+IntToStrZero(I,3)+IntToStrZero(J,3)+IntToStrZero(K,3)  ;
                           sNumeroADI := INIRec.ReadString(sSecao,'NumeroAdicao',INIRec.ReadString(sSecao,'nAdicao','FIM')) ;
                           if (sNumeroADI = 'FIM') or (Length(sNumeroADI) <= 0) then
                              break;

                           with adi.Add do
                            begin
                              nAdicao     := StrToInt(sNumeroADI);
                              nSeqAdi     := INIRec.ReadInteger( sSecao,'nSeqAdi',K);
                              cFabricante := INIRec.ReadString(  sSecao,'CodigoFabricante',INIRec.ReadString(  sSecao,'cFabricante',''));
                              vDescDI     := StringToFloatDef( INIRec.ReadString(sSecao,'DescontoADI',INIRec.ReadString(sSecao,'vDescDI','')) ,0);
                              nDraw       := INIRec.ReadString( sSecao,'nDraw','');
                            end;
                           Inc(K)
                         end;
                      end;
                   end
                  else
                    Break;
                  Inc(J);
                end;

               J := 1 ;
               while true do
                begin
                  sSecao  := 'detExport'+IntToStrZero(I,3)+IntToStrZero(J,3) ;
                  sFim    := INIRec.ReadString(sSecao,'nDraw',INIRec.ReadString(sSecao,'nRE','FIM')) ;
                  if (sFim = 'FIM') or (Length(sFim) <= 0) then
                     break ;

                  with Prod.detExport.Add do
                   begin
                     nDraw       := INIRec.ReadString( sSecao,'nDraw','');
                     nRE         := INIRec.ReadString( sSecao,'nRE','');
                     chNFe       := INIRec.ReadString( sSecao,'chNFe','');
                     qExport     := StringToFloatDef( INIRec.ReadString(sSecao,'qExport','') ,0);
                   end;
                  Inc(J);
                end;

              sSecao := 'impostoDevol'+IntToStrZero(I,3) ;
              sFim   := INIRec.ReadString( sSecao,'pDevol','FIM') ;
              if (sFim <> 'FIM') then
               begin
                 pDevol := StringToFloatDef( INIRec.ReadString(sSecao,'pDevol','') ,0);
                 vIPIDevol := StringToFloatDef( INIRec.ReadString(sSecao,'vIPIDevol','') ,0);
               end;

               sSecao := 'Combustivel'+IntToStrZero(I,3) ;
               sFim   := INIRec.ReadString( sSecao,'cProdANP','FIM') ;
               if (sFim <> 'FIM') then begin
                 with Prod.comb do begin
                    cProdANP := INIRec.ReadInteger( sSecao,'cProdANP',0) ;
                    pMixGN   := StringToFloatDef(INIRec.ReadString( sSecao,'pMixGN',''),0) ;
                    CODIF    := INIRec.ReadString(  sSecao,'CODIF'   ,'') ;
                    qTemp    := StringToFloatDef(INIRec.ReadString( sSecao,'qTemp',''),0) ;
                    UFcons   := INIRec.ReadString( sSecao,'UFCons','') ;

                    sSecao := 'CIDE'+IntToStrZero(I,3) ;
                    CIDE.qBCprod   := StringToFloatDef(INIRec.ReadString( sSecao,'qBCprod'  ,''),0) ;
                    CIDE.vAliqProd := StringToFloatDef(INIRec.ReadString( sSecao,'vAliqProd',''),0) ;
                    CIDE.vCIDE     := StringToFloatDef(INIRec.ReadString( sSecao,'vCIDE'    ,''),0) ;

                    sSecao := 'encerrante'+IntToStrZero(I,3) ;
                    encerrante.nBico    := INIRec.ReadInteger( sSecao,'nBico'  ,0) ;
                    encerrante.nBomba   := INIRec.ReadInteger( sSecao,'nBomba' ,0) ;
                    encerrante.nTanque  := INIRec.ReadInteger( sSecao,'nTanque',0) ;
                    encerrante.vEncIni  := INIRec.ReadFloat( sSecao,'vEncIni',0) ;
                    encerrante.vEncFin  := INIRec.ReadFloat( sSecao,'vEncFin',0) ;

                    sSecao := 'ICMSComb'+IntToStrZero(I,3) ;
                    ICMS.vBCICMS   := StringToFloatDef(INIRec.ReadString( sSecao,'vBCICMS'  ,''),0) ;
                    ICMS.vICMS     := StringToFloatDef(INIRec.ReadString( sSecao,'vICMS'    ,''),0) ;
                    ICMS.vBCICMSST := StringToFloatDef(INIRec.ReadString( sSecao,'vBCICMSST',''),0) ;
                    ICMS.vICMSST   := StringToFloatDef(INIRec.ReadString( sSecao,'vICMSST'  ,''),0) ;

                    sSecao := 'ICMSInter'+IntToStrZero(I,3) ;
                    sFim   := INIRec.ReadString( sSecao,'vBCICMSSTDest','FIM') ;
                    if (sFim <> 'FIM') then
                     begin
                       ICMSInter.vBCICMSSTDest := StringToFloatDef(sFim,0) ;
                       ICMSInter.vICMSSTDest   := StringToFloatDef(INIRec.ReadString( sSecao,'vICMSSTDest',''),0) ;
                     end;

                    sSecao := 'ICMSCons'+IntToStrZero(I,3) ;
                    sFim   := INIRec.ReadString( sSecao,'vBCICMSSTCons','FIM') ;
                    if (sFim <> 'FIM') then
                     begin
                       ICMSCons.vBCICMSSTCons := StringToFloatDef(sFim,0) ;
                       ICMSCons.vICMSSTCons   := StringToFloatDef(INIRec.ReadString( sSecao,'vICMSSTCons',''),0) ;
                       ICMSCons.UFcons        := INIRec.ReadString( sSecao,'UFCons','') ;
                     end;
                 end;
               end;

               with Imposto do
                begin
                   sSecao := 'ICMS'+IntToStrZero(I,3) ;
                   sFim   := INIRec.ReadString( sSecao,'CST',INIRec.ReadString(sSecao,'CSOSN','FIM')) ;
                   if (sFim <> 'FIM') then
                    begin
                      with ICMS do
                      begin
                        ICMS.orig       := StrToOrig(     OK, INIRec.ReadString(sSecao,'Origem'    ,INIRec.ReadString(sSecao,'orig'    ,'0' ) ));
                        CST             := StrToCSTICMS(  OK, INIRec.ReadString(sSecao,'CST'       ,'00'));
                        CSOSN           := StrToCSOSNIcms(OK, INIRec.ReadString(sSecao,'CSOSN'     ,''  ));     //NFe2
                        ICMS.modBC      := StrTomodBC(    OK, INIRec.ReadString(sSecao,'Modalidade',INIRec.ReadString(sSecao,'modBC','0' ) ));
                        ICMS.pRedBC     := StringToFloatDef( INIRec.ReadString(sSecao,'PercentualReducao',INIRec.ReadString(sSecao,'pRedBC','')) ,0);
                        ICMS.vBC        := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase',INIRec.ReadString(sSecao,'vBC'  ,'')) ,0);
                        ICMS.pICMS      := StringToFloatDef( INIRec.ReadString(sSecao,'Aliquota' ,INIRec.ReadString(sSecao,'pICMS','')) ,0);
                        ICMS.vICMS      := StringToFloatDef( INIRec.ReadString(sSecao,'Valor'    ,INIRec.ReadString(sSecao,'vICMS','')) ,0);
                        ICMS.modBCST    := StrTomodBCST(OK, INIRec.ReadString(sSecao,'ModalidadeST',INIRec.ReadString(sSecao,'modBCST','0')));
                        ICMS.pMVAST     := StringToFloatDef( INIRec.ReadString(sSecao,'PercentualMargemST' ,INIRec.ReadString(sSecao,'pMVAST' ,'')) ,0);
                        ICMS.pRedBCST   := StringToFloatDef( INIRec.ReadString(sSecao,'PercentualReducaoST',INIRec.ReadString(sSecao,'pRedBCST','')) ,0);
                        ICMS.vBCST      := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBaseST',INIRec.ReadString(sSecao,'vBCST','')) ,0);
                        ICMS.pICMSST    := StringToFloatDef( INIRec.ReadString(sSecao,'AliquotaST' ,INIRec.ReadString(sSecao,'pICMSST' ,'')) ,0);
                        ICMS.vICMSST    := StringToFloatDef( INIRec.ReadString(sSecao,'ValorST'    ,INIRec.ReadString(sSecao,'vICMSST'    ,'')) ,0);
                        ICMS.UFST       := INIRec.ReadString(sSecao,'UFST'    ,'');                           //NFe2
                        ICMS.pBCOp      := StringToFloatDef( INIRec.ReadString(sSecao,'pBCOp'    ,'') ,0);    //NFe2
                        ICMS.vBCSTRet   := StringToFloatDef( INIRec.ReadString(sSecao,'vBCSTRet','') ,0);     //NFe2
                        ICMS.vICMSSTRet := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSSTRet','') ,0);   //NFe2
                        ICMS.motDesICMS := StrTomotDesICMS(OK, INIRec.ReadString(sSecao,'motDesICMS','0'));   //NFe2
                        ICMS.pCredSN    := StringToFloatDef( INIRec.ReadString(sSecao,'pCredSN','') ,0);      //NFe2
                        ICMS.vCredICMSSN:= StringToFloatDef( INIRec.ReadString(sSecao,'vCredICMSSN','') ,0);  //NFe2
                        ICMS.vBCSTDest  := StringToFloatDef( INIRec.ReadString(sSecao,'vBCSTDest','') ,0);    //NFe2
                        ICMS.vICMSSTDest:= StringToFloatDef( INIRec.ReadString(sSecao,'vICMSSTDest','') ,0);   //NFe2
                        ICMS.vICMSDeson := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSDeson','') ,0);
                        ICMS.vICMSOp    := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSOp','') ,0);
                        ICMS.pDif       := StringToFloatDef( INIRec.ReadString(sSecao,'pDif','') ,0);
                        ICMS.vICMSDif   := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSDif','') ,0);
                      end;
                    end;

                   sSecao := 'ICMSUFDEST'+IntToStrZero(I,3) ;
                   sFim   := INIRec.ReadString( sSecao,'vBCUFDest','FIM') ;
                   if (sFim <> 'FIM') then
                    begin
                      with ICMSUFDest do
                      begin
                        vBCUFDest      := StringToFloatDef( INIRec.ReadString(sSecao,'vBCUFDest','') ,0);
                        pICMSUFDest    := StringToFloatDef( INIRec.ReadString(sSecao,'pICMSUFDest','') ,0);
                        pICMSInter     := StringToFloatDef( INIRec.ReadString(sSecao,'pICMSInter','') ,0);
                        pICMSInterPart := StringToFloatDef( INIRec.ReadString(sSecao,'pICMSInterPart','') ,0);
                        vICMSUFDest    := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSUFDest','') ,0);
                        vICMSUFRemet   := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSUFRemet','') ,0);
                        pFCPUFDest     := StringToFloatDef( INIRec.ReadString(sSecao,'pFCPUFDest','') ,0);
                        vFCPUFDest     := StringToFloatDef( INIRec.ReadString(sSecao,'vFCPUFDest','') ,0);
                      end;
                    end;

                   sSecao := 'IPI'+IntToStrZero(I,3) ;
                   sFim  := INIRec.ReadString( sSecao,'CST','FIM') ;
                   if (sFim <> 'FIM') then
                    begin
                     with IPI do
                      begin
                        CST      := StrToCSTIPI(OK, INIRec.ReadString( sSecao,'CST','')) ;
                        clEnq    := INIRec.ReadString(  sSecao,'ClasseEnquadramento',INIRec.ReadString(  sSecao,'clEnq'   ,''));
                        CNPJProd := INIRec.ReadString(  sSecao,'CNPJProdutor'       ,INIRec.ReadString(  sSecao,'CNPJProd',''));
                        cSelo    := INIRec.ReadString(  sSecao,'CodigoSeloIPI'      ,INIRec.ReadString(  sSecao,'cSelo'   ,''));
                        qSelo    := INIRec.ReadInteger( sSecao,'QuantidadeSelos'    ,INIRec.ReadInteger( sSecao,'qSelo'   ,0));
                        cEnq     := INIRec.ReadString(  sSecao,'CodigoEnquadramento',INIRec.ReadString(  sSecao,'cEnq'    ,''));

                        vBC    := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'   ,INIRec.ReadString(sSecao,'vBC'   ,'')) ,0);
                        qUnid  := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'  ,INIRec.ReadString(sSecao,'qUnid' ,'')) ,0);
                        vUnid  := StringToFloatDef( INIRec.ReadString(sSecao,'ValorUnidade',INIRec.ReadString(sSecao,'vUnid' ,'')) ,0);
                        pIPI   := StringToFloatDef( INIRec.ReadString(sSecao,'Aliquota'    ,INIRec.ReadString(sSecao,'pIPI'  ,'')) ,0);
                        vIPI   := StringToFloatDef( INIRec.ReadString(sSecao,'Valor'       ,INIRec.ReadString(sSecao,'vIPI'  ,'')) ,0);
                      end;
                    end;

                   sSecao   := 'II'+IntToStrZero(I,3) ;
                   sFim   := INIRec.ReadString( sSecao,'ValorBase',INIRec.ReadString( sSecao,'vBC','FIM')) ;
                   if (sFim <> 'FIM') then
                    begin
                     with II do
                      begin
                        vBc      := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'          ,INIRec.ReadString(sSecao,'vBC'     ,'')) ,0);
                        vDespAdu := StringToFloatDef( INIRec.ReadString(sSecao,'ValorDespAduaneiras',INIRec.ReadString(sSecao,'vDespAdu','')) ,0);
                        vII      := StringToFloatDef( INIRec.ReadString(sSecao,'ValorII'            ,INIRec.ReadString(sSecao,'vII'     ,'')) ,0);
                        vIOF     := StringToFloatDef( INIRec.ReadString(sSecao,'ValorIOF'           ,INIRec.ReadString(sSecao,'vIOF'    ,'')) ,0);
                      end;
                    end;

                   sSecao    := 'PIS'+IntToStrZero(I,3) ;
                   sFim   := INIRec.ReadString( sSecao,'CST','FIM') ;
                   if (sFim <> 'FIM') then
                    begin
                     with PIS do
                       begin
                        CST :=  StrToCSTPIS(OK, INIRec.ReadString( sSecao,'CST','01'));

                        PIS.vBC       := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'    ,INIRec.ReadString(sSecao,'vBC'      ,'')) ,0);
                        PIS.pPIS      := StringToFloatDef( INIRec.ReadString(sSecao,'Aliquota'     ,INIRec.ReadString(sSecao,'pPIS'     ,'')) ,0);
                        PIS.qBCProd   := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'   ,INIRec.ReadString(sSecao,'qBCProd'  ,'')) ,0);
                        PIS.vAliqProd := StringToFloatDef( INIRec.ReadString(sSecao,'ValorAliquota',INIRec.ReadString(sSecao,'vAliqProd','')) ,0);
                        PIS.vPIS      := StringToFloatDef( INIRec.ReadString(sSecao,'Valor'        ,INIRec.ReadString(sSecao,'vPIS'     ,'')) ,0);
                       end;
                    end;

                   sSecao    := 'PISST'+IntToStrZero(I,3) ;
                   sFim   := INIRec.ReadString( sSecao,'ValorBase','F')+ INIRec.ReadString( sSecao,'Quantidade','IM') ;
                   if (sFim = 'FIM') then
                      sFim   := INIRec.ReadString( sSecao,'vBC','F')+ INIRec.ReadString( sSecao,'qBCProd','IM') ;

                   if (sFim <> 'FIM') then
                    begin
                     with PISST do
                      begin
                        vBc       := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'    ,INIRec.ReadString(sSecao,'vBC'      ,'')) ,0);
                        pPis      := StringToFloatDef( INIRec.ReadString(sSecao,'AliquotaPerc' ,INIRec.ReadString(sSecao,'pPis'     ,'')) ,0);
                        qBCProd   := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'   ,INIRec.ReadString(sSecao,'qBCProd'  ,'')) ,0);
                        vAliqProd := StringToFloatDef( INIRec.ReadString(sSecao,'AliquotaValor',INIRec.ReadString(sSecao,'vAliqProd','')) ,0);
                        vPIS      := StringToFloatDef( INIRec.ReadString(sSecao,'ValorPISST'   ,INIRec.ReadString(sSecao,'vPIS'     ,'')) ,0);
                      end;
                    end;

                   sSecao    := 'COFINS'+IntToStrZero(I,3) ;
                   sFim   := INIRec.ReadString( sSecao,'CST','FIM') ;
                   if (sFim <> 'FIM') then
                    begin
                     with COFINS do
                      begin
                        CST := StrToCSTCOFINS(OK, INIRec.ReadString( sSecao,'CST','01'));

                        COFINS.vBC       := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'    ,INIRec.ReadString(sSecao,'vBC'      ,'')) ,0);
                        COFINS.pCOFINS   := StringToFloatDef( INIRec.ReadString(sSecao,'Aliquota'     ,INIRec.ReadString(sSecao,'pCOFINS'  ,'')) ,0);
                        COFINS.qBCProd   := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'   ,INIRec.ReadString(sSecao,'qBCProd'  ,'')) ,0);
                        COFINS.vAliqProd := StringToFloatDef( INIRec.ReadString(sSecao,'ValorAliquota',INIRec.ReadString(sSecao,'vAliqProd','')) ,0);
                        COFINS.vCOFINS   := StringToFloatDef( INIRec.ReadString(sSecao,'Valor'        ,INIRec.ReadString(sSecao,'vCOFINS'  ,'')) ,0);
                      end;
                    end;

                   sSecao    := 'COFINSST'+IntToStrZero(I,3) ;
                   sFim   := INIRec.ReadString( sSecao,'ValorBase','F')+ INIRec.ReadString( sSecao,'Quantidade','IM');
                   if (sFim = 'FIM') then
                      sFim   := INIRec.ReadString( sSecao,'vBC','F')+ INIRec.ReadString( sSecao,'qBCProd','IM') ;

                   if (sFim <> 'FIM') then
                    begin
                     with COFINSST do
                      begin
                         vBC       := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'    ,INIRec.ReadString(sSecao,'vBC'      ,'')) ,0);
                         pCOFINS   := StringToFloatDef( INIRec.ReadString(sSecao,'AliquotaPerc' ,INIRec.ReadString(sSecao,'pCOFINS'  ,'')) ,0);
                         qBCProd   := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'   ,INIRec.ReadString(sSecao,'qBCProd'  ,'')) ,0);
                         vAliqProd := StringToFloatDef( INIRec.ReadString(sSecao,'AliquotaValor',INIRec.ReadString(sSecao,'vAliqProd','')) ,0);
                         vCOFINS   := StringToFloatDef( INIRec.ReadString(sSecao,'ValorCOFINSST',INIRec.ReadString(sSecao,'vCOFINS'  ,'')) ,0);
                       end;
                    end;

                   sSecao    := 'ISSQN'+IntToStrZero(I,3) ;
                   sFim   := INIRec.ReadString( sSecao,'ValorBase',INIRec.ReadString(sSecao,'vBC'   ,'FIM')) ;
                   if (sFim = 'FIM') then
                      sFim   := INIRec.ReadString( sSecao,'vBC','FIM');
                   if (sFim <> 'FIM') then
                    begin
                     with ISSQN do
                      begin
                        if StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase',INIRec.ReadString(sSecao,'vBC','')) ,0) > 0 then
                         begin
                           vBC       := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'   ,INIRec.ReadString(sSecao,'vBC'   ,'')) ,0);
                           vAliq     := StringToFloatDef( INIRec.ReadString(sSecao,'Aliquota'    ,INIRec.ReadString(sSecao,'vAliq' ,'')) ,0);
                           vISSQN    := StringToFloatDef( INIRec.ReadString(sSecao,'ValorISSQN'  ,INIRec.ReadString(sSecao,'vISSQN','')) ,0);
                           cMunFG    := INIRec.ReadInteger(sSecao,'MunicipioFatoGerador',INIRec.ReadInteger(sSecao,'cMunFG',0));
                           cListServ := INIRec.ReadString(sSecao,'CodigoServico',INIRec.ReadString(sSecao,'cListServ',''));
                           cSitTrib  := StrToISSQNcSitTrib( OK,INIRec.ReadString(sSecao,'cSitTrib','')) ;
                           vDeducao    := StringToFloatDef( INIRec.ReadString(sSecao,'ValorDeducao'   ,INIRec.ReadString(sSecao,'vDeducao'   ,'')) ,0);
                           vOutro      := StringToFloatDef( INIRec.ReadString(sSecao,'ValorOutro'   ,INIRec.ReadString(sSecao,'vOutro'   ,'')) ,0);
                           vDescIncond := StringToFloatDef( INIRec.ReadString(sSecao,'ValorDescontoIncondicional'   ,INIRec.ReadString(sSecao,'vDescIncond'   ,'')) ,0);
                           vDescCond   := StringToFloatDef( INIRec.ReadString(sSecao,'vDescontoCondicional'   ,INIRec.ReadString(sSecao,'vDescCond'   ,'')) ,0);
                           vISSRet     := StringToFloatDef( INIRec.ReadString(sSecao,'ValorISSRetido'   ,INIRec.ReadString(sSecao,'vISSRet'   ,'')) ,0);
                           indISS      := StrToindISS( OK,INIRec.ReadString(sSecao,'indISS','')) ;
                           cServico    := INIRec.ReadString(sSecao,'cServico','');
                           cMun        := INIRec.ReadInteger(sSecao,'cMun',0);
                           cPais       := INIRec.ReadInteger(sSecao,'cPais',1058);
                           nProcesso   := INIRec.ReadString(sSecao,'nProcesso','');
                           indIncentivo := StrToindIncentivo( OK,INIRec.ReadString(sSecao,'indIncentivo','')) ;
                         end;
                      end;
                    end;
                end;

             end;
            Inc( I ) ;
          end ;

         Total.ICMSTot.vBC     := StringToFloatDef( INIRec.ReadString('Total','BaseICMS'     ,INIRec.ReadString('Total','vBC'     ,'')) ,0) ;
         Total.ICMSTot.vICMS   := StringToFloatDef( INIRec.ReadString('Total','ValorICMS'    ,INIRec.ReadString('Total','vICMS'   ,'')) ,0) ;
         Total.ICMSTot.vICMSDeson := StringToFloatDef( INIRec.ReadString('Total','vICMSDeson',''),0) ;
         Total.ICMSTot.vICMSUFDest := StringToFloatDef( INIRec.ReadString('Total','vICMSUFDest',''),0) ;
         Total.ICMSTot.vICMSUFRemet := StringToFloatDef( INIRec.ReadString('Total','vICMSUFRemet',''),0) ;
         Total.ICMSTot.vFCPUFDest :=  StringToFloatDef( INIRec.ReadString('Total','vFCPUFDest',''),0) ;
         Total.ICMSTot.vBCST   := StringToFloatDef( INIRec.ReadString('Total','BaseICMSSubstituicao' ,INIRec.ReadString('Total','vBCST','')) ,0) ;
         Total.ICMSTot.vST     := StringToFloatDef( INIRec.ReadString('Total','ValorICMSSubstituicao',INIRec.ReadString('Total','vST'  ,'')) ,0) ;
         Total.ICMSTot.vProd   := StringToFloatDef( INIRec.ReadString('Total','ValorProduto' ,INIRec.ReadString('Total','vProd'  ,'')) ,0) ;
         Total.ICMSTot.vFrete  := StringToFloatDef( INIRec.ReadString('Total','ValorFrete'   ,INIRec.ReadString('Total','vFrete' ,'')) ,0) ;
         Total.ICMSTot.vSeg    := StringToFloatDef( INIRec.ReadString('Total','ValorSeguro'  ,INIRec.ReadString('Total','vSeg'   ,'')) ,0) ;
         Total.ICMSTot.vDesc   := StringToFloatDef( INIRec.ReadString('Total','ValorDesconto',INIRec.ReadString('Total','vDesc'  ,'')) ,0) ;
         Total.ICMSTot.vII     := StringToFloatDef( INIRec.ReadString('Total','ValorII'      ,INIRec.ReadString('Total','vII'    ,'')) ,0) ;
         Total.ICMSTot.vIPI    := StringToFloatDef( INIRec.ReadString('Total','ValorIPI'     ,INIRec.ReadString('Total','vIPI'   ,'')) ,0) ;
         Total.ICMSTot.vPIS    := StringToFloatDef( INIRec.ReadString('Total','ValorPIS'     ,INIRec.ReadString('Total','vPIS'   ,'')) ,0) ;
         Total.ICMSTot.vCOFINS := StringToFloatDef( INIRec.ReadString('Total','ValorCOFINS'  ,INIRec.ReadString('Total','vCOFINS','')) ,0) ;
         Total.ICMSTot.vOutro  := StringToFloatDef( INIRec.ReadString('Total','ValorOutrasDespesas',INIRec.ReadString('Total','vOutro','')) ,0) ;
         Total.ICMSTot.vNF     := StringToFloatDef( INIRec.ReadString('Total','ValorNota'    ,INIRec.ReadString('Total','vNF'    ,'')) ,0) ;
         Total.ICMSTot.vTotTrib:= StringToFloatDef( INIRec.ReadString('Total','vTotTrib'     ,''),0) ;

         Total.ISSQNtot.vServ  := StringToFloatDef( INIRec.ReadString('Total','ValorServicos',INIRec.ReadString('ISSQNtot','vServ','')) ,0) ;
         Total.ISSQNTot.vBC    := StringToFloatDef( INIRec.ReadString('Total','ValorBaseISS' ,INIRec.ReadString('ISSQNtot','vBC'  ,'')) ,0) ;
         Total.ISSQNTot.vISS   := StringToFloatDef( INIRec.ReadString('Total','ValorISSQN'   ,INIRec.ReadString('ISSQNtot','vISS' ,'')) ,0) ;
         Total.ISSQNTot.vPIS   := StringToFloatDef( INIRec.ReadString('Total','ValorPISISS'  ,INIRec.ReadString('ISSQNtot','vPIS' ,'')) ,0) ;
         Total.ISSQNTot.vCOFINS := StringToFloatDef( INIRec.ReadString('Total','ValorCONFINSISS',INIRec.ReadString('ISSQNtot','vCOFINS','')) ,0) ;
         Total.ISSQNtot.dCompet     := StringToDateTime(INIRec.ReadString('ISSQNtot','dCompet','0'));
         Total.ISSQNtot.vDeducao    := StringToFloatDef( INIRec.ReadString('ISSQNtot','vDeducao'   ,'') ,0) ;
         Total.ISSQNtot.vOutro      := StringToFloatDef( INIRec.ReadString('ISSQNtot','vOutro'   ,'') ,0) ;
         Total.ISSQNtot.vDescIncond := StringToFloatDef( INIRec.ReadString('ISSQNtot','vDescIncond'   ,'') ,0) ;
         Total.ISSQNtot.vDescCond   := StringToFloatDef( INIRec.ReadString('ISSQNtot','vDescCond'   ,'') ,0) ;
         Total.ISSQNtot.vISSRet     := StringToFloatDef( INIRec.ReadString('ISSQNtot','vISSRet'   ,'') ,0) ;
         Total.ISSQNtot.cRegTrib    := StrToRegTribISSQN( OK,INIRec.ReadString('ISSQNtot','cRegTrib','1')) ;

         Total.retTrib.vRetPIS    := StringToFloatDef( INIRec.ReadString('retTrib','vRetPIS'   ,'') ,0) ;
         Total.retTrib.vRetCOFINS := StringToFloatDef( INIRec.ReadString('retTrib','vRetCOFINS','') ,0) ;
         Total.retTrib.vRetCSLL   := StringToFloatDef( INIRec.ReadString('retTrib','vRetCSLL'  ,'') ,0) ;
         Total.retTrib.vBCIRRF    := StringToFloatDef( INIRec.ReadString('retTrib','vBCIRRF'   ,'') ,0) ;
         Total.retTrib.vIRRF      := StringToFloatDef( INIRec.ReadString('retTrib','vIRRF'     ,'') ,0) ;
         Total.retTrib.vBCRetPrev := StringToFloatDef( INIRec.ReadString('retTrib','vBCRetPrev','') ,0) ;
         Total.retTrib.vRetPrev   := StringToFloatDef( INIRec.ReadString('retTrib','vRetPrev'  ,'') ,0) ;

         Transp.modFrete := StrTomodFrete(OK, INIRec.ReadString('Transportador','FretePorConta',INIRec.ReadString('Transportador','modFrete','0')));
         Transp.Transporta.CNPJCPF  := INIRec.ReadString('Transportador','CNPJCPF'  ,'');
         Transp.Transporta.xNome    := INIRec.ReadString('Transportador','NomeRazao',INIRec.ReadString('Transportador','xNome',''));
         Transp.Transporta.IE       := INIRec.ReadString('Transportador','IE'       ,'');
         Transp.Transporta.xEnder   := INIRec.ReadString('Transportador','Endereco' ,INIRec.ReadString('Transportador','xEnder',''));
         Transp.Transporta.xMun     := INIRec.ReadString('Transportador','Cidade'   ,INIRec.ReadString('Transportador','xMun',''));
         Transp.Transporta.UF       := INIRec.ReadString('Transportador','UF'       ,'');

         Transp.retTransp.vServ    := StringToFloatDef( INIRec.ReadString('Transportador','ValorServico',INIRec.ReadString('Transportador','vServ'   ,'')) ,0) ;
         Transp.retTransp.vBCRet   := StringToFloatDef( INIRec.ReadString('Transportador','ValorBase'   ,INIRec.ReadString('Transportador','vBCRet'  ,'')) ,0) ;
         Transp.retTransp.pICMSRet := StringToFloatDef( INIRec.ReadString('Transportador','Aliquota'    ,INIRec.ReadString('Transportador','pICMSRet','')) ,0) ;
         Transp.retTransp.vICMSRet := StringToFloatDef( INIRec.ReadString('Transportador','Valor'       ,INIRec.ReadString('Transportador','vICMSRet','')) ,0) ;
         Transp.retTransp.CFOP     := INIRec.ReadString('Transportador','CFOP'     ,'');
         Transp.retTransp.cMunFG   := INIRec.ReadInteger('Transportador','CidadeCod',INIRec.ReadInteger('Transportador','cMunFG',0));

         Transp.veicTransp.placa := INIRec.ReadString('Transportador','Placa'  ,'');
         Transp.veicTransp.UF    := INIRec.ReadString('Transportador','UFPlaca','');
         Transp.veicTransp.RNTC  := INIRec.ReadString('Transportador','RNTC'   ,'');

         Transp.vagao := INIRec.ReadString( 'Transportador','vagao','') ;
         Transp.balsa := INIRec.ReadString( 'Transportador','balsa','') ;

         Cobr.Fat.nFat  := INIRec.ReadString( 'Fatura','Numero',INIRec.ReadString( 'Fatura','nFat',''));
         Cobr.Fat.vOrig := StringToFloatDef( INIRec.ReadString('Fatura','ValorOriginal',INIRec.ReadString('Fatura','vOrig','')) ,0) ;
         Cobr.Fat.vDesc := StringToFloatDef( INIRec.ReadString('Fatura','ValorDesconto',INIRec.ReadString('Fatura','vDesc','')) ,0) ;
         Cobr.Fat.vLiq  := StringToFloatDef( INIRec.ReadString('Fatura','ValorLiquido' ,INIRec.ReadString('Fatura','vLiq' ,'')) ,0) ;

         I := 1 ;
         while true do
          begin
            sSecao    := 'Duplicata'+IntToStrZero(I,3) ;
            sNumDup   := INIRec.ReadString(sSecao,'Numero',INIRec.ReadString(sSecao,'nDup','FIM')) ;
            if (sNumDup = 'FIM') or (Length(sNumDup) <= 0) then
               break ;

            with Cobr.Dup.Add do
             begin
               nDup  := sNumDup;
               dVenc := StringToDateTime(INIRec.ReadString( sSecao,'DataVencimento',INIRec.ReadString( sSecao,'dVenc','0')));
               vDup  := StringToFloatDef( INIRec.ReadString(sSecao,'Valor',INIRec.ReadString(sSecao,'vDup','')) ,0) ;
             end;
            Inc(I);
          end;

         I := 1 ;
         while true do
          begin
            sSecao    := 'pag'+IntToStrZero(I,3) ;
            sFim      := INIRec.ReadString(sSecao,'tpag','FIM');
            if (sFim = 'FIM') or (Length(sFim) <= 0) then
               break ;

            with pag.Add do
             begin
               tPag  := StrToFormaPagamento(OK,sFim);
               vPag  := StringToFloatDef( INIRec.ReadString(sSecao,'vPag','') ,0) ;

               tpIntegra  := StrTotpIntegra(OK,INIRec.ReadString(sSecao,'tpIntegra',''));
               CNPJ  := INIRec.ReadString(sSecao,'CNPJ','');
               tBand := StrToBandeiraCartao(OK,INIRec.ReadString(sSecao,'tBand','99'));
               cAut  := INIRec.ReadString(sSecao,'cAut','');
             end;
            Inc(I);
          end;

         InfAdic.infAdFisco :=  INIRec.ReadString( 'DadosAdicionais','Fisco'      ,INIRec.ReadString( 'DadosAdicionais','infAdFisco',''));
         InfAdic.infCpl     :=  INIRec.ReadString( 'DadosAdicionais','Complemento',INIRec.ReadString( 'DadosAdicionais','infCpl'    ,''));

         I := 1 ;
         while true do
          begin
            sSecao     := 'InfAdic'+IntToStrZero(I,3) ;
            sCampoAdic := INIRec.ReadString(sSecao,'Campo',INIRec.ReadString(sSecao,'xCampo','FIM')) ;
            if (sCampoAdic = 'FIM') or (Length(sCampoAdic) <= 0) then
               break ;

            with InfAdic.obsCont.Add do
             begin
               xCampo := sCampoAdic;
               xTexto := INIRec.ReadString( sSecao,'Texto',INIRec.ReadString( sSecao,'xTexto',''));
             end;
            Inc(I);
          end;

         I := 1 ;
         while true do
          begin
            sSecao     := 'ObsFisco'+IntToStrZero(I,3) ;
            sCampoAdic := INIRec.ReadString(sSecao,'Campo',INIRec.ReadString(sSecao,'xCampo','FIM')) ;
            if (sCampoAdic = 'FIM') or (Length(sCampoAdic) <= 0) then
               break ;

            with InfAdic.obsFisco.Add do
             begin
               xCampo := sCampoAdic;
               xTexto := INIRec.ReadString( sSecao,'Texto',INIRec.ReadString( sSecao,'xTexto',''));
             end;
            Inc(I);
          end;

         I := 1 ;
         while true do
          begin
            sSecao     := 'procRef'+IntToStrZero(I,3) ;
            sCampoAdic := INIRec.ReadString(sSecao,'nProc','FIM') ;
            if (sCampoAdic = 'FIM') or (Length(sCampoAdic) <= 0) then
               break ;

            with InfAdic.procRef.Add do
             begin
               nProc := sCampoAdic;
               indProc := StrToindProc(OK,INIRec.ReadString( sSecao,'indProc','0'));
             end;
            Inc(I);
          end;

       end;
   finally
      INIRec.Free ;
   end;
 end;
end;


function TfrmAcbrNFe.UFparaCodigo(const UF: string): integer;
const
  (**)UFS = '.AC.AL.AP.AM.BA.CE.DF.ES.GO.MA.MT.MS.MG.PA.PB.PR.PE.PI.RJ.RN.RS.RO.RR.SC.SP.SE.TO.';
  CODIGOS = '.12.27.16.13.29.23.53.32.52.21.51.50.31.15.25.41.26.22.33.24.43.11.14.42.35.28.17.';
begin
  try
    result := StrToInt(copy(CODIGOS, pos('.' + UF + '.', UFS) + 1, 2));
  except
    result := 0;
  end;
end;


end.
