unit NFEInterfaceV3;

interface

uses
  Windows, DBTables, Controls, Forms, SysUtils, DB, Classes, StrUtils,
  XMLDoc,XMLIntf,xmldom, untCapaNFe, Dialogs, UnitDM,
  ADODB, ConexaoAcbrMonitorNfe, UntAcbrNFe, untManifestacaoDest,
  untSplashNFe, untCampos, untVariaveis, WinInet;


type
  TeeEnvV3 = (eeNotaV3,eeCanceladaV3);

  TLoteNFe = class
  private
    fNFE      : string;
    fCStat    : string;
    fXMotivo  : string;
    fChNFe    : string;
    fDhRecbto : string;
    fNProt    : string;
    fDigVal   : string;
    fArquivo  : string;
  published
    property  NFE      : string read fNFE       write fNFE;
    property  CStat    : string read fCStat     write fCStat;
    property  XMotivo  : string read fXMotivo   write fXMotivo;
    property  ChNFe    : string read fChNFe     write fChNFe;
    property  DhRecbto : string read fDhRecbto  write fDhRecbto;
    property  NProt    : string read fNProt     write fNProt;
    property  DigVal   : string read fDigVal    write fDigVal;
    property  Arquivo  : string read fArquivo   write fArquivo;
  end;


  TNFEInterfaceV3 = class(TObject)
  private
    KontNfe      : TCPClient;
    fVersaoNfe   : String;
    fCad_cgc     : String;
    fNumero_nota : String;
    fQn          : TADOQuery;
    fQf          : TADOQuery;
    fQnTemp      : TADOQuery;
    fMensCritica : TStringList;
    fCapaNFe     : TCapaNFe;
    fNFCE        : Boolean;
    fModelo      : String;
    fCampos      : TCampos;
    fXML         : AnsiString;
//    fTimeOut     : Integer;

    function  Salvar_XML(aEmpresaId,aNotaFiscal: string; pDir: String = ''; pNomeArq: String = ''; pModelo: String = '55'; pSerie: String = '1'): boolean;
    procedure SalvarXmlCancelamentoCliente(pDiretorio,empresaid,pedido,pstatus,xmlEnvio,xmlRetorno:String);

    function  SelectCapaNota:Boolean;
    function  SelectItemNota:Boolean;
    function CriarNFce(pRetornaXML: Integer; pCupomFiscal: AnsiString): Boolean;
    function SalvarXMLServidor(vArquivo: string; pXML: AnsiString): Boolean;
    function BuscaStatus(var pStatus: Integer; var pMensagem: string; pNFCE: Boolean = false): Boolean;
    function BuscaStatus2(var pStatus: Integer; var pMensagem: string; pNFCE: Boolean = false): Boolean;
    function GravaInutilizacaoNFe(cCNPJ, cJustificativa: String; nAno, nModelo, nSerie, nNumInicial, nNumFinal: integer;
                                  pXmlEnviado : WideString;
                                  pXmlRetorno : WideString;
                                  pNumeroProtocolo: string;
                                  pDate : TDateTime): boolean;
//    procedure DespausarAutomatico;
//    function PausarAutomatico: boolean;
    public
    procedure SetDestinatario(pDestinatarioCNPJ, pDestinatarioIE,
      pDestinatarioNomeRazao, pDestinatarioFone, pDestinatarioCEP,
      pDestinatarioLogradouro, pDestinatarioNumero, pDestinatarioComplemento,
      pDestinatarioBairro, pDestinatarioCidadeCod, pDestinatarioCidade,
      pDestinatarioUF, pDestinatarioindIEDest: string;
      pDestinatarioCodigo: Integer);
    procedure SetCFOP(pValue: String);
    procedure SetCNPJautXML(pValue: String);
    procedure SetCodigo(pValue: Integer);
    procedure SetCPFautXML(pValue: String);
    procedure SetDestinatarioBairro(pValue: String);
    procedure SetDestinatarioCEP(pValue: String);
    procedure SetDestinatarioCidade(pValue: String);
    procedure SetDestinatarioCidadeCod(pValue: String);
    procedure SetDestinatarioCNPJ(pValue: String);
    procedure SetDestinatarioCodigo(pValue: Integer);
    procedure SetDestinatarioComplemento(pValue: String);
    procedure SetDestinatarioFone(pValue: String);
    procedure SetDestinatarioIE(pValue: String);
    procedure SetDestinatarioindIEDest(pValue: String);
    procedure SetDestinatarioLogradouro(pValue: String);
    procedure SetDestinatarioNomeRazao(pValue: String);
    procedure SetDestinatarioNumero(pValue: String);
    procedure SetDestinatarioUF(pValue: String);
    procedure SetEmissao(pValue: String);
    procedure SetEmitenteBairro(pValue: String);
    procedure SetEmitenteCEP(pValue: String);
    procedure SetEmitenteCidade(pValue: String);
    procedure SetEmitenteCidadeCod(pValue: String);
    procedure SetEmitenteCNAE(pValue: String);
    procedure SetEmitenteCNPJ(pValue: String);
    procedure SetEmitenteComplemento(pValue: String);
    procedure SetEmitenteCRT(pValue: String);
    procedure SetEmitenteFantasia(pValue: String);
    procedure SetEmitenteFone(pValue: String);
    procedure SetEmitenteIE(pValue: String);
    procedure SetEmitenteIM(pValue: String);
    procedure SetEmitenteLogradouro(pValue: String);
    procedure SetEmitenteNumero(pValue: Integer);
    procedure SetEmitentePais(pValue: String);
    procedure SetEmitentePaisCod(pValue: String);
    procedure SetEmitenteRazao(pValue: String);
    procedure SetEmitenteUF(pValue: String);
    procedure SetFormaPag(pValue: String);
    procedure SetHoraSaida(pValue: String);
    procedure SetidDest(pValue: Integer);
    procedure SetindFinal(pValue: Integer);
    procedure SetindPres(pValue: Integer);
    procedure SetInfAdFisco(pValue: String);
    procedure SetinfCpl(pValue: String);
    procedure SetModelo(pValue: Integer);
    procedure SetNaturezaOperacao(pValue: String);
    procedure SetNumero(pValue: Integer);
    procedure SetrefNF(pValue: String);
    procedure SetSaida(pValue: String);
    procedure SetSerie(pValue: Integer);
    procedure SetStatus(pValue: String);
    procedure SetTipo(pValue: Integer);
    procedure SettpEmis(pValue: Integer);
    procedure SettpImp(pValue: Integer);
    procedure SetTransportadorFretePorConta(pValue: String);
    procedure SetfModelo;

    constructor Create(pVersao:String=''; pValidadeCertificado: string = ''; pNFCE : Boolean = False);
    destructor  Destroy; override;
    procedure TrataXML(var pMensagem  : AnsiString;
                           pChaveNfe  : string;
                           pDataHora  : string;
                           pProtocolo : string;
                           pdigVal    : string);
    procedure SalvaRetornoXML(pArquivo: string; pChaveNFE: string; pNumeroNota: String; pModelo: string; pSerie: string);
    procedure SetTotalValorProduto(pValue: Double);
    procedure SetTotalValorNota(pValue: Double);
    function Ativo: Boolean;
    function SetdhCont: Boolean;
    function SetxJust(pJust: String): Boolean;
    function  AssinarXML(pEnderecoXML : String): Boolean;
    function SetTpEmiss(pTpEmiss: Integer): Boolean;
    function SetFormaEmissao(pTpEmiss: Integer;  pAutomatico: Boolean = false): Boolean;
    function FileExists(pArquivo: string): Boolean;
    function EnviarEmail(cEmailDestino,cArqXML,cArqPDF: string; pSplash: boolean = True): Boolean;
    procedure NfeConsultaStatusServico;
    procedure EnviarEmailParaCliente(pTipoXml:TeeEnvV3;pMensagemErro:Boolean);
    procedure ValidarNfe(pTipoXml : TeeEnvV3); overload;
    function  ValidarNFe(pEnderecoXML : String): Boolean; overload;
    function  EnviarNFe(pEnderecoXML   : String;
                        var pChNFe     : string;
                        var pMotivo    : string;
                        var pRecibo    : string;
                        var pProtocolo : string;
                        pGerenciadorNFE: Boolean = False;
                        pModelo : string = '55';
                        pSerie  : string ='1';
                        pSplash : Boolean = true;
                        pAutomatico: Boolean = false;
                        pXML: Boolean = false): Boolean;
    function ImprimirDanfeNFe(cArquivo: string; pChaveNFE: STRING; pXML: AnsiString = ''; pPerguntaVisualizar: Boolean = True; pTipo: Integer = 0): Boolean;
    function ImprimirDanfeNFce(cArquivo: string; pNFE: Boolean = false): Boolean;
    function ImprimirDanfeNfceCancelado(cArquivo: string): Boolean;
    function ImprimirDanfePDF(pXML: AnsiString; var pEndPDF: string): boolean;
    function GravaNFe(pNFCE: Boolean = False): Boolean;
    function ApagaNFe: Boolean;
    function RetornaFinalidadeConfCFOP(pCFOP: string; var pTipo: Integer): Integer;
    procedure NfeCriarEnviarNFe(pTextoIni      : string;
                                pNumLote       : Integer;
                                pImprimirDanfe : Integer;
                                pSincrono      : Integer);
    function CriarNFe(pRetornaXML  : Integer;
                      pNumeroNf    : AnsiString;
                      var pEnderecoXML : string;
                      var pXML : AnsiString;
                      pModelo  : string = '55';
                      pSerie   : string ='1'): Boolean;

    function EnviarLoteNfe(pLote: string; var pRecibo: String; var pStatusRetorno: String; var pListaNFE : TList): boolean;
    function AdicionarNFe(pRetornaXML      : Integer;
                          pNumeroNf        : AnsiString;
                          var pEnderecoXML : string;
                          var pXML         : AnsiString;
                          pLote            : String;
                          pModelo          : string = '55';
                          pSerie           : string ='1'): Boolean;

    function CriaCapaNFe(pCapa: array of Variant): Boolean; overload;
    function CriaCapaNFe: Boolean; overload;
    function AddItensNfe(pItens: array of Variant): boolean;
    function AddFormaPagNfe(pFormaPag: array of Variant): boolean;
    function AddCupomRef(pCupomRef: array of Variant): boolean;
    function AddNFCeRef(pNFCeRef: string): boolean;
    function GetFormaPag: string;
    procedure SetFinalidade(pValue: Integer);
    function InutilizarNFe(cCNPJ, cJustificativa: string; nAno, nModelo, nSerie, nNumInicial, nNumFinal: integer; pDate: TDateTime; pSplash: boolean = true): Boolean;
    function NfeConsultaSituacaoNfe(pNumeroNfe: string;
                                    var pDataHora  : string;
                                    var pProtocolo : string;
                                    var pdigVal    : string;
                                    var pMSG       : string;
                                    var pStatus    : string;
                                    pSplash: Boolean = True;
                                    pMostraMSG : Boolean = true;
                                    pAutomatico: Boolean = false): Boolean;
    function CertificadoDataVencimento(pMostraSplash: Boolean = False; pAutomatico: boolean = false): TDateTime;
    function SetCertificado(pCertificado: string): boolean;
    procedure SetValidadeCertificado(pDias : string);
    function CancelarNFe(pChaveNFe, cJustificativa: string; pCNPJ: string; var pStatus, pProtocolo: string; pEvento : Boolean = True): Boolean;
    function CartaCorrecao(pChaveNFe  : string;
                           pCorrecao  : array of AnsiString;
                           pCNPJ,
                           dhEvento   : String;
                           nSeqEvento : array of Integer): Boolean;
    function ConsultaNotasPendendesManifesto(pIndicadorNFE, pIndicadorEmissor: string; var pUltimoNSU: string): TList;
    function EnviarManifestacao(pListaNFe : Tlist;
                                pTpEvento : string;
                                pJust     : string;
                                pAN       : Boolean;
                                var pListaMan : TList): Boolean;
    function Busca_XML(pEndereco: string; var pXML: AnsiString;pAutomatico: Boolean = false): Boolean;
    function ImprimirEvento(pPath: string): Boolean;
    function ImprimirTextoTEF(pTextoTEF: TStringList): boolean;
    procedure SetMensagem(pValue: String);
    procedure SetTotTrib(pTotTrib: String);
    procedure Validada(pVALIDADA: String; iNumeroNf : integer; pModelo, pSerie: string);
    function NfeConsultaXML(pEndereco: string; var pDataHora, pProtocolo,pdigVal, pMSG, pStatus: String): Boolean;

    function GetCFOP_ItensNFe(pCodigo : string; pItem: integer):string;
    function GetValorBaseICMS_ItensNFe(pCodigo : string; pItem: integer):Double;
    function GetAliquotaICMS_ItensNFe(pCodigo : string; pItem: integer):Double;
    function GetValorICMS_ItensNFe(pCodigo : string; pItem: integer):Double;

  published
    property Cad_cgc : String read fCad_cgc write fcad_cgc;
    property Numero_nota : String read fNumero_nota write fNumero_nota;
    property XML_NFE : AnsiString read fXML write fXML;
  end;

  procedure GravaHistoricoNfe(pKONT_NUMERO_NOTA : string;
                              pSTATUS_RETORNO   : Integer;
                              pXML_ENVIADO      : WideString;
                              pPesquisar        : Boolean = false;
                              pModelo           : string  = '55';
                              pSerie            : string  = '1');
  procedure GravaRetornoXML(pKONT_NUMERO_NOTA : string;
                            pXml_Retorno      : WideString = '';
                            pModelo           : string     = '55';
                            pSerie            : string     = '1');
  function TemXmlRetornoBD(pKONT_NUMERO_NOTA : string; var pXML: AnsiString; pModelo: String; pSerie: String): Boolean;

implementation

uses
  untFuncoes,MaskUtils,MacroMensagens,
  untConstante, UnitMain, MensagensSistemasSiac;

function TNFEInterfaceV3.AddItensNfe(pItens: array of Variant): boolean;
begin
  try
    Result := False;
    if fCapaNFe <> nil then
      result := fCapaNFe.AddItens(pItens);
  except
    //
  end;
end;

function TNFEInterfaceV3.AddCupomRef(pCupomRef: array of Variant): boolean;
begin
  try
    Result := False;
    if fCapaNFe <> nil then
      result := fCapaNFe.AddCupomRef(pCupomRef);
  except
    //
  end;
end;

function TNFEInterfaceV3.AddNFCeRef(pNFCeRef: string): boolean;
begin
  try
    Result := False;
    if fCapaNFe <> nil then
      if trim(pNFCeRef) <> '' then
        if not fCapaNFe.ExisteNFCeRef(pNFCeRef) then
          result := fCapaNFe.AddNFCeRef(pNFCeRef);
  except
    //
  end;
end;

procedure TNFEInterfaceV3.SetTotTrib(pTotTrib : String);
begin
  if fCapaNFe <> nil then
    fCapaNFe.TotTrib := pTotTrib;
end;

function TNFEInterfaceV3.AddFormaPagNfe(pFormaPag: array of Variant): boolean;
begin
  try
    Result := False;
    if fCapaNFe <> nil then
      result := fCapaNFe.AddFormaPagamento(pFormaPag);
  except
    //
  end;
end;

function TNFEInterfaceV3.AssinarXML(pEnderecoXML : String): Boolean;
var
  comando      : AnsiString;
  i            : Integer;
  ret          : TRetorno;
  retorno      : Boolean;
begin
  try
    result := False;

    ret.status := TStringList.Create;
    comando := 'NFe.AssinarNFe("' + pEnderecoXML + '")' ;

    try
      if not UtilizaAcbrMonitor then begin
        if frmAcbrNFe = nil then
          frmAcbrNFe := TfrmAcbrNFe.Create(nil);

        retorno := frmAcbrNFe.AssinarNFE(pEnderecoXML,ret);

      end else begin
        retorno := KontNfe.EnviarComando(comando,ret);
      end;
    except
      retorno := false;
    end;

    if retorno then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          if Pos('OK',ret.status[i]) > 0 then begin
            result := true;
            Break;
          end
          else if Pos('ERRO',ret.status[i]) > 0 then begin
            result := False;
            Break;
          end;
        end;
      end;
    end;
  except
    Result := false;
  end;
end;

function TNFEInterfaceV3.CertificadoDataVencimento(pMostraSplash: Boolean = False; pAutomatico: boolean = false): TDateTime;
var
  vMensagem : string;
  comando   : AnsiString;
  i         : Integer;
  ret       : TRetorno;
  retorno   : Boolean;
  vHoje     : TDateTime;
begin
  comando := 'NFe.CertificadoDataVencimento';

  if pMostraSplash then
    AbreSplashNFe('Aguarde, consultando Validade do Certificado!');

  ret.status := TStringList.Create;


  try
    if not UtilizaAcbrMonitor then begin
      if frmAcbrNFe = nil then
        frmAcbrNFe := TfrmAcbrNFe.Create(nil);

      ret.status.add('OK= ' + DateToStr(frmAcbrNFe.DataVencimentoCertificado));
      retorno   := TRUE;
    end else begin
      retorno := KontNfe.EnviarComando(comando,ret,2000,True,pAutomatico);
    end;
  except;
    retorno := False;
  end;

  if retorno then begin
    for I := 0 to ret.status.Count - 1 do begin
      vMensagem := ret.status[i];
      if Copy(vMensagem,1,2) = 'OK' then begin
        vMensagem := copy(vMensagem,5,length(vMensagem));
        result := StrToDateDef(vMensagem,Now);
        vHoje := StrToDateDef(FormatDateTime('DD/MM/YYYY',Now),Trunc(Now));
        KontNfe.ValidadeCertificado := IntToStr(trunc(result - vHoje));
      end;
    end;
  end;

  if pMostraSplash then
    FechaSplashNFe;
end;

Constructor TNFEInterfaceV3.Create(pVersao:String=''; pValidadeCertificado: string = ''; pNFCE : Boolean = false);
begin
  fCampos := TCampos.GetInstance;
  KontNfe := TCPClient.create(dm.EmpresasENDERECO_MONITORNFE.AsString,DM.EmpresasPORTA_MONITORNFE.AsInteger);
  DM.Empresas.Open;
  DM.Empresas.Locate('CNPJ',sEmpresaCNPJ,[]);
  fNFCE := pNFCE;
//  fModelo := IIf(fNFCE,'NFC-e','NF-e');
  fModelo := IIf(fCampos.ModeloNFe = 65,'NFC-e','NF-e');

  case StrToInt( DM.Empresas.FieldByName('VERSAO_NFE').AsString) of
    1 : fVersaoNfe := '1.00';
    2 : fVersaoNfe := '2.00';
    3 : fVersaoNfe := '3.10';
    4 : fVersaoNfe := '4.00';
  else
    fVersaoNfe := '4.00';
  end;

  if pVersao = '' then
    pVersao := '3.1';

  with KontNfe do begin
    if pNFCE then
      IPMonitorNfe  := EnderecoMonitorNFCe
    else
      IPMonitorNfe  := Trim(DM.Empresas.FieldByName('ENDERECO_MONITORNFE').AsString);
    Empresa       := Trim(DM.Empresas.FieldByName('CNPJ').AsString);
    UF            := Trim(DM.Empresas.FieldByName('CIDADE_UF').AsString);
    TipoAmbiente  := DM.Empresas.FieldByName('TIPO_AMBIENTE').AsInteger;
    CaminhoLogo   := Trim(DM.Empresas.FieldByName('CAMINHO_LOGOTIPO').AsString);
    SiglaUF       := Trim(DM.Empresas.FieldByName('CIDADE_UF').AsString);
    VersaoDados   := '3.10';
    Contingencia  := Trim(DM.Empresas.FieldByName('CONTINGENCIA').AsString);
    ValidadeCertificado := pValidadeCertificado;
    VersaoLayout  := fVersaoNfe;
    CodUF         := copy(BuscaCodigoCidadeIBGE(IEmpresaCodcidade),1,2);
  end;

  fQn                 := TADOQuery.Create(nil);
  fQn.connection      := Dm.adoconexao;
  fQnTemp             := TADOQuery.Create(nil);
  fQnTemp.connection  := Dm.adoconexao;
  fQf                 := TADOQuery.Create(nil);
  fQf.connection      := Dm.adoconexao;
  fMensCritica        := TStringList.Create;
end;

function TNFEInterfaceV3.SetFormaEmissao(pTpEmiss: Integer; pAutomatico: Boolean = false): Boolean;
var
  vMensagem : string;
  comando : AnsiString;
  ret     : TRetorno;
  i       : integer;
  retorno : Boolean;
begin
  comando    := 'NFe.SetFormaEmissao(' + IntToStr(pTpEmiss) + ')';
  ret.status := TStringList.Create;
  result     := false;

  try
    if not UtilizaAcbrMonitor then begin
      if frmAcbrNFe = nil then
        frmAcbrNFe := TfrmAcbrNFe.Create(nil);

      retorno := frmAcbrNFe.SetFormaEmisao(pTpEmiss,ret);

    end else begin
      retorno := KontNfe.EnviarComando(comando,ret,2000,True,pAutomatico);
    end;
  except
    retorno := false;
  end;

  if retorno then begin
    for I := 0 to ret.status.Count - 1 do begin
      vMensagem := ret.status[i];
      if Pos('OK:',vMensagem) > 0 then begin
        Result   := True;
        Break;
      end;
    end;
  end;
end;

function TNFEInterfaceV3.SetTpEmiss(pTpEmiss: Integer): Boolean;
begin
  try
    Result := False;
    if fcapaNfe <> nil then
      Result := fCapaNFe.SetTpEmiss(pTpEmiss);
  except
    Result := False;
  end;
end;

function TNFEInterfaceV3.SetdhCont: Boolean;
begin
  try
    Result := False;
    if fcapaNfe <> nil then
      Result := fCapaNFe.SetdhCont;
  except
    Result := False;
  end;
end;

function TNFEInterfaceV3.SetxJust(pJust: String): Boolean;
begin
  try
    Result := False;
    if fcapaNfe <> nil then
      Result := fCapaNFe.SetxJust(pJust);
  except
    Result := False;
  end;
end;

function TNFEInterfaceV3.CriaCapaNFe(pCapa : array of Variant): Boolean;
begin
  try
    fCapaNFe := TCapaNFe.Create(pCapa,Result);
  except
    on e: exception do begin
      Result := False;
      showmessage(e.Message);
    end;
  end;
end;


function TNFEInterfaceV3.CriaCapaNFe: Boolean;
begin
  try
    fCapaNFe := TCapaNFe.Create;
  except
    on e: exception do begin
      Result := False;
      showmessage(e.Message);
    end;
  end;
end;

Destructor TNFEInterfaceV3.Destroy;
begin

  FreeandNil(KontNfe);

  fQn.Close;
  FreeandNil(fQn);
  fQnTemp.Close;
  FreeandNil(fQnTemp);
  fQf.Close;
  FreeandNil(fQf);
  FreeandNil(fMensCritica);
end;

function TNFEInterfaceV3.NfeConsultaSituacaoNfe(pNumeroNfe : string;
                                                var pDataHora  : string;
                                                var pProtocolo : string;
                                                var pdigVal    : string;
                                                var pMSG       : string;
                                                var pStatus    : String;
                                                pSplash    : Boolean = True;
                                                pMostraMSG : Boolean = true;
                                                pAutomatico: Boolean = false): Boolean;
var
  vMensagem : string;
  comando   : AnsiString;
  j         : word;
  i         : Integer;
  ret       : TRetorno;
  retorno   : Boolean;
  vStatus   : Integer;
begin
  vMensagem := '';
  retorno := False;

  comando := 'NFe.ConsultarNFe("' + pNumeroNfe + '")';

  if pMostraMSG then begin
    if not pSplash then
      AbreSplashNFe('Aguarde, Enviando '+ fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55))
    else
      AbreSplashNFe('Aguarde, consultando situação ' + fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55));
  end;

  ret.status := TStringList.Create;

  try
    {Consulta}
//    if not UtilizaAcbrMonitor then begin
      if frmAcbrNFe = nil then
        frmAcbrNFe := TfrmAcbrNFe.Create(nil);

      try
        //verifica se tem internet
        if frmAcbrNFe.pingIp('google.com') then begin
  //      IF InternetGetConnectedState(@j,0) then begin
          retorno := frmAcbrNFe.ConsultaNFeChave(ret,pNumeroNfe);
        end
        else begin
          ret.status.add('[RETORNO]');
          ret.status.add('XMotivo=Sem conexão com internet!');
          retorno := false;
        end;
      except
        ret.status.add('[RETORNO]');
        ret.status.add('XMotivo=Sem conexão com internet!');
        retorno := false;
      end;
//    end
//    else begin
//      if Trim(pNumeroNfe) <> '' then
//        retorno := KontNfe.EnviarComando(comando,ret,2000,True,pAutomatico)
//      else
//        vMensagem := 'Não foi informado o numero da Chave-Nfe!';
//    end;

    if pSplash then
      FechaSplashNFe;
  except
    on e: exception do begin
      FechaSplashNFe;
      ShowMessage(e.message);
    end;
  end;

  if retorno then begin
    for I := 0 to ret.status.Count - 1 do begin
      if Pos('CSTAT',UpperCase(ret.status[i])) > 0 then begin
        vStatus := StrToIntDef(copy(UpperCase(ret.status[i]),7,length(UpperCase(ret.status[i]))-6),0);
        result  := BuscaStatus(vStatus,vMensagem);
        pStatus := IntToStr(vStatus);
      end;
      if Pos('ERRO: Autorizado o uso da NF-e',ret.status[i]) > 0 then begin
        vMensagem := Copy(ret.status[i],7,Length(ret.status[i])-6);
      end;
      if pos('DigestValue do documento',ret.status[i]) > 0 then begin
        vMensagem := 'DigestValue do documento não confere';
        pStatus   := '0';
        result    := false;
        break;
      end;
      if Pos('XMOTIVO=',UpperCase(ret.status[i])) > 0 then begin
        vMensagem := COPY(ret.status[i],9,length(ret.status[i]));
      end;
      if Pos('DHRECBTO=',UpperCase(ret.status[i])) > 0 then begin
        pDataHora := Copy(ret.status[i],10,Length(ret.status[i])-9);
      end;
      if Pos('NPROT=',UpperCase(ret.status[i])) > 0 then begin
        pProtocolo := Copy(ret.status[i],7,Length(ret.status[i])-6);
      end;
      if Pos('DIGVAL=',UpperCase(ret.status[i])) > 0 then begin
        pdigVal := Copy(ret.status[i],8,Length(ret.status[i])-7);
        Break;
      end;
    end;
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
  end
  else begin
    pMSG := vMensagem;
    if pSplash then begin
      vMensagem := vMensagem + #13 +
                   'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                   'UF................: '  +#9+KontNfe.UF+#13+
                   'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                   'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                   'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                   'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                   'Versão.........: '     +#9+KontNfe.VersaoLayout;
    end;
  end;

  if pSplash then
    _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'')
  else
    if pMostraMSG then
      FechaSplashNFe;

//  Result := true;
end;

function TNFEInterfaceV3.NfeConsultaXML(pEndereco : string;
                                        var pDataHora  : string;
                                        var pProtocolo : string;
                                        var pdigVal    : string;
                                        var pMSG       : string;
                                        var pStatus    : String): Boolean;
var
  vMensagem : string;
  comando   : AnsiString;
  i         : Integer;
  ret       : TRetorno;
  retorno   : Boolean;
  vStatus   : Integer;
  vArquivo  : string;
begin
  vArquivo := pEndereco;

  if not FileExists(vArquivo) then
    if pos('\ARQUIVOS\',UpperCase(pEndereco)) > 0 then
      vArquivo := copy(pEndereco,1,pos('\ARQUIVOS\',UpperCase(pEndereco))) +
                      'Arquivos\' + IIf(fNFCE,'NFCe','Nfe') +'\' + copy(pEndereco,pos('\ARQUIVOS\',UpperCase(pEndereco))+10,Length(pEndereco));

  if FileExists(vArquivo) then begin

    AbreSplashNFe('Aguarde, consultando situação ' + fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55));


    ret.status := TStringList.Create;

    try
      if frmAcbrNFe = nil then
        FrmAcbrNFe := TfrmAcbrNFe.Create(nil);

      retorno := frmAcbrNFe.ConsultaNFeXML(ret,vArquivo);


      FechaSplashNFe;
    except
      FechaSplashNFe;
    end;

    if retorno then begin
      for I := 0 to ret.status.Count - 1 do begin
        if Pos('CSTAT',UpperCase(ret.status[i])) > 0 then begin
          vStatus := StrToIntDef(copy(UpperCase(ret.status[i]),7,length(UpperCase(ret.status[i]))-6),0);
          result  := BuscaStatus(vStatus,vMensagem);
          pStatus := IntToStr(vStatus);
        end;
        if Pos('ERRO: Autorizado o uso da NF-e',ret.status[i]) > 0 then begin
          vMensagem := Copy(ret.status[i],7,Length(ret.status[i])-6);
        end;
        if pos('DigestValue do documento',ret.status[i]) > 0 then begin
          vMensagem := 'DigestValue do documento não confere';
          pStatus   := '0';
          result    := false;
          break;
        end;
        if Pos('XMOTIVO=',UpperCase(ret.status[i])) > 0 then begin
          vMensagem := COPY(ret.status[i],9,length(ret.status[i]));
        end;
        if Pos('DHRECBTO=',UpperCase(ret.status[i])) > 0 then begin
          pDataHora := Copy(ret.status[i],10,Length(ret.status[i])-9);
        end;
        if Pos('NPROT=',UpperCase(ret.status[i])) > 0 then begin
          pProtocolo := Copy(ret.status[i],7,Length(ret.status[i])-6);
        end;
        if Pos('DIGVAL=',UpperCase(ret.status[i])) > 0 then begin
          pdigVal := Copy(ret.status[i],8,Length(ret.status[i])-7);
          Break;
        end;
      end;
    end;
  end
  else begin
    vMensagem := 'XML não encontrado - ' + vArquivo;
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
  end
  else begin
    pMSG := vMensagem;
    vMensagem := vMensagem + #13 +
                 'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                 'UF................: '  +#9+KontNfe.UF+#13+
                 'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                 'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                 'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                 'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                 'Versão.........: '     +#9+KontNfe.VersaoLayout;
  end;

  _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'')
end;


procedure TNFEInterfaceV3.NfeConsultaStatusServico;
var
  vMensagem : string;
  vObs      : String;
  comando   : AnsiString;
  i         : Integer;
  ret       : TRetorno;
  retorno   : Boolean;
  vStatus   : integer;
begin
  vObs := '';
  vMensagem := '';
  comando := 'NFe.StatusServico';

  AbreSplashNFe('Aguarde, consultando status de serviço...');
  ret.status := TStringList.Create;
  try
 //   if not UtilizaAcbrMonitor then begin
      if frmAcbrNFe = nil then
        frmAcbrNFe := TfrmAcbrNFe.Create(nil);

//      PausarAutomatico;
      retorno := frmAcbrNFe.StatusServico(ret);
//      DespausarAutomatico;

//    end else begin
//      retorno := KontNfe.EnviarComando(comando,ret);
//    END;
    FechaSplashNFe;
  except
    retorno := false;
    FechaSplashNFe;
  end;

  if retorno then begin
    for I := 0 to ret.status.Count - 1 do begin
      if Pos('CSTAT',UpperCase(ret.status[i])) > 0 then begin
        vStatus := StrToIntDef(copy(UpperCase(ret.status[i]),7,length(UpperCase(ret.status[i]))-6),0);
        BuscaStatus(vStatus,vMensagem);
      end;
      if Pos('VERSAO=',UpperCase(ret.status[i])) > 0 then begin
        KontNfe.VersaoLayout := Copy(ret.status[i],8,Length(ret.status[i])-7);
      end;
//      if Pos('XMOTIVO=',UpperCase(ret.status[i])) > 0 then begin
//        vMensagem := Copy(ret.status[i],9,Length(ret.status[i])-8);
//      end;
      if Pos('XOBS=',UpperCase(ret.status[i])) > 0 then begin
        vObs := Copy(ret.status[i],6,Length(ret.status[i])-5);
      end;
      if Pos('ERRO',UpperCase(ret.status[i])) > 0 then begin
        vMensagem := Trim(copy(ret.status[i],Pos('ERRO',ret.status[i])+5,length(ret.status[i]))) + #13;
        if Pos('Inativo ou Inoperante tente novamente',Trim(ret.status[i+1])) > 0  then
          vMensagem := vMensagem + 'Inativo ou Inoperante tente novamente!' + #13;
        if Length(vMensagem) <=2 then
          vMensagem := 'Erro Interno: 12152 - Requisição não enviada.';
        if Pos('XMOTIVO=',UpperCase(ret.status[i])) > 0 then begin
          vMensagem := Copy(ret.status[i],9,Length(ret.status[i]))+#13;
        end;
        break;
      end;
    end;
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
  end
  else begin
    vMensagem := vMensagem + #13 +
                 'Observação..: '        +#9+vObs+#13+
                 'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                 'UF................: '  +#9+KontNfe.UF+#13+
                 'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                 'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                 'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                 'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                 'Versão.........: '     +#9+KontNfe.VersaoLayout;
  end;

  _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');

end;

procedure TNFEInterfaceV3.EnviarEmailParaCliente(pTipoXml:TeeEnvV3;pMensagemErro:Boolean);
var
  vAsunto   : String;
  vDir      : String;
  vEmailCli : String;
  vMensagem : string;
  i         : integer;
  comando   : AnsiString;
  resposta  : TRetorno;
  retorno   : boolean;
begin
  vMensagem := '';
  if not pMensagemErro then begin
    if (sEmailSMTP = '') or (sEmailEmail = '') then begin
      Exit;
    end;
  end;
  with _QExec do begin
    Close;
    Sql.Clear;
    connection := dm.adoconexao;
    SQL.Add('SELECT EMAIL                                 ');
    SQL.Add('  FROM CLIENTE                               ');
    SQL.Add(' WHERE CLI_CGCCPF = ''' + Trim(fCad_cgc) +'''');
    Open;
  end;
  vEmailCli := Trim(_QExec.FieldByName('EMAIL').AsString);

  Try
    vDir := DefDiretorioUsuario;
    if SelectCapaNota then begin
      if pTipoXml = eeNotaV3 then begin
        Salvar_Xml(DM.Empresas.FieldByName('CNPJ').AsString,
                   fQn.FieldByName('NUMERONF').AsString,
                   vDir,
                   'rec' + fQn.FieldByName('NUMERO_RECIBO').AsString + '.xml');
        vDir    := vDir + 'Rec'    + fQn.FieldByName('NUMERO_RECIBO').AsString + '.xml';
        vAsunto := 'NFe - Recibo: '+ fQn.FieldByName('NUMERO_RECIBO').AsString;
      end else begin
        SalvarXmlCancelamentoCliente(vDir,KontNfe.Empresa,'','','','');
        vAsunto := 'NFe - Cancelamento ';
      end;
    end;
    try
      resposta.status := TStringList.Create;
      comando := 'NFe.EnviarEmail("' + vEmailCli +'","'+vDir+'","1")';
      retorno := KontNfe.EnviarComando(comando,resposta);
    except
      retorno := false;
    end;
  finally
    if retorno then begin

      for I := 0 to resposta.status.Count - 1 do
        vMensagem := vMensagem + resposta.status[i]+#13;

      if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
        vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
      end
      else begin
        vMensagem := vMensagem + #13 +
                     'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                     'UF................: '  +#9+KontNfe.UF+#13+
                     'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                     'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                     'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                     'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                     'Versão.........: '     +#9+KontNfe.VersaoLayout;
      end;

      _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');
    end;
  end;

end;

function TNFEInterfaceV3.GetAliquotaICMS_ItensNFe(pCodigo: string;
  pItem: integer): Double;
begin
  try
    result :=  fCapaNFe.GetAliquotaICMS_ItensNFe(pCodigo,pItem);
  except
    result := 0;
  end;
end;

function TNFEInterfaceV3.GetCFOP_ItensNFe(pCodigo: string;
  pItem: integer): string;
begin
  try
    result :=  fCapaNFe.GetCFOP_ItensNFe(pCodigo,pItem);
  except
    result := '';
  end;
end;

function TNFEInterfaceV3.GetFormaPag: string;
begin
  result := fCapaNFe.FormaPag;
end;

function TNFEInterfaceV3.GetValorBaseICMS_ItensNFe(pCodigo: string;
  pItem: integer): Double;
begin
  try
    result :=  fCapaNFe.GetValorBaseICMS_ItensNFe(pCodigo,pItem);
  except
    result := 0;
  end;
end;

function TNFEInterfaceV3.GetValorICMS_ItensNFe(pCodigo: string;
  pItem: integer): Double;
begin
  try
    result :=  fCapaNFe.GetValorICMS_ItensNFe(pCodigo,pItem);
  except
    result := 0;
  end;
end;

procedure TNFEInterfaceV3.SetTotalValorNota(pValue : Double);
begin
  fCapaNFe.SetTotalValorNota(pValue);
end;

procedure TNFEInterfaceV3.SetMensagem(pValue : String);
begin
  fcapaNFe.Mensagem3 := pValue;
end;

procedure TNFEInterfaceV3.SetfModelo;
begin
  fModelo := IIf(fCampos.ModeloNFe = 65,'NFC-e','NF-e');
end;

procedure TNFEInterfaceV3.SetTotalValorProduto(pValue: Double);
begin
  fCapaNFe.SetTotalValorProduto(pValue);
end;

procedure TNFEInterfaceV3.SetDestinatario(pDestinatarioCNPJ        : string;
                                          pDestinatarioIE          : string;
                                          pDestinatarioNomeRazao   : string;
                                          pDestinatarioFone        : string;
                                          pDestinatarioCEP         : string;
                                          pDestinatarioLogradouro  : string;
                                          pDestinatarioNumero      : string;
                                          pDestinatarioComplemento : string;
                                          pDestinatarioBairro      : string;
                                          pDestinatarioCidadeCod   : string;
                                          pDestinatarioCidade      : string;
                                          pDestinatarioUF          : string;
                                          pDestinatarioindIEDest   : string;
                                          pDestinatarioCodigo      :Integer);
begin
  if fCapaNfe <> nil then
    fCapaNFe.SetDestinatario(pDestinatarioCNPJ,
                             pDestinatarioIE,
                             pDestinatarioNomeRazao,
                             pDestinatarioFone,
                             pDestinatarioCEP,
                             pDestinatarioLogradouro,
                             pDestinatarioNumero,
                             pDestinatarioComplemento,
                             pDestinatarioBairro,
                             pDestinatarioCidadeCod,
                             pDestinatarioCidade,
                             pDestinatarioUF,
                             pDestinatarioindIEDest,
                             pDestinatarioCodigo);
end;

function TNFEInterfaceV3.GravaNFe(pNFCE: Boolean = False): Boolean;
begin
  result := fCapaNFe.GravaNFe(pNFCE);
end;

function TNFEInterfaceV3.ApagaNFe: Boolean;
begin
  result := fCapaNFe.ApagaNFe;
end;

function TNFEInterfaceV3.SalvarXMLServidor(vArquivo: string; pXML: AnsiString): Boolean;
var
  comando   : AnsiString;
  ret       : TRetorno;
  retorno   : Boolean;
  i         : Integer;
begin
  try
    result := false;

    if not UtilizaAcbrMonitor then begin
      if frmAcbrNFe = nil then
        frmAcbrNFe := TfrmAcbrNFe.Create(nil);

      retorno := frmAcbrNFe.SavetoFile(vArquivo,pXML,ret);

    end else begin
      comando := 'NFE.SavetoFile("' + vArquivo + '","' + pXML + '")';
      retorno := KontNfe.EnviarComando(comando,ret);
    end;

    if retorno then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          if copy(UpperCase(ret.status[i]),1,2) = 'OK'  then begin
            result := true;
            break;
          end;
        end;
      end;
    end;
  except
    result := false;
  end;
end;


function TNFEInterfaceV3.Busca_XML(pEndereco: string; var pXML: AnsiString; pAutomatico: Boolean = false): Boolean;
var
  comando   : AnsiString;
  ret       : TRetorno;
  retorno   : Boolean;
  i         : Integer;
begin
  try
    result := false;
    if not UtilizaAcbrMonitor then begin
      if frmAcbrNFe = nil then
        frmAcbrNFe := TfrmAcbrNFe.Create(nil);

      ret.status := TStringList.Create;

      retorno := frmAcbrNFe.loadfromfile(pEndereco,ret);

    end else begin
      comando := 'NFE.loadfromfile("' + pEndereco + '",0)';
      retorno := KontNfe.EnviarComando(comando,ret,10000,True,pAutomatico);
    end;

    if retorno then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          if copy(UpperCase(ret.status[i]),1,2) = 'OK'  then begin
            pXML := copy(ret.status[i],5,length(ret.status[i]));
            result := true;
            break;
          end;
        end;
      end;
    end;
  except
    result := false;
  end;
end;

function TNFEInterfaceV3.ImprimirEvento(pPath: string): Boolean;
var
  comando   : AnsiString;
  ret       : TRetorno;
  retorno   : Boolean;
  vMensagem : AnsiString;
  i         : Integer;
  DeuErro   : Boolean;
begin
  try
//    AbreSplashNFe('Aguarde, Imprimindo Evento da NFe...');
    comando := 'NFE.imprimirevento("' + pPath + '")';
    try
      if not UtilizaAcbrMonitor then begin
        if frmAcbrNFe = nil then
          frmAcbrNFe := TfrmAcbrNFe.Create(nil);

        retorno := frmAcbrNFe.ImprimirEvento(ret,pPath);

      end else
        retorno := KontNfe.EnviarComando(comando,ret);
    except
      retorno := false;
    end;

    if retorno then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          if (Copy(UpperCase(ret.status[i]),1,2) = 'OK') then begin
            result := True;
            vMensagem := Copy(ret.status[i],5,Length(ret.status[i])-4)+#13;
          end;
          if (Copy(UpperCase(ret.status[i]),1,4) = 'ERRO') then begin
            result := False;
            vMensagem := Copy(ret.status[i],6,Length(ret.status[i])-5)+#13;
            DeuErro := True;
            Break;
          end;
          if Pos('VERSAO=',UpperCase(ret.status[i])) > 0 then begin
            KontNfe.VersaoLayout := Copy(ret.status[i],8,Length(ret.status[i])-7);
          end;
        end;
      end;
    end;

    if (vMensagem = '') and (not DeuErro) and (UtilizaAcbrMonitor) then begin
      vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
    end
    else begin
      vMensagem := vMensagem + #13 +
                   'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                   'UF................: '  +#9+KontNfe.UF+#13+
                   'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                   'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                   'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                   'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                   'Versão.........: '     +#9+KontNfe.VersaoLayout;
    end;
 //   FechaSplashNFe;

   _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');

  except
    on e: Exception do begin
      ShowMessage(e.Message);
    end;
  end;
end;

function TNFEInterfaceV3.ImprimirDanfeNFce(cArquivo: string; pNFE: Boolean = false): Boolean;
VAR
  vXML : STRING;
begin
  try
    if frmAcbrNFe = nil then
      frmAcbrNFe := TfrmAcbrNFe.Create(nil);

    if cArquivo = '' then
      vXML := fXML
    else
      vxml := cArquivo;

//    sleep(1000);
    frmAcbrNFe.ImprimirDanfeNfce(vXML,pNFE);
  except
    Result := false;
  end;
end;

function TNFEInterfaceV3.ImprimirDanfeNfceCancelado(cArquivo: string): Boolean;
VAR
  vXML : STRING;
begin
  try
    if frmAcbrNFe = nil then
      frmAcbrNFe := TfrmAcbrNFe.Create(nil);

    if cArquivo = '' then
      vXML := fXML
    else
      vxml := cArquivo;

    frmAcbrNFe.ImprimirDanfeNfceCancelado(vXML);
  except
    Result := false;
  end;
end;

function TNFEInterfaceV3.ImprimirTextoTEF(pTextoTEF: TStringList): boolean;
begin
  try
    result := frmAcbrNFe.ImprimeRelatorio(pTExtoTEF);
  except
    result := false;
  end;
end;

function TNFEInterfaceV3.ImprimirDanfeNFe(cArquivo: string; pChaveNFE: STRING; pXML: AnsiString = ''; pPerguntaVisualizar: Boolean = True; pTipo: Integer = 0): Boolean;
var
  comando   : AnsiString;
  ret       : TRetorno;
  vMensagem : AnsiString;
  vMSG      : string;
  retorno   : Boolean;
  i         : integer;
  DeuErro   : Boolean;
  Visualizar : Boolean;
  vArquivo   : string;
  vDataHora  : String;
  vProtocolo : String;
  vdigVal    : String;
  vStatus    : string;
begin
  try
    Visualizar := false;
    DeuErro := false;
    result  := false;
    vMensagem := '';
    retorno   := FALSE;
    ret.status := TStringList.Create;

    case pTipo of
      0 : Visualizar := False;
      1 : Visualizar := True;
      2 : Visualizar := False;
    end;

//    if pPerguntaVisualizar then
//      Visualizar := (MessageBox(0, 'Deseja visualizar o Danfe?', 'KontPosto - Confirmação', MB_ICONQUESTION or MB_YESNO) = idYes)
//    else
//      Visualizar := true;

    if (NOT Visualizar) then begin
      if pos('\ARQUIVOS\',UpperCase(cArquivo)) > 0 then begin
        vArquivo := copy(cArquivo,1,pos('\ARQUIVOS\',UpperCase(cArquivo))) +
                    'Arquivos\' + IIf(fNFCE,'NFCe','Nfe') +'\' + pChaveNFE + '-nfe.xml';

        if not FileExists(vArquivo) then  begin
          if pXML <> '' then begin
            if NfeConsultaSituacaoNfe(pChaveNFE,
                                      vDataHora,
                                      vProtocolo,
                                      vdigVal,
                                      vMSG,
                                      vStatus,
                                      False,
                                      False) then begin
              TrataXML(pXML,
                       pChaveNFE,
                       vDataHora,
                       vProtocolo,
                       vdigVal);
              if not SalvarXMLServidor(vArquivo,pXML) then begin
                vArquivo := copy(cArquivo,1,pos('\ARQUIVOS\',UpperCase(cArquivo))) +
                            'Arquivos\' + pChaveNFE + '-nfe.xml';
              end;
            end;
          end
          else begin
            if FileExists(cArquivo) then
              vArquivo := cArquivo
            else begin
              vArquivo := copy(cArquivo,1,pos('\ARQUIVOS\',UpperCase(cArquivo))) +
                          'Arquivos\' + pChaveNFE + '-nfe.xml';

            end;
          end;
        end;
      end;
    end;

    if Visualizar then begin
      if pXML = '' then begin
        AbreSplashNFe('Aguarde, Localizando XML ' + fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55));
        if Busca_XML(cArquivo,pXML) then begin
          vMensagem := pXML;
          result := true;
          retorno := False;
        end;
      end
      else begin
        vMensagem := pXML;
        result := true;
        retorno := False;
      end;
    end
    else begin
      if pTipo = 0 then
        comando := 'NFE.imprimirdanfe("' + vArquivo + '")'
      else if (pTipo = 2) then
        comando := 'NFE.ImprimirDanfePDF("' + vArquivo + '")'
    end;

    try
      if not UtilizaAcbrMonitor then begin
        if frmAcbrNFe = nil then
          frmAcbrNFe := TfrmAcbrNFe.Create(nil);

        if pTipo = 0 then
          retorno := frmAcbrNFe.ImprimirDanfe(ret,vArquivo)
        else if (pTipo = 2) then
          retorno := frmAcbrNFe.imprimirDanfePDF(ret, vArquivo)
        else if (pTipo = 1) and (not fNFCE) then begin//imprimir nfe no pdv
          retorno := frmAcbrNFe.ImprimirDanfeNfce(vMensagem,true);
        end;

      end else begin
        if (not Visualizar) then
          retorno := KontNfe.EnviarComando(comando,ret)
      end;
    except
      retorno := false;
    end;

    if retorno then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          if (Copy(UpperCase(ret.status[i]),1,2) = 'OK') then begin
            result := True;
            vMensagem := Copy(ret.status[i],5,Length(ret.status[i])-4)+#13;
          end;
          if (Copy(UpperCase(ret.status[i]),1,4) = 'ERRO') then begin
            result := False;
            vMensagem := Copy(ret.status[i],6,Length(ret.status[i])-5)+#13;
            DeuErro := True;
            Break;
          end;
          if Pos('VERSAO=',UpperCase(ret.status[i])) > 0 then begin
            KontNfe.VersaoLayout := Copy(ret.status[i],8,Length(ret.status[i])-7);
          end;
        end;
      end;
    end;

    if Result then begin
      if Visualizar then begin
        if frmAcbrNFe = nil then
           frmAcbrNFe := TfrmAcbrNFe.Create(nil);

        if fNFCE then begin
          frmAcbrNFe.ImprimirDanfeNfce(vMensagem,not fNFCE);
        end
        else begin
          if NfeConsultaSituacaoNfe(pChaveNFE,
                                    vDataHora,
                                    vProtocolo,
                                    vdigVal,
                                    vMSG,
                                    vStatus,
                                    False,
                                    False) then begin
            if vMensagem <> '' then begin
              TrataXML(vMensagem,
                       pChaveNFE,
                       vDataHora,
                       vProtocolo,
                       vdigVal);
  //            if fNFCE then
  //              frmAcbrNFe.ImprimirDanfeNfce(vMensagem)
  //            else begin
                if (UpperCase(Copy(vMensagem,Length(vMensagem)-8,9)) = '/NFEPROC>')or
                   (UpperCase(Copy(vMensagem,Length(vMensagem)-4,5)) = '/NFE>') then
                  frmAcbrNFe.ImprimirDanfe(ret,vMensagem)
                else begin
                  result := False;
                  DeuErro := true;
                  vMensagem := 'Ocorreu um erro ao imprimir o Danfe, tente novamente imprimir a NFe!' + #13;
                end;
  //            end;
            end;
          end
          else begin
            result := False;
            vMensagem := vMSG;
            Visualizar := false;
          end;
        end;
      end;
    end;

    if (vMensagem = '') and (not DeuErro) and (UtilizaAcbrMonitor) then begin
      vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
    end
    else begin
      vMensagem := vMensagem + #13 +
                   'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                   'UF................: '  +#9+KontNfe.UF+#13+
                   'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                   'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                   'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                   'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                   'Versão.........: '     +#9+KontNfe.VersaoLayout;
    end;
    FechaSplashNFe;
    if ((not Visualizar) or DeuErro) and (pTipo <> 1) then
      _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');
  except
    on e: Exception do begin
      ShowMessage(e.Message);
    end;
  end;
end;

function TNFEInterfaceV3.ImprimirDanfePDF(pXML: AnsiString; var pEndPDF: string): boolean;
VAR
  ret : TRetorno;
  i: integer;
begin
  try
    if frmAcbrNFe = nil then
      frmAcbrNFe := TfrmAcbrNFe.Create(nil);

    ret.status := TStringList.Create;
    result := frmAcbrNFe.imprimirDanfePDF(ret, pXML);

    if Result then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          if (Pos('PDF=',ret.status[i]) > 0) then begin
            pEndPDF := copy(ret.status[i],5,length(ret.status[i])-4);
          end;
        end;
      end;
    end;
  except
    result := false;
  end;
end;

function TNFEInterfaceV3.CancelarNFe(pChaveNFe,
                                     cJustificativa: string;
                                     pCNPJ: string;
                                     var pStatus,
                                     pProtocolo: string;
                                     pEvento : Boolean = True): Boolean;
var
  vMensagem    : string;
  comando      : AnsiString;
  i            : Integer;
  ret          : TRetorno;
  vCancelamento : Boolean;
  retorno       : Boolean;
  vStatus       : Integer;
  vDhEvento     : String;
begin
  try
    vCancelamento := false;
    vMensagem := '';
    AbreSplashNFe('Aguarde, Cancelando ' + fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55));
    ret.status := TStringList.Create;


    if bVoltaUmaHora then
      vDhEvento := FormatDateTime('DD/MM/YY HH:MM:SS',Now-0.04166)
    else
      vDhEvento := FormatDateTime('DD/MM/YY HH:MM:SS',Now);

    if not pEvento then
      comando := 'NFE.cancelarnfe("' + pChaveNFe + '","' + cJustificativa + '")'
    else
      comando := 'NFE.enviarevento("[EVENTO]' +sLineBreak+
                 'idLote=1'    + sLineBreak+
                 '[EVENTO001]' + sLineBreak+
                 'chNFe='      + pChaveNFe + sLineBreak+
                 'cOrgao='     + copy(BuscaCodigoCidadeIBGE(IEmpresaCodcidade),1,2) + sLineBreak+
                 'CNPJ='       + pCNPJ + sLineBreak+
                 'dhEvento='   + vDhEvento +sLineBreak+ // 15/03/15 15:34:25
                 'tpEvento=110111' + sLineBreak+
                 'nProt='          + pProtocolo + sLineBreak+
                 'xJust='          + Trim(cJustificativa) + '")';
      try
        if not UtilizaAcbrMonitor then begin
          if frmAcbrNFe = nil then
            frmAcbrNFe := TfrmAcbrNFe.Create(nil);

          retorno := frmAcbrNFe.CancelarNfe(pChaveNFe,
                                            cJustificativa,
                                            ret,
                                            StrToInt(copy(BuscaCodigoCidadeIBGE(IEmpresaCodcidade),1,2)));

        end else begin
          retorno := KontNfe.EnviarComando(comando,ret);
        END;
      except
        retorno := False;
      end;
//      retorno := False;
//      vMensagem := 'Utilize o computador onde contém o Certificado Digital' +#13+
//                   'para fazer o Cancelamento!'+#13;
//    end;

    if retorno then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          if (Pos('[CANCELAMENTO]',ret.status[i]) > 0) or
             (Pos('[EVENTO001]',ret.status[i]) > 0) then begin
            vCancelamento := True;
          end;
//          if (Copy(UpperCase(ret.status[i]),1,2) = 'OK') and (not vCancelamento) then begin
//            if Copy(UpperCase(ret.status[i]),1,2) = 'OK' then begin
//              result := true;
//            end;
//          end
//          else
          if (Copy(UpperCase(ret.status[i]),1,4) = 'ERRO') and (not vCancelamento) then begin
            result := False;
            Break;
          end;

          if vCancelamento then begin
            if Pos('XMOTIVO=',UpperCase(ret.status[i])) > 0 then begin
              vMensagem := Copy(ret.status[i],9,Length(ret.status[i])-8)+#13;
            end;
            if Pos('CSTAT=',UpperCase(ret.status[i])) > 0 then begin
              pStatus := copy(ret.status[i],Pos('CSTAT=',UpperCase(ret.status[i]))+6,3);
              result := ((pStatus = '101')OR (pStatus = '135'));
            end;
//            if (Pos('CSTAT',UpperCase(ret.status[i])) > 0) and
//               (copy(UpperCase(ret.status[i]),1,18)<>'XML=<?XML VERSION=') then begin
//              vStatus := StrToIntDef(copy(ret.status[i],7,length(ret.status[i])-6),0);
//              pStatus := copy(ret.status[i],7,length(ret.status[i])-6);
//              result := BuscaStatus(vStatus,vMensagem);
//            end;
            if Pos('NPROT=',UpperCase(ret.status[i])) > 0 then begin
              pProtocolo := copy(ret.status[i],Pos('NPROT=',UpperCase(ret.status[i]))+7,Length(ret.status[i]));
            end;
//            if Pos('XML=',UpperCase(ret.status[i])) > 0 then begin
//              pXML := copy(ret.status[i],Pos('XML=',UpperCase(ret.status[i]))+5,Length(ret.status[i]));
//            end;
          end;
        end;
      end
      else begin
        vMensagem := 'NFe não Cancelada! Tente novamente!'+#13;
      end;
    end;
  finally
    FechaSplashNFe;
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
  end
  else begin
    vMensagem := vMensagem + #13 +
                 'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                 'UF................: '  +#9+KontNfe.UF+#13+
                 'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                 'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                 'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                 'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                 'Versão.........: '     +#9+KontNfe.VersaoLayout;
  end;

  _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');
end;

function TNFEInterfaceV3.CartaCorrecao(pChaveNFe  : string;
                                       pCorrecao  : array of AnsiString;
                                       pCNPJ      : string;
                                       dhEvento   : String;
                                       nSeqEvento : array of Integer): Boolean;
var
  vMensagem    : string;
  comando      : AnsiString;
  i            : Integer;
  ret          : TRetorno;
  retorno      : Boolean;
begin
  try
    vMensagem := '';
    AbreSplashNFe('Aguarde, enviando a Carta de Correção da ' + fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55));
    ret.status := TStringList.Create;

    if not UtilizaAcbrMonitor then begin
//    if Pos(sNomeComputador,sENDERECO_MONITORNFE) > 0 then begin
      if frmAcbrNFe = nil then
        frmAcbrNFe := TfrmAcbrNFe.Create(nil);

      KontNfe.Contingencia := '1';

      KontNfe.UF := sUF_WebService;

      KontNfe.ValidadeCertificado := frmAcbrNFe.ValidadeCertificado;

      retorno := frmAcbrNFe.CartaCorrecao(1,
                                          StrToInt(copy(BuscaCodigoCidadeIBGE(IEmpresaCodcidade),1,2)),
                                          pCNPJ,
                                          pChaveNFe,
                                          nSeqEvento,
                                          pCorrecao,
                                          ret);
    end
    else begin
      comando := 'NFE.cartadecorrecao("[CCE]' +sLineBreak+
                 'idLote=1'    + sLineBreak;

      for I := 0 to High(nSeqEvento) do begin

        if i > 0 then
          comando := comando + sLineBreak;

        comando := comando + '[EVENTO' + FormatFloat('000',nSeqEvento[i]) + ']' + sLineBreak +
                             'idLote=1'    + sLineBreak +
                             'chNFe='      + pChaveNFe + sLineBreak +
                             'cOrgao='     + copy(BuscaCodigoCidadeIBGE(IEmpresaCodcidade),1,2) + sLineBreak +
                             'CNPJ='       + pCNPJ + sLineBreak +
                             'dhEvento='   + dhEvento + sLineBreak +
                             'nSeqEvento=' + IntToStr(nSeqEvento[i]) + sLineBreak +
                             'xCorrecao='  + pCorrecao[i];
      end;
      comando := comando + '")';

      try
        retorno := KontNfe.EnviarComando(comando,ret);
      except
        retorno := False;
      end;
//      retorno := false;
//      vMensagem := 'Utilize o computador onde contém o Certificado Digital' +#13+
//                   'para fazer a Carta de Correção!'+#13;
    end;

    if retorno then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          if Pos('CSTAT=135',UpperCase(ret.status[i])) > 0 then begin
            result := true;
          end;
          if (Copy(ret.status[i],1,4) = 'ERRO') then begin
            result := False;
            Break;
          end;
          if Pos('XMOTIVO=',UpperCase(ret.status[i])) > 0 then begin
            vMensagem := Copy(ret.status[i],9,Length(ret.status[i])-8)+#13;
          end;
          if Pos('VERAPLIC=',UpperCase(ret.status[i])) > 0 then begin
            KontNfe.VersaoLayout := Copy(ret.status[i],10,Length(ret.status[i])-9);
          end;
          if Pos('TPAMB=',UpperCase(ret.status[i])) > 0 then begin
            KontNfe.TipoAmbiente := StrToIntDef(Copy(ret.status[i],7,Length(ret.status[i])-6),0);
          end;
          if Pos('CORGAO=',UpperCase(ret.status[i])) > 0 then begin
            KontNfe.CodUF := Copy(ret.status[i],8,Length(ret.status[i])-7);
          end;
        end;
      end
      else begin
        vMensagem := 'Não foi enviada a carta de correção! Tente novamente!'+#13;
      end;
    end;

  finally
    FechaSplashNFe;
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
  end
  else begin
    vMensagem := vMensagem + #13 +
                 'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                 'UF................: '  +#9+KontNfe.UF+#13+
                 'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                 'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                 'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                 'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                 'Versão.........: '     +#9+KontNfe.VersaoLayout;
  end;

  _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');
end;

function TNFEInterfaceV3.InutilizarNFe(cCNPJ, cJustificativa: string; nAno,
  nModelo, nSerie, nNumInicial, nNumFinal: integer; pDate: TDateTime; pSplash: boolean = true): Boolean;
var
  vMensagem    : string;
  comando      : AnsiString;
  i            : Integer;
  ret          : TRetorno;
  vInutilizacao : Boolean;
  retorno       : Boolean;
  xml_enviado   : WideString;
  xml_retorno   : WideString;
  NumeroProtocolo : String;
begin
  try
    vInutilizacao := false;
    vMensagem := '';

    if pSplash then
      AbreSplashNFe('Aguarde, Inutilizando número(s) ' + fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55));

    ret.status := TStringList.Create;
    comando := 'NFe.InutilizarNFe("'  + cCNPJ + '",' +
                                  '"' + cJustificativa  + '",' +
                                  IntToStr(nAno)        + ',' +
                                  IntToStr(nModelo)     + ',' +
                                  IntToStr(nSerie)      + ',' +
                                  IntToStr(nNumInicial) + ',' +
                                  IntToStr(nNumFinal)   + ')' ;
    try
      if  not UtilizaAcbrMonitor then begin
        if frmAcbrNFe = nil then
          frmAcbrNFe := TfrmAcbrNFe.Create(nil);

        retorno := frmAcbrNFe.Inutilizar(ret,
                                         cCNPJ,
                                         cJustificativa,
                                         nAno,
                                         nModelo,
                                         nSerie,
                                         nNumInicial,
                                         nNumFinal);
      end
      else
        retorno := KontNfe.EnviarComando(comando,ret);
    except
      retorno := False;
    end;

    if retorno then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          vMensagem := ret.status[i]+#13;
          if Pos('[INUTILIZACAO]',vMensagem) > 0 then begin
            vInutilizacao := True;
          end;
          if (Copy(vMensagem,1,2) = 'OK') and (not vInutilizacao) then begin
            vMensagem := ret.status[i+1]+#13;
            if Copy(vMensagem,1,2) = 'OK' then begin
              result := true;
              Break;
            end;
          end
          else if (Copy(vMensagem,1,4) = 'ERRO') and (not vInutilizacao) then begin
            result := False;
            Break;
          end
          else if vInutilizacao then begin
            if Pos('XMotivo=Rejei',vMensagem) > 0 then begin
              result := False;
              vMensagem := Copy(vMensagem,9,Length(vMensagem)-8)+#13;
              Break;
            end;
            if Pos('CSTAT=102',UpperCase(vMensagem))> 0 then begin
              result := true;
              vMensagem := 'Inutilização de número homologado!' + #13;
              break;
            end;
            if Pos('XML=',vMensagem) > 0 then begin
              xml_retorno := copy(vMensagem,5,Length(vMensagem)-4);
            end;
            if Pos('XML_ENVIADO=',vMensagem) > 0 then begin
              xml_enviado := copy(vMensagem,13,Length(vMensagem)-12);
            end;
            if Pos('DhRecbto=',vMensagem) > 0 then begin
            end;
            if Pos('NProt=',vMensagem) > 0 then begin
              NumeroProtocolo := copy(vMensagem,7,Length(vMensagem)-7);
            end;
          end;
        end;
      end
      else begin
        vMensagem := 'NFe não Inutilizada! Tente novamente!'+#13;
      end;
    end;
  finally
    if pSplash then
      FechaSplashNFe;
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
  end
  else begin
    vMensagem := vMensagem + #13 +
                 'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                 'UF................: '  +#9+KontNfe.UF+#13+
                 'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                 'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                 'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                 'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                 'Versão.........: '     +#9+KontNfe.VersaoLayout;
  end;

  GravaInutilizacaoNFe(cCNPJ,
                       cJustificativa,
                       nAno,
                       nModelo,
                       nSerie,
                       nNumInicial,
                       nNumFinal,
                       xml_enviado,
                       xml_retorno,
                       NumeroProtocolo,
                       pDate);

  if pSplash then
    _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');
end;

function TNFEInterfaceV3.Salvar_XML(aEmpresaId,aNotaFiscal: string; pDir: String = ''; pNomeArq: String = ''; pModelo: String = '55'; pSerie: String = '1'): boolean;
var
  I : Integer;
  Dir,
  xCad_cgc: string;
  ret,
  env:WideString;
  xmldoc: TXMLDocument;
  RetNode,
  EnvNode:IXMLNode;
begin
  Result := False;

  xCad_cgc := '';
  if pNomeArq='' then begin
    PreencheQRep('SELECT CNPJCPF                           '+
                 '  FROM NOTAFISCAL                        '+
                 'WHERE  NUMERONF = ''' + aNotaFiscal +''' ');

    if _QExec.IsEmpty then
      Exit;
    xCad_cgc := Trim(_QExec.FieldByName('CNPJCPF').AsString);
  end;

  PreencheQRep('SELECT XML_ENVIADO,                          '+
               '       XML_RETORNO                           '+
               '  FROM NFE_HISTORICOXML                      '+
               ' WHERE EMPRESA_ID   = ''' + Trim(DM.Empresas.FieldByName('CNPJ').AsString) + ''' '+
               '   AND NOTA_FISCAL  = ''' + aNotaFiscal +''' '+
               '   AND MODELO       = ''' + pModelo     +''' '+
               '   AND SERIE        = ''' + pSerie      +''' '+
               'ORDER BY HISTORICO_ID DESC');

  if _QExec.IsEmpty then
    Exit;

  ret := '';
  env := '';

  While not _QExec.Eof do begin
    if env = '' then begin
      env := _QExec.FieldByName('XML_RETORNO').AsString;
      if Pos(WideString('<cStat>103</cStat>'),env) > 0 then begin
        env := _QExec.FieldByName('XML_ENVIADO').AsString;
      end else begin
        env := '';
      end;
    end;
    if ret = '' then begin
      ret := _QExec.FieldByName('XML_RETORNO').AsString;
      if (Pos(WideString('<cStat>100</cStat>'),ret) = 0) and
         (Pos(WideString('<cStat>150</cStat>'),ret) = 0) then begin
        ret := '';
      end;
    end;
    _QExec.Next;
  end;

  if pDir = '' then begin
    Dir := SelecionarDirStr;
  end else begin
    Dir := pDir;
  end;
  if (Trim(Dir) = '') or (ret = '') or (env = '') then begin
    MSGAtencao('A T E N Ç Ã O ! ! !'+#13+#13+
               'Arquivo não encontrado ou não pode ser gerado!');
    Exit;
  end;

  xmldoc := TXMLDocument.Create(Application);
  Try
    xmldoc.XML.Text  := ret;
//    xmldoc.DOMVendor := GetDOMVendor('Xerces XML');
    xmldoc.Active    := True;
    ret := '';

    RetNode := xmldoc.DocumentElement.ChildNodes.FindNode('protNFe');
    if RetNode = nil then
      Exit;

    xmldoc.Active    := False;
    xmldoc.XML.Clear;
    xmldoc.XML.Text  := Trim(TirarAcento(env));
    xmldoc.Active    := True;

    if xmldoc.Version<>'' then
      env := '<?xml version="'+xmldoc.Version +'" encoding="'+xmldoc.Encoding+'" ?>'
    else
      env := '';

    ret := '';
    for i:= 0 to xmldoc.DocumentElement.AttributeNodes.Count-1 do begin
      ret := ret + xmldoc.DocumentElement.AttributeNodes[i].NodeName+'="'+xmldoc.DocumentElement.AttributeNodes[i].NodeValue+'" ';
    end;

    EnvNode := xmldoc.DocumentElement.ChildNodes.FindNode('NFe');
    if EnvNode = nil then
      Exit;

    xmldoc.Active    := False;
    xmldoc.XML.Clear;
    xmldoc.XML.Append(env);
    xmldoc.XML.Append('<nfeProc '+ret+'>');
    xmldoc.XML.Append(EnvNode.XML);
    xmldoc.XML.Append(RetNode.XML);
    xmldoc.XML.Append('</nfeProc>');
    xmldoc.Active    := True;

    Dir := IncludeTrailingPathDelimiter(Dir);

    if pNomeArq = '' then begin
      xmldoc.SaveToFile(Dir+'Rec_'+Numeros(aEmpresaId)+'_'+Numeros(xCad_cgc)+'_'+Numeros(aNotaFiscal)+'.xml');
    end else begin
      if Pos('.XML',UpperCase(pNomeArq))=0 then
        xmldoc.SaveToFile(Dir+pNomeArq+'.xml')
      else
        xmldoc.SaveToFile(Dir+pNomeArq);
    end;
    Result := True;
  finally
    FreeAndNil(xmldoc);
  end;
end;

procedure TNFEInterfaceV3.SalvarXmlCancelamentoCliente(pDiretorio,empresaid,pedido,pstatus,xmlEnvio,xmlRetorno:String);
var
  xmldoc     :TXMLDocument;
  xNode      :IXMLNode;
  xxmlEnvio  :String;
  xxmlRetorno:String;
begin
  xxmlEnvio   := xmlEnvio;
  xxmlRetorno := xmlRetorno;
  if xxmlEnvio = '' then begin
    PreencheQRep('SELECT XML_ENVIADO, XML_RETORNO               ' +
                 '  FROM NFE_HISTORICOXML                       ' +
                 ' WHERE EMPRESA_ID     = ''' + empresaid + ''' ' +
                 '   AND NOTA_FISCAL    = ''' + pedido    + ''' ' +
                 '   AND KONT_STATUS    = ''' + pstatus   + ''' ' +
                 '   AND STATUS_RETORNO in (''101'',''151'')    ');
    if _QExec.IsEmpty then
      Exit;
    xxmlEnvio   := _QExec.FieldByName('XML_ENVIADO').AsString;
    xxmlRetorno := _QExec.FieldByName('XML_RETORNO').AsString;
  end;
  xmldoc := TXMLDocument.Create(Application);
  Try
    xmldoc.Active    := True;
    xNode            := xmldoc.AddChild('procCancNFe');
    xNode.SetAttribute('versao','2.00');
    xmldoc.DocumentElement.ChildNodes.Add( xmldoc.CreateNode('ini'+TirarAcento(xxmlEnvio)+'fim',ntComment	) );
    xmldoc.DocumentElement.ChildNodes.Add( xmldoc.CreateNode('ini'+TirarAcento(xxmlRetorno)+'fim',ntComment	) );
    xmldoc.XML.Text  := StringReplace(xmldoc.XML.Text, '<!--ini', '', [rfReplaceAll]);
    xmldoc.XML.Text  := StringReplace(xmldoc.XML.Text, 'fim-->',  '', [rfReplaceAll]);
    xmldoc.Active    := false;
    xmldoc.Active    := True;
//    LocTag(xmldoc.ChildNodes,'Signature',true);
    xmldoc.XML.SaveToFile( IncludeTrailingPathDelimiter(pDiretorio)+'Can_'+Numeros(empresaid)+'_'+pedido+'.xml');
    xmldoc.XML.Clear;
  finally
    FreeAndNil(xmldoc);
  end;
end;

//------------------------------selects---------------------------------------//

function  TNFEInterfaceV3.SelectCapaNota:Boolean;
  procedure ad(pString:String);
  begin
    fQn.SQL.Add(pString);
  end;
begin
  fQn.Close;
  fQn.SQL.Clear;
  ad('select                                                  ');
  ad('       nta.NumeroNF,                                    ');
  ad('       nta.Serie,                                       ');
  ad('       nta.CFOP,                                        ');
  ad('       nta.InscricaoSubstituicao,                       ');
  ad('       nta.Inscricaoestadual,                           ');
  ad('       nta.CodCliente,                                  ');
  ad('       nta.Nome,                                        ');
  ad('       nta.CnpjCpf,                                     ');
  ad('       nta.Emissao,                                     ');
  ad('       nta.Saida,                                       ');
  ad('       nta.HoraSaida,                                   ');
  ad('       nta.Endereco,                                    ');
  ad('       nta.Bairro,                                      ');
  ad('       nta.CEP,                                         ');
  ad('       nta.Municipio,                                   ');
  ad('       nta.FoneFax,                                     ');
  ad('       nta.UF,                                          ');
  ad('       nta.BaseIcms,                                    ');
  ad('       nta.Icms,                                        ');
  ad('       nta.BaseSubstituicao,                            ');
  ad('       nta.ValorSubstituicao,                           ');
  ad('       nta.TotalProduto,                                ');
  ad('       nta.TotalNota,                                   ');
  ad('       nta.Frete,                                       ');
  ad('       nta.Seguro,                                      ');
  ad('       nta.Outros,                                      ');
  ad('       nta.IPI,                                         ');
  ad('       nta.Observacao,                                  ');
  ad('       nta.Observacao,                                  ');
  ad('       nta.Mensagem1,                                   ');
  ad('       nta.Mensagem2,                                   ');
  ad('       nta.Mensagem3,                                   ');
  ad('       nta.NaturezaOperacao,                            ');
  ad('       nta.VALORDESCONTO,                               ');
  ad('       nta.Cancelada,                                   ');
  ad('       nta.Modelo,                                      ');
  ad('       nta.VALORISENTO,                                 ');
  ad('       nta.VALOROUTRAS,                                 ');
  ad('       nta.ALIQICMS,                                    ');
  ad('       nta.VERSAO_NFE,                                  ');
  ad('       nta.JUSTIFICATIVA_CONTINGENCIA,                  ');
  ad('       nta.DATA_HORA_CONTINGENCIA,                      ');
  ad('       nta.DATA_HORA_RECEB_NFE,                         ');
  ad('       nta.STATUS_CTG,                                  ');
  ad('       nta.TIPO_AMBIENTE_NFE,                           ');
  ad('       nta.NUMERO_LOTE_NFE,                             ');
  ad('       nta.NUMERO_PROTOCOLO_CANCELAMENTO,               ');
  ad('       nta.NUMERO_PROTOCOLO,                            ');
  ad('       nta.NOTA_FISCAL_NFE,                             ');
  ad('       nta.STATUS_NFE,                                  ');
  ad('       nta.NUMERO_RECIBO,                               ');
  ad('       nta.NUMERO_NFE,                                  ');
  ad('       nta.CONTINGENCIA As CONTINGENCIA_ANTES,          ');
  ad('       nta.MOTIVO_CANCELAMENTO,                         ');
  ad('       nta.COD_SIT_EFD,                                 ');
  ad('       nta.MENSAGEM_FISCO_ID,                           ');
  ad('       nta.MENSAGEM_CONTRIBUINTE_ID,                    ');
  ad('       nta.TOTAL_OUTRAS_DESP,                           ');
  ad('       nta.CONDICAO_TIPO,                               ');
  ad('       nta.CONDICAO_DESCRICAO,                          ');
  ad('       nta.DESC_NAT_OPERACAO,                           ');
  ad('       ''N'' As COMPLEMENTO_ICMS,                       ');
  ad('       0 As VALOR_FRETE,                                ');
  ad('       0 As VALOR_SEGURO,                               ');
  ad('       0 As VALOR_IPI,                                  ');
  ad('       1 AS TIPO_NOTA                                   ');
  ad('from                                                    ');
  ad('      NotaFiscal nta                                    ');
  ad('where nta.NumeroNF     = '''+ fNumero_nota+ '''         ');
  ad('  and nta.SaidaEntrada = ''S''                          ');
  fQn.Open;
  result := not fQn.IsEmpty;

  if not fQn.IsEmpty then begin
    TNumericField(fQn.FieldByName('TotalNota')).DisplayFormat := '#,###,##0.0000';
  end;
end;

function TNFEInterfaceV3.SelectItemNota:Boolean;
  procedure ad(pString:String);
  begin
    fQn.SQL.Add(pString);
  end;
begin
  fQn.Close;
  fQn.SQL.Clear;
  ad('select item.Serie,                            ');
  ad('       item.NumeroNF,                         ');
  ad('       item.CodProduto,                       ');
  ad('       item.Descricao,                        ');
  ad('       item.CST,                              ');
  ad('       item.Unidade,                          ');
  ad('       item.Quantidade,                       ');
  ad('       item.VlrUnitario,                      ');
  ad('       item.Vlrtotal,                         ');
  ad('       item.Modelo,                           ');
  ad('       prod.CODIGO_ANP,                       ');
  ad('       prod.CODIGO_NCM,                       ');
  ad('       prod.PERC_RED_ICMS,                    ');
  ad('       item.Modelo,                           ');
  ad('       item.ALIQ_ICMS,                        ');
  ad('       item.ALIQ_ICMS_CHEIA,                  ');
  ad('       item.VL_BC_ICMS,                       ');
  ad('       item.VL_REDUCAO,                       ');
  ad('       item.VL_OUT_DESP_ITEM,                 ');
  ad('       item.VL_FRETE_ITEM,                    ');
  ad('       item.VL_ICMS,                          ');
  ad('       item.ALIQ_ICMS_ST,                     ');
  ad('       item.VL_BC_ICMS_ST,                    ');
  ad('       item.VL_ICMS_ST,                       ');
  ad('       item.VL_BC_PIS,                        ');
  ad('       item.ALIQ_PIS,                         ');
  ad('       item.QUANT_BC_PIS,                     ');
  ad('       item.VL_PIS,                           ');
  ad('       item.VL_BC_COFINS,                     ');
  ad('       item.ALIQ_COFINS,                      ');
  ad('       item.QUANT_BC_COFINS,                  ');
  ad('       item.VL_COFINS,                        ');
  ad('       item.ALIQ_ISSQN,                       ');
  ad('       item.VL_BC_ISSQN,                      ');
  ad('       item.VL_ISSQN,                         ');
  ad('       item.VL_ISENTO,                        ');
  ad('       item.VL_NAO_TRIB,                      ');
  ad('       item.VL_DESC_ITEM,                     ');
  ad('       item.PERCENT_IVA,                      ');
  ad('       item.PERCENT_REDUTOR,                  ');
  ad('       item.CFOP,                             ');
  ad('       clie.CID_CODIGO,                       ');
  ad('       0 As ALIQUOTA_IPI                      ');
  ad('  from NotaItem   item,                       ');
  ad('       NotaFiscal nota,                       ');
  ad('       Tabela     prod,                       ');
  ad('       CLIENTE    clie                        ');
  ad(' where item.NumeroNF   = '''+fNumero_nota+''' ');
  ad('   and item.NumeroNF   = nota.NumeroNF        ');
  ad('   and item.CodProduto = prod.Tab_Codigo      ');
  ad('   and nota.CodCliente = clie.CLI_CODI        ');
  fQn.Open;

  result := not fQn.IsEmpty;

  if not fQn.IsEmpty then begin
    TNumericField(fQn.FieldByName('Vlrtotal')).DisplayFormat := '#,###,##0.0000';
    TNumericField(fQn.FieldByName('VlrUnitario')).DisplayFormat := '#,###,##0.0000';
  end;
end;

procedure TNFEInterfaceV3.SetValidadeCertificado(pDias: string);
begin
  KontNfe.ValidadeCertificado := pDias;
end;

function TNFEInterfaceV3.FileExists(pArquivo: string): Boolean;
var
  comando      : AnsiString;
  ret          : TRetorno;
  vRetorno     : Boolean;
  i            : integer;
  vMensagem    : string;
  retorno      : Boolean;
begin
  try
    result := false;
    ret.status := TStringList.Create;
    comando := 'NFe.FileExists("' + pArquivo + '")' ;

    try
      if not UtilizaAcbrMonitor then begin
        if frmAcbrNFe = nil then
          frmAcbrNFe := TfrmAcbrNFe.Create(nil);

        retorno := frmAcbrNFe.FileExiste(pArquivo,ret);
      end else begin
        retorno := KontNfe.EnviarComando(comando,ret);
      end;
    except
      retorno := false;
    end;

    if retorno then begin
      for I := 0 to ret.status.Count - 1 do begin
        vMensagem := ret.status[i]+#13;
        if Copy(vMensagem,1,2) = 'OK' then begin
          Result   := True;
          Break;
        end;
      end;
    end;
  except
    result := false;
  end;
end;

function TNFEInterfaceV3.ValidarNfe(pEnderecoXML: String): Boolean;
var
  vMensagem    : string;
  comando      : AnsiString;
  i            : Integer;
  ret          : TRetorno;
  retorno      : Boolean;
begin
  try
    result := False;
    vMensagem := '';
    AbreSplashNFe('Aguarde, Validando ' + fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55));

    ret.status := TStringList.Create;
    comando := 'NFe.validarnfe("' + pEnderecoXML + '")' ;

    try
      if not UtilizaAcbrMonitor then begin
        if frmAcbrNFe = nil then
          frmAcbrNFe := TfrmAcbrNFe.Create(nil);

        retorno := frmAcbrNFe.validarnfe(pEnderecoXML, ret);
      end else begin
        retorno := KontNfe.EnviarComando(comando,ret);
      end;
    except
      retorno := false;
    end;

    if retorno then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          vMensagem := ret.status[i]+#13;
          if (Copy(vMensagem,1,2) = 'OK')then begin
            vMensagem := 'NFe Validada com sucesso!'+#13;
            result := true;
            Break;
          end
          else if (Copy(vMensagem,1,4) = 'ERRO') then begin
            result := False;
            Break;
          end;
        end;
      end
      else begin
        vMensagem := 'NFe não Validada! Tente novamente!'+#13;
      end;
    end
    else begin
      vMensagem := 'NFe não Validada! Tente novamente!'+#13;
    end;
  finally
    FechaSplashNFe;
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor)  then begin
    vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
  end
  else begin
    vMensagem := vMensagem + #13 +
                 'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                 'UF................: '  +#9+KontNfe.UF+#13+
                 'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                 'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                 'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                 'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                 'Versão.........: '     +#9+KontNfe.VersaoLayout;
  end;

  if not fNFCE then
    _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');

end;

function TNFEInterfaceV3.SetCertificado(pCertificado: string): boolean;
var
  comando   : AnsiString;
  i         : Integer;
  ret       : TRetorno;
  retorno   : Boolean;
begin
  Result := false;
  comando := 'NFe.SetCertificado("' + pCertificado + '")';
  ret.status := TStringList.Create;
  try
    if not UtilizaAcbrMonitor then begin
      if frmAcbrNFe = nil then
        frmAcbrNFe := TfrmAcbrNFe.Create(nil);

      retorno := frmAcbrNFe.SetCertificado(pCertificado,'', ret);
    end else begin
      retorno := KontNfe.EnviarComando(comando,ret);
    end;
  except
    retorno := False;
  end;

  if retorno then begin
    for I := 0 to ret.status.Count - 1 do begin
      if Copy(ret.status[i],1,2) = 'OK' then begin
        result := True;
      end;
    end;
  end;
end;

procedure TNFEInterfaceV3.SetFinalidade(pValue: Integer);
begin
  fCapaNFe.Finalidade := pValue;
end;

procedure TNFEInterfaceV3.ValidarNfe(pTipoXml: TeeEnvV3);
begin
  try
//    AssinarXML(pTipoXml);
  finally

  end;
end;

function TNFEInterfaceV3.Ativo: Boolean;
var
  vMensagem : string;
  comando : AnsiString;
  ret     : TRetorno;
  i       : integer;
  retorno : Boolean;
begin
  comando    := 'NFe.Ativo';
  ret.status := TStringList.Create;
  result     := false;

  try
    if not UtilizaAcbrMonitor then begin
      retorno := true;
      ret.status.add('OK: Ativo');
    end else begin
      retorno := KontNfe.EnviarComando(comando,ret);
    end;
  except
    retorno := false;
  end;

  if retorno then begin
    for I := 0 to ret.status.Count - 1 do begin
      vMensagem := ret.status[i]+#13;
      if Pos('OK: Ativo',vMensagem) > 0 then begin
        Result   := True;
        Break;
      end;
    end;
  end;
end;

function TNFEInterfaceV3.CriarNFe(pRetornaXML  : Integer;
                                  pNumeroNf    : AnsiString;
                                  var pEnderecoXML : string;
                                  var pXML : AnsiString;
                                  pModelo  : string = '55';
                                  pSerie   : string ='1'): Boolean;
//nRetornaXML - Coloque o valor 1 se quiser que o ACBrNFeMonitor
//retorne além do Path de onde o arquivo foi criado, o XML gerado.
//Por default não retorna o XML.
var
  vMensagem    : string;
  comando      : AnsiString;
  i            : Integer;
  ret          : TRetorno;
  vRetorno     : Boolean;
  Retorno      : Boolean;
  vStatus      : String;
begin
  vMensagem := '';
  vStatus   := '';
  Result   := false;
  vRetorno := False;
//  fCapaNfe.codigo := StrToInt(pNumeroNf);
//  fCapaNfe.Numero := StrToInt(pNumeroNf);
  ret.status := TStringList.Create;
  try
    try
      if not UtilizaAcbrMonitor then begin
        if frmAcbrNFe = nil then
          frmAcbrNFe := TfrmAcbrNFe.Create(nil);
        if fCapaNFe.ValidaNota then
          retorno := frmAcbrNFe.CriarNFe(ret,fCapaNFe.GetTextoIni,pModelo,fVersaoNfe);
      end
      else begin
        comando  := 'NFe.CriarNFe("' + fCapaNFe.GetTextoIni + '",' + IntToStr(pRetornaXML) + ')' ;
        Retorno := KontNfe.EnviarComando(comando,ret);
      end;
    except
      on e: exception do begin
        Atencao(e.Message);
        Retorno := False;
      end;
    end;

    if Retorno then begin
      for I := 0 to ret.status.Count - 1 do begin
       // vMensagem := ret.status[i];
        if Copy(ret.status[i],1,2) = 'OK' then begin
          Result   := True;
          pEnderecoXML := Copy(ret.status[i],5,Length(ret.status[i])-4);
          vMensagem := '';
        end
        else if Copy(ret.status[i],1,1) = '<' then begin
          Result := True;
          pXML   := Copy(ret.status[i],1,Length(ret.status[i]));
        end
        else if Copy(ret.status[i],1,4) = 'ERRO' then begin
          Result := False;
          pEnderecoXML := '';
          vMensagem := Copy(ret.status[i],6,length(ret.status[i])-5);
          Break;
        end
        else if Pos('VERSAO=',UpperCase(ret.status[i])) > 0 then begin
          KontNfe.VersaoLayout := Copy(ret.status[i],8,Length(ret.status[i])-7);
        end
        else if Pos('CSTAT=',UpperCase(ret.status[i])) > 0 then begin
          vStatus := Copy(ret.status[i],7,Length(ret.status[i])-6);
        end
        else if Pos('Parar=True',ret.status[i]) > 0 then begin
          vMensagem := '';
          Result := True;
        end
        else if Pos('Finalizou=true',ret.status[i]) > 0 then begin
          vMensagem := '';
        end;
      end;
    end;
    if (vMensagem = '') and (not result) and (UtilizaAcbrMonitor) then begin
      vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
    end
    else begin
      if vMensagem <> '' then
        vMensagem := vMensagem + #13 +
                   'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                   'UF................: '  +#9+KontNfe.UF+#13+
                   'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                   'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                   'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                   'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                   'Versão.........: '     +#9+KontNfe.VersaoLayout;
    end;

    if (vMensagem <> '') then
      _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');

  finally
    FechaSplashNFe;
    {Atualizar a tabela de historico}
    if Result then
      GravaHistoricoNfe(pNumeroNf,
                        StrToIntDef(vStatus,0),
                        pXML,
                        true,
                        pModelo,
                        pSerie);
  end;
end;

function TNFEInterfaceV3.EnviarLoteNfe(pLote: string; var pRecibo: String; var pStatusRetorno: String; var pListaNFE : TList): boolean;
var
  vMensagem    : string;
  comando      : AnsiString;
  i, j         : Integer;
  ret          : TRetorno;
  vRetorno     : Boolean;
  Retorno      : Boolean;
  vStatus      : String;
  LoteNfe      : TLoteNfe;
  ListaLote    : Tlist;
begin
  vMensagem := '';
  vStatus   := '';
  Result   := false;
  vRetorno := False;

  ListaLote := TList.Create;
  ret.status := TStringList.Create;
  comando    := 'NFe.EnviarLoteNfe(' + pLote + ')' ;

  try
    try
      if not UtilizaAcbrMonitor then begin
//        if frmAcbrNFe = nil then
//          frmAcbrNFe := TfrmAcbrNFe.Create(nil);
//
//        retorno := frmAcbrNFe.EnviarLoteNFE(pLote, ret);
        retorno := false;
        vMensagem := 'Para EnviarLoteNFE utilize o ACBrMonitor!';

      end else begin
        Retorno := KontNfe.EnviarComando(comando,ret);
      end;
    except
      Retorno := False;
    end;

    if Retorno then begin
      for I := 0 to ret.status.Count - 1 do begin
        if not vRetorno then begin
          if Copy(ret.status[i],1,2) = 'OK' then begin
            Result   := True;
            vMensagem := '';
          end
          else if Copy(ret.status[i],1,9) = '[RETORNO]' then begin
            vRetorno := true;
            Continue;
          end;
        end;
        if vRetorno then begin
          if Pos('CSTAT=',UpperCase(ret.status[i])) > 0 then begin
            pStatusRetorno := Copy(ret.status[i],7,Length(ret.status[i])-6);
            continue;
          end
          else if Copy(ret.status[i],1,4) = 'ERRO' then begin
            Result := False;
            vMensagem := Copy(ret.status[i],6,length(ret.status[i])-5);
            Break;
          end
          else if copy(ret.status[i],1,5) = 'NRec=' then begin
            pRecibo := copy(ret.status[i],6,Length(ret.status[i]));
          end
          else if copy(ret.status[i],1,8) = 'XMotivo=' then begin
            vMensagem := copy(ret.status[i],9,Length(ret.status[i]));
          end;
          if copy(ret.status[i],1,4) = '[NFE' then begin
            for j := i to ret.status.Count - 1 do begin
              if copy(ret.status[j],1,4) = '[NFE' then begin
                LoteNfe := TLoteNFe.Create;
                LoteNfe.fNFE := SomenteNumeros(ret.status[j]);
              end
              else if copy(ret.status[j],1,5) = 'CStat' then begin
                LoteNfe.CStat := Copy(ret.status[j],7,length(ret.status[j]));
              end
              else if copy(ret.status[j],1,7) = 'XMotivo' then begin
                LoteNfe.XMotivo := Copy(ret.status[j],9,length(ret.status[j]));
              end
              else if copy(ret.status[j],1,5) = 'ChNFe' then begin
                LoteNfe.ChNFe := Copy(ret.status[j],7,length(ret.status[j]));
              end
              else if copy(ret.status[j],1,8) = 'DhRecbto' then begin
                LoteNfe.DhRecbto := Copy(ret.status[j],10,length(ret.status[j]));
              end
              else if copy(ret.status[j],1,5) = 'NProt' then begin
                LoteNfe.NProt := Copy(ret.status[j],7,length(ret.status[j]));
              end
              else if copy(ret.status[j],1,6) = 'DigVal' then begin
                LoteNfe.DigVal := Copy(ret.status[j],8,length(ret.status[j]));
              end
              else if copy(ret.status[j],1,7) = 'Arquivo' then begin
                LoteNfe.Arquivo := Copy(ret.status[j],9,length(ret.status[j]));
                ListaLote.Add(LoteNfe);
              end;
            end;
            Break;
          end;
        end;
      end;
     // if ListaLote.Count > 0 then
      pListaNFE := ListaLote;
    end;
  finally
    //
  end;
end;

function TNFEInterfaceV3.AdicionarNFe(pRetornaXML  : Integer;
                                      pNumeroNf    : AnsiString;
                                      var pEnderecoXML : string;
                                      var pXML : AnsiString;
                                      pLote    : String;
                                      pModelo  : string = '55';
                                      pSerie   : string ='1'): Boolean;
//nRetornaXML - Coloque o valor 1 se quiser que o ACBrNFeMonitor
//retorne além do Path de onde o arquivo foi criado, o XML gerado.
//Por default não retorna o XML.
var
  vMensagem    : string;
  comando      : AnsiString;
  i            : Integer;
  ret          : TRetorno;
  vRetorno     : Boolean;
  Retorno      : Boolean;
  vStatus      : String;
begin
  vMensagem := '';
  vStatus   := '';
  Result   := false;
  vRetorno := False;

  ret.status := TStringList.Create;
  comando    := 'NFe.AdicionarNFe("' + fCapaNFe.GetTextoIni + '",' +
                                       pLote + ')' ;

  try
    try
      Retorno := KontNfe.EnviarComando(comando,ret);
    except
      Retorno := False;
    end;

    if Retorno then begin
      for I := 0 to ret.status.Count - 1 do begin
       // vMensagem := ret.status[i];
        if Copy(ret.status[i],1,2) = 'OK' then begin
          Result   := True;
          pEnderecoXML := Copy(ret.status[i],5,Length(ret.status[i])-4);
          vMensagem := '';
        end
        else if Copy(ret.status[i],1,1) = '<' then begin
          Result := True;
          pXML   := Copy(ret.status[i],1,Length(ret.status[i]));
        end
        else if Copy(ret.status[i],1,4) = 'ERRO' then begin
          Result := False;
          pEnderecoXML := '';
          vMensagem := Copy(ret.status[i],6,length(ret.status[i])-5);
          Break;
        end
        else if Pos('VERSAO=',UpperCase(ret.status[i])) > 0 then begin
          KontNfe.VersaoLayout := Copy(ret.status[i],8,Length(ret.status[i])-7);
        end
        else if Pos('CSTAT=',UpperCase(ret.status[i])) > 0 then begin
          vStatus := Copy(ret.status[i],7,Length(ret.status[i])-6);
        end
        else if Pos('Parar=True',ret.status[i]) > 0 then begin
          vMensagem := '';
          Result := True;
        end
        else if Pos('Finalizou=true',ret.status[i]) > 0 then begin
          vMensagem := '';
        end;
      end;
    end;
//    if (vMensagem = '') and (not result) then begin
//      vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
//    end
//    else begin
//      if vMensagem <> '' then
//        vMensagem := vMensagem + #13 +
//                   'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
//                   'UF................: '  +#9+KontNfe.UF+#13+
//                   'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
//                   'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
//                   'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
//                   'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
//                   'Versão.........: '     +#9+KontNfe.VersaoLayout;
//    end;
//
//    if (vMensagem <> '') then
//      _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');

  finally
//    FechaSplashNFe;
    {Atualizar a tabela de historico}
    if Result then
      GravaHistoricoNfe(pNumeroNf,
                        StrToIntDef(vStatus,0),
                        pXML,
                        true,
                        pModelo,
                        pSerie);
  end;
end;

//function TNFEInterfaceV3.PausarAutomatico: boolean;
//var
//  tIni : string;
//  tFim : string;
//begin
//  try
//    result := true;
//    //se a diferença for maior do que 5 segundos tem que enviar off line
//    tIni := formatdatetime('hhmmss',Now);
//    if bUtilizaNFCE and bEnvioAutomaticoHabilitado then begin
//      fPausarAutomatico := true;
//      while not fPausadoAuto do begin
//        tFim := formatdatetime('hhmmss',Now);
//        if (StrToInt(tFim) - StrToInt(tIni)) >= 3 then begin
//          result := false;
//          break;
//        end;
//        sleep(500);
//      end;
//    end;
//  except
//    fPausarAutomatico := false;
//  end;
//end;

//procedure TNFEInterfaceV3.DespausarAutomatico;
//begin
//  try
//    if bUtilizaNFCE and bEnvioAutomaticoHabilitado then begin
//      fPausarAutomatico := false;
//      while fPausadoAuto do begin
//        sleep(100);
//      end;
//    end;
//  except
//    fPausarAutomatico := false;
//  end;
//end;

function TNFEInterfaceV3.EnviarNFe(pEnderecoXML    : String;
                                   var pChNFe      : string;
                                   var pMotivo     : string;
                                   var pRecibo     : string;
                                   var pProtocolo  : string;
                                   pGerenciadorNFE : Boolean = False;
                                   pModelo         : string = '55';
                                   pSerie          : string = '1';
                                   pSplash         : Boolean = true;
                                   pAutomatico     : Boolean = false;
                                   pXML            : Boolean = false): Boolean;
var
  vMensagem    : string;
  vMsg         : AnsiString;
  comando      : AnsiString;
  i            : Integer;
  ret          : TRetorno;
  vRetorno     : Boolean;
  Retorno      : Boolean;
  vStatus      : Integer;
  vNumNFe      : String;
  vTexto       : String;
  vPath        : String;
  j            : dword;
begin
  try
    fXML := '';
    vMensagem := '';
    result    := False;
    vRetorno  := false;

    fModelo := iif(copy(trim(somentenumeros(pEnderecoXML)),21,2) = '55','NF-e','NFC-e');

    if trim(fCampos.CupomFiscal) <> '' then
      vNumNFe := fCampos.CupomFiscal
    else begin
      if (trim(pChNFe) <> '') and (Length(trim(pChNFe))= 44) then
        vNumNFe := copy(trim(pChNFe),26,9)
      else begin
        if Length(trim(somentenumeros(pEnderecoXML))) = 44 then
          vNumNFe := copy(trim(somentenumeros(pEnderecoXML)),26,9)
        else
          vNumNFe := '';
      end;
    end;

    if pSplash then
      AbreSplashNFe('Aguarde, Enviando ' + fModelo + ' - ' + vNumNFe, IIfInt(fModelo = 'NFC-e',65,55));

    ret.status := TStringList.Create;
    comando := 'NFe.EnviarNFe("' + pEnderecoXML + '",' + '1,1,0' + ')' ;

    try
      if 1=2 then begin //Utilizando o KontServerDFE

        result := dm.EnviarNfe(fCampos.CupomFiscal, fCampos.CupomFiscal,StrToInt(pModelo),StrToInt(pSerie), vMsg );

        ret.status.Clear;

        for i := 0 to Length(vMsg) do begin
          if pos(sLineBreak,vMsg) > 0 then begin
            ret.status.add(copy(vMsg,1,pos(SlineBreak,vmsg)-1));
            Delete(vMsg,1,pos(SlineBreak,vmsg)+1);
          end;
        end;
      end
      else if (not UtilizaAcbrMonitor) or pXML then begin
        if frmAcbrNFe = nil then
          frmAcbrNFe := TfrmAcbrNFe.Create(nil);


        //verifica se tem internet
//        IF InternetGetConnectedState(@j,0) then begin
        try
          if frmAcbrNFe.pingIp('google.com') then begin
            if not pXML then begin
              if frmAcbrNFe.LerXML(pEnderecoXML) then
                retorno := frmAcbrNFe.EnviarNFCe(ret,fversaonfe);
            end else begin
              if frmAcbrNFe.LeXML(pEnderecoXML) then
                retorno := frmAcbrNFe.EnviarNFCe(ret,fversaonfe);
            end;
          end
          else begin
            retorno := true;
            ret.status.add('[RETORNO]');
            ret.status.add('XMotivo=Sem conexão com internet!');
            ret.status.add('CSTAT=999');
          end;
        except
          retorno := true;
          ret.status.add('[RETORNO]');
          ret.status.add('XMotivo=Sem conexão com internet!');
          ret.status.add('CSTAT=999');
        end;
      end
      else begin
        retorno := KontNfe.EnviarComando(comando,ret,2000,True,pAutomatico);
      end;
    except
      on e: exception do begin
        retorno := false;
        showmessage(e.message);
      end;
    end;

    if retorno then begin
      if ret.status.Count > 0  then begin
        for I := 0 to ret.status.Count - 1 do begin
          vTexto := UpperCase(ret.status[i]);
          if Pos('[RETORNO]',vTexto) > 0 then begin
            vRetorno := True;
          end;
          if Pos('ERRO',vTexto) > 0 then begin
            Result := false;
            vMensagem := Trim(copy(ret.status[i],Pos('ERRO',ret.status[i])+5,length(ret.status[i]))) + #13;
            if Pos('Inativo ou Inoperante tente novamente',Trim(ret.status[i+1])) > 0  then
              vMensagem := vMensagem + 'Inativo ou Inoperante tente novamente!' + #13;
            if Length(vMensagem) <=2 then
              vMensagem := 'Erro Interno: 12152 - Requisição não enviada.';
            pMotivo := 'E'; //ENVIAR EM OFF-LINE NFCE
            if Pos('XMotivo=',ret.status[i]) > 0 then begin
              vMensagem := Copy(ret.status[i],9,Length(ret.status[i]))+#13;
            end;
            break;
          end;
          if Pos('CSTAT=999',vTexto) > 0 then begin
            result := False;
            vStatus := 999;
          end;
          if Pos('XMOTIVO=',vTexto) > 0 then begin
            vMensagem := Copy(ret.status[i],9,Length(ret.status[i]))+#13;
          end;
          if vRetorno then begin
            if Pos('XMOTIVO=',vTexto) > 0 then begin
              vMensagem := Copy(ret.status[i],9,Length(ret.status[i]))+#13;
            end;
            if Pos('CSTAT=100',vTexto) > 0 then begin
              result := true;
              vStatus := 100;
            end;
            if Pos('CHNFE=',vTexto) > 0 then begin
              pChNFe := Copy(ret.status[i],7,Length(ret.status[i]));
            end;
            if (Pos('CSTAT=204',vTexto) > 0) then begin //Rejeicao: Duplicidade de NF-e
              result := false;
              pMotivo := 'D';
            end;
            if (Pos('CSTAT=539',vTexto) > 0) then begin //Rejeicao: Duplicidade de NF-e com diferença na Chave de Acesso [chNFe: 99999999999999999999999999999999999999999999][nRec:999999999999999]'; end;
              result := false;
              pMotivo := 'DD';
            end;
            if (Pos('CSTAT=704',vTexto) > 0) then begin //Rejeicao: 'Rejeição: NFC-e com Data-Hora de emissão atrasada'
              result := false;
              pMotivo := 'DH';
            end;
            if Pos('CSTAT=206',vTexto) > 0 then begin // NF-e já está inutilizada na Base de dadosda SEFAZ
              result := false;
              pMotivo := 'I';
            end;
            if Pos('CSTAT=105',vTexto) > 0 then begin // CStat=105 XMotivo=Lote em processamento
              result := false;
              pMotivo := 'L';
            end;
            if Pos('CSTAT=301',vTexto) > 0 then begin // Uso Denegado = Irregularidade fiscal do emitente
              result := false;
              pMotivo := 'NE';
            end;
            if Pos('CSTAT=302',vTexto) > 0 then begin // Uso Denegado = Irregularidade fiscal do DESTINATÁRIO
              result := false;
              pMotivo := 'ND';
            end;
            if Pos('CSTAT=205',vTexto) > 0 then begin // Uso Denegado = NF-e  está  denegada  na  base  de  dados  da  SEFAZ
              result := false;
              pMotivo := 'NS';
            end;
            if Pos('CSTAT=213',vTexto) > 0 then begin // CNPJ-Base do Emitente difere do CNPJ-Base do Certificado Digital
              result := false;
              pMotivo := 'R';   //ESTE MOTIVO É PRA REENVIAR O MESMO XML, CASO REPITA 3 VEZES ENVIAR EM OFF-LINE
            end;
            if Pos('CSTAT=806',vTexto) > 0 then begin // XMotivo=Rejeição: Operação com ICMS-ST sem informação do CEST
              result := false;
              pMotivo := 'C';
            end;
            if Pos('VERAPLIC=',vTexto) > 0 then begin
              KontNfe.VersaoLayout := copy(vTexto, Pos('VERAPLIC=',vTexto)+10, Length(vTexto) );
            end;
            if Pos('CUF=',vTexto) > 0 then begin
              KontNfe.UF := copy(vTexto,Pos('CUF=',vTexto)+4,Length(vTexto));
            end;
            if Pos('NREC=',vTexto) > 0 then begin
              pRecibo := Copy(vTexto,Pos('NREC=',vTexto)+5,Length(vTexto));
            end;
            if Pos('NPROT=',vTexto) > 0 then begin
              pProtocolo := somenteNumeros(ret.status[i]);
            end;
            if Pos('VERSAO=',vTexto) > 0 then begin
              KontNfe.VersaoLayout := Copy(ret.status[i],8,Length(ret.status[i])-7);
            end;
            if Pos('CSTAT=',vTexto) > 0 then begin
              vStatus := StrToInt(SOMENTENUMEROS(Copy(ret.status[i],7,Length(ret.status[i])-6)));
            end;
            if Pos('[NFE',vTexto) > 0 then begin
              vNumNFe := somentenumeros(ret.status[i]);
            end;
            if Pos('XML',vTexto) > 0 then begin
              fXML := copy(vTexto,Pos('XML=',vTexto)+4,Length(vTexto));
            end;
          end;
        end;
      end
      else begin
        vMensagem := fModelo + ' não enviada! Tente novamente!'+#13;
      end;
    end
    else begin
      vMensagem := fModelo + ' não enviada! Tente novamente!'+#13;
    end;
  finally
    if pSplash then
      FechaSplashNFe;
  end;

  //RESPOSTA DO RECEBIDA DA SEFAZ DEVERÁ SER GRAVADA NO BANCO
  if not UtilizaAcbrMonitor then begin
    if ret.status.Count = 0  then begin
      ret.status.add('Sem conteúdo!');
    end;
    vPath := ExtractFilePath(Application.ExeName) + 'arquivos\' + IIf(fModelo = 'NFC-e','NFCE\','NFE\') +
             FormatDateTime('YYYY',Now) + '\' +
             FormatDateTime('MM',Now) + '\' +
             FormatDateTime('DD',Now) + '\';
    if not DirectoryExists(vPath) then
      ForceDirectories(vPath);

    if trim(pChNFe) = '' then begin
      pChNFe := trim(somentenumeros(pEnderecoXML));
    end;

    vPath := vPath + IIf(fModelo = 'NFC-e','NFCE','NFE') + copy(trim(pChNFe),26,9) + '.txt';
    ret.status.savetofile(vPath);
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor ' + fModelo + ' não respondendo, tente novamente!';
  end
  else begin
    if trim(vNumNFe) = '' then
      vNumNFe := copy(trim(pChNFe),26,9);

    if trim(vNumNFe) <> '' then
      GravaHistoricoNfe(vNumNFe,
                        vStatus,
                        '',
                        True,
                        pModelo,
                        pSerie);
    if Result then begin
      if (trim(vNumNFe) <> '') and (Length(trim(pChNFe)) = 44) then begin
        IF StrToInt(trim(vNumNFe)) = STRTOINT(COPY(pChNFe,26,9)) then
          SalvaRetornoXML(pEnderecoXML,
                          pChNFe,
                          vNumNFe,
                          pModelo,
                          pSerie);
      end;
    end;

    if pSplash then
      vMensagem := vMensagem + #13 +
                   'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                   'UF................: '  +#9+KontNfe.UF+#13+
                   'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                   'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                   'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                   'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                   'Versão.........: '     +#9+KontNfe.VersaoLayout;
  end;

  if pSplash then begin
    if (pMotivo <> 'E') AND (not fNFCE) then
      _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'')
    else if pGerenciadorNFE then
      _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'')
  end;

  try
     if trim(vNumNFe) <> '' then
      Validada(IIf(result,'V','X'),StrToInt(vNumNFe),pModelo,pSerie);
  except
  end;

//  Result := FALSE;
//  pMotivo := 'E';

end;

function TNFEInterfaceV3.CriarNFce(pRetornaXML  : Integer;
                                   pCupomFiscal : AnsiString): Boolean;
//nRetornaXML - Coloque o valor 1 se quiser que o ACBrNFeMonitor
//retorne além do Path de onde o arquivo foi criado, o XML gerado.
//Por default não retorna o XML.
var
  vMensagem    : string;
  comando      : AnsiString;
  i            : Integer;
  ret          : TRetorno;
  vCriado      : Boolean;
  vEnderecoXML : string;
  vRetorno     : Boolean;
begin
  Result := false;
  vRetorno := False;
  vCriado  := false;
  fCapaNfe.codigo        := StrToInt(pCupomFiscal);
  fCapaNfe.Numero        := StrToInt(pCupomFiscal);
  comando := 'NFe.CriarNFe("' + fCapaNFe.GetTextoIni + '",' +
                              IntToStr(pRetornaXML) + ')' ;

  AbreSplashNFe('Aguarde, Enviando NFC-e...',65);
  ret.status := TStringList.Create;
  try
    {Consulta}
    if KontNfe.EnviarComando(comando,ret) then begin
      for I := 0 to ret.status.Count - 1 do begin
        vMensagem := ret.status[i]+#13;
        if Copy(vMensagem,1,2) = 'OK' then begin
          vCriado   := True;
          vEnderecoXML := Copy(vMensagem,5,Length(vMensagem)-5);
          vMensagem := '';
          Break;
        end
        else if Copy(vMensagem,1,4) = 'ERRO' then begin
          vCriado := False;
          vEnderecoXML := '';
          Break;
        end;
      end;

      if vCriado then begin
        comando := 'NFe.EnviarNFe("' + vEnderecoXML + '",' + '1,1,0' + ')' ;
        if KontNfe.EnviarComando(comando,ret) then begin
          if ret.status.Count > 0  then begin
            for I := 0 to ret.status.Count - 1 do begin
              vMensagem := ret.status[i]+#13;
              if Pos('[RETORNO]',vMensagem) > 0 then begin
                vRetorno := True;
              end;
              if (Copy(vMensagem,1,2) = 'OK') and (not vRetorno) then begin
                vMensagem := ret.status[i+1]+#13;
                if Copy(vMensagem,1,2) = 'OK' then begin
                  result := true;
                  Break;
                end;
              end
              else if (Copy(vMensagem,1,4) = 'ERRO') and (not vRetorno) then begin
                result := False;
                vEnderecoXML := '';
                Break;
              end
              else if vRetorno then begin
                if Pos('XMOTIVO=REJEICAO',UpperCase(vMensagem)) > 0 then begin
                  result := False;
                  vMensagem := Copy(vMensagem,9,Length(vMensagem)-9)+#13;
                  Break;
                end;
                if Pos('CSTAT=100',UpperCase(vMensagem)) > 0 then begin
                  result := true;
                  vMensagem := 'Autorizado o uso da NFC-e!' + #13;
                  break;
                end;
              end;
            end;
          end
          else begin
            vMensagem := 'Não foi Enviado a NFC-e! Tente novamente!'+#13;
          end;
        end
        else begin
          vMensagem := 'Não foi Enviado a NFC-e! Tente novamente!'+#13;
        end;
      end
      else begin
        vMensagem := 'Não foi Enviado a NFC-e! Tente novamente!'+#13;
        result := false;
      end;
    end;
  finally
    FechaSplashNFe;
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor NFce não respondendo, tente novamente!';
  end
  else begin
    vMensagem := vMensagem +
                 'Empresa.......: ' +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                 'UF............: ' +#9+KontNfe.UF+#13+
                 'Codigo UF.....: ' +#9+KontNfe.CodUF+#13+
                 'Ambiente......: ' +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                 'Contigência...: ' +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                 'Validade......: ' +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                 'Versão........: ' +#9+KontNfe.VersaoLayout;
  end;

  _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');

end;

function TNFEInterfaceV3.EnviarEmail(cEmailDestino,cArqXML,cArqPDF: string; pSplash: boolean = True): Boolean;
var
  comando   : AnsiString;
  ret       : TRetorno;
  i         : integer;
  vMensagem : string;
  retorno   : Boolean;
  MSG       : TStrings;
begin
  vMensagem := '';
  comando := 'NFe.EnviarEmail("' + cEmailDestino + '","' + cArqXML + '",1)';
  if pSplash then
    AbreSplashNFe('Aguarde, Enviando e-mail da ' + fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55));
  ret.status := TStringList.Create;

  try
//  if not UtilizaAcbrMonitor then begin
      if frmAcbrNFe = nil then
        frmAcbrNFe := TfrmAcbrNFe.Create(nil);

      retorno := frmAcbrNFe.EnviaEmail(cEmailDestino,cArqXML,cArqPDF,ret);
//
//    end else begin
//      retorno := KontNfe.EnviarComando(comando,ret);
//    end;
    if pSplash then
      FechaSplashNFe;
  except
    if pSplash then
      FechaSplashNFe;
    retorno := false;
  end;
//  END;

  if retorno then begin
    if ret.status.Count > 0  then begin
      for I := 0 to ret.status.Count - 1 do begin
        if (Copy(ret.status[i],1,4) = 'ERRO') then begin
          result := False;
          vMensagem := Copy(ret.status[i],6,Length(ret.status[i])-5)+#13;
          Break;
        end
        else if pos('OK:',ret.status[i]) > 0 then begin
          result := True;
          vMensagem := ret.status[i];
          Break;
        end;
      end;
    end;
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
  end
  else begin
    vMensagem := vMensagem + #13 +
                 'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                 'UF................: '  +#9+KontNfe.UF+#13+
                 'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                 'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                 'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                 'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                 'Versão.........: '     +#9+KontNfe.VersaoLayout;
  end;

  if pSplash then
    _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');


end;

procedure TNFEInterfaceV3.NfeCriarEnviarNFe(pTextoIni      : string;
                                            pNumLote       : Integer;
                                            pImprimirDanfe : Integer;
                                            pSincrono      : Integer);
var
  vMensagem : string;
  comando   : AnsiString;
  i         : Integer;
  ret       : TRetorno;
  retorno   : Boolean;
begin
  vMensagem := '';
  comando := 'NFe.CriarEnviarNFe("' + pTextoIni                + '",' +
                                     IntToStr(pNumLote)       + ',' +
                                     IntToStr(pImprimirDanfe) + ','+
                                     IntToStr(pSincrono)      + ')' ;

  AbreSplashNFe('Aguarde, Enviando ' + fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55));
  ret.status := TStringList.Create;
  try
    retorno := KontNfe.EnviarComando(comando,ret);
    FechaSplashNFe;
  except
    FechaSplashNFe;
  end;

  if retorno then
    for I := 0 to ret.status.Count - 1 do
      vMensagem := Ret.status[i]+#13;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
  end
  else begin
    vMensagem := vMensagem + #13 +
                 'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                 'UF................: '  +#9+KontNfe.UF+#13+
                 'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                 'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                 'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                 'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                 'Versão.........: '     +#9+KontNfe.VersaoLayout;
  end;

  _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');

end;

procedure TNFEInterfaceV3.SalvaRetornoXML(pArquivo: string; pChaveNFE: string; pNumeroNota: String; pModelo: string; pSerie: string);
var
  vArquivo : string;
  pXml     : AnsiString;
begin
  try
    if pos('\ARQUIVOS\',UpperCase(pArquivo)) > 0 then begin
      vArquivo := copy(pArquivo,1,pos('\ARQUIVOS\',UpperCase(pArquivo))) +
                  'Arquivos\' + IIf(fNFCE,'NFCe','Nfe') +'\' + pChaveNFE + '-nfe.xml';

      if FileExists(vArquivo) then begin
        Busca_XML(vArquivo, pXml);
        if trim(pXml) <> '' then
          GravaRetornoXML(pNumeroNota,pXml,pModelo,pSerie);
      end;
    end;
  except
    //
  end;
end;

function TemXmlRetornoBD(pKONT_NUMERO_NOTA : string; var pXML: AnsiString; pModelo: String; pSerie: String): Boolean;
var
  QR      : TADOQuery;
  xHistId : Integer;
  vXML    : WideString;
begin
  try
    Result := False;

    QR            := TADOQuery.Create(Application);
    QR.connection := dm.adoconexao;
    QR.Close;
    QR.SQL.Clear;

    Try
      Qr.Sql.Clear;
      Qr.Sql.add('SELECT XML_RETORNO                                   ');
      Qr.Sql.add('  FROM NFE_HISTORICOXML                              ');
      Qr.Sql.add(' WHERE EMPRESA_ID = ''' + SomenteNumeros(sEmpresaCNPJ)+'''');
      Qr.Sql.add('   AND NOTA_FISCAL     = ''' + pKONT_NUMERO_NOTA     +'''');
      QR.sql.add('   AND MODELO          = ''' + pModelo + '''');
      QR.SQL.Add('   AND SERIE           = ''' + pSerie  + '''');
      Qr.Open;

      if Length(TRIM(QR.FieldByName('XML_RETORNO').AsString)) > 0 THEN begin
        pXML := TRIM(QR.FieldByName('XML_RETORNO').AsString);
        Result := True;
      end;

    finally
      Qr.Close;
      Qr.Sql.Clear;
    end;
  except
    Result := False;
  end;
end;

procedure GravaRetornoXML(pKONT_NUMERO_NOTA : string;
                          pXml_Retorno      : WideString = '';
                          pModelo           : string     = '55';
                          pSerie            : string     = '1');
var
  QR : TADOQuery;
  xHistId : Integer;
  vXML : WideString;
begin
  try
    QR            := TADOQuery.Create(Application);
    QR.connection := dm.adoconexao;
    QR.Close;
    QR.SQL.Clear;

    vXML := pXml_Retorno;

    Try
      xHistId:=0;

      Qr.Sql.Clear;
      Qr.Sql.add('SELECT HISTORICO_ID, XML_ENVIADO                           ');
      Qr.Sql.add('  FROM NFE_HISTORICOXML                                    ');
      Qr.Sql.add(' WHERE NOTA_FISCAL     = ''' + pKONT_NUMERO_NOTA      +'''');
//      EMPRESA_ID = ''' + SomenteNumeros(sEmpresaCNPJ)+'''');
//      Qr.Sql.add('   AND NOTA_FISCAL     = ''' + pKONT_NUMERO_NOTA      +'''');
      QR.sql.add('   AND MODELO          = ''' + pModelo + '''');
      QR.SQL.Add('   AND SERIE           = ''' + pSerie  + '''');
      Qr.Open;
      if not qr.isempty then begin
        xHistId := Qr.FieldByName('HISTORICO_ID').AsInteger;
      end;
      Qr.Close;
      Qr.Sql.Clear;

      if xHistId <> 0 then begin
        QR.SQL.add('UPDATE NFE_HISTORICOXML           ');
        QR.SQL.add('   SET XML_RETORNO = :XML_RETORNO, ');
        QR.SQL.add('       STATUS_RETORNO = :STATUS_RETORNO ');
        QR.SQL.add(' where NOTA_FISCAL = :NOTA_FISCAL ');
//        Qr.Sql.add('   AND EMPRESA_ID  = :EMPRESA_ID  ');
        qr.sql.add('   AND MODELO      = :MODELO      ');
        QR.SQL.ADD('   AND SERIE       = :SERIE       ');
        QR.Parameters.ParamByName('NOTA_FISCAL').value    := StrToIntDef(pKONT_NUMERO_NOTA,0);
        QR.Parameters.ParamByName('XML_RETORNO').value    := vXML;
//        QR.Parameters.ParamByName('EMPRESA_ID').value     := SomenteNumeros(sEmpresaCNPJ);
        QR.Parameters.ParamByName('STATUS_RETORNO').value := IIF(vXml <> '',100,0);
        QR.Parameters.ParamByName('MODELO').Value         := pModelo;
        QR.Parameters.ParamByName('SERIE').Value          := pSerie;
        QR.ExecSQL;
        QR.Sql.Clear;
      end;
    finally
//    FreeandNil(QR);
    end;
  except
    on e: Exception do begin
      QR.SQL.SaveToFile(ExtractFilePath(Application.ExeName) + 'HistoricoNfe_log.txt');
      if QR <> nil then
        FreeandNil(QR);
    end;
  end;
end;

function TNFEInterfaceV3.GravaInutilizacaoNFe(cCNPJ,
                                               cJustificativa: String;
                                               nAno,
                                               nModelo,
                                               nSerie,
                                               nNumInicial,
                                               nNumFinal: integer;
                                               pXmlEnviado : WideString;
                                               pXmlRetorno : WideString;
                                               pNumeroProtocolo: string;
                                               pDate : TDateTime): boolean;
var
  qr : TADOQuery;
  vXmlEnviado : WideString;
  vXmlRetorno : WideString;
begin
  try
    vXmlEnviado := pXmlEnviado;
    vXmlRetorno := pXmlRetorno;

    result := false;
    qr := TADOQuery.Create(Application);
    qr.connection := dm.adoconexao;
    qr.close;
    qr.sql.clear;
    qr.sql.add('SELECT COUNT(NOTA_INICIO) AS QTDE ');
    QR.SQL.ADD('  FROM NFE_INUTILIZACAO ');
    QR.SQL.ADD(' WHERE MODELO = :MODELO ');
    QR.SQL.ADD('   AND SERIE  = :SERIE');
    QR.SQL.ADD('   AND NOTA_INICIO = :NOTA_INICIO');
    QR.SQL.ADD('   AND EMPRESA_ID = :EMPRESA_ID');
    QR.Parameters.ParamByName('MODELO').Value := nModelo;
    QR.Parameters.ParamByName('SERIE').Value := nSerie;
    QR.Parameters.ParamByName('NOTA_INICIO').Value := nNumInicial;
    QR.Parameters.ParamByName('EMPRESA_ID').Value := cCNPJ;
    QR.Open;

    if qr.FieldByName('QTDE').AsInteger > 0 THEN
      exit;

    qr.close;
    qr.sql.clear;
    qr.sql.add(' INSERT INTO NFE_INUTILIZACAO(EMPRESA_ID,      ');
    qr.sql.add('                              DATA_ENVIO,      ');
    qr.sql.add('                              MOTIVO,          ');
    qr.sql.add('                              ANO,             ');
    qr.sql.add('                              MODELO,          ');
    qr.sql.add('                              SERIE,           ');
    qr.sql.add('                              NOTA_INICIO,     ');
    qr.sql.add('                              NOTA_FIM,        ');
    qr.sql.add('                              XML_ENVIADO,     ');
    qr.sql.add('                              XML_RETORNO,     ');
    qr.sql.add('                              NUMERO_PROTOCOLO)');
    qr.sql.add('                      VALUES(:EMPRESA_ID,      ');
    qr.sql.add('                             :DATA_ENVIO,      ');
    qr.sql.add('                             :MOTIVO,          ');
    qr.sql.add('                             :ANO,             ');
    qr.sql.add('                             :MODELO,          ');
    qr.sql.add('                             :SERIE,           ');
    qr.sql.add('                             :NOTA_INICIO,     ');
    qr.sql.add('                             :NOTA_FIM,        ');
    qr.sql.add('                             :XML_ENVIADO,     ');
    qr.sql.add('                             :XML_RETORNO,     ');
    qr.sql.add('                             :NUMERO_PROTOCOLO)');
    qr.Parameters.ParamByName('EMPRESA_ID').Value       := copy(somentenumeros(cCNPJ),1,18);
    qr.Parameters.ParamByName('DATA_ENVIO').Value       := pDate;
    qr.Parameters.ParamByName('MOTIVO').Value           := copy(cJustificativa,1,255);
    qr.Parameters.ParamByName('ANO').Value              := IntToStr(nAno);
    qr.Parameters.ParamByName('MODELO').Value           := IntToStr(nModelo);
    qr.Parameters.ParamByName('SERIE').Value            := InttoStr(nSerie);
    qr.Parameters.ParamByName('NOTA_INICIO').Value      := InttoStr(nNumInicial);
    qr.Parameters.ParamByName('NOTA_FIM').Value         := InttoStr(nNumFinal);

    qr.Parameters.ParamByName('XML_ENVIADO').DataType := ftString;
    qr.Parameters.ParamByName('XML_ENVIADO').Size     := Length(vXmlEnviado);
    qr.Parameters.ParamByName('XML_ENVIADO').Value    := vXmlEnviado;

    qr.Parameters.ParamByName('XML_RETORNO').DataType := ftString;
    qr.Parameters.ParamByName('XML_RETORNO').Size     := Length(vXmlRetorno);
    qr.Parameters.ParamByName('XML_RETORNO').Value    := vXmlRetorno;

    qr.Parameters.ParamByName('NUMERO_PROTOCOLO').Value := copy(pNumeroProtocolo,1,15);
    qr.ExecSQL;
    result := true;
  except
    on e: exception do begin
      result := false;
      showmessage(e.message);
    end;
  end;
end;

procedure TNFEInterfaceV3.Validada(pVALIDADA: String; iNumeroNf : integer; pModelo, pSerie: string);
var
  QR : TADOQuery;
begin
  TRY
    QR := TADOQuery.Create(Application);
    QR.connection := dm.adoconexao;
    QR.Close;
    QR.SQL.Clear;
    with QR do begin
      Close;
      SQL.Clear;
      SQL.Add('UPDATE NOTAFISCAL');
      SQL.Add('   SET VALIDADA = :VALIDADA');
      SQL.Add(' WHERE NUMERONF = :PNUMERONF');
      SQL.Add('   AND MODELO   = :MODELO');
      SQL.ADD('   AND SERIE    = :SERIE');
      Parameters.ParamByName('VALIDADA').Value  := pVALIDADA;
      Parameters.ParamByName('PNUMERONF').Value := iNumeroNf;
      Parameters.ParamByName('MODELO').Value    := pMODELO;
      Parameters.ParamByName('SERIE').Value     := pSERIE;
      ExecSQL;
    end;
  except
    ON E: Exception DO begin
     // ShowMessage(E.Message);
    end;
  END;
end;

procedure GravaHistoricoNfe(pKONT_NUMERO_NOTA : string;
                            pSTATUS_RETORNO   : Integer;
                            pXML_ENVIADO      : WideString;
                            pPesquisar        : Boolean = false;
                            pModelo           : string = '55';
                            pSerie            : string ='1');
var
  QR : TADOQuery;
  xHistId : Integer;
  vXML : WideString;
  vNumeroNota : Integer;
begin
  try
    QR            := TADOQuery.Create(Application);
    QR.connection := dm.adoconexao;
    QR.Close;
    QR.SQL.Clear;

    vXML := pXML_ENVIADO;

    vNumeroNota := StrToIntDef(pKONT_NUMERO_NOTA,0);

    if vNumeroNota = 0 then
      exit;

    Try
      xHistId:=0;
      if pPesquisar then begin
        Qr.Sql.Clear;
        Qr.Sql.add('SELECT HISTORICO_ID, XML_ENVIADO                                   ');
        Qr.Sql.add('  FROM NFE_HISTORICOXML                                            ');
        Qr.Sql.add(' WHERE EMPRESA_ID = ''' + SomenteNumeros(sEmpresaCNPJ)+'''');
        Qr.Sql.add('   AND NOTA_FISCAL     = ' + IntToStr(vNumeroNota) );
        Qr.sql.add('   AND MODELO = ''' + pModelo + '''');
        Qr.sql.add('   AND SERIE  = ''' + pSerie  + '''');
        Qr.Open;
//        qr.sql.savetofile('c:\texto.txt');
        if not qr.isempty then begin
          xHistId := Qr.FieldByName('HISTORICO_ID').AsInteger;
//          if trim(Qr.fieldbyname('XML_ENVIADO').asString) <> '' then
//            vXML := QR.FieldByName('XML_ENVIADO').AsWideString;
        end;
        Qr.Close;
        Qr.Sql.Clear;
      end;
      if xHistId = 0 then begin
        QR.SQL.add('INSERT INTO NFE_HISTORICOXML             ');
        QR.SQL.add('           (EMPRESA_ID,                  ');
        QR.SQL.add('            NOTA_FISCAL,                 ');

        if vXML <> '' then
          QR.SQL.add(' XML_ENVIADO, ');

        QR.SQL.add('            STATUS_RETORNO,              ');
        QR.SQL.add('            MODELO,                      ');
        QR.SQL.add('            SERIE,                       ');
        QR.SQL.add('            DATA_PROCESSO)               ');
        QR.SQL.add('VALUES                                   ');
        QR.SQL.add('           (:EMPRESA_ID,                 ');
        QR.SQL.Add('            :NOTA_FISCAL,                ');

        if vXML <> '' then
          QR.SQL.Add(' :XML_ENVIADO,');

        QR.SQL.Add('            :STATUS_RETORNO,             ');
        QR.SQL.add('            :MODELO,                      ');
        QR.SQL.add('            :SERIE,                       ');
        QR.SQL.Add('            :DATA_PROCESSO)              ');
      end else begin
        QR.SQL.add('UPDATE NFE_HISTORICOXML SET              ');
        QR.SQL.add('EMPRESA_ID          = :EMPRESA_ID,       ');
        QR.SQL.add('NOTA_FISCAL         = :NOTA_FISCAL,      ');
        if vXML <> '' then
          QR.SQL.add('XML_ENVIADO       = :XML_ENVIADO,      ');
        QR.SQL.add('STATUS_RETORNO      = :STATUS_RETORNO,   ');
        QR.SQL.add('MODELO              = :MODELO,           ');
        QR.SQL.add('SERIE               = :SERIE             ');
//        QR.SQL.add('DATA_PROCESSO       = :DATA_PROCESSO     ');
        QR.SQL.add('where HISTORICO_ID  = :HISTORICO_ID      ');
      end;

      if xHistId > 0 then
        QR.Parameters.ParamByName('HISTORICO_ID').value := xHistId;

      QR.Parameters.ParamByName('EMPRESA_ID').value     := SomenteNumeros(sEmpresaCNPJ);
      QR.Parameters.ParamByName('NOTA_FISCAL').value    := StrToIntDef(pKONT_NUMERO_NOTA,0);
      QR.Parameters.ParamByName('STATUS_RETORNO').value := pSTATUS_RETORNO;

      if vXML <> '' then
        QR.Parameters.ParamByName('XML_ENVIADO').value  := vXML;

      QR.Parameters.ParamByName('MODELO').value         := pModelo;
      QR.Parameters.ParamByName('SERIE').value          := pSerie;

      if xHistId = 0 then
        QR.Parameters.ParamByName('DATA_PROCESSO').value  := Date;
      QR.ExecSQL;
      QR.Sql.Clear;
    finally
//    FreeandNil(QR);
    end;
  except
    on e: Exception do begin
      QR.SQL.SaveToFile(ExtractFilePath(Application.ExeName) + 'HistoricoNfe_log.txt');
      if QR <> nil then
        FreeandNil(QR);
    end;
  end;
end;

function TNFEInterfaceV3.ConsultaNotasPendendesManifesto(pIndicadorNFE, pIndicadorEmissor: string; var pUltimoNSU: string): TList;
var
  vMensagem : string;
  comando   : AnsiString;
  i         : Integer;
  ret       : TRetorno;
  retorno   : Boolean;
//  manifestar : Boolean;
  ListaManifestacao : TList;
  fNfMan   : TManifestacao;
  RESNFE   : Boolean;
  vLinha   : String;
  vStatus  : Integer;
begin
  try
    vMensagem := '';
    AbreSplashNFe('Aguarde, Consultando ' + fModelo + '...', IIfInt(fModelo = 'NFC-e',65,55));
    ret.status := TStringList.Create;

    RESNFE := False;

    if not UtilizaAcbrMonitor then begin
      if frmAcbrNFe = nil then
        frmAcbrNFe := TfrmAcbrNFe.Create(nil);

      KontNfe.Contingencia := '1';

      KontNfe.UF := sUF_WebService;

      KontNfe.ValidadeCertificado := frmAcbrNFe.ValidadeCertificado;

      retorno := frmAcbrNFe.ConsultaNfeDest(pIndicadorNFE,
                                            pIndicadorEmissor,
                                            pUltimoNSU,
                                            ret);
    end
    else begin
      comando := 'NFE.consultanfedest(' + sEmpresaCNPJ + ','
                                        + pIndicadorNFE + ','
                                        + pIndicadorEmissor + ','
                                        + pUltimoNSU + ')';

      try
        retorno := KontNfe.EnviarComando(comando,ret);
      except
        retorno := False;
      end;

    end;

    vMensagem := '';

    ListaManifestacao := TList.Create;

    if retorno then begin
        for I := 0 to ret.status.Count - 1 do begin
        vLinha := UpperCase(ret.status[i]);
        if Pos('[RESNFE',vLinha) > 0 then begin
          RESNFE := True;
          fNfMan := TManifestacao.Create;
          fNfMan.Tipo := 1;
        end;
        if Pos('[RESCAN',vLinha) > 0 then begin
          RESNFE := false;
        end;
        if Pos('[RESCCE',vLinha) > 0 then begin
          RESNFE := false;
        end;
        if RESNFE then begin
          if Pos('CHNFE=',vLinha) > 0 then begin
            fNfMan.chNFe    := Copy(vLinha,7,Length(vLinha)-6);
            fNfMan.NumeroNF := Copy(vLinha,42,8);
          end;
          if Pos('DEMI=',vLinha) > 0 then begin
            fNfMan.dEmi := Copy(vLinha,6,Length(vLinha)-5);
          end;
          if Pos('CNPJ=',vLinha) > 0 then begin
            fNfMan.CNPJ := Copy(vLinha,6,Length(vLinha)-5);
          end;
          if Pos('XNOME=',vLinha) > 0 then begin
            fNfMan.xNome := Copy(vLinha,7,Length(vLinha)-6);
          end;
          if Pos('VNF=',vLinha) > 0 then begin
            fNfMan.vrNF := Copy(vLinha,5,Length(vLinha)-4);
          end;
          if Pos('CSITCONF=',vLinha) > 0 then begin
            ListaManifestacao.Add(fNfMan);
            RESNFE := false;
          end;
        end;
        if Pos('CSTAT',vLinha) > 0 then begin
          vStatus := StrToIntDef(copy(vLinha,7,length(vLinha)-6),0);
          BuscaStatus(vStatus,vMensagem);
        end;
//        if Pos('XMOTIVO=',vLinha) > 0 then begin
//          vMensagem := copy(ret.status[i],9,length(ret.status[i])-8);
//        end;
        if Pos('ULTNSU=',vLinha) > 0 then begin
          pUltimoNSU := copy(vLinha,8,length(vLinha)-7);
        end;
        if Pos('ERRO',vLinha) > 0  then begin
          vMensagem := 'Mensagem Sefaz: ' + Copy(ret.status[i],7,Length(ret.status[i]))+#13;
          Break;
        end;
      end;
    end;

    if vMensagem <> '' then begin
      vMensagem := vMensagem + #13 +
                   'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                   'UF................: '  +#9+KontNfe.UF+#13+
                   'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                   'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                   'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                   'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                   'Versão.........: '     +#9+KontNfe.VersaoLayout;
      FechaSplashNFe;
      _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');
    end;
  finally
    FechaSplashNFe;
    result := ListaManifestacao;
  end;
end;

function TNFEInterfaceV3.EnviarManifestacao(pListaNFe : Tlist;
                                            pTpEvento : string;
                                            pJust     : string;
                                            pAN       : Boolean;
                                            var pListaMan : TList): Boolean;
var
  vMensagem    : string;
  comando      : AnsiString;
  i            : Integer;
  ret          : TRetorno;
  retorno      : Boolean;
  vOrgao       : integer;
//  vManifestado : boolean;
  fNfMan       : TManifestacao;
  vLinha       : string;
  vStatus      : Integer;
begin
  try
    retorno := False;
    vMensagem := '';
    AbreSplashNFe('Aguarde, enviando a Manifestação do Destinatário...');
    ret.status := TStringList.Create;

    if pAN then
      vOrgao := 91
    else
      vOrgao := StrToInt(copy(BuscaCodigoCidadeIBGE(IEmpresaCodcidade),1,2));

      if not UtilizaAcbrMonitor then begin

        if frmAcbrNFe = nil then
          frmAcbrNFe := TfrmAcbrNFe.Create(nil);

        KontNfe.Contingencia := '1';

        KontNfe.UF := sUF_WebService;

        KontNfe.ValidadeCertificado := frmAcbrNFe.ValidadeCertificado;

        retorno := frmAcbrNFe.EnviarManifestacao(pListaNFe,
                                                 pTpEvento,
                                                 pJust,
                                                 vOrgao,
                                                 ret);

      end
      else begin

        comando := 'NFE.enviarevento("[EVENTO]' + sLineBreak +
                                     'idLote=1' + sLineBreak;

//        for I := 0 to pListaNFe.Count -1 do begin

          if i > 0 then
            comando := comando + sLineBreak;

          comando := comando + '[EVENTO' + FormatFloat('000',i+1) + ']' + sLineBreak +
                               'idLote=1'    + sLineBreak +
                               'tpEvento='   + pTpEvento + sLineBreak +
                               'chNFe='      + TManifestacao(pListaNFe).chNFe + sLineBreak +
                               'cOrgao='     + IntToStr(vOrgao) + sLineBreak +
                               'CNPJ='       + sEmpresaCNPJ  + sLineBreak + //TManifestacao(pListaNFe.Items[i]).CNPJ + sLineBreak +
                               'dhEvento='   + FormatDateTime('DD/MM/YYYY',Now) + sLineBreak +
                               'nSeqEvento=1';

 //       end;
        comando := comando + '")';

        try
          retorno := KontNfe.EnviarComando(comando,ret);
        except
          retorno := False;
        end;
      end;
//    end;

    if retorno then begin
//      if ret.status.Count > 0  then begin
      fNfMan := TManifestacao.Create;
      for I := 0 to ret.status.Count - 1 do begin
        vLinha := UpperCase(ret.status[i]);
        if Pos('ERRO',vLinha) > 0  then begin
          vMensagem := Copy(ret.status[i],6,Length(ret.status[i])-5)+#13;
          result := False;
          Break;
        end;
        if Pos('CSTAT',vLinha) > 0 then begin
          vStatus := StrToIntDef(copy(vLinha,7,length(vLinha)-6),0);
          Result  := BuscaStatus(vStatus,vMensagem);
          fNfMan.Status := vStatus;
          if not Result then begin
            pListaMan.Add(fNfMan);
          end;
//          vManifestado := result;
        end;
        if result then begin
          if Pos('VERAPLIC=',vLinha) > 0 then begin
            KontNfe.VersaoLayout := Copy(ret.status[i],10,Length(ret.status[i])-9);
          end;
          if Pos('TPAMB=',vLinha) > 0 then begin
            KontNfe.TipoAmbiente := StrToIntDef(Copy(ret.status[i],7,Length(ret.status[i])-6),0);
          end;
          if Pos('CORGAO=',vLinha) > 0 then begin
            KontNfe.CodUF := Copy(ret.status[i],8,Length(ret.status[i])-7);
          end;
          if Pos('DHREGEVENTO=',vLinha) > 0 then begin
            fNfMan.dhRegEvento := Copy(ret.status[i],13,Length(ret.status[i])-12);
          end;
          if Pos('CHNFE=',vLinha) > 0 then begin
            fNfMan.chNFe := Copy(ret.status[i],7,Length(ret.status[i])-6);
          end;
          if Pos('NPROT=',vLinha) > 0 then begin
            fNfMan.nProt := Copy(ret.status[i],7,Length(ret.status[i])-6);
            pListaMan.Add(fNfMan);
          end;
        end;
      end;
    end
    else begin
      vMensagem := 'Não foi enviada a Manifestação do Destinatário! Tente novamente!'+#13;
    end;
  finally
    FechaSplashNFe;
  end;

  if (vMensagem = '') and (UtilizaAcbrMonitor) then begin
    vMensagem := 'Servidor Monitor Nfe não respondendo, tente novamente!';
  end
  else begin
    vMensagem := vMensagem + #13 +
                 'Empresa......: '       +#9+FormataCNPJCPF( KontNfe.Empresa )+#13+
                 'UF................: '  +#9+KontNfe.UF+#13+
                 'Codigo UF....: '       +#9+KontNfe.CodUF+#13+
                 'Ambiente.....: '       +#9+Decode(KontNfe.TipoAmbiente,[1,'Produção',2,'Homologação',IntToStr(KontNfe.TipoAmbiente)+'-Desconhecido'])+#13+
                 'Contigência.: '        +#9+Decode(KontNfe.Contingencia,[1,'Não',2,'Sim - Segurança',3,'Sim - SCAN','Desconhecido'])+#13+
                 'Validade......: '      +#9+KontNfe.ValidadeCertificado+' Dias '+#13+
                 'Versão.........: '     +#9+KontNfe.VersaoLayout;
  end;

  _ATENCAO(vMensagem,'0',1,TMConsistencia,0,'');
end;

//função para corrigir o &amp;
//e verificar se tem o protocolo de autorização
procedure TNFEInterfaceV3.TrataXML(var pMensagem  : AnsiString;
                                       pChaveNfe  : string;
                                       pDataHora  : string;
                                       pProtocolo : string;
                                       pdigVal    : string);
var
  vXML : AnsiString;
  vDataHora : String;
//  texto : Tstringlist;
begin
  vDataHora := Copy(pDataHora,7,4) + '-' +
               Copy(pDataHora,4,2) + '-' +
               Copy(pDataHora,1,2) + 'T' +
               Trim(Copy(pDataHora,11,9)) + '-03:00';

  if Pos('&AMP;AMP;',UpperCase(pMensagem)) > 0 then begin
    vXML := Copy(pMensagem,1,Pos('&AMP;AMP;',UpperCase(pMensagem))-1) +
            '&amp;' + Copy(pMensagem,Pos('&AMP;AMP;',UpperCase(pMensagem))+9,Length(pMensagem));
    pMensagem := vXML;
  end;
  if (NOT (Pos('protNFe',(pMensagem)) > 0)) or
     ((Pos('protNFeversao=',(pMensagem)) > 0)) then begin
    if not ((Pos('protNFeversao=',(pMensagem)) > 0)) then begin
      vXML := '<protNFe versao="3.10"><infProt Id="NFe00"><tpAmb>1</tpAmb><verAplic>GO3.0</verAplic><chNFe>';
      vXML := vXML + pChaveNfe + '</chNFe><dhRecbto>';
      vXML := vXML + vDataHora + '</dhRecbto><nProt>';
      vXML := vXML + pProtocolo + '</nProt><digVal>';
      vXML := vXML + pdigVal + '</digVal><cStat>100</cStat><xMotivo>Autorizado o uso da NF-e</xMotivo></infProt></protNFe>';
      if (Pos('/nfeProc',(pMensagem)) > 0) then begin
        vXML := Copy(pMensagem,1,Pos('/nfeProc',(pMensagem))) +
                vXML + '</nfeProc>';

      end else begin
        if not (Copy(pMensagem,1,5) = '<?xml') then begin
          vXML := '<?xml version="1.0" encoding="UTF-8"?><nfeProc versao="3.10" xmlns="http://www.portalfiscal.inf.br/nfe">' +
                  pMensagem + vXML + '</nfeProc>';
        end
        else
          vXML := Copy(pMensagem,1,POs('"UTF-8"?>',pMensagem)+8) +
                       '<nfeProc versao="3.10" xmlns="http://www.portalfiscal.inf.br/nfe">' +
                       Copy(pMensagem,POs('"UTF-8"?>',pMensagem)+9,Length(pMensagem)) +
                       vXML + '</nfeProc>';
      end;
    end
    else begin
      vXML := Copy(pMensagem,1,Pos('protNFeversao',(pMensagem))-1) + 'protNFe versao';
      vXml := vXML + Copy(pMensagem, Pos('protNFeversao',(pMensagem))+13, Length(pMensagem));
    end;
    pMensagem := vXML;
//    texto := TStringList.Create;
//    texto.Add(pMensagem);
//    texto.SaveToFile('c:\pXML.xml');

  end;
end;

//pTipo = 1 - Devolução de venda, 2 - Devolução de compra,
function TNFEInterfaceV3.RetornaFinalidadeConfCFOP(pCFOP: string; var pTipo: Integer): Integer;
VAR
  vCFOP : string;
begin
  vCFOP := SomenteNumeros(pCFOP);

  pTipo := 0;

  IF EM(vCFOP,['1201','1202','1203','1204','1208','1209','1410','1411','1503','1504','1505','1506',
        '1553','1660','1661','1662','1918','1919','2201','2202','2203','2204','2208','2209','2410',
        '2411','2503','2504','2505','2506','2553','2660','2661','2662','2918','2919','3201','3202',
        '3211','3503','3553']) THEN begin
    result := 4;
    pTipo := 1;
  end
  else if EM(vCFOP,['5201','5202','5208','5209','5210','5410','5411','5412','5413','5503','5553',
              '5555','5556','5660','5661','5662','5918','5919','5921','6201','6202','6208','6209',
              '6210','6410','6411','6412','6413','6503','6553','6555','6556','6660','6661','6662',
              '6918','6919','6921','7201','7202','7210','7211','7553','7556']) THEN begin
    result := 4;
    pTipo := 2;
  end
  else begin
    result := 1;
  end;
end;

function TNFEInterfaceV3.BuscaStatus(var pStatus: Integer; var pMensagem: string; pNFCE: Boolean = false): Boolean;
begin   //maximo 470 constantes
  case pStatus of
    //RESULTADO DO PROCESSAMENTO DA SOLICITAÇÃO
      0: begin Result := False; pMensagem := 'Rejeição: Erro não catalogado '; end;
    100: begin result := True; pMensagem :=  'Autorizado o uso da NF-e'; end;
    101: begin result := True; pMensagem :=  'Cancelamento de NF-e homologado'; end;
    102: begin result := True; pMensagem :=  'Inutilização de número homologado'; end;
    103: begin result := True; pMensagem :=  'Lote recebido com sucesso'; end;
    104: begin result := True; pMensagem :=  'Lote processado'; end;
    105: begin result := True; pMensagem :=  'Lote em processamento'; end;
    106: begin result := True; pMensagem :=  'Lote não localizado'; end;
    107: begin result := True; pMensagem :=  'Serviço em Operação'; end;
    108: begin result := false; pMensagem :=  'Serviço Paralisado Momentaneamente (curto prazo)'; end;
    109: begin result := false; pMensagem :=  'Serviço Paralisado sem Previsão'; end;
    110: begin result := false; pMensagem :=  'Uso Denegado'; end;
    111: begin result := True; pMensagem :=  'Consulta cadastro com uma ocorrência'; end;
    112: begin result := True; pMensagem :=  'Consulta cadastro com mais de uma ocorrência'; end;
    124: begin result := True; pMensagem :=  'EPEC Autorizado';end;
    128: begin result := True; pMensagem :=  'Lote de Evento Processado'; end;
    135: begin result := True; pMensagem :=  'Evento registrado e vinculado a ' + IIf(pNFCE,'NFC-e','NF-e'); end;
    136: begin result := True; pMensagem :=  'Evento registrado, mas não vinculado a ' + IIf(pNFCE,'NFC-e','NF-e'); end;
    137: begin result := false; pMensagem := 'Nenhum documento localizado para o destinatário'; end;
    138: begin result := True; pMensagem :=  'Documento localizado para o destinatário'; end;
    139: begin result := True; pMensagem :=  'Pedido de Download processado'; end;
    140: begin result := True; pMensagem :=  'Download disponibilizado'; end;
    142: begin result := false; pMensagem := 'Ambiente de Contingência EPEC bloqueado para o Emitente';end;
    150: begin result := true; pMensagem :=  'Autorizado o uso da NF-e, autorização fora de prazo';end;
    151: begin result := True; pMensagem :=  'Cancelamento de NF-e homologado fora de prazo';end;

    //MOTIVOS DE NÃO ATENDIMENTO DA SOLICITAÇÃO
    201: begin result := false; pMensagem := 'Rejeição: O numero máximo de numeração de ' + IIf(pNFCE,'NFC-e','NF-e') + ' a inutilizar ultrapassou o limite'; end;
    202: begin result := false; pMensagem :=  'Rejeição: Falha no reconhecimento da autoria ou integridade do arquivo digital'; end;
    203: begin result := false; pMensagem :=  'Rejeição: Emissor não habilitado para emissão da ' + IIf(pNFCE,'NFC-e','NF-e'); end;
    204: begin result := false; pMensagem :=  'Duplicidade de ' + IIf(pNFCE,'NFC-e','NF-e') +  ' [nRec:999999999999999]'; end;
    205: begin result := false; pMensagem :=  'NF-e está denegada na base de dados da SEFAZ [nRec:999999999999999]'; end;
    206: begin result := false; pMensagem :=  'Rejeição: ' + IIf(pNFCE,'NFC-e','NF-e') + ' já está inutilizada na Base de dados da SEFAZ'; end;
    207: begin result := false; pMensagem :=  'Rejeição: CNPJ do emitente inválido'; end;
    208: begin result := false; pMensagem :=  'Rejeição: CNPJ do destinatário inválido'; end;
    209: begin result := false; pMensagem :=  'Rejeição: IE do emitente inválida'; end;
    210: begin result := false; pMensagem :=  'Rejeição: IE do destinatário inválida'; end;
    211: begin result := false; pMensagem :=  'Rejeição: IE do substituto inválida'; end;
    212: begin result := false; pMensagem :=  'Rejeição: Data de emissão NF-e posterior a data de recebimento'; end;
    213: begin result := false; pMensagem :=  'Rejeição: CNPJ-Base do Emitente difere do CNPJ-Base do Certificado Digital'; end;
    214: begin result := false; pMensagem :=  'Rejeição: Tamanho da mensagem excedeu o limite estabelecido'; end;
    215: begin result := false; pMensagem :=  'Rejeição: Falha no schema XML'; end;
    216: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso difere da cadastrada'; end;
    217: begin result := false; pMensagem :=  'Rejeição: ' + IIf(pNFCE,'NFC-e','NF-e') + ' não consta na base de dados da SEFAZ'; end;
    218: begin result := false; pMensagem :=  'NF-e já está cancelada na base de dados da SEFAZ [nRec:999999999999999]'; end;
    219: begin result := false; pMensagem :=  'Rejeição: Circulação da NF-e verificada'; end;
    220: begin result := false; pMensagem :=  'Rejeição: Prazo de Cancelamento superior ao previsto na Legislação'; end;
    221: begin result := false; pMensagem :=  'Rejeição: Confirmado o recebimento da ' + IIf(pNFCE,'NFC-e','NF-e') + ' pelo destinatário'; end;
    222: begin result := false; pMensagem :=  'Rejeição: Protocolo de Autorização de Uso difere do cadastrado'; end;
    223: begin result := false; pMensagem :=  'Rejeição: CNPJ do transmissor do lote difere do CNPJ do transmissor da consulta'; end;
    224: begin result := false; pMensagem :=  'Rejeição: A faixa inicial é maior que a faixa final'; end;
    225: begin result := false; pMensagem :=  'Rejeição: Falha no Schema XML do lote de ' + IIf(pNFCE,'NFC-e','NF-e'); end;
    226: begin result := false; pMensagem :=  'Rejeição: Código da UF do Emitente diverge da UF autorizadora'; end;
    227: begin result := false; pMensagem :=  'Rejeição: Erro na Chave de Acesso - Campo Id – falta a literal ' + IIf(pNFCE,'NFC-e','NF-e'); end;
    228: begin result := false; pMensagem :=  'Rejeição: Data de Emissão muito atrasada'; end;
    229: begin result := false; pMensagem :=  'Rejeição: IE do emitente não informada'; end;
    230: begin result := false; pMensagem :=  'Rejeição: IE do emitente não cadastrada'; end;
    231: begin result := false; pMensagem :=  'Rejeição: IE do emitente não vinculada ao CNPJ'; end;
    232: begin result := false; pMensagem :=  'Rejeição: IE do destinatário não informada'; end;
    233: begin result := false; pMensagem :=  'Rejeição: IE do destinatário não cadastrada'; end;
    234: begin result := false; pMensagem :=  'Rejeição: IE do destinatário não vinculada ao CNPJ'; end;
    235: begin result := false; pMensagem :=  'Rejeição: Inscrição SUFRAMA inválida'; end;
    236: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso com dígito verificador inválido'; end;
    237: begin result := false; pMensagem :=  'Rejeição: CPF do destinatário inválido'; end;
    238: begin result := false; pMensagem :=  'Rejeição: Cabeçalho - Versão do arquivo XML superior a Versão vigente'; end;
    239: begin result := false; pMensagem :=  'Rejeição: Cabeçalho - Versão do arquivo XML não suportada'; end;
    240: begin result := false; pMensagem :=  'Rejeição: Cancelamento/Inutilização - Irregularidade Fiscal do Emitente'; end;
    241: begin result := false; pMensagem :=  'Rejeição: Um número da faixa já foi utilizado'; end;
    242: begin result := false; pMensagem :=  'Rejeição: Cabeçalho - Falha no Schema XML'; end;
    243: begin result := false; pMensagem :=  'Rejeição: XML Mal Formado'; end;
    244: begin result := false; pMensagem :=  'Rejeição: CNPJ do Certificado Digital difere do CNPJ da Matriz e do CNPJ do Emitente'; end;
    245: begin result := false; pMensagem :=  'Rejeição: CNPJ Emitente não cadastrado'; end;
    246: begin result := false; pMensagem :=  'Rejeição: CNPJ Destinatário não cadastrado'; end;
    247: begin result := false; pMensagem :=  'Rejeição: Sigla da UF do Emitente diverge da UF autorizadora'; end;
    248: begin result := false; pMensagem :=  'Rejeição: UF do Recibo diverge da UF autorizadora'; end;
    249: begin result := false; pMensagem :=  'Rejeição: UF da Chave de Acesso diverge da UF autorizadora'; end;
    250: begin result := false; pMensagem :=  'Rejeição: UF diverge da UF autorizadora'; end;
    251: begin result := false; pMensagem :=  'Rejeição: UF/Município destinatário não pertence a SUFRAMA'; end;
    252: begin result := false; pMensagem :=  'Rejeição: Ambiente informado diverge do Ambiente de recebimento'; end;
    253: begin result := false; pMensagem :=  'Rejeição: Digito Verificador da chave de acesso composta inválida'; end;
    254: begin result := false; pMensagem :=  'Rejeição: NF-e complementar não possui NF referenciada'; end;
    255: begin result := false; pMensagem :=  'Rejeição: NF-e complementar possui mais de uma NF referenciada'; end;
    256: begin result := false; pMensagem :=  'Rejeição: Uma NF-e da faixa já está inutilizada na Base de dados da SEFAZ'; end;
    257: begin result := false; pMensagem :=  'Rejeição: Solicitante não habilitado para emissão da NF-e'; end;
    258: begin result := false; pMensagem :=  'Rejeição: CNPJ da consulta inválido'; end;
    259: begin result := false; pMensagem :=  'Rejeição: CNPJ da consulta não cadastrado como contribuinte na UF'; end;
    260: begin result := false; pMensagem :=  'Rejeição: IE da consulta inválida'; end;
    261: begin result := false; pMensagem :=  'Rejeição: IE da consulta não cadastrada como contribuinte na UF'; end;
    262: begin result := false; pMensagem :=  'Rejeição: UF não fornece consulta por CPF'; end;
    263: begin result := false; pMensagem :=  'Rejeição: CPF da consulta inválido'; end;
    264: begin result := false; pMensagem :=  'Rejeição: CPF da consulta não cadastrado como contribuinte na UF'; end;
    265: begin result := false; pMensagem :=  'Rejeição: Sigla da UF da consulta difere da UF do Web Service'; end;
    266: begin result := false; pMensagem :=  'Rejeição: Série utilizada não permitida no Web Service'; end;
    267: begin result := false; pMensagem :=  'Rejeição: NF Complementar referencia uma NF-e inexistente'; end;
    268: begin result := false; pMensagem :=  'Rejeição: NF Complementar referencia uma outra NF-e Complementar'; end;
    269: begin result := false; pMensagem :=  'Rejeição: CNPJ Emitente da NF Complementar difere do CNPJ da NF Referenciada'; end;
    270: begin result := false; pMensagem :=  'Rejeição: Código Município do Fato Gerador: dígito inválido'; end;
    271: begin result := false; pMensagem :=  'Rejeição: Código Município do Fato Gerador: difere da UF do emitente'; end;
    272: begin result := false; pMensagem :=  'Rejeição: Código Município do Emitente: dígito inválido'; end;
    273: begin result := false; pMensagem :=  'Rejeição: Código Município do Emitente: difere da UF do emitente'; end;
    274: begin result := false; pMensagem :=  'Rejeição: Código Município do Destinatário: dígito inválido'; end;
    275: begin result := false; pMensagem :=  'Rejeição: Código Município do Destinatário: difere da UF do Destinatário'; end;
    276: begin result := false; pMensagem :=  'Rejeição: Código Município do Local de Retirada: dígito inválido'; end;
    277: begin result := false; pMensagem :=  'Rejeição: Código Município do Local de Retirada: difere da UF do Local de Retirada'; end;
    278: begin result := false; pMensagem :=  'Rejeição: Código Município do Local de Entrega: dígito inválido'; end;
    279: begin result := false; pMensagem :=  'Rejeição: Código Município do Local de Entrega: difere da UF do Local de Entrega'; end;
    280: begin result := false; pMensagem :=  'Rejeição: Certificado Transmissor inválido'; end;
    281: begin result := false; pMensagem :=  'Rejeição: Certificado Transmissor Data Validade'; end;
    282: begin result := false; pMensagem :=  'Rejeição: Certificado Transmissor sem CNPJ'; end;
    283: begin result := false; pMensagem :=  'Rejeição: Certificado Transmissor - erro Cadeia de Certificação'; end;
    284: begin result := false; pMensagem :=  'Rejeição: Certificado Transmissor revogado'; end;
    285: begin result := false; pMensagem :=  'Rejeição: Certificado Transmissor difere ICP-Brasil'; end;
    286: begin result := false; pMensagem :=  'Rejeição: Certificado Transmissor erro no acesso a LCR'; end;
    287: begin result := false; pMensagem :=  'Rejeição: Código Município do FG - ISSQN: dígito inválido'; end;
    288: begin result := false; pMensagem :=  'Rejeição: Código Município do FG - Transporte: dígito inválido'; end;
    289: begin result := false; pMensagem :=  'Rejeição: Código da UF informada diverge da UF solicitada'; end;
    290: begin result := false; pMensagem :=  'Rejeição: Certificado Assinatura inválido'; end;
    291: begin result := false; pMensagem :=  'Rejeição: Certificado Assinatura Data Validade'; end;
    292: begin result := false; pMensagem :=  'Rejeição: Certificado Assinatura sem CNPJ'; end;
    293: begin result := false; pMensagem :=  'Rejeição: Certificado Assinatura - erro Cadeia de Certificação'; end;
    294: begin result := false; pMensagem :=  'Rejeição: Certificado Assinatura revogado'; end;
    295: begin result := false; pMensagem :=  'Rejeição: Certificado Assinatura difere ICP-Brasil'; end;
    296: begin result := false; pMensagem :=  'Rejeição: Certificado Assinatura erro no acesso a LCR'; end;
    297: begin result := false; pMensagem :=  'Rejeição: Assinatura difere do calculado'; end;
    298: begin result := false; pMensagem :=  'Rejeição: Assinatura difere do padrão do Sistema'; end;
    299: begin result := false; pMensagem :=  'Rejeição: XML da área de cabeçalho com codificação diferente de UTF-8'; end;
    301: begin result := false; pMensagem :=  'Uso Denegado: Irregularidade fiscal do emitente'; end;
    302: begin result := false; pMensagem :=  'Uso Denegado: Irregularidade fiscal do destinatário'; end;
    303: begin result := false; pMensagem :=  'Uso Denegado: Destinatário não habilitado a operar na UF'; end;
    304: begin result := false; pMensagem :=  'Rejeição: Pedido de Cancelamento para ' + IIf(pNFCE,'NFC-e','NF-e') + ' com evento da Suframa'; end;
    315: begin result := false; pMensagem :=  'Rejeição: Data de Emissão anterior ao início da autorização de Nota Fiscal na UF'; end;
    316: begin result := false; pMensagem :=  'Rejeição: Nota Fiscal referenciada com a mesma Chave de Acesso da Nota Fiscal atual'; end;
    317: begin result := false; pMensagem :=  'Rejeição: NF modelo 1 referenciada com data de emissão inválida'; end;
    318: begin result := false; pMensagem :=  'Rejeição: Contranota de Produtor sem Nota Fiscal referenciada'; end;
    319: begin result := false; pMensagem :=  'Rejeição: Contranota de Produtor não pode referenciar somente Nota Fiscal de entrada'; end;
    320: begin result := false; pMensagem :=  'Rejeição: Contranota de Produtor referencia somente NF de outro emitente'; end;
    321: begin result := false; pMensagem :=  'Rejeição: NF-e de devolução de mercadoria não possui documento fiscal referenciado'; end;
    322: begin result := false; pMensagem :=  'Rejeição: NF de produtor referenciada com data de emissão inválida'; end;
    323: begin result := false; pMensagem :=  'Rejeição: CNPJ autorizado para download inválido'; end;
    324: begin result := false; pMensagem :=  'Rejeição: CNPJ do destinatário já autorizado para download'; end;
    325: begin result := false; pMensagem :=  'Rejeição: CPF autorizado para download inválido'; end;
    326: begin result := false; pMensagem :=  'Rejeição: CPF do destinatário já autorizado para download'; end;
    327: begin result := false; pMensagem :=  'Rejeição: CFOP inválido para Nota Fiscal com finalidade de devolução de mercadoria [nItem:nnn]'; end;
    328: begin result := false; pMensagem :=  'Rejeição: CFOP de devolução de mercadoria para NF-e que não tem finalidade de devolução de mercadoria'; end;
    329: begin result := false; pMensagem :=  'Rejeição: Número da DI /DSI inválido'; end;
    330: begin result := false; pMensagem :=  'Rejeição: Informar o Valor da AFRMM na importação por via marítima'; end;
    331: begin result := false; pMensagem :=  'Rejeição: Informar o CNPJ do adquirente ou do encomendante nesta forma de importação'; end;
    332: begin result := false; pMensagem :=  'Rejeição: CNPJ do adquirente ou do encomendante da importação inválido'; end;
    333: begin result := false; pMensagem :=  'Rejeição: Informar a UF do adquirente ou do encomendante nesta forma de importação'; end;
    334: begin result := false; pMensagem :=  'Rejeição: Número do processo de drawback não informado na importação'; end;
    335: begin result := false; pMensagem :=  'Rejeição: Número do processo de drawback na importação inválido'; end;
    336: begin result := false; pMensagem :=  'Rejeição: Informado o grupo de exportação no item para CFOP que não é de exportação'; end;
    337: begin result := false; pMensagem :=  'Rejeição: NFC-e para emitente pessoa física'; end;
    338: begin result := false; pMensagem :=  'Rejeição: Número do processo de drawback não informado na exportação'; end;
    339: begin result := false; pMensagem :=  'Rejeição: Número do processo de drawback na exportação inválido'; end;
    340: begin result := false; pMensagem :=  'Rejeição: Não informado o grupo de exportação indireta no item'; end;
    341: begin result := false; pMensagem :=  'Rejeição: Número do registro de exportação inválido'; end;
    342: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso informada na Exportação Indireta com DV inválido'; end;
    343: begin result := false; pMensagem :=  'Rejeição: Modelo da NF-e informada na Exportação Indireta diferente de 55'; end;
    344: begin result := false; pMensagem :=  'Rejeição: Duplicidade de NF-e informada na Exportação Indireta (Chave de Acesso informada mais de uma vez)'; end;
    345: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso informada na Exportação Indireta não consta como NF-e referenciada'; end;
    346: begin result := false; pMensagem :=  'Rejeição: Somatório das quantidades informadas na Exportação Indireta não corresponde a quantidade total do item'; end;
    347: begin result := false; pMensagem :=  'Rejeição: Descrição do combustível diverge da descrição adotada pela ANP'; end;
    348: begin result := false; pMensagem :=  'Rejeição: NFC-e com grupo RECOPI'; end;
    349: begin result := false; pMensagem :=  'Rejeição: Número RECOPI não informado'; end;
    350: begin result := false; pMensagem :=  'Rejeição: Número RECOPI inválido'; end;
    351: begin result := false; pMensagem :=  'Rejeição: Valor do ICMS da Operação no CST=51 difere do produto BC e Alíquota'; end;
    352: begin result := false; pMensagem :=  'Rejeição: Valor do ICMS Diferido no CST=51 difere do produto Valor ICMS Operação e percentual diferimento'; end;
    353: begin result := false; pMensagem :=  'Rejeição: Valor do ICMS no CST=51 não corresponde a diferença do ICMS operação e ICMS diferido'; end;
    354: begin result := false; pMensagem :=  'Rejeição: Informado grupo de devolução de tributos para NF-e que não tem finalidade de devolução de mercadoria'; end;
    355: begin result := false; pMensagem :=  'Rejeição: Informar o local de saída do Pais no caso da exportação'; end;
    356: begin result := false; pMensagem :=  'Rejeição: Informar o local de saída do Pais somente no caso da exportação'; end;
    357: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso do grupo de Exportação Indireta inexistente [nRef: xxx]'; end;
    358: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso do grupo de Exportação Indireta cancelada ou denegada [nRef: xxx]'; end;
    359: begin result := false; pMensagem :=  'Rejeição: NF-e de venda a Órgão Público sem informar a Nota de Empenho'; end;
    360: begin result := false; pMensagem :=  'Rejeição: NF-e com Nota de Empenho inválida para a UF.'; end;
    361: begin result := false; pMensagem :=  'Rejeição: NF-e com Nota de Empenho inexistente na UF.'; end;
    362: begin result := false; pMensagem :=  'Rejeição: Venda de combustível sem informação do Transportador'; end;
    364: begin result := false; pMensagem :=  'Rejeição: Total do valor da dedução do ISS difere do somatório dos itens'; end;
    365: begin result := false; pMensagem :=  'Rejeição: Total de outras retenções difere do somatório dos itens'; end;
    366: begin result := false; pMensagem :=  'Rejeição: Total do desconto incondicionado ISS difere do somatório dos itens'; end;
    367: begin result := false; pMensagem :=  'Rejeição: Total do desconto condicionado ISS difere do somatório dos itens'; end;
    368: begin result := false; pMensagem :=  'Rejeição: Total de ISS retido difere do somatório dos itens'; end;
    369: begin result := false; pMensagem :=  'Rejeição: Não informado o grupo avulsa na emissão pelo Fisco'; end;
    370: begin result := false; pMensagem :=  'Rejeição: Nota Fiscal Avulsa com tipo de emissão inválido'; end;
    372: begin result := false; pMensagem :=  'Rejeição: Destinatário com identificação de estrangeiro com caracteres inválidos'; end;
    373: begin result := false; pMensagem :=  'Rejeição: Descrição do primeiro item diferente de NOTA FISCAL EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL'; end;
    374: begin result := false; pMensagem :=  'Rejeição: CFOP incompatível com o grupo de tributação [nItem:nnn]'; end;
    375: begin result := false; pMensagem :=  'Rejeição: NF-e com CFOP 5929 (Lançamento relativo a Cupom Fiscal) referencia uma NFC-e [nItem:nnn]'; end;
    376: begin result := false; pMensagem :=  'Rejeição: Data do Desembaraço Aduaneiro inválida [nItem:nnn]'; end;
    378: begin result := false; pMensagem :=  'Rejeição: Grupo de Combustível sem a informação de Encerrante [nItem:nnn]'; end;
    379: begin result := false; pMensagem :=  'Rejeição: Grupo de Encerrante na NF-e (modelo 55) para CFOP diferente de venda de combustível para consumidor final [nItem:nnn]'; end;
    380: begin result := false; pMensagem :=  'Rejeição: Valor do Encerrante final não é superior ao Encerrante inicial [nItem:nnn]'; end;
    381: begin result := false; pMensagem :=  'Rejeição:Grupo de tributação ICMS90, informando dados do ICMS-ST [nItem:nnn]'; end;
    382: begin result := false; pMensagem :=  'Rejeição:CFOP não permitido para o CST informado [nItem:nnn]'; end;
    383: begin result := false; pMensagem :=  'Rejeição: Item com CSOSN indevido [nItem:nnn]'; end;
    384: begin result := false; pMensagem :=  'Rejeição: CSOSN não permitido para a UF [nItem:nnn]'; end;
    385: begin result := false; pMensagem :=  'Rejeição:Grupo de tributação ICMS900, informando dados do ICMS-ST [nItem:nnn]'; end;
    386: begin result := false; pMensagem :=  'Rejeição: CFOP não permitido para o CSOSN informado [nItem:nnn]'; end;
    387: begin result := false; pMensagem :=  'Rejeição: Código de Enquadramento Legal do IPI inválido [nItem:nnn]'; end;
    388: begin result := false; pMensagem :=  'Rejeição: Código de Situação Tributária do IPI incompatível com o Código de Enquadramento Legal do IPI [nItem:nnn]'; end;
    389: begin result := false; pMensagem :=  'Rejeição: Código Município ISSQN inexistente [nItem:nnn]'; end;
    390: begin result := false; pMensagem :=  'Rejeição: Nota Fiscal com grupo de devolução de tributos [nItem:nnn]'; end;
    391: begin result := false; pMensagem :=  'Rejeição: Não informados os dados do cartão de crédito / débito nas Formas de Pagamento da Nota Fiscal'; end;
    392: begin result := false; pMensagem :=  'Rejeição: Não informados os dados da operação de pagamento por cartão de crédito / débito'; end;
    393: begin result := false; pMensagem :=  'Rejeição: NF-e com o grupo de Informações Suplementares'; end;
    394: begin result := false; pMensagem :=  'Rejeição: Nota Fiscal sem a informação do QR-Code'; end;
    395: begin result := false; pMensagem :=  'Rejeição: Endereço do site da UF da Consulta via QRCode diverge do previsto'; end;
    396: begin result := false; pMensagem :=  'Rejeição: Parâmetro do QR-Code inexistente (chAcesso)'; end;
    397: begin result := false; pMensagem :=  'Rejeição: Parâmetro do QR-Code divergente da Nota Fiscal (chAcesso)'; end;
    398: begin result := false; pMensagem :=  'Rejeição: Parâmetro nVersao do QR-Code difere do previsto'; end;
    399: begin result := false; pMensagem :=  'Rejeição: Parâmetro de Identificação do destinatário no QR-Code para Nota Fiscal sem identificação do destinatário'; end;
    400: begin result := false; pMensagem :=  'Rejeição: Parâmetro do QR-Code não está no formato hexadecimal (dhEmi)'; end;
    401: begin result := false; pMensagem :=  'Rejeição: CPF do remetente inválido'; end;
    402: begin result := false; pMensagem :=  'Rejeição: XML da área de dados com codificação diferente de UTF-8'; end;
    403: begin result := false; pMensagem :=  'Rejeição: O grupo de informações da NF-e avulsa é de uso exclusivo do Fisco'; end;
    404: begin result := false; pMensagem :=  'Rejeição: Uso de prefixo de namespace não permitido'; end;
    405: begin result := false; pMensagem :=  'Rejeição: Código do país do emitente: dígito inválido'; end;
    406: begin result := false; pMensagem :=  'Rejeição: Código do país do destinatário: dígito inválido'; end;
    407: begin result := false; pMensagem :=  'Rejeição: O CPF só pode ser informado no campo emitente para a NF-e avulsa'; end;
    409: begin result := false; pMensagem :=  'Rejeição: Campo cUF inexistente no elemento nfeCabecMsg do SOAP Header'; end;
    410: begin result := false; pMensagem :=  'Rejeição: UF informada no campo cUF não é atendida pelo Web Service'; end;
    411: begin result := false; pMensagem :=  'Rejeição: Campo versaoDados inexistente no elemento nfeCabecMsg do SOAP Header'; end;
    417: begin result := false; pMensagem :=  'Rejeição: Total do ICMS superior ao valor limite estabelecido'; end;
    418: begin result := false; pMensagem :=  'Rejeição: Total do ICMS ST superior ao valor limite estabelecido'; end;
    420: begin result := false; pMensagem :=  'Rejeição: Cancelamento para ' + IIf(pNFCE,'NFC-e','NF-e') + ' já cancelada'; end;
    450: begin result := false; pMensagem :=  'Rejeição: Modelo da NF-e diferente de 55'; end;
    451: begin result := false; pMensagem :=  'Rejeição: Processo de emissão informado inválido'; end;
    452: begin result := false; pMensagem :=  'Rejeição: Tipo Autorizador do Recibo diverge do Órgão Autorizador'; end;
    453: begin result := false; pMensagem :=  'Rejeição: Ano de inutilização não pode ser superior ao Ano atual'; end;
    454: begin result := false; pMensagem :=  'Rejeição: Ano de inutilização não pode ser inferior a 2006'; end;
    455: begin result := false; pMensagem :=  'Rejeição: Órgão Autor do evento diferente da UF da Chave de Acesso'; end;
    461: begin result := false; pMensagem :=  'Rejeição: Informado percentual de Gás Natural na mistura para produto diferente de GLP'; end;
    462: begin result := false; pMensagem :=  'Rejeição:Código Identificador do CSC no QR-Code não cadastrado na SEFAZ'; end;
    463: begin result := false; pMensagem :=  'Rejeição:Código Identificador do CSC no QR-Code foi revogado pela empresa'; end;
    464: begin result := false; pMensagem :=  'Rejeição: Código de Hash no QR-Code difere do calculado'; end;
    465: begin result := false; pMensagem :=  'Rejeição: Número de Controle da FCI inexistente'; end;
    466: begin result := false; pMensagem :=  'Rejeição: Evento com Tipo de Autor incompatível'; end;
    467: begin result := false; pMensagem :=  'Rejeição: Dados da NF-e divergentes do EPEC'; end;
    468: begin result := false; pMensagem :=  'Rejeição: NF-e com Tipo Emissão = 4, sem EPEC correspondente'; end;
    471: begin result := false; pMensagem :=  'Rejeição: Informado NCM=00 indevidamente'; end;
    476: begin result := false; pMensagem :=  'Rejeição: Código da UF diverge da UF da primeira NF-e do Lote'; end;
    477: begin result := false; pMensagem :=  'Rejeição: Código do órgão diverge do órgão do primeiro evento do Lote'; end;
    478: begin result := false; pMensagem :=  'Rejeição: Local da entrega não informado para faturamento direto de veículos novos'; end;
    479: begin result := false; pMensagem :=  'Rejeição: Data de Emissão anterior a data de credenciamento ou anterior a Data de Abertura do estabelecimento'; end;
    480: begin result := false; pMensagem :=  'Rejeição: Código Município do Emitente diverge do cadastrado na UF'; end;
    481: begin result := false; pMensagem :=  'Rejeição: Código Regime Tributário do emitente diverge do cadastro na SEFAZ'; end;
    482: begin result := false; pMensagem :=  'Rejeição: Código do Município do Destinatário diverge do cadastrado na UF'; end;
    483: begin result := false; pMensagem :=  'Rejeição: Valor do desconto maior que valor do produto [nItem:nnn]'; end;
    484: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso com tipo de emissão diferente de 4 (posição 35 da Chave de Acesso)'; end;
    485: begin result := false; pMensagem :=  'Rejeição: Duplicidade de numeração do EPEC (Modelo, CNPJ, Série e Número)'; end;
    486: begin result := false; pMensagem :=  'Rejeição: Não informado o Grupo de Autorização para UF que exige a identificação'; end;
    487: begin result := false; pMensagem :=  'Rejeição: Escritório de Contabilidade não cadastrado na SEFAZ'; end;
    488: begin result := false; pMensagem :=  'Rejeição: Vendas do Emitente incompatíveis com o Porte da Empresa'; end;
    489: begin result := false; pMensagem :=  'Rejeição: CNPJ informado inválido (DV ou zeros)'; end;
    490: begin result := false; pMensagem :=  'Rejeição: CPF informado inválido (DV ou zeros)'; end;
    491: begin result := false; pMensagem :=  'Rejeição: O tpEvento informado inválido'; end;
    492: begin result := false; pMensagem :=  'Rejeição: O verEvento informado inválido'; end;
    493: begin result := false; pMensagem :=  'Rejeição: Evento não atende o Schema XML específico'; end;
    494: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso inexistente'; end;
    496: begin result := false; pMensagem :=  'Rejeição: Não informado o tipo de integração no pagamento com cartão de crédito / débito'; end;
    501: begin result := false; pMensagem :=  'Rejeição: Pedido de Cancelamento intempestivo (NF-e autorizada a mais de 7 dias)'; end;
    502: begin result := false; pMensagem :=  'Rejeição: Erro na Chave de Acesso - Campo Id não corresponde à concatenação dos campos correspondentes'; end;
    503: begin result := false; pMensagem :=  'Rejeição: Série utilizada fora da faixa permitida no SCAN (900-999)'; end;
    504: begin result := false; pMensagem :=  'Rejeição: Data de Entrada/Saída posterior ao permitido'; end;
    505: begin result := false; pMensagem :=  'Rejeição: Data de Entrada/Saída anterior ao permitido'; end;
    506: begin result := false; pMensagem :=  'Rejeição: Data de Saída menor que a Data de Emissão'; end;
    507: begin result := false; pMensagem :=  'Rejeição: O CNPJ do destinatário/remetente não deve ser informado em operação com o exterior'; end;
    508: begin result := false; pMensagem :=  'Rejeição: O CNPJ com conteúdo nulo só é válido em operação com exterior'; end;
    509: begin result := false; pMensagem :=  'Rejeição: Informado código de município diferente de “9999999” para operação com o exterior'; end;
    510: begin result := false; pMensagem :=  'Rejeição: Operação com Exterior e Código País destinatário é 1058 (Brasil) ou não informado'; end;
    511: begin result := false; pMensagem :=  'Rejeição: Não é de Operação com Exterior e Código País destinatário difere de 1058 (Brasil)'; end;
    512: begin result := false; pMensagem :=  'Rejeição: CNPJ do Local de Retirada inválido'; end;
    513: begin result := false; pMensagem :=  'Rejeição: Código Município do Local de Retirada deve ser 9999999 para UF retirada = EX'; end;
    514: begin result := false; pMensagem :=  'Rejeição: CNPJ do Local de Entrega inválido'; end;
    515: begin result := false; pMensagem :=  'Rejeição: Código Município do Local de Entrega deve ser 9999999 para UF entrega = EX'; end;
    516: begin result := false; pMensagem :=  'Rejeição: Falha no schema XML – inexiste a tag raiz esperada para a mensagem'; end;
    517: begin result := false; pMensagem :=  'Rejeição: Falha no schema XML – inexiste atributo versao na tag raiz da mensagem'; end;
    518: begin result := false; pMensagem :=  'Rejeição: CFOP de entrada para NF-e de saída'; end;
    519: begin result := false; pMensagem :=  'Rejeição: CFOP de saída para NF-e de entrada'; end;
    520: begin result := false; pMensagem :=  'Rejeição: CFOP de Operação com Exterior e UF destinatário difere de EX'; end;
    521: begin result := false; pMensagem :=  'Rejeição:  CFOP  de  Operação  Estadual  e  UF  do  emitente  difere  da  UF  do  destinatário para destinatário contribuinte do ICMS.'; end;
    522: begin result := false; pMensagem :=  'Rejeição: CFOP de Operação Estadual e UF emitente difere da UF remetente para remetente contribuinte do ICMS.'; end;
    523: begin result := false; pMensagem :=  'Rejeição: CFOP não é de Operação Estadual e UF emitente igual a UFdestinatário.'; end;
    524: begin result := false; pMensagem :=  'Rejeição: CFOP de Operação com Exterior e não informado NCM'; end;
    525: begin result := false; pMensagem :=  'Rejeição: CFOP de Importação e não informado dados da DI'; end;
    526: begin result := false; pMensagem :=  'Rejeição: CFOP de Exportação e não informado Local de Embarque'; end;
    527: begin result := false; pMensagem :=  'Rejeição: Operação de Exportação com informação de ICMS incompatível'; end;
    528: begin result := false; pMensagem :=  'Rejeição: Valor do ICMS difere do produto BC e Alíquota'; end;
    529: begin result := false; pMensagem :=  'Rejeição: NCM de informação obrigatória para produto tributado pelo IPI'; end;
    530: begin result := false; pMensagem :=  'Rejeição: Operação com tributação de ISSQN sem informar a Inscrição Municipal'; end;
    531: begin result := false; pMensagem :=  'Rejeição: Total da BC ICMS difere do somatório dos itens'; end;
    532: begin result := false; pMensagem :=  'Rejeição: Total do ICMS difere do somatório dos itens'; end;
    533: begin result := false; pMensagem :=  'Rejeição: Total da BC ICMS-ST difere do somatório dos itens'; end;
    534: begin result := false; pMensagem :=  'Rejeição: Total do ICMS-ST difere do somatório dos itens'; end;
    535: begin result := false; pMensagem :=  'Rejeição: Total do Frete difere do somatório dos itens'; end;
    536: begin result := false; pMensagem :=  'Rejeição: Total do Seguro difere do somatório dos itens'; end;
    537: begin result := false; pMensagem :=  'Rejeição: Total do Desconto difere do somatório dos itens'; end;
    538: begin result := false; pMensagem :=  'Rejeição: Total do IPI difere do somatório dos itens'; end;
    539: begin result := false; pMensagem :=  'Duplicidade de NF-e com diferença na Chave de Acesso [chNFe: 99999999999999999999999999999999999999999999][nRec:999999999999999]'; end;
    540: begin result := false; pMensagem :=  'Rejeição: CPF do Local de Retirada inválido'; end;
    541: begin result := false; pMensagem :=  'Rejeição: CPF do Local de Entrega inválido'; end;
    542: begin result := false; pMensagem :=  'Rejeição: CNPJ do Transportador inválido'; end;
    543: begin result := false; pMensagem :=  'Rejeição: CPF do Transportador inválido'; end;
    544: begin result := false; pMensagem :=  'Rejeição: IE do Transportador inválida'; end;
    545: begin result := false; pMensagem :=  'Rejeição: Falha no schema XML – versão informada na versaoDados do SOAPHeader diverge da versão da mensagem'; end;
    546: begin result := false; pMensagem :=  'Rejeição: Erro na Chave de Acesso – Campo Id – falta a literal NFe'; end;
    547: begin result := false; pMensagem :=  'Rejeição: Dígito Verificador da Chave de Acesso da NF-e Referenciada inválido'; end;
    548: begin result := false; pMensagem :=  'Rejeição: CNPJ da NF referenciada inválido.'; end;
    549: begin result := false; pMensagem :=  'Rejeição: CNPJ da NF referenciada de produtor inválido.'; end;
    550: begin result := false; pMensagem :=  'Rejeição: CPF da NF referenciada de produtor inválido.'; end;
    551: begin result := false; pMensagem :=  'Rejeição: IE da NF referenciada de produtor inválido.'; end;
    552: begin result := false; pMensagem :=  'Rejeição: Dígito Verificador da Chave de Acesso do CT-e Referenciado inválido'; end;
    553: begin result := false; pMensagem :=  'Rejeição: Tipo autorizador do recibo diverge do Órgão Autorizador.'; end;
    554: begin result := false; pMensagem :=  'Rejeição: Série difere da faixa 0-899'; end;
    555: begin result := false; pMensagem :=  'Rejeição: Tipo autorizador do protocolo diverge do Órgão Autorizador.'; end;
    556: begin result := false; pMensagem :=  'Rejeição: Justificativa de entrada em contingência não deve ser informada para tipo de emissão normal.'; end;
    557: begin result := false; pMensagem :=  'Rejeição: A Justificativa de entrada em contingência deve ser informada.'; end;
    558: begin result := false; pMensagem :=  'Rejeição: Data de entrada em contingência posterior a data de recebimento.'; end;
    559: begin result := false; pMensagem :=  'Rejeição: UF do Transportador não informada'; end;
    560: begin result := false; pMensagem :=  'Rejeição: CNPJ base do emitente difere do CNPJ base da primeira NF-e do lote recebido'; end;
    561: begin result := false; pMensagem :=  'Rejeição: Mês de Emissão informado na Chave de Acesso difere do Mês de Emissão da NF-e'; end;
    562: begin result := false; pMensagem :=  'Rejeição: Código Numérico informado na Chave de Acesso difere do Código  Numérico da NF-e [chNFe:99999999999999999999999999999999999999999999]'; end;
    563: begin result := false; pMensagem :=  'Rejeição: Já existe pedido de Inutilização com a mesma faixa de inutilização'; end;
    564: begin result := false; pMensagem :=  'Rejeição: Total do Produto / Serviço difere do somatório dos itens'; end;
    565: begin result := false; pMensagem :=  'Rejeição: Falha no schema XML – inexiste a tag raiz esperada para o lote de NF-e'; end;
    567: begin result := false; pMensagem :=  'Rejeição: Falha no schema XML – versão informada na versaoDados do SOAPHeader diverge da versão do lote de NF-e'; end;
    568: begin result := false; pMensagem :=  'Rejeição: Falha no schema XML – inexiste atributo versao na tag raiz do lote de NF-e'; end;
    569: begin result := false; pMensagem :=  'Rejeição: Data de entrada em contingência muito atrasada'; end;
    570: begin result := false; pMensagem :=  'Rejeição: tpEmis = 3 só é válido na contingência SCAN'; end;
    571: begin result := false; pMensagem :=  'Rejeição: O tpEmis informado diferente de 3 para contingência SCAN'; end;
    572: begin result := false; pMensagem :=  'Rejeição: Erro Atributo ID do evento não corresponde a concatenação dos campos (“ID” + tpEvento + chNFe + nSeqEvento)'; end;
    573: begin result := false; pMensagem :=  'Rejeição: Duplicidade de Evento'; end;
    574: begin result := false; pMensagem :=  'Rejeição: O autor do evento diverge do emissor da NF-e'; end;
    575: begin result := false; pMensagem :=  'Rejeição: O autor do evento diverge do destinatário da NF-e'; end;
    576: begin result := false; pMensagem :=  'Rejeição: O autor do evento não é um órgão autorizado a gerar o evento'; end;
    577: begin result := false; pMensagem :=  'Rejeição: A data do evento não pode ser menor que a data de emissão da NF-e'; end;
    578: begin result := false; pMensagem :=  'Rejeição: A data do evento não pode ser maior que a data do processamento'; end;
    579: begin result := false; pMensagem :=  'Rejeição: A data do evento não pode ser menor que a data de autorização para NF-e não emitida em contingência'; end;
    580: begin result := false; pMensagem :=  'Rejeição: O evento exige uma NF-e autorizada'; end;
    587: begin result := false; pMensagem :=  'Rejeição: Usar somente o namespace padrão da NF-e'; end;
    588: begin result := false; pMensagem :=  'Rejeição: Não é permitida a presença de caracteres de edição no início/fim da mensagem ou entre as tags da mensagem'; end;
    589: begin result := false; pMensagem :=  'Rejeição: Número do NSU informado superior ao maior NSU da base de dados da SEFAZ';end;
    590: begin result := false; pMensagem :=  'Rejeição: Informado CST para emissor do Simples Nacional (CRT=1)'; end;
    591: begin result := false; pMensagem :=  'Rejeição: Informado CSOSN para emissor que não é do Simples Nacional (CRT diferente de 1)'; end;
    592: begin result := false; pMensagem :=  'Rejeição: A NF-e deve ter pelo menos um item de produto sujeito ao ICMS'; end;
    593: begin result := false; pMensagem :=  'Rejeição: CNPJ-Base consultado difere do CNPJ-Base do Certificado Digita'; end;
    594: begin result := false; pMensagem :=  'Rejeição: O número de sequencia do evento informado é maior que o permitido'; end;
    595: begin result := false; pMensagem :=  'Rejeição: A versão do leiaute da NF-e utilizada não é mais válida'; end;
    596: begin result := false; pMensagem :=  'Rejeição: Ambiente de homologação indisponível para recepção de NF-e da versão 1.10.'; end;
    597: begin result := false; pMensagem :=  'Rejeição: CFOP de Importação e não informado dados de IPI'; end;
    598: begin result := false; pMensagem :=  'Rejeição: NF-e emitida em ambiente de homologação com Razão Social do destinatário diferente de NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL'; end;
    599: begin result := false; pMensagem :=  'Rejeição: CFOP de Importação e não informado dados de II'; end;
    601: begin result := false; pMensagem :=  'Rejeição: Total do II difere do somatório dos itens'; end;
    602: begin result := false; pMensagem :=  'Rejeição: Total do PIS difere do somatório dos itens sujeitos ao ICMS'; end;
    603: begin result := false; pMensagem :=  'Rejeição: Total do COFINS difere do somatório dos itens sujeitos ao ICMS'; end;
    604: begin result := false; pMensagem :=  'Rejeição: Total do vOutro difere do somatório dos itens'; end;
    605: begin result := false; pMensagem :=  'Rejeição: Total do vISS difere do somatório do vProd dos itens sujeitos ao ISSQN'; end;
    606: begin result := false; pMensagem :=  'Rejeição: Total do vBC do ISS difere do somatório dos itens'; end;
    607: begin result := false; pMensagem :=  'Rejeição: Total do ISS difere do somatório dos itens'; end;
    608: begin result := false; pMensagem :=  'Rejeição: Total do PIS difere do somatório dos itens sujeitos ao ISSQN'; end;
    609: begin result := false; pMensagem :=  'Rejeição: Total do COFINS difere do somatório dos itens sujeitos ao ISSQN'; end;
    610: begin result := false; pMensagem :=  'Rejeição: Total da NF difere do somatório dos Valores compõe o valor Total da NF.'; end;
    611: begin result := false; pMensagem :=  'Rejeição: cEAN inválido'; end;
    612: begin result := false; pMensagem :=  'Rejeição: cEANTrib inválido'; end;
    613: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso difere da existente em BD'; end;
    614: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso inválida (Código UF inválido)'; end;
    615: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso inválida (Ano < 05 ou Ano maior que Ano corrente)'; end;
    616: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso inválida (Mês < 1 ou Mês > 12)'; end;
    617: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso inválida (CNPJ zerado ou dígito inválido)'; end;
    618: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso inválida (modelo diferente de 55)'; end;
    619: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso inválida (número NF = 0)'; end;
    620: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso difere da existente em BD'; end;
    621: begin result := false; pMensagem :=  'Rejeição: CPF Emitente não cadastrado'; end;
    622: begin result := false; pMensagem :=  'Rejeição: IE emitente não vinculada ao CPF'; end;
    623: begin result := false; pMensagem :=  'Rejeição: CPF Destinatário não cadastrado'; end;
    624: begin result := false; pMensagem :=  'Rejeição: IE Destinatário não vinculada ao CPF'; end;
    625: begin result := false; pMensagem :=  'Rejeição: Inscrição SUFRAMA deve ser informada na venda com isenção para ZFM'; end;
    626: begin result := false; pMensagem :=  'Rejeição: O CFOP de operação isenta para ZFM deve ser 6109 ou 6110'; end;
    627: begin result := false; pMensagem :=  'Rejeição: O valor do ICMS desonerado deve ser informado'; end;
    628: begin result := false; pMensagem :=  'Rejeição: Total da NF superior ao valor limite estabelecido pela SEFAZ [Limite]'; end;
    629: begin result := false; pMensagem :=  'Rejeição:  Valor  do  Produto  difere  do  produto  Valor  Unitário  de  Comercialização e Quantidade Comercial'; end;
    630: begin result := false; pMensagem :=  'Rejeição:  Valor do Produto  difere do produto  Valor Unitário de Tributação  e  Quantidade Tributável'; end;
    631: begin result := false; pMensagem :=  'Rejeição: CNPJ-Base do Destinatário difere do CNPJ-Base do Certificado Digital';end;
    632: begin result := false; pMensagem :=  'Rejeição: Solicitação fora de prazo, a NF-e não está mais disponível para download';end;
    633: begin result := false; pMensagem :=  'Rejeição: NF-e indisponível para download devido a ausência de Manifestação do Destinatário';end;
    634: begin result := false; pMensagem :=  'Rejeição: Destinatário da NF-e não tem o mesmo CNPJ raiz do solicitante do download';end;
    635: begin result := false; pMensagem :=  'Rejeição: NF-e com mesmo número e série já transmitida e aguardando processamento'; end;
    650: begin result := false; pMensagem :=  'Rejeição: Evento de "Ciência da Operação" para NF-e Cancelada ou Denegada';end;
    651: begin result := false; pMensagem :=  'Rejeição: Evento de "Desconhecimento da Operação" para NF-e Cancelada ou Denegada';end;
    653: begin result := false; pMensagem :=  'Rejeição: NF-e Cancelada, arquivo indisponível para download';end;
    654: begin result := false; pMensagem :=  'Rejeição: NF-e Denegada, arquivo indisponível para download';end;
    655: begin result := false; pMensagem :=  'Rejeição: Evento de Ciência da Operação informado após a manifestação final do destinatário';end;
    656: begin result := false; pMensagem :=  'Rejeição: Consumo Indevido';end;
    657: begin result := false; pMensagem :=  'Rejeição: Código do Órgão diverge do órgão autorizador';end;
    658: begin result := false; pMensagem :=  'Rejeição: UF do destinatário da Chave de Acesso diverge da UF autorizadora';end;
    660: begin result := false; pMensagem :=  'Rejeição: CFOP de Combustível e não informado grupo de combustível [nItem:nnn]'; end;
    661: begin result := false; pMensagem :=  'Rejeição: NF-e já existente para o número do EPEC informado'; end;
    662: begin result := false; pMensagem :=  'Rejeição: Numeração do EPEC está inutilizada na Base de Dados da SEFAZ'; end;
    663: begin result := false; pMensagem :=  'Rejeição: Alíquota do ICMS com valor superior a 4 por cento na operação de saída interestadual com produtos importados [nItem:999]'; end;
    678: begin result := false; pMensagem :=  'Rejeição: NF referenciada com UF diferente da NF-e complementar'; end;
    679: begin result := false; pMensagem :=  'Rejeição: Modelo de DF-e referenciado inválido'; end;
    680: begin result := false; pMensagem :=  'Rejeição: Duplicidade de NF-e referenciada (Chave de Acesso referenciada mais de uma vez)'; end;
    681: begin result := false; pMensagem :=  'Rejeição: Duplicidade de NF Modelo 1 referenciada (CNPJ, Modelo, Série e Número)'; end;
    682: begin result := false; pMensagem :=  'Rejeição: Duplicidade de NF de Produtor referenciada (IE, Modelo, Série e Número)'; end;
    683: begin result := false; pMensagem :=  'Rejeição: Modelo do CT-e referenciado diferente de 57'; end;
    684: begin result := false; pMensagem :=  'Rejeição: Duplicidade de Cupom Fiscal referenciado (Modelo, Número de Ordem e COO)'; end;
    685: begin result := false; pMensagem :=  'Rejeição: Total do Valor Aproximado dos Tributos difere do somatório dos itens'; end;
    686: begin result := false; pMensagem :=  'Rejeição: NF Complementar referencia uma NF-e cancelada'; end;
    687: begin result := false; pMensagem :=  'Rejeição: NF Complementar referencia uma NF-e denegada'; end;
    688: begin result := false; pMensagem :=  'Rejeição: NF referenciada de Produtor com IE inexistente [nRef: xxx]'; end;
    689: begin result := false; pMensagem :=  'Rejeição: NF referenciada de Produtor com IE não vinculada ao CNPJ/CPF informado [nRef: xxx]'; end;
    690: begin result := false; pMensagem :=  'Rejeição: Pedido de Cancelamento para NF-e com CT-e'; end;
    691: begin result := false; pMensagem :=  'Rejeição: Chave de Acesso da NF-e diverge da Chave de Acesso do EPEC'; end;
    693: begin result := false; pMensagem :=  'Rejeição: Alíquota de ICMS superior a definida para a operação interestadual [nItem:999]'; end;
    694: begin result := false; pMensagem :=  'Rejeição: Não informado o grupo de ICMS para a UF de destino [nItem:999]'; end;
    695: begin result := false; pMensagem :=  'Rejeição: Informado indevidamente o grupo de ICMS para a UF de destino [nItem:999]'; end;
    697: begin result := false; pMensagem :=  'Rejeição: Alíquota interestadual do ICMS com origem diferente do previsto [nItem:999]'; end;
    698: begin result := false; pMensagem :=  'Rejeição: Alíquota interestadual do ICMS incompatível com as UF envolvidas na operação [nItem:999]'; end;
    699: begin result := false; pMensagem :=  'Rejeição: Percentual do ICMS Interestadual para a UF de destino difere do previsto para o ano da Data de Emissão [nItem:999]'; end;
    700: begin result := false; pMensagem :=  'Rejeição: Mensagem de Lote versão 3.xx. Enviar para o Web Service nfeAutorizacao'; end;
    701: begin result := false; pMensagem :=  'Rejeição: Não informado Nota Fiscal referenciada (CFOP de Exportação Indireta)'; end;
    702: begin result := false; pMensagem :=  'Rejeição: NFC-e não é aceita pela UF do Emitente'; end;
    703: begin result := false; pMensagem :=  'Rejeição: Data-Hora de Emissão posterior ao horário de recebimento'; end;
    704: begin result := false; pMensagem :=  'Rejeição: NFC-e com Data-Hora de emissão atrasada'; end;
    705: begin result := false; pMensagem :=  'Rejeição: NFC-e com data de entrada/saída'; end;
    706: begin result := false; pMensagem :=  'Rejeição: NFC-e para operação de entrada'; end;
    707: begin result := false; pMensagem :=  'Rejeição: NFC-e para operação interestadual ou com o exterior'; end;
    708: begin result := false; pMensagem :=  'Rejeição: NFC-e não pode referenciar documento fiscal'; end;
    709: begin result := false; pMensagem :=  'Rejeição: NFC-e com formato de DANFE inválido'; end;
    710: begin result := false; pMensagem :=  'Rejeição: NF-e com formato de DANFE inválido'; end;
    711: begin result := false; pMensagem :=  'Rejeição: NF-e com contingência off-line'; end;
    712: begin result := false; pMensagem :=  'Rejeição: NFC-e com contingência off-line para a UF'; end;
    713: begin result := false; pMensagem :=  'Rejeição: Tipo de Emissão diferente de 6 ou 7 para contingência da SVC acessada'; end;
    714: begin result := false; pMensagem :=  'Rejeição: NFC-e com opção de contingência inválida (tpEmis=2, 4 (a critério da UF) ou 5)'; end;
    715: begin result := false; pMensagem :=  'Rejeição: NFC-e com finalidade inválida'; end;
    716: begin result := false; pMensagem :=  'Rejeição: NFC-e em operação não destinada a consumidor final'; end;
    717: begin result := false; pMensagem :=  'Rejeição: NFC-e em operação não presencial'; end;
    718: begin result := false; pMensagem :=  'Rejeição: NFC-e não deve informar IE de Substituto Tributário'; end;
    719: begin result := false; pMensagem :=  'Rejeição: NF-e sem a identificação do destinatário'; end;
    720: begin result := false; pMensagem :=  'Rejeição: Na operação com Exterior deve ser informada tag idEstrangeiro'; end;
    721: begin result := false; pMensagem :=  'Rejeição: Operação interestadual deve informar CNPJ ou CPF'; end;
    723: begin result := false; pMensagem :=  'Rejeição: Operação interna com idEstrangeiro informado deve ser para consumidor final'; end;
    724: begin result := false; pMensagem :=  'Rejeição: NF-e sem o nome do destinatário'; end;
    725: begin result := false; pMensagem :=  'Rejeição: NFC-e com CFOP inválido [nItem:nnn]'; end;
    726: begin result := false; pMensagem :=  'Rejeição: NF-e sem a informação de endereço do destinatário'; end;
    727: begin result := false; pMensagem :=  'Rejeição: Operação com Exterior e UF diferente de EX'; end;
    728: begin result := false; pMensagem :=  'Rejeição: NF-e sem informação da IE do destinatário'; end;
    729: begin result := false; pMensagem :=  'Rejeição: NFC-e com informação da IE do destinatário'; end;
    730: begin result := false; pMensagem :=  'Rejeição: NFC-e com Inscrição Suframa'; end;
    731: begin result := false; pMensagem :=  'Rejeição: CFOP de operação com Exterior e idDest <> 3'; end;
    732: begin result := false; pMensagem :=  'Rejeição: CFOP de operação interestadual e idDest <> 2'; end;
    733: begin result := false; pMensagem :=  'Rejeição: CFOP de operação interna e idDest <> 1'; end;
    734: begin result := false; pMensagem :=  'Rejeição: NFC-e com Unidade de Comercialização inválida'; end;
    735: begin result := false; pMensagem :=  'Rejeição: NFC-e com Unidade de Tributação inválida'; end;
    736: begin result := false; pMensagem :=  'Rejeição: NFC-e com grupo de Veículos novos'; end;
    737: begin result := false; pMensagem :=  'Rejeição: NFC-e com grupo de Medicamentos'; end;
    738: begin result := false; pMensagem :=  'Rejeição: NFC-e com grupo de Armamentos'; end;
    740: begin result := false; pMensagem :=  'Rejeição: Item com Repasse de ICMS retido por Substituto Tributário [nItem:nnn]'; end;
    741: begin result := false; pMensagem :=  'Rejeição: NFC-e com Partilha de ICMS entre UF'; end;
    742: begin result := false; pMensagem :=  'Rejeição: NFC-e com grupo do IPI'; end;
    743: begin result := false; pMensagem :=  'Rejeição: NFC-e com grupo do II'; end;
    745: begin result := false; pMensagem :=  'Rejeição: NF-e sem grupo do PIS'; end;
    746: begin result := false; pMensagem :=  'Rejeição: NFC-e com grupo do PIS-ST'; end;
    748: begin result := false; pMensagem :=  'Rejeição: NF-e sem grupo da COFINS'; end;
    749: begin result := false; pMensagem :=  'Rejeição: NFC-e com grupo da COFINS-ST'; end;
    750: begin result := false; pMensagem :=  'Rejeição: NFC-e com valor total superior ao permitido para destinatário não identificado (Código) [Limite]'; end;
    751: begin result := false; pMensagem :=  'Rejeição: NFC-e com valor total superior ao permitido para destinatário não identificado (Nome) [Limite]'; end;
    752: begin result := false; pMensagem :=  'Rejeição: NFC-e com valor total superior ao permitido para destinatário não identificado (Endereço) [Limite]'; end;
    753: begin result := false; pMensagem :=  'Rejeição: NFC-e com Frete'; end;
    754: begin result := false; pMensagem :=  'Rejeição: NFC-e com dados do Transportador'; end;
    755: begin result := false; pMensagem :=  'Rejeição: NFC-e com dados de Retenção do ICMS no Transporte'; end;
    756: begin result := false; pMensagem :=  'Rejeição: NFC-e com dados do veículo de Transporte'; end;
  else
    begin
      result := BuscaStatus2(pStatus,pMensagem, pNFCE);
    end;
  end;
end;

function TNFEInterfaceV3.BuscaStatus2(var pStatus: Integer; var pMensagem: string; pNFCE: Boolean = false): Boolean;
begin   //maximo 470 constantes
  case pStatus of
    757: begin result := false; pMensagem :=  'Rejeição: NFC-e com dados de Reboque do veículo de Transporte'; end;
    758: begin result := false; pMensagem :=  'Rejeição: NFC-e com dados do Vagão de Transporte'; end;
    759: begin result := false; pMensagem :=  'Rejeição: NFC-e com dados da Balsa de Transporte'; end;
    760: begin result := false; pMensagem :=  'Rejeição: NFC-e com dados de cobrança (Fatura, Duplicata)'; end;
    762: begin result := false; pMensagem :=  'Rejeição: NFC-e com dados de compras (Empenho, Pedido, Contrato)'; end;
    763: begin result := false; pMensagem :=  'Rejeição: NFC-e com dados de aquisição de Cana'; end;
    764: begin result := false; pMensagem :=  'Rejeição: Solicitada resposta síncrona para Lote com mais de uma NF-e (indSinc=1)'; end;
    765: begin result := false; pMensagem :=  'Rejeição: Lote só poderá conter NF-e ou NFC-e'; end;
    766: begin result := false; pMensagem :=  'Rejeição: Item com CST indevido [nItem:nnn]'; end;
    767: begin result := false; pMensagem :=  'Rejeição: NFC-e com somatório dos pagamentos diferente do total da Nota Fiscal'; end;
    768: begin result := false; pMensagem :=  'Rejeição: NF-e não deve possuir o grupo de Formas de Pagamento'; end;
    769: begin result := false; pMensagem :=  'Rejeição: A critério da UF NFC-e deve possuir o grupo de Formas de Pagamento'; end;
    770: begin result := false; pMensagem :=  'Rejeição: NFC-e autorizada há mais de 24 horas.'; end;
    771: begin result := false; pMensagem :=  'Rejeição: Operação Interestadual e UF de destino com EX'; end;
    772: begin result := false; pMensagem :=  'Rejeição: Operação Interestadual e UF de destino igual à UF do emitente'; end;
    773: begin result := false; pMensagem :=  'Rejeição: Operação Interna e UF de destino difere da UF do emitente'; end;
    774: begin result := false; pMensagem :=  'Rejeição: NFC-e com indicador de item não participante do total'; end;
    775: begin result := false; pMensagem :=  'Rejeição: Modelo da NFC-e diferente de 65'; end;
    776: begin result := false; pMensagem :=  'Rejeição: Solicitada resposta síncrona para UF que não disponibiliza este atendimento (indSinc=1)'; end;
    777: begin result := false; pMensagem :=  'Rejeição: Obrigatória a informação do NCM completo'; end;
    778: begin result := false; pMensagem :=  'Rejeição: Informado NCM inexistente [nItem:nnn]'; end;
    779: begin result := false; pMensagem :=  'Rejeição: NFC-e com NCM incompatível'; end;
    780: begin result := false; pMensagem :=  'Rejeição: Total da NFC-e superior ao valor limite estabelecido pela SEFAZ [Limite]'; end;
    781: begin result := false; pMensagem :=  'Rejeição: Emissor não habilitado para emissão da NFC-e'; end;
    782: begin result := false; pMensagem :=  'Rejeição: NFC-e não é autorizada pelo SCAN'; end;
    783: begin result := false; pMensagem :=  'Rejeição: NFC-e não é autorizada pela SVC'; end;
    784: begin result := false; pMensagem :=  'Rejeição: NFC-e não permite o evento de Carta de Correção'; end;
    785: begin result := false; pMensagem :=  'Rejeição: NFC-e com entrega a domicílio não permitida pela UF'; end;
    786: begin result := false; pMensagem :=  'Rejeição: NFC-e de entrega a domicílio sem dados do Transportador'; end;
    787: begin result := false; pMensagem :=  'Rejeição: NFC-e de entrega a domicílio sem a identificação do destinatário'; end;
    788: begin result := false; pMensagem :=  'Rejeição: NFC-e de entrega a domicílio sem o endereço do destinatário'; end;
    789: begin result := false; pMensagem :=  'Rejeição: NFC-e para destinatário contribuinte de ICMS'; end;
    790: begin result := false; pMensagem :=  'Rejeição: Operação com Exterior para destinatário Contribuinte de ICMS'; end;
    791: begin result := false; pMensagem :=  'Rejeição: NF-e com indicação de destinatário isento de IE, com a informação da IE do destinatário'; end;
    792: begin result := false; pMensagem :=  'Rejeição: Informada a IE do destinatário para operação com destinatário no Exterior'; end;
    793: begin result := false; pMensagem :=  'Rejeição: Valor do ICMS relativo ao Fundo de Combate à Pobreza na UF de destino difere do calculado [nItem:999]'; end;
    794: begin result := false; pMensagem :=  'Rejeição: NF-e com indicativo de NFC-e com entrega a domicílio'; end;
    795: begin result := false; pMensagem :=  'Rejeição: Total do ICMS desonerado difere do somatório dos itens'; end;
    796: begin result := false; pMensagem :=  'Rejeição: Empresa sem Chave de Segurança para o QR-Code'; end;
    798: begin result := false; pMensagem :=  'Rejeição: Valor total do ICMS relativo Fundo de Combate à Pobreza (FCP) da UF de destino difere do somatório do valor dos itens'; end;
    799: begin result := false; pMensagem :=  'Rejeição: Valor total do ICMS Interestadual da UF de destino difere do somatório dos itens'; end;
    800: begin result := false; pMensagem :=  'Rejeição: Valor total do ICMS Interestadual da UF do remetente difere do somatório dos itens'; end;
    805: begin result := false; pMensagem :=  'Rejeição: A SEFAZ do destinatário não permite Contribuinte Isento de Inscrição Estadual'; end;
    806: begin result := false; pMensagem :=  'Rejeição: Operação com ICMS-ST sem informação do CEST'; end;
    807: begin result := false; pMensagem :=  'Rejeição: NFC-e com grupo de ICMS para a UF do destinatário'; end;
    999: begin result := false; pMensagem :=  'Rejeição: Erro não catalogado'; end;
  end;
end;

procedure TNFEInterfaceV3.SetNaturezaOperacao(pValue: String);
begin
  fCapaNFe.NaturezaOperacao := pValue;
end;
procedure TNFEInterfaceV3.SetModelo(pValue: Integer);
begin
  fCapaNFe.Modelo := pValue;
end;
procedure TNFEInterfaceV3.SetCodigo(pValue: Integer);
begin
  fCapaNfe.Codigo := pValue;
end;
procedure TNFEInterfaceV3.SetNumero(pValue: Integer);
begin
  fCapaNfe.Numero := pValue;
end;
procedure TNFEInterfaceV3.SetSerie(pValue: Integer);
begin
  fCapaNfe.Serie := pValue;
end;
procedure TNFEInterfaceV3.SetEmissao(pValue: String);
begin
  fCapaNfe.Emissao := pValue;
end;
procedure TNFEInterfaceV3.SetSaida(pValue: String);
begin
  fCapaNfe.Saida := pValue;
end;
procedure TNFEInterfaceV3.SetTipo(pValue: Integer);
begin
  fCapaNfe.Tipo := pValue;
end;
procedure TNFEInterfaceV3.SetFormaPag(pValue: String);
begin
  fCapaNfe.FormaPag := pValue;
end;
procedure TNFEInterfaceV3.SetidDest(pValue: Integer);
begin
  fCapaNfe.idDest := pValue;
end;
procedure TNFEInterfaceV3.SetindFinal(pValue: Integer);
begin
  fCapaNfe.indFinal := pValue;
end;
procedure TNFEInterfaceV3.SetindPres(pValue: Integer);
begin
  fCapaNfe.indPres := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteCNPJ(pValue: String);
begin
  fCapaNfe.EmitenteCNPJ := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteIE(pValue: String);
begin
  fCapaNfe.EmitenteIE := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteIM(pValue: String);
begin
  fCapaNfe.EmitenteIM := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteCNAE(pValue: String);
begin
  fCapaNfe.EmitenteCNAE := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteRazao(pValue: String);
begin
  fCapaNfe.EmitenteRazao := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteFantasia(pValue: String);
begin
  fCapaNfe.EmitenteFantasia := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteFone(pValue: String);
begin
  fCapaNfe.EmitenteFone := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteCEP(pValue: String);
begin
  fCapaNfe.EmitenteCEP := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteLogradouro(pValue: String);
begin
  fCapaNfe.EmitenteLogradouro := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteNumero(pValue: Integer);
begin
  fCapaNfe.EmitenteNumero := IntToStr(pValue);
end;
procedure TNFEInterfaceV3.SetEmitenteComplemento(pValue: String);
begin
  fCapaNfe.EmitenteComplemento := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteBairro(pValue: String);
begin
  fCapaNfe.EmitenteBairro := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteCidadeCod(pValue: String);
begin
  fCapaNfe.EmitenteCidadeCod := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteCidade(pValue: String);
begin
  fCapaNfe.EmitenteCidade := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteUF(pValue: String);
begin
  fCapaNfe.EmitenteUF := pValue;
end;
procedure TNFEInterfaceV3.SetEmitentePaisCod(pValue: String);
begin
  fCapaNfe.EmitentePaisCod := pValue;
end;
procedure TNFEInterfaceV3.SetEmitentePais(pValue: String);
begin
  fCapaNfe.EmitentePais := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioCNPJ(pValue: String);
begin
  fCapaNfe.DestinatarioCNPJ := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioIE(pValue: String);
begin
  fCapaNfe.DestinatarioIE := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioNomeRazao(pValue: String);
begin
  fCapaNfe.DestinatarioNomeRazao := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioFone(pValue: String);
begin
  fCapaNfe.DestinatarioFone := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioCEP(pValue: String);
begin
  fCapaNfe.DestinatarioCEP := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioLogradouro(pValue: String);
begin
  fCapaNfe.DestinatarioLogradouro := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioNumero(pValue: String);
begin
  fCapaNfe.DestinatarioNumero := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioComplemento(pValue: String);
begin
  fCapaNfe.DestinatarioComplemento := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioBairro(pValue: String);
begin
  fCapaNfe.DestinatarioBairro := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioCidadeCod(pValue: String);
begin
  fCapaNfe.DestinatarioCidadeCod := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioCidade(pValue: String);
begin
  fCapaNfe.DestinatarioCidade := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioUF(pValue: String);
begin
  fCapaNfe.DestinatarioUF := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioindIEDest(pValue: String);
begin
  fCapaNfe.DestinatarioindIEDest := pValue;
end;
procedure TNFEInterfaceV3.SetCNPJautXML(pValue: String);
begin
  fCapaNfe.CNPJautXML := pValue;
end;
procedure TNFEInterfaceV3.SetCPFautXML(pValue: String);
begin
  fCapaNfe.CPFautXML := pValue;
end;
procedure TNFEInterfaceV3.SetDestinatarioCodigo(pValue: Integer);
begin
  fCapaNfe.DestinatarioCodigo := pValue;
end;
procedure TNFEInterfaceV3.SettpImp(pValue: Integer);
begin
  fCapaNfe.tpImp := pValue;
end;
procedure TNFEInterfaceV3.SetinfCpl(pValue: String);
begin
  fCapaNfe.infCpl := pValue;
end;
procedure TNFEInterfaceV3.SetCFOP(pValue: String);
begin
  fCapaNfe.CFOP := pValue;
end;
procedure TNFEInterfaceV3.SetrefNF(pValue: String);
begin
  fCapaNfe.refNF := pValue;
end;
procedure TNFEInterfaceV3.SetInfAdFisco(pValue: String);
begin
  fCapaNfe.InfAdFisco := pValue;
end;
procedure TNFEInterfaceV3.SettpEmis(pValue: Integer);
begin
  fCapaNfe.tpEmis := pValue;
end;
procedure TNFEInterfaceV3.SetEmitenteCRT(pValue: String);
begin
  fCapaNfe.EmitenteCRT := pValue;
end;
procedure TNFEInterfaceV3.SetHoraSaida(pValue: String);
begin
  fCapaNfe.HoraSaida := pValue;
end;
procedure TNFEInterfaceV3.SetTransportadorFretePorConta(pValue: String);
begin
  fCapaNfe.TransportadorFretePorConta := pValue;
end;
procedure TNFEInterfaceV3.SetStatus(pValue: String);
begin
  fCapaNfe.SetStatus(pValue);
end;
end.
