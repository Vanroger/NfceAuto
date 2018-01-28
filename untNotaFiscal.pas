unit untNotaFiscal;

interface

uses
   Adodb, untdm, untNotaItem, Classes, sysutils, dialogs, //untFuncoes,
   pcnConversaoNFe,  untHistoricoXml;

Type
  TNotaFiscal = class
    private
    fListaNotaItem         : TList;
    fListaNFeHistoricoXML  : Tlist;
    fQryNota               : TADOQuery;
    fQryNotaItem           : TADOQuery;
    fQryNFeHistorico_XML   : TADOQuery;
    fNumeroNF              : integer;   //  int NOT NULL,
    fSaidaEntrada          : string;    //  char(1) COLLATE Latin1_General_CI_AS NULL,
    fSerie                 : string;    //  char(3) COLLATE Latin1_General_CI_AS NOT NULL,
    fCFOP                  : string;    //  char(4) COLLATE Latin1_General_CI_AS NULL,
    fInscricaoSubstituicao : string;    //  char(10) COLLATE Latin1_General_CI_AS NULL,
    fInscricaoestadual     : string;    //  char(20) COLLATE Latin1_General_CI_AS NULL,
    fCodCliente            : integer;   //  int NOT NULL,
    fNome                  : string;    //  varchar(100) COLLATE Latin1_General_CI_AS NULL,
    fCnpjCpf               : string;    //  char(18) COLLATE Latin1_General_CI_AS NULL,
    fEmissao               : TDateTime; //  datetime NULL,
    fSaida                 : TDateTime; //  datetime NULL,
    fHoraSaida             : string;    //  char(10) COLLATE Latin1_General_CI_AS NULL,
    fEndereco              : String;    //  char(40) COLLATE Latin1_General_CI_AS NULL,
    fDest_complemento      : String;    //  varchar(50) COLLATE Latin1_General_CI_AS NULL,
    fDest_Numero           : Integer;   //  int NULL,
    fBairro                : String;    //  char(20) COLLATE Latin1_General_CI_AS NULL,
    fCEP                   : String;    //  char(10) COLLATE Latin1_General_CI_AS NULL,
    fMunicipio             : string;    //  char(40) COLLATE Latin1_General_CI_AS NULL,
    fFoneFax               : string;    //  char(12) COLLATE Latin1_General_CI_AS NULL,
    fUF                    : string;    //  char(10) COLLATE Latin1_General_CI_AS NULL,
    fBaseIcms              : Double;    //  float NULL,
    fIcms                  : Double;    //  float NULL,
    fBaseSubstituicao      : Double;    //  float NULL,
    fValorSubstituicao     : Double;    //  float NULL,
    fTotalProduto          : Double;    //  float NULL,
    fTotalNota             : Double;    //  float NULL,
    fFrete                 : Double;    //  float NULL,
    fSeguro                : Double;    //  float NULL,
    fOutros                : Double;    //  float NULL,
    fIPI                   : Double;    //  float NULL,
    fObservacao            : string;    //  ntext COLLATE Latin1_General_CI_AS NULL,
    fNaturezaOperacao      : string;    //  char(50) COLLATE Latin1_General_CI_AS NULL,
    fVALORDESCONTO         : Double;    //  real DEFAULT 0 NOT NULL,
    fCancelada             : string;    //  varchar(1) COLLATE Latin1_General_CI_AS DEFAULT 'N' NOT NULL,
    fModelo                : string;    //  varchar(4) COLLATE Latin1_General_CI_AS DEFAULT '55' NULL,
    fVALORISENTO           : Double;    //  numeric(13, 2) NULL,
    fVALOROUTRAS           : Double;    //  numeric(13, 2) NULL,
    fALIQICMS              : Double;    //  numeric(13, 2) NULL,
    fVERSAO_NFE            : string;    //  varchar(4) COLLATE Latin1_General_CI_AS NULL,
    fJUSTIFICATIVA_CONTINGENCIA    : string;    //  varchar(1000) COLLATE Latin1_General_CI_AS NULL,
    fDATA_HORA_CONTINGENCIA        : TDateTime; //  datetime NULL,
    fDATA_HORA_RECEB_NFE           : string;    //  varchar(19) COLLATE Latin1_General_CI_AS NULL,
    fSTATUS_CTG                    : string;    //  char(1) COLLATE Latin1_General_CI_AS NULL,
    fTIPO_AMBIENTE_NFE             : string;    //  char(1) COLLATE Latin1_General_CI_AS NULL,
    fNUMERO_LOTE_NFE               : string;   //  varchar(15) COLLATE Latin1_General_CI_AS NULL,
    fNUMERO_PROTOCOLO_CANCELAMENTO : string;  //  varchar(15) COLLATE Latin1_General_CI_AS NULL,
    fNUMERO_PROTOCOLO              : string;  //  varchar(15) COLLATE Latin1_General_CI_AS NULL,
    fNOTA_FISCAL_NFE               : string;  //  varchar(9) COLLATE Latin1_General_CI_AS NULL,
    fSTATUS_NFE                    : string;  //  char(1) COLLATE Latin1_General_CI_AS NULL,
    fNUMERO_RECIBO                 : string;  //  varchar(15) COLLATE Latin1_General_CI_AS NULL,
    fNUMERO_NFE                    : string;  //  varchar(44) COLLATE Latin1_General_CI_AS NULL,
    fCONTINGENCIA                  : string;  //  char(1) COLLATE Latin1_General_CI_AS NULL,
    fMOTIVO_CANCELAMENTO           : string;  //  varchar(50) COLLATE Latin1_General_CI_AS NULL,
    fCOD_SIT_EFD                   : string;  //  char(2) COLLATE Latin1_General_CI_AS NULL,
    fMENSAGEM_FISCO_ID             : integer; //  int NULL,
    fMENSAGEM_CONTRIBUINTE_ID      : integer; //  int NULL,
    fCONDICAO_TIPO                 : string;  //  int NULL,
    fCONDICAO_DESCRICAO            : string;  //  varchar(20) COLLATE Latin1_General_CI_AS NULL,
    fTOTAL_OUTRAS_DESP             : Double;  //  float NULL,
    fEMPRESA_ID                    : string;   //  varchar(18) COLLATE Latin1_General_CI_AS NULL,
    fSTATUS                        : string;    //  char(1) COLLATE Latin1_General_CI_AS NULL,
    fEMITIU                        : string;    //  char(1) COLLATE Latin1_General_CI_AS NULL,
    fMENSAGEM_ID1                  : Integer;   //  int NULL,
    fMENSAGEM_ID2                  : Integer;   //  int NULL,
    fDATA_CANCELAMENTO             : TdateTime; //  datetime NULL,
    fDESC_NAT_OPERACAO             : string;    //  varchar(100) COLLATE Latin1_General_CI_AS NULL,
    fENDERECO_XML                  : string;    //  varchar(300) COLLATE Latin1_General_CI_AS NULL,
    fMENSAGEM1                     : string;    //  varchar(400) COLLATE Latin1_General_CI_AS NULL,
    fMENSAGEM2                     : string;    //  varchar(400) COLLATE Latin1_General_CI_AS NULL,
    fMENSAGEM3                     : string;    //  varchar(400) COLLATE Latin1_General_CI_AS NULL,
    fMensagemComplementar          : string;    //  text COLLATE Latin1_General_CI_AS NULL,
    frefNF                         : string;    //  varchar(44) COLLATE Latin1_General_CI_AS NULL,
    fTIPO_DENEGADA                 : string;    //  varchar(2) COLLATE Latin1_General_CI_AS NULL,

    //CAMPOS QUE NÃO TEM NO BANCO DE DADOS
    fInfoAdicional                 : String;
    fInfAdFisco                    : String;
    fTpEmiss                       : Integer;   //Tipo de Emissao 1=Emissão normal 9=Contingência off-line da NFC-e;
    fCodCidade                     : String;
    fFinalidade                    : TpcnFinalidadeNFe;
    fTipo                          : integer;
    fIE_Destinatario               : String;
    fIndicador                     : Integer;
    fIndFinal                      : Integer;

    function PreencherNotaItem(pNota_Fiscal: string; pModelo, pSerie: integer): boolean;
    procedure LimpaListaNotaItem;
    procedure MensagemAdicional;
    procedure MensagemFisco;
    procedure LimpaCampos;
    procedure SetTipoEmissao;
    function preencheNFE_HISTORICO_XML(pNota_fiscal: integer; pModelo, pSerie: String): boolean;
    public
      constructor create;
      function PreencherNota(pNota_Fiscal : string; pNota_Fiscal_Nfe : string; pModelo : integer; pSerie  : integer): boolean;
    published
      property NumeroNF                      : integer   read fNumeroNF                      write fNumeroNF;
      property SaidaEntrada                  : string    read fSaidaEntrada                  write fSaidaEntrada;
      property Serie                         : string    read fSerie                         write fSerie;
      property CFOP                          : string    read fCFOP                          write fCFOP;
      property InscricaoSubstituicao         : string    read fInscricaoSubstituicao         write fInscricaoSubstituicao;
      property Inscricaoestadual             : string    read fInscricaoestadual             write fInscricaoestadual;
      property CodCliente                    : integer   read fCodCliente                    write fCodCliente;
      property Nome                          : string    read fNome                          write fNome;
      property CnpjCpf                       : string    read fCnpjCpf                       write fCnpjCpf;
      property Emissao                       : TDateTime read fEmissao                       write fEmissao;
      property Saida                         : TDateTime read fSaida                         write fSaida;
      property HoraSaida                     : string    read fHoraSaida                     write fHoraSaida;
      property Endereco                      : String    read fEndereco                      write fEndereco;
      property Dest_complemento              : String    read fDest_complemento              write fDest_complemento;
      property Dest_Numero                   : Integer   read fDest_Numero                   write fDest_Numero;
      property Bairro                        : String    read fBairro                        write fBairro;
      property CEP                           : String    read fCEP                           write fCEP;
      property Municipio                     : string    read fMunicipio                     write fMunicipio;
      property FoneFax                       : string    read fFoneFax                       write fFoneFax;
      property UF                            : string    read fUF                            write fUF;
      property BaseIcms                      : Double    read fBaseIcms                      write fBaseIcms;
      property Icms                          : Double    read fIcms                          write fIcms;
      property BaseSubstituicao              : Double    read fBaseSubstituicao              write fBaseSubstituicao;
      property ValorSubstituicao             : Double    read fValorSubstituicao             write fValorSubstituicao;
      property TotalProduto                  : Double    read fTotalProduto                  write fTotalProduto;
      property TotalNota                     : Double    read fTotalNota                     write fTotalNota;
      property Frete                         : Double    read fFrete                         write fFrete;
      property Seguro                        : Double    read fSeguro                        write fSeguro;
      property Outros                        : Double    read fOutros                        write fOutros;
      property IPI                           : Double    read fIPI                           write fIPI;
      property Observacao                    : string    read fObservacao                    write fObservacao;
      property NaturezaOperacao              : string    read fNaturezaOperacao              write fNaturezaOperacao;
      property VALORDESCONTO                 : Double    read fVALORDESCONTO                 write fVALORDESCONTO;
      property Cancelada                     : string    read fCancelada                     write fCancelada;
      property Modelo                        : string    read fModelo                        write fModelo;
      property VALORISENTO                   : Double    read fVALORISENTO                   write fVALORISENTO;
      property VALOROUTRAS                   : Double    read fVALOROUTRAS                   write fVALOROUTRAS;
      property ALIQICMS                      : Double    read fALIQICMS                      write fALIQICMS;
      property VERSAO_NFE                    : string    read fVERSAO_NFE                    write fVERSAO_NFE;
      property JUSTIFICATIVA_CONTINGENCIA    : string    read fJUSTIFICATIVA_CONTINGENCIA    write fJUSTIFICATIVA_CONTINGENCIA;
      property DATA_HORA_CONTINGENCIA        : TDateTime read fDATA_HORA_CONTINGENCIA        write fDATA_HORA_CONTINGENCIA;
      property DATA_HORA_RECEB_NFE           : string    read fDATA_HORA_RECEB_NFE           write fDATA_HORA_RECEB_NFE;
      property STATUS_CTG                    : string    read fSTATUS_CTG                    write fSTATUS_CTG;
      property TIPO_AMBIENTE_NFE             : string    read fTIPO_AMBIENTE_NFE             write fTIPO_AMBIENTE_NFE;
      property NUMERO_LOTE_NFE               : string    read fNUMERO_LOTE_NFE               write fNUMERO_LOTE_NFE;
      property NUMERO_PROTOCOLO_CANCELAMENTO : string    read fNUMERO_PROTOCOLO_CANCELAMENTO write fNUMERO_PROTOCOLO_CANCELAMENTO;
      property NUMERO_PROTOCOLO              : string    read fNUMERO_PROTOCOLO              write fNUMERO_PROTOCOLO;
      property NOTA_FISCAL_NFE               : string    read fNOTA_FISCAL_NFE               write fNOTA_FISCAL_NFE;
      property STATUS_NFE                    : string    read fSTATUS_NFE                    write fSTATUS_NFE;
      property NUMERO_RECIBO                 : string    read fNUMERO_RECIBO                 write fNUMERO_RECIBO;
      property NUMERO_NFE                    : string    read fNUMERO_NFE                    write fNUMERO_NFE;
      property CONTINGENCIA                  : string    read fCONTINGENCIA                  write fCONTINGENCIA;
      property MOTIVO_CANCELAMENTO           : string    read fMOTIVO_CANCELAMENTO           write fMOTIVO_CANCELAMENTO;
      property COD_SIT_EFD                   : string    read fCOD_SIT_EFD                   write fCOD_SIT_EFD;
      property MENSAGEM_FISCO_ID             : integer   read fMENSAGEM_FISCO_ID             write fMENSAGEM_FISCO_ID;
      property MENSAGEM_CONTRIBUINTE_ID      : integer   read fMENSAGEM_CONTRIBUINTE_ID      write fMENSAGEM_CONTRIBUINTE_ID;
      property CONDICAO_TIPO                 : string    read fCONDICAO_TIPO                 write fCONDICAO_TIPO;
      property CONDICAO_DESCRICAO            : string    read fCONDICAO_DESCRICAO            write fCONDICAO_DESCRICAO;
      property TOTAL_OUTRAS_DESP             : Double    read fTOTAL_OUTRAS_DESP             write fTOTAL_OUTRAS_DESP;
      property EMPRESA_ID                    : string    read fEMPRESA_ID                    write fEMPRESA_ID;
      property STATUS                        : string    read fSTATUS                        write fSTATUS;
      property EMITIU                        : string    read fEMITIU                        write fEMITIU;
      property MENSAGEM_ID1                  : Integer   read fMENSAGEM_ID1                  write fMENSAGEM_ID1;
      property MENSAGEM_ID2                  : Integer   read fMENSAGEM_ID2                  write fMENSAGEM_ID2;
      property DATA_CANCELAMENTO             : TdateTime read fDATA_CANCELAMENTO             write fDATA_CANCELAMENTO;
      property DESC_NAT_OPERACAO             : string    read fDESC_NAT_OPERACAO             write fDESC_NAT_OPERACAO;
      property ENDERECO_XML                  : string    read fENDERECO_XML                  write fENDERECO_XML;
      property MENSAGEM1                     : string    read fMENSAGEM1                     write fMENSAGEM1;
      property MENSAGEM2                     : string    read fMENSAGEM2                     write fMENSAGEM2;
      property MENSAGEM3                     : string    read fMENSAGEM3                     write fMENSAGEM3;
      property MensagemComplementar          : string    read fMensagemComplementar          write fMensagemComplementar;
      property refNF                         : string    read frefNF                         write frefNF;
      property TIPO_DENEGADA                 : string    read fTIPO_DENEGADA                 write fTIPO_DENEGADA;
      property InfoAdicional                 : String    read fInfoAdicional                 write fInfoAdicional;
      property InfAdFisco                    : String    read fInfAdFisco                    write fInfAdFisco;
      property TpEmiss                       : Integer   read fTpEmiss                       write fTpEmiss;
      property CodCidade                     : String    read fCodCidade                     write fCodCidade;
      property Finalidade                    : TpcnFinalidadeNFe   read fFinalidade          write fFinalidade;
      property Tipo                          : integer   read fTipo                          write fTipo;
      property IE_Destinatario               : String    read fIE_Destinatario               write fIE_Destinatario;
      property Indicador                     : Integer   read fIndicador                     write fIndicador;
      property IndFinal                      : Integer   read fIndFinal                      write fIndFinal;
  end;

implementation

{ TNotaFiscal }

constructor TNotaFiscal.create;
begin
  fListaNotaItem := Tlist.Create;
  fListaNFeHistoricoXML := TList.Create;

  fQryNota := TADOQuery.Create(Self);
  fQryNota.Connection := dmConection.ADOConnection;

  fQryNotaItem := TADOQuery.Create(Self);
  fQryNotaItem.Connection := dmConection.ADOConnection;

  fQryNFeHistorico_XML := TADOQuery.Create(Self);
  fQryNFeHistorico_XML.Connection := dmConection.ADOConnection;

  LimpaCampos;
end;

procedure TNotaFiscal.LimpaCampos;
begin
  fInfoAdicional := '';
  fInfAdFisco    := '';
  fTpEmiss       := 1;
end;

procedure TNotaFiscal.SetTipoEmissao;
begin
  fTpEmiss := IIf(fCONTINGENCIA = '4',9,1);
end;

function TNotaFiscal.PreencherNota(pNota_Fiscal,
                                   pNota_Fiscal_Nfe: string;
                                   pModelo,
                                   pSerie: integer): boolean;
var
  vNotaItem : TNotaItem;
begin
  try
    result := false;

    with fQryNota do begin
      Close;
      sql.Clear;
      sql.Add(' select * ');
      sql.Add('   from NOTAFISCAL N');
      sql.Add('   where N.NumeroNF = :NUMERONF');
      SQL.Add('     and N.MODELO = :modelo ');
      SQL.Add('     and N.SERIE = :serie ');
      Parameters.ParamByName('NUMERONF').Value := pNota_Fiscal;
      Parameters.ParamByName('MODELO').Value := pModelo;
      Parameters.ParamByName('SERIE').Value := pSerie;
      Open;

      if not isEmpty then begin
        fNumeroNF              :=  FieldByName('NumeroNF').AsInteger;
        fSaidaEntrada          :=  FieldByName('SaidaEntrada').Asstring;
        fSerie                 :=  FieldByName('Serie').Asstring;
        fCFOP                  :=  FieldByName('CFOP').Asstring;
        fInscricaoSubstituicao :=  FieldByName('InscricaoSubstituicao').Asstring;
        fInscricaoestadual     :=  TRIM(FieldByName('Inscricaoestadual').Asstring);
        fCodCliente            :=  FieldByName('CodCliente').Asinteger;
        fNome                  :=  FieldByName('Nome').Asstring;
        fCnpjCpf               :=  FieldByName('CnpjCpf').AsString;
        fEmissao               :=  FieldByName('Emissao').AsDateTime;
        fSaida                 :=  FieldByName('Saida').AsDateTime;
        fHoraSaida             :=  TRIM(FieldByName('HoraSaida').Asstring);
        fEndereco              :=  FieldByName('Endereco').AsString;
        fDest_complemento      :=  FieldByName('Dest_complemento').AsString;
        fDest_Numero           :=  FieldByName('Dest_Numero').AsInteger;
        fBairro                :=  FieldByName('Bairro').AsString;
        fCEP                   :=  FieldByName('CEP').AsString;
        fMunicipio             :=  FieldByName('Municipio').Asstring;
        fFoneFax               :=  FieldByName('FoneFax').Asstring;
        fUF                    :=  FieldByName('UF').Asstring;
        fBaseIcms              :=  FieldByName('BaseIcms').AsFloat;
        fIcms                  :=  FieldByName('Icms').AsFloat;
        fBaseSubstituicao      :=  FieldByName('BaseSubstituicao').AsFloat;
        fValorSubstituicao     :=  FieldByName('ValorSubstituicao').AsFloat;
        fTotalProduto          :=  FieldByName('TotalProduto').AsFloat;
        fTotalNota             :=  FieldByName('TotalNota').AsFloat;
        fFrete                 :=  FieldByName('Frete').AsFloat;
        fSeguro                :=  FieldByName('Seguro').AsFloat;
        fOutros                :=  FieldByName('Outros').AsFloat;
        fIPI                   :=  FieldByName('IPI').AsFloat;
        fObservacao            :=  FieldByName('Observacao').Asstring;
        fNaturezaOperacao      :=  FieldByName('NaturezaOperacao').Asstring;
        fVALORDESCONTO         :=  FieldByName('VALORDESCONTO').AsFloat;
        fCancelada             :=  FieldByName('Cancelada').Asstring;
        fModelo                :=  FieldByName('Modelo').Asstring;
        fVALORISENTO           :=  FieldByName('VALORISENTO').AsFloat;
        fVALOROUTRAS           :=  FieldByName('VALOROUTRAS').AsFloat;
        fALIQICMS              :=  FieldByName('ALIQICMS').AsFloat;
        fVERSAO_NFE            :=  FieldByName('VERSAO_NFE').Asstring;
        fJUSTIFICATIVA_CONTINGENCIA     := FieldByName('JUSTIFICATIVA_CONTINGENCIA').Asstring;
        fDATA_HORA_CONTINGENCIA         := FieldByName('DATA_HORA_CONTINGENCIA').AsDateTime;
        fDATA_HORA_RECEB_NFE            := FieldByName('DATA_HORA_RECEB_NFE').Asstring;
        fSTATUS_CTG                     := FieldByName('STATUS_CTG').Asstring;
        fTIPO_AMBIENTE_NFE              := FieldByName('TIPO_AMBIENTE_NFE').Asstring;
        fNUMERO_LOTE_NFE                := FieldByName('NUMERO_LOTE_NFE').Asstring;
        fNUMERO_PROTOCOLO_CANCELAMENTO  := FieldByName('NUMERO_PROTOCOLO_CANCELAMENTO').Asstring;
        fNUMERO_PROTOCOLO               := FieldByName('NUMERO_PROTOCOLO').Asstring;
        fNOTA_FISCAL_NFE                := FieldByName('NOTA_FISCAL_NFE').Asstring;
        fSTATUS_NFE                     := FieldByName('STATUS_NFE').Asstring;
        fNUMERO_RECIBO                  := FieldByName('NUMERO_RECIBO').Asstring;
        fNUMERO_NFE                     := FieldByName('NUMERO_NFE').Asstring;
        fCONTINGENCIA                   := FieldByName('CONTINGENCIA').Asstring;
        fMOTIVO_CANCELAMENTO            := FieldByName('MOTIVO_CANCELAMENTO').Asstring;
        fCOD_SIT_EFD                    := FieldByName('COD_SIT_EFD').Asstring;
        fMENSAGEM_FISCO_ID              := FieldByName('MENSAGEM_FISCO_ID').Asinteger;
        fMENSAGEM_CONTRIBUINTE_ID       := FieldByName('MENSAGEM_CONTRIBUINTE_ID').Asinteger;
        fCONDICAO_TIPO                  := FieldByName('CONDICAO_TIPO').Asstring;
        fCONDICAO_DESCRICAO             := FieldByName('CONDICAO_DESCRICAO').Asstring;
        fTOTAL_OUTRAS_DESP              := FieldByName('TOTAL_OUTRAS_DESP').AsFloat;
        fEMPRESA_ID                     := FieldByName('EMPRESA_ID').Asstring;
        fSTATUS                         := FieldByName('STATUS').Asstring;
        fEMITIU                         := FieldByName('EMITIU').Asstring;
        fMENSAGEM_ID1                   := FieldByName('MENSAGEM_ID1').AsInteger;
        fMENSAGEM_ID2                   := FieldByName('MENSAGEM_ID2').AsInteger;
        fDATA_CANCELAMENTO              := FieldByName('DATA_CANCELAMENTO').AsDateTime;
        fDESC_NAT_OPERACAO              := FieldByName('DESC_NAT_OPERACAO').Asstring;
        fENDERECO_XML                   := FieldByName('ENDERECO_XML').Asstring;
        fMENSAGEM1                      := FieldByName('MENSAGEM1').Asstring;
        fMENSAGEM2                      := FieldByName('MENSAGEM2').Asstring;
        fMENSAGEM3                      := FieldByName('MENSAGEM3').Asstring;
        fMensagemComplementar           := FieldByName('MensagemComplementar').Asstring;
        frefNF                          := FieldByName('refNF').Asstring;
        fTIPO_DENEGADA                  := FieldByName('TIPO_DENEGADA').Asstring;
        result := true;
      end;

      if result then begin
        //TEM QUE TER RETORNO DOS ITENS
        PreencherNotaItem(pNota_Fiscal,pModelo,pSerie);
        preencheNFE_HISTORICO_XML(pNota_Fiscal,pModelo,pSerie);
        MensagemAdicional;
        MensagemFisco;
        SetTipoEmissao;
      end;

    end;
  except
    result := false;
  end;
end;

function TNotaFiscal.preencheNFE_HISTORICO_XML(pNota_fiscal : integer;
                                               pModelo      : String;
                                               pSerie       : String): boolean;
var
  fHistoricoXml : THistoricoXml;
begin
  try
    result := false;

    if fQryNFeHistorico_XML = NIL THEN begin
      fQryNFeHistorico_XML := TADOQuery.Create(Self);
      fQryNFeHistorico_XML.Connection := dmConection.ADOConnection;
    end;

    with fQryNFeHistorico_XML do begin
      Close;
      sql.Clear;
      sql.Add(' select * ');
      sql.Add('   from NFE_HISTORICOXML N');
      sql.Add('  where N.NumeroNF = :NUMERONF');
      SQL.Add('    and N.MODELO   = :modelo ');
      SQL.Add('    and N.SERIE    = :serie ');
      Parameters.ParamByName('NUMERONF').Value := pNota_Fiscal;
      Parameters.ParamByName('MODELO').Value := pModelo;
      Parameters.ParamByName('SERIE').Value := pSerie;
      Open;

      if not isEmpty then begin
        LimpaListaHistoricoXML;
        fQryNFeHistorico_XML.First;
        while not fQryNFeHistorico_XML.Eof do begin
          fHistoricoXml := THistoricoXml.Create;
          fHistoricoXml.ENDERECO_WS      := FieldByName('ENDERECO_WS').AsString; //varchar(200) COLLATE Latin1_General_CI_AS NULL,
          fHistoricoXml.HISTORICO_ID     := FieldByName('HISTORICO_ID').AsInteger; //int IDENTITY(1, 1) NOT NULL,
          fHistoricoXml.EMPRESA_ID       := FieldByName('EMPRESA_ID').AsString; //varchar(18) COLLATE Latin1_General_CI_AS NOT NULL,
          fHistoricoXml.TIPO_MOVIMENTO   := FieldByName('TIPO_MOVIMENTO').AsString; //char(1) COLLATE Latin1_General_CI_AS NULL,
          fHistoricoXml.PC_EMITENTE_NFE  := FieldByName('PC_EMITENTE_NFE').AsString; //varchar(40) COLLATE Latin1_General_CI_AS NULL,
          fHistoricoXml.USUARIO_EMITENTE := FieldByName('USUARIO_EMITENTE').AsString; //varchar(40) COLLATE Latin1_General_CI_AS NULL,
          fHistoricoXml.STATUS_RETORNO   := FieldByName('STATUS_RETORNO').AsInteger; //int NULL,
          fHistoricoXml.DATA_PROCESSO    := FieldByName('DATA_PROCESSO').AsDateTime; //datetime NULL,
          fHistoricoXml.XML_RETORNO      := FieldByName('XML_RETORNO').AsString; //text COLLATE Latin1_General_CI_AS NULL,
          fHistoricoXml.NOTA_FISCAL      := FieldByName('NOTA_FISCAL').AsInteger; //int NOT NULL,
          fHistoricoXml.XML_ENVIADO      := FieldByName('XML_ENVIADO').AsString; //text COLLATE Latin1_General_CI_AS NULL,
          fHistoricoXml.KONT_STATUS      := FieldByName('KONT_STATUS').AsString; //varchar(2) COLLATE Latin1_General_CI_AS NULL,
          fHistoricoXml.MODELO           := FieldByName('MODELO').AsString; //varchar(2) COLLATE Latin1_General_CI_AS DEFAULT '55' NOT NULL,
          fHistoricoXml.SERIE            := FieldByName('SERIE').AsString; //varchar(1) COLLATE Latin1_General_CI_AS DEFAULT '1' NOT NULL,
          fHistoricoXml.ENVIO_RETORNO_SEFAZ := FieldByName('ENVIO_RETORNO_SEFAZ').AsString; //text COLLATE Latin1_General_CI_AS NULL,
          fListaNFeHistoricoXML.Add(fHistoricoXml);
          fQryNFeHistorico_XML.Next;
          Result := true;
        end;
      end;
    end;
  except
    Result := False;
  end;
end;

procedure TNotaFiscal.LimpaListaNotaItem;
var
  i : integer;
begin
  if fListaNotaItem.Count > 0 then
    for I := 0 to flistaNotaItem.Count do
      fListaNotaItem.Delete(i);
end;

procedure TNotaFiscal.LimpaListaHistoricoXML;
var
  i : integer;
begin
  if fListaNFeHistoricoXML.Count > 0 then
    for I := 0 to fListaNFeHistoricoXML.Count do
      fListaNFeHistoricoXML.Delete(i);
end;

procedure TNotaFiscal.MensagemAdicional;
begin
  if fMensagemComplementar <> '' then begin
    if (COPY(fInfoAdicional,Length(fInfoAdicional),1) <> ';') then
      fInfoAdicional := fInfoAdicional + ';';

    fInfoAdicional := fInfoAdicional + fMensagemComplementar;
  end;
end;


procedure TNotaFiscal.MensagemFisco;
begin
  if fMensagem1 <> '' then
   fInfAdFisco := fMensagem1;

  if fMensagem2 <> '' then
   fInfAdFisco := IIf(fInfAdFisco <> '',fInfAdFisco + ';' + fMensagem2,fInfAdFisco + '');
end;


function TNotaFiscal.PreencherNotaItem(pNota_Fiscal: string;
                                       pModelo,
                                       pSerie: integer): boolean;
var
  vNotaItem : TNotaItem;
begin
  try
    result := false;

    with fQryNotaItem do begin
      Close;
      sql.Clear;

      sql.Add('  SELECT NI.Serie,                                   ');
      sql.Add('         NI.NumeroNF,                                ');
      sql.Add('         NI.CodProduto,                              ');
      sql.Add('         NI.Descricao,                               ');
      sql.Add('         NI.CST,                                     ');
      sql.Add('         NI.Unidade,                                 ');
      sql.Add('         ROUND(SUM(NI.Quantidade),2,1) AS QUANTIDADE,');
      sql.Add('         NI.VlrUnitario,                             ');
      sql.Add('         SUM(NI.Vlrtotal) AS VLRTOTAL,               ');
      sql.Add('         NI.ICMS,                                    ');
      sql.Add('         NI.Aliquota,                                ');
      sql.Add('         NI.Modelo,                                  ');
      sql.Add('         NI.CFOP,                                    ');
      sql.Add('         NI.ALIQ_ICMS,                               ');
      sql.Add('         NI.ALIQ_ICMS_CHEIA,                         ');
      sql.Add('         NI.VL_BC_ICMS,                              ');
      sql.Add('         NI.VL_REDUCAO,                              ');
      sql.Add('         NI.VL_OUT_DESP_ITEM,                        ');
      sql.Add('         NI.VL_FRETE_ITEM,                           ');
      sql.Add('         NI.VL_ICMS,                                 ');
      sql.Add('         NI.ALIQ_ICMS_ST,                            ');
      sql.Add('         NI.VL_BC_ICMS_ST,                           ');
      sql.Add('         NI.VL_ICMS_ST,                              ');
      sql.Add('         NI.VL_BC_PIS,                               ');
      sql.Add('         NI.QUANT_BC_PIS,                            ');
      sql.Add('         NI.VL_PIS,                                  ');
      sql.Add('         NI.ALIQ_PIS,                                ');
      sql.Add('         NI.ALIQ_COFINS,                             ');
      sql.Add('         NI.QUANT_BC_COFINS,                         ');
      sql.Add('         NI.VL_COFINS,                               ');
      sql.Add('         NI.ALIQ_ISSQN,                              ');
      sql.Add('         NI.VL_BC_ISSQN,                             ');
      sql.Add('         NI.VL_ISSQN,                                ');
      sql.Add('         NI.VL_ISENTO,                               ');
      sql.Add('         NI.VL_NAO_TRIB,                             ');
      sql.Add('         NI.VL_DESC_ITEM,                            ');
      sql.Add('         NI.VL_BC_COFINS,                            ');
      sql.Add('         NI.CST_PIS,                                 ');
      sql.Add('         NI.CST_COFINS,                              ');
      sql.Add('         NI.PERCENT_IVA,                             ');
      sql.Add('         NI.PERCENT_REDUTOR,                         ');
      sql.Add('         NI.TIPO_ALIQUOTA_PIS_COFINS,                ');
      sql.Add('         NI.PIS_COFINS_CUMULATIVO,                   ');
      sql.Add('         NI.TIPO_ALIQUOTA_ENT_COFINS_IMP,            ');
      sql.Add('         NI.TIPO_ALIQUOTA_ENT_PIS_IMP,               ');
      sql.Add('         NI.NATUREZA_PISCOFINS,                      ');
      sql.Add('         T.CODIGO_NCM,                               ');
      sql.Add('         T.TAB_CODRED                                ');
      sql.Add('    FROM NOTAITEM NI, TABELA T                       ');
      sql.Add('   WHERE T.TAB_CODIGO = NI.CODPRODUTO                ');
      SQL.Add('     AND NI.NumeroNF  = :NUMERONF                    ');
      SQL.Add('     AND NI.MODELO    = :MODELO                      ');
      SQL.Add('     AND NI.SERIE     = :SERIE                       ');
      sql.Add('   GROUP BY NI.Serie,                                ');
      sql.Add('            NI.NumeroNF,                             ');
      sql.Add('            NI.CodProduto,                           ');
      sql.Add('            NI.Descricao,                            ');
      sql.Add('            NI.CST,                                  ');
      sql.Add('            NI.Unidade,                              ');
      sql.Add('            NI.VlrUnitario,                          ');
      sql.Add('            NI.ICMS,                                 ');
      sql.Add('            NI.Aliquota,                             ');
      sql.Add('            NI.Modelo,                               ');
      sql.Add('            NI.CFOP,                                 ');
      sql.Add('            NI.ALIQ_ICMS,                            ');
      sql.Add('            NI.ALIQ_ICMS_CHEIA,                      ');
      sql.Add('            NI.VL_BC_ICMS,                           ');
      sql.Add('            NI.VL_REDUCAO,                           ');
      sql.Add('            NI.VL_OUT_DESP_ITEM,                     ');
      sql.Add('            NI.VL_FRETE_ITEM,                        ');
      sql.Add('            NI.VL_ICMS,                              ');
      sql.Add('            NI.ALIQ_ICMS_ST,                         ');
      sql.Add('            NI.VL_BC_ICMS_ST,                        ');
      sql.Add('            NI.VL_ICMS_ST,                           ');
      sql.Add('            NI.VL_BC_PIS,                            ');
      sql.Add('            NI.QUANT_BC_PIS,                         ');
      sql.Add('            NI.VL_PIS,                               ');
      sql.Add('            NI.ALIQ_PIS,                             ');
      sql.Add('            NI.ALIQ_COFINS,                          ');
      sql.Add('            NI.QUANT_BC_COFINS,                      ');
      sql.Add('            NI.VL_COFINS,                            ');
      sql.Add('            NI.ALIQ_ISSQN,                           ');
      sql.Add('            NI.VL_BC_ISSQN,                          ');
      sql.Add('            NI.VL_ISSQN,                             ');
      sql.Add('            NI.VL_ISENTO,                            ');
      sql.Add('            NI.VL_NAO_TRIB,                          ');
      sql.Add('            NI.VL_DESC_ITEM,                         ');
      sql.Add('            NI.VL_BC_COFINS,                         ');
      sql.Add('            NI.CST_PIS,                              ');
      sql.Add('            NI.CST_COFINS,                           ');
      sql.Add('            NI.PERCENT_IVA,                          ');
      sql.Add('            NI.PERCENT_REDUTOR,                      ');
      sql.Add('            NI.TIPO_ALIQUOTA_PIS_COFINS,             ');
      sql.Add('            NI.PIS_COFINS_CUMULATIVO,                ');
      sql.Add('            NI.TIPO_ALIQUOTA_ENT_COFINS_IMP,         ');
      sql.Add('            NI.TIPO_ALIQUOTA_ENT_PIS_IMP,            ');
      sql.Add('            NI.NATUREZA_PISCOFINS,                   ');
      sql.Add('            T.CODIGO_NCM,                            ');
      sql.Add('            T.TAB_CODRED                             ');
      Parameters.ParamByName('NUMERONF').Value := pNota_Fiscal;
      Parameters.ParamByName('MODELO').Value := pModelo;
      Parameters.ParamByName('SERIE').Value := pSerie;
      Open;

      if not isEmpty then begin
        LimpaListaNotaItem;
        first;
        while not eof do begin
          vNotaItem := TNotaitem.Create;
          vNotaItem.Serie                        := fieldByName('Serie').asString;
          vNotaItem.NumeroNF                     := fieldByName('NumeroNF').asInteger;
          vNotaItem.CodProduto                   := fieldByName('CodProduto').asString;
          vNotaItem.Descricao                    := fieldByName('Descricao').asString;
          vNotaItem.CST                          := fieldByName('CST').asString;
          vNotaItem.Unidade                      := fieldByName('Unidade').asString;
          vNotaItem.Quantidade                   := fieldByName('Quantidade').AsFloat;
          vNotaItem.VlrUnitario                  := fieldByName('VlrUnitario').AsFloat;
          vNotaItem.Vlrtotal                     := fieldByName('Vlrtotal').AsFloat;
          vNotaItem.ICMS                         := fieldByName('ICMS').AsFloat;
          vNotaItem.Aliquota                     := fieldByName('Aliquota').asString;
          vNotaItem.Modelo                       := fieldByName('Modelo').AsString;
          vNotaItem.CFOP                         := fieldByName('CFOP').AsString;
          vNotaItem.ALIQ_ICMS                    := fieldByName('ALIQ_ICMS').AsFloat;
          vNotaItem.ALIQ_ICMS_CHEIA              := fieldByName('ALIQ_ICMS_CHEIA').AsFloat;
          vNotaItem.VL_BC_ICMS                   := fieldByName('VL_BC_ICMS').AsFloat;
          vNotaItem.VL_REDUCAO                   := fieldByName('VL_REDUCAO').AsFloat;
          vNotaItem.VL_OUT_DESP_ITEM             := fieldByName('VL_OUT_DESP_ITEM').AsFloat;
          vNotaItem.VL_FRETE_ITEM                := fieldByName('VL_FRETE_ITEM').AsFloat;
          vNotaItem.VL_ICMS                      := fieldByName('VL_ICMS').AsFloat;
          vNotaItem.ALIQ_ICMS_ST                 := fieldByName('ALIQ_ICMS_ST').AsFloat;
          vNotaItem.VL_BC_ICMS_ST                := fieldByName('VL_BC_ICMS_ST').AsFloat;
          vNotaItem.VL_ICMS_ST                   := fieldByName('VL_ICMS_ST').AsFloat;
          vNotaItem.VL_BC_PIS                    := fieldByName('VL_BC_PIS').AsFloat;
          vNotaItem.QUANT_BC_PIS                 := fieldByName('QUANT_BC_PIS').AsFloat;
          vNotaItem.VL_PIS                       := fieldByName('VL_PIS').AsFloat;
          vNotaItem.ALIQ_PIS                     := fieldByName('ALIQ_PIS').AsFloat;
          vNotaItem.ALIQ_COFINS                  := fieldByName('ALIQ_COFINS').AsFloat;
          vNotaItem.QUANT_BC_COFINS              := fieldByName('QUANT_BC_COFINS').AsFloat;
          vNotaItem.VL_COFINS                    := fieldByName('VL_COFINS').AsFloat;
          vNotaItem.ALIQ_ISSQN                   := fieldByName('ALIQ_ISSQN').AsFloat;
          vNotaItem.VL_BC_ISSQN                  := fieldByName('VL_BC_ISSQN').AsFloat;
          vNotaItem.VL_ISSQN                     := fieldByName('VL_ISSQN').AsFloat;
          vNotaItem.VL_ISENTO                    := fieldByName('VL_ISENTO').AsFloat;
          vNotaItem.VL_NAO_TRIB                  := fieldByName('VL_NAO_TRIB').AsFloat;
          vNotaItem.VL_DESC_ITEM                 := fieldByName('VL_DESC_ITEM').AsFloat;
          vNotaItem.VL_BC_COFINS                 := fieldByName('VL_BC_COFINS').AsFloat;
          vNotaItem.CST_PIS                      := fieldByName('CST_PIS').AsString;
          vNotaItem.CST_COFINS                   := fieldByName('CST_COFINS').AsFloat;
          vNotaItem.PERCENT_IVA                  := fieldByName('PERCENT_IVA').AsFloat;
          vNotaItem.PERCENT_REDUTOR              := fieldByName('PERCENT_REDUTOR').AsFloat;
          vNotaItem.TIPO_ALIQUOTA_PIS_COFINS     := fieldByName('TIPO_ALIQUOTA_PIS_COFINS').AsString;
          vNotaItem.PIS_COFINS_CUMULATIVO        := fieldByName('PIS_COFINS_CUMULATIVO').AsString;
          vNotaItem.TIPO_ALIQUOTA_ENT_COFINS_IMP := fieldByName('TIPO_ALIQUOTA_ENT_COFINS_IMP').AsString;
          vNotaItem.TIPO_ALIQUOTA_ENT_PIS_IMP    := fieldByName('TIPO_ALIQUOTA_ENT_PIS_IMP').AsString;
          vNotaItem.NATUREZA_PISCOFINS           := fieldByName('NATUREZA_PISCOFINS').AsString;
          vNotaItem.CODIGO_NCM                   := fieldByName('CODIGO_NCM').AsString;
          vNotaitem.TAB_CODRED                   := fieldByName('TAB_CODRED').AsString;
          fListaNotaItem.Add(vNotaItem);
          next;
        end;
      end;
    end;
  finally

  end;
end;

end.





































