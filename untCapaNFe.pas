unit untCapaNFe;

interface

uses
  Classes, Generics.Collections, untItensNfe, SysUtils, Dialogs, untFuncoes,
  untPagNfe, ADODB, UnitDM, untRefEcf, untRefNFCe, untConstante;

type
  TCapaNFe = class(TObject)
    private
       // Itens
      fListaItens       : TObjectList<TItensNFe>;
      fListaPag         : TObjectList<TPagNFe>;

      //Lista referente a Cupom
      fListaRefECF      : TObjectList<TRefEcf>;
      fListaRefNFCe     : TObjectList<TRefNFCe>;

      frefNF            : String; //Chave de acesso da NF-e referenciada

      fDataCaixa        : TDateTime;
      fIDNotaFiscal     : Integer;

      // Identificacao
      fVersao           : string;
      fNaturezaOperacao : string;
      fModelo           : Integer; // 65 NFC-e ou 55 NFe
      fCodigo           : Integer;
      fNumero           : Integer;
      fSerie            : Integer;
      fEmissao          : String;//TDateTime;
      fSaida            : String;//TDateTime;
      fHoraSaida        : string;
      fTipo             : Integer;
      fFormaPag         : String;   //0=Pagamento à vista;
                                    //1=Pagamento a prazo;
                                    //2=Outros.


      fdhCont           : String; //data e hora em contingencia
      fxJust            : String; //Justificativa em contingencia

      fFinalidade       : Integer;  //1=NF-e normal;
                                    //2=NF-e complementar;
                                    //3=NF-e de ajuste;
                                    //4=Devolução de mercadoria.

      fidDest           : Integer; //Identificador de local de destino da operação
                                   //1=Operação interna;
                                   //2=Operação interestadual;
                                   //3=Operação com exterior
      findFinal         : Integer; //Indica operação com Consumidor final
                                   //0=Não;
                                   //1=Consumidor final;
      findPres          : Integer; //Indicador de presença do comprador no estabelecimento comercial nomomento da operação
                                   //0=Não se aplica (por exemplo, Nota Fiscal complementar ou de ajuste);
                                   //1=Operação presencial;
                                   //2=Operação não presencial, pela Internet;
                                   //3=Operação não presencial, Teleatendimento;
                                   //4=NFC-e em operação com entrega a domicílio;
                                   //9=Operação não presencial, outros.

    ftpEmis             : Integer; //1=Emissão normal (não em contingência);
                                   //2=Contingência FS-IA, com impressão do DANFE em formulário de segurança;
                                   //3=Contingência SCAN (Sistema de Contingência do Ambiente Nacional);
                                   //4=Contingência DPEC (Declaração Prévia da Emissão em Contingência);
                                   //5=Contingência FS-DA, com impressão do DANFE em formulário de segurança;
                                   //6=Contingência SVC-AN (SEFAZ Virtual de Contingência do AN);
                                   //7=Contingência SVC-RS (SEFAZ Virtual de Contingência do RS);
                                   //9=Contingência off-line da NFC-e;
                                   //Observação: Para a NFC-e somente estão disponíveis e são válidas as opções de contingência 5 e 9.


     ftpImp             : Integer; // 0=Sem geração de DANFE;
                                   // 1=DANFE normal, Retrato;
                                   // 2=DANFE normal, Paisagem;
                                   // 3=DANFE Simplificado;
                                   // 4=DANFE NFC-e;
                                   // 5=DANFE NFC-e em mensagem eletrônica (o envio de
                                   // mensagem eletrônica pode ser feita de forma simultânea
                                   // com a impressão do DANFE; usar o tpImp=5 quando
                                   // esta for a única forma de disponibilização do DANFE).

      // Emitente
      fEmitenteCNPJ             : string;
      fEmitenteIE               : string;
      fEmitenteIM               : string;
      fEmitenteCNAE             : string;
      fEmitenteRazao            : String;
      fEmitenteFantasia         : string;
      fEmitenteFone             : String;
      fEmitenteCEP              : String;
      fEmitenteLogradouro       : String;
      fEmitenteNumero           : String;
      fEmitenteComplemento      : String;
      fEmitenteBairro           : String;
      fEmitenteCidadeCod        : String;
      fEmitenteCidade           : String;
      fEmitenteUF               : String;
      fEmitentePaisCod          : String;
      fEmitentePais             : String;
      fEmitenteCRT              : String; //Regime Tributário

      //Destinatario
      fDestinatarioCodigo       : Integer;
      fDestinatarioCNPJ         : String;
      fDestinatarioIE           : String;
      fDestinatarioISUF         : String; //*
      fDestinatarioNomeRazao    : String;
      fDestinatarioFone         : String;
      fDestinatarioCEP          : String;
      fDestinatarioLogradouro   : String;
      fDestinatarioNumero       : String;
      fDestinatarioComplemento  : String;
      fDestinatarioBairro       : String;
      fDestinatarioCidadeCod    : String;
      fDestinatarioCidade       : String;
      fDestinatarioUF           : String;
      fDestinatarioPaisCod      : String; //*
      fDestinatarioPais         : String; //*
      fDestinatarioindIEDest    : String; // Indicador da IE do Destinatário
                                          //1=Contribuinte ICMS (informar a IE dodestinatário);
                                          //2=Contribuinte isento de Inscrição no cadastro de
                                          //Contribuintes do ICMS;
                                          //9=Não Contribuinte, que pode ou não possuir Inscrição
                                          //Estadual no Cadastro de Contribuintes do ICMS;
                                          //Nota 1: No caso de NFC-e informar indIEDest=9 e não
                                          //informar a tag IE do destinatário;
                                          //Nota 2: No caso de operação com o Exterior informar
                                          //indIEDest=9 e não informar a tag IE do destinatário;
                                          //Nota 3: No caso de Contribuinte Isento de Inscrição
                                          //(indIEDest=2), não informar a tag IE do destinatário.


      //Transportador
      fTransportadorFretePorConta : String; //*
      fTransportadorCnpjCpf       : String; //*
      fTransportadorNomeRazao     : String; //*
      fTransportadorIE            : String; //*
      fTransportadorEndereco      : String; //*
      fTransportadorCidade        : String; //*
      fTransportadorUF            : String; //*
      fTransportadorValorServico  : String; //*
      fTransportadorValorBase     : String; //*
      fTransportadorAliquota      : String; //*
      fTransportadorValor         : String; //*
      fTransportadorCFOP          : String; //*
      fTransportadorCidadeCod     : String; //*
      fTransportadorPlaca         : String; //*
      fTransportadorUFPlaca       : String; //*
      fTransportadorRNTC          : String; //*

      //Volume
      fVolumeQuantidade           : String; //*
      fVolumeEspecie              : String; //*
      fVolumeMarca                : String; //*
      fVolumeNumeracao            : String; //*
      fVolumePesoLiquido          : String; //*
      fVolumePesoBruto            : String; //*

      //Fatura
      fFaturaNumero               : String; //*
      fFaturaValorOriginal        : String; //*
      fFaturaValorDesconto        : String; //*
      fFaturaValorLiquido         : String; //*

      //Duplicata
      fDuplicataNumero            : String; //*
      fDuplicataDataVencimento    : String; //*
      fDuplicataValor             : String; //*

      //DadosAdicionais
      fDadosAdicionaisComplemento : String; //*

      //InfAdic
      finfCpl                     : AnsiString;
      fInfAdicCampo               : String; //*
      fInfAdicTexto               : String; //*
      fInfAdFisco                 : AnsiString;

      //autXMLxxx
      fCNPJautXML                 : string;
      fCPFautXML                  : string;

      //Total
      fTotalBaseICMS               : Double;
      fTotalValorICMS              : Double;
      fTotalValorProduto           : Double;
      fTotalValorNota              : Double;
      fValorBaseICMS_ST            : Double;
      fValorICMS_ST                : Double;
//      fTotalValorFrete             : string; // * - este campo é opcional
//      fTotalValorSeguro            : string; // * - este campo é opcional
//      fTotalValorDesconto          : string; // * - este campo é opcional
//      fTotalValorII                : string; // * - este campo é opcional
//      fTotalValorIPI               : string; // * - este campo é opcional
//      fTotalValorPIS               : string; // * - este campo é opcional
//      fTotalValorCOFINS            : string; // * - este campo é opcional
//      fTotalValorOutrasDespesas    : string; // * - este campo é opcional

      fTotTrib    : String; //vTotTrib Valor aproximado total de tributos federais, estaduais e municipais.

      fMensagem3  : string;
      fCFOP       : string;
      fStatus_Ctg : string;
      fStatus_Nfe : String;

      fTemItemCombustivel : Boolean;

      function AddCapa(pCapa: array of Variant): boolean;
      function CalculaTotaisNfe: string;
      procedure SetCFOP(value: string = '');
      function BuscaCuponsReferenciados(var pIndex: integer): string;
      function BuscaNFCeReferenciadas(pIndex: integer): string;
    function BuscaCest(pNCM: string): STRING;
    function CFOPCombustivel(pCfop: string): boolean;
    function ValidaICMS(pCSOSN, pCST, pCFOP, pCODIGO,
      pDescricao: STRING): boolean;
    public
      constructor create(pCapa: array of Variant; var retorno: Boolean); overload;
      constructor create; overload;
      procedure SetStatus(value: string);
      function AddItens(pItens: array of Variant): Boolean;
      function AddFormaPagamento(pFpgto: array of Variant): Boolean;
      function AddCupomRef(pCupomRef: array of Variant): boolean;
      function AddNFCeRef(pNFCeRef: string): boolean;
      function ExisteNFCeRef(pNFCeRef: String): boolean;
      destructor Destroy;
      function GetTextoIni: string;
      function ValidaNota: boolean;
      function GravaNFe(pNFCe: Boolean = False): Boolean;
      function ApagaNFe: Boolean;
      function GravaItensNfe(pNFCe : Boolean = False): Boolean;
      procedure SetTotalValorProduto(pValue : Double);
      procedure SetTotalValorNota(pValue : Double);
      function SetTpEmiss(pTpEmiss: Integer): Boolean;
      function SetdhCont: Boolean;
      function SetxJust(pJust: String): Boolean;
      procedure SetDestinatario(pDestinatarioCNPJ        : string;
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

    function GetCFOP_ItensNFe(pCodigo : string; pItem: integer):string;
    function GetValorBaseICMS_ItensNFe(pCodigo : string; pItem: integer):Double;
    function GetAliquotaICMS_ItensNFe(pCodigo : string; pItem: integer):Double;
    function GetValorICMS_ItensNFe(pCodigo : string; pItem: integer):Double;


    published
      property Versao           : string    read fVersao             write fVersao;
      property NaturezaOperacao  : string    read fNaturezaOperacao   write fNaturezaOperacao;
      property refNF            : String    read frefNF              write frefNF;
      property Modelo          : Integer   read fModelo             write fModelo;
      property Codigo         : Integer   read fCodigo             write fCodigo;
      property Numero        : Integer   read fNumero             write fNumero;
      property Serie        : Integer   read fSerie              write fSerie;
      property Emissao     : String    read fEmissao            write fEmissao;
      property Saida      : String    read fSaida              write fSaida;
      property HoraSaida : string    read fHoraSaida          write fHoraSaida;
      property Tipo       : Integer   read fTipo               write fTipo;
      property FormaPag    : String    read fFormaPag           write fFormaPag;
      property Finalidade   : Integer   read fFinalidade         write fFinalidade;
      property idDest        : Integer   read fidDest             write fidDest;
      property indFinal       : Integer   read findFinal           write findFinal;
      property indPres         : Integer   read findPres            write findPres;
      property EmitenteCNPJ     : string    read fEmitenteCNPJ       write fEmitenteCNPJ;
      property EmitenteIE        : string    read fEmitenteIE         write fEmitenteIE;
      property EmitenteIM         : String    read fEmitenteIM         write fEmitenteIM;
      property EmitenteCNAE        : String    read fEmitenteCNAE       write fEmitenteCNAE;
      property EmitenteRazao        : String    read fEmitenteRazao      write fEmitenteRazao;
      property EmitenteFantasia      : string    read fEmitenteFantasia   write fEmitenteFantasia;
      property EmitenteFone           : String    read fEmitenteFone       write fEmitenteFone;
      property EmitenteCEP             : String    read fEmitenteCEP        write fEmitenteCEP;
      property EmitenteLogradouro       : String    read fEmitenteLogradouro write fEmitenteLogradouro;
      property EmitenteNumero          : String  read fEmitenteNumero       write fEmitenteNumero;
      property EmitenteComplemento    : String  read fEmitenteComplemento  write fEmitenteComplemento;
      property EmitenteBairro        : String  read fEmitenteBairro       write fEmitenteBairro;
      property EmitenteCidadeCod    : String  read fEmitenteCidadeCod    write fEmitenteCidadeCod;
      property EmitenteCidade      : String  read fEmitenteCidade       write fEmitenteCidade;
      property EmitenteUF        : String   read fEmitenteUF           write fEmitenteUF;
      property EmitentePaisCod  : String   read fEmitentePaisCod      write fEmitentePaisCod;
      property EmitentePais     : String read fEmitentePais           write fEmitentePais;
      property DestinatarioCNPJ  : String read fDestinatarioCNPJ       write fDestinatarioCNPJ;
      property DestinatarioCodigo : Integer read fDestinatarioCodigo     write fDestinatarioCodigo;
      property DestinatarioIE      : String read fDestinatarioIE         write fDestinatarioIE;
      property DestinatarioISUF     : String read fDestinatarioISUF       write fDestinatarioISUF;
      property DestinatarioNomeRazao : String read fDestinatarioNomeRazao  write fDestinatarioNomeRazao;
      property DestinatarioFone       : String read fDestinatarioFone       write fDestinatarioFone;
      property DestinatarioCEP         : String read fDestinatarioCEP        write fDestinatarioCEP;
      property DestinatarioLogradouro   : String read fDestinatarioLogradouro write fDestinatarioLogradouro;
      property DestinatarioNumero        : String read fDestinatarioNumero     write fDestinatarioNumero;
      property DestinatarioComplemento  : String read fDestinatarioComplemento  write fDestinatarioComplemento;
      property DestinatarioBairro      : String read fDestinatarioBairro       write fDestinatarioBairro;
      property DestinatarioCidadeCod  : String read fDestinatarioCidadeCod    write fDestinatarioCidadeCod;
      property DestinatarioCidade    : String read fDestinatarioCidade         write fDestinatarioCidade;
      property DestinatarioUF         : String read fDestinatarioUF             write fDestinatarioUF;
      property DestinatarioPaisCod     : String read fDestinatarioPaisCod        write fDestinatarioPaisCod;
      property DestinatarioPais         : String read fDestinatarioPais           write fDestinatarioPais;
      property DestinatarioindIEDest     : String read fDestinatarioindIEDest      write fDestinatarioindIEDest;
      property TransportadorFretePorConta : String read fTransportadorFretePorConta write fTransportadorFretePorConta;
      property TransportadorCnpjCpf      : String read fTransportadorCnpjCpf       write fTransportadorCnpjCpf;
      property TransportadorNomeRazao   : String read fTransportadorNomeRazao     write fTransportadorNomeRazao;
      property TransportadorIE         : String read fTransportadorIE            write fTransportadorIE;
      property TransportadorEndereco  : String read fTransportadorEndereco      write fTransportadorEndereco;
      property TransportadorCidade     : String read fTransportadorCidade        write fTransportadorCidade;
      property TransportadorUF          : String read fTransportadorUF            write fTransportadorUF;
      property TransportadorValorServico : String read fTransportadorValorServico  write fTransportadorValorServico;
      property TransportadorValorBase     : String read fTransportadorValorBase   write fTransportadorValorBase;
      property TransportadorAliquota     : String read fTransportadorAliquota    write fTransportadorAliquota;
      property TransportadorValor       : String read fTransportadorValor       write fTransportadorValor;
      property TransportadorCFOP       : String read fTransportadorCFOP        write fTransportadorCFOP;
      property TransportadorCidadeCod : String read fTransportadorCidadeCod   write fTransportadorCidadeCod;
      property TransportadorPlaca    : String read fTransportadorPlaca       write fTransportadorPlaca;
      property TransportadorUFPlaca : String read fTransportadorUFPlaca     write fTransportadorUFPlaca;
      property TransportadorRNTC   : String read fTransportadorRNTC        write fTransportadorRNTC;
      property VolumeQuantidade   : String read fVolumeQuantidade         write fVolumeQuantidade;
      property VolumeEspecie     : String read fVolumeEspecie            write fVolumeEspecie;
      property VolumeMarca      : String read fVolumeMarca                write fVolumeMarca;
      property VolumeNumeracao   : String read fVolumeNumeracao            write fVolumeNumeracao;
      property VolumePesoLiquido  : String read fVolumePesoLiquido          write fVolumePesoLiquido;
      property VolumePesoBruto     : String read fVolumePesoBruto            write fVolumePesoBruto;
      property FaturaNumero         : String read fFaturaNumero               write fFaturaNumero;
      property FaturaValorOriginal   : String read fFaturaValorOriginal        write fFaturaValorOriginal;
      property FaturaValorDesconto    : String read fFaturaValorDesconto        write fFaturaValorDesconto;
      property FaturaValorLiquido      : String read fFaturaValorLiquido         write fFaturaValorLiquido;
      property DuplicataNumero          : String read fDuplicataNumero            write fDuplicataNumero;
      property DuplicataDataVencimento   : String read fDuplicataDataVencimento    write fDuplicataDataVencimento;
      property DuplicataValor             : String read fDuplicataValor             write fDuplicataValor;
      property DadosAdicionaisComplemento : String read fDadosAdicionaisComplemento  write fDadosAdicionaisComplemento;
      property InfAdicCampo              : String read fInfAdicCampo                write fInfAdicCampo;
      property InfAdicTexto             : String read fInfAdicTexto                write fInfAdicTexto;
      property CNPJautXML              : string read fCNPJautXML                  write fCNPJautXML;
      property CPFautXML              : string read fCPFautXML                   write fCPFautXML;
      property Mensagem3             : string read fMensagem3                   write fMensagem3;
      property CFOP                 : string read fCFOP                        write SetCFOP;
      property Status_Ctg         : string read fStatus_Ctg                   write SetStatus;
      property Status_Nfe        : String read fStatus_Nfe                   write SetStatus;
      property TotTrib          : String read fTotTrib                      write fTotTrib;
      property tpImp           : Integer  read ftpImp                      write ftpImp;
      property infCpl         : AnsiString read finfCpl                   write finfCpl;
      property InfAdFisco    : AnsiString  read fInfAdFisco              write fInfAdFisco;
      property tpEmis       : Integer      read ftpEmis                 write ftpEmis;
      property EmitenteCRT  : String       read fEmitenteCRT           write fEmitenteCRT;
      property DataCaixa    : TDateTime    read fDataCaixa            write fDataCaixa;
      property IDNotaFiscal : Integer      read fIDNotaFiscal        write fIDNotaFiscal;


  end;

implementation

{ TCapaNFe }

function TCapaNFe.AddItens(pItens: array of Variant): boolean;
VAR
  vItensNfe : TItensNFe;
begin
  try
    vItensNfe := TItensNFe.Create(pItens);
    flistaItens.Add(vItensNfe);
    Result := True;
  except
    on e: Exception do begin
      Result := False;
      ShowMessage(e.Message + #13#10 + 'Não foi possível informar os itens da Nfc-e, entre em contato com o suporte da Kontrol!');
    end;
  end;
end;

//adiciona na classe de cupom referenciado
function TCapaNFe.AddCupomRef(pCupomRef: array of Variant): boolean;
VAR
  vRefEcf : TRefEcf;
begin
  try
    vRefEcf := TRefEcf.Create(pCupomRef);
    fListaRefECF.Add(vRefEcf);
    Result := True;
  except
    on e: Exception do begin
      Result := False;
      ShowMessage(e.Message + #13#10 + 'Não foi possível informar os cupons Referenciado!, entre em contato com o suporte da Kontrol!');
    end;
  end;
end;

function TCapaNFe.ExisteNFCeRef(pNFCeRef: String): boolean;
var
  fRefNFCe : TRefNFCe;
begin
  result := false;
  for fRefNFCe in fListaRefNFCE  do begin
    if frefNFCE.chave = pNFCeRef then begin
      result := true;
      break;
    end;
  end;
end;

//adiciona na classe de NFCe referenciado
function TCapaNFe.AddNFCeRef(pNFCeRef: string): boolean;
VAR
  vRefNFCe : TRefNFCe;
begin
  try
    vRefNFCe := TRefNFCe.Create(pNFCeRef);
    fListaRefNFCe.Add(vRefNFCe);
    Result := True;
  except
    on e: Exception do begin
      Result := False;
      ShowMessage(e.Message + #13#10 + 'Não foi possível informar os cupons Referenciado!, entre em contato com o suporte da Kontrol!');
    end;
  end;
end;

function TCapaNFe.AddFormaPagamento(pFpgto: array of Variant): Boolean;
var
  vFormaPagNfe : TPagNFe;
begin
  try
    vFormaPagNfe := TPagNFe.Create(pFpgto);
    fListaPag.Add(vFormaPagNfe);
    Result := True;
  except
    on e: Exception do begin
      Result := False;
      ShowMessage(e.Message + #13#10 + 'Não foi possível informar a forma de Pagamento da Nfc-e, entre em contato com o suporte da Kontrol!');
    end;
  end;
end;

constructor TCapaNFe.create(pCapa: array of Variant; var retorno: Boolean);
begin
  fVersao := '3.10';
  fListaItens  := TObjectList<TItensNFe>.Create;
  fListaPag    := TObjectList<TPagNFe>.Create;
  fListaRefECF := TObjectList<TRefEcf>.Create;
  fListaRefNFCe := TObjectList<TRefNFCe>.create;
  fTemItemCombustivel := false;
  retorno := AddCapa(pCapa);
end;

function TCapaNFe.SetTpEmiss(pTpEmiss: Integer): Boolean;
begin
  try
    ftpEmis := pTpEmiss;
    Result := true;
  except
    Result := false;
  end;
end;

function TCapaNFe.SetdhCont: Boolean;
begin
  try
    fdhCont := fEmissao;
    Result := true;
  except
    Result := false;
  end;
end;

function TCapaNFe.SetxJust(pJust: String): Boolean;
begin
  try
    fxJust := pJust;
    Result := true;
  except
    Result := false;
  end;
end;

function TCapaNFe.AddCapa(pCapa: array of Variant): boolean;
var
  i : integer;
begin
  try
    i := 0;
    fNaturezaOperacao        := pCapa[i]; inc(i);        //string       0
    fModelo                  := pCapa[i]; inc(i);        //Integer      1
    fCodigo                  := pCapa[i]; inc(i);        //Integer      2
    fNumero                  := pCapa[i]; inc(i);        //Integer      3
    fSerie                   := pCapa[i]; inc(i);        //Integer      4
    fEmissao                 := pCapa[i]; inc(i);        //String       5
    fSaida                   := pCapa[i]; inc(i);        //String       6
    fTipo                    := pCapa[i]; inc(i);        //Integer      7
    fFormaPag                := pCapa[i]; inc(i);        //String       8
    fFinalidade              := pCapa[i]; inc(i);        //Integer      9
    fidDest                  := pCapa[i]; inc(i);       //Integer       10
    findFinal                := pCapa[i]; inc(i);       //Integer       11
    findPres                 := pCapa[i]; inc(i);       //Integer       12
    fEmitenteCNPJ            := pCapa[i]; inc(i);       //string        13
    fEmitenteIE              := pCapa[i]; inc(i);       //string        15
    fEmitenteIM              := pCapa[i]; inc(i);       //string        14
    fEmitenteCNAE            := pCapa[i]; inc(i);       //string        16
    fEmitenteRazao           := pCapa[i]; inc(i);       //string        17
    fEmitenteFantasia        := pCapa[i]; inc(i);       //string        18
    fEmitenteFone            := pCapa[i]; inc(i);       //string        19
    fEmitenteCEP             := pCapa[i]; inc(i);       //string        20
    fEmitenteLogradouro      := pCapa[i]; inc(i);      //string        21
    fEmitenteNumero          := pCapa[i]; inc(i);     //string        22
    fEmitenteComplemento     := pCapa[i]; inc(i);     //string        23
    fEmitenteBairro          := pCapa[i]; inc(i);     //string        24
    fEmitenteCidadeCod       := pCapa[i]; inc(i);     //string        25
    fEmitenteCidade          := pCapa[i]; inc(i);     //string        26
    fEmitenteUF              := pCapa[i]; inc(i);     //string        27
    fEmitentePaisCod         := pCapa[i]; inc(i);     //string        28
    fEmitentePais            := pCapa[i]; inc(i);     //string        29
    fDestinatarioCNPJ        := pCapa[i]; inc(i);     //string        30
    fDestinatarioIE          := pCapa[i]; inc(i); //string        31
    fDestinatarioNomeRazao   := pCapa[i]; inc(i); //string        32
    fDestinatarioFone        := pCapa[i]; inc(i); //string        33
    fDestinatarioCEP         := pCapa[i]; inc(i); //string        34
    fDestinatarioLogradouro  := pCapa[i]; inc(i); //string        35
    fDestinatarioNumero      := pCapa[i]; inc(i); //string        36
    fDestinatarioComplemento := pCapa[i]; inc(i); //string        37
    fDestinatarioBairro      := pCapa[i]; inc(i); //string        38
    fDestinatarioCidadeCod   := pCapa[i]; inc(i); //string        39
    fDestinatarioCidade      := pCapa[i]; inc(i); //string        40
    fDestinatarioUF          := pCapa[i]; inc(i); //string        41
    fDestinatarioindIEDest   := pCapa[i]; inc(i); //string        42
    fCNPJautXML              := pCapa[i]; inc(i); //string        43
    fCPFautXML               := pCapa[i]; inc(i); //string        44
    fDestinatarioCodigo      := pCapa[i]; inc(i); //Integer;      45
    ftpImp                   := pCapa[i]; inc(i); //Integer;      46
    finfCpl                  := pCapa[i]; inc(i); //AnsiString;   47
    SetCFOP(pCapa[i]); inc(i);                    //              48
    frefNF                   := pCapa[i]; Inc(i); //String;       49
    fInfAdFisco              := pCapa[i]; Inc(i); //String;       50
    ftpEmis                  := pCapa[i]; Inc(i); //String;       51
    fEmitenteCRT             := pCapa[i]; Inc(i); //String;       52
    fHoraSaida               := pCapa[i]; Inc(i); //String;       53
    if trim(fHoraSaida) = '' then
      fHoraSaida := TimeToStr(Time);

    if fModelo = 65 then
      fTransportadorFretePorConta := '9';
//      0=Contratação do Frete por conta do Remetente (CIF);
//      1=Contratação do Frete por conta do Destinatário (FOB);
//      2=Contratação do Frete por conta de Terceiros;
//      3=Transporte Próprio por conta do Remetente;
//      4=Transporte Próprio por conta do Destinatário;
//      9=Sem Ocorrência de Transporte.

    if fModelo = 65 then begin
      fDataCaixa    := pCapa[i]; Inc(i); //Datetime 54
      fIDNotaFiscal := pCapa[i]; Inc(i); //Integer  55
    end;

    SetStatus('P');

    result := True;
  except
    result := False;
  end;
end;

destructor TCapaNFe.destroy;
begin
  fListaItens.Free;
  fListaRefECF.Free;
  fListaRefECF.Free;
  fTemItemCombustivel := false;
end;

function TCapaNFe.GetAliquotaICMS_ItensNFe(pCodigo: string;
  pItem: integer): Double;
var
  I : INTEGER;
begin
  try
    for I := 0 TO fListaItens.Count - 1  do begin
      IF ((fListaItens.Items[I] AS TItensNFe).Codigo = pCodigo) AND (I = pItem) Then begin
        Result := (fListaItens.Items[I] AS TItensNFe).AliquotaICMS;
      end;
    end;
  except
    result := 0;
  end;
end;

function TCapaNFe.GetCFOP_ItensNFe(pCodigo: string; pItem: integer): string;
var
  I : INTEGER;
begin
  try
    for I := 0 TO fListaItens.Count - 1  do begin
      IF ((fListaItens.Items[I] AS TItensNFe).Codigo = pCodigo) AND (I = pItem) Then begin
        Result := (fListaItens.Items[I] AS TItensNFe).CFOP;
      end;
    end;
  except
    result := '';
  end;
end;

function TCapaNFe.GetTextoIni: string;
var
  fItens     : TItensNFe;
  fFormaPgto : TPagNFe;
  vIndex     : Integer;
  vBand      : String;
  vTotalProduto   : String;
  vTotalValorNota : String;
  vTime           : TTime;
begin

  fTotalValorProduto := 0;
  fTotalValorNota := 0;

  fTotalBaseICMS := 0;
  fTotalValorICMS := 0;

  fValorBaseICMS_ST := 0;
  fValorICMS_ST     := 0;

  fDestinatarioCNPJ := SomenteNumeros(fDestinatarioCNPJ);

  if (Length(Trim(fDestinatarioCNPJ)) < 11) or
     (Trim(fDestinatarioCNPJ) = '11111111111') then
     fDestinatarioCNPJ := '';

  vTime := StrToTime(fHoraSaida);

  //0.000694444 um minuto

  if (Time - vTime) > 0.003124998 then begin  //direfença de 5 minutos
    fEmissao := FormatDateTime('DD/MM/YYYY hh:mm:ss',now);
    fSaida   := FormatDateTime('DD/MM/YYYY hh:mm:ss',now);
  end;


  Result := '[infNFe]'          + sLineBreak +
            'versao=3.10'       + sLineBreak +
            '[Identificacao]'   + sLineBreak +
            'NaturezaOperacao=' + fNaturezaOperacao + sLineBreak +
            'Modelo='           + IntToStr(fModelo) + sLineBreak +
            'Codigo='           + Padl(IntToStr(fCodigo),9,'0')  + sLineBreak +
            'Numero='           + Padl(IntToStr(fNumero),9,'0')  + sLineBreak +
            'Serie='            + IntToStr(fSerie) +  sLineBreak +
            'Emissao='          + fEmissao         +  sLineBreak +
            'Saida='            + fSaida           +  sLineBreak +
            'Tipo='             + IntToStr(fTipo)  +  sLineBreak +
            'FormaPag='         + fFormaPag        +  sLineBreak +
            'Finalidade='       + IntToStr(fFinalidade) + sLineBreak +
            'idDest='           + IntToStr(fidDest)     + sLineBreak +
            'indFinal='         + IntToStr(findFinal)   + sLineBreak +
            'indPres='          + IntToStr(findPres)    + sLineBreak +
            'tpImp='            + IntToStr(ftpImp)      + sLineBreak +
            'tpEmis='           + IntToStr(ftpEmis)     + sLineBreak;

            if ftpEmis = 9 then begin
              Result := Result + 'dhCont=' + fdhCont + sLineBreak +
                                 'xJust='  + fxJust  + sLineBreak;
            end;

            if fFinalidade = 4 then begin
              Result := Result + '[NFRef001]' +  sLineBreak +
                                 'Tipo=NFE' + sLineBreak +
                                 'refNFe='   + frefNF +  sLineBreak;
            end
            else begin
              result  :=  Result  + BuscaCuponsReferenciados(vIndex);
              result  :=  Result  + BuscaNFCeReferenciadas(vIndex);
            end;

            result  :=  Result  +
            '[Emitente]'        + sLineBreak    +
            'CNPJ='             + fEmitenteCNPJ +  sLineBreak +
            'IE='               + fEmitenteIE   +  sLineBreak;

            if (fEmitenteIM <> '') and (fEmitenteCNAE <> '') then begin
              Result :=  Result + 'IM=' + fEmitenteIM + sLineBreak +
                                  'CNAE=' + fEmitenteCNAE + sLineBreak;
            end;

            result  :=  Result  +
            'Razao='            + fEmitenteRazao       + sLineBreak +
            'Fantasia='         + fEmitenteFantasia    + sLineBreak +
            'Fone='             + fEmitenteFone        + sLineBreak +
            'CEP='              + fEmitenteCEP         + sLineBreak +
            'Logradouro='       + fEmitenteLogradouro  + sLineBreak +
            'Numero='           + fEmitenteNumero      + sLineBreak +
            'Complemento='      + fEmitenteComplemento + sLineBreak +
            'Bairro='           + fEmitenteBairro      + sLineBreak +
            'CidadeCod='        + fEmitenteCidadeCod   + sLineBreak +
            'Cidade='           + fEmitenteCidade      + sLineBreak +
            'UF='               + fEmitenteUF          + sLineBreak +
            'PaisCod='          + fEmitentePaisCod     + sLineBreak +
            'Pais='             + fEmitentePais        + sLineBreak +
            'CRT='              + fEmitenteCRT         + sLineBreak;



            if fDestinatarioCNPJ <> '' then begin
              Result := Result  +
              '[Destinatario]'  + sLineBreak +
              'CNPJ='           + fDestinatarioCNPJ +  sLineBreak;

              if NOT (fDestinatarioIE = '') then
                Result := Result + 'IE=' + fDestinatarioIE +  sLineBreak;

              Result := Result +
              'NomeRazao='     + fDestinatarioNomeRazao   + sLineBreak +
              'Fone='          + fDestinatarioFone        + sLineBreak +
              'CEP='           + fDestinatarioCEP         + sLineBreak +
              'Logradouro='    + fDestinatarioLogradouro  + sLineBreak +
              'Numero='        + fDestinatarioNumero      + sLineBreak +
              'Complemento='   + fDestinatarioComplemento + sLineBreak +
              'Bairro='        + fDestinatarioBairro      + sLineBreak +
              'CidadeCod='     + fDestinatarioCidadeCod   + sLineBreak +
              'Cidade='        + fDestinatarioCidade      + sLineBreak +
              'UF='            + fDestinatarioUF          + sLineBreak +
              'indIEDest='     + fDestinatarioindIEDest   + sLineBreak;
            end;

            Result := Result +
            '[DadosAdicionais]' + sLineBreak;

            if (fModelo <> 65) then
              Result := Result + 'infCpl=' + finfCpl + sLineBreak
            else
              Result := Result + 'infCpl=' + finfCpl + fMensagem3 + sLineBreak;

            if fInfAdFisco <> '' then
              Result := Result + 'infAdFisco=' + fInfAdFisco + sLineBreak;

            if (fModelo = 65) and (trim(fMensagem3) <> '') then begin
              Result := Result +
              '[InfAdic]' + sLineBreak +
              'xCampo=Dados' + sLineBreak +
              'xTexto=' + fMensagem3 + sLineBreak;
            end;

  if (fCNPJautXML <> '') or (fCPFautXML <> '') then begin
    Result := Result + '[autXML001]' + sLineBreak;
    if (fCNPJautXML <> '') then
      Result := Result + 'CNPJautXML=' + fCNPJautXML + sLineBreak;
    if (fCPFautXML <> '') then
      Result := Result + 'CPFautXML='  + fCPFautXML + sLineBreak;
  end;

  result := Result + CalculaTotaisNfe;

  vTotalProduto   := TrocarCaracter(FormatFloat('##########0.00',fTotalValorProduto),'.',',');
  vTotalValorNota := TrocarCaracter(FormatFloat('##########0.00',fTotalValorNota),'.',',');

  result := Result + '[Total]'       + sLineBreak +
                     'BaseICMS='     + TrocarCaracter(FormatFloat('######0.00',fTotalBaseICMS),',','.')  +  sLineBreak +
                     'ValorICMS='    + TrocarCaracter(FormatFloat('######0.00',fTotalValorICMS),',','.') +  sLineBreak +
                     'ValorProduto=' + TrocarCaracter(vTotalProduto,',','.') + sLineBreak +
                     'ValorNota='    + TrocarCaracter(vTotalValorNota,',','.')   +  sLineBreak;

  if fValorBaseICMS_ST > 0 then begin
    result := result + 'BaseICMSSubstituicao=' + TrocarCaracter(FormatFloat('######0.00',fValorBaseICMS_ST),',','.') + sLineBreak +
                       'ValorICMSSubstituicao=' + TrocarCaracter(FormatFloat('######0.00',fValorICMS_ST),',','.') + sLineBreak;
  end;

  result := Result + '[Transportador]' + sLineBreak +
                     'FretePorConta='  + fTransportadorFretePorConta + sLineBreak;

  IF (fFinalidade = 1) AND (fModelo = 55) then begin  //se finalidade for normal
    for fItens in fListaItens  do begin //e o item for combustivel e o cfop for de combustivel
      if fItens.Combustivel and (CFOPCombustivel(fItens.CFOP)) then begin
        result := result + 'CNPJCPF=' + fDestinatarioCNPJ + sLineBreak +
                           'xNome='   + fDestinatarioNomeRazao + sLineBreak +
                           'xEnder='  + fDestinatarioLogradouro + sLineBreak +
                           'xMun='    + fDestinatarioCidade + sLineBreak +
                           'UF='      + fDestinatarioUF + sLineBreak;
        break;
      end;
    end;
  end;

  vIndex := 1;
  for fFormaPgto in fListaPag do begin
    result := Result + '[pag' + Padl(IntToStr(vIndex),3,'0') + ']' +  sLineBreak +
                       'tPag='  + fFormaPgto.tPag  + sLineBreak +
                       'vPag='  + fFormaPgto.vPag  + sLineBreak;


    if fFormaPgto.CNPJ <> '' then begin
      if fFormaPgto.tBand = 'Visa' then
        vBand := '01'
      else if fFormaPgto.tBand = 'RedeCard' then
        vBand := '02'
      else if fFormaPgto.tBand = 'America Express' then
        vBand := '03'
      else
        vBand := '99';

      result := Result + 'tpIntegra=1' + sLineBreak + // 1=TEF, 2=POS
                         'CNPJ='  + fFormaPgto.CNPJ  + sLineBreak +
                         'tBand=' + vBand + sLineBreak +
                         'cAut='  + fFormaPgto.cAut  + sLineBreak;
    end;
    Inc(vIndex);
  end;
end;

function TCapaNFe.GetValorBaseICMS_ItensNFe(pCodigo: string;
  pItem: integer): Double;
var
  I : INTEGER;
begin
  try
    for I := 0 TO fListaItens.Count - 1  do begin
      IF ((fListaItens.Items[I] AS TItensNFe).Codigo = pCodigo) AND (I = pItem) Then begin
        Result := (fListaItens.Items[I] AS TItensNFe).ValorBaseICMS;
      end;
    end;
  except
    result := 0;
  end;
end;

function TCapaNFe.GetValorICMS_ItensNFe(pCodigo: string;
  pItem: integer): Double;
var
  I : INTEGER;
begin
  try
    for I := 0 TO fListaItens.Count - 1  do begin
      IF ((fListaItens.Items[I] AS TItensNFe).Codigo = pCodigo) AND (I = pItem) Then begin
        Result := (fListaItens.Items[I] AS TItensNFe).ValorICMS;
      end;
    end;
  except
    result := 0;
  end;
end;

function TCapaNFe.ValidaNota: boolean;
begin
  try
    if fModelo = 65 then begin
      result := (fListaPag.Count >= 1);
      if not Result then begin
        Atencao('Não foi informado a forma de pagamento!' + sLineBreak +
                'Não será possível emitir essa nota!');
      end;
    end
    else
      result := true;
  except
  end;
end;

function TCapaNFe.GravaItensNfe(pNFCe : Boolean = False): Boolean;
var
//  qryBuscaItens : TADOQuery;
  qryNotaItem   : TADOQuery;
  fItens        : TItensNFe;
begin
  try



//    qryBuscaItens := TADOQuery.Create(nil);
//    qryBuscaItens.Connection := DM.ADOconexao;
//
    qryNotaItem := TADOQuery.Create(nil);
    qryNotaItem.Connection := DM.ADOconexao;


//    with qryBuscaItens do begin
//      Close;
//      SQL.Clear;
//      if pNFCe then begin
//        SQL.Add('SELECT C.CODIGO AS CFOP,          ');
//        SQL.Add('       COALESCE(T.PERC_RED_ICMS,0) AS PERCENT_REDUTOR, ');
//        SQL.Add('       B.NUMERO_NFCE AS NUMERONF, ');
//        SQL.Add('       B.Serie,                   ');
//        SQL.Add('       A.TAB_CODIGO AS CodProduto,');
//        SQL.Add('       T.TAB_DESCR AS Descricao,  ');
//        SQL.Add('       T.CodSituacaoTributaria,   ');
//        SQL.Add('       T.UNI_SIGLA AS Unidade,    ');
//        SQL.Add('       ALIQUOTA = COALESCE (CASE WHEN ICM_PERC IN (''FF'', ''II'', ''NN'', ''0000'') THEN '''' ');
//        SQL.Add('                            WHEN ICM_PERC IS NULL THEN '''' ');
//        SQL.Add('                            ELSE RTRIM( ARQ_ICMS ) END, ''''), ');
//        SQL.Add('        A.ARQ_QTDE as Quantidade, ');
//        SQL.Add('        A.ARQ_VRUNITL AS VlrUnitario,');
//        SQL.Add('        ROUND((A.ARQ_QTDE * A.ARQ_VRUNITL),2,case A.round when 0 then 0 else 1 END) as ValorTotal,');
//        SQL.Add('        ICMS = COALESCE (CASE WHEN ICM_PERC IN (''FF'', ''II'', ''NN'', ''0000'') THEN 0 ');
//        SQL.Add('                         WHEN ICM_PERC IS NULL 						 		 THEN 0 ');
//        SQL.Add('                         ELSE CAST (ICM_PERC AS NUMERIC (19, 2)) * ARQ_QTDE * ARQ_VRUNITL / 100 END, 0), ');
//        SQL.Add('        BASEICMS = COALESCE (CASE WHEN ICM_PERC IN (''FF'', ''II'', ''NN'', ''0000'') THEN 0 ');
//        SQL.Add('                             WHEN ICM_PERC IS NULL 						 	     THEN 0 ');
//        SQL.Add('                             ELSE ARQ_QTDE * ARQ_VRUNITL END, 0), ');
//        SQL.ADD('                             A.EncerranteInicial, ');
//        SQL.ADD('                             A.EncerranteFinal    ');
//        SQL.Add(' FROM BASICO B, ARQUIVO A, TABELA T, NOTAFISCAL NF, TABGRUPO G, CFOPS C ');
//        SQL.Add(' WHERE B.IdNotaFiscal = A.IdNotaFiscal ');
//        SQL.Add(' AND T.TAB_CODIGO = A.TAB_CODIGO ');
//        SQL.Add(' AND NF.NumeroNF = B.NUMERO_NFCE ');
//        SQL.Add(' and nf.Modelo = b.MODELONF ');
//        SQL.Add(' and nf.Serie = b.Serie ');
//        SQL.Add(' AND NF.Modelo = ''65'' ');
//        SQL.Add(' AND NF.Serie = :serie ');
//        SQL.Add(' AND T.GRU_CODIGO = G.GRU_CODIGO ');
//        SQL.Add(' AND G.CFOP_ID = C.CFOPS_ID ');
//        SQL.Add(' AND Not A.TAB_CODIGO is null ');
//        SQL.Add(' AND B.NUMERO_NFCE = :NumeroNFCe ');
//        Parameters.ParamByName('NumeroNFCe').Value := fNumero;
//        Parameters.ParamByName('serie').Value := fSerie;
//      end
//      else begin
//        SQL.Add(' Select NumeroNF,                                        ');
//        SQL.Add('        Serie,                                           ');
//        SQL.Add('        Modelo,                                          ');
//        SQL.Add('        CodProduto,                                      ');
//        SQL.Add('        Descricao,                                       ');
//        SQL.Add('        CodSituacaoTributaria,                           ');
//        SQL.Add('        Unidade,                                         ');
//        SQL.Add('        Aliquota,                                        ');
//        SQL.Add('        Sum( Quantidade  ) as Quantidade,                ');
//        SQL.Add('        VlrUnitario,                                     ');
//        SQL.Add('        Sum( ValorTotal  ) as ValorTotal,                ');
//        SQL.Add('        Sum( ICMS )        as ICMS,                      ');
//        SQL.Add('        Sum( BaseICms )    as BaseICms                   ');
//        SQL.Add('   From vw_NotaItem                                      ');
//        SQL.Add('  WHERE Not CodProduto is null and  NUMERONF = :NumeroNF ');
//        sql.add('    And Modelo = :modelo  AND Serie = :serie             ');
//        SQL.Add('  Group By NumeroNF,                                     ');
//        SQL.Add('           Serie,                                        ');
//        SQL.Add('           Modelo,                                       ');
//        SQL.Add('           CodSituacaoTributaria,                        ');
//        SQL.Add('           Unidade,                                      ');
//        SQL.Add('           Aliquota,                                     ');
//        SQL.Add('           CodGrupo,                                     ');
//        SQL.Add('           Descricao,                                    ');
//        SQL.Add('           CodProduto,                                   ');
//        SQL.Add('           VlrUnitario                                   ');
//        SQL.Add('  Order By CodProduto, VlrUnitario                       ');
//        Parameters.ParamByName('NumeroNF').Value := fCodigo;
//        Parameters.ParamByName('Modelo').Value := fModelo;
//        Parameters.ParamByName('serie').Value := fSerie;
//      end;
//      Open;
//    end;

//    if qryBuscaItens.IsEmpty then begin
//      Result := False;
//      exit;
//    END;
//
//    qryBuscaItens.First;


    for fItens in fListaItens  do begin
//    while not qryBuscaItens.Eof do begin
      qryNotaItem.Close;
      qryNotaItem.sql.clear;
      qryNotaItem.SQL.Add(' INSERT INTO NOTAITEM(ALIQUOTA,          ');
      qryNotaItem.SQL.Add('                      CodProduto,        ');
      qryNotaItem.SQL.Add('                      CST,               ');
      qryNotaItem.SQL.Add('                      Descricao,         ');
      qryNotaItem.SQL.Add('                      ICMS,              ');
      qryNotaItem.SQL.Add('                      NumeroNF,          ');
      qryNotaItem.SQL.Add('                      Quantidade,        ');
      qryNotaItem.SQL.Add('                      Serie,             ');
      qryNotaItem.SQL.Add('                      Modelo,            ');
      qryNotaItem.SQL.Add('                      Unidade,           ');
      qryNotaItem.SQL.Add('                      Vlrtotal,          ');
      qryNotaItem.SQL.Add('                      VlrUnitario,       ');
      qryNotaItem.SQL.Add('                      CFOP,              ');
      qryNotaItem.SQL.Add('                      ALIQ_ICMS,         ');
      qryNotaItem.SQL.Add('                      ALIQ_ICMS_CHEIA,   ');
      qryNotaItem.SQL.Add('                      VL_BC_ICMS,        ');
      qryNotaItem.SQL.Add('                      VL_REDUCAO,        ');
      qryNotaItem.SQL.Add('                      VL_OUT_DESP_ITEM,  ');
      qryNotaItem.SQL.Add('                      VL_FRETE_ITEM,     ');
      qryNotaItem.SQL.Add('                      VL_ICMS,           ');
      qryNotaItem.SQL.Add('                      ALIQ_ICMS_ST,      ');
      qryNotaItem.SQL.Add('                      VL_BC_ICMS_ST,     ');
      qryNotaItem.SQL.Add('                      VL_ICMS_ST,        ');
      qryNotaItem.SQL.Add('                      VL_BC_PIS,         ');
      qryNotaItem.SQL.Add('                      CST_PIS,           ');
      qryNotaItem.SQL.Add('                      ALIQ_PIS,          ');
      qryNotaItem.SQL.Add('                      QUANT_BC_PIS,      ');
      qryNotaItem.SQL.Add('                      VL_PIS,            ');
      qryNotaItem.SQL.Add('                      VL_BC_COFINS,      ');
      qryNotaItem.SQL.Add('                      CST_COFINS,        ');
      qryNotaItem.SQL.Add('                      ALIQ_COFINS,       ');
      qryNotaItem.SQL.Add('                      QUANT_BC_COFINS,   ');
      qryNotaItem.SQL.Add('                      VL_COFINS,         ');
      qryNotaItem.SQL.Add('                      ALIQ_ISSQN,        ');
      qryNotaItem.SQL.Add('                      VL_BC_ISSQN,       ');
      qryNotaItem.SQL.Add('                      VL_ISSQN,          ');
      qryNotaItem.SQL.Add('                      VL_ISENTO,         ');
      qryNotaItem.SQL.Add('                      VL_NAO_TRIB,       ');
      qryNotaItem.SQL.Add('                      VL_DESC_ITEM,      ');
      qryNotaItem.SQL.Add('                      PERCENT_IVA,       ');
      if pNFCe then begin
        qryNotaItem.SQL.Add('                    ENCERRANTEINICIAL, ');
        qryNotaItem.SQL.Add('                    ENCERRANTEFINAL,   ');
      end;
      if fModelo = 65 then begin
        qryNotaItem.SQL.ADD('                    IDNOTAFISCAL,      ');
      END;
      qryNotaItem.SQL.Add('                      PIS_COFINS_CUMULATIVO,    ');
      qryNotaItem.SQL.Add('                      TIPO_ALIQUOTA_PIS_COFINS, ');
      qryNotaItem.SQL.Add('                      PERCENT_REDUTOR)   ');
      qryNotaItem.SQL.Add('VALUES (:ALIQUOTA,                       ');
      qryNotaItem.SQL.Add('        :CodProduto,                     ');
      qryNotaItem.SQL.Add('        :CST,                            ');
      qryNotaItem.SQL.Add('        :Descricao,                      ');
      qryNotaItem.SQL.Add('        :ICMS,                           ');
      qryNotaItem.SQL.Add('        :NumeroNF,                       ');
      qryNotaItem.SQL.Add('        :Quantidade,                     ');
      qryNotaItem.SQL.Add('        :Serie,                          ');
      qryNotaItem.SQL.Add('        :Modelo,                         ');
      qryNotaItem.SQL.Add('        :Unidade,                        ');
      qryNotaItem.SQL.Add('        :Vlrtotal,                       ');
      qryNotaItem.SQL.Add('        :VlrUnitario,                    ');
      qryNotaItem.SQL.Add('        :CFOP,                           ');
      qryNotaItem.SQL.Add('        :ALIQ_ICMS,                      ');
      qryNotaItem.SQL.Add('        :ALIQ_ICMS_CHEIA,                ');
      qryNotaItem.SQL.Add('        :VL_BC_ICMS,                     ');
      qryNotaItem.SQL.Add('        :VL_REDUCAO,                     ');
      qryNotaItem.SQL.Add('        :VL_OUT_DESP_ITEM,               ');
      qryNotaItem.SQL.Add('        :VL_FRETE_ITEM,                  ');
      qryNotaItem.SQL.Add('        :VL_ICMS,                        ');
      qryNotaItem.SQL.Add('        :ALIQ_ICMS_ST,                   ');
      qryNotaItem.SQL.Add('        :VL_BC_ICMS_ST,                  ');
      qryNotaItem.SQL.Add('        :VL_ICMS_ST,                     ');
      qryNotaItem.SQL.Add('        :VL_BC_PIS,                      ');
      qryNotaItem.SQL.Add('        :CST_PIS,                        ');
      qryNotaItem.SQL.Add('        :ALIQ_PIS,                       ');
      qryNotaItem.SQL.Add('        :QUANT_BC_PIS,                   ');
      qryNotaItem.SQL.Add('        :VL_PIS,                         ');
      qryNotaItem.SQL.Add('        :VL_BC_COFINS,                   ');
      qryNotaItem.SQL.Add('        :CST_COFINS,                     ');
      qryNotaItem.SQL.Add('        :ALIQ_COFINS,                    ');
      qryNotaItem.SQL.Add('        :QUANT_BC_COFINS,                ');
      qryNotaItem.SQL.Add('        :VL_COFINS,                      ');
      qryNotaItem.SQL.Add('        :ALIQ_ISSQN,                     ');
      qryNotaItem.SQL.Add('        :VL_BC_ISSQN,                    ');
      qryNotaItem.SQL.Add('        :VL_ISSQN,                       ');
      qryNotaItem.SQL.Add('        :VL_ISENTO,                      ');
      qryNotaItem.SQL.Add('        :VL_NAO_TRIB,                    ');
      qryNotaItem.SQL.Add('        :VL_DESC_ITEM,                   ');
      qryNotaItem.SQL.Add('        :PERCENT_IVA,                    ');
      if pNFCe then begin
        qryNotaItem.SQL.Add('      :ENCERRANTEINICIAL,              ');
        qryNotaItem.SQL.Add('      :ENCERRANTEFINAL,                ');
      end;
      if fModelo = 65 then begin
        qryNotaItem.SQL.Add('        :IDNOTAFISCAL,                 ');
      END;
      qryNotaItem.SQL.Add('        :PIS_COFINS_CUMULATIVO,          ');
      qryNotaItem.SQL.Add('        :TIPO_ALIQUOTA_PIS_COFINS,       ');
      qryNotaItem.SQL.Add('        :PERCENT_REDUTOR)                ');
      qryNotaItem.Parameters.ParamByName('Aliquota').Value        := FormatFloat('#######00',fitens.AliquotaICMS);     // qryBuscaItens.FieldByName('Aliquota').AsString;
      qryNotaItem.Parameters.ParamByName('CodProduto').Value      := fItens.Codigo;          //qryBuscaItens.FieldByName('CodProduto').AsString;
      qryNotaItem.Parameters.ParamByName('CST').Value             := Copy(fItens.CST,2,2);  //Copy(qryBuscaItens.FieldByName('CodSituacaoTributaria').AsString,2,2);
      qryNotaItem.Parameters.ParamByName('Descricao').Value       := fItens.Descricao;      //qryBuscaItens.FieldByName('Descricao').AsString;
      qryNotaItem.Parameters.ParamByName('ICMS').Value            := fitens.ValorICMS;      //qryBuscaItens.FieldByName('ICMS').AsFloat;
      qryNotaItem.Parameters.ParamByName('NumeroNF').Value        := fCodigo; //esse é o numero nf
      qryNotaItem.Parameters.ParamByName('Quantidade').Value      := fitens.Quantidade;     //qryBuscaItens.FieldByName('ValorTotal').AsFloat / qryBuscaItens.FieldByName('VlrUnitario').AsFloat;//qryBuscaItens.FieldByName('Quantidade').AsFloat;
      qryNotaItem.Parameters.ParamByName('Serie').Value           := IntToStr(fserie);
      qryNotaItem.Parameters.ParamByName('Modelo').Value          := IntToStr(fModelo);
      qryNotaItem.Parameters.ParamByName('Unidade').Value         := fItens.Unidade;        //qryBuscaItens.FieldByName('Unidade').AsString;
      qryNotaItem.Parameters.ParamByName('Vlrtotal').Value        := fitens.ValorTotal;     //qryBuscaItens.FieldByName('ValorTotal').AsFloat;
      qryNotaItem.Parameters.ParamByName('VlrUnitario').Value     := fitens.ValorUnitario;  //qryBuscaItens.FieldByName('VlrUnitario').AsFloat;
      qryNotaItem.Parameters.ParamByName('CFOP').Value            := fItens.CFOP;           //qryBuscaItens.FieldByName('CFOP').AsString;
      qryNotaItem.Parameters.ParamByName('ALIQ_ICMS').Value       := fitens.AliquotaICMS;           {Alíquota ICMS }
      qryNotaItem.Parameters.ParamByName('ALIQ_ICMS_CHEIA').Value := 0;           {Alíquota ICMS Cheia             }

      if fitens.ValorICMS > 0 then
        qryNotaItem.Parameters.ParamByName('VL_BC_ICMS').Value    := fitens.ValorBaseICMS; {Base de ICMS           }

      qryNotaItem.Parameters.ParamByName('VL_REDUCAO').Value      := 0;           {Valor de redução                }
      qryNotaItem.Parameters.ParamByName('VL_OUT_DESP_ITEM').Value:= 0;           {Valor de outras despesa por item}
      qryNotaItem.Parameters.ParamByName('VL_FRETE_ITEM').Value   := 0;           {Valor do frete por item         }
      qryNotaItem.Parameters.ParamByName('VL_ICMS').Value         := 0;           {Valor de ICMS do item           }
      qryNotaItem.Parameters.ParamByName('ALIQ_ICMS_ST').Value    := 0;           {Alíquota de ICMS ST             }
      qryNotaItem.Parameters.ParamByName('VL_BC_ICMS_ST').Value   := 0;           {Valor base ICMS ST              }
      qryNotaItem.Parameters.ParamByName('VL_ICMS_ST').Value      := 0;           {Valor do ICMS ST                }
      qryNotaItem.Parameters.ParamByName('VL_BC_PIS').Value       := 0;           {Valor da base do PIS            }
      qryNotaItem.Parameters.ParamByName('CST_PIS').Value         := fItens.PISCST;
      qryNotaItem.Parameters.ParamByName('ALIQ_PIS').Value        := fItens.PISAliquota; {Alíquota do PIS          }
      qryNotaItem.Parameters.ParamByName('QUANT_BC_PIS').Value    := 0;           {Quantidade da base do PIS       }
      qryNotaItem.Parameters.ParamByName('VL_PIS').Value          := fItens.PISValor;    {Valor do PIS             }
      qryNotaItem.Parameters.ParamByName('VL_BC_COFINS').Value    := 0;           {Valor da base do COFINS         }
      qryNotaItem.Parameters.ParamByName('CST_COFINS').Value      := fItens.COFINSCST;
      qryNotaItem.Parameters.ParamByName('ALIQ_COFINS').Value     := fItens.COFINSAliquota; {Aliquota do COFINS    }
      qryNotaItem.Parameters.ParamByName('QUANT_BC_COFINS').Value := 0;           {Quantidade da base do COFINS    }
      qryNotaItem.Parameters.ParamByName('VL_COFINS').Value       := fItens.COFINSValor;    {Valor do COFINS       }
      qryNotaItem.Parameters.ParamByName('ALIQ_ISSQN').Value      := 0;           {Alíquota de ISSQN               }
      qryNotaItem.Parameters.ParamByName('VL_BC_ISSQN').Value     := 0;           {Valor da base de ISSQN          }
      qryNotaItem.Parameters.ParamByName('VL_ISSQN').Value        := 0;           {Valor de ISSQN                  }
      qryNotaItem.Parameters.ParamByName('VL_ISENTO').Value       := 0;           {Valor de Isento                 }
      qryNotaItem.Parameters.ParamByName('VL_NAO_TRIB').Value     := 0;           {Valor de não tributados         }
      qryNotaItem.Parameters.ParamByName('VL_DESC_ITEM').Value    := 0;           {Valor do desconto do item       }
      qryNotaItem.Parameters.ParamByName('PERCENT_IVA').Value     := 0;           {Percentual de IVA               }
      qryNotaItem.Parameters.ParamByName('PERCENT_REDUTOR').Value := fItens.PercentualReducao;// qryBuscaItens.FieldByName('PERCENT_REDUTOR').AsFloat;  {Percentual de redutor           }
      qryNotaItem.Parameters.ParamByName('PIS_COFINS_CUMULATIVO').Value    := fItens.PIS_COFINS_CUMULATIVO;
      qryNotaItem.Parameters.ParamByName('TIPO_ALIQUOTA_PIS_COFINS').Value := fItens.TIPO_ALIQUOTA_PIS_COFINS;

      if fModelo = 65 then begin
        qryNotaItem.Parameters.ParamByName('IDNOTAFISCAL').Value    := fIDNotaFiscal;
      END;
      if pNFCe then begin
        qryNotaItem.Parameters.ParamByName('ENCERRANTEINICIAL').Value := fItens.EncInicial; // qryBuscaItens.FieldByName('ENCERRANTEINICIAL').AsFloat;
        qryNotaItem.Parameters.ParamByName('ENCERRANTEFINAL').Value   := fItens.EncFinal; //qryBuscaItens.FieldByName('ENCERRANTEFINAL').AsFloat;
      end;
      qryNotaItem.ExecSQL;

//      qryBuscaItens.Next;
    end;

    {Final do processamento dos itens da nota fiscal}
    Result := True;
  except
    on e: exception do begin
      Result := False;
      Atencao(e.message);
    end;
  end;
end;

function TCapaNFe.ApagaNFe: Boolean;
var
  qryCapa : TADOQuery;
  qryNotaItem   : TADOQuery;
begin
  try
    qryCapa := TADOQuery.Create(nil);
    qryCapa.Connection := DM.ADOconexao;

    qryNotaItem := TADOQuery.Create(nil);
    qryNotaItem.Connection := DM.ADOconexao;

    with qryNotaItem do begin
      Close;
      SQL.Clear;
      sql.Add(' DELETE NOTAITEM ');
      SQL.Add('  WHERE NUMERONF = :NumeroNf ');
      SQL.Add('    AND SERIE    = :SERIE ');
      SQL.Add('    AND MODELO   = :MODELO');
      Parameters.ParamByName('NumeroNf').Value := fCodigo;
      Parameters.ParamByName('SERIE').Value    := fSerie;
      Parameters.ParamByName('Modelo').Value   := IntToStr(fModelo);
      ExecSQL;
    end;

    with qryCapa DO begin
      Close;
      SQL.Clear;
      SQL.Add(' DELETE NOTAFISCAL');
      SQL.Add('  WHERE NUMERONF   = :NUMERONF');
      SQL.Add('    AND SERIE      = :SERIE ');
      SQL.Add('    AND CODCLIENTE = :CODCLIENTE');
      SQL.Add('    AND MODELO     = :MODELO');
      Parameters.ParamByName('NumeroNf').Value := fCodigo;
      Parameters.ParamByName('SERIE').Value := fSerie;
      Parameters.ParamByName('CODCLIENTE').Value := fDestinatarioCodigo;
      Parameters.ParamByName('Modelo').Value := IntToStr(fModelo);
      ExecSQL;
    end;
    Result := True;
  except
    Result := False;
  end;

end;

function TCapaNFe.GravaNFe(pNFCe: Boolean = False): Boolean;
var
  qryCapa : TADOQuery;
begin
  try
    SetStatus('P');
    Result := true;
    qryCapa := TADOQuery.Create(nil);
    qryCapa.Connection := DM.ADOconexao;

//    CalculaTotaisNfe;

    with qryCapa do begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO NOTAFISCAL (NumeroNF,                      ');
      SQL.Add('                        SaidaEntrada,                  ');
      SQL.Add('                        Serie,                         ');
      SQL.Add('                        CFOP,                          ');
      SQL.Add('                        InscricaoSubstituicao,         ');
      SQL.Add('                        Inscricaoestadual,             ');
      SQL.Add('                        CodCliente,                    ');
      SQL.Add('                        Nome,                          ');
      SQL.Add('                        CnpjCpf,                       ');
      SQL.Add('                        Emissao,                       ');
      SQL.Add('                        Saida,                         ');
      SQL.Add('                        HoraSaida,                     ');
      SQL.Add('                        Endereco,                      ');
      SQL.Add('                        Bairro,                        ');
      SQL.Add('                        CEP,                           ');
      SQL.Add('                        Municipio,                     ');
      SQL.Add('                        FoneFax,                       ');
      SQL.Add('                        UF,                            ');
      SQL.Add('                        BaseIcms,                      ');
      SQL.Add('                        Icms,                          ');
      SQL.Add('                        BaseSubstituicao,              ');
      SQL.Add('                        ValorSubstituicao,             ');
      SQL.Add('                        TotalProduto,                  ');
      SQL.Add('                        TotalNota,                     ');
      SQL.Add('                        Frete,                         ');
      SQL.Add('                        Seguro,                        ');
      SQL.Add('                        Outros,                        ');
      SQL.Add('                        IPI,                           ');
      SQL.Add('                        Observacao,                    ');
      SQL.Add('                        Mensagem1,                     ');
      SQL.Add('                        Mensagem2,                     ');
      SQL.Add('                        Mensagem3,                     ');
      SQL.Add('                        NaturezaOperacao,              ');
      SQL.Add('                        VALORDESCONTO,                 ');
      SQL.Add('                        Cancelada,                     ');
      SQL.Add('                        Modelo,                        ');
//      SQL.Add('                        VALORISENTO,                   ');
//      SQL.Add('                        VALOROUTRAS,                   ');
//      SQL.Add('                        ALIQICMS,                      ');
      SQL.Add('                        VERSAO_NFE,                    ');
      SQL.Add('                        JUSTIFICATIVA_CONTINGENCIA,    ');
//      SQL.Add('                        DATA_HORA_CONTINGENCIA,        ');
//      SQL.Add('                        DATA_HORA_RECEB_NFE,           ');
      SQL.Add('                        STATUS_CTG,                    ');
      SQL.Add('                        TIPO_AMBIENTE_NFE,             ');
      SQL.Add('                        NUMERO_LOTE_NFE,               ');
      SQL.Add('                        NUMERO_PROTOCOLO_CANCELAMENTO, ');
      SQL.Add('                        NUMERO_PROTOCOLO,              ');
      SQL.Add('                        NOTA_FISCAL_NFE,               ');
      SQL.Add('                        STATUS_NFE,                    ');
      SQL.Add('                        NUMERO_RECIBO,                 ');
      SQL.Add('                        NUMERO_NFE,                    ');
      SQL.Add('                        CONTINGENCIA,                  ');
      SQL.Add('                        MOTIVO_CANCELAMENTO,           ');
      SQL.Add('                        COD_SIT_EFD,                   ');
//      SQL.Add('                        MENSAGEM_FISCO_ID,             ');
//      SQL.Add('                        MENSAGEM_CONTRIBUINTE_ID,      ');
//      SQL.Add('                        CONDICAO_TIPO,                 ');
      SQL.Add('                        CONDICAO_DESCRICAO,            ');
//      SQL.Add('                        TOTAL_OUTRAS_DESP,             ');
      SQL.Add('                        EMPRESA_ID,                    ');
      SQL.Add('                        STATUS,                        ');
      SQL.Add('                        EMITIU,                        ');
//      SQL.Add('                        MENSAGEM_ID1,                  ');
//      SQL.Add('                        MENSAGEM_ID2,                  ');
//      SQL.Add('                        DATA_CANCELAMENTO,             ');
      if fModelo = 65 then begin
        SQL.Add('                        DATACAIXA,                   ');
        SQL.Add('                        IDNOTAFISCAL,                ');
      end;
      SQL.Add('                        DESC_NAT_OPERACAO)             ');
      SQL.Add('                 VALUES(:NumeroNF,                     ');
      SQL.Add('                        :SaidaEntrada,                 ');
      SQL.Add('                        :Serie,                        ');
      SQL.Add('                        :CFOP,                         ');
      SQL.Add('                        :InscricaoSubstituicao,        ');
      SQL.Add('                        :Inscricaoestadual,            ');
      SQL.Add('                        :CodCliente,                   ');
      SQL.Add('                        :Nome,                         ');
      SQL.Add('                        :CnpjCpf,                      ');
      SQL.Add('                        :Emissao,                      ');
      SQL.Add('                        :Saida,                        ');
      SQL.Add('                        :HoraSaida,                    ');
      SQL.Add('                        :Endereco,                     ');
      SQL.Add('                        :Bairro,                       ');
      SQL.Add('                        :CEP,                          ');
      SQL.Add('                        :Municipio,                    ');
      SQL.Add('                        :FoneFax,                      ');
      SQL.Add('                        :UF,                           ');
      SQL.Add('                        :BaseIcms,                     ');
      SQL.Add('                        :Icms,                         ');
      SQL.Add('                        :BaseSubstituicao,             ');
      SQL.Add('                        :ValorSubstituicao,            ');
      SQL.Add('                        :TotalProduto,                 ');
      SQL.Add('                        :TotalNota,                    ');
      SQL.Add('                        :Frete,                        ');
      SQL.Add('                        :Seguro,                       ');
      SQL.Add('                        :Outros,                       ');
      SQL.Add('                        :IPI,                          ');
      SQL.Add('                        :Observacao,                   ');
      SQL.Add('                        :Mensagem1,                    ');
      SQL.Add('                        :Mensagem2,                    ');
      SQL.Add('                        :Mensagem3,                    ');
      SQL.Add('                        :NaturezaOperacao,             ');
      SQL.Add('                        :VALORDESCONTO,                ');
      SQL.Add('                        :Cancelada,                    ');
      SQL.Add('                        :Modelo,                       ');
//      SQL.Add('                        :VALORISENTO,                  ');
//      SQL.Add('                        :VALOROUTRAS,                  ');
//      SQL.Add('                        :ALIQICMS,                     ');
      SQL.Add('                        :VERSAO_NFE,                   ');
      SQL.Add('                        :JUSTIFICATIVA_CONTINGENCIA,   ');
//      SQL.Add('                        :DATA_HORA_CONTINGENCIA,       ');
//      SQL.Add('                        :DATA_HORA_RECEB_NFE,          ');
      SQL.Add('                        :STATUS_CTG,                   ');
      SQL.Add('                        :TIPO_AMBIENTE_NFE,            ');
      SQL.Add('                        :NUMERO_LOTE_NFE,              ');
      SQL.Add('                        :NUMERO_PROTOCOLO_CANCELAMENTO,');
      SQL.Add('                        :NUMERO_PROTOCOLO,             ');
      SQL.Add('                        :NOTA_FISCAL_NFE,              ');
      SQL.Add('                        :STATUS_NFE,                   ');
      SQL.Add('                        :NUMERO_RECIBO,                ');
      SQL.Add('                        :NUMERO_NFE,                   ');
      SQL.Add('                        :CONTINGENCIA,                 ');
      SQL.Add('                        :MOTIVO_CANCELAMENTO,          ');
      SQL.Add('                        :COD_SIT_EFD,                  ');
//      SQL.Add('                        :MENSAGEM_FISCO_ID,           ');
//      SQL.Add('                        :MENSAGEM_CONTRIBUINTE_ID,    ');
//      SQL.Add('                        :CONDICAO_TIPO,               ');
      SQL.Add('                        :CONDICAO_DESCRICAO,            ');
//      SQL.Add('                        :TOTAL_OUTRAS_DESP,           ');
      SQL.Add('                        :EMPRESA_ID,                    ');
      SQL.Add('                        :STATUS,                        ');
      SQL.Add('                        :EMITIU,                        ');
//      SQL.Add('                        :MENSAGEM_ID1,                ');
//      SQL.Add('                        :MENSAGEM_ID2,                ');
//      SQL.Add('                        :DATA_CANCELAMENTO,           ');
      if fModelo = 65 then begin
        SQL.Add('                        :DATACAIXA,                   ');
        SQL.Add('                        :IDNOTAFISCAL,                ');
      end;
      SQL.Add('                        :DESC_NAT_OPERACAO)             ');
      Parameters.ParamByName('NumeroNF').Value                      := fCodigo;    //int NOT NULL,
      Parameters.ParamByName('SaidaEntrada').Value                  := 'S';    //char(1)
      Parameters.ParamByName('Serie').Value                         := iif(fModelo=55,'001',fSerie);    //char(3)
      Parameters.ParamByName('CFOP').Value                          := fCFOP;
      Parameters.ParamByName('InscricaoSubstituicao').Value         := '';    //char(10)
      Parameters.ParamByName('Inscricaoestadual').Value             := Trim(Copy(fDestinatarioIE,1,20));    //char(20)
      Parameters.ParamByName('CodCliente').Value                    := fDestinatarioCodigo;    //int NOT NULL,

      if length(Trim(copy(fDestinatarioNomeRazao,1,40))) <= 2 then
        Parameters.ParamByName('Nome').Value  := 'CONSUMIDOR FINAL'
      else
        Parameters.ParamByName('Nome').Value  := Trim(copy(fDestinatarioNomeRazao,1,40));   //char(40)

      Parameters.ParamByName('CnpjCpf').Value                       := Trim(copy(fDestinatarioCNPJ,1,18));    //char(18)
      Parameters.ParamByName('Emissao').Value                       := fEmissao;         //datetime
      Parameters.ParamByName('Saida').Value                         := fSaida;           //datetime

      Parameters.ParamByName('HoraSaida').Value                     := Trim(copy(fHoraSaida,1,10));              //char(10)
      Parameters.ParamByName('Endereco').Value                      := Trim(copy(fDestinatarioLogradouro,1,40)); //char(40)
      Parameters.ParamByName('Bairro').Value                        := Trim(copy(fDestinatarioBairro,1,20));     //char(20)
      Parameters.ParamByName('CEP').Value                           := Trim(copy(fDestinatarioCEP,1,10));        //char(10)
      Parameters.ParamByName('Municipio').Value                     := Trim(copy(fDestinatarioCidade,1,40));     //char(40)
      Parameters.ParamByName('FoneFax').Value                       := Trim(copy(fDestinatarioFone,1,12));       //char(12)
      Parameters.ParamByName('UF').Value                            := Trim(copy(fDestinatarioUF,1,10));         //char(10)

      Parameters.ParamByName('BaseIcms').Value                      := fTotalBaseICMS;     //float
      Parameters.ParamByName('Icms').Value                          := fTotalValorICMS;    //float
      Parameters.ParamByName('BaseSubstituicao').Value              := 0;                  //float
      Parameters.ParamByName('ValorSubstituicao').Value             := 0;                  //float
      Parameters.ParamByName('TotalProduto').Value                  := fTotalValorProduto; //float
      Parameters.ParamByName('TotalNota').Value                     := fTotalValorNota;    //float
      Parameters.ParamByName('Frete').Value                         := 0;                  //float
      Parameters.ParamByName('Seguro').Value                        := 0;                  //float
      Parameters.ParamByName('Outros').Value                        := 0;                  //float
      Parameters.ParamByName('IPI').Value                           := 0;                  //float
      Parameters.ParamByName('Observacao').Value                    := '';                 //ntext
      Parameters.ParamByName('Mensagem1').Value                     := '';                 //varchar(400)
      Parameters.ParamByName('Mensagem2').Value                     := '';                 //varchar(400)
      Parameters.ParamByName('Mensagem3').Value                     := Trim(Copy(fMensagem3,1,400));         //varchar(400)
      Parameters.ParamByName('NaturezaOperacao').Value              := Trim(Copy(fCFOP,1,50));              //char(50)
      Parameters.ParamByName('VALORDESCONTO').Value                 := 0;                  //real DEFAULT 0 NOT NULL,
      Parameters.ParamByName('Cancelada').Value                     := 'N';                //varchar(1)  DEFAULT 'N' NOT NULL,
      Parameters.ParamByName('Modelo').Value                        := IntToStr(fModelo);  //varchar(4)
//      Parameters.ParamByName('VALORISENTO').Value                   := 0;                //numeric(13, 2)
//      Parameters.ParamByName('VALOROUTRAS').Value                   := 0;                //numeric(13, 2)
//      Parameters.ParamByName('ALIQICMS').Value                      := 0;                //numeric(13, 2)

      Parameters.ParamByName('VERSAO_NFE').Value                    := fVersao;            //int
      Parameters.ParamByName('JUSTIFICATIVA_CONTINGENCIA').Value    := '';                 //varchar(1000)
//      Parameters.ParamByName('DATA_HORA_CONTINGENCIA').Value        := now;              //datetime
//      Parameters.ParamByName('DATA_HORA_RECEB_NFE').Value           := '';               //varchar(19)

      Parameters.ParamByName('STATUS_CTG').Value                    := Trim(Copy(fStatus_Ctg,1,1));    //char(1)
      Parameters.ParamByName('TIPO_AMBIENTE_NFE').Value             := '';    //char(1)
      Parameters.ParamByName('NUMERO_LOTE_NFE').Value               := '';    //varchar(15)
      Parameters.ParamByName('NUMERO_PROTOCOLO_CANCELAMENTO').Value := '';    //varchar(15)
      Parameters.ParamByName('NUMERO_PROTOCOLO').Value              := '';    //varchar(15)
      Parameters.ParamByName('NOTA_FISCAL_NFE').Value               := IntToStr(fNumero) ;    //varchar(9)
      Parameters.ParamByName('STATUS_NFE').Value                    := Trim(Copy(fStatus_Nfe,1,1));    //char(1)
      Parameters.ParamByName('NUMERO_RECIBO').Value                 := '';    //varchar(15)
      Parameters.ParamByName('NUMERO_NFE').Value                    := '';    //varchar(44)
      Parameters.ParamByName('CONTINGENCIA').Value                  := '';    //char(1)
      Parameters.ParamByName('MOTIVO_CANCELAMENTO').Value           := '';    //varchar(50)
      Parameters.ParamByName('COD_SIT_EFD').Value                   := '';    //char(2)
//      Parameters.ParamByName('MENSAGEM_FISCO_ID').Value             := '';    //int
//      Parameters.ParamByName('MENSAGEM_CONTRIBUINTE_ID').Value      := '';    //int
//      Parameters.ParamByName('CONDICAO_TIPO').Value                 := '';    //int
      Parameters.ParamByName('CONDICAO_DESCRICAO').Value            := '';    //varchar(20)
//      Parameters.ParamByName('TOTAL_OUTRAS_DESP').Value             := '';    //float
      Parameters.ParamByName('EMPRESA_ID').Value                    := Trim(Copy(fEmitenteCNPJ,1,18));    //varchar(18)
      Parameters.ParamByName('STATUS').Value                        := 'N';    //char(1)
      Parameters.ParamByName('EMITIU').Value                        := 'N';    //char(1)
//      Parameters.ParamByName('MENSAGEM_ID1').Value                  := ;    //int
//      Parameters.ParamByName('MENSAGEM_ID2').Value                  := ;    //int
//      Parameters.ParamByName('DATA_CANCELAMENTO').Value             := '';    //datetime
      Parameters.ParamByName('DESC_NAT_OPERACAO').Value             := fNaturezaOperacao;    //varchar(100)

      if fModelo = 65 then begin
        Parameters.ParamByName('DATACAIXA').Value     := fDataCaixa;
        Parameters.ParamByName('IDNOTAFISCAL').Value  := fIDNotaFiscal;
      end;

      ExecSQL;

      result := GravaItensNfe(pNFCe);

      if not Result then
        raise Exception.Create('Não foi possível gravar os itens da ' + IIf(pNFCe,'NFC-e','NFe!'));
    end;
  except
    on e: Exception do begin
      result := False;
      ShowMessage(e.Message);
    end;
  end;
end;

procedure TCapaNFe.SetCFOP(value: string = '');
begin
  if value <> '' then begin
    fCFOP := value;
  end
  else begin
    if fEmitenteUF <> fDestinatarioUF then
      fCFOP := '6656'
    else
      fCFOP := '5656';
  end;
end;

procedure TCapaNFe.SetDestinatario(pDestinatarioCNPJ, pDestinatarioIE,
  pDestinatarioNomeRazao, pDestinatarioFone, pDestinatarioCEP,
  pDestinatarioLogradouro, pDestinatarioNumero, pDestinatarioComplemento,
  pDestinatarioBairro, pDestinatarioCidadeCod, pDestinatarioCidade,
  pDestinatarioUF, pDestinatarioindIEDest: string;
  pDestinatarioCodigo: Integer);
begin
  fDestinatarioCNPJ        := pDestinatarioCNPJ;
  fDestinatarioIE          := pDestinatarioIE;
  fDestinatarioNomeRazao   := pDestinatarioNomeRazao;
  fDestinatarioFone        := pDestinatarioFone;
  fDestinatarioCEP         := pDestinatarioCEP;
  fDestinatarioLogradouro  := pDestinatarioLogradouro;
  fDestinatarioNumero      := pDestinatarioNumero;
  fDestinatarioComplemento := pDestinatarioComplemento;
  fDestinatarioBairro      := pDestinatarioBairro;
  fDestinatarioCidadeCod   := pDestinatarioCidadeCod;
  fDestinatarioCidade      := pDestinatarioCidade;
  fDestinatarioUF          := pDestinatarioUF;
  fDestinatarioindIEDest   := pDestinatarioindIEDest;
  fDestinatarioCodigo      := pDestinatarioCodigo;
end;

function TCapaNFe.BuscaNFCeReferenciadas(pIndex: integer): string;
var
  fRefNFCe : TRefNFCe;
  vIndex : Integer;
begin
  if pIndex > 0 then
    vIndex := pIndex
  else
    vIndex := 1;

  result := '';
  for fRefNFCe in fListaRefNFCE  do begin
    if trim(frefNFCE.chave) <> '' then
      Result := Result + '[NFRef'    + Padl(IntToStr(vIndex),3,'0') + ']' +  sLineBreak +
                         'Tipo=NFE'  + sLineBreak +
                         'refNFe='   + frefNFCE.chave +  sLineBreak;
    Inc(vIndex);
  end;
end;

function TCapaNFe.BuscaCuponsReferenciados(var pIndex: integer): string;
var
  fRefECF : TRefEcf;
  vIndex : Integer;
begin
  vIndex := 1;
  result := '';
  for fRefECF in fListaRefECF  do begin
    if trim(fRefECF.nCOO) <> '' then
      Result := Result + '[NFRef'  + Padl(IntToStr(vIndex),3,'0') + ']' +  sLineBreak +
                         'Tipo='   + fRefECF.Tipo   +  sLineBreak +
                         'ModECF=' + fRefECF.ModECF +  sLineBreak +
                         'nECF='   + fRefECF.nECF   +  sLineBreak +
                         'nCOO='   + fRefECF.nCOO   +  sLineBreak;
    Inc(vIndex);
  end;
  pIndex := vIndex;
end;

function TCapaNFe.CFOPCombustivel(pCfop: string): boolean;
begin
  result := false;
  if StrToIntDef(trim(pCfop),0) >= 5650 then begin
    if StrToIntDef(trim(pCfop),0) <= 5667 then begin
      result := true;
    end;
  end;
  if (fidDest = 2) then begin
    if StrToIntDef(trim(pCfop),0) >= 6650 then begin
      if StrToIntDef(trim(pCfop),0) <= 6667 then begin
        result := true;
      end;
    end;
  end;
end;

function TCapaNFe.CalculaTotaisNfe: string;
var
  fItens : TItensNFe;
  vIndex  : Integer;
  vGTICMS  : String; //grupo tributação icms
  vCST      : String;
  vVrBC      : Double;
  vVrTotal    : String;
  vQtde        : string;
  vVrUnit       : String;
  vValorDesconto : String;
  vValorBase    : String;
  vAliquota    : String;
  vValor      : String;
  vRedBC      : String;
  vCest       : String;
  vqTrib       : String;
begin
  vIndex := 1;

  for fItens in fListaItens  do begin
    vGTICMS := copy(fItens.CST,2,2);
    if fModelo = 65 then begin
      //NT-2015-002 N12-30 Entra em vigor em 01/01/2016 em relação a NFC-e
      //NFC-e com CST diferente da relação abaixo:
      //- 00-Tributada integralmente;
      //- 20-Com redução da Base de Cálculo;
      //- 40-Isenta;
      //- 41-Não tributada;
      //- 60-ICMS cobrado anteriormente por substituição tributária;
      //Exceção 1: Aceitar CST=90-Outros, a critério da UF.
      //Exceção 2: A regra de validação não se aplica, em produção, para
      //Nota Fiscal com Data de Emissão anterior a 01/01/2016.

      //No sistema para a NFC-e está padrão CST=60 unit untImprimeItens procedure ImprimeItemNFCe
      if not (strToIntDef(vGTICMS,0) in[0,20,40,41,60,90]) then begin
        vGTICMS := '40';
      end;
    end;

    if ((vGTICMS = '60') or
        (vGTICMS = '41') or
        (vGTICMS = '40') or
        (vGTICMS = '30')) or
       (fEmitenteCRT = '1') then
      vVrBC := 0
    else
      vVrBC := fItens.ValorBaseICMS;

//    vVrBC := IIf(vGTICMS in ['60','41','40'],0,fItens.ValorBase);
//    vVrBC := IIf(vGTICMS = '60',0,fItens.ValorBase);

    vVrTotal := TrocarCaracter(FormatFloat('##########0.00',fItens.ValorTotal),'.',',');
    vQtde    := TrocarCaracter(FormatFloat('##########0.000',fItens.Quantidade),'.',',');
    vVrUnit  := TrocarCaracter(FormatFloat('##########0.0000000000',fItens.ValorUnitario),'.',',');

    vValorDesconto := TrocarCaracter(FormatFloat('##########0.00',fItens.ValorDesconto),'.',',');
    vValorBase := TrocarCaracter(FormatFloat('##########0.00',fItens.ValorBaseICMS),'.',',');
    vAliquota := TrocarCaracter(FormatFloat('##########0.00',fItens.AliquotaICMS),'.',',');
    vValor  := TrocarCaracter(FormatFloat('##########0.00',fItens.ValorICMS),'.',',');
    vRedBC  := TrocarCaracter(FormatFloat('##########0.0000',fItens.PercentualReducao/100),'.',',');

    vqTrib := TrocarCaracter(FormatFloat('##########0.00',fItens.ValorBaseICMS / fItens.ValorTotal),'.',',');

    Result := Result + '[PRODUTO' + Padl(IntToStr(vIndex),3,'0') + ']' +  sLineBreak +
                       'CFOP='          + fItens.CFOP +  sLineBreak +
                       'Codigo='        + fItens.Codigo +  sLineBreak +
                       'Descricao='     + fItens.Descricao +  sLineBreak;

                       if (fItens.EAN <> '') and (Length(fItens.EAN) = 13) then
                             Result := Result + 'EAN=' + fItens.EAN +  sLineBreak;

    //ESTÁ SENDO FEITO EM HOMOLOGAÇÃO, MAS TEM QUE TESTAR EM PRODUÇÃO
//    if fEmitenteCRT <> '1' then BEGIN
      //"I - ao § 1º da cláusula terceira, a partir de:
      //a) 1º de julho de 2017, para a indústria e o importador;
      //b) 1º de outubro de 2017, para o atacadista;
      //c) 1ª de abril de 2018, para os demais segmentos econômicos;".
      if (sAmbiente = 1) OR ((Date >= strtodate('01/04/2018')) AND (sAmbiente = 0)) then begin
      // sAmbiente -  homologação = 1, Produção = 0;
        if StrToIntDef(vGTICMS,60) in [10,30,60,70,90] then begin
          vCest := BuscaCest(fItens.NCM);
          if vCest <> '' then
            Result := Result + 'CEST=' + vCest + sLineBreak;
        end;
      end;
//    end;

    Result := Result +
                       'NCM='           + fItens.NCM +  sLineBreak +
                       'Unidade='       + fItens.Unidade +  sLineBreak +
                       'Quantidade='    + TrocarCaracter(vQtde,',','.') +  sLineBreak +
                       'ValorUnitario=' + TrocarCaracter(vVrUnit,',','.')  +  sLineBreak +
                       'ValorTotal='    + TrocarCaracter(vVrTotal,',','.')  +  sLineBreak +
                       'ValorDesconto=' + TrocarCaracter(vValorDesconto,',','.') +  sLineBreak +
                       'NVE='           + fItens.NVE +  sLineBreak +
                       'nFCI='          + fItens.nFCI +  sLineBreak +
                       'nRECOPI='       + fItens.nRECOPI +  sLineBreak +
                       'pDevol='        + fItens.pDevol +  sLineBreak +
                       'vIPIDevol='     + fItens.vIPIDevol +  sLineBreak +
                       'indTot=1'       + sLineBreak;

    if fEmitenteCRT <> '1' then BEGIN

      Result := Result + '[ICMS'          + Padl(IntToStr(vIndex),3,'0') + ']' +  sLineBreak +
                         'CST='           + vGTICMS + sLineBreak +
                         'ValorBase='     + TrocarCaracter(vValorBase,',','.') + sLineBreak +
                         'Aliquota='      + TrocarCaracter(vAliquota,',','.') + sLineBreak +
                         'Valor='         + TrocarCaracter(vValor,',','.') + sLineBreak +
                         'pRedBC='        + TrocarCaracter(vRedBC,',','.') + sLineBreak;

      if fitens.ValorST > 0 then begin
        Result := Result + 'vBCST=' + TrocarCaracter(FormatFloat('######0.00',fitens.ValorBaseST),',','.')  + sLineBreak +
                           'vICMSST=' + TrocarCaracter(FormatFloat('######0.00',fItens.ValorST),',','.') + sLineBreak;
      end;

    end else if fEmitenteCRT = '1' then BEGIN
      if fModelo = 65 then begin
        ValidaICMS(fItens.CSOSN, fItens.CST, fItens.CFOP, fItens.Codigo, fItens.Descricao);
      END;
        Result := Result + '[ICMS'          + Padl(IntToStr(vIndex),3,'0') + ']' +  sLineBreak +
                           'Origem=' + fitens.CSOSN_ORIGEM + sLineBreak +
                           'CSOSN=' + fItens.CSOSN + sLineBreak;
//      end;
    END;

    if fItens.Combustivel or (CFOPCombustivel(fItens.CFOP)) then begin
      Result := Result + '[Combustivel' + Padl(IntToStr(vIndex),3,'0') + ']' +  sLineBreak +
                         'cProdANP='    + IntToStr(fItens.cProdANP) +  sLineBreak +
                         'UFCons='      + fItens.UFCons +  sLineBreak;

//    fazer teste aqui, mas ainda ESta errado
      if fModelo = 65 then begin  //NT2015-002 LA11-10 entra em vigor em 01/01/2016
        if fItens.Combustivel then begin
          Result := Result + '[encerrante'  + Padl(IntToStr(vIndex),3,'0') + ']' +  sLineBreak +
                             'nBico='       + copy(fItens.Codigo,3,2) + sLineBreak +
                             'nTanque=1'    + sLineBreak +
                             'vEncIni='     + FormatFloat('############.000',fItens.EncInicial) + sLineBreak +
                             'vEncFin='     + FormatFloat('############.000',fItens.EncFinal)   + sLineBreak;
        end;
      end;
    end;

    if bEnviaPisCofinsNFCe then begin
      if ((fItens.PISCST <> '') and (StrToIntDef(fitens.PISCST,0)>0)) and ((fItens.COFINSCST <> '') and (StrToIntDef(fitens.COFINSCST,0)>0)) and (fItens.PISAliquota > 0) and (fItens.PISValor > 0) and (    fItens.COFINSAliquota > 0) and (fItens.COFINSValor > 0) then begin
        Result := Result + '[PIS'+ Padl(IntToStr(vIndex),3,'0') + ']' +  sLineBreak +
                           'CST='  + fItens.PISCST + sLineBreak;
        if StrToIntDef(fitens.PISCST,0) in[1,2] then begin
          Result := Result + 'ValorBase=' + TrocarCaracter(vVrTotal,',','.') + sLineBreak+
                             'Aliquota='  + TrocarCaracter(FormatFloat('######0.00',fitens.PISAliquota),',','.') + sLineBreak+
                             'Valor='     + TrocarCaracter(FormatFloat('######0.00',fitens.PISValor),',','.') +sLineBreak;
        end
        else if StrToIntDef(fitens.PISCST,0) in[3] then begin
          Result := Result + 'Quantidade='    + TrocarCaracter(FormatFloat('######0.00',fitens.Quantidade),',','.') +sLineBreak+
                             'ValorAliquota=' + TrocarCaracter(FormatFloat('######0.00',fitens.PISAliquota),',','.') +sLineBreak+
                             'Valor='         + TrocarCaracter(FormatFloat('######0.00',fitens.PISValor),',','.') +sLineBreak;

        end
        else if StrToIntDef(fitens.PISCST,0) in[4,5,6,7,8,9] then begin
          //Nessa condição envia somente o CST que já foi informado acima.
        end
        else begin
          Result := Result + 'ValorBase='     + TrocarCaracter(vVrTotal,',','.') +sLineBreak+
                             'Aliquota='      + TrocarCaracter(FormatFloat('######0.00',fitens.PISAliquota),',','.') + sLineBreak+
                             'Quantidade='    + TrocarCaracter(FormatFloat('######0.00',fitens.Quantidade),',','.')  + sLineBreak+
                             'ValorAliquota=' + TrocarCaracter(FormatFloat('######0.00',fitens.PISAliquota),',','.') + sLineBreak+
                             'Valor='         + TrocarCaracter(FormatFloat('######0.00',fitens.PISValor),',','.')    + sLineBreak;
        end;

        Result := Result + '[COFINS'+ Padl(IntToStr(vIndex),3,'0') + ']' +  sLineBreak +
                           'CST='  + fItens.COFINSCST + sLineBreak;
        if StrToIntDef(fitens.COFINSCST,0) in[1,2] then begin
          Result := Result + 'ValorBase=' + TrocarCaracter(vVrTotal,',','.') +sLineBreak+
                             'Aliquota='  + TrocarCaracter(FormatFloat('######0.00',fitens.COFINSAliquota),',','.') +sLineBreak+
                             'Valor='     + TrocarCaracter(FormatFloat('######0.00',fitens.COFINSValor),',','.')    + sLineBreak;
        end
        else if StrToIntDef(fitens.COFINSCST,0) in[3] then begin
          Result := Result + 'Quantidade=' + TrocarCaracter(FormatFloat('######0.00',fitens.Quantidade),',','.')     +sLineBreak+
                             'Aliquota='   + TrocarCaracter(FormatFloat('######0.00',fitens.COFINSAliquota),',','.') +sLineBreak+
                             'Valor='      + TrocarCaracter(FormatFloat('######0.00',fitens.COFINSValor),',','.')    +sLineBreak;

        end
        else if StrToIntDef(fitens.COFINSCST,0) in[4,5,6,7,8,9] then begin
          //Nessa condição envia somente o COFINS que já foi informado acima.
        end
        else begin
          Result := Result + 'Valor='         + TrocarCaracter(FormatFloat('######0.00',fitens.COFINSValor),',','.')    +sLineBreak+
                             'Aliquota='      + TrocarCaracter(FormatFloat('######0.00',fitens.COFINSAliquota),',','.') +sLineBreak+
                             'Quantidade='    + TrocarCaracter(FormatFloat('######0.00',fitens.Quantidade),',','.')     +sLineBreak+
                             'ValorAliquota=' + TrocarCaracter(FormatFloat('######0.00',fitens.COFINSAliquota),',','.') +sLineBreak+
                             'ValorBase='     + TrocarCaracter(vVrTotal,',','.') + sLineBreak;
        end;
      end;
    end;

    fValorBaseICMS_ST := fValorBaseICMS_ST + fitens.ValorBaseST;
    fValorICMS_ST     := fValorICMS_ST + fitens.ValorST;

    fTotalBaseICMS     := fTotalBaseICMS + vVrBC;
    fTotalValorICMS    := fTotalValorICMS +  iif(vVrBC > 0,fItens.ValorICMS,0);
    fTotalValorProduto := fTotalValorProduto + fItens.ValorTotal;

    if fitens.ValorST > 0 then
      fTotalValorNota    := fTotalValorNota + fitens.ValorST + fItens.ValorTotal
    else
      fTotalValorNota    := fTotalValorNota + fItens.ValorTotal;


    Inc(vIndex);
  end;
end;

constructor TCapaNFe.create;
begin
  fVersao := '3.10';
  fListaItens  := TObjectList<TItensNFe>.Create;
  fListaPag    := TObjectList<TPagNFe>.Create;
  fListaRefECF := TObjectList<TRefEcf>.Create;
  fListaRefNFCe := TObjectList<TRefNFCe>.create;
  fTemItemCombustivel := false;
end;

procedure TCapaNFe.SetStatus(value: string);
begin
//  Notas não validadas  N
//  Notas enviadas com erros E
//  Notas validadas           V
//  Notas pendentes de envio   P
//  Notas Canceladas

  if Dm.Empresas.FieldByName('CONTINGENCIA').AsString = '2' then begin
    fStatus_Ctg := value;
    fStatus_Nfe := '';
  end else begin
    fStatus_Ctg := '';
    fStatus_Nfe := value;
  end;
end;

procedure TCapaNFe.SetTotalValorNota(pValue: Double);
begin
  fTotalValorNota := pValue;
end;

procedure TCapaNFe.SetTotalValorProduto(pValue: Double);
begin
  fTotalValorProduto := pValue;
end;

function TCapaNFe.BuscaCest(pNCM: string): STRING;
var
  qryCest : TADOQuery;
begin
  try
    qryCest := TADOQuery.Create(nil);
    qryCest.Connection := DM.ADOconexao;

    with qryCest do begin
      Close;
      SQL.Clear;
      sql.Add(' SELECT CEST FROM CEST WHERE NCM_SH = :NCM');
      Parameters.ParamByName('NCM').Value := pNCM;
      Open;
      Result := SomenteNumeros(Trim(FieldByName('CEST').Value));
    end;
  except
    result := '';
  end;
end;

function TcapaNFe.ValidaICMS(pCSOSN, pCST, pCFOP, pCODIGO, pDescricao: STRING): boolean;
begin
  //verifica cfop e csosn e cst
  result := false;
  //NOTA TECNICA 2015-002 CAMPO SEQUENCIA N12-30
  if NOT em(pCST,['000','020','040','041','060','090'])   then begin
    raise exception.Create('Rejeição: Para o item ' + pCODIGO + ' - ' + pDescricao + sLineBreak +
                           'com CST ' + pCST + ' indevido!');
  end;
  //NOTA TECNICA 2015-002 CAMPO SEQUENCIA N12-44
  if (pCST = '060') and (not em(pCFOP,['5405','5656','5667'])) then begin
    raise exception.Create('Rejeição: Para o item ' + pCODIGO + ' - ' + pDescricao + sLineBreak +
                           'O CFOP ' + pCFOP + ' não é permitido para o CST ' + pCST);
  end;
  //NOTA TECNICA 2015-002 CAMPO SEQUENCIA N12-40
  if (not em(pCFOP,['5101','5102','5103','5104','5115'])) and em(pCST,['000','020','040','041','090'])   then begin
    raise exception.Create('Rejeição: Para o item ' + pCODIGO + ' - ' + pDescricao + sLineBreak +
                           'O CFOP ' + pCFOP + ' não é permitido para o CST ' + pCST);
  end;
  //NOTA TECNICA 2015-002 CAMPO SEQUENCIA N12a-20
  if not em(pCSOSN,['102','103','300','400','500','900']) then begin
    raise exception.Create('Rejeição: Para o item ' + pCODIGO + ' - ' + pDescricao + sLineBreak +
                           'com o CSOSN ' + pCSOSN + ' indevido!');
  end;
  //NOTA TECNICA 2015-002 CAMPO SEQUENCIA N12-40a
  if (not em(pCFOP,['5101','5102','5103','5104','5115'])) and em(pCSOSN,['102','103','300','400','900'])   then begin
    raise exception.Create('Rejeição: Para o item ' + pCODIGO + ' - ' + pDescricao + sLineBreak +
                           'O CFOP ' + pCFOP + ' não é permitido para o CSOSN ' + pCSOSN);
  end;
  result := true;
end;

end.
