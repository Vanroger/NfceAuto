unit untImprimeItens;

interface

uses
  StdCtrls, Classes, ADODB, UnitDM, SysUtils, UnitMain, untImpressoraFiscal, untConstante,
  UNTVARIAveis, UnitFuncao, unitIBPT, untFuncoes, Dialogs, ExtCtrls, Types,
  Forms, Graphics, untCampos, DateUtils, untItensImpresso, DB, Windows, NFEInterfaceV3,
  untSplashNFe, pcnAuxiliar, untFpgtoEfetivadas;

type
  TImprimeItens = class
  private
    fqryPLM       : TADOQuery;
    fqryCabecalho : TADOQuery;
    fQryItens     : TADOQuery;
    fQryExiste    : TADOQuery;
    fQryDBFTemp   : TADOQuery;
    fQryStatusCupom : TADOQuery;
    fQryPontos_Cli  : TADOQuery;
    fQryFpgto       : TADOQuery;
    fCampos       : Tcampos;
    fListaItens   : TList;
//    fQTdItemLista : Integer;
    fVrTotal             : Double;
    fFechamentoAbortado  : Boolean;

    fInterfaceNfCe    : TNFEInterfaceV3;

    fDescricaoProduto : string;
    fItem             : Integer;
    fStringList       : TStrings;
    fPlacamov_id      : Integer;

    fQtdePto          : Double;
    fQtdePtoTotal     : Double;

    fEncInicial       : Double;
    fEncFinal         : Double;

    fNUMABASTEC        : Integer;
    fCSOSN             : String;
    fCSOSN_ORIGEM      : String;
    fCSOSN_MODBCICMS   : String;
    fCSOSN_MODBCICMSST : String;

    fCupomAberto      : Boolean;
    fGravouCabecalho  : Boolean;

    fPrecoEspecial    : Boolean;

    fAcbrAtivo        : Boolean;
    fJaIniciouVndNFCe : Boolean;

    fCODPRODUTO      : String;
    fQUANTIDADE      : Double;
    fValorUnitario   : Double;
    fValorUnitarioL  : Double;
    fVALORTOTAL      : Double;
    fVrTotalProdutos : Double;

    fDbfTempID       : integer;
    fArqItem         : Integer;

    fVndTEFfinalizada : Boolean;

    //essa é a soma total pago no fechamento do cupom das formas de pgto
    fVrTotalPago : Double;

    fUtilizandoTef : Boolean;
    fVndUsouTEF    : Boolean; //quando utilizar a forma de pgto cartao
    fCancelando    : Boolean;
    fTextoComprovante : TStringList;
    fMsgComprovante   : String;

    fDescarregaNFCe   : boolean;

    function AtribuiValores(pValorTotal: double = 0): Boolean;
    function ImprimeItem(pCodVendedor, pCodProduto: string; pQuantidade,
      pUnitario, pUnitarioL, pValorTotal: Double;
      var pArredonda: Boolean): Boolean;
    function GravaBasicoArquivo(pNumeroCupom   : AnsiString;
                                pIdNotaFiscal  : Integer;
                                pCodVendedor   : Integer;
                                pIdentificador : String;
                                pCODPRODUTO    : String;
                                pQUANTIDADE    : Double;
                                pValorUnitario : Double;
                                pValorUnitarioL: Double;
                                pVALORTOTAL    : Double;
                                pArredonda     : Boolean;
                                pPlacaMov      : Integer = 0;
                                pDbfTempID     : integer = 0;
                                intNumAbastec  : integer = 0): Boolean;
    function ApagaBasicoArquivo: boolean;
    function RetornaPlacaMov: Boolean;
    function GravaCabecalhoPedido(strBas_NumeroNF  : String;
                                  intIDNotaFiscal  : Integer;
                                  strBas_Tipo      : String;
                                  dtBas_Data       : TDate;
                                  intCodVend       : Integer;
                                  intCodFornec     : Integer;
                                  strCodClie       : String;
                                  intFormaPgto     : Variant;
                                  intPista         : Variant;
                                  intTurno         : Variant;
                                  strCodUser       : String;
                                  intIDMovimento   : Integer;
                                  strSerieECF      : String;
                                  strCFOP          : String;
                                  strSerie         : String;
                                  fltBaseICMSSub   : Real;
                                  fltICMSSUB       : Real;
                                  fltBaseCalcICMS  : Real;
                                  fltBaseICMS      : Real;
                                  strNumeroECF     : String;
                                  HoraEmissao      : TDatetime;
                                  dtCupom          : TDate): Boolean;
    function GravaItensPedidos(intIDNotaFiscal  : Integer;
                               strTabCodigo     : String;
                               intArqItem       : Integer;
                               fltQtde          : Real;
                               intPlacaMov_ID   : Integer;
                               intDBFTEMP_ID    : Integer;
                               strIdentificador : String;
                               fltArqVrUnit     : Real;
                               fltArqVrUnitL    : Real;
                               intNumAbastec    : Integer;
                               fltArqSaldo      : Real;
                               fltArqEstoque    : Real;
                               strArqICMS       : String;
                               intPistaConsumo  : Variant;
                               strTipo          : String  = '';
                               pArredonda       : Boolean = True): Boolean;

    function SubstituiVirgulaPorPontos(sString: String): String;
    procedure ImprimeCupom(var pntPapel: TImage; Texto: TStrings);
    function ApagaItemArquivo(pCombustivel: Boolean; pDocID: integer): Boolean;
    function GetArquivoID(pIdNF: Integer; pCodProduto: string; var pItem: Integer): Integer;
    procedure SetImpressoLista(pCODPRODUTO: string; pPlacamov_id: integer; pItem : Integer);
    procedure MostraPrecoEspecial(var pPrecoEspecial: TLabel);
    function AlteraItemLista(pPlacaMovID: Integer; pUnit,
      pVrtotal: Double): Boolean;
    function VerificaDescontoVendedor: boolean;
    function GravaDBFTemp(pArredonda: integer): Integer;
    function ImprimeItemNFCe(pCodVendedor,
                             pCodProduto    : string;
                             pQuantidade,
                             pUnitario,
                             pUnitarioL,
                             pValorTotal    : Double;
                             var pArredonda : Boolean;
                             pEncInicial    : double = 0;
                             pEncFinal      : double = 0): Boolean;
    function EmituNfe(sEmitiu : String; iNumeroNf : Integer; vChNFe: string = ''; pNumero_Recibo : string = ''; pProtocolo : String = ''): Boolean;
    procedure SetEnderecoXML(pNumeroNf : Integer; pEnderecoXML : string = '');
    function SetChaveNFe(pNumeroNf : Integer; pChaveNfe : string): boolean;
    procedure AlteraNumeroNFe(var pNotaFiscal: Integer);
    function AlterarSTATUS_NFE(valor: string; pNumeroNF: integer): Boolean;
    function AlterarSTATUS_CTG(sEmitiu        : String;
                               iNumeroNf      : Integer;
                               pContingencia  : Integer;
                               pSTATUS_NFE    : string = '';
                               pSTATUS_CTG    : string = '';
             //                  vChNFe         : String = '';
                               pNumero_Recibo : String = '';
                               pProtocolo     : String = ''): Boolean;
    function MonitorAtivo: boolean;
    procedure LimpaTextoComprovante;
    function ValidaCST(var pCST: string; pCFOP: String): boolean;
    procedure AlterarTIPO_DENEGADA(valor: string; pNumeroNF: integer; pModelo,
      pSerie: String);
    function GetUtilizandoTef: Boolean;
    procedure SetUtilizandoTef(const Value: Boolean);
    function ConverteCSOSN(pValue: integer): string;
    function ConverteCSOSN_MODBCICMS(pValue: integer): string;
    function ConverteCSOSN_MODBCICMSST(pValue: integer): string;
    function ConverteCSOSN_ORIGEM(pValue: integer): string;
  public
    constructor Create(pModeloECF: string);
    destructor Destroy;

    procedure SetDestinatario(pDestinatarioCNPJ, pDestinatarioIE,
      pDestinatarioNomeRazao, pDestinatarioFone, pDestinatarioCEP,
      pDestinatarioLogradouro, pDestinatarioNumero, pDestinatarioComplemento,
      pDestinatarioBairro, pDestinatarioCidadeCod, pDestinatarioCidade,
      pDestinatarioUF: string;
      pDestinatarioCodigo: Integer);
    function GravaFormaPagamento: Boolean;
    function SetPtosPremia: Boolean;
    //funçoes para NFCe
    function GravaNFCe(pNFCE: Boolean = False): Boolean;
    function ApagaNFce: Boolean;
    function EnviarOffLine: boolean;
    function CriarEnviarNFCe: Boolean;
    function IniciaVendaNfce(pTpEmiss: Integer = 1): Boolean;
    procedure AddFormaPagNfe(pFormaPag: array of Variant);

    procedure ImprimeImage(var pImage: TImage; pImprimeCabecalho: Boolean = False; pCancelar: Boolean = False);
    function ImprimePlacaMov(var pImage: TImage;
                             pIdentifid: boolean;
                             var pPrecoEspecial: TLabel;
                             pVerificaIdentificadorVenda : Boolean;
                             pDescarrega : Boolean = false): Boolean; //imprime itens do placamov que foram selecionados
    function ImprimeProduto(pCodProduto: string;           //imprime produto chamado pelo F6
                            var pImage: TImage;
                            var pPrecoEspecial: TLabel;
                            pQTDe: Double = 1;
                            pValorTotal: double = 0;
                            pVndCombF6comAutomacao: boolean = false): Boolean;  //quando está configurado com automação mas pode vender combustivel pela automação
    function RetornaDadosAfericao(pEscolha: Integer; pIdNotaFiscal : Integer; pNumero,pMotivoIntervencao,pNomeInterventor,pCNPJInterventora,pCPFTecnico: string): Boolean;
    function AlteraBasTipo(pTipo: string): Boolean;
    function AlteraIdMovimento(pIdMovimento: string): Boolean;
    function AlteraFpgto(pFpgto: string): Boolean;
    function IniciaVenda(pPlacaMovID: string): Boolean;
    function RetornaListaItens: TList;
    procedure SetItemCancelado(pPlacaMovId: Integer = 0; pArquivoId: Integer = 0; pCombustivel: Boolean=False);
    procedure SetItemLista(pItem : Integer; pCodProduto: String; pDescricao: string; pQuantidade, pVrUnit, pVrTotal: double; pArquivoId: Integer);
    procedure RetiraImageCupom(pIndex: Integer; var pImage: TImage);
    function GetCountLista: integer;
    function RetornaVenda: Boolean;
    function TodosItensCancelado: Boolean;
    procedure SetCupomAberto;
    function RetornaCheque: Boolean;
    function RetornaMoviment: Boolean;
    function FinalizaDbfTemp: Boolean;
    function CancelaBasico: Boolean;
    function InsereStatusCupom: Boolean;
    function AlteraStatusCupom(pValue: Integer): Boolean;
    function GetQtdePontos: Double;
    procedure SetTextoComprovante(pTexto : String);
    procedure SetComprovanteNFCe(pLista: TStringList);
    procedure SetMsgComprovante(pValue: String);
    function GetTextoTEF: TStringList;
    procedure SetMensagem(pMensagem : string);
    procedure SetTotTrib(pTotTrib: String);
    function ImprimeComprovante: boolean;
    procedure SetAlteraFonte(pValue: Boolean);
    function AlterarLiberadoAutomatico(pValor: Integer; pNumeroNF: integer): Boolean;
  published
    property CupomAberto      : Boolean     read fCupomAberto      write fCupomAberto;
    property VrTotalProdutos  : Double      read fVrTotalProdutos  write fVrTotalProdutos;
    property VrTotalPago      : Double      read fVrTotalPago      write fVrTotalPago;
    property VrTotal          : Double      read fVrTotal          write fVrTotal; //este campo se refere ao valor total exibido na tela de vendas
    property UtilizandoTef    : Boolean     read GetUtilizandoTef    write SetUtilizandoTef;
    property Cancelando       : Boolean     read fCancelando       write fCancelando;
    property TextoComprovante : TStringList read fTextoComprovante write fTextoComprovante;
    property FechamentoAbortado : Boolean   read fFechamentoAbortado write fFechamentoAbortado;//se o fechamento foi abortado ou não
    property JaIniciouVndNFCe   : Boolean   read fJaIniciouVndNFCe   write fJaIniciouVndNFCe;

  end;

implementation

uses untRelComprovanteNFCe;

function TImprimeItens.VerificaDescontoVendedor: boolean;
var
  sSenhaUsuario : String;
  sSenhaInput   : String;
  i             : Integer;
begin
  try
    sSenhaUsuario := '';
    sSenhaInput   := PasswordInputBox('Senha do Vendedor Autorizado a dar desconto', 'Senha ', True);
    for i := 1 to Length(sSenhaInput) do begin
      sSenhaUsuario := sSenhaUsuario + Criptografa(Copy(sSenhaInput, i, 1));
    end;
    dm.qryGeral.Close;
    dm.qryGeral.SQL.Clear;
    dm.qryGeral.SQL.Add('SELECT');
    dm.qryGeral.SQL.Add('       V.VEN_CODVEND');
    dm.qryGeral.SQL.Add('  FROM VENDEDOR V');
    dm.qryGeral.SQL.Add(' WHERE V.VEN_SENHA    = ' + StringToSql(sSenhaUsuario));
    dm.qryGeral.SQL.Add('   AND V.PERMDESCONTO = 1');
    dm.qryGeral.SQL.Add('   AND V.ATIVO = 0');
    dm.qryGeral.Open;
    if dm.qryGeral.RecordCount < 1 then begin

      if MensagemSimNao('Vendedor não autorizado a conceder desconto!' + slinebreak +
                     'Deseja informar a senha novamente?') = IDYES then begin
        result := VerificaDescontoVendedor
      end
      else
        Result := False;
    end
    else begin
      Result := True;
    end;
  except
    Result := false;
    MensagemErro('Ocorreu um erro na tabela de vendedor, não foi possível conceder desconto!');
  end;
end;

//este método é para verificar se usa classe de preço ou preço especial
//ou se tem desconto no momento da impressao, que deve ser utilizado como
//base o preço vendido na bomba.
//primeiro - classe de preço
//segundo - preço especial
//terceiro - desconto - pelo cadastro do produto
//quarto - preço a prazo - pelo cadastro do produto e cofiguração mestre/vendas/8
function TImprimeItens.AtribuiValores(pValorTotal: double = 0): Boolean;
var
  vDesconto : Double;
  sDesconto : String;
  vUsouPrecoPrazo : Boolean;

  procedure RetValores;
  begin
    if fQUANTIDADE = 1 then begin
      if pValorTotal > 0 then
        fVALORTOTAL := pValorTotal
      else
        fValorTotal := fqryPLM.FieldByName('ValorUnitario').AsFloat;

      fQUANTIDADE     := TruncaFloat(fValortotal / fqryPLM.FieldByName('ValorUnitario').AsFloat,3,false);
      fValorUnitarioL := fqryPLM.FieldByName('ValorUnitario').AsFloat;
      fValorUnitario  := fValorUnitarioL;
    end
    else begin
      fValorUnitarioL := fqryPLM.FieldByName('ValorUnitario').AsFloat;
      if fVALORTOTAL = 0 then
        fVALORTOTAL := fqryPLM.FieldByName('VALORTOTAL').AsFloat;
    end;
  end;
begin
  try
    fValorUnitario  := 0;
    fValorUnitarioL := 0;
    fVALORTOTAL     := 0;
    vUsouPrecoPrazo := false;
    fPlacamov_id    := fqryPLM.FieldByName('PLACAMOV_ID').AsInteger;
    fCODPRODUTO     := fqryPLM.FieldByName('CodProduto').AsString;
    fQUANTIDADE     := fqryPLM.FieldByName('QUANTIDADE').AsFloat;
    fQtdePtoTotal   := fQtdePtoTotal + fqryPLM.FieldByName('QUANTIDADE').AsFloat * fqryPLM.FieldByName('QtdePontos').AsFloat;
    fQtdePto        := fqryPLM.FieldByName('QtdePontos').AsFloat;
    fEncInicial     := fqryPLM.FieldByName('EncInicial').AsFloat;
    fEncFinal       := fqryPLM.FieldByName('EncFinal').AsFloat;
    fNUMABASTEC     := fqryPLM.FieldByName('NUMABASTEC').AsInteger;

    if fqryPLM.FieldByName('USACLASSEPRECO').AsBoolean then begin
      if fqryPLM.FieldByName('VrUntCLASSEPRECO').AsFloat > 0 then begin
        fValorUnitario  := fqryPLM.FieldByName('ValorUnitario').AsFloat;// fqryPLM.FieldByName('VrUntCLASSEPRECO').AsFloat;
        fValorUnitarioL := fqryPLM.FieldByName('VrUntCLASSEPRECO').AsFloat; //fValorUnitario;

        if pValorTotal > 0 then begin
          fVALORTOTAL := pValorTotal;
          fQUANTIDADE := TruncaFloat(pValorTotal / fqryPLM.FieldByName('VrUntCLASSEPRECO').AsFloat,3,False);
        end
        else
          fVALORTOTAL     := fqryPLM.FieldByName('VRTOTALCLASSEPRECO').AsFloat;
        fPrecoEspecial := True;
      end
      else begin
//        fValorUnitario  := fqryPLM.FieldByName('ValorUnitario').AsFloat;
//        fValorUnitarioL := fValorUnitario;

        if pValorTotal > 0 then begin
          fVALORTOTAL := pValorTotal;
          fQUANTIDADE := TruncaFloat(pValorTotal / fqryPLM.FieldByName('ValorUnitario').AsFloat,3,False);
        end
        else
          fVALORTOTAL     := fqryPLM.FieldByName('VALORTOTAL').AsFloat;
        fPrecoEspecial := False;
      end;
    end
    else begin
      if fqryPLM.FieldByName('VrUntPrecoEsp').AsFloat > 0 then begin
        fValorUnitario  := fqryPLM.FieldByName('ValorUnitario').AsFloat;
        fValorUnitarioL := fqryPLM.FieldByName('VrUntPrecoEsp').AsFloat;

        if pValorTotal > 0 then begin
          fVALORTOTAL := pValorTotal;
          fQUANTIDADE := TruncaFloat(pValorTotal / fqryPLM.FieldByName('VrUntPrecoEsp').AsFloat,3,False);
        end
        else
          fVALORTOTAL     := fqryPLM.FieldByName('VRTOTALPRECOESPEC').AsFloat;
        fPrecoEspecial := True;
      end
      else begin
        if fqryPLM.FieldByName('TAB_DESCONTO').AsFloat = 0 then BEGIN
//          fValorUnitario  := fqryPLM.FieldByName('ValorUnitario').AsFloat;
//          fValorUnitarioL := fValorUnitario;
          if pValorTotal > 0 then begin
            fVALORTOTAL := pValorTotal;
            fQUANTIDADE := TruncaFloat(pValorTotal / fqryPLM.FieldByName('ValorUnitario').AsFloat,3,False)
          end
          else
            fVALORTOTAL     := fqryPLM.FieldByName('VALORTOTAL').AsFloat;
        END;
        fPrecoEspecial := False;
      end;
    end;

//    //fazer teste aqui
//    if ((fCampos.PerguntaPgtoCartao) and (fCampos.AplicaPrecoPrazoCartao)) and (fqryPLM.FieldByName('ValorPrazo').AsFloat > 0) then begin
//      fValorUnitarioL := fqryPLM.FieldByName('ValorPrazo').AsFloat;
//      fValorUnitario  := fqryPLM.FieldByName('ValorPrazo').AsFloat;
//      vUsouPrecoPrazo := True;
//      if fQUANTIDADE = 1 then begin
//        if pValorTotal > 0 then
//          fVALORTOTAL := pValorTotal
//        else
//          fValorTotal := fValorUnitarioL;
//
//        fQUANTIDADE := TruncaFloat(fVALORTOTAL / fValorUnitarioL,3,False)
//      end
//      else begin
//        fVALORTOTAL := TruncaFloat(fValorUnitarioL * fQUANTIDADE,2,False);
//      end;
//    end //fim teste
//    else

    if ((fCampos.AplicaPrecoPrazoCliente) and (not fPrecoEspecial)) then begin
      if fqryPLM.FieldByName('ValorPrazo').AsFloat > 0 then begin
        fValorUnitarioL := fqryPLM.FieldByName('ValorPrazo').AsFloat;
        fValorUnitario  := fqryPLM.FieldByName('ValorUnitario').AsFloat;
        vUsouPrecoPrazo := True;
        if fQUANTIDADE = 1 then begin
          if pValorTotal > 0 then
            fVALORTOTAL := pValorTotal
          else
            fValorTotal := fValorUnitarioL;

          fQUANTIDADE := TruncaFloat(fVALORTOTAL / fValorUnitarioL,3,False)
        end
        else begin
          fVALORTOTAL := TruncaFloat(fValorUnitarioL * fQUANTIDADE,2,False);
        end;
      end;
    end;

    if (fValorUnitario <= 0) then begin
      RetValores; //esta chamada aqui somente serve quando as duas condições abaixo for false
//     if ((bPrecoVistaPrazo) and (not fPrecoEspecial)) then begin
      if (not fPrecoEspecial) then begin
        if (not vUsouPrecoPrazo) then begin
          fValorUnitario  := fqryPLM.FieldByName('ValorUnitario').AsFloat;
          vDesconto := fqryPLM.FieldByName('TAB_DESCONTO').AsFloat;

          if bPrecoVistaPrazo and (pValorTotal <= 0)  then begin
            if fqryPLM.FieldByName('ValorPrazo').AsFloat > 0 then begin
              if (MensagemSimNao('Usar preço à vista?', False) = idyes) then begin
                RetValores;
              end
              else begin
                fValorUnitarioL := fqryPLM.FieldByName('ValorPrazo').AsFloat;
                vUsouPrecoPrazo := True;
                if fQUANTIDADE = 1 then begin
                  if pValorTotal > 0 then
                    fVALORTOTAL := pValorTotal
                  else
                    fValorTotal := fValorUnitarioL;

                  fQUANTIDADE := TruncaFloat(fVALORTOTAL / fValorUnitarioL,3,False)
                end
                else begin
                  fVALORTOTAL := TruncaFloat(fValorUnitarioL * fQUANTIDADE,2,False);
                end;
              end;
            end
            else begin
              RetValores;
            end;
          end;

          if (vDesconto > 0) and (not fCampos.AplicaPrecoPrazoCliente) and (pValorTotal <= 0) and (not fDescarregaNFCe) then begin
            if Questionar('Deseja efetuar DESCONTO no valor do produto?')  then begin
              if VerificaDescontoVendedor then begin
                repeat
                  if (sTipoDesconto = '$') then begin
                    if not InputQuery('Valor do Desconto', 'Desconto (R$) ', sDesconto) then begin
                      result := true;
                      RetValores;
                      Exit;
                    end;
                  end else begin
                    if not InputQuery('Percentual do Desconto', 'Desconto (%)', sDesconto) then begin
                      Result := True;
                      RetValores;
                      exit;
                    end;
                  end;

                  if (StrToFloat(sDesconto) > vDesconto) then
                    MSGAtencao('A T E N Ç Ã O ! ! !'+ slineBreak +'Desconto maior que o permitido!');

                until (vDesconto >= StrToFloat(sDesconto)) ;

                if (sTipoDesconto = '$') then begin
                  fValorUnitarioL := TruncaFloat(IIf(vUsouPrecoPrazo,fValorUnitarioL,fValorUnitario) - StrToFloat(sDesconto),3,False);
                  if fQUANTIDADE = 1 then begin
                    if pValorTotal > 0 then
                      fVALORTOTAL := pValorTotal
                    else
                      fValorTotal := fValorUnitarioL;

                    fQUANTIDADE := TruncaFloat(fVALORTOTAL / fValorUnitarioL,3,False)
                  end
                  else
                    fVALORTOTAL := TruncaFloat(fValorUnitarioL * fQUANTIDADE,2,False);
                end else begin
                  If vUsouPrecoPrazo then
                    fValorUnitarioL := TruncaFloat(fValorUnitarioL - ((fValorUnitarioL * StrToFloat(sDesconto))/100),3,false)
                  else
                    fValorUnitarioL := TruncaFloat(fValorUnitario - ((fValorUnitario * StrToFloat(sDesconto))/100),3,false);
                  if fQUANTIDADE = 1 then begin
                    if pValorTotal > 0 then
                      fVALORTOTAL := pValorTotal
                    else
                      fValorTotal := fValorUnitarioL;

                    fQUANTIDADE := TruncaFloat(fVALORTOTAL / fValorUnitarioL,3,False)
                  end
                  else
                    fVALORTOTAL := TruncaFloat(fValorUnitarioL * fQUANTIDADE,2,False);
                end;
              end
              else begin
                if not vUsouPrecoPrazo then begin
                  RetValores;
                end;
              end;
            end
            else begin
              if not vUsouPrecoPrazo then begin
                RetValores;
              end;
            end;
          end
          else begin
            if not vUsouPrecoPrazo then begin
              fValorUnitarioL := fqryPLM.FieldByName('ValorUnitario').AsFloat;
              if fQUANTIDADE = 1 then begin
                if pValorTotal > 0 then
                  fVALORTOTAL := pValorTotal
                else
                  fValorTotal := fValorUnitarioL;

                fQUANTIDADE := TruncaFloat(fVALORTOTAL / fValorUnitarioL,3,False)
              end
              else if fVALORTOTAL <= 0 then
                fVALORTOTAL := TruncaFloat(fValorUnitarioL * fQUANTIDADE,2,False);
            end;
          end;
        end;
      end;
    end;
    Result := True;
  except
    result := false;
  end;
end;

{ESTE METÓDO É RESPONSAVEL POR IMPRIMIR OS COMBUSTIVEIS DO PLACAMOV SELECIONADOS}
function TImprimeItens.ImprimePlacaMov(var pImage: TImage;
                                           pIdentifid: boolean;
                                       var pPrecoEspecial: TLabel;
                                           pVerificaIdentificadorVenda : Boolean;
                                           pDescarrega : Boolean = false): Boolean;
var
  pArredonda  : Boolean;
  NomeVendedor : String;
  DeuErro      : boolean;
  vItem        : integer;
  vArquivoId   : integer;
  vImprimiuItem : Boolean;
begin
  try
    DeuErro := false;
    fDescarregaNFCe := pDescarrega;
    with fqryPLM do begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT ');
      SQL.ADD('      QtdePontos = COALESCE(ROUND((SELECT T.QtdePontos ');
      SQL.Add('                                     FROM TABELA  T ');
      SQL.Add('                                    WHERE T.TAB_CODIGO  = P.CodProduto AND T.GRU_CODIGO = ''040''),3,1),0), ');
      SQL.Add('      USACLASSEPRECO = (select USARCLASSESPRECOS from config), ');
      SQL.Add('      P.CodProduto, ');

      if not pDescarrega then begin
        SQL.ADD('      1 as NUMABASTEC,');
        SQL.Add('      P.PLACAMOV_ID, ');
        SQL.Add('      (p.Quantidade) as QUANTIDADE, ');
        SQL.Add('      p.Encerrante - p.Quantidade as EncInicial, ');
        SQL.Add('      p.Encerrante as EncFinal, ');
      end
      else begin
        SQL.ADD('      count(p.codproduto) as NUMABASTEC,');
        SQL.Add('      0 AS PLACAMOV_ID,  ');
        SQL.Add('      sum(p.Quantidade) as QUANTIDADE,  ');
        SQL.Add('      max(p.Encerrante) - sum(p.Quantidade) as EncInicial,  ');
        SQL.Add('      MAX(p.Encerrante) as EncFinal,  ');
      end;
      SQL.Add('      p.ValorUnitario, ');
      SQL.ADD('      TAB_DESCONTO = COALESCE(ROUND((SELECT T.TAB_DESCONTO ');
      SQL.Add('                               FROM TABELA  T ');
      SQL.Add('                              WHERE T.TAB_CODIGO  = P.CodProduto AND T.GRU_CODIGO = ''040''),3,1),0), ');
      SQL.ADD('      ValorPrazo = COALESCE(ROUND((SELECT T.TAB_VRUNITP ');
      SQL.Add('                               FROM TABELA  T ');
      SQL.Add('                              WHERE T.TAB_CODIGO  = P.CodProduto AND T.GRU_CODIGO = ''040''),3,1),0), ');
      SQL.Add('     VrUntCLASSEPRECO = COALESCE(ROUND((SELECT UNITARIO = CASE WHEN UNITARIO > 0 THEN ');
      SQL.Add('                                                                    UNITARIO          ');
      SQL.Add('                                                          ELSE                        ');
      SQL.Add('                                                               round(((100-PERCDESC)/100) * (SELECT pL.ValorUnitario ');
      SQL.Add('                                                                                         FROM PLACAMOV PL      ');
      SQL.Add('                                                                                        WHERE PL.CupomFecha = 0 ');
      SQL.Add('                                                                                          AND PL.CupomFiscal = ' + IntToStr(StrtoInt(fCampos.CupomFiscal)));
      SQL.Add('                                                                                          AND PL.IdNotaFiscal = ' + IntToStr(fCampos.IdNotaFiscal) );
      SQL.Add('                                                                                          AND PL.CodProduto = P.CodProduto ');
      SQL.Add('                                                                                     GROUP BY PL.CodProduto, PL.ValorUnitario),3,0)  ');
      SQL.Add('                                                          END ');
      SQL.Add('                                         FROM PROCLASSES ');
      SQL.Add('                                        WHERE PRODUTO_COD = P.CODPRODUTO ');
      SQL.Add('                                          AND CLASSES_ID =  (SELECT CLASSES_ID ');
      SQL.Add('                                                               FROM CLASSES    ');
      SQL.Add('                                                              WHERE CODIGO = (SELECT CLASSES_ID ');
      SQL.Add('                                                                                FROM CLIENTE ');
      SQL.Add('                                                                               WHERE CLI_CODI = ' + StringToSql(fCampos.CODCliente) + ' ))),3,1),0), ');
      SQL.Add('     VrUntPrecoEsp =  COALESCE(ROUND((SELECT PRE_VRUNIT FROM CLIPRECO ');
      SQL.Add('                                                     WHERE TAB_CODIGO = P.CODPRODUTO ');
      SQL.Add('                                                       AND CLI_CODI = ' + StringToSql(fCampos.CODCliente) + ' ),3,1),0), ');

      if not pDescarrega then
        SQL.Add('     VALORTOTAL  = ROUND((p.ValorTotal),3,1), ')
      else
        SQL.Add('     VALORTOTAL  = ROUND(sum(p.ValorTotal),3,1), ');

      SQL.Add('     VRTOTALCLASSEPRECO = ROUND(COALESCE((SELECT UNITARIO = CASE WHEN UNITARIO > 0 THEN ');
      SQL.Add('                                                                      UNITARIO ');
      SQL.Add('                                                            ELSE ');
      SQL.Add('                                                               round(((100-PERCDESC)/100) * (SELECT pL.ValorUnitario ');
      SQL.Add('                                                                                         FROM PLACAMOV PL      ');
      SQL.Add('                                                                                        WHERE PL.CupomFecha = 0 ');
      SQL.Add('                                                                                          AND PL.CupomFiscal = ' + IntToStr(StrtoInt(fCampos.CupomFiscal)));
      SQL.Add('                                                                                          AND PL.IdNotaFiscal = ' + IntToStr(fCampos.IdNotaFiscal) );
      SQL.Add('                                                                                          AND PL.CodProduto = P.CodProduto ');
      SQL.Add('                                                                                     GROUP BY PL.CodProduto, PL.ValorUnitario),3,0)  ');
      SQL.Add('                                                             END ');
      SQL.Add('                                                      FROM PROCLASSES ');
      SQL.Add('                                                     WHERE PRODUTO_COD = P.CODPRODUTO ');
      SQL.Add('                                                       AND CLASSES_ID = (SELECT CLASSES_ID ');
      SQL.Add('                                                                           FROM CLASSES    ');
      SQL.Add('                                                                          WHERE CODIGO =(SELECT CLASSES_ID ');
      SQL.Add('                                                                                           FROM CLIENTE ');

      if not pDescarrega then
        SQL.Add('                                                                                        WHERE CLI_CODI = ' + StringToSql(fCampos.CODCliente) + ' ))),0) * P.Quantidade,2,0), ')
      else
        SQL.Add('                                                                                        WHERE CLI_CODI = ' + StringToSql(fCampos.CODCliente) + ' ))),0) * sum(P.Quantidade),2,0), ');

      SQL.Add('     VRTOTALPRECOESPEC =  round(COALESCE((SELECT PRE_VRUNIT FROM CLIPRECO ');
      SQL.Add('                                                     WHERE TAB_CODIGO = P.CODPRODUTO ');

      if not pDescarrega then
        SQL.Add('                                                       AND CLI_CODI = ' + StringToSql(fCampos.CODCliente) + ' ),0) * P.Quantidade,3,1) ')
      else
        SQL.Add('                                                       AND CLI_CODI = ' + StringToSql(fCampos.CODCliente) + ' ),0) * sum(P.Quantidade),3,1) ');

      SQL.Add(' From placamov p ');

      if pIdentifid and pVerificaIdentificadorVenda then
        SQL.Add(' , vendedor v ');

      SQL.Add(' where ');

      if pIdentifid and pVerificaIdentificadorVenda then
        SQL.Add(' p.IDENTIFICADOR = v.IDENTIFICADOR and');

      SQL.Add('     p.CupomFecha   = 0');
      SQL.Add(' and p.CupomFiscal  = ' + IntToStr(StrtoInt(fCampos.CupomFiscal)));
      SQL.Add(' and p.IdNotaFiscal = ' + IntToStr(fCampos.IdNotaFiscal) );

      if pIdentifid and pVerificaIdentificadorVenda then
        SQL.Add(' and v.VEN_CODVEND  = ' + fCampos.CodVendedor);

      SQL.Add(' AND P.DATACAIXA    = ' + DateToSql(fCampos.DataCaixa));

      if pDescarrega then
        SQL.Add(' group by p.CodProduto, p.ValorUnitario');

      SQL.Add(' order by p.CodProduto');
      Open;
      First;
    end;

    if not fqryPLM.IsEmpty then begin
      if bUtilizaNFCE then begin
        IniciaVendaNfce;
      end;
      while not fqryPLM.Eof do begin
        if AtribuiValores then begin
          MostraPrecoEspecial(pPrecoEspecial);
          if not bUtilizaNFCE then
            vImprimiuItem := ImprimeItem(fCampos.CodVendedor,
                                         fCODPRODUTO,
                                         fQUANTIDADE,
                                         fValorUnitario,
                                         fValorUnitarioL,
                                         fVALORTOTAL,
                                         pArredonda)
          else
            vImprimiuItem := ImprimeItemNFCe(fCampos.CodVendedor,
                                             fCODPRODUTO,
                                             fQUANTIDADE,
                                             fValorUnitario,
                                             fValorUnitarioL,
                                             fVALORTOTAL,
                                             pArredonda,
                                             fEncInicial,
                                             fEncFinal);
          if vImprimiuItem then begin
            ImprimeImage(pImage);
            pImage.Refresh;
            AlteraItemLista(fPlacamov_id,fValorUnitarioL,fVALORTOTAL);
            fDbfTempID := GravaDBFTemp(iiF(pArredonda,0,1));
            if not GravaBasicoArquivo(fCampos.CupomFiscal,
                                      fCampos.IdNotaFiscal,
                                      Strtoint(fCampos.CodVendedor),
                                      fCampos.Identificador,
                                      fCODPRODUTO,
                                      fQUANTIDADE,
                                      fValorUnitario,
                                      fValorUnitarioL,
                                      fVALORTOTAL,
                                      pArredonda,
                                      fPlacamov_id,
                                      fDbfTempID,
                                      fNUMABASTEC) then begin
              DeuErro := true;
              break;
            end;

            vArquivoId := GetArquivoID(fCampos.IdNotaFiscal,
                                       Copy(Trim(fCODPRODUTO),1,20),
                                       vItem);

            SetItemLista(vItem,
                         fCODPRODUTO,
                         fDescricaoProduto,
                         TruncaFloat(fQUANTIDADE,3,false),
                         fValorUnitarioL,
                         fVALORTOTAL,
                         vArquivoId);

            SetImpressoLista(fCODPRODUTO,fPlacamov_id, vItem);

          end
          else begin
            DeuErro := true;
            Break;
          end;
        end
        else begin
          DeuErro := true;
          Break;
        end;
        fqryPLM.next;
      end;

      if not DeuErro then begin
        result := true;
//        if fStringList.Count > 0 then
//          fStringList.clear;
      end
      else begin
        //caso já tenha impresso mais de um item, então terá que
        //apagar o basico e arquivo, com isso o registros do placamov
        //serão alterados por uma trigger da tabela arquivo
        ApagaBasicoArquivo;
        RetornaPlacaMov;
        //esta chamada a função abaixo verifica se a ECF pode cancelar o
        //cupom, caso possa o parametro tem que estar setado para
        //true para cancelar o cupom fiscal aberto ou o ultimo impresso
        ImpressoraFiscal.CupomFiscalAberto(true);
        result := false;
      end;
    end;
  except
    result := false;
  end;
end;

function TImprimeItens.SetPtosPremia: Boolean;
begin
  if bUTILIZA_PREMIACAO then begin
    with fQryPontos_Cli do begin
      result := false;
      if (fCampos.CNPJCpf = '') or (fCampos.IdNotaFiscal <= 0) or (fCampos.CupomFiscal = '') or
         (fQtdePtoTotal <= 0) or (fCODPRODUTO = '') or (sCPFCNPJClientePadrao = fCampos.CNPJCpf) then
        Exit;

      try
        close;
        sql.Clear;
        sql.Add(' INSERT INTO PONTOS_CLI(CNPJ_CPF, ');
        sql.Add('                        IDNOTAFISCAL, ');
        sql.Add('                        CUPOMFISCAL, ');
        sql.Add('                        DATA_EXPIRACAO, ');
        sql.Add('                        DATA_CADASTRO, ');
        sql.Add('                        QTDE_PONTOS, ');
        sql.Add('                        VEN_CODVEND_VENDA)');
        sql.Add('                 VALUES(:CNPJ_CPF, ');
        sql.Add('                        :IDNOTAFISCAL, ');
        sql.Add('                        :CUPOMFISCAL, ');
        sql.Add('                        :DATA_EXPIRACAO, ');
        sql.Add('                        :DATA_CADASTRO, ');
        sql.Add('                        :QTDE_PONTOS, ');
        sql.Add('                        :VEN_CODVEND_VENDA)');
        Parameters.ParamByName('CNPJ_CPF').Value          := fCampos.CNPJCpf;
        Parameters.ParamByName('IDNOTAFISCAL').Value      := fCampos.IdNotaFiscal;
        Parameters.ParamByName('CUPOMFISCAL').Value       := fCampos.CupomFiscal;
        Parameters.ParamByName('DATA_EXPIRACAO').Value    := FormatDateTime('DD/MM/YYYY',Now + bDiasExpirarPtosPremia);
        Parameters.ParamByName('DATA_CADASTRO').Value     := now;
        Parameters.ParamByName('QTDE_PONTOS').Value       := fQtdePtoTotal;
        Parameters.ParamByName('VEN_CODVEND_VENDA').Value := Strtoint(fCampos.CodVendedor);
        ExecSQL;
        result := true;
        fQtdePtoTotal := 0;
      except
        on e: exception do begin
          result := False;
        end;
      end;
    end;
  end;
end;

procedure TImprimeItens.SetTextoComprovante(pTexto: String);
begin
  fTextoComprovante.Add(pTexto);
end;

procedure TImprimeItens.SetComprovanteNFCe(pLista : TStringList);
begin
  if fTextoComprovante = nil then
    fTextoComprovante := TStringList.Create;

  fTextoComprovante := pLista;
end;

function TImprimeItens.FinalizaDbfTemp: Boolean;
begin
  try
    if fQryDBFTemp = nil then begin
      fQryDBFTemp := TADOQuery.Create(nil);
      fQryDBFTemp.Connection := DM.ADOconexao;
    end;

    with fQryDBFTemp do begin
      Close;
      SQL.Clear;
      SQL.Add(' SELECT COUNT(QTDE)');
      SQL.Add('   FROM DBFTEMP' );
      SQL.Add(' WHERE CUPOMFISCAL = :CF ');
      SQL.Add('   AND TMP_SERIAL = :NOMEPC');
      Parameters.ParamByName('CF').Value := StrtoInt(fCampos.CupomFiscal);
      Parameters.ParamByName('NOMEPC').Value := sNomeComputador;
      open;

      if not IsEmpty then begin
        Close;
        SQL.Clear;
        SQL.Add(' DELETE DBFTEMP' );
        SQL.Add(' WHERE CUPOMFISCAL = :CF ');
        SQL.Add('   AND TMP_SERIAL = :NOMEPC');
        SQL.Add('   AND TMP_PISTA2 = :PISTA ');
        SQL.Add('   AND TMP_TURNO2 = :TURNO ');
        Parameters.ParamByName('CF').Value := StrtoInt(fCampos.CupomFiscal);
        Parameters.ParamByName('NOMEPC').Value := sNomeComputador;
        Parameters.parambyName('PISTA').Value  := IntToStr(fCampos.Pista);
        Parameters.ParamByName('TURNO').Value  := IntToStr(fCampos.Turno);
        ExecSQL;
      end;
    end;
    Result := true;
  except
    on e: exception do begin
      result := False;
//      ShowMessage(e.Message);
    end;
  end;
end;

function TImprimeItens.GravaDBFTemp(pArredonda: integer): integer;
begin
  try
    Result := 0;
    if fQryDBFTemp = nil then
      fQryDBFTemp := TADOQuery.Create(nil);

    fQryDBFTemp.Connection := DM.ADOconexao;

    //Setando o sequencial dos Itens
    with fQryItens do begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT COALESCE(MAX(ARQ_ITEM),0) + 1 AS CONT');
      SQL.Add('  FROM ARQUIVO');
      SQL.Add(' WHERE IDNOTAFISCAL = ' + IntToStr(fCampos.IDNotaFiscal));
      Open;
      fArqItem := FieldByName('CONT').AsInteger;
    end;

    with fQryDBFTemp do begin
      Close;
      sql.Clear;
      SQL.Add('INSERT INTO DBFTEMP(TMP_ORDS2,          ');  //  ARQ_ITEM,');
      SQL.Add('                    TMP_SERIAL,         ');
      SQL.Add('                    TMP_CODI2,          ');  //  TAB_CODIGO,');
      SQL.Add('                    TMP_DESC2,          ');  //  PRODUTO,');
      SQL.Add('                    TMP_QTDE2,          ');  //  ARQ_QTDE,');
      SQL.Add('                    TMP_UNIT2,          ');  //  ARQ_VRUNITL,');
      SQL.Add('                    TMP_DESCO2,         ');
      SQL.Add('                    TMP_VALO2,          ');
      SQL.Add('                    TMP_ITEMCANCELADO,  ');
      SQL.Add('                    CUPOMFISCAL,        ');  //  CUPOMFISCAL,');
      SQL.Add('                    ItemECF,            ');
      SQL.Add('                    EmitidoItem,        ');
      SQL.Add('                    round,              ');
      SQL.Add('                    TMP_VEND2,          ');  //  VEN_CODVEND,');
      SQL.Add('                    TMP_CLI2,           ');  //  CLI_CODI,');
      SQL.Add('                    TMP_PISTA2,         ');  //  PIS_CODIGO,');
      SQL.Add('                    TMP_TURNO2,         ');  //  TUR_TURNO
      SQL.Add('                    PLACAMOV_ID,        ');  //  PLACAMOVID
      SQL.Add('                    IDENTIFICADOR,      ');  //  IDENTIFICADOR
      SQL.Add('                    TMP_SALDO2,         ');  //  ARQ_SALDO
      SQL.Add('                    TMP_ESTOQ2,         ');  //  ARQ_ESTOQ
      SQL.Add('                    TMP_ALIQ2,          ');  //  ARQ_ICMS
      SQL.Add('                    Qtde)               ');
//      SQL.Add(' VALUES( ' + Copy(IntToStr(fArqItem),1,8) );
//      SQL.Add(',' + Copy(sNomeComputador,1,19) );
//      SQL.Add(',' + Copy(fCODPRODUTO,1,15) );
//      SQL.Add(',' + Copy(fDescricaoProduto,1,32) );
//      SQL.Add(',' + FormatFloat('######0.00',fQUANTIDADE));
//      SQL.Add(',' + FormatFloat('######0.00',fValorUnitarioL));
//      SQL.Add(',' + FormatFloat('######0.00',fValorUnitario - fValorUnitarioL));
//      SQL.Add(',' + '0');
//      SQL.Add(',' + IntToStr(0));
//      SQL.Add(',' + IntToStr(StrToIntDef(fCampos.CupomFiscal,0)));
//      SQL.Add(',' + IntToStr(fArqItem));
//      SQL.Add(',' + IntToStr(1));
//      SQL.Add(',' + IntToStr(pArredonda));
//      SQL.Add(',' + fCampos.CodVendedor);
//      SQL.Add(',' + fCampos.CodCliente);
//      SQL.Add(',' + IntToStr(fCampos.Pista));
//      SQL.Add(',' + IntToStr(fCampos.Turno));
//      SQL.Add(',' + IntToStr(fPlacamov_id));
//      SQL.Add(',' + copy(fCampos.Identificador,1,20));
//      SQL.Add(',' + '0');
//      SQL.Add(',' + '0');
//      SQL.Add(',' + '0');
//      SQL.Add(',' + FormatFloat('######0.000',fQUANTIDADE)+ ')');
      SQL.Add(' VALUES(            :TMP_ORDS2,         ');
      SQL.Add('                    :TMP_SERIAL,        ');
      SQL.Add('                    :TMP_CODI2,         ');
      SQL.Add('                    :TMP_DESC2,         ');
      SQL.Add('                    :TMP_QTDE2,         ');
      SQL.Add('                    :TMP_UNIT2,         ');
      SQL.Add('                    :TMP_DESCO2,        ');
      SQL.Add('                    :TMP_VALO2,         ');
      SQL.Add('                    :TMP_ITEMCANCELADO, ');
      SQL.Add('                    :CUPOMFISCAL,       ');
      SQL.Add('                    :ItemECF,           ');
      SQL.Add('                    :EmitidoItem,       ');
      SQL.Add('                    :round,             ');
      SQL.Add('                    :TMP_VEND2,         ');
      SQL.Add('                    :TMP_CLI2,          ');
      SQL.Add('                    :TMP_PISTA2,        ');
      SQL.Add('                    :TMP_TURNO2,        ');
      SQL.Add('                    :PLACAMOV_ID,       ');
      SQL.Add('                    :IDENTIFICADOR,     ');
      SQL.Add('                    :TMP_SALDO2,        ');
      SQL.Add('                    :TMP_ESTOQ2,        ');
      SQL.Add('                    :TMP_ALIQ2,         ');
      SQL.Add('                    :Qtde)              ');
      Parameters.ParamByName('TMP_ORDS2').Value  :=  Copy(IntToStr(fArqItem),1,8);
      Parameters.ParamByName('TMP_SERIAL').Value :=  Copy(sNomeComputador,1,19);
      Parameters.ParamByName('TMP_CODI2').Value  :=  Copy(fCODPRODUTO,1,15);
      Parameters.ParamByName('TMP_DESC2').value  :=  Copy(fDescricaoProduto,1,32);
      Parameters.ParamByName('TMP_QTDE2').value  :=  FormatFloat('######0.00',fQUANTIDADE);
      Parameters.ParamByName('TMP_UNIT2').value  :=  FormatFloat('######0.00',fValorUnitarioL);
      Parameters.ParamByName('TMP_DESCO2').value :=  FormatFloat('######0.00',fValorUnitario - fValorUnitarioL);
      Parameters.ParamByName('TMP_VALO2').value  :=  FormatFloat('######0.00',0);
      Parameters.ParamByName('TMP_ITEMCANCELADO').value := 0;
      Parameters.ParamByName('CUPOMFISCAL').value       := StrToIntDef(fCampos.CupomFiscal,0);
      Parameters.ParamByName('ItemECF').value           := fArqItem;
      Parameters.ParamByName('EmitidoItem').value       := 1;
      Parameters.ParamByName('round').value             := pArredonda;
      Parameters.ParamByName('TMP_VEND2').value         := Copy(fCampos.CodVendedor,1,2);
      Parameters.ParamByName('TMP_CLI2').value          := Copy(fCampos.CodCliente,1,10);
      Parameters.ParamByName('TMP_PISTA2').value        := IntToStr(fCampos.Pista);
      Parameters.ParamByName('TMP_TURNO2').value        := IntToStr(fCampos.Turno);
      Parameters.ParamByName('PLACAMOV_ID').value       := fPlacamov_id;
      Parameters.ParamByName('IDENTIFICADOR').value     := copy(fCampos.Identificador,1,20);
      Parameters.ParamByName('TMP_SALDO2').value        := '0';
      Parameters.ParamByName('TMP_ESTOQ2').value        := '0';
      Parameters.ParamByName('TMP_ALIQ2').value         := '0';
      Parameters.ParamByName('Qtde').value              := fQUANTIDADE;
      ExecSQL;

    end;

    with fQryItens do begin
      Close;
      sql.Clear;
      sql.Add('SELECT MAX(DBFTEMP_ID) AS DBFTEMP_ID');
      SQL.Add('  FROM DBFTEMP ');
      SQL.Add(' WHERE CUPOMFISCAL = :CF ');
      SQL.Add('   AND TMP_SERIAL = :NOMEPC');
      Parameters.ParamByName('CF').Value := StrtoInt(fCampos.CupomFiscal);
      Parameters.ParamByName('NOMEPC').Value := sNomeComputador;
      open;

      if not IsEmpty then
        Result := FieldByName('DBFTEMP_ID').AsInteger;

    end;

  except
    on e: exception do begin
      result := 0;
      ShowMessage(e.Message);
    end;
  end;
end;

function TImprimeItens.GravaFormaPagamento: Boolean;
var
  i : integer;
//  fpgto : TFpgtoEfetivadas;
begin
  try
//    fpgto := TFpgtoEfetivadas.Create;
    with fQryFpgto do begin
      for i := 0 to fCampos.FpgtoLista.Count -1 do begin
        close;
        sql.clear;
        sql.add(' INSERT INTO FPGTO_EFETIVADAS(DOC_CODIGO, IDNOTAFISCAL, VALOR) ');
        SQL.ADD(' VALUES (:DOC_CODIGO, :IDNOTAFISCAL, :VALOR)');
        Parameters.ParamByName('DOC_CODIGO').Value := TFpgtoEfetivadas(fCampos.FpgtoLista[i]).DOC_CODIGO;
        Parameters.ParamByName('IDNOTAFISCAL').Value := TFpgtoEfetivadas(fCampos.FpgtoLista[i]).IDNOTAFISCAL;
        Parameters.ParamByName('VALOR').Value := TFpgtoEfetivadas(fCampos.FpgtoLista[i]).VALOR;
        ExecSQL;
        result := true;
      end;
    end;
  except
    result := false;
  end;
end;

function TImprimeItens.TodosItensCancelado: Boolean;
var
  i : integer;
begin
  Result := True;
  for I := 0 to fListaItens.Count - 1 do begin
    if not (TItensImpresso(fListaItens[i]).Cancelado) then begin
      Result := false;
      Break;
    end;
  end;
end;

procedure TImprimeItens.SetCupomAberto;
begin
  if bUtilizaNFCE then
    fCupomAberto := true
  else
    fCupomAberto := ImpressoraFiscal.CupomFiscalAberto(False);
end;

procedure TImprimeItens.SetImpressoLista(pCODPRODUTO : string;
                                         pPlacamov_id  : integer;
                                         pItem : Integer);
var
  i : integer;
begin
  for I := 0 to fListaItens.Count - 1 do begin
    if (TItensImpresso(fListaItens[i]).Codigo = pCodProduto) and
       (TItensImpresso(fListaItens[i]).PlacaMovId = pPlacamov_id) then begin
      TItensImpresso(fListaItens[i]).impresso := true;
      TItensImpresso(fListaItens[i]).ItemECF  := pItem;
      TItensImpresso(fListaItens[i]).QtdPto   := fQtdePto;
    end;
  end;
end;

function TImprimeItens.RetornaDadosAfericao(pEscolha: Integer;
                                            pIdNotaFiscal : Integer;
                                            pNumero,
                                            pMotivoIntervencao,
                                            pNomeInterventor,
                                            pCNPJInterventora,
                                            pCPFTecnico: string): Boolean;
begin
  try
    if pEscolha = 88 then begin
      with fqryCabecalho do begin
        Close;
        SQL.Clear;
        SQL.Add('UPDATE BASICO');
        SQL.Add('   SET NUM_INTERV   = ''' + pNumero                       + '''');
        SQL.Add('     ,TAF_FORMAPGTO = 88');
        SQL.Add('      ,MOT_INTERV   = ''' + pMotivoIntervencao            + '''');
        SQL.Add('      ,NOM_INTERV   = ''' + pNomeInterventor              + '''');
        SQL.Add('      ,CNPJ_INTERV  = ''' + SoNumeros(pCNPJInterventora)  + '''');
        SQL.Add('      ,CPF_INTERV   = ''' + SoNumeros(pCPFTecnico)        + '''');
        SQL.Add(' WHERE IDNOTAFISCAL = '   + IntegerSQL(pIdNotaFiscal));
        SQL.Add('   AND BAS_DATA     = '   + DateToSql(DateOf(dDataCaixa)));
        SQL.Add('   AND PIS_CODIGO   = 1');
        SQL.Add('   AND TUR_TURNO    = ' + IntToStr(iTurno));
        SQL.Add('   AND BAS_TIPO     = ''D''');
        ExecSQL;
      end;
    end;
    Result := True;
  except
    Result := False;
  end;
end;

function TImprimeItens.RetornaListaItens: TList;
begin
  if Assigned(fListaItens) then
    Result := fListaItens
  else
    Result := TList.Create;
end;

{ESTE METÓDO É RESPONSAVEL POR IMPRIMIR O PRODUTO INFORMADO PELO F6}
function TImprimeItens.ImprimeProduto(pCodProduto: string;
                                      var pImage: TImage;
                                      var pPrecoEspecial: TLabel;
                                      pQTDe: Double = 1;
                                      pValorTotal: double = 0;
                                      pVndCombF6comAutomacao: boolean = false): Boolean;
var
  pArredonda  : Boolean;
  NomeVendedor : String;
  DeuErro      : boolean;
  vItem        : integer;
  vArquivoId   : integer;
  sQtde        : string;
  vImprimiuItem : Boolean;
begin
  try
    DeuErro := false;
    sQtde := TrocarCaracter(FloatToStr(pQTDe),',','.');
    with fqryPLM do begin
      Close;
      SQL.Clear;
      SQL.Add('     SELECT ');
      SQL.ADD('            T.QtdePontos, ');
      SQL.Add('            USACLASSEPRECO = (select USARCLASSESPRECOS from config),    ');
      SQL.ADD('            0 as NUMABASTEC,                                            ');
      SQL.Add('            0 AS PLACAMOV_ID,                                           ');
      SQL.Add('            T.TAB_CODIGO AS CodProduto,                                 ');
      SQL.Add('            QUANTIDADE = ' + sQtde + ' ,');
      SQL.Add('            EncInicial = CASE T.GRU_CODIGO WHEN ''040'' THEN T.TAB_NUMINIC ELSE 0 END,   ');
      SQL.Add('            EncFinal = CASE T.GRU_CODIGO WHEN ''040'' THEN T.TAB_NUMINIC + ' + sQtde + ' ELSE 0 END,   ');
      SQL.Add('            T.TAB_VRUNITV AS ValorUnitario,                             ');
      SQL.Add('            T.TAB_DESCONTO,                                             ');
      SQL.Add('            T.TAB_VRUNITP AS ValorPrazo,                                ');
      SQL.Add('            VrUntCLASSEPRECO = COALESCE(ROUND((SELECT UNITARIO FROM PROCLASSES');
      SQL.Add('                                              WHERE PRODUTO_COD = T.TAB_CODIGO');
      SQL.Add('                                                AND CLASSES_ID = (SELECT CLASSES_ID');
      SQL.Add('                                                                    FROM CLASSES');
      SQL.Add('                                                                   WHERE CODIGO = (SELECT CLASSES_ID ');
      SQL.Add('                                                                                     FROM CLIENTE');
      SQL.Add('                                                                                    WHERE CLI_CODI = ' + StringToSql(fCampos.CODCliente) + ' ))),3,1),0),');
      SQL.Add('           VrUntPrecoEsp =  COALESCE(ROUND((SELECT PRE_VRUNIT FROM CLIPRECO');
      SQL.Add('                                                           WHERE TAB_CODIGO = T.TAB_CODIGO');
      SQL.Add('                                                             AND CLI_CODI = ' + StringToSql(fCampos.CODCliente) + ' ),3,1),0),');
      SQL.Add('           VALORTOTAL  = ROUND((T.TAB_VRUNITV * ' + sQtde + '),3,1),');
      SQL.Add('           VRTOTALCLASSEPRECO = COALESCE(ROUND((SELECT UNITARIO FROM PROCLASSES');
      SQL.Add('                                                           WHERE PRODUTO_COD = T.TAB_CODIGO');
      SQL.Add('                                                             AND CLASSES_ID = (SELECT CLASSES_ID');
      SQL.Add('                                                                                 FROM CLASSES');
      SQL.Add('                                                                                WHERE CODIGO = (SELECT CLASSES_ID ');
      SQL.Add('                                                                                                  FROM CLIENTE');
      SQL.Add('                                                                                                 WHERE CLI_CODI = ' + StringToSql(fCampos.CODCliente) + ' ))),3,1),0) * ' + sQtde + ',');
      SQL.Add('           VRTOTALPRECOESPEC =  round(COALESCE((SELECT PRE_VRUNIT FROM CLIPRECO');
      SQL.Add('                                                           WHERE TAB_CODIGO = T.TAB_CODIGO');
      SQL.Add('                                                             AND CLI_CODI = ' + StringToSql(fCampos.CODCliente) + ' ),0) * ' + sQtde + ',3,1)');
      SQL.Add('       From TABELA T');
      SQL.Add('        WHERE ');

      if not pVndCombF6comAutomacao then
        if bAtivarPlacaConcentra then
          SQL.Add('  T.GRU_CODIGO <> ''040'' AND ' );

      SQL.Add('              T.INATIVO = 0');
      SQL.Add('          AND T.TAB_CODIGO = ''' + pCodProduto + '''');
      SQL.Add('          AND ((T.PISTA_ID  IS NULL) ');
      SQL.Add('           OR (T.PISTA_ID   = 0) ');
      SQL.Add('           OR (T.PISTA_ID   = ' + IntToStr(iPistaVenda) + '))');
      Open;
    end;

    if not fqryPLM.IsEmpty then begin
      if bUtilizaNFCE then begin
        IniciaVendaNfce;
      end;
      if AtribuiValores(pValorTotal) then begin
        MostraPrecoEspecial(pPrecoEspecial);
        if not bUtilizaNFCE then
          vImprimiuItem := ImprimeItem( fCampos.CodVendedor,
                                        fCODPRODUTO,
                                        fQUANTIDADE,
                                        fValorUnitario,
                                        fValorUnitarioL,
                                        fVALORTOTAL,
                                        pArredonda)
        else
          vImprimiuItem := ImprimeItemNFCe(fCampos.CodVendedor,
                                           fCODPRODUTO,
                                           fQUANTIDADE,
                                           fValorUnitario,
                                           fValorUnitarioL,
                                           fVALORTOTAL,
                                           pArredonda,
                                           fEncInicial,
                                           fEncFinal);

        if vImprimiuItem then begin
          ImprimeImage(pImage);
          pImage.Refresh;
          fDbfTempID := GravaDBFTemp(iiF(pArredonda,0,1));
          if not GravaBasicoArquivo(fCampos.CupomFiscal,
                                    fCampos.IdNotaFiscal,
                                    Strtoint(fCampos.CodVendedor),
                                    fCampos.Identificador,
                                    fCODPRODUTO,
                                    fQUANTIDADE,
                                    fValorUnitario,
                                    fValorUnitarioL,
                                    fVALORTOTAL,
                                    pArredonda) then begin
            DeuErro := true;
          end;

          vArquivoId := GetArquivoID(fCampos.IdNotaFiscal,
                                     Copy(Trim(fCODPRODUTO),1,20),
                                     vItem);

          SetItemLista(vItem,
                       fCODPRODUTO,
                       fDescricaoProduto,
                       TruncaFloat(fQUANTIDADE,3,false),
                       fValorUnitarioL,
                       fVALORTOTAL,
                       vArquivoId);
        end
        else begin
          DeuErro := true;
        end;
      end
      else begin
        DeuErro := true;
      end;

      if DeuErro then begin
        //caso já tenha impresso mais de um item, então terá que
        //apagar o basico e arquivo
        ApagaBasicoArquivo;
        RetornaPlacaMov;
        //esta chamada a função abaixo verifica se a ECF pode cancelar o
        //cupom, caso possa o parametro tem que estar setado para
        //true para cancelar o cupom fiscal aberto ou o ultimo impresso
        if not bUtilizaNFCE then begin
          if ImpressoraFiscal.CupomFiscalAberto(False) then begin
            if ImpressoraFiscal.ImpressoraPoucoPapel = 'FIM DE PAPEL' then begin
              Atencao('Coloque o papel na impressora fiscal!');
            end
            else begin
              ImpressoraFiscal.CupomFiscalAberto(true);
            end;
          end;
        end;
        result := false;
      end
      else begin
        result := true;
      end;
    end;
  except
    on e: exception do begin
      ApagaBasicoArquivo;
      ImpressoraFiscal.CupomFiscalAberto(true);
      result := false;
      Atencao('Erro na venda do item!'+ sLineBreak + e.message);
    end;
  end;
end;

function TImprimeItens.IniciaVenda(pPlacaMovID: string): Boolean;
var
  fItens : TItensImpresso;
begin
  try
    with fQryItens do begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT P.CodProduto,');
      SQL.Add('       T.TAB_DESCR,');
      SQL.Add('       P.Quantidade,');
      SQL.Add('       P.ValorUnitario,');
      SQL.Add('       P.ValorTotal,');
      SQL.Add('       P.PlacaMov_id');
      SQL.Add('  FROM PLACAMOV p, TABELA T');
      SQL.Add(' WHERE T.TAB_CODIGO = P.CodProduto');
      SQL.Add('   AND PLACAMOV_ID IN ('+pPlacaMovID+')');
      SQL.Add('   AND CUPOMFISCAL <> ''''');
      SQL.Add('   AND DATACAIXA   = ' + DateToSql(fCampos.DataCaixa));
      SQL.Add('   AND PLPISTA     = 1');
      SQL.Add('   AND TURNO_ID    = ' + IntToStr(fCampos.Turno));
      sql.Add('  ORDER BY P.CODPRODUTO DESC, P.DATAHORA');
      Open;
    end;

//    if not fQryItens.IsEmpty then begin
//      fQryItens.First;
//      while not fQryItens.Eof do begin
//        fItens := TItensImpresso.Create;
//        fItens.ItemECF    := 0;
//        fItens.Codigo     := StrToInt(Trim(fQryItens.FieldByName('CodProduto').AsString ));
//        fItens.Descricao  := Trim(fQryItens.FieldByName('TAB_DESCR').AsString);
//        fItens.Quantidade := TruncaFloat(fQryItens.FieldByName('Quantidade').AsFloat,3,false);
//        fItens.Unitario   := fQryItens.FieldByName('ValorUnitario').AsFloat;
//        fItens.Valor      := fQryItens.FieldByName('ValorTotal').AsFloat;
//        fItens.PlacaMovId := fQryItens.FieldByName('Placamov_id').AsInteger;
//        fItens.Cancelado  := False;
//        fItens.Combustivel := True;
//        fItens.ArquivoID   := 0;
//        fItens.impresso    := False;
//        fListaItens.Add(fItens);
//        fQryItens.Next;
//      end;
//    end;

    Result := True;
  except
    Result := False;
  end;
end;

procedure TImprimeItens.SetItemLista(pItem       : Integer;
                                     pCodProduto : String;
                                     pDescricao  : string;
                                     pQuantidade : double;
                                     pVrUnit     : double;
                                     pVrTotal    : double;
                                     pArquivoId  : Integer);
var
  fItens : TItensImpresso;
begin
  fItens := TItensImpresso.Create;
  if Assigned(fListaItens) then begin
    fItens.ItemECF     := pItem;
    fItens.Codigo      := Trim(pCodProduto);
    fItens.Descricao   := pDescricao;
    fItens.Quantidade  := TruncaFloat(pQuantidade,3,false);
    fItens.Unitario    := pVrUnit;
    fItens.Valor       := pVrTotal;
    fItens.PlacaMovId  := 0;
    fItens.Cancelado   := False;
    fItens.Combustivel := false;
    fItens.ArquivoID   := pArquivoId;
    fItens.impresso    := true;
    fItens.QtdPto      := fQtdePto;
    fListaItens.Add(fItens);
  end;
end;

procedure TImprimeItens.SetMensagem(pMensagem : string);
begin
  fInterfaceNfCe.SetMensagem(pMensagem);
end;

procedure TImprimeItens.SetDestinatario(pDestinatarioCNPJ        : string;
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
                                          pDestinatarioCodigo      :Integer);
var
  vIndicador : Integer;
  vIndFinal  : integer;
  InscEstDest : String;
begin
  InscEstDest := TrataIE_NFe(vIndicador,fCampos.Cli_InscEst,IntToStr(fCampos.ModeloNFe),vIndFinal);

  fInterfaceNfCe.SetDestinatario(pDestinatarioCNPJ,
                                 InscEstDest,
                                 pDestinatarioNomeRazao,
                                 pDestinatarioFone,
                                 pDestinatarioCEP,
                                 pDestinatarioLogradouro,
                                 pDestinatarioNumero,
                                 pDestinatarioComplemento,
                                 pDestinatarioBairro,
                                 BuscaCodigoCidadeIBGE(StrToInt(pDestinatarioCidadeCod)),
                                 pDestinatarioCidade,
                                 pDestinatarioUF,
                                 IntToStr(vIndicador),
                                 pDestinatarioCodigo);
end;

constructor TImprimeItens.Create(pModeloECF: string);
begin
  fJaIniciouVndNFCe := false;
  fCampos   := TCampos.GetInstance;
  fCampos.ModeloECF := pModeloECF;
  fQtdePto          := 0;
  fQtdePtoTotal     := 0;
  fItem             := 0;
  fVrTotalProdutos  := 0;
  fVrTotalPago      := 0;
  fDescarregaNFCe   := false;
  fStringList       := TStringList.Create;
  fTextoComprovante := TStringList.Create;

  fqryPLM := TADOQuery.Create(nil);
  fqryPLM.Connection := DM.ADOconexao;

  fqryCabecalho := TADOQuery.Create(nil);
  fqryCabecalho.Connection := dm.ADOconexao;

  fQryFpgto := TADOQuery.Create(nil);
  fQryFpgto.Connection := dm.ADOconexao;

  fQryItens := TADOQuery.Create(nil);
  fQryItens.Connection := dm.ADOconexao;

  fQryExiste := TADOQuery.Create(nil);
  fQryExiste.Connection := dm.ADOconexao;

  fQryDBFTemp := TADOQuery.Create(nil);
  fQryDBFTemp.Connection := dm.ADOconexao;

  fQryStatusCupom := TADOQuery.Create(nil);
  fQryStatusCupom.Connection := dm.ADOconexao;

  fQryPontos_Cli := TADOQuery.Create(nil);
  fQryPontos_Cli.Connection := dm.ADOconexao;

  fListaItens := TList.Create;

  fGravouCabecalho := false;
  fPrecoEspecial := False;
  fCupomAberto   := False;
  fFechamentoAbortado := false;
  fUtilizandoTef      := false;
  fCancelando         := false;
  fVndUsouTEF         := false;

  if bUtilizaNFCE then begin
    fInterfaceNfCe  := TNFEInterfaceV3.create('','',True);
    fAcbrAtivo      := true;// fInterfaceNfCe.Ativo;
    frmComprovanteNFCe := TfrmComprovanteNFCe.Create(nil);
    if not fAcbrAtivo then begin
      Atencao('O AcbrNfeMonitor não está ativo!' + sLineBreak +
              'Não será possível emitir a NFc-e!');
    end;
  end;
end;

function TImprimeItens.MonitorAtivo: boolean;
begin
  result := true;
  exit;

//  if bUtilizaNFCE then begin
//    fAcbrAtivo := fInterfaceNfCe.Ativo;
//    result     := fAcbrAtivo;
//    if not fAcbrAtivo then begin
//      Atencao('O AcbrNfeMonitor não está ativo!' + sLineBreak +
//              'Não será possível emitir a NFc-e!');
//    end;
//  end;
end;

destructor TImprimeItens.Destroy;
begin
  if fStringList.Count > 0 then
    fStringList.Clear;

  if fTextoComprovante.count > 0 then
    fTextoComprovante.clear;

  if Assigned(fStringList)then
    FreeAndNil(fStringList);

  if Assigned(fqryPLM)then
    FreeAndNil(fqryPLM);

  if Assigned(fqryCabecalho)then
    FreeAndNil(fqryCabecalho);

  if Assigned(fQryFpgto)then
    FreeAndNil(fQryFpgto);

  if Assigned(fQryItens)then
    FreeAndNil(fQryItens);

  if Assigned(fQryExiste)then
    FreeAndNil(fQryExiste);

  if Assigned(fQryDBFTemp) then
    FreeAndNil(fqryDBFTemp);

   fVrTotalPago := 0;
   fJaIniciouVndNFCe := false;
   fDescarregaNFCe   := false;
   fVndUsouTEF       := false;

end;

function TImprimeItens.ImprimeItemNFCe(pCodVendedor,
                                       pCodProduto    : string;
                                       pQuantidade,
                                       pUnitario,
                                       pUnitarioL,
                                       pValorTotal    : Double;
                                       var pArredonda : Boolean;
                                       pEncInicial    : double = 0;
                                       pEncFinal      : double = 0): Boolean;
var
//  vDif         : Double;
//  vMR_OK       : Integer;
  vVrUnit      : Double;
  vDecimais    : Integer;
  vVrTotal     : Double;
  vAliq        : Double;
  vrAliq       : Double;
  qryTabela    : TADOQuery;
  vCFOP        : String;
  vCombustivel : Boolean;
  vcProdANP     : Integer;
  vUFCons       : string;
  vCST          : String;
begin
  try
    Result  := False;

    qryTabela := TADOQuery.Create(nil);
    qryTabela.Connection := DM.ADOconexao;

    with qryTabela do begin
      Close;
      SQL.Clear;
      SQL.Add(' SELECT T.CODIGO_NCM, ');
      SQL.Add('        T.TAB_CODRED, ');
      SQL.Add('        T.ICM_PERC,   ');
      SQL.Add('        T.TAB_VRUNITV AS VALORUNITARIO,');
      SQL.Add('        T.TAB_DESCR   AS DESCRICAO,');
      SQL.Add('        T.GRU_CODIGO, ');
      SQL.Add('        T.UNI_SIGLA,  ');
      Sql.Add('        Coalesce(T.CODIGO_ANP,0) as CODIGO_ANP, ');
      SQL.Add('        T.GRU_CODIGO,  ');
      SQL.Add('        T.CodSituacaoTributaria, ');
      sql.add('        COALESCE(T.PERC_RED_ICMS,0) AS PERC_RED_ICMS, ');
      sql.add('        T.CST_PIS_SAIDA as CST_PIS,');
      sql.add('        T.CST_COFINS_SAIDA AS CST_COFINS,');
      sql.add('        E.PERCENT_PIS AS ALIQ_PIS,');
      sql.add('        E.PERCENT_COFINS AS ALIQUOTA_COFINS, ');
      sql.add('        CASE WHEN T.PIS_COFINS_CUMULATIVO = 1 THEN ''S'' ELSE ''N'' END AS PIS_COFINS_CUMULATIVO, ');
      sql.add('        T.CSOSN, ');
      sql.add('        T.CSOSN_ORIGEM, ');
      sql.add('        T.CSOSN_BCICMS, ');
      sql.add('        T.CSOSN_BCICMSST ');
      SQL.Add('   FROM TABELA T, EMPRESA E');
      SQL.Add('  WHERE T.TAB_CODIGO = ''' + Trim(pCodProduto) + '''' );
      Open;
    end;

    fDescricaoProduto  := Trim(qryTabela.FieldByName('DESCRICAO').AsString);
    vCombustivel       := IIf(qrytabela.fieldbyname('GRU_CODIGO').AsString = '040',true,false);
    vCST               := qryTabela.FieldByName('CodSituacaoTributaria').AsString;
    fCSOSN             := ConverteCSOSN(qryTabela.FieldByName('CSOSN').AsInteger);
    fCSOSN_ORIGEM      := ConverteCSOSN_ORIGEM(qryTabela.FieldByName('CSOSN_ORIGEM').AsInteger);
    fCSOSN_MODBCICMS   := ConverteCSOSN_MODBCICMS(qryTabela.FieldByName('CSOSN_BCICMS').AsInteger);
    fCSOSN_MODBCICMSST := ConverteCSOSN_MODBCICMSST(qryTabela.FieldByName('CSOSN_BCICMSST').AsInteger);

    if (sEmpresaUf = fCampos.Cli_UF) then begin
      if vCombustivel  then begin
        vCFOP := '5656';
        vCST  := '060';
      end
      else
        vCFOP := RetornaCFOP_GRUPO(Trim(qryTabela.FieldByName('GRU_CODIGO').AsString));
    end
    else begin
      if vCombustivel then
        if fCampos.Cli_UF = '' then //dentro do estado
            vCFOP := '5656'
        else
          vCFOP := '5667'// '6656'
      else
        vCFOP := RetornaCFOP_GRUPO(Trim(qryTabela.FieldByName('GRU_CODIGO').AsString));

      //nesta condição é uma nfce o cliente é fora do estado e o produto está configurado
      //com 5656 que é para combustivel ou lubrificante então o cfop será 5667
      if vCFOP = '5656' then
        vCFOP := '5667';
    end;

    if not ValidaCST(vCST, vCFOP) then begin
      Atencao('CFOP: ' + vCFOP + ' não é compatível com CST: ' + vCST + #13#10 +
              'Configure o produto: ' + pCodProduto + ' ' + fDescricaoProduto + #13#10 +
              'No cadastro do produto verifique o CST!' + #13#10 +
              'No cadastro do grupo de produto verifique o CFOP!');
      result := false;
      exit;
    end;

    if not Em(TRIM(qryTabela.FieldByName('ICM_PERC').AsString),['FF', 'II', 'NN', '0000']) and
          (TRIM(qryTabela.FieldByName('ICM_PERC').AsString) <> '') and
          (not (StrToIntDef(vCST,40) in[40,41,60])) then begin
      vAliq  := StrToFloatDef(qryTabela.FieldByName('ICM_PERC').AsString,0) / 10000;
      vrAliq := fVALORTOTAL * vAliq;
      vAliq := vAliq * 100;
    end
    else begin
      vrAliq := 0;
      vAliq  := 0;
    end;

    if ((pUnitario - pUnitarioL) <> 0) then
      vVrUnit := pUnitarioL
    else
      vVrUnit := pUnitario;

    vDecimais  := 3;

    //ESTAVA OCORRENDO MUITOS ERROS NA VISUALIZAÇÃO DE RELATORIOS E CAIXA
    //MOSTRANDO MUITOS ITENS DO MESMO PRODUTO, PARA FAZER O CALCULO NÃO SE
    //DEVE ALTERAR O VALOR UNITARIO E NEM O VALOR TOTAL.


    IF EM(vCST,['020']) THEN BEGIN
      vVrTotal := pValorTotal * ((100-qryTabela.FieldByName('PERC_RED_ICMS').AsFloat)/100);
      vrAliq   := vVrTotal * (vAliq / 100);
      fQUANTIDADE := TruncaFloat((pValorTotal/fValorUnitarioL),3,False);
    END
    eLSE  begin
      vVrTotal := pValorTotal;
      fQUANTIDADE := TruncaFloat((vVrTotal/fValorUnitarioL),3,False);
    end;

    vVrTotal   := TruncaFloat(vVrTotal,2,false);

    //AQUI CALCULA O VALOR UNITÁRIO
    //fValorUnitarioL := TruncaFloat((vVrTotal / pQuantidade),10,False);

    pArredonda := TruncaArredonda(vVrTotal,fQUANTIDADE,fValorUnitarioL);

    if qryTabela.FieldByName('PERC_RED_ICMS').AsFloat <= 0 then
      fVALORTOTAL := vVrTotal;

    ImpDocFiscal.BuscaAliquota(Trim(pCodProduto),fQUANTIDADE * vVrUnit);

    vCombustivel  := IIf(Trim(qryTabela.FieldByName('GRU_CODIGO').AsString) = '040', True, False);
    vcProdANP     := qryTabela.FieldByName('CODIGO_ANP').AsInteger;
    vUFCons       := IIf(fCampos.Cli_UF = '','GO',fCampos.Cli_UF);


    Result := fInterfaceNfCe.AddItensNfe([vCFOP,                                             //CFOP
                                          Trim(fCODPRODUTO),                                 //Codigo
                                          Trim(qryTabela.FieldByName('DESCRICAO').AsString), //Descricao
                                          qryTabela.FieldByName('TAB_CODRED').AsString,      //EAN
                                          qryTabela.FieldByName('CODIGO_NCM').AsString,      //NCM
                                          qryTabela.FieldByName('UNI_SIGLA').AsString,       //Unidade
                                          TruncaFloat(fQUANTIDADE,3,false),                  //Quantidade
                                          fValorUnitarioL, //vVrUnit,                        //ValorUnitario
                                          fVALORTOTAL,                                       //ValorTotal
                                          0,                                                 //ValorDesconto
                                          '',                                                //NVE
                                          '',                                                //nFCI
                                          '',                                                //fnRECOPI
                                          '',                                                //fpDevol
                                          '',                                                //fvIPIDevol
                                          vCST,                                              //CST              15
                                          vVrTotal,                                          //ValorBase        16
                                          vAliq,                                             //Aliquota         17
                                          vrAliq,                                            //ValorAliq        18
                                          qryTabela.FieldByName('PERC_RED_ICMS').AsFloat,    //Percentual da base de calculo 19
                                          vCombustivel,                                      //se é combustivel
                                          vcProdANP,                                         //Codigo ANP
                                          vUFCons,                                           //Sigla da UF de consumo
                                          pEncInicial,                                       //Encerrante Inicial
                                          pEncFinal,                                         //Encerrante Final
                                          0,                                                 //ValorBaseST
                                          0,                                                 //AliquotaST
                                          0,                                                 //ValorST
                                          qryTabela.FieldByName('CST_PIS').AsString,                            //PISCST
                                          qryTabela.FieldByName('ALIQ_PIS').AsFloat,                            //PISAliquota
                                          (fVALORTOTAL * qryTabela.FieldByName('ALIQ_PIS').AsFloat)/100,        //PISValor
                                          qryTabela.FieldByName('CST_COFINS').AsString,                         //COFINSCST
                                          qryTabela.FieldByName('ALIQUOTA_COFINS').AsFloat,                     //COFINSAliquota
                                          (fVALORTOTAL * qryTabela.FieldByName('ALIQUOTA_COFINS').AsFloat)/100, //COFINSValor
                                          qryTabela.FieldByName('PIS_COFINS_CUMULATIVO').AsString,              //PIS_COFINS_CUMULATIVO
                                          'B',                                                                  //TIPO_ALIQUOTA_PIS_COFINS
                                          fCSOSN,
                                          fCSOSN_ORIGEM,
                                          fCSOSN_MODBCICMS,
                                          fCSOSN_MODBCICMSST]);
    Inc(fItem);
  except
    on e: exception do begin
      result := false;
      MSGAtencao(e.Message);
    end;
  end;
end;


//Este Método é responsável por imprimir o item do cupom
//logo após de confirmado a impressao deverá gravar o placamov
function TImprimeItens.ImprimeItem(pCodVendedor,
                                   pCodProduto: string;
                                   pQuantidade,
                                   pUnitario,
                                   pUnitarioL,
                                   pValorTotal: Double;
                                   var pArredonda: Boolean): Boolean;
var
  sAliq        : AnsiString;
  vDif         : Double;
//  fltVlrItem,
//  fltTotal,
//  fltVlrDesc   : Real;
//  intI         : Integer;
  vMR_OK       : Integer;
  vContinuar   : Boolean;
  vVrUnit      : Double;
  vVrUnitSTR   : String;
  vDecimais    : Integer;
  vVrTotal     : Double;

  qryTabela    : TADOQuery;
  bAliISS      : Boolean;
  strModeloECF : STRING;
  pPoucoPapel  : String;
begin
  try
    Result  := False;
    bAliISS := false;
    vContinuar := True;

    qryTabela := TADOQuery.Create(nil);
    qryTabela.Connection := DM.ADOconexao;

    with qryTabela do begin
      Close;
      SQL.Clear;
      SQL.Add(' SELECT ');
      SQL.Add('        T.ICM_PERC    AS ALIQUOTA,');
      SQL.Add('        T.TAB_VRUNITV AS VALORUNITARIO,');
      SQL.Add('        T.TAB_DESCR   AS DESCRICAO,');
      SQL.Add('        T.GRU_CODIGO ');
      SQL.Add('   FROM TABELA T ');
      SQL.Add('  WHERE TAB_CODIGO = ''' + Trim(pCodProduto) + '''' );
      Open;
    end;

    fDescricaoProduto := Trim(qryTabela.FieldByName('DESCRICAO').AsString);

    sAliq := trim(qryTabela.FieldByName('ALIQUOTA').AsString);

    with DM.qryGeral do begin
      if (Length(sAliq) = 2) and (sAliq <> 'FF') and (sAliq <> 'II') and (sAliq <> 'NN') then begin
        if (iEcf <> 8) and (iEcf <> 9) then begin
          sAliq := sAliq + '00';
        end;
      end;
      Close;
      SQL.Clear;
      SQL.Add('SELECT');
      SQL.Add('       ICM_ISS');
      SQL.Add('  FROM TABICMS');
      SQL.Add(' WHERE ICM_ALIQUOTA =  ' + StringToSql(sAliq));
      Open;
      if not IsEmpty then begin
        if FieldByName('ICM_ISS').Value = 'S' then begin
          ImpressoraFiscal.RetornaIndiceAliquotaISS(sAliq);
          bAliISS := True;
        end;
      end;
    end;

    if qryTabela.RecordCount = 1 then begin
      if not bAliISS then begin
        SetCupomAberto;
        if not fCupomAberto then begin
          ImpressoraFiscal.AbreCupom(fCampos.CNPJCpf, fcampos.CodCliente + ' - ' + fCampos.NomeCliente, fCampos.CliEnderecoComercial, '0000');
          fCampos.NumSerieECF := ImpressoraFiscal.NumeroSerie;
          fCampos.NumECF      := ImpressoraFiscal.NumeroCaixa;
          SetCupomAberto;
          InsereStatusCupom;
          if FindWindow('TAppBuilder', nil) = 0 then begin
            //desabilita ctrl+alt+del
            CTRLALTDEL(False);
          end;
        end;
      end;
    end
    else begin
      Result := False;
      raise exception.Create('Ocorreu um erro na impressão do item!');
      exit;
    end;

    vDif := 0;
    if ((not bAliISS)) then begin

      if ((pUnitario - pUnitarioL) <> 0) then
        vVrUnit := pUnitarioL
      else
        vVrUnit := pUnitario;

      vDecimais  := 3;

      vVrUnitSTR := Trim(FloatToStr(TruncaFloat(vVrUnit,vDecimais)));

      vVrTotal   := TruncaFloat(pValorTotal,2,false);

      pArredonda := TruncaArredonda(vVrTotal,pQuantidade,vVrUnit);

      fVALORTOTAL := vVrTotal;

      strModeloECF := ImpressoraFiscal.ModeloECF;

      ImpDocFiscal.BuscaAliquota(Trim(pCodProduto),pquantidade * vVrUnit);

      repeat
        Result := ImpressoraFiscal.VendeItem(Trim(pCodProduto),
                                             Trim(qryTabela.FieldByName('DESCRICAO').AsString),
                                             sAliq,
                                             'F', {(F)FRACIONADO (I) INTEIRO}
                                             Trim(FloatToStr(TruncaFloat(pquantidade,3,false))),
                                             vVrUnitSTR,
                                             sTipoDesconto,
                                             SomenteNumeros(FormatFloat('00.00', TruncaFloat((vDif * -1),3))),
                                             Trim(qryTabela.FieldByName('GRU_CODIGO').AsString),
                                             bAliISS,
                                             '0000',
                                             strModeloECF,
                                             pArredonda);


        pPoucoPapel := ImpressoraFiscal.ImpressoraPoucoPapel;

        if (pPoucoPapel = 'FIM DE PAPEL') then begin
          vMR_OK := MessageBox(0, 'Para continuar troque o papel da impressora!',
                           '    I M P R E S S O R A   S E M   P A P E L ! ! !',
                           MB_ICONWARNING or MB_OKCANCEL);

          IF vMR_OK = 2 then begin
            vContinuar := False;
            Break;
          end;
        end;
      until (pPoucoPapel <> 'FIM DE PAPEL');

      if not vContinuar then begin
        Result := False;
        Exit;
      end;

      Inc(fItem);
    end;
  except
    on e: exception do begin
      result := false;
      MSGAtencao(e.Message);
    end;
  end;
end;

function TImprimeItens.AlteraItemLista(pPlacaMovID: Integer; pUnit, pVrtotal: Double): Boolean;
var
  i : Integer;
begin
  try
    for I := 0 to fListaItens.Count - 1 do begin
      if TItensImpresso(flistaItens[i]).PlacaMovId = pPlacaMovID then begin
        TItensImpresso(flistaItens[i]).Unitario := pUnit;
        TItensImpresso(flistaItens[i]).Valor    := pVrtotal;
      end;
    end;
  except
    result := false;
  end;
end;

function TImprimeItens.GravaBasicoArquivo(pNumeroCupom   : AnsiString;
                                          pIdNotaFiscal  : Integer;
                                          pCodVendedor   : Integer;
                                          pIdentificador : String;
                                          pCODPRODUTO    : String;
                                          pQUANTIDADE    : Double;
                                          pValorUnitario : Double;
                                          pValorUnitarioL: Double;
                                          pVALORTOTAL    : Double;
                                          pArredonda     : Boolean;
                                          pPlacaMov      : Integer = 0;
                                          pDbfTempID     : integer = 0;
                                          intNumAbastec  : integer = 0): Boolean;
var
  sDataCupom : AnsiString;
  vDataCupom : TDateTime;
begin
  try
    result := true;

    if not fGravouCabecalho then begin

      if not bUtilizaNFCE then
        sDataCupom := ImpressoraFiscal.ObtemDataUltimoCupomFiscal
      else
        sDataCupom := FormatDateTime('DDMMYY',Now);

      vDataCupom := StrToDate(Copy(sDataCupom,1,2) + '/' +
                              Copy(sDataCupom,3,2) + '/' +
                              IntToStr((StrToInt(Copy(sDataCupom,5,2)) + 2000)));

      result := GravaCabecalhoPedido(trim(pNumeroCupom),
                                     pIdNotaFiscal,
                                     'D',
                                     fCampos.DataCaixa,
                                     pCodVendedor,
                                     0,
                                     fCampos.CODCliente,
                                     1,
                                     fCampos.Pista,
                                     fCampos.Turno,
                                     '2', //codigousuario,
                                     0,   //iIdMovimentoFecha,
                                     TRIM(fCampos.NumSerieECF), //NumserialECF,
                                     '5102',
                                     fCampos.NumSerieNFCe,
                                     0,
                                     0,
                                     0,
                                     0,
                                     TRIM(fCampos.NumECF),
                                     NOW,
                                     vDataCupom);
      fGravouCabecalho := result;
    end;

    if not Result then
      Exit;

    result := GravaItensPedidos(pIdNotaFiscal,
                                pCODPRODUTO,
                                0,
                                pQUANTIDADE,
                                pPlacaMov,
                                pDbfTempID,
                                pIdentificador,
                                pValorUnitario,
                                pValorUnitarioL,
                                intNumAbastec,
                                0,
                                0,
                                '',
                                IntToStr(fCampos.pista),
                                '',
                                pArredonda);
  except
    Result := false;
  end;
end;

function TImprimeItens.CancelaBasico: Boolean;
var
  AdoQryBasico  : TADOQuery;
begin
  try
    AdoQryBasico := TADOQuery.Create(nil);
    AdoQryBasico.Connection := dm.ADOconexao;

    with AdoQryBasico do begin
      Close;
      SQL.Clear;
      SQL.Add(' select COUNT(IDNOTAFISCAL) as QTDE ');
      SQL.Add('  from BASICO     ');
      SQL.Add(' where IDNOTAFISCAL = ' + IntToStr(fCampos.IdNotaFiscal));
      Open;

      if FieldByName('QTDE').AsInteger > 0 then begin
        Close;
        SQL.Clear;
        SQL.Add(' UPDATE BASICO     ');
        SQL.Add(' SET BAS_CANCELADO = 1 ');
        SQL.Add(' where IDNOTAFISCAL = ' + IntToStr(fCampos.IdNotaFiscal));
        ExecSQL;
      end;
    end;
    Result := True;
  except
    result := False;
  end;
end;

function TImprimeItens.ApagaBasicoArquivo: boolean;
var
  AdoQryGeral    : TADOQuery;
begin
  try
    AdoQryGeral := TADOQuery.Create(nil);
    AdoQryGeral.Connection := dm.ADOconexao;

    with AdoQryGeral do begin
      Close;
      SQL.Clear;
      SQL.Add(' select COUNT(IDNOTAFISCAL) as QTDE ');
      SQL.Add('  from ARQUIVO     ');
      SQL.Add(' where IDNOTAFISCAL = ' + IntToStr(fCampos.IdNotaFiscal));
      Open;

      if FieldByName('QTDE').AsInteger > 0 then begin
        Close;
        SQL.Clear;
        SQL.Add(' DELETE ARQUIVO     ');
        SQL.Add(' where IDNOTAFISCAL = ' + IntToStr(fCampos.IdNotaFiscal));
        ExecSQL;
      end;

      Close;
      SQL.Clear;
      SQL.Add(' select COUNT(IDNOTAFISCAL) as QTDE ');
      SQL.Add('  from BASICO     ');
      SQL.Add(' where IDNOTAFISCAL = ' + IntToStr(fCampos.IdNotaFiscal));
      Open;

      if FieldByName('QTDE').AsInteger > 0 then begin
        Close;
        SQL.Clear;
        SQL.Add(' DELETE BASICO     ');
        SQL.Add(' where IDNOTAFISCAL = ' + IntToStr(fCampos.IdNotaFiscal));
        ExecSQL;
      end;
    end;
    Result := True;
  except
    Result := False;
  end;
end;

function TImprimeItens.RetornaVenda: Boolean;
begin
  try
    Result := False;
    RetornaMoviment;
    RetornaCheque;
    ApagaBasicoArquivo;
    RetornaPlacaMov;
    Result := true;
  except
    Result := False;
  end;
end;


function TImprimeItens.RetornaPlacaMov: Boolean;
var
  vqryAlteraPLM : TADOQuery;
  vPlacaMovId   : String;
  index         : integer;
begin
  try
    vqryAlteraPLM := TADOQuery.Create(nil);
    vqryAlteraPLM.Connection := dm.ADOconexao;

    with vqryAlteraPLM do begin

      Close;
      SQL.Clear;
      SQL.Add(' select COUNT(IDNOTAFISCAL) as QTDE ');
      SQL.Add('  from PLACAMOV     ');
      SQL.Add(' where IDNOTAFISCAL = ' + IntToStr(fCampos.IdNotaFiscal));
      Open;

      if FieldByName('QTDE').AsInteger > 0 then begin
        Close;
        sql.Clear;
        sql.Add(' UPDATE PLACAMOV ');
        SQL.Add('    SET IDNOTAFISCAL = 0, ');
        SQL.Add('        CUPOMFISCAL  = 0, ');
        sql.Add('        CUPOMFECHA   = 0 ');
        SQL.Add('  WHERE IDNOTAFISCAL = ' + IntToStr(fCampos.IdNotaFiscal));
        ExecSQL;
      end;

      Result := true;
    end;
  except
    Result := False;
  end;
end;

function TImprimeItens.RetornaMoviment: Boolean;
var
  vqryMovimet   : TADOQuery;
  vIdMoviment   : Integer;

begin
  try
    vIdMoviment := 0;
    vqryMovimet := TADOQuery.Create(nil);
    vqryMovimet.Connection := dm.ADOconexao;

    with vqryMovimet do begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT IdMovimento ');
      SQL.Add(' FROM BASICO ');
      SQL.Add(' WHERE IdNotaFiscal = '  + IntToStr(fCampos.IDNotaFiscal));
      Open;

      if not IsEmpty then
        vIdMoviment := FieldByName('IdMovimento').AsInteger;

      if vIdMoviment > 0 then begin
        Close;
        SQL.Clear;
        SQL.Add('DELETE MOVIMENT');
        SQL.Add(' WHERE IDMOVIMENTO = ' + IntToStr(vIdMoviment));
        ExecSQL;
      end;
    end;
    Result := true;
  except
    Result := False;
  end;
end;

function TImprimeItens.RetornaCheque: Boolean;
var
  vqryCheque   : TADOQuery;
begin
  try
    vqryCheque := TADOQuery.Create(nil);
    vqryCheque.Connection := dm.ADOconexao;

    with vqryCheque do begin
      Close;
      SQL.Clear;
      SQL.Add('DELETE CHEQUE');
      SQL.Add(' WHERE IDDOCFISCAL = ' + IntToStr(fCampos.IDNotaFiscal));
      ExecSQL;
    end;
    Result := true;
  except
   result := false;
  end;
end;

function TImprimeItens.AlteraBasTipo(pTipo: string): Boolean;
begin
  try
    with fqryCabecalho do begin
      Close;
      SQL.Clear;
      SQL.Add('UPDATE BASICO');
      SQL.Add('   SET BAS_TIPO = ''' + pTipo + '''');
      SQL.Add('WHERE IDNOTAFISCAL = ' + IntToStr(fCampos.IDNotaFiscal));
      ExecSQL;

      result := true;
    end;
  except
    Result := False;
  end;
end;

function TImprimeItens.AlteraIdMovimento(pIdMovimento: string): Boolean;
begin
  try
    with fqryCabecalho do begin
      Close;
      SQL.Clear;
      SQL.Add('UPDATE BASICO');
      SQL.Add('   SET IdMovimento = ' + pIdMovimento);
      SQL.Add('WHERE IDNOTAFISCAL = ' + IntToStr(fCampos.IDNotaFiscal));
      ExecSQL;

      result := true;
    end;
  except
    Result := False;
  end;
end;

function TImprimeItens.AlteraFpgto(pFpgto: string): Boolean;
begin
  try
    with fqryCabecalho do begin
      Close;
      SQL.Clear;
      SQL.Add('UPDATE BASICO');
      SQL.Add('   SET TAF_FORMAPGTO = ' + pFpgto);
      SQL.Add('WHERE IDNOTAFISCAL = ' + IntToStr(fCampos.IDNotaFiscal));
      ExecSQL;

      result := true;
    end;
  except
    Result := False;
  end;
end;

function TImprimeItens.GravaCabecalhoPedido(strBas_NumeroNF  : String;
                                            intIDNotaFiscal  : Integer;
                                            strBas_Tipo      : String;
                                            dtBas_Data       : TDate;
                                            intCodVend,
                                            intCodFornec     : Integer;
                                            strCodClie       : String;
                                            intFormaPgto,
                                            intPista         : Variant;
                                            intTurno         : Variant;
                                            strCodUser       : String;
                                            intIDMovimento   : Integer;
                                            strSerieECF,
                                            strCFOP,
                                            strSerie         : String;
                                            fltBaseICMSSub,
                                            fltICMSSUB,
                                            fltBaseCalcICMS,
                                            fltBaseICMS      : Real;
                                            strNumeroECF     : String;
                                            HoraEmissao      : TDatetime;
                                            dtCupom          : TDate) : Boolean;
VAR
  vAltPlacaMovId : boolean;
  i              : integer;
  vTexto         : string;
begin
  try
    vAltPlacaMovId := false;
    vTexto := '';

    if intIDNotaFiscal = 999999 then
      intIDNotaFiscal := 0;

    Result := False;

    with fqryCabecalho do begin
      Close;
      if Length(fCampos.CNPJCpf) > 0 then begin
        for I := 0 to length(fCampos.CNPJCpf) do begin
          if fCampos.CNPJCpf[i] in['0'..'9'] then
            vTexto := vTexto + fCampos.CNPJCpf[i];
        end;
      end;

      if vTexto <> '' then
        fCampos.CNPJCpf := vTexto;

      //Caso não tenha sido informado um vendedor busca o vendedor padrão
      if intCodVend <= 0 then begin
        SQL.Clear;
        SQL.Add('SELECT COALESCE(VEN_CODVEND_PADRAO,0) AS VEN_CODVEND_PADRAO FROM CONFIG');
        Open;
        intCodVend := FieldByName('VEN_CODVEND_PADRAO').AsInteger;
      end;

      //Buscando forma de pagamento igual a dinheiro
      if intFormaPgto = -1 then begin
        SQL.Clear;
        SQL.Add('SELECT DOC_CODIGO');
        SQL.Add('  FROM DOCFISCA');
        SQL.Add(' WHERE DOC_DESC LIKE ''%DINHEIRO%''');
        Open;
        intFormaPgto := FieldByName('DOC_CODIGO').AsInteger;
      end;

      SQL.Clear;
      SQL.Add('INSERT INTO BASICO(IDNOTAFISCAL,');
      SQL.Add('                   BAS_NUMERONF,');
      SQL.Add('                   BAS_TIPO,');
      SQL.Add('                   VEN_CODVEND,');
      if intCodFornec > 0 then
        SQL.Add('                 FOR_CODFORNEC,');
      SQL.Add('                   CLI_CODI,');
      SQL.Add('                   TAF_FORMAPGTO,');
      SQL.Add('                   PIS_CODIGO,');
      SQL.Add('                   TUR_TURNO,');
      SQL.Add('                   BAS_DATA,');
      SQL.Add('                   CODUSUARIO,');
      SQL.Add('                   IDMOVIMENTO,');
      SQL.Add('                   SERIEECF,');
      SQL.Add('                   CFOP,');
      SQL.Add('                   SERIE,');
      SQL.Add('                   BASEICMSSUBSTITUICAO,');
      SQL.Add('                   ICMSSUBSTITUICAO,');
      SQL.Add('                   BASECALCULOICMS,');
      SQL.Add('                   BAS_CANCELADO,');
      SQL.Add('                   BAS_ICM,');
      SQL.Add('                   NUMERO_ECF, ');
      if bUtilizaNFCE then begin
        if fCampos.ModeloNFe = 65 then
          sql.Add(' NUMERO_NFCE,')
        ELSE
          SQL.Add(' NUMERONOTA,');

        SQL.Add(' MODELONF, ');
      end;
      SQL.Add('                   HORAEMISSAO) ');
      SQL.Add('VALUES(:IDNOTAFISCAL,');
      SQL.Add('       :BAS_NUMERONF,');
      SQL.Add('       :BAS_TIPO,');
      SQL.Add('       :VEN_CODVEND,');
      if intCodFornec > 0 then
        SQL.Add('       :FOR_CODFORNEC,');
      SQL.Add('       :CLI_CODI,');
      SQL.Add('       :TAF_FORMAPGTO,');
      SQL.Add('       :PIS_CODIGO,');
      SQL.Add('       :TUR_TURNO,');
      SQL.Add('       :BAS_DATA,');
      SQL.Add('       :CODUSUARIO,');
      SQL.Add('       :IDMOVIMENTO,');
      SQL.Add('       :SERIEECF,');
      SQL.Add('       :CFOP,');
      SQL.Add('       :SERIE,');
      SQL.Add('       :BASEICMSSUBSTITUICAO,');
      SQL.Add('       :ICMSSUBSTITUICAO,');
      SQL.Add('       :BASECALCULOICMS,');
      SQL.Add('       0,');
      SQL.Add('       :BAS_ICM,');
      SQL.Add('       :NUMERO_ECF,');

      if bUtilizaNFCE then begin
        if fCampos.ModeloNFe = 65 then
          sql.Add(' :NUMERO_NFCE,')
        ELSE
          SQL.Add(' :NUMERONOTA,');

        SQL.Add(' :MODELONF, ');
      end;

      SQL.Add('       :HORAEMISSAO) ');
      Parameters.ParamByName('IDNOTAFISCAL').Value          := intIDNotaFiscal;                     //int
      Parameters.ParamByName('BAS_NUMERONF').Value          := Copy(Trim(strBas_NumeroNF),1,12);    //string
      Parameters.ParamByName('BAS_TIPO').Value              := Copy(Trim(strBas_Tipo),1,1);         //string
      Parameters.ParamByName('VEN_CODVEND').Value           := intCodVend;                          //int
      if intCodFornec > 0 then
        Parameters.ParamByName('FOR_CODFORNEC').Value       := intCodFornec;                        //int
      Parameters.ParamByName('CLI_CODI').Value              := Copy(Trim(strCodClie),1,7);          //string
      Parameters.ParamByName('TAF_FORMAPGTO').Value         := intFormaPgto;                        //int
      Parameters.ParamByName('PIS_CODIGO').Value            := intPista;                            //Variant
      Parameters.ParamByName('TUR_TURNO').Value             := intTurno;                            //Variant
      Parameters.ParamByName('BAS_DATA').Value              := dtBas_Data;                          //date
      Parameters.ParamByName('CODUSUARIO').Value            := Copy(Trim(strCodUser),1,2);          //string
      Parameters.ParamByName('IDMOVIMENTO').Value           := intIDMovimento;                      //int
      Parameters.ParamByName('SERIEECF').Value              := Copy(Trim(strSerieECF),1,20);        //string
      Parameters.ParamByName('CFOP').Value                  := Copy(Trim(strCFOP),1,4);             //string
      Parameters.ParamByName('SERIE').Value                 := Copy(Trim(strSerie),1,2);            //string
      Parameters.ParamByName('BASEICMSSUBSTITUICAO').Value  := fltBaseICMSSub;                      //float
      Parameters.ParamByName('ICMSSUBSTITUICAO').Value      := fltICMSSUB;                          //float
      Parameters.ParamByName('BASECALCULOICMS').Value       := fltBaseCalcICMS;                     //float
      Parameters.ParamByName('BAS_ICM').Value               := fltBaseICMSSub;                      //float
      Parameters.ParamByName('NUMERO_ECF').Value            := Copy(Trim(strNumeroECF),1,3);        //string

      if bUtilizaNFCE then begin
        if fCampos.ModeloNFe = 65 then
          Parameters.ParamByName('NUMERO_NFCE').Value         := Copy(Trim(strBas_NumeroNF),1,12)
        ELSE
          Parameters.ParamByName('NUMERONOTA').Value         := Copy(Trim(strBas_NumeroNF),1,12);

        Parameters.ParamByName('MODELONF').Value            := IntToStr(fCampos.ModeloNFe);
      end;

      Parameters.ParamByName('HORAEMISSAO').Value           := FormatDateTime('hhmm', HoraEmissao); //date
      ExecSQL();

      Close;
      SQL.Clear;
      SQL.Add('UPDATE BASICO');
      SQL.Add('   SET DATACUPOM    = ' + DateTimeToSql(dtCupom));
      SQL.Add(' WHERE IDNOTAFISCAL = ' + IntegerSQL(intIDNotaFiscal));
      SQL.Add('   AND BAS_DATA     = ' + DateToSql(dtBas_Data));
      SQL.Add('   AND PIS_CODIGO   = 1');
      SQL.Add('   AND TUR_TURNO    = ' + IntToStr(iTurno));
      SQL.Add('   AND BAS_TIPO     = ''D''');
      ExecSQL();


      Result := true;

    end;
  except
    on e: Exception do begin
      result := False;
      MSGAtencao(e.Message);
    end;
  end;
end;

function TImprimeItens.GravaItensPedidos(intIDNotaFiscal  : Integer;
                                         strTabCodigo     : String;
                                         intArqItem       : Integer;
                                         fltQtde          : Real;
                                         intPlacaMov_ID   : Integer;
                                         intDBFTEMP_ID    : Integer;
                                         strIdentificador : String;
                                         fltArqVrUnit,
                                         fltArqVrUnitL    : Real;
                                         intNumAbastec    : Integer;
                                         fltArqSaldo,
                                         fltArqEstoque    : Real;
                                         strArqICMS       : String;
                                         intPistaConsumo  : Variant;
                                         strTipo          : String = '';
                                         pArredonda       : Boolean = True) : Boolean;
var
  strCodRed,
  strDescProd : String;
  intProdID   : Integer;
  vArredonda  : integer;
  vCFOP       : String;
  vCST        : STRING;
  vPercRed    : double;

  function ItemJaExiste() : Boolean;
  begin
    Result := False;
//    if dDataCaixa <= 0 then Exit;
//    with fQryExiste do begin
//      Close;
//      SQL.Clear;
//      SQL.Add('SELECT COUNT(*) AS QTDE');
//      SQL.Add('  FROM ARQUIVO');
//      SQL.Add(' WHERE TAB_CODIGO   = ' + QuotedStr(strTabCodigo));
//      SQL.Add('   AND DBFTEMP_ID   = ' + IntToStr(intDBFTEMP_ID));
//      sql.Add('   AND IDNOTAFISCAL = ' + IntToStr(intIDNotaFiscal));
//      SQL.Add('   AND ARQ_VRUNITL  = ' + SubstituiVirgulaPorPontos(FormatFloat('###,###',fltArqVrUnitL)));
//      Open;
//      Result := (FieldByName('QTDE').AsInteger > 0);
//      Close;
//    end;
  end;
begin
  Result := false;

  if pArredonda then
    vArredonda := 0  //arredonda
  else
    vArredonda := 1; //trunca
  try
    if Trim(strArqICMS) = '' then begin
       strArqICMS := 'FF';
    end;
    with fQryItens do begin
      Close;

      //não pode vir zerado, caso venha tem que retornar falso;
      if intIDNotaFiscal = 0 then begin
        Result := False;
        exit;
      end;

      ///Setando o sequencial dos Itens
      SQL.Clear;
      SQL.Add('SELECT COALESCE(MAX(ARQ_ITEM),0) + 1 AS CONT');
      SQL.Add('  FROM ARQUIVO');
      SQL.Add(' WHERE IDNOTAFISCAL = ' + IntToStr(intIDNotaFiscal));
      Open;
      intArqItem := FieldByName('CONT').AsInteger;

      //Buscando informações do produto
      SQL.Clear;
      SQL.Add('SELECT T.ICM_PERC,              ');
      SQL.Add('	      T.TAB_DESCR,             ');
      SQL.Add('	      T.PRODUTO_ID,            ');
      SQL.Add('	      T.TAB_CODRED,            ');
      SQL.Add('	      T.CodSituacaoTributaria, ');
      SQL.Add('	      T.PERC_RED_ICMS          ');
      SQL.Add('  FROM TABELA T  ');
      SQL.Add(' WHERE T.TAB_CODIGO = ' + QuotedStr(Trim(strTabCodigo)));
      Open;
      strArqICMS  := Trim(FieldByName('ICM_PERC').AsString);
      strDescProd := Trim(FieldByName('TAB_DESCR').AsString);
      intProdID   := FieldByName('PRODUTO_ID').AsInteger;
      strCodRed   := Trim(FieldByName('TAB_CODRED').AsString);
      vCST        := Trim(FieldByName('CodSituacaoTributaria').AsString);
      vPercRed    := FieldByName('PERC_RED_ICMS').AsFloat;

      if (Trim(strIdentificador) = '') then begin
        SQL.Clear;
        SQL.Add('SELECT IDENTIFICADOR');
        SQL.Add('  FROM VENDEDOR');
        SQL.Add(' WHERE VEN_CODVEND = (SELECT VEN_CODVEND_PADRAO FROM CONFIG)');
        Open;
        strIdentificador := Trim(FieldByName('IDENTIFICADOR').AsString);
      end;

      if (not ItemJaExiste()) then begin
         //Inserindo o produto
        SQL.Clear;
        SQL.Add('INSERT INTO ARQUIVO(IDNOTAFISCAL,');
        SQL.Add('                    TAB_CODIGO,');
        SQL.Add('                    ARQ_ITEM,');
        SQL.Add('                    ARQ_QTDE,');
        SQL.Add('                    ARQ_VRUNIT,');
        SQL.Add('                    ARQ_SALDO,');
        SQL.Add('                    ARQ_VRUNITL,');
        SQL.Add('                    ARQ_ESTOQ,');
        SQL.Add('                    ARQ_ICMS,');
        SQL.Add('                    PISTACONSUMO,');
        SQL.Add('                    PLACAMOV_ID,');
        SQL.Add('                    NUMABASTEC,');
        if Copy(Trim(strIdentificador),1,20) <> '' then begin
          SQL.Add('                    IDENTIFICADOR,');
        end;
        SQL.Add('                    PRODUTO_ID,');
        SQL.Add('                    TAB_CODRED,');
        SQL.Add('                    PRODUTO,');
        SQL.Add('                    VALORUNITARIOCUSTO,');
        SQL.Add('                    VALORUNITARIOMEDIO,');
        SQL.Add('                    DBFTEMP_ID,');
        SQL.ADD('                    CFOP,');
        SQL.ADD('                    ICMSBSVALOR,');
        SQL.ADD('                    ICMSALIQ,');
        SQL.ADD('                    ICMSVALOR,');
        SQL.ADD('                    PERC_RED_ICMS,');
        SQL.ADD('                    CLASSE_FISCAL,');
        SQL.Add('                    ROUND)');
        SQL.Add('VALUES(:IDNOTAFISCAL,');
        SQL.Add('       :TAB_CODIGO,');
        SQL.Add('       :ARQ_ITEM,');
        SQL.Add('       :ARQ_QTDE,');
        SQL.Add('       :ARQ_VRUNIT,');
        SQL.Add('       :ARQ_SALDO,');
        SQL.Add('       :ARQ_VRUNITL,');
        SQL.Add('       :ARQ_ESTOQ,');
        SQL.Add('       :ARQ_ICMS,');
        SQL.Add('       :PISTACONSUMO,');
        SQL.Add('       :PLACAMOV_ID,');
        SQL.Add('       :NUMABASTEC,');
        if Copy(Trim(strIdentificador),1,20) <> '' then begin
          SQL.Add('       :IDENTIFICADOR,');
        end;
        SQL.Add('       :PRODUTOID,');
        SQL.Add('       :CODRED,');
        SQL.Add('       :PRODUTO,');
        SQL.Add('       :VALORUNITARIOCUSTO,');
        SQL.Add('       :VALORUNITARIOMEDIO,');
        SQL.Add('       :DBFTEMP_ID,');
        SQL.ADD('       :CFOP,');
        SQL.ADD('       :ICMSBSVALOR,');
        SQL.ADD('       :ICMSALIQ,');
        SQL.ADD('       :ICMSVALOR,');
        SQL.ADD('       :PERC_RED_ICMS,');
        SQL.ADD('       :CLASSE_FISCAL,');
        SQL.Add('       :ROUND)');
        Parameters.ParamByName('IDNOTAFISCAL').Value       := intIDNotaFiscal;
        Parameters.ParamByName('TAB_CODIGO').Value         := Copy(Trim(strTabCodigo),1,20);
        Parameters.ParamByName('ARQ_ITEM').Value           := intArqItem;
        Parameters.ParamByName('ARQ_QTDE').Value           := fltQtde;
        Parameters.ParamByName('ARQ_VRUNIT').Value         := fltArqVrUnit;
        Parameters.ParamByName('ARQ_SALDO').Value          := fltArqSaldo;

        Parameters.ParamByName('ARQ_VRUNITL').DataType     := ftFloat;
        Parameters.ParamByName('ARQ_VRUNITL').Size         := 15;
        Parameters.ParamByName('ARQ_VRUNITL').Precision    := 10;
        Parameters.ParamByName('ARQ_VRUNITL').Value        := fltArqVrUnitL;

        Parameters.ParamByName('ARQ_ESTOQ').Value          := fltArqEstoque;
        Parameters.ParamByName('ARQ_ICMS').Value           := Copy(Trim(strArqICMS),1,4);
        Parameters.ParamByName('PISTACONSUMO').Value       := intPistaConsumo;
        Parameters.ParamByName('PLACAMOV_ID').Value        := intPlacaMov_ID;
        Parameters.ParamByName('NUMABASTEC').Value         := intNumAbastec;
        if Copy(Trim(strIdentificador),1,20) <> '' then begin
          Parameters.ParamByName('IDENTIFICADOR').Value      := Copy(Trim(strIdentificador),1,20);
        end;
        Parameters.ParamByName('PRODUTOID').Value          := intProdID;
        if trim(strCodRed) = '' then
          Parameters.ParamByName('CODRED').Value := '0'
        else
          Parameters.ParamByName('CODRED').Value           := Copy(Trim(strCodRed),1,50);
        Parameters.ParamByName('PRODUTO').Value            := Copy(Trim(strDescProd),1,40);
        Parameters.ParamByName('VALORUNITARIOCUSTO').Value := RetornaCustoUnitario(Copy(Trim(strTabCodigo),1,20));
        Parameters.ParamByName('VALORUNITARIOMEDIO').Value := RetornaCustoMedio(Copy(Trim(strTabCodigo),1,20));
        Parameters.ParamByName('DBFTEMP_ID').Value         := intDBFTEMP_ID;
        Parameters.ParamByName('CFOP').Value               := fInterfaceNfCe.GetCFOP_ItensNFe(strTabCodigo,intArqItem-1);
        if fInterfaceNfCe.GetAliquotaICMS_ItensNFe(strTabCodigo,intArqItem-1) > 0 then
          Parameters.ParamByName('ICMSBSVALOR').Value      := fInterfaceNfCe.GetValorBaseICMS_ItensNFe(strTabCodigo,intArqItem-1);
        Parameters.ParamByName('ICMSALIQ').Value           := fInterfaceNfCe.GetAliquotaICMS_ItensNFe(strTabCodigo,intArqItem-1);
        Parameters.ParamByName('ICMSVALOR').Value          := fInterfaceNfCe.GetValorICMS_ItensNFe(strTabCodigo,intArqItem-1);
        Parameters.ParamByName('PERC_RED_ICMS').Value      := vPercRed;
        Parameters.ParamByName('CLASSE_FISCAL').Value      := vCST;
        Parameters.ParamByName('ROUND').Value              := vArredonda;
        ExecSQL();
      end else begin

      end;
      Close;
    end;
    Result := true;
  except
    on e: Exception do begin
      Result := false;
      MSGAtencao(e.Message);
    end;
  end;
end;

function TImprimeItens.GetArquivoID(pIdNF: Integer; pCodProduto: string; var pItem: Integer): integer;
VAR
  I: Integer;
  vItem : integer;
begin
  with fQryItens do begin
    Close;
    SQL.Clear;
    SQL.Add('SELECT COALESCE(MAX(ARQ_ITEM),0) AS ITEM');
    SQL.Add('  FROM ARQUIVO');
    SQL.Add(' WHERE IDNOTAFISCAL = ' + IntToStr(pIdNF));
    Open;
    if not IsEmpty then begin
      vItem := FieldByName('ITEM').AsInteger;

      Close;
      SQL.Clear;
      sql.Add('SELECT ARQUIVO_ID ');
      SQL.Add('  FROM ARQUIVO ');
      SQL.Add(' WHERE IDNOTAFISCAL = ' + IntToStr(PIDNF));
      SQL.Add('   AND TAB_CODIGO = ''' + Trim(pCodProduto) + '''');
      SQL.Add('   AND ARQ_ITEM = ' + IntToStr(vItem));
      Open;

      if not IsEmpty then begin
        result := FieldByName('ARQUIVO_ID').AsInteger;
        pItem  := fListaItens.Count + 1;//vItem;
//        fQTdItemLista := fListaItens.Count;
      end;
    end;
  end;
end;


function TImprimeItens.GetCountLista: integer;
begin
  Result := fListaItens.Count;
end;

procedure TImprimeItens.SetItemCancelado(pPlacaMovId: Integer = 0; pArquivoId: Integer = 0; pCombustivel: Boolean=False);
var
  i : integer;
begin
  try
    if Assigned(fListaItens)  then begin
      for I := 0 to fListaItens.Count - 1 do begin
        if pCombustivel then begin
          if pPlacaMovId > 0 then begin
            if TItensImpresso(fListaItens[i]).PlacaMovId = pPlacamovId then begin
              TItensImpresso(fListaItens[i]).Cancelado := True;
              FVrTotalProdutos := fVrTotalProdutos - TItensImpresso(fListaItens[i]).Valor;
              ApagaItemArquivo(pCombustivel,pPlacamovId);
              fQtdePtoTotal := fQtdePtoTotal - TItensImpresso(fListaItens[i]).QtdPto;
            end;
          end;
        end
        else begin
          if pArquivoId > 0 then begin
            if TItensImpresso(fListaItens[i]).ArquivoID = pArquivoId then begin
              TItensImpresso(fListaItens[i]).Cancelado := True;
              FVrTotalProdutos := fVrTotalProdutos - TItensImpresso(fListaItens[i]).Valor;
              ApagaItemArquivo(pCombustivel, pArquivoId);
              fQtdePtoTotal := fQtdePtoTotal - TItensImpresso(fListaItens[i]).QtdPto;
            end;
          end;
        end;
      end;
    end;
  except
    //
  end;
end;

function TImprimeItens.ApagaItemArquivo(pCombustivel: Boolean; pDocID: integer): Boolean;
begin
   try
    with fQryItens do begin
      Close;
      SQL.Clear;
      SQL.Add('DELETE ARQUIVO');
      SQL.Add(' WHERE IDNOTAFISCAL = ' + IntToStr(fCampos.IDNotaFiscal));
      if pCombustivel then
        SQL.Add('   AND PLACAMOV_ID  = ' + IntToStr(pDocID))
      else
        SQL.Add('  AND ARQUIVO_ID = ' + IntToStr(pDocID));
      ExecSQL();
      result := true;
    end;
  except
    on e: Exception do begin
      Result := False;
      MSGAtencao('ERRO AO RETORNAR ITEM' + sLineBreak +
                  E.Message);
    end;
  end;
end;

Function TImprimeItens.SubstituiVirgulaPorPontos(sString: String): String;
Var
  i                      : Integer;
  sStrin                 : String;
Begin
  sStrin := '';
  For i := 1 To Length(sString) Do Begin
    If (sString[i] = ',') Then Begin
      sStrin := sStrin + '.';
    End Else Begin
      sStrin := sStrin + sString[i];
    End;
  End;
  Result := sStrin;

End;

procedure TImprimeItens.RetiraImageCupom(pIndex: Integer; var pImage: TImage);
var
  vItem : string;
  i     : integer;
begin
  vItem := FormatFloat('0000',pIndex);
  if fStringList.Count > 0 then begin
    for I := 0 to fStringList.Count - 1 do begin
      if Copy(fStringList[i],1,4) = vitem then begin
        fVALORTOTAL := StrToFloat(Copy(fStringList[i],41,11));
        fStringList.Delete(i);
        ImprimeImage(pImage, false, True);
        break;
      end;
    end;
  end;
end;

procedure TImprimeItens.MostraPrecoEspecial(var pPrecoEspecial: TLabel);
begin
  if fPrecoEspecial then
    pPrecoEspecial.Caption := 'Preço Especial: ' + Formatfloat('#,###.###', fValorUnitarioL)
  else
    pPrecoEspecial.Caption := '';

  pPrecoEspecial.Visible := fPrecoEspecial;
  pPrecoEspecial.Refresh;
end;

procedure TImprimeItens.ImprimeImage(var pImage: TImage; pImprimeCabecalho: Boolean = False; pCancelar: Boolean = False);
var
  StringList : TStrings;
  i : integer;
  vTotal : Double;
begin
  StringList := TStringList.Create;
  StringList.Add(' ');

  if pImprimeCabecalho then begin
    if Length(fCampos.CupomFiscal) > 1 then begin
      if not bUtilizaNFCE then
        fStringList.Add('         C U P O M    F I S C A L  ' + fCampos.CupomFiscal)
      else
        fStringList.Add('    NOTA FISCAL DO CONSUMIDOR ELETRÔNICA  ' + fCampos.CupomFiscal)
    end else begin
      if not bUtilizaNFCE then
        fStringList.Add('         C U P O M    F I S C A L  ')
      else
        fStringList.Add('  NOTA FISCAL DO CONSUMIDOR ELETRÔNICA ');
    end;
    fStringList.Add('Item  Descrição            Quantidade     Total');
    fStringList.Add('-----------------------------------------------');
  end;

  if (not pImprimeCabecalho) and (not pCancelar) then begin
    fStringList.Add(Format('%0.4d', [fItem]) +
                    Format('%24.22s', [Copy(fDescricaoProduto, 1, 21)]) +
                    Format('%9.3f', [fQUANTIDADE]) +
                    Format('%9.2f', [fVALORTOTAL]));

    fVrTotalProdutos := fVrTotalProdutos + fVALORTOTAL;
  end;

  if pCancelar then
    vTotal := fVrTotalProdutos - fVALORTOTAL
  else
    vTotal := fVrTotalProdutos;

  if fStringList.Count > 0 then begin
    for I := 0 to fStringList.Count - 1 do begin
      StringList.Add(fStringList.Strings[i]);
    end;
  end;

  StringList.Add('-----------------------------------------------');
  StringList.Add('T O T A L' + Format('%36.2f', [vTotal]));
  StringList.Add('');

  fVrTotal := vTotal;

  ImprimeCupom(pImage, StringList);
  pImage.Refresh;
end;

procedure TImprimeItens.ImprimeCupom(var pntPapel: TImage; Texto: TStrings);
var
  i,
  i1,
  tam,
  y,
  k : Integer;
  V : Array of TPoint;
begin
  Tam := 27;
  SetLength(v, pntPapel.Width);
  k   := Trunc(High(v) / 2) + 1;
  with pntPapel do begin
    y := 0;
    for i := Low(V) to Trunc(High(v) / 2) do begin
      if i mod 2 = 1 then begin
        V[i] := Point(i * 2, y)
      end else begin
        V[i] := Point(i * 2, y + 4);
      end;
    end;
    i1 := Trunc(High(v) / 2);
    if Texto.Count > Tam then begin
      k := Tam * Canvas.TextHeight('A') - 3;
    end else begin
      y := Texto.Count * Canvas.TextHeight('A') - 3;
    end;
    V[i1] := Point(i1 * 2, y);
    for i := Trunc(High(v) / 2) to High(v) do begin
      k := k - 1;
      if i mod 2 = 1 then begin
        V[i] := Point(k * 2, y)
      end else begin
        V[i] := Point(k * 2, y + 4);
      end;
    end;
    V[High(v)] := Point(0, y);
    Canvas.Polygon(v);
    if Texto.Count > tam then begin
      k := Texto.Count - Tam;
    end else begin
      k := 0;
      Height := Texto.Count * Canvas.TextHeight('A');
    end;
    Y                 := Canvas.TextHeight('A');
    Canvas.Font.Name  := 'Courier New';
    Canvas.Font.size  := 9;
    Canvas.Font.Color := clBlue;
    Canvas.Font.Style := [];
    for i := 0 to Texto.Count - 1 do begin
      if k > Texto.Count - 2 then begin
        Exit;
      end;
      Canvas.TextOut(10, (i * Y) + 5, Texto.Strings[k + 1]);
      k := k + 1;
      if Texto.Count > 3 then begin
        if k + 5 = Texto.Count then begin
          Canvas.Font.Color := clRed;
        end else begin
          Canvas.Font.Color := clBlue;
        end;
        if k + 3 >= Texto.Count then begin
          Canvas.Font.Style := [fsBold]
        end else begin
          Canvas.Font.Style := [];
        end;
      end;
    end;
  end;
  pntPapel.Enabled := False;
  pntPapel.Refresh;
  Application.ProcessMessages;
end;

function TImprimeItens.IniciaVendaNfce(pTpEmiss: Integer = 1): Boolean;
var
  vIndicador : integer;
  vNumeroNFce : integer;
  vNumeroNf   : integer; //usado para nfe emitida no pdv
//  vCodigo     : Integer;
  vInfoAdicional : string;
  vCFOP : String;
  CodCidIBGE : string;
  InscEstDest : String;
  vIndFinal : Integer;
  idDest : INTEGER;
  vEmissao : String;
begin
  try
    if not fJaIniciouVndNFCe then begin
      fJaIniciouVndNFCe := True;
      InsereStatusCupom;
    end
    else if pTpEmiss <> 9 then
      exit;

    if pTpEmiss = 1 then begin
      vNumeroNFce := StrtoInt(fCampos.CupomFiscal);

//      if fcampos.ModeloNFe = 55 then
//        vCodigo := ProximaNotaFiscal(True);
//      else
//        vCodigo := vNumeroNFce;

      vCFOP := IIf(sEmpresaUf = fCampos.Cli_UF,'5656','5667');

      InscEstDest := TrataIE_NFe(vIndicador,fCampos.Cli_InscEst,IntToStr(fCampos.ModeloNFe),vIndFinal); //DestinatarioIE


      IF (fCampos.CodCliente = sCodClientePadrao) then begin
        if not fCampos.CliSemCadastro then begin
          fCampos.CNPJCpf     := '';
          fCampos.NomeCliente := '';                        //DestinatarioNomeRazao
        end;
        fCampos.Cli_Fone    := '';                        //DestinatarioFone
        fCampos.Cli_CEP     := '';                        //DestinatarioCEP
        fCampos.CliEnderecoComercial := '';               //DestinatarioLogradouro
        fCampos.Cli_Bairro := '';                         //DestinatarioBairro
        CodCidIBGE         := '';                         //DestinatarioCidadeCod
        fCampos.Cli_Cidade := '';                         //DestinatarioCidade
        fCampos.Cli_UF     := sEmpresaUf;                 //DestinatarioUF
      end
      ELSE begin
//        If (sEmpresaUf <> fCampos.Cli_UF) and (fCampos.MODELONFE = 65) then begin
//          fCampos.Cli_Fone    := '';                        //DestinatarioFone
//          fCampos.Cli_CEP     := '';                        //DestinatarioCEP
//          fCampos.CliEnderecoComercial := '';               //DestinatarioLogradouro
//          fCampos.Cli_Bairro := '';                         //DestinatarioBairro
//          CodCidIBGE         := '';                         //DestinatarioCidadeCod
//          fCampos.Cli_Cidade := '';                         //DestinatarioCidade
//          fCampos.Cli_UF     := sEmpresaUf;                 //DestinatarioUF
//        end
//        else
          CodCidIBGE :=  BuscaCodigoCidadeIBGE(fCampos.Cli_Cid_Codigo);
      end;

      IF Length(TRIM(fCampos.NomeCliente))  <= 2 then
        fCampos.NomeCliente := 'CONSUMIDOR FINAL';

      if fCampos.ModeloNFe = 65 then begin
        idDest := 1;
        vNumeroNf := vNumeroNFce;
      end
      else begin
        if sEmpresaUf = fCampos.Cli_UF then
          idDest := 1
        else
          idDest := 1;

        vNumeroNf :=  ProximaNotaFiscal(True);
      end;

      if bVoltaUmaHora then
        vEmissao := FormatDateTime('DD/MM/YYYY HH:MM:SS',Now-0.04166)
      else
        vEmissao := FormatDateTime('DD/MM/YYYY HH:MM:SS',Now);

      Result := fInterfaceNfCe.CriaCapaNFe(['VENDA',                    //NaturezaOperacao
                                             fcampos.ModeloNFe,         //Modelo
                                             vNumeroNf,               //Codigo
                                             vNumeroNFce,               //Numero
                                             fCampos.NumSerieNFCe,      //Serie
                                             vEmissao,                  //Emissao
                                             FormatDateTime('DD/MM/YYYY HH:MM:SS',Now),//Saida
                                             1,                         //Tipo de operação 0 - entrada 1 - saida
                                             '01',                      //FormaPag
                                             1,                         //Finalidade
                                             idDest,                    //idDest
                                             1,  {vIndFinal}            //indFinal  Indica operação com Consumidor final  0=nao 1=consumidor final
                                             1,                         //indPres  1=Operação presencial;
                                             DM.EmpresasCNPJ.AsString,               //EmitenteCNPJ
                                             DM.EmpresasInscricaoEstadual.AsString, //EmitenteIE
                                             dm.EmpresasInscMunicipal.AsString,     //EmitenteIM
                                             DM.EmpresasCNAE.AsString,              //EmitenteCNAE
                                             DM.EmpresasRazaoSocial.AsString,      //EmitenteRazao
                                             DM.EmpresasNomeFantasia.AsString,    //EmitenteFantasia
                                             sEmpresaFone,                        //EmitenteFone
                                             sEmpresaCep,                         //EmitenteCEP
                                             COPY(DM.Empresasendereco.AsString,1,60),   //EmitenteLogradouro
                                             'S/N',                                     //EmitenteNumero
                                             sEmpresaComplemento,                       //EmitenteComplemento
                                             sEmpresaBairro,                            //EmitenteBairro
                                             BuscaCodigoCidadeIBGE(IEmpresaCodcidade),  //EmitenteCidadeCod
                                             sEmpresaCidade,                            //EmitenteCidade
                                             sEmpresaUf,                                //EmitenteUF
                                             1058,                                      //EmitentePaisCod
                                             'BRASIL',                                  //EmitentePais
                                             fCampos.CNPJCpf,                           //DestinatarioCNPJ
                                             InscEstDest,                               //DestinatarioIE
                                             fCampos.NomeCliente,                       //DestinatarioNomeRazao
                                             fCampos.Cli_Fone,                          //DestinatarioFone
                                             fCampos.Cli_CEP,                           //DestinatarioCEP
                                             fCampos.CliEnderecoComercial,              //DestinatarioLogradouro
                                             'S/N',                                     //DestinatarioNumero
                                             '',                                        //DestinatarioComplemento
                                             fCampos.Cli_Bairro,                        //DestinatarioBairro
                                             CodCidIBGE,                                //DestinatarioCidadeCod
                                             fCampos.Cli_Cidade,                        //DestinatarioCidade
                                             fCampos.Cli_UF,                            //DestinatarioUF
                                             vIndicador,                                //DestinatarioindIEDest
                                             '',
                                             '',
                                             StrtoInt(fCampos.CodCliente),
                                             iif(fcampos.ModeloNFe = 55,1,4),                                           //0=Sem geração de DANFE; 1=DANFE normal, Retrato; 2=DANFE normal, Paisagem; 3=DANFE Simplificado;
                                             vInfoAdicional,
                                             vCFOP,
                                             '',
                                             '',  //finfAdFisco - informação adicional ao fisco
                                             pTpEmiss,//Tipo de Emissao 1=Emissão normal 9=Contingência off-line da NFC-e;
                                             IntToStr(DM.EmpresasREG_TRIBUTARIO.AsInteger), //CRT - Regime Tributário
                                             FormatDateTime('HH:MM:SS',Now),                //datahora
                                             fCampos.DataCaixa,                             //datacaixa
                                             fcampos.IDNotaFiscal]);                        //idnotafiscal
    end
    else if pTpEmiss = 9 then begin
      IF fcampos.ModeloNFe = 65 THEN BEGIN
        Result :=  fInterfaceNfCe.SetTpEmiss(pTpEmiss);
        fInterfaceNfCe.SetdhCont;
        fInterfaceNfCe.SetxJust('Falha na conexao com a internet');
      END;
    end;
  except
    on e: Exception do begin
      Result := False;
      ShowMessage(e.Message + #13#10 + 'Não foi possível iniciar a Nfc-e, entre em contato com o suporte da Kontrol!');
    end;
  end;
end;

function TImprimeItens.AlteraStatusCupom(pValue: Integer): Boolean;
begin
  try
    with fQryStatusCupom do begin
      Close;
      SQL.Clear;
      SQL.Add(' UPDATE STATUSCUPOM');
      SQL.Add('    SET Status = :STATUS');
      SQL.Add('  WHERE COMPUTADOR = :Computador');
      SQL.Add('    AND ECF = :Ecf');
      SQL.Add('    AND NUMEROCUPOM = :NumeroCupom');
      Parameters.ParamByName('Computador').Value  := fCampos.NomeComputador;
      Parameters.ParamByName('Ecf').Value         := fCampos.NumSerieECF;
      Parameters.ParamByName('NumeroCupom').Value := fCampos.CupomFiscal;
      Parameters.ParamByName('Status').Value      := pValue;
      ExecSQL;
    end;
  except
    on e: Exception do begin
      Result := False;
      ShowMessage(e.Message + #13#10 + 'Não foi possível alterar o status do cupom, entre em contato com o suporte da Kontrol!');
    end;
  end;
end;

function TImprimeItens.InsereStatusCupom: Boolean;
begin
  try
    with fQryStatusCupom do begin
      Close;
      SQL.Clear;
      SQL.Add(' INSERT INTO STATUSCUPOM');
      SQL.Add('    (Computador,        ');
      SQL.Add('     Ecf,               ');
      SQL.Add('     NumeroCupom,       ');
      SQL.Add('     Status) VALUES (   ');
      SQL.Add('    :Computador,        ');
      SQL.Add('    :Ecf,               ');
      SQL.Add('    :NumeroCupom,       ');
      SQL.Add('    :Status)            ');
      Parameters.ParamByName('Computador').Value  := fCampos.NomeComputador;
      Parameters.ParamByName('Ecf').Value         := fCampos.NumSerieECF;
      Parameters.ParamByName('NumeroCupom').Value := fCampos.CupomFiscal;
      Parameters.ParamByName('Status').Value      := 1;
      ExecSQL;
    end;
  except
    on e: Exception do begin
      Result := False;
      ShowMessage(e.Message + #13#10 + 'Não foi possível gravar o status do cupom, entre em contato com o suporte da Kontrol!');
    end;
  end;
end;

function TImprimeItens.GravaNFCe(pNFCE: Boolean = False): Boolean;
begin
  try
    fInterfaceNfCe.SetTotalValorProduto(VrTotalProdutos);
    fInterfaceNfCe.SetTotalValorNota(VrTotal);
    result := fInterfaceNfCe.GravaNFe(pNFCE);
  except
    on e: Exception do begin
      Result := False;
      Atencao('Ocorreu um erro ao gravar a NFC-e! ' + sLineBreak +
               e.Message);
    end;
  end;
end;

function TImprimeItens.ApagaNFce: Boolean;
begin
  try
    result := fInterfaceNfCe.ApagaNFe;
  except
    on e: Exception do begin
      Result := False;
      Atencao('Ocorreu um erro ao apagar a NFC-e! ' + sLineBreak +
               e.Message);
    end;
  end;
end;

function TImprimeItens.EnviarOffLine: boolean;
var
  vAnterior : boolean;
begin
  try
    vAnterior := bEmitirNFCeOffLine;
    bEmitirNFCeOffLine := true;
//    raise exception.create('erro');
    Result := CriarEnviarNFCe;
    bEmitirNFCeOffLine := vAnterior;
  except
    on e: exception do begin
      bEmitirNFCeOffLine := vAnterior;
      result := false;
      atencao('Ocorreu um erro ao enviar a NFCe Off-line, a venda será retornada! ' + sLineBreak +
              'Entre em contato com o suporte da Kontrol! ' + sLineBreak +
              e.Message);
    end;
  end;
end;

function TImprimeItens.CriarEnviarNFCe: Boolean;
var
  vEnderecoXML   : string;
  vXML           : AnsiString;
  vCh_NFe        : AnsiString;
  vChNFe         : string;
  vMotivo        : string;
  vNumero_Recibo : string;
  vProtocolo     : string;
  sXml           : TStringList;
  vNumeroNF      : Integer;
  vModelo        : Integer;
  vSerie         : Integer;
  vTpEmissao     : Integer;
  vMsg           : AnsiString;
  vReenvia       : Integer;
  vEnviou        : Boolean;
  vDataHora  : string;
  vdigVal    : string;
  vStatus    : string;
  vMsgem     : String;
  i          : integer;
  vEmail     : STRING;
  vEnderecoPDF : String;

  label envia;
  label reenvia;


  function EnviarModoOffLine: boolean;
  begin
    //CONFIGURADO NA OPÇÃO MESTRE PARA EMITIR NFCE EM MODO OFF-LINE
    //NÃO ESQUEÇER QUE O XML DEVE SER ENVIADO EM 24 HORAS
    try
      AbreSplashNFe('Aguarde, criando a NFC-e em modo Off-Line...',fCampos.ModeloNFe);
      if IniciaVendaNfce(9) then begin
        if fInterfaceNfCe.SetFormaEmissao(9) then begin
          vEnderecoXML := '';
          vXML := '';
          result := fInterfaceNfCe.CriarNFe(1,fCampos.CupomFiscal,vEnderecoXML,vXML,IntToStr(fCampos.ModeloNFe),fCampos.NumSerieNFCe);
          SetEnderecoXML(StrToIntDef(fCampos.CupomFiscal,0),vEnderecoXML);
          vChNFe := SomenteNumeros(vEnderecoXML);
          SetChaveNFe(StrToIntDef(fCampos.CupomFiscal,0),vChNFe);
          AbreSplashNFe('Aguarde, Assinando a NFC-e em modo Off-Line...',fCampos.ModeloNFe);
          Result := IIf(Result, fInterfaceNfCe.AssinarXML(vEnderecoXML),False);
          FechaSplashNFe;
          AbreSplashNFe('Aguarde, validando a NFC-e em modo Off-Line...',fCampos.ModeloNFe);
          Result := IIf(Result, fInterfaceNfCe.ValidarNFe(vEnderecoXML), False);
          AlterarSTATUS_CTG('E',StrToIntDef(fCampos.CupomFiscal,0),4,IIf(Result,'P','N'),IIf(Result,'P','N'));
          FechaSplashNFe;
          if Result then begin
            fInterfaceNfCe.Busca_XML(vEnderecoXML,vXML);
            FechaSplashNFe;
            fInterfaceNfCe.ImprimirDanfeNFce(vXML,fCampos.ModeloNFe = 55);
            ImprimeComprovante;
            AlterarLiberadoAutomatico(1,StrToIntDef(fCampos.CupomFiscal,0));
            GravaHistoricoNfe(fCampos.CupomFiscal,
                              0,
                              vXml,
                              true,
                              IntToStr(fCampos.ModeloNFe),
                              fCampos.NumSerieNFCe);
          end;
        end;
      end;
    finally
      FechaSplashNFe;
    end;
  end;

  procedure GerarChaveNFe(pTpEmissao: Integer = 1);
  begin
    if trim(vChNFe) = '' then begin
      vModelo    := fCampos.ModeloNFe;
      vSerie     := StrToIntDef(fCampos.NumSerieNFCe,1);
      vTpEmissao := pTpEmissao;

      GerarChave(vCh_NFe,
                 StrToInt(Copy(BuscaCodigoCidadeIBGE(IEmpresaCodcidade),1,2)),
                 StrToInt(fCampos.CupomFiscal),
                 vModelo,
                 vSerie,
                 StrToInt(fCampos.CupomFiscal),
                 vTpEmissao,
                 Date,
                 sEmpresaCNPJ);
      vChNFe := SomenteNumeros(vCh_NFe);
    end;
  end;
begin
  try
    vReenvia := 0;
    vChNFe   := '';
    fInterfaceNfCe.SetfModelo;
    if MonitorAtivo then begin
      envia:
      Result := True;
      //quando a fomra de pagamento for TEF irá enviar OFF-line
      if (not bEmitirNFCeOffLine) then //and (not fVndUsouTEF) then
        Result := fInterfaceNfCe.CriarNFe(1,fCampos.CupomFiscal,vEnderecoXML,vXML,IntToStr(fCampos.ModeloNFe),fCampos.NumSerieNFCe);

      if Result then begin
        SetEnderecoXML(StrToIntDef(fCampos.CupomFiscal,0),vEnderecoXML);
        MSGAguarde(false);
        reenvia:
        if (not bEmitirNFCeOffLine) then begin// and (not fVndUsouTEF) then begin
//          raise exception.create('erro');
          vEnviou := fInterfaceNfCe.EnviarNFe(vEnderecoXML,vChNFe,vMotivo,vNumero_Recibo,vProtocolo,False,IntToStr(fCampos.ModeloNFe),fCampos.NumSerieNFCe);

          if (not vEnviou) and (vMotivo = 'L') then begin
            for i := 0 to 2 do begin
              sleep(2000);
              if vChNFe = '' then
                GerarChaveNFe;

              vEnviou := fInterfaceNfCe.NfeConsultaSituacaoNfe(vChNFe, vDataHora, vProtocolo, vdigVal, vMsgem, vStatus, false, True, false);
              if vStatus = '100' then begin
                vEnviou := true;
                Break;
              end
              else if em(vStatus,['110','205','301','302','303']) then begin
                case StrToInt(vStatus) of
                  110 : vMotivo := 'NS';  //'Uso Denegado'
                  205 : vMotivo := 'NS';  //'NF-e está denegada na base de dados da SEFAZ [nRec:999999999999999]'
                  301 : vMotivo := 'NE';  //'Uso Denegado: Irregularidade fiscal do emitente'
                  302 : vMotivo := 'ND';  //'Uso Denegado: Irregularidade fiscal do destinatário'; end;
                  303 : vMotivo := 'NU';  //'Uso Denegado: Destinatário não habilitado a operar na UF';
                end;
              end;
            end;
          end;

          if vEnviou then begin
            if EmituNfe('S',StrToIntDef(fCampos.CupomFiscal,0),vChNFe,vNumero_Recibo,vProtocolo) then begin
              AlterarSTATUS_NFE('V',StrToIntDef(fCampos.CupomFiscal,0));
              fInterfaceNfCe.Busca_XML(vEnderecoXML,vXML);
              fInterfaceNfCe.ImprimirDanfeNFce(vXML,fCampos.ModeloNFe = 55);
              ImprimeComprovante;
              AlterarLiberadoAutomatico(1,StrToIntDef(fCampos.CupomFiscal,0));

              if fCampos.EnviarEmail then begin
                vEmail := BuscaEmailCLIENTE(StrToInt(trim(fCampos.CodCliente)));
                if ValidarEMail(Trim(vEmail)) then begin
                  //cria o PDF
                  fInterfaceNfCe.Busca_XML(vEnderecoXML,vXML);
                  fInterfaceNfCe.imprimirDanfePDF(vXML,vEnderecoPDF);
                  if FindWindow('TAppBuilder', nil) = 0 then begin //quando estiver usando o delphi não irá enviar o email
                    fInterfaceNfCe.EnviarEmail(Trim(vEmail),vEnderecoXML,vEnderecoPDF);
                  end;
                end
              end;
            end;
          end
          else begin
            if vMotivo = 'D' then begin
              EmituNfe('S',StrToIntDef(fCampos.CupomFiscal,0),vChNFe);
              AlterarLiberadoAutomatico(1,StrToIntDef(fCampos.CupomFiscal,0));
            end
            else if vMotivo = 'I' then begin //NF-e Inutilizada na base da sefaz
              vNumeroNF := StrToIntDef(fCampos.CupomFiscal,0);
              AlteraNumeroNFe(vNumeroNF);
              if vNumeroNF <> StrToIntDef(fCampos.CupomFiscal,0) then
                fCampos.CupomFiscal := IntToStr(vNumeroNF);
              goto envia;
            end
            else if vMotivo = 'L' then begin //LOTE EM PROCESSAMENTO
              vChNFe         := '';
              vNumero_Recibo := '';
              vProtocolo     := '';
              GerarChaveNFe;
              EmituNfe('L',StrToIntDef(fCampos.CupomFiscal,0),vChNFe,vNumero_Recibo,vProtocolo);
//              ConsultaSituacao;
              //tem que ser impresso os comprovantes aqui pois
              //não sao ficam cadastrados no bd
              ImprimeComprovante;
              AlterarLiberadoAutomatico(1,StrToIntDef(fCampos.CupomFiscal,0));
            end
            else if vMotivo = 'R' then begin //deve ser reenviado o mesmo xml, no caso de reenviar 3x deve reenviar em offline
              inc(vReenvia);
              if vReenvia > 3 then begin
                vReenvia := 0;
                EnviarModoOffLine;
              end
              else begin
                goto reenvia;
              end;
            end
            else if vMotivo = 'E' then begin
              //DEU ERRO E DEVE SER EMITIDO EM MODO OFF-LINE
              //NÃO ESQUEÇER QUE O XML DEVE SER ENVIADO EM 24 HORAS
              try
                AbreSplashNFe('Aguarde, criando a NFC-e em modo Off-Line...',fCampos.ModeloNFe);
                if IniciaVendaNfce(9) then begin
                  if fInterfaceNfCe.SetFormaEmissao(9) then begin
                    vEnderecoXML := '';
                    vXML := '';
                    result := fInterfaceNfCe.CriarNFe(1,fCampos.CupomFiscal,vEnderecoXML,vXML,IntToStr(fCampos.ModeloNFe),fCampos.NumSerieNFCe);
                    SetEnderecoXML(StrToIntDef(fCampos.CupomFiscal,0),vEnderecoXML);
                    AbreSplashNFe('Aguarde, Assinando a NFC-e em modo Off-Line...',fCampos.ModeloNFe);
                    Result := IIf(Result, fInterfaceNfCe.AssinarXML(vEnderecoXML),False);
                    FechaSplashNFe;
                    AbreSplashNFe('Aguarde, validando a NFC-e em modo Off-Line...',fCampos.ModeloNFe);
                    Result := IIf(Result, fInterfaceNfCe.ValidarNFe(vEnderecoXML), False);
                    vChNFe := SomenteNumeros(vEnderecoXML);
                    if length(vChNFE) <> 44 then begin
                      vChNFe := '';
                      GerarChaveNFe(9);
                    end;
                    SetChaveNFe(StrToIntDef(fCampos.CupomFiscal,0),vChNFe);
                    AlterarSTATUS_CTG('E',StrToIntDef(fCampos.CupomFiscal,0),4,IIf(Result,'P','N'),IIf(Result,'P','N'));
                    FechaSplashNFe;
                    if Result then begin
                      fInterfaceNfCe.Busca_XML(vEnderecoXML,vXML);
                      FechaSplashNFe;
                      fInterfaceNfCe.ImprimirDanfeNFce(vXML,fCampos.ModeloNFe = 55);
                      ImprimeComprovante;
                      AlterarLiberadoAutomatico(1,StrToIntDef(fCampos.CupomFiscal,0));
                    end;
                  end;
                end;
              finally
                FechaSplashNFe;
              end;                                                //NU - Uso Denegado: Destinatário não habilitado a operar na UF
            end                                                   //NE - Uso Denegado: Irregularidade fiscal do emitente
            else if em(vMotivo,['NE','ND','NS','NU']) then begin  //ND - Uso Denegado = Irregularidade fiscal do DESTINATÁRIO
                                                                  //NS - Uso Denegado = NF-e  está  denegada  na  base  de  dados  da  SEFAZ
              AlterarSTATUS_NFE('D',StrToIntDef(fCampos.CupomFiscal,0));
              AlterarTIPO_DENEGADA(vMotivo,StrToIntDef(fCampos.CupomFiscal,0),IntToStr(fCampos.ModeloNFe),fCampos.NumSerieNFCe);
              vNumeroNF := StrToIntDef(fCampos.CupomFiscal,0);
              RetornaVenda;

              if vMotivo = 'NU' then
                vMsgem := 'Destinatário não habilitado a operar na UF!'
              else If vMotivo = 'NE' then
                vMsgem :=  'Emitente com irregularidade fiscal!'
              else If vMotivo = 'ND' then
                vMsgem :=  'Destinatário com irregularidade fiscal!'
              else If vMotivo = 'NS' then
                vMsgem :=  'NF-e  está  denegada  na  base  de  dados  da  SEFAZ';


              Atencao(vMsgem + sLineBreak +
                     'Corrija a irregularidade e depois emita a ' + iif(fCampos.ModeloNFe=65,'NFC-e!','NF-e!') + sLineBreak +
                     'A venda será retornada!');
            end
            else begin
              result := EnviarModoOffLine;
//              Result := False;
//              if not fCampos.IniciouComF6 then
//                RetornaPlacaMov;
//              ApagaBasicoArquivo;
//              Atencao('Ocorreu um erro ao enviar a NFC-e! ');
            end;
          end;
        end
        else begin
          //CONFIGURADO NA OPÇÃO MESTRE PARA EMITIR NFCE EM MODO OFF-LINE
          //NÃO ESQUEÇER QUE O XML DEVE SER ENVIADO EM 24 HORAS
          result := EnviarModoOffLine;
        end;
      end;
    end;
  except
    on e: Exception do begin
      Result := False;
      Atencao('Ocorreu um erro ao enviar a NFC-e! ' + sLineBreak +
               e.Message);
    end;
  end;
end;


//function TImprimeItens.ConsultaSituacao: Boolean;
//var
//  vDataHora  : string;
//  vProtocolo : string;
//  vdigVal    : string;
//  vArquivo   : String;
//  vTemProt   : Boolean;
//  vMSG       : String;
//  vStatus    : string;
//  vXML       : AnsiString;
//  vChaveNFe  : String;
//begin
//  try
//    vChaveNFe := fConsulta.fieldByName('NUMERO_NFE').asString;
//    IF VerificaNUMERO_NFE(vChaveNFe) then begin
//      if fInterfaceNfCe.NfeConsultaSituacaoNfe(vChaveNFe,vDataHora,vProtocolo,vdigVal,vMSG,vStatus,False,False,True) then begin
//        AlterarSTATUS_NFE('V', StrToIntDef(fConsulta.fieldByName('NOTA_FISCAL_NFE').asString,0),'65',fConsulta.fieldByName('SERIE').asString);
//        if trim(vProtocolo) <> '' then
//          AlterarNUMERO_PROTOCOLO(vProtocolo,StrToIntDef(fConsulta.fieldByName('NOTA_FISCAL_NFE').asString,0),'65',fConsulta.fieldByName('SERIE').asString);
//          //TEM QUE SER INCLUIDO O XML DE RETORNO COM O PROTOCOLO NA TABELA NFE_HISTORICOXLM
//          //PRIMEIRO CONSULTA NA TABELA PARA VERIFICAR SE TEM O XML DE RETORNO E COM O PROTOCOLO
//        BuscaXMLHistorico(vXml,'65',fConsulta.fieldByName('SERIE').asString);
//        vTemProt := (Pos('<PROTNFE',UpperCase(vXml)) > 0);
//        IF Pos('<PROTNFE',UpperCase(vXml)) <= 0 then begin
//          vArquivo := copy(fConsulta.fieldByName('ENDERECO_XML').asString,1,pos('\ARQUIVOS\',UpperCase(fConsulta.fieldByName('ENDERECO_XML').asString))) +
//                      'Arquivos\NFCe\' + vChaveNFe + '-nfe.xml';
//          //SENAO BUSCA PELO ACBR E VERIFICA O PROTOCOLO
//          fInterfaceNfCe.Busca_XML(vArquivo,vXml,true);
//          IF Pos('<PROTNFE',UpperCase(vXml)) <= 0 then begin
//            //SENÃO PEGA O XMLENVIADO E INCLUI O PROTOCOLO
//            fInterfaceNfCe.TrataXML(vXml,vChaveNFe,vDataHora,vProtocolo,vdigVal);
//          end;
//        end;
//        if not vTemProt then
//          IF Pos('<PROTNFE',UpperCase(vXml)) > 0 then
//            GravaRetornoXML(fConsulta.fieldByName('NOTA_FISCAL_NFE').asString,vXml);
//
//         Result := True;
//
//      end
//      else begin
//        if vStatus = '217' then begin //Não Consta na base de dados a sefaz
//          AlterarSTATUS_NFE('P', StrToIntDef(fConsulta.fieldByName('NOTA_FISCAL_NFE').asString,0),'65',fConsulta.fieldByName('SERIE').asString);
//        end;
//      end;
//
//      try
//        //verifica se tem o historico gravado o xml
//        vXml:='';
//        if fInterfaceNfCe.Busca_XML(fConsulta.fieldByName('ENDERECO_XML').asString,vXml,True) then begin
//          NFEInterfaceV3.GravaHistoricoNfe(fConsulta.fieldByName('NOTA_FISCAL_NFE').asString,
//                                           StrToIntDef(Copy(vxml,Pos('<cStat>',vXml)+7,3),0),
//                                           vXml,
//                                           true,
//                                           '65',
//                                           fConsulta.fieldByName('SERIE').asString);
//        end;
//    //    VerificaDadosNfe;
//      except
//        //
//      end;
//    end;
//  except
//    //
//  end;
//end;

procedure TImprimeItens.AddFormaPagNfe(pFormaPag: array of Variant);
begin
  fInterfaceNfCe.AddFormaPagNfe(pFormaPag);
end;

procedure TImprimeItens.SetTotTrib(pTotTrib : String);
begin
  fInterfaceNfCe.SetTotTrib(pTotTrib);
end;

procedure TImprimeItens.SetUtilizandoTef(const Value: Boolean);
begin
  fUtilizandoTef := value;
  if Value then
    fVndUsouTEF := Value;
end;

function TImprimeItens.EmituNfe(sEmitiu : String; iNumeroNf : Integer; vChNFe: string = ''; pNumero_Recibo : string = ''; pProtocolo : String = ''): Boolean;
VAR
  vStatusNFE : STRING;
begin
  TRY
    result := false;
    if sEmitiu = 'S' then
      vStatusNFE := 'N'
    else if sEmitiu = 'L' then
      vStatusNFE := 'L'
    else if sEmitiu <> 'S' then
      vStatusNFE := 'P';

    with DM do begin
      qryAux3.Close;
      qryAux3.SQL.Clear;
      qryAux3.SQL.Add('UPDATE NOTAFISCAL');
      qryAux3.SQL.Add('   SET EMITIU     = :PEMITIU,');
      qryAux3.SQL.Add('       STATUS_NFE = :PSTATUS_NFE,');
      qryAux3.SQL.Add('       NUMERO_NFE = :NUMERO_NFE,');
      qryAux3.SQL.Add('       NUMERO_RECIBO = :NUMERO_RECIBO,');
      qryAux3.SQL.Add('       NUMERO_PROTOCOLO = :PPROTOCOLO');
      if fcampos.ModeloNFe = 55 then begin
        qryAux3.SQL.Add(' WHERE NOTA_FISCAL_NFE   = :PNUMERONF');
        qryAux3.SQL.Add('   AND SERIE = ''001'' ')
      end else begin
        qryAux3.SQL.Add(' WHERE NUMERONF   = :PNUMERONF');
        qryAux3.SQL.Add('   AND SERIE = ' + fCampos.NumSerieNFCe);
      end;
      qryAux3.SQL.Add('   AND MODELO = ''' + IntToStr(fCampos.ModeloNFe)+ ''' ');
      qryAux3.Parameters.ParamByName('PEMITIU').Value       := sEmitiu;
      qryAux3.Parameters.ParamByName('PSTATUS_NFE').Value   := vStatusNFE;//IIf(sEmitiu = 'S','N','P');
      qryAux3.Parameters.ParamByName('NUMERO_NFE').Value    := vChNFe;
      qryAux3.Parameters.ParamByName('NUMERO_RECIBO').Value := pNumero_Recibo;
      qryAux3.Parameters.ParamByName('PNUMERONF').Value     := iNumeroNf;
      qryAux3.Parameters.ParamByName('PPROTOCOLO').Value    := pProtocolo;
      qryAux3.ExecSQL;
      result := true;
    end;
  except
    ON E: Exception DO begin
      result := false;
      ShowMessage(E.Message);
    end;
  END;
end;

procedure TImprimeItens.SetEnderecoXML(pNumeroNf : Integer; pEnderecoXML : string = '');
begin
  TRY
    with DM do begin
      qryAux3.Close;
      qryAux3.SQL.Clear;
      qryAux3.SQL.Add('UPDATE NOTAFISCAL');
      qryAux3.SQL.Add('   SET ENDERECO_XML = :PENDERECOXML');
      if fcampos.ModeloNFe = 55 then begin
        qryAux3.SQL.Add(' WHERE NOTA_FISCAL_NFE = :PNUMERONF');
        qryAux3.SQL.Add('   AND SERIE = ''001'' ')
      end else begin
        qryAux3.SQL.Add(' WHERE NUMERONF = :PNUMERONF');
        qryAux3.SQL.Add('   AND SERIE = ' + fCampos.NumSerieNFCe);
      end;
      qryAux3.SQL.Add('   AND MODELO = '''+IntToStr(fCampos.ModeloNFe)+''' ');
      qryAux3.Parameters[0].Value := pEnderecoXML;
      qryAux3.Parameters[1].Value := pNumeroNf;
      qryAux3.ExecSQL;
    end;
  except
//    ON E: Exception DO begin
//      ShowMessage(E.Message);
//    end;
  END;
end;

function TImprimeItens.SetChaveNFe(pNumeroNf : Integer; pChaveNfe : string): boolean;
begin
  TRY
    result := false;
    with DM do begin
      if trim(pChaveNfe) <> '' then begin
        qryAux3.Close;
        qryAux3.SQL.Clear;
        qryAux3.SQL.Add('UPDATE NOTAFISCAL');
        qryAux3.SQL.Add('   SET NUMERO_NFE = :PCHAVE');
        if fcampos.ModeloNFe = 55 then begin
          qryAux3.SQL.Add(' WHERE NOTA_FISCAL_NFE = :PNUMERONF');
          qryAux3.SQL.Add('   AND SERIE = ''001''')
        end else begin
          qryAux3.SQL.Add(' WHERE NUMERONF = :PNUMERONF');
          qryAux3.SQL.Add('   AND SERIE = ' + fCampos.NumSerieNFCe);
        end;
        qryAux3.SQL.Add('   AND MODELO = '''+IntToStr(fCampos.ModeloNFe)+''' ');
        qryAux3.Parameters[0].Value := pChaveNfe;
        qryAux3.Parameters[1].Value := pNumeroNf;
        qryAux3.ExecSQL;
        result := true;
      end;
    end;
  except
    result := false;
  END;
end;

procedure TImprimeItens.AlteraNumeroNFe(var pNotaFiscal: Integer);
var
  vNotaFiscal : integer;
  vQtde       : integer;//qtd max de loop 50
begin
  try
    ///Pegando o número da próxima nota fiscal e validando se já está sendo utiliza
    ///enquanto não achar um numero de nfce sem q tenha gerado uma nfce irá ficar tentand
    vQtde := 0;
    while vQtde <= 50 do begin
      Inc(vQtde);
      vNotaFiscal := DM.ProximoIdNFcE();
      with DM.qryGeral do begin
        Close;
        SQL.Clear;
        SQL.Add('SELECT');
        SQL.Add('       NOTA_FISCAL_NFE');
        SQL.Add('  FROM NOTAFISCAL');
        SQL.Add(' WHERE NOTA_FISCAL_NFE = ' + IntToStr(vNotaFiscal));
        SQL.Add('   AND MODELO = '''+IntToStr(fCampos.ModeloNFe)+''' ');
        SQL.Add('   AND SERIE = ' + fCampos.NumSerieNFCe);
        Open;
      end;
      if dm.qryGeral.IsEmpty then begin
        with dm.qryAux3 do begin
          Close;
          SQL.Clear;
          SQL.Add(' UPDATE NOTAFISCAL');
          sQL.Add('    SET ENDERECO_XML = '''', ');
          sQL.Add('        NUMERO_NFE = '''', ' );
          SQL.Add('        NUMERONF = ' + IntToStr(vNotaFiscal) + ',');
          SQL.Add('        NOTA_FISCAL_NFE = ' + IntToStr(vNotaFiscal));
          SQL.Add('  WHERE NUMERONF   = ' + IntToStr(pNotaFiscal));
          SQL.Add('   AND MODELO = '''+IntToStr(fCampos.ModeloNFe)+''' ');
          SQL.Add('   AND SERIE = ' + fCampos.NumSerieNFCe);
          ExecSQL;
          pNotaFiscal := vNotaFiscal;
          break;
        end;
      end;
    end;
  except
    ON E: Exception DO begin
      ShowMessage(E.Message);
    end;
  END;
end;


function TImprimeItens.AlterarLiberadoAutomatico(pValor: Integer; pNumeroNF: integer): Boolean;
var
  vValor : Integer;
  vNumeroNF : iNTEGER;
  qryAuto : TADOQuery;
begin
  try
    try
      vValor := StrToIntDef(IntToStr(pvalor),1);
    except
      vValor := 1;
    end;

    try
      if pNumeroNF > 0 then
        vNumeroNF := pNumeroNF
      else if StrToInt(fCampos.CupomFiscal) > 0 then
        vNumeroNF := StrToInt(fCampos.CupomFiscal);
    except
      Result := false;
      exit;
    end;

    result := false;

    qryAuto := TADOQuery.Create(nil);
    qryAuto.Connection := dm.ADOconexao;

    with qryAuto do begin
      SQL.Clear;
      Close;
      SQL.Add('UPDATE NOTAFISCAL ');
      SQL.Add('   SET LiberadoAutomatico = :VALOR ');
      SQL.Add(' WHERE NUMERONF = :PNUMERONF');
      SQL.Add('   AND MODELO = '''+IntToStr(fCampos.ModeloNFe)+''' ');
      SQL.Add('   AND SERIE = ' + iif(fcampos.modelonfe=55,'001',fCampos.NumSerieNFCe));
      Parameters.ParamByName('VALOR').Value := vValor;
      Parameters.ParamByName('PNUMERONF').Value := vNumeroNF;
      ExecSQL;
    end;
    result := true;
  except
    on e: Exception do begin
      ShowMessage('Ocorreu um erro ao alterar o Liberação Automatica da nota' + #13 +
                  'entre em contato com o suporte da Kontrol!' + #13 +
                   e.Message);
      result := false;
    end;
  end;
end;

function TImprimeItens.AlterarSTATUS_NFE(valor: string; pNumeroNF: integer): Boolean;
begin
  try
    result := false;
    with DM.qryGeral do begin
      SQL.Clear;
      Close;
      SQL.Add('UPDATE NOTAFISCAL ');
      SQL.Add('   SET STATUS_NFE   = :STATUS,');
      SQL.Add('       CONTINGENCIA = 1 ');
      if fcampos.ModeloNFe = 55 then begin
        SQL.Add(' WHERE NOTA_FISCAL_NFE = :PNUMERONF');
        SQL.Add('   AND SERIE = ''001'' ');
      end else begin
        SQL.Add(' WHERE NUMERONF = :PNUMERONF');
        SQL.Add('   AND SERIE = ' + fCampos.NumSerieNFCe);
      end;
      SQL.Add('   AND MODELO = '''+IntToStr(fCampos.ModeloNFe)+''' ');
      Parameters.ParamByName('STATUS').Value := valor;
      Parameters.ParamByName('PNUMERONF').Value := pNumeroNF;
      ExecSQL;
    end;
    result := true;
  except
    on e: Exception do begin
      ShowMessage('Ocorreu um erro ao alterar o status da nota' + #13 +
                  'entre em contato com o suporte da Kontrol!' + #13 +
                   e.Message);
      result := false;
    end;
  end;

end;

procedure TImprimeItens.AlterarTIPO_DENEGADA(valor     : string;
                                             pNumeroNF : integer;
                                             pModelo   : string;
                                             pSerie    : String);
var
  vQryTIPO_DENEGADA : TADOQuery;
begin
  try
    vQryTIPO_DENEGADA := TADOQuery.Create(nil);
    vQryTIPO_DENEGADA.Connection := dm.ADOconexao;

    vQryTIPO_DENEGADA.SQL.Clear;
    vQryTIPO_DENEGADA.Close;
    vQryTIPO_DENEGADA.SQL.Add('UPDATE NOTAFISCAL ');
    vQryTIPO_DENEGADA.SQL.Add('   SET TIPO_DENEGADA = :TIPO_DENEGADA');
    vQryTIPO_DENEGADA.SQL.Add(' WHERE NUMERONF = :PNUMERONF');
    vQryTIPO_DENEGADA.SQL.Add('   AND SERIE  = :SERIE');
    vQryTIPO_DENEGADA.SQL.Add('   AND MODELO = :MODELO');
    vQryTIPO_DENEGADA.Parameters.ParamByName('TIPO_DENEGADA').Value := valor;
    vQryTIPO_DENEGADA.Parameters.ParamByName('PNUMERONF').Value := pNumeroNF;
    vQryTIPO_DENEGADA.Parameters.ParamByName('SERIE').Value := pSerie;
    vQryTIPO_DENEGADA.Parameters.ParamByName('MODELO').Value := pModelo;
    vQryTIPO_DENEGADA.ExecSQL;

    FreeAndNil(vQryTIPO_DENEGADA);
  except
    on e: Exception do begin
      ShowMessage('Ocorreu um erro ao alterar o tipo_denegada da nota' + #13 +
                  'entre em contato com o suporte da Kontrol!' + #13 +
                   e.Message);
    end;
  end;

end;

function TImprimeItens.AlterarSTATUS_CTG(sEmitiu        : String;
                                         iNumeroNf      : Integer;
                                         pContingencia  : Integer;
                                         pSTATUS_NFE    : string = '';
                                         pSTATUS_CTG    : string = '';
//                                         vChNFe         : string = '';
                                         pNumero_Recibo : string = '';
                                         pProtocolo     : String = ''): Boolean;
begin
  TRY
    result := false;
    with DM do begin
      qryAux3.Close;
      qryAux3.SQL.Clear;
      qryAux3.SQL.Add('UPDATE NOTAFISCAL');
      qryAux3.SQL.Add('   SET EMITIU     = :PEMITIU,');
      qryAux3.SQL.Add('       STATUS_NFE = :PSTATUS_NFE,');
//      qryAux3.SQL.Add('       NUMERO_NFE = :NUMERO_NFE,');
      qryAux3.SQL.Add('       NUMERO_RECIBO = :NUMERO_RECIBO,');
      qryAux3.SQL.Add('       NUMERO_PROTOCOLO = :PPROTOCOLO,');
      qryAux3.SQL.Add('       CONTINGENCIA = :PCONTINGENCIA,');
      qryAux3.SQL.Add('       STATUS_CTG = :PSTATUS_CTG, ');
      qryAux3.SQL.Add('       LiberadoAutomatico = 1 ');
      qryAux3.SQL.Add(' WHERE NUMERONF   = :PNUMERONF');
      qryAux3.SQL.Add('   AND MODELO = '''+IntToStr(fCampos.ModeloNFe)+''' ');
      qryAux3.SQL.Add('   AND SERIE = ' + fCampos.NumSerieNFCe);
      qryAux3.Parameters.ParamByName('PEMITIU').Value       := sEmitiu;
      qryAux3.Parameters.ParamByName('PSTATUS_NFE').Value   := pSTATUS_NFE;
//      qryAux3.Parameters.ParamByName('NUMERO_NFE').Value    := vChNFe;
      qryAux3.Parameters.ParamByName('NUMERO_RECIBO').Value := pNumero_Recibo;
      qryAux3.Parameters.ParamByName('PNUMERONF').Value     := iNumeroNf;
      qryAux3.Parameters.ParamByName('PPROTOCOLO').Value    := pProtocolo;
      qryAux3.Parameters.ParamByName('PSTATUS_CTG').Value   := pSTATUS_CTG;
      qryAux3.Parameters.ParamByName('PCONTINGENCIA').Value := pContingencia;
      qryAux3.ExecSQL;
      result := true;
    end;
  except
    ON E: Exception DO begin
      result := false;
      ShowMessage(E.Message);
    end;
  END;
end;

function TImprimeItens.GetQtdePontos: Double;
begin
  if (sCPFCNPJClientePadrao = fCampos.CNPJCpf) then
    Result := 0
  else
    result := fQtdePtoTotal;
end;

function TImprimeItens.GetTextoTEF: TStringList;
begin
  Result := fTextoComprovante;
end;

function TImprimeItens.GetUtilizandoTef: Boolean;
begin
  Result := fUtilizandoTef;
end;

procedure TImprimeItens.SetMsgComprovante(pValue : String);
begin
  fMsgComprovante := pValue;
end;

procedure TImprimeItens.SetAlteraFonte(pValue : Boolean);
begin
  frmComprovanteNFCe.pAlteraFonte := pValue;
end;

function TimprimeItens.ImprimeComprovante: boolean;
var
  qtde : integer;
  i    : integer;
begin
  try
    QTDE := 1;
    if frmComprovanteNFCe = nil then
      frmComprovanteNFCe := TfrmComprovanteNFCe.Create(nil);

    if (fTextoComprovante <> nil) then begin
      if fTextoComprovante.Count > 0 then begin
        frmComprovanteNFCe.Comprovante := fTextoComprovante;
        if TRim(fMsgComprovante) <> '' then
          frmComprovanteNFCe.MsgComprovante := fMsgComprovante;

        if (frmComprovanteNFCe.MsgComprovante = 'SANGRIA') OR (frmComprovanteNFCe.MsgComprovante = 'SUPRIMENTO') THEN
          qtde := 2;

        for i := 1 to qtde do begin
          frmComprovanteNFCe.ImprimeComprovante;
        end;

        SetMsgComprovante('');
        frmComprovanteNFCe.LimpaTextoComprovanteNFCe;
        LimpaTextoComprovante;
        result := true;
      end;
    end;
  except
    result := false;
  end;
end;

procedure TimprimeItens.LimpaTextoComprovante;
var
  i : integer;
begin
  for i := fTextoComprovante.count -1 downto 0 do
    fTextoComprovante.Delete(i);
end;

function TImprimeItens.ValidaCST(var pCST :string; pCFOP: String): boolean;
begin
  result := false;
  if pCFOP = '5656' then begin
    if (pCST <> '060')  then
      pCST := '060';
    result := true;
    exit;
  end;
  if (pCFOP = '5102') and EM(pCST,['000','020']) then begin
    result := true;
    exit;
  end;
  if (pCFOP = '5405') and (pCST = '060') then begin
    result := true;
    exit;
  end;
  if (pCFOP = '5667') and (pCST = '060') then begin
    result := true;
    exit;
  end;
  if (pCFOP = '6667') and (pCST = '060') then begin
    result := true;
    exit;
  end;
  if (pCFOP = '6102') and (pCST = '000') then begin
    result := true;
    exit;
  end;
  if (pCFOP = '6656') and (pCST = '060') then begin
    result := true;
    exit;
  end;
  if (pCFOP = '6403') and (pCST = '060') then begin
    result := true;
    exit;
  end;
end;

function TImprimeItens.ConverteCSOSN(pValue : integer): string;
begin
  try
    result := '';
    if pValue > 0 then begin
      case pValue of
        1 : result := '101';
        2 : result := '102';
        3 : result := '103';
        4 : result := '201';
        5 : result := '202';
        6 : result := '203';
        7 : result := '300';
        8 : result := '400';
        9 : result := '500';
       10 : result := '900';
      end;
    end;
  except
    result := '';
  end;
end;

function TImprimeItens.ConverteCSOSN_ORIGEM(pValue : integer): string;
begin
  try
    result := '';
    if pValue > 0 then begin
      case pValue of
        1 : result := '0';
        2 : result := '1';
        3 : result := '2';
        4 : result := '3';
        5 : result := '4';
        6 : result := '5';
        7 : result := '6';
        8 : result := '7';
        9 : result := '8';
      end;
    end;
  except
    result := '';
  end;
end;

function TImprimeItens.ConverteCSOSN_MODBCICMS(pValue : integer): string;
begin
  try
    result := '';
    if pValue > 0 then begin
      case pValue of
        1 : result := '0';
        2 : result := '1';
        3 : result := '2';
        4 : result := '3';
      end;
    end;
  except
    result := '';
  end;
end;

function TImprimeItens.ConverteCSOSN_MODBCICMSST(pValue : integer): string;
begin
  try
    result := '';
    if pValue > 0 then begin
      case pValue of
        1 : result := '0';
        2 : result := '1';
        3 : result := '2';
        4 : result := '3';
        5 : result := '4';
        6 : result := '5';
      end;
    end;
  except
    result := '';
  end;
end;

end.
