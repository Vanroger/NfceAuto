unit uIdentificacao;

interface

uses
  pcnConversaoNfe;

type
  TIdentificacao = class
  private
    FindPag: string;
    FxJust: String;
    FtpEmis: String;
    FtpNF: String;
    FdhEmi: tdatetime;
    FcMunFG: integer;
    FnatOp: string;
    FindPres: string;
    FfinNFe: String;
    FtpAmb: String;
    Fserie: integer;
    Fmodelo: integer;
    FindFinal: string;
    FnNF: integer;
    FcUF: integer;
    FprocEmi: string;
    FcNF: integer;
    FtpImp: String;
    FdhCont: tdatetime;
    FverProc: String;
    FidDest: String;
    Fversao: TpcnVersaoDF;
    FmodFrete: integer;
    procedure SetcMunFG(const Value: integer);
    procedure SetcNF(const Value: integer);
    procedure SetcUF(const Value: integer);
    procedure SetdhCont(const Value: tdatetime);
    procedure SetdhEmi(const Value: tdatetime);
    procedure SetfinNFe(const Value: String);
    procedure SetidDest(const Value: String);
    procedure SetindFinal(const Value: string);
    procedure SetindPag(const Value: string);
    procedure SetindPres(const Value: string);
    procedure Setmodelo(const Value: integer);
    procedure SetnatOp(const Value: string);
    procedure SetnNF(const Value: integer);
    procedure SetprocEmi(const Value: string);
    procedure Setserie(const Value: integer);
    procedure SettpAmb(const Value: String);
    procedure SettpEmis(const Value: String);
    procedure SettpImp(const Value: String);
    procedure SettpNF(const Value: String);
    procedure SetverProc(const Value: String);
    procedure SetxJust(const Value: String);
    procedure Setversao(const Value: TpcnVersaoDF);
    procedure SetmodFrete(const Value: integer);
  public
  published
  property cUF      : integer    read FcUF      write SetcUF;      //C�digo da UF do emitente do Documento Fisca
  property cMunFG   : integer   read FcMunFG   write SetcMunFG;   //C�digo do Munic�pio de Ocorr�ncia do Fato Gerador
  property cNF      : integer   read FcNF      write SetcNF;      //C�digo Num�rico que comp�e a Chave de Acesso
  property natOp    : string    read FnatOp    write SetnatOp;    //NATUREZA DA OPERA��O
  property indPag   : string   read FindPag   write SetindPag;   //Indicador da forma de pagamento 0=Pagamento � vista; 1=Pagamento a prazo; 2=Outros.
  property modelo   : integer   read Fmodelo   write Setmodelo;   //65
  property serie    : integer   read Fserie    write Setserie;    //1
  property nNF      : integer   read FnNF      write SetnNF;      //N�mero do Documento Fiscal
  property dhEmi    : tdatetime read FdhEmi    write SetdhEmi;    //Data e hora de emiss�o do Documento Fiscal
  property tpNF     : String   read FtpNF     write SettpNF;     //Tipo de opera��o 0 - entrada 1 - saida
  property idDest   : String   read FidDest   write SetidDest;   //Identificador de local de destino da opera��o 1=Opera��o interna; 2=Opera��o interestadual; 3=Opera��o com exterior.
  property tpImp    : String   read FtpImp    write SettpImp;    //Formato de Impress�o do DANFE 0=Sem gera��o de DANFE; 1=DANFE normal, Retrato; 2=DANFE normal, Paisagem; 3=DANFE Simplificado; 4=DANFE NFC-e; 5=DANFE NFC-e em mensagem eletr�nica (o envio de mensagem eletr�nica pode ser feita de forma simult�nea com a impress�o do DANFE; usar o tpImp=5 quando esta for a �nica forma de disponibiliza��o do DANFE).
  property tpEmis   : String   read FtpEmis   write SettpEmis;   //Tipo de Emiss�o da NF-e 1=Emiss�o normal (n�o em conting�ncia); 2=Conting�ncia FS-IA, com impress�o do DANFE em formul�rio de seguran�a; 3=Conting�ncia SCAN (Sistema de Conting�ncia do Ambiente Nacional); 4=Conting�ncia DPEC (Declara��o Pr�via da Emiss�o em Conting�ncia); 5=Conting�ncia FS-DA, com impress�o do DANFE em formul�rio de seguran�a; 6=Conting�ncia SVC-AN (SEFAZ Virtual de Conting�ncia do AN); 7=Conting�ncia SVC-RS (SEFAZ Virtual de Conting�ncia do RS);
  property tpAmb    : String    read FtpAmb    write SettpAmb;    //Identifica��o do Ambiente
  property finNFe   : String   read FfinNFe   write SetfinNFe;   //Finalidade de emiss�o da NF-e 1=NF-e normal; 2=NF-e complementar; 3=NF-e de ajuste; 4=Devolu��o de mercadoria
  property indFinal : string   read FindFinal write SetindFinal; //Indica opera��o com Consumidor final 0=Normal; 1=Consumidor final; 29
  property indPres  : string   read FindPres  write SetindPres;  //Indicador de presen�a do comprador no estabelecimento comercial no momento da opera��o 0=N�o se aplica (por exemplo, Nota Fiscal complementar ou de ajuste); 1=Opera��o presencial; 2=Opera��o n�o presencial, pela Internet; 3=Opera��o n�o presencial, Teleatendimento; 4=NFC-e em opera��o com entrega a domic�lio; 9=Opera��o n�o presencial, outros.
  property procEmi  : string    read FprocEmi  write SetprocEmi;  //Processo de emiss�o da NF-e 0=Emiss�o de NF-e com aplicativo do contribuinte; 1=Emiss�o de NF-e avulsa pelo Fisco; 2=Emiss�o de NF-e avulsa, pelo contribuinte com seu certificado digital, atrav�s do site do Fisco; 3=Emiss�o NF-e pelo contribuinte com aplicativo fornecido pelo Fisco.
  property verProc  : String    read FverProc  write SetverProc;  //Vers�o do Processo de emiss�o da NF-e Informar a vers�o do aplicativo emissor de NF-e.
  property dhCont   : tdatetime read FdhCont   write SetdhCont;   //Data e Hora da entrada em conting�ncia
  property xJust    : String    read FxJust    write SetxJust;    //Justificativa da entrada em conting�ncia "Falha na conexao com a internet"
  property versao   : TpcnVersaoDF read Fversao write Setversao;
  property modFrete : integer read FmodFrete write SetmodFrete;  //para nfce o valor padrao � 9

  end;

implementation

{ TIdentificacao }

procedure TIdentificacao.SetcMunFG(const Value: integer);
begin
  FcMunFG := Value;
end;

procedure TIdentificacao.SetcNF(const Value: integer);
begin
  FcNF := Value;
end;

procedure TIdentificacao.SetcUF(const Value: integer);
begin
  FcUF := Value;
end;

procedure TIdentificacao.SetdhCont(const Value: tdatetime);
begin
  FdhCont := Value;
end;

procedure TIdentificacao.SetdhEmi(const Value: tdatetime);
begin
  FdhEmi := Value;
end;

procedure TIdentificacao.SetfinNFe(const Value: String);
begin
  FfinNFe := Value;
end;

procedure TIdentificacao.SetidDest(const Value: String);
begin
  FidDest := Value;
end;

procedure TIdentificacao.SetindFinal(const Value: string);
begin
  FindFinal := Value;
end;

procedure TIdentificacao.SetindPag(const Value: string);
begin
  FindPag := Value;
end;

procedure TIdentificacao.SetindPres(const Value: string);
begin
  FindPres := Value;
end;

procedure TIdentificacao.Setmodelo(const Value: integer);
begin
  Fmodelo := Value;
end;

procedure TIdentificacao.SetmodFrete(const Value: integer);
begin
  FmodFrete := Value;
end;

procedure TIdentificacao.SetnatOp(const Value: string);
begin
  FnatOp := Value;
end;

procedure TIdentificacao.SetnNF(const Value: integer);
begin
  FnNF := Value;
end;

procedure TIdentificacao.SetprocEmi(const Value: string);
begin
  FprocEmi := Value;
end;

procedure TIdentificacao.Setserie(const Value: integer);
begin
  Fserie := Value;
end;

procedure TIdentificacao.SettpAmb(const Value: String);
begin
  FtpAmb := Value;
end;

procedure TIdentificacao.SettpEmis(const Value: String);
begin
  FtpEmis := Value;
end;

procedure TIdentificacao.SettpImp(const Value: String);
begin
  FtpImp := Value;
end;

procedure TIdentificacao.SettpNF(const Value: String);
begin
  FtpNF := Value;
end;

procedure TIdentificacao.SetverProc(const Value: String);
begin
  FverProc := Value;
end;

procedure TIdentificacao.Setversao(const Value: TpcnVersaoDF);
begin
  Fversao := Value;
end;

procedure TIdentificacao.SetxJust(const Value: String);
begin
  FxJust := Value;
end;

end.
