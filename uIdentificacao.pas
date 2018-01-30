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
  property cUF      : integer    read FcUF      write SetcUF;      //Código da UF do emitente do Documento Fisca
  property cMunFG   : integer   read FcMunFG   write SetcMunFG;   //Código do Município de Ocorrência do Fato Gerador
  property cNF      : integer   read FcNF      write SetcNF;      //Código Numérico que compõe a Chave de Acesso
  property natOp    : string    read FnatOp    write SetnatOp;    //NATUREZA DA OPERAÇÃO
  property indPag   : string   read FindPag   write SetindPag;   //Indicador da forma de pagamento 0=Pagamento à vista; 1=Pagamento a prazo; 2=Outros.
  property modelo   : integer   read Fmodelo   write Setmodelo;   //65
  property serie    : integer   read Fserie    write Setserie;    //1
  property nNF      : integer   read FnNF      write SetnNF;      //Número do Documento Fiscal
  property dhEmi    : tdatetime read FdhEmi    write SetdhEmi;    //Data e hora de emissão do Documento Fiscal
  property tpNF     : String   read FtpNF     write SettpNF;     //Tipo de operação 0 - entrada 1 - saida
  property idDest   : String   read FidDest   write SetidDest;   //Identificador de local de destino da operação 1=Operação interna; 2=Operação interestadual; 3=Operação com exterior.
  property tpImp    : String   read FtpImp    write SettpImp;    //Formato de Impressão do DANFE 0=Sem geração de DANFE; 1=DANFE normal, Retrato; 2=DANFE normal, Paisagem; 3=DANFE Simplificado; 4=DANFE NFC-e; 5=DANFE NFC-e em mensagem eletrônica (o envio de mensagem eletrônica pode ser feita de forma simultânea com a impressão do DANFE; usar o tpImp=5 quando esta for a única forma de disponibilização do DANFE).
  property tpEmis   : String   read FtpEmis   write SettpEmis;   //Tipo de Emissão da NF-e 1=Emissão normal (não em contingência); 2=Contingência FS-IA, com impressão do DANFE em formulário de segurança; 3=Contingência SCAN (Sistema de Contingência do Ambiente Nacional); 4=Contingência DPEC (Declaração Prévia da Emissão em Contingência); 5=Contingência FS-DA, com impressão do DANFE em formulário de segurança; 6=Contingência SVC-AN (SEFAZ Virtual de Contingência do AN); 7=Contingência SVC-RS (SEFAZ Virtual de Contingência do RS);
  property tpAmb    : String    read FtpAmb    write SettpAmb;    //Identificação do Ambiente
  property finNFe   : String   read FfinNFe   write SetfinNFe;   //Finalidade de emissão da NF-e 1=NF-e normal; 2=NF-e complementar; 3=NF-e de ajuste; 4=Devolução de mercadoria
  property indFinal : string   read FindFinal write SetindFinal; //Indica operação com Consumidor final 0=Normal; 1=Consumidor final; 29
  property indPres  : string   read FindPres  write SetindPres;  //Indicador de presença do comprador no estabelecimento comercial no momento da operação 0=Não se aplica (por exemplo, Nota Fiscal complementar ou de ajuste); 1=Operação presencial; 2=Operação não presencial, pela Internet; 3=Operação não presencial, Teleatendimento; 4=NFC-e em operação com entrega a domicílio; 9=Operação não presencial, outros.
  property procEmi  : string    read FprocEmi  write SetprocEmi;  //Processo de emissão da NF-e 0=Emissão de NF-e com aplicativo do contribuinte; 1=Emissão de NF-e avulsa pelo Fisco; 2=Emissão de NF-e avulsa, pelo contribuinte com seu certificado digital, através do site do Fisco; 3=Emissão NF-e pelo contribuinte com aplicativo fornecido pelo Fisco.
  property verProc  : String    read FverProc  write SetverProc;  //Versão do Processo de emissão da NF-e Informar a versão do aplicativo emissor de NF-e.
  property dhCont   : tdatetime read FdhCont   write SetdhCont;   //Data e Hora da entrada em contingência
  property xJust    : String    read FxJust    write SetxJust;    //Justificativa da entrada em contingência "Falha na conexao com a internet"
  property versao   : TpcnVersaoDF read Fversao write Setversao;
  property modFrete : integer read FmodFrete write SetmodFrete;  //para nfce o valor padrao é 9

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
