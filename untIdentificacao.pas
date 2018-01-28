unit untIdentificacao;

interface

type
  TIdentificacao = class
  private
    FindPag: String;
    FxJust: String;
    FdEmi: TDateTime;
    FtpEmis: String;
    FtpNF: String;
    FcMunFG: Integer;
    FnatOp: String;
    FindPres: String;
    FfinNFe: String;
    Fserie: Integer;
    Fmodelo: Integer;
    FcUF: Integer;
    FindFinal: String;
    FnNF: Integer;
    FprocEmi: String;
    FhSaiEnt: TDateTime;
    FcNF: Integer;
    FtpImp: String;
    FdhCont: TDateTime;
    FverProc: String;
    FidDest: String;
    FdSaiEnt: TDateTime;
    procedure SetcMunFG(const Value: Integer);
    procedure SetcNF(const Value: Integer);
    procedure SetcUF(const Value: Integer);
    procedure SetdEmi(const Value: TDateTime);
    procedure SetdhCont(const Value: TDateTime);
    procedure SetdSaiEnt(const Value: TDateTime);
    procedure SetfinNFe(const Value: String);
    procedure SethSaiEnt(const Value: TDateTime);
    procedure SetidDest(const Value: String);
    procedure SetindFinal(const Value: String);
    procedure SetindPag(const Value: String);
    procedure SetindPres(const Value: String);
    procedure Setmodelo(const Value: Integer);
    procedure SetnatOp(const Value: String);
    procedure SetnNF(const Value: Integer);
    procedure SetprocEmi(const Value: String);
    procedure Setserie(const Value: Integer);
    procedure SettpEmis(const Value: String);
    procedure SettpImp(const Value: String);
    procedure SettpNF(const Value: String);
    procedure SetverProc(const Value: String);
    procedure SetxJust(const Value: String);
  public
    property cNF      : Integer read FcNF write SetcNF;
    property natOp    : String  read FnatOp write SetnatOp;
    property indPag   : String  read FindPag write SetindPag;
    property modelo   : Integer read Fmodelo write Setmodelo;
    property serie    : Integer read Fserie write Setserie;
    property nNF      : Integer read FnNF write SetnNF;
    property dEmi     : TDateTime read FdEmi write SetdEmi;
    property dSaiEnt  : TDateTime Read FdSaiEnt write SetdSaiEnt;
    property hSaiEnt  : TDateTime read FhSaiEnt write SethSaiEnt;
    property tpNF     : String read FtpNF write SettpNF;
    property idDest   : String read FidDest write SetidDest;
    property tpImp    : String read FtpImp write SettpImp;
    property tpEmis   : String read FtpEmis write SettpEmis;
    property finNFe   : String read FfinNFe write SetfinNFe;
    property indFinal : String read FindFinal write SetindFinal;
    property indPres  : String read FindPres write SetindPres;
    property procEmi  : String read FprocEmi write SetprocEmi;
    property verProc  : String read FverProc write SetverProc;
    property dhCont   : TDateTime read FdhCont write SetdhCont;
    property xJust    : String read FxJust write SetxJust;
    property cUF      : Integer read FcUF write SetcUF;
    property cMunFG   : Integer read FcMunFG write SetcMunFG;
  end;



implementation

{ TIdentificacao }

procedure TIdentificacao.SetcMunFG(const Value: Integer);
begin
  FcMunFG := Value;
end;

procedure TIdentificacao.SetcNF(const Value: Integer);
begin
  FcNF := Value;
end;

procedure TIdentificacao.SetcUF(const Value: Integer);
begin
  FcUF := Value;
end;

procedure TIdentificacao.SetdEmi(const Value: TDateTime);
begin
  FdEmi := Value;
end;

procedure TIdentificacao.SetdhCont(const Value: TDateTime);
begin
  FdhCont := Value;
end;

procedure TIdentificacao.SetdSaiEnt(const Value: TDateTime);
begin
  FdSaiEnt := Value;
end;

procedure TIdentificacao.SetfinNFe(const Value: String);
begin
  FfinNFe := Value;
end;

procedure TIdentificacao.SethSaiEnt(const Value: TDateTime);
begin
  FhSaiEnt := Value;
end;

procedure TIdentificacao.SetidDest(const Value: String);
begin
  FidDest := Value;
end;

procedure TIdentificacao.SetindFinal(const Value: String);
begin
  FindFinal := Value;
end;

procedure TIdentificacao.SetindPag(const Value: String);
begin
  FindPag := Value;
end;

procedure TIdentificacao.SetindPres(const Value: String);
begin
  FindPres := Value;
end;

procedure TIdentificacao.Setmodelo(const Value: Integer);
begin
  Fmodelo := Value;
end;

procedure TIdentificacao.SetnatOp(const Value: String);
begin
  FnatOp := Value;
end;

procedure TIdentificacao.SetnNF(const Value: Integer);
begin
  FnNF := Value;
end;

procedure TIdentificacao.SetprocEmi(const Value: String);
begin
  FprocEmi := Value;
end;

procedure TIdentificacao.Setserie(const Value: Integer);
begin
  Fserie := Value;
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

procedure TIdentificacao.SetxJust(const Value: String);
begin
  FxJust := Value;
end;

end.
