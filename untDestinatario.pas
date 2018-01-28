unit untDestinatario;

interface

type
  TDestinatario = class
  private
    FFone: String;
    FISUF: String;
    FcPais: Integer;
    FEmail: String;
    FUF: String;
    FxNome: String;
    FidEstrangeiro: String;
    FxPais: String;
    FCEP: Integer;
    FcMun: Integer;
    FCNPJCPF: String;
    FIE: String;
    FindIEDest: String;
    FxBairro: String;
    FxCpl: String;
    FxMun: String;
    FxLgr: String;
    Fnro: String;
    procedure SetCEP(const Value: Integer);
    procedure SetcMun(const Value: Integer);
    procedure SetCNPJCPF(const Value: String);
    procedure SetcPais(const Value: Integer);
    procedure SetEmail(const Value: String);
    procedure SetFone(const Value: String);
    procedure SetidEstrangeiro(const Value: String);
    procedure SetIE(const Value: String);
    procedure SetindIEDest(const Value: String);
    procedure SetISUF(const Value: String);
    procedure Setnro(const Value: String);
    procedure SetUF(const Value: String);
    procedure SetxBairro(const Value: String);
    procedure SetxCpl(const Value: String);
    procedure SetxLgr(const Value: String);
    procedure SetxMun(const Value: String);
    procedure SetxNome(const Value: String);
    procedure SetxPais(const Value: String);
  public
    property idEstrangeiro : String read FidEstrangeiro write SetidEstrangeiro;
    property CNPJCPF       : String read FCNPJCPF write SetCNPJCPF;
    property xNome         : String read FxNome write SetxNome;
    property indIEDest     : String read FindIEDest write SetindIEDest;
    property IE            : String read FIE write SetIE;
    property ISUF          : String read FISUF write SetISUF;
    property Email         : String read FEmail write SetEmail;
    property xLgr          : String read FxLgr write SetxLgr;
    property nro           : String read Fnro write Setnro;
    property xCpl          : String read FxCpl write SetxCpl;
    property xBairro       : String read FxBairro write SetxBairro;
    property cMun          : Integer read FcMun write SetcMun;
    property xMun          : String read FxMun write SetxMun;
    property UF            : String read FUF write SetUF;
    property CEP           : Integer read FCEP write SetCEP;
    property cPais         : Integer read FcPais write SetcPais;
    property xPais         : String read FxPais write SetxPais;
    property Fone          : String read FFone write SetFone;
  end;

implementation

{ TDestinatario }

procedure TDestinatario.SetCEP(const Value: Integer);
begin
  FCEP := Value;
end;

procedure TDestinatario.SetcMun(const Value: Integer);
begin
  FcMun := Value;
end;

procedure TDestinatario.SetCNPJCPF(const Value: String);
begin
  FCNPJCPF := Value;
end;

procedure TDestinatario.SetcPais(const Value: Integer);
begin
  FcPais := Value;
end;

procedure TDestinatario.SetEmail(const Value: String);
begin
  FEmail := Value;
end;

procedure TDestinatario.SetFone(const Value: String);
begin
  FFone := Value;
end;

procedure TDestinatario.SetidEstrangeiro(const Value: String);
begin
  FidEstrangeiro := Value;
end;

procedure TDestinatario.SetIE(const Value: String);
begin
  FIE := Value;
end;

procedure TDestinatario.SetindIEDest(const Value: String);
begin
  FindIEDest := Value;
end;

procedure TDestinatario.SetISUF(const Value: String);
begin
  FISUF := Value;
end;

procedure TDestinatario.Setnro(const Value: String);
begin
  Fnro := Value;
end;

procedure TDestinatario.SetUF(const Value: String);
begin
  FUF := Value;
end;

procedure TDestinatario.SetxBairro(const Value: String);
begin
  FxBairro := Value;
end;

procedure TDestinatario.SetxCpl(const Value: String);
begin
  FxCpl := Value;
end;

procedure TDestinatario.SetxLgr(const Value: String);
begin
  FxLgr := Value;
end;

procedure TDestinatario.SetxMun(const Value: String);
begin
  FxMun := Value;
end;

procedure TDestinatario.SetxNome(const Value: String);
begin
  FxNome := Value;
end;

procedure TDestinatario.SetxPais(const Value: String);
begin
  FxPais := Value;
end;

end.
