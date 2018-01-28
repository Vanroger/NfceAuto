unit uEmitente;

interface

type
  TEmitente = class
  private
    Ffone: String;
    FCNAE: String;
    FcPais: Integer;
    FIEST: String;
    FxFant: String;
    FIM: String;
    FUF: String;
    FxNome: String;
    FxPais: String;
    FCEP: Integer;
    FcMun: Integer;
    FCNPJCPF: String;
    FIE: String;
    FxBairro: String;
    FxCpl: String;
    FxMun: String;
    FxLgr: String;
    FCRT: String;
    Fnro: String;
    procedure SetCEP(const Value: Integer);
    procedure SetcMun(const Value: Integer);
    procedure SetCNAE(const Value: String);
    procedure SetCNPJCPF(const Value: String);
    procedure SetcPais(const Value: Integer);
    procedure SetCRT(const Value: String);
    procedure Setfone(const Value: String);
    procedure SetIE(const Value: String);
    procedure SetIEST(const Value: String);
    procedure SetIM(const Value: String);
    procedure Setnro(const Value: String);
    procedure SetUF(const Value: String);
    procedure SetxBairro(const Value: String);
    procedure SetxCpl(const Value: String);
    procedure SetxFant(const Value: String);
    procedure SetxLgr(const Value: String);
    procedure SetxMun(const Value: String);
    procedure SetxNome(const Value: String);
    procedure SetxPais(const Value: String);
  public
    property CNPJCPF :String read FCNPJCPF write SetCNPJCPF;
    property xNome   :String read FxNome write SetxNome;
    property xFant   :String read FxFant write SetxFant;
    property IE      :String read FIE write SetIE;
    property IEST    :String read FIEST write SetIEST;
    property IM      :String read FIM write SetIM;
    property CNAE    :String read FCNAE write SetCNAE;
    property CRT     :String read FCRT write SetCRT;
    property xLgr    :String read FxLgr write SetxLgr;
    property nro     :String read Fnro write Setnro;
    property xCpl    :String read FxCpl write SetxCpl;
    property xBairro :String read FxBairro write SetxBairro;
    property cMun    :Integer read FcMun write SetcMun;
    property xMun    :String read FxMun write SetxMun;
    property UF      :String read FUF write SetUF;
    property CEP     :Integer read FCEP write SetCEP;
    property cPais   :Integer read FcPais write SetcPais;
    property xPais   :String read FxPais write SetxPais;
    property fone    :String read Ffone write Setfone;
  end;

implementation

{ TEmitente }

procedure TEmitente.SetCEP(const Value: Integer);
begin
  FCEP := Value;
end;

procedure TEmitente.SetcMun(const Value: Integer);
begin
  FcMun := Value;
end;

procedure TEmitente.SetCNAE(const Value: String);
begin
  FCNAE := Value;
end;

procedure TEmitente.SetCNPJCPF(const Value: String);
begin
  FCNPJCPF := Value;
end;

procedure TEmitente.SetcPais(const Value: Integer);
begin
  FcPais := Value;
end;

procedure TEmitente.SetCRT(const Value: String);
begin
  FCRT := Value;
end;

procedure TEmitente.Setfone(const Value: String);
begin
  Ffone := Value;
end;

procedure TEmitente.SetIE(const Value: String);
begin
  FIE := Value;
end;

procedure TEmitente.SetIEST(const Value: String);
begin
  FIEST := Value;
end;

procedure TEmitente.SetIM(const Value: String);
begin
  FIM := Value;
end;

procedure TEmitente.Setnro(const Value: String);
begin
  Fnro := Value;
end;

procedure TEmitente.SetUF(const Value: String);
begin
  FUF := Value;
end;

procedure TEmitente.SetxBairro(const Value: String);
begin
  FxBairro := Value;
end;

procedure TEmitente.SetxCpl(const Value: String);
begin
  FxCpl := Value;
end;

procedure TEmitente.SetxFant(const Value: String);
begin
  FxFant := Value;
end;

procedure TEmitente.SetxLgr(const Value: String);
begin
  FxLgr := Value;
end;

procedure TEmitente.SetxMun(const Value: String);
begin
  FxMun := Value;
end;

procedure TEmitente.SetxNome(const Value: String);
begin
  FxNome := Value;
end;

procedure TEmitente.SetxPais(const Value: String);
begin
  FxPais := Value;
end;

end.
