unit uDestinatario;

interface

type
  TDestinatario = class
  private
    FCNPJ: STRING;
    FBAIRRO: STRING;
    FUF: STRING;
    FXNOME: STRING;
    FCEP: STRING;
    FCMUN: STRING;
    FXMUN: STRING;
    FXLGR: STRING;
    FNRO: STRING;
    FxCpl: String;
    procedure SetBAIRRO(const Value: STRING);
    procedure SetCEP(const Value: STRING);
    procedure SetCMUN(const Value: STRING);
    procedure SetCNPJ(const Value: STRING);
    procedure SetNRO(const Value: STRING);
    procedure SetUF(const Value: STRING);
    procedure SetXLGR(const Value: STRING);
    procedure SetXMUN(const Value: STRING);
    procedure SetXNOME(const Value: STRING);
    procedure SetxCpl(const Value: String);
    public
    procedure LimpaCampos;
    property CNPJ   : STRING read FCNPJ   write SetCNPJ;
    property XNOME  : STRING read FXNOME  write SetXNOME;
    property XLGR   : STRING read FXLGR   write SetXLGR;
    property NRO    : STRING read FNRO    write SetNRO;
    property BAIRRO : STRING read FBAIRRO write SetBAIRRO;
    property CMUN   : STRING read FCMUN   write SetCMUN;
    property XMUN   : STRING read FXMUN   write SetXMUN;
    property UF     : STRING read FUF     write SetUF;
    property CEP    : STRING read FCEP    write SetCEP;
    property xCpl   : String read FxCpl   write SetxCpl; //complemento
  end;

implementation

{ TDestinatario }

procedure TDestinatario.LimpaCampos;
begin
  FCNPJ := '';
  FBAIRRO := '';
  FUF := '';
  FXNOME := '';
  FCEP := '';
  FCMUN := '';
  FXMUN := '';
  FXLGR := '';
  FNRO := '';
end;

procedure TDestinatario.SetBAIRRO(const Value: STRING);
begin
  FBAIRRO := Value;
end;

procedure TDestinatario.SetCEP(const Value: STRING);
begin
  FCEP := Value;
end;

procedure TDestinatario.SetCMUN(const Value: STRING);
begin
  FCMUN := Value;
end;

procedure TDestinatario.SetCNPJ(const Value: STRING);
begin
  FCNPJ := Value;
end;

procedure TDestinatario.SetNRO(const Value: STRING);
begin
  FNRO := Value;
end;

procedure TDestinatario.SetUF(const Value: STRING);
begin
  FUF := Value;
end;

procedure TDestinatario.SetxCpl(const Value: String);
begin
  FxCpl := Value;
end;

procedure TDestinatario.SetXLGR(const Value: STRING);
begin
  FXLGR := Value;
end;

procedure TDestinatario.SetXMUN(const Value: STRING);
begin
  FXMUN := Value;
end;

procedure TDestinatario.SetXNOME(const Value: STRING);
begin
  FXNOME := Value;
end;

end.
