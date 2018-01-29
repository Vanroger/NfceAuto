unit uCertificado;

interface

type
  TCertificado = class
  private
    FCSC: string;
    FIdCsc: String;
    Fcertificado: string;
    FSenha: String;
    procedure Setcertificado(const Value: string);
    procedure SetCSC(const Value: string);
    procedure SetIdCsc(const Value: String);
    procedure SetSenha(const Value: String);
  public
    property certificado : string read Fcertificado write Setcertificado;
    property CSC : string read FCSC write SetCSC;
    property IdCsc : String read FIdCsc write SetIdCsc;
    property Senha : String read FSenha write SetSenha;
  end;

implementation

{ TCertificado }

procedure TCertificado.Setcertificado(const Value: string);
begin
  Fcertificado := Value;
end;

procedure TCertificado.SetCSC(const Value: string);
begin
  FCSC := Value;
end;

procedure TCertificado.SetIdCsc(const Value: String);
begin
  FIdCsc := Value;
end;

procedure TCertificado.SetSenha(const Value: String);
begin
  FSenha := Value;
end;

end.
