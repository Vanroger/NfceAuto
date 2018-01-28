unit uCertificado;

interface

type
  TCertificado = class
  private
    FCSC: string;
    FIdCsc: String;
    Fcertificado: string;
    procedure Setcertificado(const Value: string);
    procedure SetCSC(const Value: string);
    procedure SetIdCsc(const Value: String);
  public
    property certificado : string read Fcertificado write Setcertificado;
    property CSC : string read FCSC write SetCSC;
    property IdCsc : String read FIdCsc write SetIdCsc;
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

end.
