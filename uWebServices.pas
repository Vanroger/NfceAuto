unit uWebServices;

interface

uses
  pcnConversao;

type
  TWebServices = class
  private
    FIntervaloTentativas: integer;
    FUF: sTRING;
    FpAmbiente: TpcnTipoAmbiente;
    FAguardarConsultaRet: integer;
    procedure SetAguardarConsultaRet(const Value: integer);
    procedure SetIntervaloTentativas(const Value: integer);
    procedure SetpAmbiente(const Value: TpcnTipoAmbiente);
    procedure SetUF(const Value: STRING);
  public
  property pAmbiente : TpcnTipoAmbiente read FpAmbiente write SetpAmbiente;
  property UF        : STRING read FUF write SetUF;
  property AguardarConsultaRet : integer read FAguardarConsultaRet write SetAguardarConsultaRet;
  property IntervaloTentativas : integer read FIntervaloTentativas write SetIntervaloTentativas;
  end;

implementation

{ TWebServices }

procedure TWebServices.SetAguardarConsultaRet(const Value: integer);
begin
  FAguardarConsultaRet := Value;
end;

procedure TWebServices.SetIntervaloTentativas(const Value: integer);
begin
  FIntervaloTentativas := Value;
end;

procedure TWebServices.SetpAmbiente(const Value: TpcnTipoAmbiente);
begin
  FpAmbiente := Value;
end;

procedure TWebServices.SetUF(const Value: STRING);
begin
  FUF := Value;
end;

end.
