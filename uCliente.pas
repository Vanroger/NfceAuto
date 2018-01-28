unit uCliente;

interface

TYPE
  TCliente = class
  private
    Fid: integer;
    Fnome: string;
    procedure Setid(const Value: integer);
    procedure Setnome(const Value: string);
  public
    property id : integer read Fid write Setid;
    property nome : string read Fnome write Setnome;
  end;

implementation

{ TCliente }

procedure TCliente.Setid(const Value: integer);
begin
  Fid := Value;
end;

procedure TCliente.Setnome(const Value: string);
begin
  Fnome := Value;
end;

end.
