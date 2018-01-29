unit uItem;

interface

  type
    TItem = class
  private
    FfQuantidade: double;
    FfNome: String;
    FfCST: STRING;
    FfOrigem: String;
    FfCSOSN: string;
    FfUnitario: double;
    FfCodigo: String;
    FfTotal: double;
    FfNCM: STRING;
    FfAliquota: double;
    procedure SetfAliquota(const Value: double);
    procedure SetfCodigo(const Value: String);
    procedure SetfCSOSN(const Value: string);
    procedure SetfCST(const Value: STRING);
    procedure SetfNCM(const Value: STRING);
    procedure SetfNome(const Value: String);
    procedure SetfOrigem(const Value: String);
    procedure SetfQuantidade(const Value: double);
    procedure SetfTotal(const Value: double);
    procedure SetfUnitario(const Value: double);
    public
     property Codigo     : String read FfCodigo write SetfCodigo;
     property Nome       : String read FfNome write SetfNome;
     property Quantidade : double read FfQuantidade write SetfQuantidade;
     property Unitario   : double read FfUnitario write SetfUnitario;
     property Total      : double read FfTotal write SetfTotal;
     property NCM        : STRING read FfNCM write SetfNCM;
     property Origem     : String read FfOrigem write SetfOrigem;
     property CST        : STRING read FfCST write SetfCST;
     property CSOSN      : string read FfCSOSN write SetfCSOSN;
     property Aliquota   : double read FfAliquota write SetfAliquota;
    end;

implementation

{ TItem }

procedure TItem.SetfAliquota(const Value: double);
begin
  FfAliquota := Value;
end;

procedure TItem.SetfCodigo(const Value: String);
begin
  FfCodigo := Value;
end;

procedure TItem.SetfCSOSN(const Value: string);
begin
  FfCSOSN := Value;
end;

procedure TItem.SetfCST(const Value: STRING);
begin
  FfCST := Value;
end;

procedure TItem.SetfNCM(const Value: STRING);
begin
  FfNCM := Value;
end;

procedure TItem.SetfNome(const Value: String);
begin
  FfNome := Value;
end;

procedure TItem.SetfOrigem(const Value: String);
begin
  FfOrigem := Value;
end;

procedure TItem.SetfQuantidade(const Value: double);
begin
  FfQuantidade := Value;
end;

procedure TItem.SetfTotal(const Value: double);
begin
  FfTotal := Value;
end;

procedure TItem.SetfUnitario(const Value: double);
begin
  FfUnitario := Value;
end;

end.
