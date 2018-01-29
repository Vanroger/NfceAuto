unit uArquivoJson;

interface

uses uDestinatario, uItem;


type

  TArrayItens = array of TItem;

  TArquivo = class
  private
    FDestinatario: TDestinatario;
    FItens: TArrayItens;
    procedure SetDestinatario(const Value: TDestinatario);
    procedure SetItens(const Value: TArrayItens);
  public
    property Destinatario : TDestinatario read FDestinatario write SetDestinatario;
    property Itens : TArrayItens read FItens write SetItens;
  end;

implementation

{ TArquivo }

procedure TArquivo.SetDestinatario(const Value: TDestinatario);
begin
  FDestinatario := Value;
end;

procedure TArquivo.SetItens(const Value: TArrayItens);
begin
  FItens := Value;
end;

end.
