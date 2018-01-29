unit untGerenciadorNFCe;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, ACBrBase, ACBrDFe, ACBrNFe, uCertificado,
  Data.DB, Datasnap.DBClient, uEmitente, rest.json, System.JSON, uitem,
  system.generics.collections, uDestinatario;

type
  TfrmGerenciadorNFCe = class(TForm)
    ACBrNFe1: TACBrNFe;
    cdsCertificado: TClientDataSet;
    cdsCertificadoCertificado: TStringField;
    cdsCertificadocsc: TStringField;
    cdsCertificadoidcsc: TStringField;
    cdsEmitente: TClientDataSet;
    cdsEmitentecnpj: TStringField;
    cdsEmitentenome: TStringField;
    cdsEmitentefantasia: TStringField;
    cdsEmitenteInscEstadual: TStringField;
    cdsEmitenteInscMun: TStringField;
    cdsEmitenteCNAE: TStringField;
    cdsEmitenteendereco: TStringField;
    cdsEmitentenumero: TStringField;
    cdsEmitenteComplemento: TStringField;
    cdsEmitentebairro: TStringField;
    cdsEmitentecodMun: TIntegerField;
    cdsEmitentemunicipio: TStringField;
    cdsEmitenteuf: TStringField;
    cdsEmitenteCEP: TIntegerField;
    cdsEmitenteFONE: TStringField;
    cdsEmitenteCRT: TStringField;
    procedure FormCreate(Sender: TObject);
  private
    fCertificado : TCertificado;
    fEmitente    : TEmitente;
    fItem        : Titem;
    fItens        : TList;
    fDestinatario : TDestinatario;
    procedure LeCertificado;
    procedure LeEmitente;
    function InformarItens(value: TJSONObject): boolean;
    { Private declarations }
  public
    { Public declarations }
    function EnviarNFCe(value : TJSONObject): TJsonObject;
  end;

var
  frmGerenciadorNFCe: TfrmGerenciadorNFCe;

implementation

{$R *.dfm}

function TfrmGerenciadorNFCe.EnviarNFCe(value: TJSONObject): TJsonObject;
begin
  try
    fDestinatario.limpaCampos;
    InformarItens(value);
  finally

  end;
end;

function TfrmGerenciadorNFCe.InformarItens(value: TJSONObject): boolean;
var
  i: integer;
  valRoot : TJSONValue;
  objRoot : TJSONObject;
  valItens : TJSONValue;
  valDestinatario : TJSONValue;
  arrItens : TJsonArray;
begin
  try
    if fitens.Count > 0 then
      for I := fitens.Count-1 downto 0 do
        fitens.Delete(i);

    valRoot := TJSONObject.ParseJSONValue(value.ToString);
    if valRoot <> nil  then begin
      objRoot := TJSONObject(valRoot);
      if objRoot.Count > 0 then begin
        valDestinatario := objRoot.Values['destinatario'];
        if valDestinatario <> nil then begin
          if valDestinatario is TJSONObject then begin
            fDestinatario := TJson.JsonToObject<TDestinatario>(valDestinatario.tostring);
          end;
        end;

        valItens := objRoot.Values['itens'];
        if valItens <> nil then begin
          if valItens is TJSONArray then begin
            arrItens := TJSONArray(valItens);
            for I := 0 to arrItens.Count -1 do begin
              if arrItens.Items[i] is TJSONObject then begin
                fItem := TJson.JsonToObject<TItem>(arrItens.items[i].tostring);
                fitens.Add(fItem);
              end;
            end;
          end;
        end;
      end;
    end;
  except

  end;

end;

procedure TfrmGerenciadorNFCe.FormCreate(Sender: TObject);
begin
  fitens := TList.Create;
  fCertificado := TCertificado.Create;
  LeCertificado;
  fEmitente := TEmitente.Create;
  LeEmitente;
  fDestinatario := TDestinatario.create;

end;

procedure TFrmGerenciadorNFCe.LeEmitente;
begin
  try
    if FileExists(ExtractFileDir(ParamStr(0))+'\Emitente\Emitente.xml') then begin
      cdsEmitente.Open;
      cdsEmitente.LoadFromFile(ExtractFileDir(ParamStr(0))+'\Emitente\Emitente.xml');
      fEmitente.CNPJCPF := cdsEmitentecnpj.AsString;
      fEmitente.xNome   := cdsEmitentenome.AsString;
      fEmitente.xFant   := cdsEmitentefantasia.AsString;
      fEmitente.IE      := cdsEmitenteInscEstadual.AsString;
      fEmitente.IM      := cdsEmitenteInscMun.AsString;
      fEmitente.CNAE    := cdsEmitenteCNAE.AsString;
      fEmitente.CRT     := cdsEmitenteCRT.AsString;
      fEmitente.xLgr    := cdsEmitenteendereco.AsString;
      fEmitente.nro     := cdsEmitentenumero.AsString;
      fEmitente.xCpl    := cdsEmitenteComplemento.AsString;
      fEmitente.xBairro := cdsEmitentebairro.AsString;
      fEmitente.cMun    := cdsEmitentecodMun.AsInteger;
      fEmitente.xMun    := cdsEmitentemunicipio.AsString;
      fEmitente.UF      := cdsEmitenteuf.AsString;
      fEmitente.CEP     := cdsEmitenteCEP.AsInteger;
      fEmitente.cPais   := 1058;
      fEmitente.xPais   := '1';
      fEmitente.fone    := cdsEmitenteFONE.AsString;
    end;
  except
    on e: exception do begin
      showmessage(e.Message);
    end;
  end;
end;

procedure TFrmGerenciadorNFCe.LeCertificado;
begin
  try
    if FileExists(ExtractFileDir(ParamStr(0))+'\Certificado\certificado.xml') then begin
      cdscertificado.Open;
      cdscertificado.LoadFromFile(ExtractFileDir(ParamStr(0))+'\Certificado\certificado.xml');
      fcertificado.certificado := cdsCertificadoCertificado.AsString;
      fcertificado.CSC := cdsCertificadocsc.AsString;
      fcertificado.IDcsc := cdsCertificadoidcsc.AsString;
    end;
  except
    on e: exception do begin
      showmessage(e.Message);
    end;
  end;
end;

end.
