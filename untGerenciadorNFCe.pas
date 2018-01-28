unit untGerenciadorNFCe;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, ACBrBase, ACBrDFe, ACBrNFe, uCertificado,
  Data.DB, Datasnap.DBClient, uEmitente;

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
    procedure LeCertificado;
    procedure LeEmitente;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmGerenciadorNFCe: TfrmGerenciadorNFCe;

implementation

{$R *.dfm}

procedure TfrmGerenciadorNFCe.FormCreate(Sender: TObject);
begin
  fCertificado := TCertificado.Create;
  LeCertificado;
  fEmitente := TEmitente.Create;
  LeEmitente;

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
