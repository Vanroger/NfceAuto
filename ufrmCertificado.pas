unit ufrmCertificado;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Buttons, Vcl.StdCtrls, Data.DB,
  Datasnap.DBClient, ACBrCAPICOM_TLB, MidasLib;

type
  TfrmCertificado = class(TForm)
    edtCertificado: TEdit;
    Certificado: TLabel;
    edtCSC: TEdit;
    Label1: TLabel;
    edtID: TEdit;
    Label2: TLabel;
    SpeedButton1: TSpeedButton;
    btnGravar: TSpeedButton;
    SpeedButton3: TSpeedButton;
    cdsCertificado: TClientDataSet;
    cdsCertificadoCertificado: TStringField;
    cdsCertificadocsc: TStringField;
    cdsCertificadoidcsc: TStringField;
    procedure btnGravarClick(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCertificado: TfrmCertificado;

implementation

{$R *.dfm}

procedure TfrmCertificado.btnGravarClick(Sender: TObject);
begin
  try
    cdscertificado.EmptyDataSet;
    cdscertificado.Open;
    cdscertificado.Insert;
    if trim(edtCertificado.Text) <> '' then
      cdsCertificadoCertificado.AsString := trim(edtCertificado.Text);
    if trim(edtCSC.Text) <> '' then
      cdsCertificadocsc.AsString := TRIM(edtCSC.Text);
    if TRIM(edtID.Text) <> '' then
      cdsCertificadoidcsc.Text := trim(edtid.Text);


    if not DirectoryExists(ExtractFileDir(ParamStr(0))+'\Certificado\') then begin
      ForceDirectories(ExtractFileDir(ParamStr(0))+'\Certificado\');
    end;

    cdsCertificado.LogChanges := false;

    cdsCertificado.SaveToFile( ExtractFileDir(ParamStr(0))+'\Certificado\certificado.xml',dfXMLUTF8);

    showmessage('Gravação executada com sucesso!');
    self.Close;
  except
    on e: exception do begin
      showmessage('Erro ao gravar dados!');
    end;
  end;
end;

procedure TfrmCertificado.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action := caFree;
end;

procedure TfrmCertificado.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  canclose := true;
end;

procedure TfrmCertificado.FormShow(Sender: TObject);
begin
  try
    if FileExists(ExtractFileDir(ParamStr(0))+'\Certificado\certificado.xml') then begin
 //     cdscertificado.CreateDataSet;
      cdscertificado.Open;
      cdscertificado.LoadFromFile(ExtractFileDir(ParamStr(0))+'\Certificado\certificado.xml');
      edtCertificado.Text := cdsCertificadoCertificado.AsString;
      edtCSC.Text := cdsCertificadocsc.AsString;
      edtID.Text  := cdsCertificadoidcsc.AsString;
    end;
  except
    on e: exception do begin
      showmessage(e.Message);
    end;
  end;
end;

procedure TfrmCertificado.SpeedButton1Click(Sender: TObject);
var
  Store : IStore3;
  CertsLista, CertsSelecionado : ICertificates2;
  CertDados : ICertificate;
  lSigner     : TSigner;
  lSignedData : TSignedData;
  vEmpresaCnpj : WideString;
  Cert         : ICertificate2;

  CertContext : ICertContext;
  PCertContext : Pointer;
begin
  inherited;
//  vEmpresaCnpj := sEmpresaCNPJ;
  Store := CoStore.Create;
  Store.Open(CAPICOM_CURRENT_USER_STORE, 'My', CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED);

  CertsLista := Store.Certificates as ICertificates2;
  CertsSelecionado := CertsLista.Select('Certificado(s) Digital(is) disponível(is)', 'Selecione o Certificado Digital para uso no aplicativo', false);

  if not(CertsSelecionado.Count = 0) then begin
    CertDados := IInterface(CertsSelecionado.Item[1]) as ICertificate2;

    if CertDados.ValidFromDate > Now then begin
      showmessage('certificado não liberado. aguardar' +datetostr(CertDados.ValidFromDate));
      exit;
    end;

    if CertDados.ValidToDate < Now then
    begin
      showmessage('certificado expirado');
      exit;
    end;

//    if Pos(vEmpresaCNPJ,CertDados.SubjectName) = 0 then begin
//      showmessage('certificado pertencente a outra empresa / pessoa '+chr(13)+CertDados.SubjectName);
//      exit;
//    end;

    if not(CertsSelecionado.Count = 0) then begin
      Cert := IInterface(CertsSelecionado.Item[1]) as ICertificate2;
      edtcertificado.text := Cert.SerialNumber;
//      FDataVenc    := Cert.ValidToDate;
//      FSubjectName := Cert.SubjectName;
    end;

  end;

end;

procedure TfrmCertificado.SpeedButton3Click(Sender: TObject);
begin
  self.close;
end;

end.
