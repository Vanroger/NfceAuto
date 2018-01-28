unit ufrmEmitente;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Data.DB, Datasnap.DBClient,
  Vcl.Buttons;

type
  TfrmEmitente = class(TForm)
    Label1: TLabel;
    edtCNPJ: TEdit;
    Label2: TLabel;
    edtNome: TEdit;
    Label3: TLabel;
    edtFantasia: TEdit;
    Label4: TLabel;
    edtIE: TEdit;
    Label6: TLabel;
    edtInscMun: TEdit;
    Label7: TLabel;
    edtCNAE: TEdit;
    Label8: TLabel;
    Label9: TLabel;
    edtLgr: TEdit;
    Label10: TLabel;
    edtNumero: TEdit;
    Label11: TLabel;
    edtComplemento: TEdit;
    Label12: TLabel;
    edtBairro: TEdit;
    Label13: TLabel;
    edtCodMun: TEdit;
    Label14: TLabel;
    edtCidade: TEdit;
    Label15: TLabel;
    edtUF: TEdit;
    Label16: TLabel;
    edtCep: TEdit;
    Label17: TLabel;
    edtFone: TEdit;
    cbxCRT: TComboBox;
    cdsEmitente: TClientDataSet;
    cdsEmitentecnpj: TStringField;
    cdsEmitentenome: TStringField;
    cdsEmitentefantasia: TStringField;
    cdsEmitenteInscEstadual: TStringField;
    cdsEmitenteendereco: TStringField;
    cdsEmitentenumero: TStringField;
    cdsEmitenteInscMun: TStringField;
    cdsEmitenteCNAE: TStringField;
    cdsEmitenteComplemento: TStringField;
    cdsEmitentebairro: TStringField;
    cdsEmitentecodMun: TIntegerField;
    cdsEmitentemunicipio: TStringField;
    cdsEmitenteuf: TStringField;
    cdsEmitenteCEP: TIntegerField;
    cdsEmitenteFONE: TStringField;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    cdsEmitenteCRT: TStringField;
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEmitente: TfrmEmitente;

implementation

{$R *.dfm}

procedure TfrmEmitente.FormShow(Sender: TObject);
begin
                              try
    if FileExists(ExtractFileDir(ParamStr(0))+'\Emitente\Emitente.xml') then begin
      cdsemitente.Open;
      cdsemitente.LoadFromFile(ExtractFileDir(ParamStr(0))+'\Emitente\Emitente.xml');
      edtCNPJ.Text := cdsEmitentecnpj.AsString;
      edtNome.Text := cdsEmitentenome.AsString;
      edtFantasia.Text := cdsEmitentefantasia.AsString;
      edtIE.Text := cdsEmitenteInscEstadual.AsString;
      edtLgr.Text := cdsEmitenteendereco.AsString;
      edtNumero.Text := cdsEmitentenumero.AsString;
      edtInscMun.Text := cdsEmitenteInscMun.AsString;
      edtCNAE.Text := cdsEmitenteCNAE.AsString;
      case STRTOINT(cdsEmitenteCRT.AsString) of
        1 : cbxcrt.ItemIndex := 0;
        2 : cbxcrt.ItemIndex := 1;
        3 : cbxcrt.ItemIndex := 2;
      end;
      edtComplemento.Text := cdsEmitenteComplemento.AsString;
      edtBairro.Text := cdsEmitentebairro.AsString;
      edtCodMun.Text := IntToStr(cdsEmitentecodMun.AsInteger);
      edtCidade.Text := cdsEmitentemunicipio.AsString;
      edtuf.Text     := cdsEmitenteuf.AsString;
      edtCep.Text    := IntToStr(cdsEmitenteCEP.AsInteger);
      edtFone.Text   := cdsEmitenteFONE.AsString;
    end;
  except
    on e: exception do begin
      showmessage(e.Message);
    end;
  end;
end;

procedure TfrmEmitente.SpeedButton1Click(Sender: TObject);
begin
  try

    cdsemitente.EmptyDataSet;

    cdsemitente.Open;
    cdsemitente.Insert;

    cdsEmitentecnpj.AsString         := edtCNPJ.Text;
    cdsEmitentenome.AsString         := edtNome.Text;
    cdsEmitentefantasia.AsString     := edtFantasia.Text;
    cdsEmitenteInscEstadual.AsString := edtIE.Text;
    cdsEmitenteendereco.AsString     :=  edtLgr.Text;
    cdsEmitentenumero.AsString       := edtNumero.Text;
    cdsEmitenteInscMun.AsString      := edtInscMun.Text;
    cdsEmitenteCNAE.AsString         := edtCNAE.Text;
    case cbxcrt.ItemIndex of
      0 : cdsEmitenteCRT.AsString := '1';
      1 : cdsEmitenteCRT.AsString := '2';
      2 : cdsEmitenteCRT.AsString := '3';
    end;
    cdsEmitenteComplemento.AsString := edtComplemento.Text;
    cdsEmitentebairro.AsString      := edtBairro.Text;
    cdsEmitentecodMun.AsInteger     := strtoint(edtCodMun.Text);
    cdsEmitentemunicipio.AsString   := edtCidade.Text;
    cdsEmitenteuf.AsString          := edtuf.Text;
    cdsEmitenteCEP.AsInteger        := strtoint(edtCep.Text);
    cdsEmitenteFONE.AsString        := edtFone.Text;
    cdsEmitente.Post;

    if not DirectoryExists(ExtractFileDir(ParamStr(0))+'\Emitente\') then begin
      ForceDirectories(ExtractFileDir(ParamStr(0))+'\Emitente\');
    end;

    if FileExists(ExtractFileDir(ParamStr(0))+'\Emitente\Emitente.xml') then
      DeleteFile(ExtractFileDir(ParamStr(0))+'\Emitente\Emitente.xml');

    cdsemitente.LogChanges := false;

    cdsemitente.SaveToFile( ExtractFileDir(ParamStr(0))+'\Emitente\Emitente.xml',dfXMLUTF8);

    showmessage('Gravação executada com sucesso!');
    self.Close;
  except
    on e: exception do begin
      showmessage('Erro ao gravar dados!');
    end;
  end;
end;

procedure TfrmEmitente.SpeedButton2Click(Sender: TObject);
begin
  self.Close;
end;

end.
