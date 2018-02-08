unit uSM;

interface

uses System.SysUtils, System.Classes, System.Json,
    DataSnap.DSProviderDataModuleAdapter,
    Datasnap.DSServer, Datasnap.DSAuth, REST.Json,
    untGerenciadorNFCe;

type
  TServerMethods1 = class(TDSServerModule)
  private
    { Private declarations }
  public
    { Public declarations }
    function EnviarNFCe(pDados: TJsonObject): TJsonObject;
  end;

implementation


{$R *.dfm}


{ TServerMethods1 }

function TServerMethods1.EnviarNFCe(pDados: TJsonObject): TJsonObject;
begin
  result := frmGerenciadorNFCe.EnviarNFCe(pDados);
end;

end.

