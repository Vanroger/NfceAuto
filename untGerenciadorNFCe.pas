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
    procedure PreencheNota;
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

procedure TfrmGerenciadorNFCe.PreencheNota;
begin
  try
    ACBrNFe1.NotasFiscais.Clear;
    with ACBrNFe1.NotasFiscais.Add.NFe do begin
//      versao        :=                   INIRec.ReadString('infNFe','versao', VersaoDFToStr(ACBrNFe1.Configuracoes.Geral.VersaoDF));
//      infNFe.versao := StringToFloatDef( INIRec.ReadString('infNFe','versao', VersaoDFToStr(ACBrNFe1.Configuracoes.Geral.VersaoDF)),0) ;
//
//      versao := infNFe.VersaoStr;
//      versao := StringReplace(versao,'versao="','',[rfReplaceAll,rfIgnoreCase]);
//      versao := StringReplace(versao,'"','',[rfReplaceAll,rfIgnoreCase]);
//
//      Ide.cNF        := INIRec.ReadInteger( 'Identificacao','Codigo' ,INIRec.ReadInteger( 'Identificacao','cNF' ,0));
//      Ide.natOp      := INIRec.ReadString(  'Identificacao','NaturezaOperacao' ,INIRec.ReadString(  'Identificacao','natOp' ,''));
//      Ide.indPag     := StrToIndpag(OK,INIRec.ReadString( 'Identificacao','FormaPag',INIRec.ReadString( 'Identificacao','indPag','0')));
//      Ide.modelo     := INIRec.ReadInteger( 'Identificacao','Modelo' ,INIRec.ReadInteger( 'Identificacao','mod' ,55));
//      ACBrNFe1.Configuracoes.Geral.ModeloDF := StrToModeloDF(OK,IntToStr(Ide.modelo));
//      ACBrNFe1.Configuracoes.Geral.VersaoDF := StrToVersaoDF(OK,versao);
//      Ide.serie      := INIRec.ReadInteger( 'Identificacao','Serie'  ,1);
//      Ide.nNF        := INIRec.ReadInteger( 'Identificacao','Numero' ,INIRec.ReadInteger( 'Identificacao','nNF' ,0));
//      Ide.dEmi       := StringToDateTime(INIRec.ReadString( 'Identificacao','Emissao',INIRec.ReadString( 'Identificacao','dEmi',INIRec.ReadString( 'Identificacao','dhEmi',FormatDateTimeBr(Now)))));
//      Ide.dSaiEnt    := StringToDateTime(INIRec.ReadString( 'Identificacao','Saida'  ,INIRec.ReadString( 'Identificacao','dSaiEnt'  ,INIRec.ReadString( 'Identificacao','dhSaiEnt','0'))));
//      Ide.hSaiEnt    := StringToDateTime(INIRec.ReadString( 'Identificacao','hSaiEnt','0'));  //NFe2
//      Ide.tpNF       := StrToTpNF(OK,INIRec.ReadString( 'Identificacao','Tipo',INIRec.ReadString( 'Identificacao','tpNF','1')));
//
//      Ide.idDest     := StrToDestinoOperacao(OK,INIRec.ReadString( 'Identificacao','idDest','1'));
//
//      Ide.tpImp      := StrToTpImp(  OK, INIRec.ReadString( 'Identificacao','tpImp',TpImpToStr(ACBrNFe1.DANFE.TipoDANFE)));  //NFe2
//      Ide.tpEmis     := StrToTpEmis( OK,INIRec.ReadString( 'Identificacao','tpEmis',IntToStr(ACBrNFe1.Configuracoes.Geral.FormaEmissaoCodigo)));
////      Ide.cDV
////      Ide.tpAmb
//      Ide.finNFe     := StrToFinNFe( OK,INIRec.ReadString( 'Identificacao','Finalidade',INIRec.ReadString( 'Identificacao','finNFe','0')));
//      Ide.indFinal   := StrToConsumidorFinal(OK,INIRec.ReadString( 'Identificacao','indFinal','0'));
//      Ide.indPres    := StrToPresencaComprador(OK,INIRec.ReadString( 'Identificacao','indPres','0'));
//
//      Ide.procEmi    := StrToProcEmi(OK,INIRec.ReadString( 'Identificacao','procEmi','0')); //NFe2
//      Ide.verProc    := INIRec.ReadString(  'Identificacao','verProc' ,'Kontrol Sistemas' );
//      Ide.dhCont     := StringToDateTime(INIRec.ReadString( 'Identificacao','dhCont'  ,'0')); //NFe2
//      Ide.xJust      := INIRec.ReadString(  'Identificacao','xJust' ,'' ); //NFe2
//
//      I := 1 ;
//      while true do begin
//         sSecao := 'NFRef'+IntToStrZero(I,3) ;
//         sFim   := INIRec.ReadString(  sSecao,'Tipo'  ,'FIM');
//         sTipo := UpperCase(INIRec.ReadString(  sSecao,'Tipo'  ,'NFe')); //NFe2 NF NFe NFP CTe ECF)
//         if (sFim = 'FIM') or (Length(sFim) <= 0) then begin
//           if INIRec.ReadString(sSecao,'refNFe','') <> '' then
//             sTipo := 'NFE';
//           break ;
//         end;
//
//         with Ide.NFref.Add do begin
//           if sTipo = 'NFE' then
//             refNFe :=  INIRec.ReadString(sSecao,'refNFe','');
//         end;
//         Inc(I);
//      end;

      Emit.CNPJCPF           := fEmitente.CNPJCPF;
      Emit.xNome             := fEmitente.xNome;
      Emit.xFant             := fEmitente.xFant;
      Emit.IE                := fEmitente.IE;
      Emit.IEST              := fEmitente.IEST;
      Emit.IM                := fEmitente.IM;
      Emit.CNAE              := fEmitente.CNAE;
      Emit.CRT               := fEmitente.CRT;
      Emit.EnderEmit.xLgr    := fEmitente.xLgr;
      Emit.EnderEmit.nro     := fEmitente.nro;
      Emit.EnderEmit.xCpl    := fEmitente.xCpl;
      Emit.EnderEmit.xBairro := fEmitente.xBairro;
      Emit.EnderEmit.cMun    := fEmitente.cMun;
      Emit.EnderEmit.xMun    := fEmitente.xMun;
      Emit.EnderEmit.UF      := fEmitente.UF;
      Emit.EnderEmit.CEP     := fEmitente.cep;
      Emit.EnderEmit.cPais   := '1058';   // fEmitente.cPais;
      Emit.EnderEmit.xPais   := 'BRASIL'; // fEmitente.xPais;
      Emit.EnderEmit.fone    := fEmitente.fone;

//      Ide.cUF       := INIRec.ReadInteger( 'Identificacao','cUF'       ,UFparaCodigo(Emit.EnderEmit.UF));
//      Ide.cMunFG    := INIRec.ReadInteger( 'Identificacao','CidadeCod' ,INIRec.ReadInteger( 'Identificacao','cMunFG' ,Emit.EnderEmit.cMun));
//
//      if INIRec.ReadString(  'Avulsa','CNPJ','') <> '' then
//       begin
//         Avulsa.CNPJ    := INIRec.ReadString(  'Avulsa','CNPJ','');
//         Avulsa.xOrgao  := INIRec.ReadString(  'Avulsa','xOrgao','');
//         Avulsa.matr    := INIRec.ReadString(  'Avulsa','matr','');
//         Avulsa.xAgente := INIRec.ReadString(  'Avulsa','xAgente','');
//         Avulsa.fone    := INIRec.ReadString(  'Avulsa','fone','');
//         Avulsa.UF      := INIRec.ReadString(  'Avulsa','UF','');
//         Avulsa.nDAR    := INIRec.ReadString(  'Avulsa','nDAR','');
//         Avulsa.dEmi    := StringToDateTime(INIRec.ReadString(  'Avulsa','dEmi','0'));
//         Avulsa.vDAR    := StringToFloatDef(INIRec.ReadString(  'Avulsa','vDAR',''),0);
//         Avulsa.repEmi  := INIRec.ReadString(  'Avulsa','repEmi','');
//         Avulsa.dPag    := StringToDateTime(INIRec.ReadString(  'Avulsa','dPag','0'));
//       end;
//
//      Dest.idEstrangeiro     := INIRec.ReadString(  'Destinatario','idEstrangeiro','');
//      Dest.CNPJCPF           := INIRec.ReadString(  'Destinatario','CNPJ'       ,INIRec.ReadString(  'Destinatario','CNPJCPF',INIRec.ReadString(  'Destinatario','CPF','')));
//      Dest.xNome             := INIRec.ReadString(  'Destinatario','NomeRazao'  ,INIRec.ReadString(  'Destinatario','xNome'  ,''));
//      Dest.indIEDest         := StrToindIEDest(OK,INIRec.ReadString( 'Destinatario','indIEDest','1'));
//      Dest.IE                := INIRec.ReadString(  'Destinatario','IE'         ,'');
//      Dest.ISUF              := INIRec.ReadString(  'Destinatario','ISUF'       ,'');
//      Dest.Email             := INIRec.ReadString(  'Destinatario','Email'      ,'');  //NFe2
//
//      Dest.EnderDest.xLgr    := INIRec.ReadString(  'Destinatario','Logradouro' ,INIRec.ReadString(  'Destinatario','xLgr' ,''));
//      if (INIRec.ReadString('Destinatario','Numero','') <> '') or (INIRec.ReadString('Destinatario','nro','') <> '') then
//         Dest.EnderDest.nro     := INIRec.ReadString(  'Destinatario','Numero'     ,INIRec.ReadString('Destinatario','nro',''));
//      if (INIRec.ReadString('Destinatario','Complemento','') <> '') or (INIRec.ReadString('Destinatario','xCpl','') <> '') then
//         Dest.EnderDest.xCpl    := INIRec.ReadString(  'Destinatario','Complemento',INIRec.ReadString('Destinatario','xCpl',''));
//      Dest.EnderDest.xBairro := INIRec.ReadString(  'Destinatario','Bairro'     ,INIRec.ReadString(  'Destinatario','xBairro',''));
//      Dest.EnderDest.cMun    := INIRec.ReadInteger( 'Destinatario','CidadeCod'  ,INIRec.ReadInteger( 'Destinatario','cMun'   ,0));
//      Dest.EnderDest.xMun    := INIRec.ReadString(  'Destinatario','Cidade'     ,INIRec.ReadString(  'Destinatario','xMun'   ,''));
//      Dest.EnderDest.UF      := INIRec.ReadString(  'Destinatario','UF'         ,'');
//      Dest.EnderDest.CEP     := INIRec.ReadInteger( 'Destinatario','CEP'       ,0);
////      if Dest.EnderDest.cMun <= 0 then
////         Dest.EnderDest.cMun := ObterCodigoMunicipio(Dest.EnderDest.xMun,Dest.EnderDest.UF);
//      Dest.EnderDest.cPais   := INIRec.ReadInteger( 'Destinatario','PaisCod'    ,INIRec.ReadInteger('Destinatario','cPais',1058));
//      Dest.EnderDest.xPais   := INIRec.ReadString(  'Destinatario','Pais'       ,INIRec.ReadString( 'Destinatario','xPais','BRASIL'));
//      Dest.EnderDest.Fone    := INIRec.ReadString(  'Destinatario','Fone'       ,'');
//
//      I := 1 ;
//
//      while true do begin
//         sSecao    := 'Produto'+IntToStrZero(I,3) ;
//         sCodPro   := INIRec.ReadString(sSecao,'Codigo',INIRec.ReadString( sSecao,'cProd','FIM')) ;
//         if sCodPro = 'FIM' then
//            break ;
//
//         with Det.Add do begin
//            Prod.nItem := I;
//            infAdProd      := INIRec.ReadString(sSecao,'infAdProd','');
//
//            Prod.cProd    := INIRec.ReadString( sSecao,'Codigo'   ,INIRec.ReadString( sSecao,'cProd'   ,''));
//            if (Length(INIRec.ReadString( sSecao,'EAN','')) > 0) or (Length(INIRec.ReadString( sSecao,'cEAN','')) > 0)  then
//               Prod.cEAN      := INIRec.ReadString( sSecao,'EAN'      ,INIRec.ReadString( sSecao,'cEAN'      ,''));
//            Prod.xProd    := INIRec.ReadString( sSecao,'Descricao',INIRec.ReadString( sSecao,'xProd',''));
//            Prod.NCM      := INIRec.ReadString( sSecao,'NCM'      ,'');
//            Prod.CEST     := INIRec.ReadString( sSecao,'CEST'      ,'');
//            Prod.EXTIPI   := INIRec.ReadString( sSecao,'EXTIPI'      ,'');
//            Prod.CFOP     := INIRec.ReadString( sSecao,'CFOP'     ,'');
//            Prod.uCom     := INIRec.ReadString( sSecao,'Unidade'  ,INIRec.ReadString( sSecao,'uCom'  ,''));
//            Prod.qCom     := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'   ,INIRec.ReadString(sSecao,'qCom'  ,'')) ,0) ;
//            Prod.vUnCom   := StringToFloatDef( INIRec.ReadString(sSecao,'ValorUnitario',INIRec.ReadString(sSecao,'vUnCom','')) ,0) ;
//            Prod.vProd    := StringToFloatDef( INIRec.ReadString(sSecao,'ValorTotal'   ,INIRec.ReadString(sSecao,'vProd' ,'')) ,0) ;
//
//            if Length(INIRec.ReadString( sSecao,'cEANTrib','')) > 0 then
//               Prod.cEANTrib      := INIRec.ReadString( sSecao,'cEANTrib'      ,'');
//            Prod.uTrib     := INIRec.ReadString( sSecao,'uTrib'  , Prod.uCom);
//            Prod.qTrib     := StringToFloatDef( INIRec.ReadString(sSecao,'qTrib'  ,''), Prod.qCom);
//            Prod.vUnTrib   := StringToFloatDef( INIRec.ReadString(sSecao,'vUnTrib','') ,Prod.vUnCom) ;
//
//            Prod.vFrete    := StringToFloatDef( INIRec.ReadString(sSecao,'vFrete','') ,0) ;
//            Prod.vSeg      := StringToFloatDef( INIRec.ReadString(sSecao,'vSeg','') ,0) ;
//            Prod.vDesc     := StringToFloatDef( INIRec.ReadString(sSecao,'ValorDesconto',INIRec.ReadString(sSecao,'vDesc','')) ,0) ;
//            Prod.vOutro    := StringToFloatDef( INIRec.ReadString(sSecao,'vOutro','') ,0) ; //NFe2
//            Prod.IndTot    := StrToindTot(OK,INIRec.ReadString(sSecao,'indTot','1'));       //NFe2
//
//            Prod.xPed      := INIRec.ReadString( sSecao,'xPed'    , '');  //NFe2
//            Prod.nItemPed  := INIRec.ReadString( sSecao,'nItemPed', '');  //NFe2
//
//            Prod.nFCI      := INIRec.ReadString( sSecao,'nFCI','');  //NFe3
//            Prod.nRECOPI   := INIRec.ReadString( sSecao,'nRECOPI','');  //NFe3
//
//            pDevol := StringToFloatDef( INIRec.ReadString(sSecao,'pDevol','') ,0);
//            vIPIDevol := StringToFloatDef( INIRec.ReadString(sSecao,'vIPIDevol','') ,0);
//
//            Imposto.vTotTrib := StringToFloatDef( INIRec.ReadString(sSecao,'vTotTrib','') ,0) ; //NFe2
//
//            J := 1 ;
//            while true do
//             begin
//               sSecao  := 'NVE'+IntToStrZero(I,3)+IntToStrZero(J,3) ;
//               sNVE    := INIRec.ReadString(sSecao,'NVE','') ;
//               if (sNVE <> '') then
//                  Prod.NVE.Add.NVE := sNVE
//               else
//                  Break;
//               Inc(J);
//             end;
//
//            J := 1 ;
//            while true do
//             begin
//               sSecao      := 'DI'+IntToStrZero(I,3)+IntToStrZero(J,3) ;
//               sNumeroDI := INIRec.ReadString(sSecao,'NumeroDI',INIRec.ReadString(sSecao,'nDi','')) ;
//
//               if sNumeroDI <> '' then
//                begin
//                  with Prod.DI.Add do
//                   begin
//                     nDi         := sNumeroDI;
//                     dDi         := StringToDateTime(INIRec.ReadString(sSecao,'DataRegistroDI'  ,INIRec.ReadString(sSecao,'dDi'  ,'0')));
//                     xLocDesemb  := INIRec.ReadString(sSecao,'LocalDesembaraco',INIRec.ReadString(sSecao,'xLocDesemb',''));
//                     UFDesemb    := INIRec.ReadString(sSecao,'UFDesembaraco'   ,INIRec.ReadString(sSecao,'UFDesemb'   ,''));
//                     dDesemb     := StringToDateTime(INIRec.ReadString(sSecao,'DataDesembaraco',INIRec.ReadString(sSecao,'dDesemb','0')));
//
//                     tpViaTransp  := StrToTipoViaTransp(OK,INIRec.ReadString(sSecao,'tpViaTransp',''));
//                     vAFRMM       := StringToFloatDef( INIRec.ReadString(sSecao,'vAFRMM','') ,0) ;
//                     tpIntermedio := StrToTipoIntermedio(OK,INIRec.ReadString(sSecao,'tpIntermedio',''));
//                     CNPJ         := INIRec.ReadString(sSecao,'CNPJ','');
//                     UFTerceiro   := INIRec.ReadString(sSecao,'UFTerceiro','');
//
//                     cExportador := INIRec.ReadString(sSecao,'CodigoExportador',INIRec.ReadString(sSecao,'cExportador',''));
//
//                     K := 1 ;
//                     while true do
//                      begin
//                        sSecao      := 'LADI'+IntToStrZero(I,3)+IntToStrZero(J,3)+IntToStrZero(K,3)  ;
//                        sNumeroADI := INIRec.ReadString(sSecao,'NumeroAdicao',INIRec.ReadString(sSecao,'nAdicao','FIM')) ;
//                        if (sNumeroADI = 'FIM') or (Length(sNumeroADI) <= 0) then
//                           break;
//
//                        with adi.Add do
//                         begin
//                           nAdicao     := StrToInt(sNumeroADI);
//                           nSeqAdi     := INIRec.ReadInteger( sSecao,'nSeqAdi',K);
//                           cFabricante := INIRec.ReadString(  sSecao,'CodigoFabricante',INIRec.ReadString(  sSecao,'cFabricante',''));
//                           vDescDI     := StringToFloatDef( INIRec.ReadString(sSecao,'DescontoADI',INIRec.ReadString(sSecao,'vDescDI','')) ,0);
//                           nDraw       := INIRec.ReadString( sSecao,'nDraw','');
//                         end;
//                        Inc(K)
//                      end;
//                   end;
//                end
//               else
//                 Break;
//               Inc(J);
//             end;
//
//            J := 1 ;
//            while true do
//             begin
//               sSecao  := 'detExport'+IntToStrZero(I,3)+IntToStrZero(J,3) ;
//               sFim    := INIRec.ReadString(sSecao,'nDraw',INIRec.ReadString(sSecao,'nRE','FIM')) ;
//               if (sFim = 'FIM') or (Length(sFim) <= 0) then
//                  break ;
//
//               with Prod.detExport.Add do
//                begin
//                  nDraw       := INIRec.ReadString( sSecao,'nDraw','');
//                  nRE         := INIRec.ReadString( sSecao,'nRE','');
//                  chNFe       := INIRec.ReadString( sSecao,'chNFe','');
//                  qExport     := StringToFloatDef( INIRec.ReadString(sSecao,'qExport','') ,0);
//                end;
//               Inc(J);
//             end;
//
//           sSecao := 'impostoDevol'+IntToStrZero(I,3) ;
//           sFim   := INIRec.ReadString( sSecao,'pDevol','FIM') ;
//           if (sFim <> 'FIM') then
//            begin
//              pDevol := StringToFloatDef( INIRec.ReadString(sSecao,'pDevol','') ,0);
//              vIPIDevol := StringToFloatDef( INIRec.ReadString(sSecao,'vIPIDevol','') ,0);
//            end;
//
//            sSecao := 'Combustivel'+IntToStrZero(I,3) ;
//            sFim   := INIRec.ReadString( sSecao,'cProdANP','FIM') ;
//            if (sFim <> 'FIM') then begin
//              with Prod.comb do begin
//                 cProdANP := INIRec.ReadInteger( sSecao,'cProdANP',0) ;
//                 pMixGN   := StringToFloatDef(INIRec.ReadString( sSecao,'pMixGN',''),0) ;
//                 CODIF    := INIRec.ReadString(  sSecao,'CODIF'   ,'') ;
//                 qTemp    := StringToFloatDef(INIRec.ReadString( sSecao,'qTemp',''),0) ;
//                 UFcons   := INIRec.ReadString( sSecao,'UFCons','') ;
//
//                 sSecao := 'CIDE'+IntToStrZero(I,3) ;
//                 CIDE.qBCprod   := StringToFloatDef(INIRec.ReadString( sSecao,'qBCprod'  ,''),0) ;
//                 CIDE.vAliqProd := StringToFloatDef(INIRec.ReadString( sSecao,'vAliqProd',''),0) ;
//                 CIDE.vCIDE     := StringToFloatDef(INIRec.ReadString( sSecao,'vCIDE'    ,''),0) ;
//
//                 sSecao := 'encerrante'+IntToStrZero(I,3) ;
//                 encerrante.nBico    := INIRec.ReadInteger( sSecao,'nBico'  ,0) ;
//                 encerrante.nBomba   := INIRec.ReadInteger( sSecao,'nBomba' ,0) ;
//                 encerrante.nTanque  := INIRec.ReadInteger( sSecao,'nTanque',0) ;
//                 encerrante.vEncIni  := INIRec.ReadFloat( sSecao,'vEncIni',0) ;
//                 encerrante.vEncFin  := INIRec.ReadFloat( sSecao,'vEncFin',0) ;
//
//                 sSecao := 'ICMSComb'+IntToStrZero(I,3) ;
//                 ICMS.vBCICMS   := StringToFloatDef(INIRec.ReadString( sSecao,'vBCICMS'  ,''),0) ;
//                 ICMS.vICMS     := StringToFloatDef(INIRec.ReadString( sSecao,'vICMS'    ,''),0) ;
//                 ICMS.vBCICMSST := StringToFloatDef(INIRec.ReadString( sSecao,'vBCICMSST',''),0) ;
//                 ICMS.vICMSST   := StringToFloatDef(INIRec.ReadString( sSecao,'vICMSST'  ,''),0) ;
//
//                 sSecao := 'ICMSInter'+IntToStrZero(I,3) ;
//                 sFim   := INIRec.ReadString( sSecao,'vBCICMSSTDest','FIM') ;
//                 if (sFim <> 'FIM') then
//                  begin
//                    ICMSInter.vBCICMSSTDest := StringToFloatDef(sFim,0) ;
//                    ICMSInter.vICMSSTDest   := StringToFloatDef(INIRec.ReadString( sSecao,'vICMSSTDest',''),0) ;
//                  end;
//
//                 sSecao := 'ICMSCons'+IntToStrZero(I,3) ;
//                 sFim   := INIRec.ReadString( sSecao,'vBCICMSSTCons','FIM') ;
//                 if (sFim <> 'FIM') then
//                  begin
//                    ICMSCons.vBCICMSSTCons := StringToFloatDef(sFim,0) ;
//                    ICMSCons.vICMSSTCons   := StringToFloatDef(INIRec.ReadString( sSecao,'vICMSSTCons',''),0) ;
//                    ICMSCons.UFcons        := INIRec.ReadString( sSecao,'UFCons','') ;
//                  end;
//              end;
//            end;
//
//            with Imposto do
//             begin
//                sSecao := 'ICMS'+IntToStrZero(I,3) ;
//                sFim   := INIRec.ReadString( sSecao,'CST',INIRec.ReadString(sSecao,'CSOSN','FIM')) ;
//                if (sFim <> 'FIM') then
//                 begin
//                   with ICMS do
//                   begin
//                     ICMS.orig       := StrToOrig(     OK, INIRec.ReadString(sSecao,'Origem'    ,INIRec.ReadString(sSecao,'orig'    ,'0' ) ));
//                     CST             := StrToCSTICMS(  OK, INIRec.ReadString(sSecao,'CST'       ,'00'));
//                     CSOSN           := StrToCSOSNIcms(OK, INIRec.ReadString(sSecao,'CSOSN'     ,''  ));     //NFe2
//                     ICMS.modBC      := StrTomodBC(    OK, INIRec.ReadString(sSecao,'Modalidade',INIRec.ReadString(sSecao,'modBC','0' ) ));
//                     ICMS.pRedBC     := StringToFloatDef( INIRec.ReadString(sSecao,'PercentualReducao',INIRec.ReadString(sSecao,'pRedBC','')) ,0);
//                     ICMS.vBC        := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase',INIRec.ReadString(sSecao,'vBC'  ,'')) ,0);
//                     ICMS.pICMS      := StringToFloatDef( INIRec.ReadString(sSecao,'Aliquota' ,INIRec.ReadString(sSecao,'pICMS','')) ,0);
//                     ICMS.vICMS      := StringToFloatDef( INIRec.ReadString(sSecao,'Valor'    ,INIRec.ReadString(sSecao,'vICMS','')) ,0);
//                     ICMS.modBCST    := StrTomodBCST(OK, INIRec.ReadString(sSecao,'ModalidadeST',INIRec.ReadString(sSecao,'modBCST','0')));
//                     ICMS.pMVAST     := StringToFloatDef( INIRec.ReadString(sSecao,'PercentualMargemST' ,INIRec.ReadString(sSecao,'pMVAST' ,'')) ,0);
//                     ICMS.pRedBCST   := StringToFloatDef( INIRec.ReadString(sSecao,'PercentualReducaoST',INIRec.ReadString(sSecao,'pRedBCST','')) ,0);
//                     ICMS.vBCST      := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBaseST',INIRec.ReadString(sSecao,'vBCST','')) ,0);
//                     ICMS.pICMSST    := StringToFloatDef( INIRec.ReadString(sSecao,'AliquotaST' ,INIRec.ReadString(sSecao,'pICMSST' ,'')) ,0);
//                     ICMS.vICMSST    := StringToFloatDef( INIRec.ReadString(sSecao,'ValorST'    ,INIRec.ReadString(sSecao,'vICMSST'    ,'')) ,0);
//                     ICMS.UFST       := INIRec.ReadString(sSecao,'UFST'    ,'');                           //NFe2
//                     ICMS.pBCOp      := StringToFloatDef( INIRec.ReadString(sSecao,'pBCOp'    ,'') ,0);    //NFe2
//                     ICMS.vBCSTRet   := StringToFloatDef( INIRec.ReadString(sSecao,'vBCSTRet','') ,0);     //NFe2
//                     ICMS.vICMSSTRet := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSSTRet','') ,0);   //NFe2
//                     ICMS.motDesICMS := StrTomotDesICMS(OK, INIRec.ReadString(sSecao,'motDesICMS','0'));   //NFe2
//                     ICMS.pCredSN    := StringToFloatDef( INIRec.ReadString(sSecao,'pCredSN','') ,0);      //NFe2
//                     ICMS.vCredICMSSN:= StringToFloatDef( INIRec.ReadString(sSecao,'vCredICMSSN','') ,0);  //NFe2
//                     ICMS.vBCSTDest  := StringToFloatDef( INIRec.ReadString(sSecao,'vBCSTDest','') ,0);    //NFe2
//                     ICMS.vICMSSTDest:= StringToFloatDef( INIRec.ReadString(sSecao,'vICMSSTDest','') ,0);   //NFe2
//                     ICMS.vICMSDeson := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSDeson','') ,0);
//                     ICMS.vICMSOp    := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSOp','') ,0);
//                     ICMS.pDif       := StringToFloatDef( INIRec.ReadString(sSecao,'pDif','') ,0);
//                     ICMS.vICMSDif   := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSDif','') ,0);
//                   end;
//                 end;
//
//                sSecao := 'ICMSUFDEST'+IntToStrZero(I,3) ;
//                sFim   := INIRec.ReadString( sSecao,'vBCUFDest','FIM') ;
//                if (sFim <> 'FIM') then
//                 begin
//                   with ICMSUFDest do
//                   begin
//                     vBCUFDest      := StringToFloatDef( INIRec.ReadString(sSecao,'vBCUFDest','') ,0);
//                     pICMSUFDest    := StringToFloatDef( INIRec.ReadString(sSecao,'pICMSUFDest','') ,0);
//                     pICMSInter     := StringToFloatDef( INIRec.ReadString(sSecao,'pICMSInter','') ,0);
//                     pICMSInterPart := StringToFloatDef( INIRec.ReadString(sSecao,'pICMSInterPart','') ,0);
//                     vICMSUFDest    := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSUFDest','') ,0);
//                     vICMSUFRemet   := StringToFloatDef( INIRec.ReadString(sSecao,'vICMSUFRemet','') ,0);
//                     pFCPUFDest     := StringToFloatDef( INIRec.ReadString(sSecao,'pFCPUFDest','') ,0);
//                     vFCPUFDest     := StringToFloatDef( INIRec.ReadString(sSecao,'vFCPUFDest','') ,0);
//                   end;
//                 end;
//
//                sSecao := 'IPI'+IntToStrZero(I,3) ;
//                sFim  := INIRec.ReadString( sSecao,'CST','FIM') ;
//                if (sFim <> 'FIM') then
//                 begin
//                  with IPI do
//                   begin
//                     CST      := StrToCSTIPI(OK, INIRec.ReadString( sSecao,'CST','')) ;
//                     clEnq    := INIRec.ReadString(  sSecao,'ClasseEnquadramento',INIRec.ReadString(  sSecao,'clEnq'   ,''));
//                     CNPJProd := INIRec.ReadString(  sSecao,'CNPJProdutor'       ,INIRec.ReadString(  sSecao,'CNPJProd',''));
//                     cSelo    := INIRec.ReadString(  sSecao,'CodigoSeloIPI'      ,INIRec.ReadString(  sSecao,'cSelo'   ,''));
//                     qSelo    := INIRec.ReadInteger( sSecao,'QuantidadeSelos'    ,INIRec.ReadInteger( sSecao,'qSelo'   ,0));
//                     cEnq     := INIRec.ReadString(  sSecao,'CodigoEnquadramento',INIRec.ReadString(  sSecao,'cEnq'    ,''));
//
//                     vBC    := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'   ,INIRec.ReadString(sSecao,'vBC'   ,'')) ,0);
//                     qUnid  := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'  ,INIRec.ReadString(sSecao,'qUnid' ,'')) ,0);
//                     vUnid  := StringToFloatDef( INIRec.ReadString(sSecao,'ValorUnidade',INIRec.ReadString(sSecao,'vUnid' ,'')) ,0);
//                     pIPI   := StringToFloatDef( INIRec.ReadString(sSecao,'Aliquota'    ,INIRec.ReadString(sSecao,'pIPI'  ,'')) ,0);
//                     vIPI   := StringToFloatDef( INIRec.ReadString(sSecao,'Valor'       ,INIRec.ReadString(sSecao,'vIPI'  ,'')) ,0);
//                   end;
//                 end;
//
//                sSecao   := 'II'+IntToStrZero(I,3) ;
//                sFim   := INIRec.ReadString( sSecao,'ValorBase',INIRec.ReadString( sSecao,'vBC','FIM')) ;
//                if (sFim <> 'FIM') then
//                 begin
//                  with II do
//                   begin
//                     vBc      := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'          ,INIRec.ReadString(sSecao,'vBC'     ,'')) ,0);
//                     vDespAdu := StringToFloatDef( INIRec.ReadString(sSecao,'ValorDespAduaneiras',INIRec.ReadString(sSecao,'vDespAdu','')) ,0);
//                     vII      := StringToFloatDef( INIRec.ReadString(sSecao,'ValorII'            ,INIRec.ReadString(sSecao,'vII'     ,'')) ,0);
//                     vIOF     := StringToFloatDef( INIRec.ReadString(sSecao,'ValorIOF'           ,INIRec.ReadString(sSecao,'vIOF'    ,'')) ,0);
//                   end;
//                 end;
//
//                sSecao    := 'PIS'+IntToStrZero(I,3) ;
//                sFim   := INIRec.ReadString( sSecao,'CST','FIM') ;
//                if (sFim <> 'FIM') then
//                 begin
//                  with PIS do
//                    begin
//                     CST :=  StrToCSTPIS(OK, INIRec.ReadString( sSecao,'CST','01'));
//
//                     PIS.vBC       := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'    ,INIRec.ReadString(sSecao,'vBC'      ,'')) ,0);
//                     PIS.pPIS      := StringToFloatDef( INIRec.ReadString(sSecao,'Aliquota'     ,INIRec.ReadString(sSecao,'pPIS'     ,'')) ,0);
//                     PIS.qBCProd   := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'   ,INIRec.ReadString(sSecao,'qBCProd'  ,'')) ,0);
//                     PIS.vAliqProd := StringToFloatDef( INIRec.ReadString(sSecao,'ValorAliquota',INIRec.ReadString(sSecao,'vAliqProd','')) ,0);
//                     PIS.vPIS      := StringToFloatDef( INIRec.ReadString(sSecao,'Valor'        ,INIRec.ReadString(sSecao,'vPIS'     ,'')) ,0);
//                    end;
//                 end;
//
//                sSecao    := 'PISST'+IntToStrZero(I,3) ;
//                sFim   := INIRec.ReadString( sSecao,'ValorBase','F')+ INIRec.ReadString( sSecao,'Quantidade','IM') ;
//                if (sFim = 'FIM') then
//                   sFim   := INIRec.ReadString( sSecao,'vBC','F')+ INIRec.ReadString( sSecao,'qBCProd','IM') ;
//
//                if (sFim <> 'FIM') then
//                 begin
//                  with PISST do
//                   begin
//                     vBc       := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'    ,INIRec.ReadString(sSecao,'vBC'      ,'')) ,0);
//                     pPis      := StringToFloatDef( INIRec.ReadString(sSecao,'AliquotaPerc' ,INIRec.ReadString(sSecao,'pPis'     ,'')) ,0);
//                     qBCProd   := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'   ,INIRec.ReadString(sSecao,'qBCProd'  ,'')) ,0);
//                     vAliqProd := StringToFloatDef( INIRec.ReadString(sSecao,'AliquotaValor',INIRec.ReadString(sSecao,'vAliqProd','')) ,0);
//                     vPIS      := StringToFloatDef( INIRec.ReadString(sSecao,'ValorPISST'   ,INIRec.ReadString(sSecao,'vPIS'     ,'')) ,0);
//                   end;
//                 end;
//
//                sSecao    := 'COFINS'+IntToStrZero(I,3) ;
//                sFim   := INIRec.ReadString( sSecao,'CST','FIM') ;
//                if (sFim <> 'FIM') then
//                 begin
//                  with COFINS do
//                   begin
//                     CST := StrToCSTCOFINS(OK, INIRec.ReadString( sSecao,'CST','01'));
//
//                     COFINS.vBC       := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'    ,INIRec.ReadString(sSecao,'vBC'      ,'')) ,0);
//                     COFINS.pCOFINS   := StringToFloatDef( INIRec.ReadString(sSecao,'Aliquota'     ,INIRec.ReadString(sSecao,'pCOFINS'  ,'')) ,0);
//                     COFINS.qBCProd   := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'   ,INIRec.ReadString(sSecao,'qBCProd'  ,'')) ,0);
//                     COFINS.vAliqProd := StringToFloatDef( INIRec.ReadString(sSecao,'ValorAliquota',INIRec.ReadString(sSecao,'vAliqProd','')) ,0);
//                     COFINS.vCOFINS   := StringToFloatDef( INIRec.ReadString(sSecao,'Valor'        ,INIRec.ReadString(sSecao,'vCOFINS'  ,'')) ,0);
//                   end;
//                 end;
//
//                sSecao    := 'COFINSST'+IntToStrZero(I,3) ;
//                sFim   := INIRec.ReadString( sSecao,'ValorBase','F')+ INIRec.ReadString( sSecao,'Quantidade','IM');
//                if (sFim = 'FIM') then
//                   sFim   := INIRec.ReadString( sSecao,'vBC','F')+ INIRec.ReadString( sSecao,'qBCProd','IM') ;
//
//                if (sFim <> 'FIM') then
//                 begin
//                  with COFINSST do
//                   begin
//                      vBC       := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'    ,INIRec.ReadString(sSecao,'vBC'      ,'')) ,0);
//                      pCOFINS   := StringToFloatDef( INIRec.ReadString(sSecao,'AliquotaPerc' ,INIRec.ReadString(sSecao,'pCOFINS'  ,'')) ,0);
//                      qBCProd   := StringToFloatDef( INIRec.ReadString(sSecao,'Quantidade'   ,INIRec.ReadString(sSecao,'qBCProd'  ,'')) ,0);
//                      vAliqProd := StringToFloatDef( INIRec.ReadString(sSecao,'AliquotaValor',INIRec.ReadString(sSecao,'vAliqProd','')) ,0);
//                      vCOFINS   := StringToFloatDef( INIRec.ReadString(sSecao,'ValorCOFINSST',INIRec.ReadString(sSecao,'vCOFINS'  ,'')) ,0);
//                    end;
//                 end;
//
//                sSecao    := 'ISSQN'+IntToStrZero(I,3) ;
//                sFim   := INIRec.ReadString( sSecao,'ValorBase',INIRec.ReadString(sSecao,'vBC'   ,'FIM')) ;
//                if (sFim = 'FIM') then
//                   sFim   := INIRec.ReadString( sSecao,'vBC','FIM');
//                if (sFim <> 'FIM') then
//                 begin
//                  with ISSQN do
//                   begin
//                     if StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase',INIRec.ReadString(sSecao,'vBC','')) ,0) > 0 then
//                      begin
//                        vBC       := StringToFloatDef( INIRec.ReadString(sSecao,'ValorBase'   ,INIRec.ReadString(sSecao,'vBC'   ,'')) ,0);
//                        vAliq     := StringToFloatDef( INIRec.ReadString(sSecao,'Aliquota'    ,INIRec.ReadString(sSecao,'vAliq' ,'')) ,0);
//                        vISSQN    := StringToFloatDef( INIRec.ReadString(sSecao,'ValorISSQN'  ,INIRec.ReadString(sSecao,'vISSQN','')) ,0);
//                        cMunFG    := INIRec.ReadInteger(sSecao,'MunicipioFatoGerador',INIRec.ReadInteger(sSecao,'cMunFG',0));
//                        cListServ := INIRec.ReadString(sSecao,'CodigoServico',INIRec.ReadString(sSecao,'cListServ',''));
//                        cSitTrib  := StrToISSQNcSitTrib( OK,INIRec.ReadString(sSecao,'cSitTrib','')) ;
//                        vDeducao    := StringToFloatDef( INIRec.ReadString(sSecao,'ValorDeducao'   ,INIRec.ReadString(sSecao,'vDeducao'   ,'')) ,0);
//                        vOutro      := StringToFloatDef( INIRec.ReadString(sSecao,'ValorOutro'   ,INIRec.ReadString(sSecao,'vOutro'   ,'')) ,0);
//                        vDescIncond := StringToFloatDef( INIRec.ReadString(sSecao,'ValorDescontoIncondicional'   ,INIRec.ReadString(sSecao,'vDescIncond'   ,'')) ,0);
//                        vDescCond   := StringToFloatDef( INIRec.ReadString(sSecao,'vDescontoCondicional'   ,INIRec.ReadString(sSecao,'vDescCond'   ,'')) ,0);
//                        vISSRet     := StringToFloatDef( INIRec.ReadString(sSecao,'ValorISSRetido'   ,INIRec.ReadString(sSecao,'vISSRet'   ,'')) ,0);
//                        indISS      := StrToindISS( OK,INIRec.ReadString(sSecao,'indISS','')) ;
//                        cServico    := INIRec.ReadString(sSecao,'cServico','');
//                        cMun        := INIRec.ReadInteger(sSecao,'cMun',0);
//                        cPais       := INIRec.ReadInteger(sSecao,'cPais',1058);
//                        nProcesso   := INIRec.ReadString(sSecao,'nProcesso','');
//                        indIncentivo := StrToindIncentivo( OK,INIRec.ReadString(sSecao,'indIncentivo','')) ;
//                      end;
//                   end;
//                 end;
//             end;
//
//          end;
//         Inc( I ) ;
//       end ;
//
//      Total.ICMSTot.vBC     := StringToFloatDef( INIRec.ReadString('Total','BaseICMS'     ,INIRec.ReadString('Total','vBC'     ,'')) ,0) ;
//      Total.ICMSTot.vICMS   := StringToFloatDef( INIRec.ReadString('Total','ValorICMS'    ,INIRec.ReadString('Total','vICMS'   ,'')) ,0) ;
//      Total.ICMSTot.vICMSDeson := StringToFloatDef( INIRec.ReadString('Total','vICMSDeson',''),0) ;
//      Total.ICMSTot.vICMSUFDest := StringToFloatDef( INIRec.ReadString('Total','vICMSUFDest',''),0) ;
//      Total.ICMSTot.vICMSUFRemet := StringToFloatDef( INIRec.ReadString('Total','vICMSUFRemet',''),0) ;
//      Total.ICMSTot.vFCPUFDest :=  StringToFloatDef( INIRec.ReadString('Total','vFCPUFDest',''),0) ;
//      Total.ICMSTot.vBCST   := StringToFloatDef( INIRec.ReadString('Total','BaseICMSSubstituicao' ,INIRec.ReadString('Total','vBCST','')) ,0) ;
//      Total.ICMSTot.vST     := StringToFloatDef( INIRec.ReadString('Total','ValorICMSSubstituicao',INIRec.ReadString('Total','vST'  ,'')) ,0) ;
//      Total.ICMSTot.vProd   := StringToFloatDef( INIRec.ReadString('Total','ValorProduto' ,INIRec.ReadString('Total','vProd'  ,'')) ,0) ;
//      Total.ICMSTot.vFrete  := StringToFloatDef( INIRec.ReadString('Total','ValorFrete'   ,INIRec.ReadString('Total','vFrete' ,'')) ,0) ;
//      Total.ICMSTot.vSeg    := StringToFloatDef( INIRec.ReadString('Total','ValorSeguro'  ,INIRec.ReadString('Total','vSeg'   ,'')) ,0) ;
//      Total.ICMSTot.vDesc   := StringToFloatDef( INIRec.ReadString('Total','ValorDesconto',INIRec.ReadString('Total','vDesc'  ,'')) ,0) ;
//      Total.ICMSTot.vII     := StringToFloatDef( INIRec.ReadString('Total','ValorII'      ,INIRec.ReadString('Total','vII'    ,'')) ,0) ;
//      Total.ICMSTot.vIPI    := StringToFloatDef( INIRec.ReadString('Total','ValorIPI'     ,INIRec.ReadString('Total','vIPI'   ,'')) ,0) ;
//      Total.ICMSTot.vPIS    := StringToFloatDef( INIRec.ReadString('Total','ValorPIS'     ,INIRec.ReadString('Total','vPIS'   ,'')) ,0) ;
//      Total.ICMSTot.vCOFINS := StringToFloatDef( INIRec.ReadString('Total','ValorCOFINS'  ,INIRec.ReadString('Total','vCOFINS','')) ,0) ;
//      Total.ICMSTot.vOutro  := StringToFloatDef( INIRec.ReadString('Total','ValorOutrasDespesas',INIRec.ReadString('Total','vOutro','')) ,0) ;
//      Total.ICMSTot.vNF     := StringToFloatDef( INIRec.ReadString('Total','ValorNota'    ,INIRec.ReadString('Total','vNF'    ,'')) ,0) ;
//      Total.ICMSTot.vTotTrib:= StringToFloatDef( INIRec.ReadString('Total','vTotTrib'     ,''),0) ;
//
//      Total.ISSQNtot.vServ  := StringToFloatDef( INIRec.ReadString('Total','ValorServicos',INIRec.ReadString('ISSQNtot','vServ','')) ,0) ;
//      Total.ISSQNTot.vBC    := StringToFloatDef( INIRec.ReadString('Total','ValorBaseISS' ,INIRec.ReadString('ISSQNtot','vBC'  ,'')) ,0) ;
//      Total.ISSQNTot.vISS   := StringToFloatDef( INIRec.ReadString('Total','ValorISSQN'   ,INIRec.ReadString('ISSQNtot','vISS' ,'')) ,0) ;
//      Total.ISSQNTot.vPIS   := StringToFloatDef( INIRec.ReadString('Total','ValorPISISS'  ,INIRec.ReadString('ISSQNtot','vPIS' ,'')) ,0) ;
//      Total.ISSQNTot.vCOFINS := StringToFloatDef( INIRec.ReadString('Total','ValorCONFINSISS',INIRec.ReadString('ISSQNtot','vCOFINS','')) ,0) ;
//      Total.ISSQNtot.dCompet     := StringToDateTime(INIRec.ReadString('ISSQNtot','dCompet','0'));
//      Total.ISSQNtot.vDeducao    := StringToFloatDef( INIRec.ReadString('ISSQNtot','vDeducao'   ,'') ,0) ;
//      Total.ISSQNtot.vOutro      := StringToFloatDef( INIRec.ReadString('ISSQNtot','vOutro'   ,'') ,0) ;
//      Total.ISSQNtot.vDescIncond := StringToFloatDef( INIRec.ReadString('ISSQNtot','vDescIncond'   ,'') ,0) ;
//      Total.ISSQNtot.vDescCond   := StringToFloatDef( INIRec.ReadString('ISSQNtot','vDescCond'   ,'') ,0) ;
//      Total.ISSQNtot.vISSRet     := StringToFloatDef( INIRec.ReadString('ISSQNtot','vISSRet'   ,'') ,0) ;
//      Total.ISSQNtot.cRegTrib    := StrToRegTribISSQN( OK,INIRec.ReadString('ISSQNtot','cRegTrib','1')) ;
//
//      Total.retTrib.vRetPIS    := StringToFloatDef( INIRec.ReadString('retTrib','vRetPIS'   ,'') ,0) ;
//      Total.retTrib.vRetCOFINS := StringToFloatDef( INIRec.ReadString('retTrib','vRetCOFINS','') ,0) ;
//      Total.retTrib.vRetCSLL   := StringToFloatDef( INIRec.ReadString('retTrib','vRetCSLL'  ,'') ,0) ;
//      Total.retTrib.vBCIRRF    := StringToFloatDef( INIRec.ReadString('retTrib','vBCIRRF'   ,'') ,0) ;
//      Total.retTrib.vIRRF      := StringToFloatDef( INIRec.ReadString('retTrib','vIRRF'     ,'') ,0) ;
//      Total.retTrib.vBCRetPrev := StringToFloatDef( INIRec.ReadString('retTrib','vBCRetPrev','') ,0) ;
//      Total.retTrib.vRetPrev   := StringToFloatDef( INIRec.ReadString('retTrib','vRetPrev'  ,'') ,0) ;
//
//      Transp.modFrete := StrTomodFrete(OK, INIRec.ReadString('Transportador','FretePorConta',INIRec.ReadString('Transportador','modFrete','0')));
//      Transp.Transporta.CNPJCPF  := INIRec.ReadString('Transportador','CNPJCPF'  ,'');
//      Transp.Transporta.xNome    := INIRec.ReadString('Transportador','NomeRazao',INIRec.ReadString('Transportador','xNome',''));
//      Transp.Transporta.IE       := INIRec.ReadString('Transportador','IE'       ,'');
//      Transp.Transporta.xEnder   := INIRec.ReadString('Transportador','Endereco' ,INIRec.ReadString('Transportador','xEnder',''));
//      Transp.Transporta.xMun     := INIRec.ReadString('Transportador','Cidade'   ,INIRec.ReadString('Transportador','xMun',''));
//      Transp.Transporta.UF       := INIRec.ReadString('Transportador','UF'       ,'');
//
//      Transp.retTransp.vServ    := StringToFloatDef( INIRec.ReadString('Transportador','ValorServico',INIRec.ReadString('Transportador','vServ'   ,'')) ,0) ;
//      Transp.retTransp.vBCRet   := StringToFloatDef( INIRec.ReadString('Transportador','ValorBase'   ,INIRec.ReadString('Transportador','vBCRet'  ,'')) ,0) ;
//      Transp.retTransp.pICMSRet := StringToFloatDef( INIRec.ReadString('Transportador','Aliquota'    ,INIRec.ReadString('Transportador','pICMSRet','')) ,0) ;
//      Transp.retTransp.vICMSRet := StringToFloatDef( INIRec.ReadString('Transportador','Valor'       ,INIRec.ReadString('Transportador','vICMSRet','')) ,0) ;
//      Transp.retTransp.CFOP     := INIRec.ReadString('Transportador','CFOP'     ,'');
//      Transp.retTransp.cMunFG   := INIRec.ReadInteger('Transportador','CidadeCod',INIRec.ReadInteger('Transportador','cMunFG',0));
//
//      Transp.veicTransp.placa := INIRec.ReadString('Transportador','Placa'  ,'');
//      Transp.veicTransp.UF    := INIRec.ReadString('Transportador','UFPlaca','');
//      Transp.veicTransp.RNTC  := INIRec.ReadString('Transportador','RNTC'   ,'');
//
//      Transp.vagao := INIRec.ReadString( 'Transportador','vagao','') ;
//      Transp.balsa := INIRec.ReadString( 'Transportador','balsa','') ;
//
//      Cobr.Fat.nFat  := INIRec.ReadString( 'Fatura','Numero',INIRec.ReadString( 'Fatura','nFat',''));
//      Cobr.Fat.vOrig := StringToFloatDef( INIRec.ReadString('Fatura','ValorOriginal',INIRec.ReadString('Fatura','vOrig','')) ,0) ;
//      Cobr.Fat.vDesc := StringToFloatDef( INIRec.ReadString('Fatura','ValorDesconto',INIRec.ReadString('Fatura','vDesc','')) ,0) ;
//      Cobr.Fat.vLiq  := StringToFloatDef( INIRec.ReadString('Fatura','ValorLiquido' ,INIRec.ReadString('Fatura','vLiq' ,'')) ,0) ;
//
//      I := 1 ;
//      while true do
//       begin
//         sSecao    := 'Duplicata'+IntToStrZero(I,3) ;
//         sNumDup   := INIRec.ReadString(sSecao,'Numero',INIRec.ReadString(sSecao,'nDup','FIM')) ;
//         if (sNumDup = 'FIM') or (Length(sNumDup) <= 0) then
//            break ;
//
//         with Cobr.Dup.Add do
//          begin
//            nDup  := sNumDup;
//            dVenc := StringToDateTime(INIRec.ReadString( sSecao,'DataVencimento',INIRec.ReadString( sSecao,'dVenc','0')));
//            vDup  := StringToFloatDef( INIRec.ReadString(sSecao,'Valor',INIRec.ReadString(sSecao,'vDup','')) ,0) ;
//          end;
//         Inc(I);
//       end;
//
//      I := 1 ;
//      while true do
//       begin
//         sSecao    := 'pag'+IntToStrZero(I,3) ;
//         sFim      := INIRec.ReadString(sSecao,'tpag','FIM');
//         if (sFim = 'FIM') or (Length(sFim) <= 0) then
//            break ;
//
//         with pag.Add do
//          begin
//            tPag  := StrToFormaPagamento(OK,sFim);
//            vPag  := StringToFloatDef( INIRec.ReadString(sSecao,'vPag','') ,0) ;
//
//            tpIntegra  := StrTotpIntegra(OK,INIRec.ReadString(sSecao,'tpIntegra',''));
//            CNPJ  := INIRec.ReadString(sSecao,'CNPJ','');
//            tBand := StrToBandeiraCartao(OK,INIRec.ReadString(sSecao,'tBand','99'));
//            cAut  := INIRec.ReadString(sSecao,'cAut','');
//          end;
//         Inc(I);
//       end;
//
//      InfAdic.infAdFisco :=  INIRec.ReadString( 'DadosAdicionais','Fisco'      ,INIRec.ReadString( 'DadosAdicionais','infAdFisco',''));
//      InfAdic.infCpl     :=  INIRec.ReadString( 'DadosAdicionais','Complemento',INIRec.ReadString( 'DadosAdicionais','infCpl'    ,''));
//
//      I := 1 ;
//      while true do
//       begin
//         sSecao     := 'InfAdic'+IntToStrZero(I,3) ;
//         sCampoAdic := INIRec.ReadString(sSecao,'Campo',INIRec.ReadString(sSecao,'xCampo','FIM')) ;
//         if (sCampoAdic = 'FIM') or (Length(sCampoAdic) <= 0) then
//            break ;
//
//         with InfAdic.obsCont.Add do
//          begin
//            xCampo := sCampoAdic;
//            xTexto := INIRec.ReadString( sSecao,'Texto',INIRec.ReadString( sSecao,'xTexto',''));
//          end;
//         Inc(I);
//       end;
//
//      I := 1 ;
//      while true do
//       begin
//         sSecao     := 'ObsFisco'+IntToStrZero(I,3) ;
//         sCampoAdic := INIRec.ReadString(sSecao,'Campo',INIRec.ReadString(sSecao,'xCampo','FIM')) ;
//         if (sCampoAdic = 'FIM') or (Length(sCampoAdic) <= 0) then
//            break ;
//
//         with InfAdic.obsFisco.Add do
//          begin
//            xCampo := sCampoAdic;
//            xTexto := INIRec.ReadString( sSecao,'Texto',INIRec.ReadString( sSecao,'xTexto',''));
//          end;
//         Inc(I);
//       end;
//
//      I := 1 ;
//      while true do
//       begin
//         sSecao     := 'procRef'+IntToStrZero(I,3) ;
//         sCampoAdic := INIRec.ReadString(sSecao,'nProc','FIM') ;
//         if (sCampoAdic = 'FIM') or (Length(sCampoAdic) <= 0) then
//            break ;
//
//         with InfAdic.procRef.Add do
//          begin
//            nProc := sCampoAdic;
//            indProc := StrToindProc(OK,INIRec.ReadString( sSecao,'indProc','0'));
//          end;
//         Inc(I);
//       end;
    end;
  finally

  end;

end;

end.
