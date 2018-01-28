object frmGerenciadorNFCe: TfrmGerenciadorNFCe
  Left = 0
  Top = 0
  Caption = 'Gerenciador NFCe'
  ClientHeight = 201
  ClientWidth = 447
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object ACBrNFe1: TACBrNFe
    Configuracoes.Geral.SSLLib = libNone
    Configuracoes.Geral.SSLCryptLib = cryNone
    Configuracoes.Geral.SSLHttpLib = httpNone
    Configuracoes.Geral.SSLXmlSignLib = xsNone
    Configuracoes.Geral.FormatoAlerta = 'TAG:%TAGNIVEL% ID:%ID%/%TAG%(%DESCRICAO%) - %MSG%.'
    Configuracoes.Arquivos.OrdenacaoPath = <>
    Configuracoes.WebServices.UF = 'SP'
    Configuracoes.WebServices.AguardarConsultaRet = 0
    Configuracoes.WebServices.QuebradeLinha = '|'
    Left = 24
    Top = 8
  end
  object cdsCertificado: TClientDataSet
    PersistDataPacket.Data = {
      6C0000009619E0BD0100000018000000030000000000030000006C000B436572
      746966696361646F010049000000010005574944544802000200640003637363
      0100490000000100055749445448020002006400056964637363010049000000
      01000557494454480200020001000000}
    Active = True
    Aggregates = <>
    Params = <>
    Left = 376
    Top = 24
    object cdsCertificadoCertificado: TStringField
      FieldName = 'Certificado'
      Size = 100
    end
    object cdsCertificadocsc: TStringField
      FieldName = 'csc'
      Size = 100
    end
    object cdsCertificadoidcsc: TStringField
      FieldName = 'idcsc'
      Size = 1
    end
  end
  object cdsEmitente: TClientDataSet
    PersistDataPacket.Data = {
      B30100009619E0BD010000001800000010000000000003000000B30104636E70
      6A0100490000000100055749445448020002001200046E6F6D65010049000000
      01000557494454480200020064000866616E7461736961010049000000010005
      57494454480200020064000C496E7363457374616475616C0100490000000100
      05574944544802000200140007496E73634D756E010049000000010005574944
      544802000200140004434E41450100490000000100055749445448020002000A
      0008656E64657265636F0100490000000100055749445448020002006400066E
      756D65726F01004900000001000557494454480200020006000B436F6D706C65
      6D656E746F01004900000001000557494454480200020064000662616972726F
      010049000000010005574944544802000200640006636F644D756E0400010000
      000000096D756E69636970696F01004900000001000557494454480200020032
      0002756601004900000001000557494454480200020002000343455004000100
      0000000004464F4E450100490000000100055749445448020002001400034352
      5401004900000001000557494454480200020001000000}
    Active = True
    Aggregates = <>
    Params = <>
    Left = 376
    Top = 80
    object cdsEmitentecnpj: TStringField
      FieldName = 'cnpj'
      Size = 18
    end
    object cdsEmitentenome: TStringField
      FieldName = 'nome'
      Size = 100
    end
    object cdsEmitentefantasia: TStringField
      FieldName = 'fantasia'
      Size = 100
    end
    object cdsEmitenteInscEstadual: TStringField
      FieldName = 'InscEstadual'
    end
    object cdsEmitenteInscMun: TStringField
      FieldName = 'InscMun'
    end
    object cdsEmitenteCNAE: TStringField
      FieldName = 'CNAE'
      Size = 10
    end
    object cdsEmitenteendereco: TStringField
      FieldName = 'endereco'
      Size = 100
    end
    object cdsEmitentenumero: TStringField
      FieldName = 'numero'
      Size = 6
    end
    object cdsEmitenteComplemento: TStringField
      FieldName = 'Complemento'
      Size = 100
    end
    object cdsEmitentebairro: TStringField
      FieldName = 'bairro'
      Size = 100
    end
    object cdsEmitentecodMun: TIntegerField
      FieldName = 'codMun'
    end
    object cdsEmitentemunicipio: TStringField
      FieldName = 'municipio'
      Size = 50
    end
    object cdsEmitenteuf: TStringField
      FieldName = 'uf'
      Size = 2
    end
    object cdsEmitenteCEP: TIntegerField
      FieldName = 'CEP'
    end
    object cdsEmitenteFONE: TStringField
      FieldName = 'FONE'
    end
    object cdsEmitenteCRT: TStringField
      FieldName = 'CRT'
      Size = 1
    end
  end
end
