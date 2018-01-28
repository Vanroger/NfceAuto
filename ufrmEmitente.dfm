object frmEmitente: TfrmEmitente
  Left = 0
  Top = 0
  Caption = 'Cadastro do Emitente'
  ClientHeight = 344
  ClientWidth = 465
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 8
    Width = 25
    Height = 13
    Caption = 'CNPJ'
  end
  object Label2: TLabel
    Left = 143
    Top = 8
    Width = 60
    Height = 13
    Caption = 'Raz'#227'o Social'
  end
  object Label3: TLabel
    Left = 16
    Top = 56
    Width = 71
    Height = 13
    Caption = 'Nome Fantasia'
  end
  object Label4: TLabel
    Left = 255
    Top = 56
    Width = 87
    Height = 13
    Caption = 'Inscri'#231'ao Estadual'
  end
  object Label6: TLabel
    Left = 16
    Top = 104
    Width = 89
    Height = 13
    Caption = 'Inscri'#231#227'o Municipal'
  end
  object Label7: TLabel
    Left = 143
    Top = 104
    Width = 27
    Height = 13
    Caption = 'CNAE'
  end
  object Label8: TLabel
    Left = 270
    Top = 104
    Width = 120
    Height = 13
    Caption = 'C'#243'digo Regime Tribut'#225'rio'
  end
  object Label9: TLabel
    Left = 16
    Top = 152
    Width = 45
    Height = 13
    Caption = 'Endereco'
  end
  object Label10: TLabel
    Left = 382
    Top = 152
    Width = 37
    Height = 13
    Caption = 'Numero'
  end
  object Label11: TLabel
    Left = 16
    Top = 202
    Width = 65
    Height = 13
    Caption = 'Complemento'
  end
  object Label12: TLabel
    Left = 270
    Top = 202
    Width = 28
    Height = 13
    Caption = 'Bairro'
  end
  object Label13: TLabel
    Left = 16
    Top = 250
    Width = 79
    Height = 13
    Caption = 'C'#243'digo Municipio'
  end
  object Label14: TLabel
    Left = 111
    Top = 250
    Width = 43
    Height = 13
    Caption = 'Municipio'
  end
  object Label15: TLabel
    Left = 238
    Top = 250
    Width = 13
    Height = 13
    Caption = 'UF'
  end
  object Label16: TLabel
    Left = 272
    Top = 250
    Width = 19
    Height = 13
    Caption = 'CEP'
  end
  object Label17: TLabel
    Left = 359
    Top = 250
    Width = 24
    Height = 13
    Caption = 'Fone'
  end
  object SpeedButton1: TSpeedButton
    Left = 336
    Top = 312
    Width = 103
    Height = 25
    Caption = 'Confirmar'
    OnClick = SpeedButton1Click
  end
  object SpeedButton2: TSpeedButton
    Left = 227
    Top = 312
    Width = 103
    Height = 25
    Caption = 'Cancelar'
    OnClick = SpeedButton2Click
  end
  object edtCNPJ: TEdit
    Left = 16
    Top = 27
    Width = 121
    Height = 21
    TabOrder = 0
  end
  object edtNome: TEdit
    Left = 143
    Top = 27
    Width = 296
    Height = 21
    TabOrder = 1
  end
  object edtFantasia: TEdit
    Left = 16
    Top = 75
    Width = 233
    Height = 21
    TabOrder = 2
  end
  object edtIE: TEdit
    Left = 255
    Top = 75
    Width = 184
    Height = 21
    TabOrder = 3
  end
  object edtInscMun: TEdit
    Left = 16
    Top = 123
    Width = 121
    Height = 21
    TabOrder = 4
  end
  object edtCNAE: TEdit
    Left = 143
    Top = 123
    Width = 121
    Height = 21
    TabOrder = 5
  end
  object edtLgr: TEdit
    Left = 16
    Top = 171
    Width = 360
    Height = 21
    TabOrder = 6
  end
  object edtNumero: TEdit
    Left = 382
    Top = 171
    Width = 57
    Height = 21
    TabOrder = 7
  end
  object edtComplemento: TEdit
    Left = 16
    Top = 221
    Width = 248
    Height = 21
    TabOrder = 8
  end
  object edtBairro: TEdit
    Left = 270
    Top = 221
    Width = 169
    Height = 21
    TabOrder = 9
  end
  object edtCodMun: TEdit
    Left = 16
    Top = 269
    Width = 89
    Height = 21
    NumbersOnly = True
    TabOrder = 10
  end
  object edtCidade: TEdit
    Left = 111
    Top = 269
    Width = 121
    Height = 21
    TabOrder = 11
  end
  object edtUF: TEdit
    Left = 238
    Top = 269
    Width = 28
    Height = 21
    TabOrder = 12
  end
  object edtCep: TEdit
    Left = 272
    Top = 269
    Width = 81
    Height = 21
    NumbersOnly = True
    TabOrder = 13
  end
  object edtFone: TEdit
    Left = 359
    Top = 269
    Width = 80
    Height = 21
    TabOrder = 14
  end
  object cbxCRT: TComboBox
    Left = 270
    Top = 123
    Width = 169
    Height = 21
    TabOrder = 15
    Text = '1 - Simples Nacional'
    Items.Strings = (
      '1 - Simples Nacional'
      '2 - Simples Nacional Excesso de Sublimite de Receita Bruta'
      '3 - Regime Normal')
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
    Left = 408
    Top = 24
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
