object frmCertificado: TfrmCertificado
  Left = 0
  Top = 0
  BorderIcons = []
  Caption = 'Informa'#231'oes Certificado'
  ClientHeight = 206
  ClientWidth = 349
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Certificado: TLabel
    Left = 16
    Top = 13
    Width = 52
    Height = 13
    Caption = 'Certificado'
  end
  object Label1: TLabel
    Left = 16
    Top = 61
    Width = 53
    Height = 13
    Caption = 'CSC/Token'
  end
  object Label2: TLabel
    Left = 16
    Top = 117
    Width = 71
    Height = 13
    Caption = 'IdCSC/idToken'
  end
  object SpeedButton1: TSpeedButton
    Left = 318
    Top = 32
    Width = 23
    Height = 22
    OnClick = SpeedButton1Click
  end
  object btnGravar: TSpeedButton
    Left = 127
    Top = 166
    Width = 105
    Height = 30
    Caption = 'Confirmar'
    OnClick = btnGravarClick
  end
  object SpeedButton3: TSpeedButton
    Left = 16
    Top = 166
    Width = 105
    Height = 30
    Caption = 'Cancelar'
    OnClick = SpeedButton3Click
  end
  object edtCertificado: TEdit
    Left = 16
    Top = 32
    Width = 296
    Height = 21
    TabOrder = 0
  end
  object edtCSC: TEdit
    Left = 16
    Top = 80
    Width = 325
    Height = 21
    TabOrder = 1
  end
  object edtID: TEdit
    Left = 16
    Top = 136
    Width = 71
    Height = 21
    Alignment = taCenter
    TabOrder = 2
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
    Left = 264
    Top = 112
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
end
