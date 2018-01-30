object frmPrincipal: TfrmPrincipal
  Left = 271
  Top = 114
  ActiveControl = ButtonStart
  Caption = 'Gerenciador de NFCe'
  ClientHeight = 158
  ClientWidth = 368
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 299
    Top = 8
    Width = 26
    Height = 13
    Caption = 'Porta'
  end
  object SpeedButton1: TSpeedButton
    Left = 186
    Top = 63
    Width = 107
    Height = 22
    Caption = 'Certificado'
    OnClick = SpeedButton1Click
  end
  object btnEmitente: TSpeedButton
    Left = 186
    Top = 91
    Width = 107
    Height = 22
    Caption = 'Emitente'
    OnClick = btnEmitenteClick
  end
  object SpeedButton2: TSpeedButton
    Left = 186
    Top = 119
    Width = 107
    Height = 22
    Caption = 'Enviar NFCe Teste'
    OnClick = SpeedButton2Click
  end
  object ButtonStart: TButton
    Left = 24
    Top = 24
    Width = 75
    Height = 25
    Caption = 'Start'
    TabOrder = 0
    OnClick = ButtonStartClick
  end
  object ButtonStop: TButton
    Left = 105
    Top = 24
    Width = 75
    Height = 25
    Caption = 'Stop'
    TabOrder = 1
    OnClick = ButtonStopClick
  end
  object EditPort: TEdit
    Left = 299
    Top = 24
    Width = 53
    Height = 21
    TabOrder = 2
    Text = '211'
  end
  object ButtonOpenBrowser: TButton
    Left = 186
    Top = 24
    Width = 107
    Height = 25
    Caption = 'Open Browser'
    TabOrder = 3
    OnClick = ButtonOpenBrowserClick
  end
  object GroupBox1: TGroupBox
    Left = 24
    Top = 55
    Width = 156
    Height = 86
    Caption = 'Dados NFce'
    TabOrder = 4
    object Label2: TLabel
      Left = 14
      Top = 15
      Width = 59
      Height = 13
      Caption = 'NumeroNfce'
    end
    object Label3: TLabel
      Left = 79
      Top = 15
      Width = 24
      Height = 13
      Caption = 'Serie'
    end
    object Edit1: TEdit
      Left = 79
      Top = 34
      Width = 26
      Height = 21
      Alignment = taCenter
      TabOrder = 0
      Text = '1'
    end
    object Edit2: TEdit
      Left = 14
      Top = 34
      Width = 59
      Height = 21
      TabOrder = 1
      Text = '1503'
    end
  end
  object ApplicationEvents1: TApplicationEvents
    OnIdle = ApplicationEvents1Idle
    Left = 48
    Top = 8
  end
end
