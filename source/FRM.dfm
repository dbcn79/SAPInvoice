object frmPrincipal: TfrmPrincipal
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'srvSAPInvoice_Payment TEST'
  ClientHeight = 440
  ClientWidth = 419
  Color = clBtnFace
  CustomTitleBar.CaptionAlignment = taCenter
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 419
    Height = 154
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object lbSesion: TLabel
      Left = 151
      Top = 118
      Width = 3
      Height = 13
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object btTEST: TButton
      Left = 151
      Top = 87
      Width = 127
      Height = 25
      Caption = 'Importar a SAP'
      TabOrder = 0
      OnClick = btTESTClick
    end
    object cbGenerar: TCheckBox
      Left = 284
      Top = 87
      Width = 89
      Height = 17
      Caption = 'Generar JSON'
      Checked = True
      State = cbChecked
      TabOrder = 1
      Visible = False
      OnClick = cbFechaClick
    end
    object rgImportacion: TRadioGroup
      Left = 8
      Top = 24
      Width = 137
      Height = 57
      Caption = ' Importar a SAP seg'#250'n '
      ItemIndex = 0
      Items.Strings = (
        'Por formas de pago'
        'Por banco')
      TabOrder = 2
    end
    object Panel2: TPanel
      Left = 0
      Top = 0
      Width = 419
      Height = 18
      Align = alTop
      BevelOuter = bvNone
      Caption = 'Filtros para importar a SAP'
      Color = clSilver
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 3
    end
    object GroupBox1: TGroupBox
      Left = 151
      Top = 24
      Width = 261
      Height = 57
      Caption = ' Filtro Fechas '
      TabOrder = 4
      object lbFecha: TLabel
        Left = 12
        Top = 22
        Width = 30
        Height = 13
        Caption = 'Desde'
      end
      object Label1: TLabel
        Left = 133
        Top = 22
        Width = 30
        Height = 13
        Caption = 'Desde'
      end
      object deFechaDesde: TDateTimePicker
        Left = 48
        Top = 18
        Width = 80
        Height = 21
        Date = 44927.000000000000000000
        Time = 0.670568680558062600
        TabOrder = 0
      end
      object deFechaHasta: TDateTimePicker
        Left = 169
        Top = 18
        Width = 80
        Height = 21
        Date = 44927.000000000000000000
        Time = 0.670568680558062600
        TabOrder = 1
      end
    end
    object gbOpciones: TGroupBox
      Left = 8
      Top = 87
      Width = 137
      Height = 57
      Caption = ' Importar '
      TabOrder = 5
      object cbFacturas: TCheckBox
        Left = 16
        Top = 18
        Width = 97
        Height = 17
        Caption = 'Facturas'
        Checked = True
        State = cbChecked
        TabOrder = 0
      end
      object cbCobros: TCheckBox
        Left = 16
        Top = 36
        Width = 97
        Height = 17
        Caption = 'Cobros'
        Checked = True
        State = cbChecked
        TabOrder = 1
      end
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 154
    Width = 419
    Height = 245
    Align = alClient
    BevelOuter = bvNone
    Caption = 'Panel3'
    TabOrder = 1
    object Panel4: TPanel
      Left = 0
      Top = 0
      Width = 419
      Height = 18
      Align = alTop
      BevelOuter = bvNone
      Caption = 'Registro de Log'
      Color = clSilver
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 0
    end
    object Panel5: TPanel
      Left = 0
      Top = 18
      Width = 419
      Height = 227
      Align = alClient
      BevelOuter = bvNone
      Caption = 'Panel5'
      TabOrder = 1
      object memLog: TMemo
        Left = 0
        Top = 0
        Width = 419
        Height = 227
        Align = alClient
        ScrollBars = ssVertical
        TabOrder = 0
      end
    end
  end
  object Panel6: TPanel
    Left = 0
    Top = 399
    Width = 419
    Height = 41
    Align = alBottom
    TabOrder = 2
    object btnSalir: TButton
      Left = 273
      Top = 8
      Width = 127
      Height = 25
      Caption = 'Salir'
      TabOrder = 0
      OnClick = btnSalirClick
    end
  end
  object ConexionGesden: TFDConnection
    ConnectionName = 'GELITE'
    Params.Strings = (
      'DriverID=MSSQL')
    Transaction = FDTransaction
    UpdateTransaction = FDTransaction
    Left = 72
    Top = 144
  end
  object FDTransaction: TFDTransaction
    Connection = ConexionGesden
    Left = 72
    Top = 256
  end
  object FDPhysMSSQLDriverLink: TFDPhysMSSQLDriverLink
    DriverID = 'GELITE'
    Left = 72
    Top = 200
  end
  object FDGUIxWaitCursor: TFDGUIxWaitCursor
    Provider = 'Forms'
    Left = 72
    Top = 312
  end
  object qrySEL: TFDQuery
    Connection = ConexionGesden
    Left = 112
    Top = 56
  end
end
