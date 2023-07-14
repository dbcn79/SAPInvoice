object frmSRV_SAPInvoice: TfrmSRV_SAPInvoice
  OnDestroy = ServiceDestroy
  DisplayName = 'srvSAPInvoice'
  OnContinue = ServiceContinue
  OnExecute = ServiceExecute
  OnPause = ServicePause
  OnStart = ServiceStart
  OnStop = ServiceStop
  Height = 241
  Width = 285
  PixelsPerInch = 96
  object tmrTemporizador: TTimer
    Enabled = False
    Interval = 1800
    OnTimer = tmrTemporizadorTimer
    Left = 176
    Top = 128
  end
  object ConexionGesden: TFDConnection
    ConnectionName = 'GELITE'
    Params.Strings = (
      'DriverID=MSSQL')
    Transaction = FDTransaction
    UpdateTransaction = FDTransaction
    Left = 56
    Top = 16
  end
  object FDTransaction: TFDTransaction
    Connection = ConexionGesden
    Left = 56
    Top = 128
  end
  object FDPhysMSSQLDriverLink: TFDPhysMSSQLDriverLink
    DriverID = 'GELITE'
    Left = 56
    Top = 72
  end
  object FDGUIxWaitCursor: TFDGUIxWaitCursor
    Provider = 'Forms'
    Left = 56
    Top = 184
  end
  object qrySEL: TFDQuery
    Connection = ConexionGesden
    Left = 232
    Top = 16
  end
end
