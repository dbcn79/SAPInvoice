unit SRV;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.SvcMgr, Vcl.Dialogs,
  edFuncionesConexion, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.MSSQL, FireDAC.Phys.MSSQLDef,
  FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt, Data.DB,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, FireDAC.Comp.UI, FireDAC.Phys.ODBCBase, Vcl.ExtCtrls, System.IniFiles,
  Vcl.Forms, IPPeerClient, Data.Bind.Components, Data.Bind.ObjectScope, REST.Client, edFuncionesJournal;

const
  // Password base de datos
  PASSSWORD_SA    = 'NEWYORK1930';
  TIEMPO_UNA_HORA = 3600000;

type
  TfrmSRV_SAPInvoice = class(TService)
    tmrTemporizador: TTimer;
    ConexionGesden: TFDConnection;
    FDTransaction: TFDTransaction;
    FDPhysMSSQLDriverLink: TFDPhysMSSQLDriverLink;
    FDGUIxWaitCursor: TFDGUIxWaitCursor;
    qrySEL: TFDQuery;
    procedure ServiceContinue(Sender: TService; var Continued: Boolean);
    procedure ServiceDestroy(Sender: TObject);
    procedure ServiceExecute(Sender: TService);
    procedure ServicePause(Sender: TService; var Paused: Boolean);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure tmrTemporizadorTimer(Sender: TObject);
  private
    { Private declarations }
    ServicioCfg: TServicioCfg;
    journalEntry: TJournalEntry;
    Reconectar: Boolean;

    {$region 'FICHERO INI'}
    procedure ParametrizacionServicio;
    function ReconectarConBaseDatos: Boolean;
    procedure MarcarReconexion(AValue: Boolean);
    {$endregion}

    {$region 'CONEXION BASE DE DATOS'}
    function ConectarBaseDatosGesden: Boolean;
    procedure DesconectarBaseDatosGesden;
    {$endregion}
  public
    function GetServiceController: TServiceController; override;
    { Public declarations }
  end;

var
  frmSRV_SAPInvoice: TfrmSRV_SAPInvoice;

implementation

{$R *.dfm}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  frmSRV_SAPInvoice.Controller(CtrlCode);
end;

function TfrmSRV_SAPInvoice.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TfrmSRV_SAPInvoice.ServiceContinue(Sender: TService; var Continued: Boolean);
begin
  ServicioCfg.GrabarLOG('Servicio srvSAPInvoice reanudado');
  // Volvemos a leer la configuración
  ParametrizacionServicio;
  // Volvemos a poner en marcha el timer
  tmrTemporizador.Enabled := True;
end;

procedure TfrmSRV_SAPInvoice.ServiceDestroy(Sender: TObject);
begin
  ServicioCfg.Free;
end;

procedure TfrmSRV_SAPInvoice.ServiceExecute(Sender: TService);
begin
  while not Terminated do
    ServiceThread.ProcessRequests(True);
end;

procedure TfrmSRV_SAPInvoice.ServicePause(Sender: TService; var Paused: Boolean);
begin
  ServicioCfg.GrabarLOG('Servicio srvSAPInvoice pausado');
  Paused := True;
  // Pausamos el timer
  tmrTemporizador.Enabled := False;
end;

procedure TfrmSRV_SAPInvoice.ServiceStart(Sender: TService; var Started: Boolean);
begin
  if ServicioCfg = nil then begin
    ServicioCfg := TServicioCfg.Create;
  end;

  ServicioCfg.GrabarLOG('Servicio srvSAPInvoice iniciado');
  ParametrizacionServicio;
  tmrTemporizador.Interval := TIEMPO_UNA_HORA;
  tmrTemporizador.Enabled := True;
end;

procedure TfrmSRV_SAPInvoice.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
  ServicioCfg.GrabarLOG('Servicio svrSAPInvoice parado');
  if Assigned(ServicioCfg) then begin
    ServicioCfg.Free;
  end;
  Stopped := True;
end;

{$region ' FICHERO INI'}
procedure TfrmSRV_SAPInvoice.ParametrizacionServicio;
var
  IniFile: TIniFile;
begin
  ServicioCfg.FicheroIni := ExtractFilePath(Application.ExeName) + 'SAPInvoice.config';
  IniFile := TIniFile.Create(ServicioCfg.FicheroIni);
  try
    // Recogemos los valores del Config
    ServicioCfg.servidor      := IniFile.ReadString('SITE', 'Server', '');
    ServicioCfg.baseDatos     := IniFile.ReadString('SITE', 'BBDD', '');
    ServicioCfg.usuario       := IniFile.ReadString('SITE', 'Login', '');
    ServicioCfg.password      := IniFile.ReadString('SITE', 'Password', '');
    ServicioCfg.baseurl       := IniFile.ReadString('SAP', 'URL', '');
    ServicioCfg.baseUrlLogin  := IniFile.ReadString('SAP', 'URL_LOGIN', '');
    ServicioCfg.userName      := IniFile.ReadString('SAP', 'USERNAME', '');
    ServicioCfg.passwordLogin := IniFile.ReadString('SAP', 'PASSWORD', '');
    ServicioCfg.companyDB     := IniFile.ReadString('SAP', 'COMPANYDB', '');
    ServicioCfg.horaConsultar := IniFile.ReadInteger('SAP', 'HORA_CONSULTAR', 22);
    ServicioCfg.host          := IniFile.ReadString('EMAIL', 'HOST', '');
    ServicioCfg.userEmail     := IniFile.ReadString('EMAIL', 'USER', '');
    ServicioCfg.passwordEmail := IniFile.ReadString('EMAIL', 'PASSWORD', '');
    ServicioCfg.puerto        := IniFile.ReadInteger('EMAIL', 'PORT', 0);
    ServicioCfg.destinatarios := IniFile.ReadString('EMAIL', 'DESTINATARIOS', '');

    if ServicioCfg.password = '' then
      ServicioCfg.password := PASSSWORD_SA;

    ServicioCfg.conexionBaseDatos := ConexionGesden;
    ServicioCfg.transaccionGlobal := FDTransaction;
  finally
    IniFile.Free;
  end;

  ServicioCfg.GrabarLOG('');
  ServicioCfg.GrabarLOG('  ** Parámetros fichero SAPInvoice.config **');
  ServicioCfg.GrabarLOG('        Servidor: ' + ServicioCfg.servidor);
  ServicioCfg.GrabarLOG('        BaseDatos: ' + ServicioCfg.baseDatos);
  ServicioCfg.GrabarLOG('  ** Fin Parámetros fichero SAPInvoice.config **');
end;

function TfrmSRV_SAPInvoice.ReconectarConBaseDatos: Boolean;
var
  IniFile: TIniFile;
begin
  ServicioCfg.FicheroIni := ExtractFilePath(Application.ExeName) + 'SAPInvoice.config';
  IniFile := TIniFile.Create(ServicioCfg.FicheroIni);
  try
    Result := (IniFile.ReadInteger('SITE', 'Reconectar', 0) = 1);
  finally
    IniFile.Free;
  end;
end;

procedure TfrmSRV_SAPInvoice.MarcarReconexion(AValue: Boolean);
var
  IniFile: TIniFile;
begin
  ServicioCfg.FicheroIni := ExtractFilePath(Application.ExeName) + 'SAPInvoice.config';
  IniFile := TIniFile.Create(ServicioCfg.FicheroIni);
  try
    IniFile.WriteInteger('SITE', 'Reconectar', Ord(AValue));
  finally
    IniFile.Free;
  end;
end;
{$endregion}

{$region 'CONEXION BASE DE DATOS'}
function TfrmSRV_SAPInvoice.ConectarBaseDatosGesden: Boolean;
begin
  Result := False;

  ConexionGesden.Connected:= False;
  ConexionGesden.Params.Clear;
  ConexionGesden.Params.Add('DriverId=GELITE');
  ConexionGesden.Params.Add('Server=' + ServicioCfg.servidor);
  ConexionGesden.Params.Add('Database=' + ServicioCfg.baseDatos);
  ConexionGesden.Params.Add('OSAuthent=No');
  ConexionGesden.Params.Add('User_Name=' + ServicioCfg.usuario);
  ConexionGesden.Params.Add('Password=' + ServicioCfg.password);
  ConexionGesden.Params.Add('MetaDefSchema=dbo');
  ConexionGesden.Params.Add('MetaDefCatalog=' + ServicioCfg.baseDatos);
  try
    ConexionGesden.Connected := True;
    ServicioCfg.GrabarLOG('Conectado: ' + ServicioCfg.baseDatos + ' en ' + ServicioCfg.servidor);
    Result := True;
  except
    on E: Exception do
      ServicioCfg.GrabarLOG('ERROR: No se pudo establecer conexión con la base de datos ' + ServicioCfg.baseDatos + ' (' + E.Message + ')');
  end;
end;

procedure TfrmSRV_SAPInvoice.DesconectarBaseDatosGesden;
begin
  if qrySEL.Active then qrySEL.Close;
  if ConexionGesden.Connected then begin
    ServicioCfg.GrabarLOG('Desconectado: ' + ServicioCfg.baseDatos + ' en ' + ServicioCfg.servidor);
    ConexionGesden.Connected := False;
  end;
end;
{$endregion'}

procedure TfrmSRV_SAPInvoice.tmrTemporizadorTimer(Sender: TObject);
var hora: Integer;
begin
  tmrTemporizador.Enabled := False;
  // Averiguamos la hora actual
  hora := FormatDateTime('hh', Now).ToInteger;
  if (hora = ServicioCfg.horaConsultar) or (ReconectarConBaseDatos) then begin
    journalEntry := TJournalEntry.Create;
    try
      if ConectarBaseDatosGesden then begin
        try
          ServicioCfg.createJsonObjects;
          ServicioCfg.createRESTObjects;
          ServicioCfg.inicializaValores;
          try
            // Obtenemos el id de la session
            ServicioCfg.dameLogin;
            // La fecha tiene que ser el día anterior por el tema de la diferencia horaria, en el servidor local está el dío actual pero
            // al restar la diferencia horaria el día es el día anterior
            ServicioCfg.fecha := Date;
            // Creamos el directorio Generados, que es donde dejaremos los ficheros JSON generados de cobros y facturas
            ServicioCfg.directorio := IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName)) + 'Generados\';
            if not DirectoryExists(ServicioCfg.directorio) then
              ForceDirectories(ServicioCfg.directorio);
            // Pasamos el object de conexión al TJournalEntry
            journalEntry.srvCfg := ServicioCfg;
            // Hacemos el posible cambio de Bancos.CodBIC a Bancos.SCCodSCta
            journalEntry.CambiarCuentaSAP;
            // Exportamos las facturas
            ServicioCfg.GrabarLOG('*** Exportando facturas ***');
            journalEntry.ExportarExcelFacturas;
            // Exportamos los cobros
            ServicioCfg.GrabarLOG('*** Exportando cobros ***');
            journalEntry.ExportarExcelCobrosPorFormaPago;
            // Si se ha conectado se coloca como a False la reconexión
            MarcarReconexion(false)
          except
            on E: Exception do begin
              ServicioCfg.GrabarLOG('ERROR: ' + E.Message);
              MarcarReconexion(true);
            end;
          end;
        finally
          DesconectarBaseDatosGesden;
          tmrTemporizador.Enabled := True;
          ServicioCfg.destroyJsonObjects;
          ServicioCfg.destroyRESTObjects;
        end;
      end else begin
        tmrTemporizador.Enabled := True;
        MarcarReconexion(true);
      end;
    finally
      journalEntry.Free;
    end;
  end else begin
    tmrTemporizador.Enabled := True;
  end;
end;

end.
