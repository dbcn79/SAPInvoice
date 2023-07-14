unit FRM;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms,
  Vcl.Dialogs, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.MSSQL, FireDAC.Phys.MSSQLDef, FireDAC.VCLUI.Wait, IPPeerClient,
  FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt, Vcl.StdCtrls, Vcl.ExtCtrls, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, FireDAC.Comp.UI, FireDAC.Phys.ODBCBase, edFuncionesConexion, edFuncionesJournal, System.IniFiles,
  Vcl.ComCtrls, System.DateUtils, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase,
  IdMessageClient, IdSMTPBase, IdSMTP, IdMessage, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL;

const
  // Password base de datos
  PASSSWORD_SA = 'NEWYORK1930';

type
  TfrmPrincipal = class(TForm)
    ConexionGesden: TFDConnection;
    FDTransaction: TFDTransaction;
    FDPhysMSSQLDriverLink: TFDPhysMSSQLDriverLink;
    FDGUIxWaitCursor: TFDGUIxWaitCursor;
    qrySEL: TFDQuery;
    Panel1: TPanel;
    btTEST: TButton;
    cbGenerar: TCheckBox;
    rgImportacion: TRadioGroup;
    Panel2: TPanel;
    GroupBox1: TGroupBox;
    lbFecha: TLabel;
    deFechaDesde: TDateTimePicker;
    Label1: TLabel;
    deFechaHasta: TDateTimePicker;
    gbOpciones: TGroupBox;
    cbFacturas: TCheckBox;
    cbCobros: TCheckBox;
    lbSesion: TLabel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    memLog: TMemo;
    Panel6: TPanel;
    btnSalir: TButton;
    procedure btTESTClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure cbFechaClick(Sender: TObject);
    procedure btnSalirClick(Sender: TObject);
  private
    { Private declarations }
    {$region 'FICHERO INI'}
    procedure ParametrizacionServicio;
    {$endregion}

    {$region 'CONEXION BASE DE DATOS'}
    function ConectarBaseDatosGesden: Boolean;
    procedure DesconectarBaseDatosGesden;
    {$endregion}
  public
    { Public declarations }
    ServicioCfg: TServicioCfg;
    journalEntry: TJournalEntry;
end;

var
  frmPrincipal: TfrmPrincipal;

implementation

{$R *.dfm}

procedure TfrmPrincipal.FormCreate(Sender: TObject);
begin
  deFechaDesde.Date := Date -1;
  deFechaHasta.Date := Date;
end;

procedure TfrmPrincipal.cbFechaClick(Sender: TObject);
begin
end;

{$region ' FICHERO INI'}
procedure TfrmPrincipal.ParametrizacionServicio;
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
    ServicioCfg.host          := IniFile.ReadString('EMAIL', 'HOST', '');
    ServicioCfg.userEmail     := IniFile.ReadString('EMAIL', 'USER', '');
    ServicioCfg.passwordEmail := IniFile.ReadString('EMAIL', 'PASSWORD', '');
    ServicioCfg.puerto        := IniFile.ReadInteger('EMAIL', 'PORT', 0);
    ServicioCfg.destinatarios := IniFile.ReadString('EMAIL', 'DESTINATARIOS', '');

    if ServicioCfg.password = '' then
      ServicioCfg.password      := PASSSWORD_SA;

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
{$endregion}

{$region 'CONEXION BASE DE DATOS'}

function TfrmPrincipal.ConectarBaseDatosGesden: Boolean;
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

procedure TfrmPrincipal.DesconectarBaseDatosGesden;
begin
  if qrySEL.Active then qrySEL.Close;
  if ConexionGesden.Connected then begin
    ServicioCfg.GrabarLOG('Desconectado: ' + ServicioCfg.baseDatos + ' en ' + ServicioCfg.servidor);
    ConexionGesden.Connected := False;
  end;
end;
{$endregion'}

procedure TfrmPrincipal.btnSalirClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmPrincipal.btTESTClick(Sender: TObject);
var dias, i: Integer; fechaActual: TDateTime; bicCambiado: Boolean;
begin
  ServicioCfg := TServicioCfg.Create;
  journalEntry := TJournalEntry.Create;
  try
    ParametrizacionServicio;
    memLog.Lines.Add('** Servidor: ' + ServicioCfg.servidor + ' **');
    memLog.Lines.Add('** Base de Datos: ' + ServicioCfg.baseDatos + ' **');
    memLog.Lines.Add('');
    memLog.Lines.Add('************* INICIO PROCESO *************');

    if ConectarBaseDatosGesden then begin

      try
        ServicioCfg.createJsonObjects;
        ServicioCfg.createRESTObjects;
        ServicioCfg.inicializaValores;
        try
          // Obtenemos el id de la session
          ServicioCfg.dameLogin;
          lbSesion.Caption := ServicioCfg.sessionId;
          Application.ProcessMessages;
          dias := Trunc(deFechaHasta.Date) - Trunc(deFechaDesde.Date);
          fechaActual := deFechaDesde.Date;
          ServicioCfg.directorio := IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName)) + 'Generados\';
          if not DirectoryExists(ServicioCfg.directorio) then
            ForceDirectories(ServicioCfg.directorio);
          bicCambiado := false;
          // Pasamos el object de conexión al TJournalEntry
          for i := 0 to dias do begin
            ServicioCfg.fecha := deFechaDesde.Date + i;
            journalEntry.srvCfg := ServicioCfg;

            if not bicCambiado then begin
              // Hacemos el posible cambio de Bancos.CodBIC a Bancos.SCCodSCta
              journalEntry.CambiarCuentaSAP;
              bicCambiado := true;
            end;

            if cbFacturas.Checked then begin
              memLog.Lines.Add('** Exportando facturas día ' + FormatDateTime('dd/mm/yyyy', ServicioCfg.fecha) + ' **');
              ServicioCfg.GrabarLOG('** Exportando facturas día ' + FormatDateTime('dd/mm/yyyy', ServicioCfg.fecha) + ' **');
              journalEntry.ExportarExcelFacturas;
            end;
            if cbCobros.Checked then begin
              memLog.Lines.Add('** Exportando cobros día ' + FormatDateTime('dd/mm/yyyy', ServicioCfg.fecha) + ' **');
              ServicioCfg.GrabarLOG('** Exportando cobros día ' + FormatDateTime('dd/mm/yyyy', ServicioCfg.fecha) + ' **');
               case rgImportacion.ItemIndex of
                 0: journalEntry.ExportarExcelCobrosPorFormaPago;
                 1: journalEntry.ExportarExcelCobrosPorBanco;
               end;
            end;
            Application.ProcessMessages;
          end;
        except
          on E: Exception do
            ServicioCfg.GrabarLOG('ERROR: ' + E.Message);
        end;
      finally
        DesconectarBaseDatosGesden;
        memLog.Lines.Add('************* FIN PROCESO *************');
      end;
    end else begin
      memLog.Lines.Add('No se pudo conectar con la base de datos!');
    end;
  finally
    journalEntry.Free;
    ServicioCfg.destroyJsonObjects;
    ServicioCfg.destroyRESTObjects;
    ServicioCfg.Free;
  end;
end;

end.
