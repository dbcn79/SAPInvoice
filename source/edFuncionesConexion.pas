unit edFuncionesConexion;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.SvcMgr, Vcl.Dialogs,
  Vcl.ExtCtrls, System.IniFiles, Vcl.Forms, System.StrUtils, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error,
  FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.MSSQL,
  FireDAC.Phys.MSSQLDef, FireDAC.Comp.Client, Data.DB, FireDAC.Phys.ODBCBase, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.VCLUI.Wait, FireDAC.Comp.UI, REST.Client, REST.Authenticator.Simple, REST.Authenticator.Basic,
  System.NetEncoding, REST.Types, System.Generics.Collections, System.JSON, REST.Authenticator.OAuth, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase, IdSMTP, IdMessage, IdIOHandler, IdIOHandlerSocket,
  IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdAttachmentFile, IdGlobal;

const
  ARCHIVO_LOG_NAME = 'SAPInvoice.txt';
  JSON_ID_HTTPS    = 0;
  JSON_ID_SESSION  = 1;

type
  TServicioCfg = class(TObject)
  public
    ficheroIni: string;
    servidor: string;
    baseDatos: string;
    usuario: string;
    password: string;
    baseUrl: string;
    baseUrlLogin: string;
    userName: string;
    passwordLogin: string;
    companyDB: string;
    odataMetadata: string;
    sessionId: string;
    numAsiento: string;
    host: string;
    userEmail: string;
    passwordEmail: string;
    puerto: Integer;
    destinatarios: string;
    error: string;
    directorio: string;
    nombreFichero: string;
    tratamientos: string;
    horaConsultar: Integer;
    statusCode: Integer;
    envioOk: Boolean;
    fecha: TDateTime;
    generarJson: Boolean;
    srvLOG: TStringList;
    conexionBaseDatos: TFDConnection;
    transaccionGlobal: TFDTransaction;
    restClient: TRESTClient;
    restResponse: TRESTResponse;
    restRequest: TRESTRequest;
    jsonRequest: TJSONObject;
    Authenticator: TOAuth2Authenticator;

    constructor Create;
    destructor Destroy; override;

    procedure CreateRESTObjects;
    procedure DestroyRESTObjects;
    procedure InicializaValores;

    procedure createJsonObjects;
    procedure destroyJsonObjects;
    procedure rellenarJSONObject;

    procedure DameLogin;
    procedure montarMensajeError(var lista: TStringList);
    procedure GrabarLOG(parMensaje: string);
    procedure SendEmail(Subject: string; Body: TStrings);
  end;

implementation

//uses PrincipalF;

{ TServicioCfg }

constructor TServicioCfg.Create;
begin
  ficheroIni    := '';
  servidor      := '';
  baseDatos     := '';
  usuario       := '';
  password      := '';
  error         := '';
  baseUrl       := '';
  baseUrlLogin  := '';
  userName      := '';
  passwordLogin := '';
  companyDB     := '';
  odataMetadata := '';
  sessionId     := '';
  numAsiento    := '';
  host          := '';
  userEmail     := '';
  passwordEmail := '';
  puerto        := 0;
  destinatarios := '';
  directorio    := '';
  nombreFichero := '';
  tratamientos  := '';
  envioOk       := false;
  statusCode    := 0;
  fecha         := Date;
  horaConsultar := 22;
  generarJson   := true;
  srvLOG        := TStringList.Create;
  restClient    := nil;
  restResponse  := nil;
  restRequest   := nil;
end;

destructor TServicioCfg.Destroy;
begin
  srvLOG.Free;
  inherited;
end;

procedure TServicioCfg.CreateRESTObjects;
begin
  restClient    := TrestClient.Create(nil);
  restResponse  := TrestResponse.Create(nil);
  restRequest   := TrestRequest.Create(nil);
  Authenticator := TOAuth2Authenticator.Create(nil);
end;

procedure TServicioCfg.DestroyRESTObjects;
begin
  Authenticator.Free;
  restClient.Free;
  restResponse.Free;
  restRequest.Free;
end;

procedure TServicioCfg.InicializaValores;
begin
  restClient.Accept := 'application/json';
  restClient.ContentType := 'application/json';
  restClient.UserAgent := 'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.13 (KHTML, like';
  restClient.HandleRedirects := True;
  restRequest.Accept := 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8';
  restRequest.Method := rmPOST;
  restRequest.Client := restClient;
  restRequest.Response := restResponse;
  restResponse.ContentType := 'txt/json';
end;

procedure TServicioCfg.createJsonObjects;
begin
  jsonRequest   := TJSONObject.Create;
end;

procedure TServicioCfg.destroyJsonObjects;
begin
  FreeAndNil(jsonRequest);
end;

procedure TServicioCfg.rellenarJSONObject;
begin
  jsonRequest.AddPair('UserName', userName);
  jsonRequest.AddPair('Password', passwordLogin);
  jsonRequest.AddPair('CompanyDB', companyDB);
end;

procedure TServicioCfg.dameLogin;
var request: string; body: TStringList;
begin
  body := TStringList.Create;
  try
    GrabarLOG('urlLogin :' +  baseUrlLogin);
    restClient.BaseURL := baseUrlLogin;
    restRequest.Method := rmPOST;
    try
      rellenarJSONObject;
      request := jsonRequest.ToString;
      request := request.Replace('\\\\', '\\');
      jsonRequest := TJSONObject.ParseJSONValue(TEncoding.ASCII.GetBytes(request), 0) as TJSONObject;
      restRequest.AddBody(jsonRequest);
      restRequest.Execute;

      jsonRequest   := TJSONObject.ParseJSONValue(restResponse.Content) as TJSONObject;
      if restResponse.StatusCode = 200 then begin
        odataMetadata := jsonRequest.Pairs[JSON_ID_HTTPS].JsonValue.ToString.Replace('"', '');
        sessionId     := jsonRequest.Pairs[JSON_ID_SESSION].JsonValue.ToString.Replace('"', '');
      end else begin
        GrabarLOG('json :' +  restResponse.StatusText);
      end;
      restRequest.Params.Clear;
    except
      on E: Exception do begin
        error := e.Message;
        montarMensajeError(body);
        SendEmail('Error en el Login (' + FormatDateTime('dd/mm/yyyy', fecha) + ')', body);
        raise Exception.Create('DameLogin: ' + E.Message);
      end;
    end;
  finally
    body.Free;
  end;
end;

procedure TServicioCfg.montarMensajeError(var lista: TStringList);
begin
  lista.Add('************************************************');
  lista.Add('********************  ERROR ********************');
  lista.Add('');
  lista.Add('Messsage: ' + error);
  lista.Add('');
  lista.Add('************************************************');
end;

procedure TServicioCfg.GrabarLOG(parMensaje: string);
begin
  with TStringList.Create do begin
    try
      if (FileExists(ExtractFilePath(Application.ExeName) + ARCHIVO_LOG_NAME)) then begin
        LoadFromFile(ExtractFilePath(Application.ExeName) + ARCHIVO_LOG_NAME);
        if (Count >= 200) then begin
          Delete(200 - 1);
        end;
        Insert(0, FormatDateTime('dd/mm/yyyy hh:mm:ss' , Now) + ' - ' + parMensaje);
      end else begin
        Add(FormatDateTime('dd/mm/yyyy hh:mm:ss' , Now) + ' - ' + parMensaje);
      end;
      SaveToFile(ExtractFilePath(Application.ExeName) + ARCHIVO_LOG_NAME);
    finally
      Free;
    end;
  end;
end;

procedure TServicioCfg.SendEmail(Subject: string; Body: TStrings);
var
  SMTP: TIdSMTP;
  Email: TIdMessage;
  SSLHandler: TIdSSLIOHandlerSocketOpenSSL;
begin
  SMTP := TIdSMTP.Create(nil);
  Email := TIdMessage.Create(nil);
  SSLHandler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);

  try
    SSLHandler.SSLOptions.Method := sslvTLSv1_2;
    SSLHandler.SSLOptions.Mode := sslmUnassigned;
    SSLHandler.SSLOptions.VerifyMode := [];
    SSLHandler.SSLOptions.VerifyDepth := 0;

    SMTP.IOHandler := SSLHandler;
    SMTP.Host := host;
    SMTP.Port := puerto;
    SMTP.Username := userEmail;
    SMTP.Password := passwordEmail;
    SMTP.UseTLS := utUseExplicitTLS;

    Email.From.Address := userEmail;
    Email.From.Name := 'SERVICIOS AXIONNET';
    Email.Recipients.EmailAddresses := destinatarios;
    Email.Subject := Subject;
    Email.Body := Body;
    TIdAttachmentFile.Create(Email.MessageParts, directorio + nombreFichero);
    SMTP.Connect;
    try
      SMTP.Send(Email);
    except
      on E: Exception do begin
        ShowMessage(E.Message);
      end;
    end;
    SMTP.Disconnect;
  finally
    SMTP.Free;
    Email.Free;
    SSLHandler.Free;
  end;
end;

end.
