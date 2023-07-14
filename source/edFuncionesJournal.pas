unit edFuncionesJournal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.SvcMgr, Vcl.Dialogs,
  Vcl.ExtCtrls, System.IniFiles, Vcl.Forms, System.StrUtils, Data.DB, REST.Client, REST.Authenticator.Simple, System.Variants,
  REST.Authenticator.Basic, System.NetEncoding, REST.Types, Xml.xmldom, Xml.XMLIntf, Xml.XMLDoc,  System.JSON,
  System.Generics.Collections, edFuncionesConexion, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error,
  FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.Phys.MSSQL, FireDAC.Phys.MSSQLDef, FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.Comp.Client, FireDAC.Comp.UI, FireDAC.Phys.ODBCBase, System.Math;

const
  TXT_TITULO_COBROS   = 'COBROS ';
  TXT_TITULO_FACTURAS = 'FACTURACION ';
  ACCOUNT_DEFAULT     = '103001';

type
  TJournalEntry = class(TObject)
    numLinea: Integer;
    accounting: string;
    contraAccount: string;
    debit: Extended;
    credit: Extended;
    taxGroup: string;
    officeGroup: string;
    departament: string;
    division: string;
    branch: string;
    sapProvider: string;
    commentary: string;
    hayRegistros: Boolean;
    srvCfg: TServicioCfg;

    elementoLines: TJSONObject;
    jsonRequest: TJSONObject;
    arLines: TJSONArray;

    constructor Create;
    procedure Clear;

    function iif(AValue: Boolean; Str1, Str2: Variant): Variant;
    {$region 'CREATE / DESTROY JSON OBJECTS'}
    procedure CreateJsonObjects;
    procedure DestroyJsonObjects;
    {$endregion}

    {$region 'JSON HEADER'}
    procedure AddJsonHeader(tituloCabecera: string);
    {$endregion}

    {$region 'JSON SAP'}
    procedure AddSAPInfo;
    procedure AddLineInfoSAP;
    {$endregion}

    {$region 'GENERA Y ENVIA'}
    procedure EnviarJSON(titulo: string; porPost: boolean = True);
    procedure GeneraEnviaJson(TituloHeader: string);
    {$endregion}

    {$region 'EXPORTAR COBROS POR FORMA DE PAGO'}
    procedure ExportarExcelCobrosPorFormaPago;
    procedure ExportarLineasCobrosPorFormaPago;
    {$endregion}

    {$region 'EXPORTAR COBROS POR BANCO'}
    procedure ExportarExcelCobrosPorBanco;
    procedure ExportarLineasCobrosPorBanco;
    {$endregion}

    {$region 'EXPORTAR FACTURAS'}
    procedure ExportarExcelFacturas;
    procedure ExportarLineasFacturas;
    {$endregion}

    procedure CambiarCuentaSAP;
    procedure montarMensajeError(var lista: TStringList);

    function HayRegistrosParaExportar(nombreTabla, campoFecha: string): Boolean;

    destructor Destroy; override;
  end;

implementation


{ TJournalEntry }

constructor TJournalEntry.Create;
begin
end;

destructor TJournalEntry.Destroy;
begin

  inherited;
end;

procedure TJournalEntry.Clear;
begin
  accounting    := '';
  contraAccount := '';
  debit         := 0;
  credit        := 0;
  taxGroup      := '';
  officeGroup   := '';
  departament   := '';
  division      := '';
  branch        := '';
  sapProvider   := '';
  commentary    := '';
end;

function TJournalEntry.iif(AValue: Boolean; Str1, Str2: Variant): Variant;
begin
  if AValue then
    Result := Str1
  else
    Result := Str2;
end;

{$region 'CREATE / DESTROY JSON OBJECTS'}
procedure TJournalEntry.CreateJsonObjects;
begin
  jsonRequest     := TJSONObject.Create;
  arLines         := TJSONArray.Create;
end;

procedure TJournalEntry.DestroyJsonObjects;
begin
  FreeAndNil(jsonRequest);
end;
{$endregion}

{$region 'JSON HEADER'}
procedure TJournalEntry.AddJsonHeader(tituloCabecera: string);
begin
  jsonRequest.AddPair('ReferenceDate', FormatDateTime('yyyy-mm-dd', iif((srvCfg.fecha = 0), Date, srvCfg.fecha)));
  jsonRequest.AddPair('Memo', tituloCabecera + FormatDateTime('dd/mm/yyyy', iif((srvCfg.fecha = 0), Date, srvCfg.fecha)));
  jsonRequest.AddPair('DueDate', FormatDateTime('yyyy-mm-dd', iif((srvCfg.fecha = 0), Date, srvCfg.fecha)));
end;
{$endregion}

{$region 'JSON PAYMENT'}
procedure TJournalEntry.AddSAPInfo;
begin
  SetRoundMode(rmDown);
  elementoLines := TJSONObject.Create;
  elementoLines.AddPair('Line_ID', TJSONNumber.Create(numLinea));
  elementoLines.AddPair('AccountCode', accounting);
  elementoLines.AddPair('Debit', TJSONNumber.Create(SimpleRoundTo(debit, -2)));
  elementoLines.AddPair('Credit', TJSONNumber.Create(SimpleRoundTo(credit, -2)));
  elementoLines.AddPair('DueDate', FormatDateTime('yyyy-mm-dd', iif((srvCfg.fecha = 0), Date, srvCfg.fecha)));
  elementoLines.AddPair('LineMemo', commentary);
  elementoLines.AddPair('CostingCode', officeGroup);
  elementoLines.AddPair('CostingCode2', departament);
  elementoLines.AddPair('CostingCode3', division);
  elementoLines.AddPair('CostingCode4', branch);
  elementoLines.AddPair('CostingCode5', sapProvider);
  arLines.AddElement(elementoLines);
end;

procedure TJournalEntry.AddLineInfoSAP;
begin
  jsonRequest.AddPair('JournalEntryLines', arLines);
end;
{$endregion}

{$region 'GENERA Y ENVIA'}
procedure TJournalEntry.EnviarJSON(titulo: string; porPost: boolean);
var jsonResponse: TJSONObject; lstGeneraJson, jsonResponseToFile: TStringList; request: string;
begin
  jsonResponse  := TJSONObject.Create;
  lstGeneraJson := TStringList.Create;
  try
    srvCfg.restClient.BaseURL := srvCfg.baseUrl;
    if porPost then begin
      srvCfg.restRequest.Method := rmPOST;
      srvCfg.restRequest.AddBody(jsonRequest);
    end else begin
      srvCfg.restRequest.Method := rmGET;
      srvCfg.restClient.BaseURL := srvCfg.baseurl + '(' + srvCfg.numAsiento + ')';
    end;

    if (porPost) then begin
      lstGeneraJson.Add(jsonRequest.ToString);
      srvCfg.nombreFichero := titulo.Trim + '_' + FormatDateTime('yyyymmdd', srvCfg.fecha) + '.json';
      lstGeneraJson.SaveToFile(srvCfg.directorio + srvCfg.nombreFichero);
    end;
    srvCfg.restRequest.Execute;
    jsonResponse := TJSONObject.ParseJSONValue(TEncoding.ASCII.GetBytes(srvCfg.restResponse.Content), 0) as TJSONObject;

    if (Assigned(jsonResponse)) then begin
      if not porPost then begin
        jsonResponseToFile := TStringList.Create;
        try
          jsonResponseToFile.Text := jsonResponse.ToString;
          jsonResponseToFile.SaveToFile('response_' + FormatDateTime('yyyymmddhhmmss', Now) + '.json');
        finally
          FreeAndNil(jsonResponseToFile);
        end;
      end else begin
        if (srvCfg.restResponse.StatusCode = 200) or (srvCfg.restResponse.StatusCode = 201) then begin
          srvCfg.envioOk := true;
          srvCfg.GrabarLog('Fichero enviado!');
        end
        else begin
          srvCfg.envioOk := false;
          srvCfg.statusCode := srvCfg.restResponse.StatusCode;
          srvCfg.error := jsonResponse.ToString;
          srvCfg.GrabarLog('StatusCode: ' + srvCfg.restResponse.StatusCode.ToString);
          srvCfg.GrabarLog(jsonResponse.ToString);
        end;
      end;
    end else begin
      srvCfg.GrabarLog('JSONResponse is nil');
    end;
  finally
    jsonResponse.Free;
    lstGeneraJson.Free;
  end;
end;

procedure TJournalEntry.GeneraEnviaJson(TituloHeader: string);
begin
  AddJsonHeader(TituloHeader);
  AddLineInfoSAP;
  EnviarJSON(TituloHeader);
end;
{$endregion}

{$region 'EXPORTAR FACTURAS'}
procedure TJournalEntry.ExportarExcelFacturas;
var body: TStringList;
begin
  if HayRegistrosParaExportar('DocAdmin', 'FecDoc') then begin
    body := TStringList.Create;
    CreateJsonObjects;
    try
      try
        ExportarLineasFacturas;
        if hayRegistros then begin
          GeneraEnviaJson(TXT_TITULO_FACTURAS);
          if not srvCfg.envioOk then begin
            montarMensajeError(body);
            srvCfg.SendEmail('Error Envío Facturas a SAP (' + FormatDateTime('dd/mm/yyyy', iif((srvCfg.fecha = 0), Date, srvCfg.fecha)) + ')', body);
          end;
        end;
      except
        on E: Exception do begin
          raise
        end;
      end;
    finally
      body.Free;
      DestroyJsonObjects;
    end;
  end;
end;

procedure TJournalEntry.ExportarLineasFacturas;
var qGenerica: TFDQuery; numRegistros: Integer;
begin
  qGenerica := TFDQuery.Create(nil);
  try
    qGenerica.Connection := srvCfg.conexionBaseDatos;
    try
      qGenerica.SQL.Add('SELECT ''103001'' Accounting,');
      qGenerica.SQL.Add('       SUM(l.Importe) Debit,');
      qGenerica.SQL.Add('	      0 Credit,');
      qGenerica.SQL.Add('	      '''' TaxGroup,');
      qGenerica.SQL.Add('	      '''' OfficeGroup,');
      qGenerica.SQL.Add('	      '''' Department,');
      qGenerica.SQL.Add('	      '''' Division,');
      qGenerica.SQL.Add('       '''' Branch,');
      qGenerica.SQL.Add('	      '''' SAPProvider');
      qGenerica.SQL.Add('  FROM DocAdmin d');
      qGenerica.SQL.Add('       INNER JOIN LinAdmin l on l.IdDocAdmin = d.Ident');
      qGenerica.SQL.Add(' 	    INNER JOIN Emisores e ON e.IdEmisor = d.IdEmisor');
      qGenerica.SQL.Add('	      INNER JOIN Centros c ON c.IdCentro = e.IdCentro');
      qGenerica.SQL.Add(' WHERE d.FecDoc = ' + QuotedStr(FormatDateTime('yyyymmdd', iif((srvCfg.fecha = 0), Date, srvCfg.fecha))));
      qGenerica.SQL.Add('   AND d.Doc = ''F''');
      qGenerica.SQL.Add('   AND l.IdPagoCli IS NULL');
      qGenerica.SQL.Add('UNION ALL');
      qGenerica.SQL.Add('SELECT SUBSTRING(t.CodInt, 1, 6) Accounting,');
      qGenerica.SQL.Add('       SUM(ABS(l.Importe)) Debit,');
      qGenerica.SQL.Add('	      0 Credit,');
      qGenerica.SQL.Add('	      '''' TaxGroup,');
      qGenerica.SQL.Add('	      ISNULL(c.Delegacion, '''') OfficeGroup,');
      qGenerica.SQL.Add('       ISNULL(g.Descripcio, '''') Department,');
      qGenerica.SQL.Add('	      ISNULL(omc.Descripcio, '''') Division,');
      qGenerica.SQL.Add('	      ISNULL(c.DescCorta, '''') Branch,');
      qGenerica.SQL.Add('	      ISNULL(col.CodInt, '''') SAPProvider');
      qGenerica.SQL.Add('  FROM DocAdmin d');
      qGenerica.SQL.Add('	      INNER JOIN LinAdmin l ON l.IdDocAdmin = d.Ident');
      qGenerica.SQL.Add('	      INNER JOIN Emisores e ON e.IdEmisor = d.IdEmisor');
      qGenerica.SQL.Add('       LEFT JOIN DeudaCli dc ON dc.IdDeudaCli = l.IdDeudaCli');
      qGenerica.SQL.Add('	      INNER JOIN TtosMed tm ON tm.IdPac = dc.IdPac AND tm.NumTto = dc.NumTto');
      qGenerica.SQL.Add('  	    INNER JOIN Centros c ON c.IdCentro = e.IdCentro');
      qGenerica.SQL.Add('	      LEFT JOIN TColabos col ON col.IdCol = tm.IdCol');
      qGenerica.SQL.Add('	      INNER JOIN Tratamientos_Tarifas tt ON tt.IdTratamientoTarifa = tm.IdTto');
      qGenerica.SQL.Add('	      INNER JOIN Tratamientos t ON t.IdTratamiento = tt.IdTratamiento');
      qGenerica.SQL.Add('	      INNER JOIN TGrupos g ON g.IdGrupo = t.IdGrupo');
      qGenerica.SQL.Add('	      INNER JOIN TEspecOMC omc ON omc.IdTipoEspec = t.IdTipoEspec');
      qGenerica.SQL.Add(' WHERE d.FecDoc = ' + QuotedStr(FormatDateTime('yyyymmdd', iif((srvCfg.fecha = 0), Date, srvCfg.fecha))));
      qGenerica.SQL.Add('   AND d.Doc = ''F''');
      qGenerica.SQL.Add('   AND l.Importe <> 0');
      qGenerica.SQL.Add('   AND d.Abono IS NOT NULL');
      qGenerica.SQL.Add('GROUP BY SUBSTRING(t.CodInt, 1, 6),');
      qGenerica.SQL.Add('     	  ISNULL(c.Delegacion, ''''),');
      qGenerica.SQL.Add('         ISNULL(g.Descripcio, ''''),');
      qGenerica.SQL.Add(' 	      ISNULL(omc.Descripcio, ''''),');
      qGenerica.SQL.Add(' 	      ISNULL(c.DescCorta, ''''),');
      qGenerica.SQL.Add(' 	      ISNULL(col.CodInt, '''')');
      qGenerica.SQL.Add('UNION ALL');
      qGenerica.SQL.Add('SELECT SUBSTRING(t.CodInt, 9, 14) Accounting,');
      qGenerica.SQL.Add('       0 Debit,');
      qGenerica.SQL.Add('	      SUM(l.Importe) Credit,');
      qGenerica.SQL.Add('	      '''' TaxGroup,');
      qGenerica.SQL.Add('	      ISNULL(c.Delegacion, '''') OfficeGroup,');
      qGenerica.SQL.Add('       ISNULL(g.Descripcio, '''') Department,');
      qGenerica.SQL.Add('	      ISNULL(omc.Descripcio, '''') Division,');
      qGenerica.SQL.Add('	      ISNULL(c.DescCorta, '''') Branch,');
      qGenerica.SQL.Add('	      ISNULL(col.CodInt, '''') SAPProvider');
      qGenerica.SQL.Add('  FROM DocAdmin d');
      qGenerica.SQL.Add('	      INNER JOIN LinAdmin l ON l.IdDocAdmin = d.Ident');
      qGenerica.SQL.Add('	      INNER JOIN Emisores e ON e.IdEmisor = d.IdEmisor');
      qGenerica.SQL.Add('       LEFT JOIN DeudaCli dc ON dc.IdDeudaCli = l.IdDeudaCli');
      qGenerica.SQL.Add('	      INNER JOIN TtosMed tm ON tm.IdPac = dc.IdPac AND tm.NumTto = dc.NumTto');
      qGenerica.SQL.Add('	      INNER JOIN Centros c ON c.IdCentro = e.IdCentro');
      qGenerica.SQL.Add('	      LEFT JOIN TColabos col ON col.IdCol = tm.IdCol');
      qGenerica.SQL.Add('	      INNER JOIN Tratamientos_Tarifas tt ON tt.IdTratamientoTarifa = tm.IdTto');
      qGenerica.SQL.Add('	      INNER JOIN Tratamientos t ON t.IdTratamiento = tt.IdTratamiento');
      qGenerica.SQL.Add('	      INNER JOIN TGrupos g ON g.IdGrupo = t.IdGrupo');
      qGenerica.SQL.Add('	      INNER JOIN TEspecOMC omc ON omc.IdTipoEspec = t.IdTipoEspec');
      qGenerica.SQL.Add(' WHERE d.FecDoc = ' + QuotedStr(FormatDateTime('yyyymmdd', iif((srvCfg.fecha = 0), Date, srvCfg.fecha))));
      qGenerica.SQL.Add('   AND l.Importe > 0');
      qGenerica.SQL.Add('   AND d.Doc = ''F''');
      qGenerica.SQL.Add('   AND l.IdPagoCli IS NULL');
      qGenerica.SQL.Add('GROUP BY SUBSTRING(t.CodInt, 9, 14),');
      qGenerica.SQL.Add('	        ISNULL(c.Delegacion, ''''),');
      qGenerica.SQL.Add('         ISNULL(g.Descripcio, ''''),');
      qGenerica.SQL.Add('	        ISNULL(omc.Descripcio, ''''),');
      qGenerica.SQL.Add('	        ISNULL(c.DescCorta, ''''),');
      qGenerica.SQL.Add('	        ISNULL(col.CodInt, '''')');
      qGenerica.Open;

      {$region 'SQL SERVER'}
        {SELECT '103001' Accounting,
                SUM(l.Importe) Debit,
                0 Credit,
                '' TaxGroup,
                '' OfficeGroup,
                '' Department,
                '' Division,
                '' Branch,
                '' SAPProvider
           FROM DocAdmin d
                INNER JOIN LinAdmin l on l.IdDocAdmin = d.Ident
                INNER JOIN Emisores e ON e.IdEmisor = d.IdEmisor
                INNER JOIN Centros c ON c.IdCentro = e.IdCentro
          WHERE d.fecdoc = '20230321'
            AND d.Doc = 'F'
            AND l.IdPagoCli IS NULL
          UNION ALL
         SELECT SUBSTRING(t.CodInt, 1, 6) Accounting,
                SUM(ABS(l.Importe)) Debit,
                0 Credit,
                '' TaxGroup,
                ISNULL(c.Delegacion, '') OfficeGroup,
                ISNULL(g.Descripcio, '') Department,
                ISNULL(omc.Descripcio, '') Division,
                ISNULL(c.DescCorta, '') Branch,
                ISNULL(col.CodInt, '') SAPProvider
           FROM DocAdmin d
                INNER JOIN LinAdmin l ON l.IdDocAdmin = d.Ident
                INNER JOIN Emisores e ON e.IdEmisor = d.IdEmisor
                LEFT JOIN DeudaCli dc ON dc.IdDeudaCli = l.IdDeudaCli
                INNER JOIN TtosMed tm ON tm.IdPac = dc.IdPac AND tm.NumTto = dc.NumTto
                INNER JOIN Centros c ON c.IdCentro = e.IdCentro
                LEFT JOIN TColabos col ON col.IdCol = tm.IdCol
                INNER JOIN Tratamientos_Tarifas tt ON tt.IdTratamientoTarifa = tm.IdTto
                INNER JOIN Tratamientos t ON t.IdTratamiento = tt.IdTratamiento
                INNER JOIN TGrupos g ON g.IdGrupo = t.IdGrupo
                INNER JOIN TEspecOMC omc ON omc.IdTipoEspec = t.IdTipoEspec
          WHERE d.fecdoc = '20230321'
            AND d.Doc = 'F'
            AND d.Abono IS NOT NULL
            AND l.Importe <> 0
         GROUP BY SUBSTRING(t.CodInt, 1, 6),
                  ISNULL(c.Delegacion, ''),
                  ISNULL(g.Descripcio, ''),
                  ISNULL(omc.Descripcio, ''),
                  ISNULL(c.DescCorta, ''),
                  ISNULL(col.CodInt, '')
         UNION ALL
         SELECT SUBSTRING(t.CodInt, 9, 14) Accounting,
                0 Debit,
                SUM(l.Importe) Credit,
                '' TaxGroup,
                ISNULL(c.Delegacion, '') OfficeGroup,
                ISNULL(g.Descripcio, '') Department,
                ISNULL(omc.Descripcio, '') Division,
                ISNULL(c.DescCorta, '') Branch,
                ISNULL(col.CodInt, '') SAPProvider
           FROM DocAdmin d
                INNER JOIN LinAdmin l ON l.IdDocAdmin = d.Ident
                INNER JOIN Emisores e ON e.IdEmisor = d.IdEmisor
                LEFT JOIN DeudaCli dc ON dc.IdDeudaCli = l.IdDeudaCli
                INNER JOIN TtosMed tm ON tm.IdPac = dc.IdPac AND tm.NumTto = dc.NumTto
                INNER JOIN Centros c ON c.IdCentro = e.IdCentro
                LEFT JOIN TColabos col ON col.IdCol = tm.IdCol
                INNER JOIN Tratamientos_Tarifas tt ON tt.IdTratamientoTarifa = tm.IdTto
                INNER JOIN Tratamientos t ON t.IdTratamiento = tt.IdTratamiento
                INNER JOIN TGrupos g ON g.IdGrupo = t.IdGrupo
                INNER JOIN TEspecOMC omc ON omc.IdTipoEspec = t.IdTipoEspec
          WHERE d.FecDoc = '20230321'
            AND l.Importe > 0
            AND l.IdPagoCli IS NULL
            AND d.Doc = 'F'
          GROUP BY SUBSTRING(t.CodInt, 9, 14),
                   ISNULL(c.Delegacion, ''),
                   ISNULL(g.Descripcio, ''),
                   ISNULL(omc.Descripcio, ''),
                   ISNULL(c.DescCorta, ''),
                   ISNULL(col.CodInt, '') }
         {$endregion}
      hayRegistros := not qGenerica.IsEmpty;

      numlinea := 0; numRegistros := 0;
      while not qGenerica.Eof do begin
        Inc(numRegistros);
        Clear;

        accounting  := qGenerica.FieldByName('Accounting').AsString;
        debit       := qGenerica.FieldByName('Debit').AsFloat;
        credit      := qGenerica.FieldByName('Credit').AsFloat;
        taxGroup    := qGenerica.FieldByName('TaxGroup').AsString;
        officeGroup := qGenerica.FieldByName('OfficeGroup').AsString;
        departament := qGenerica.FieldByName('Department').AsString;
        division    := qGenerica.FieldByName('Division').AsString;
        branch      := qGenerica.FieldByName('Branch').AsString;
        sapProvider := qGenerica.FieldByName('SAPProvider').AsString;
        commentary  := TXT_TITULO_FACTURAS + FormatDateTime('dd/mm/yyyy', iif((srvCfg.fecha = 0), Date, srvCfg.fecha));
        AddSAPInfo;

        qGenerica.Next;
        Inc(numLinea);
      end;
      if numRegistros > 0 then
        srvCfg.GrabarLog('Exportando ' + numRegistros.ToString + ' facturas a SAP (' + FormatDateTime('dd/mm/yyyy', srvCfg.fecha) + ')');
    except
      on E: Exception do begin
        raise
      end;
    end;
  finally
    FreeAndNil(qGenerica);
  end;
end;
{$endregion}

{$region 'EXPORTAR COBROS POR FORMA DE PAGO'}
procedure TJournalEntry.ExportarExcelCobrosPorFormaPago;
var body: TStringList;
begin
  if HayRegistrosParaExportar('PagoCli', 'FecPago') then begin
    body := TStringList.Create;
    CreateJsonObjects;
    try
      try
        ExportarLineasCobrosPorFormaPago;
        if hayRegistros then
          GeneraEnviaJson(TXT_TITULO_COBROS);
          if not srvCfg.envioOk then begin
            montarMensajeError(body);
            srvCfg.SendEmail('Error Envío Cobros a SAP (' + FormatDateTime('dd/mm/yyyy', srvCfg.fecha) + ')', body);
          end;
          if FileExists(srvCfg.nombreFichero) then
            DeleteFile(srvCfg.nombreFichero);
      except
        on E: Exception do begin
          raise
        end;
      end;
    finally
      body.Free;
      DestroyJsonObjects;
    end;
  end;
end;

procedure TJournalEntry.ExportarLineasCobrosPorFormaPago;
var qGenerica: TFDQuery; numRegistros: Integer;
begin
  qGenerica := TFDQuery.Create(nil);
  try
    qGenerica.Connection := srvCfg.conexionBaseDatos;
    try
      qGenerica.SQL.Add('DECLARE @PaymentTemp TABLE (Accounting VARCHAR(10),');
      qGenerica.SQL.Add('                            Tipo VARCHAR(1),');
      qGenerica.SQL.Add('							               Debit FLOAT,');
      qGenerica.SQL.Add('				               			 Credit FLOAT,');
      qGenerica.SQL.Add('							               TaxGroup VARCHAR(3),');
      qGenerica.SQL.Add('							               OfficeGroup VARCHAR(20),');
      qGenerica.SQL.Add('							               Department VARCHAR(30),');
      qGenerica.SQL.Add('							               Division VARCHAR(30),');
      qGenerica.SQL.Add('						                 Branch VARCHAR(4),');
      qGenerica.SQL.Add('						               	 SAPProvider VARCHAR(30),');
      qGenerica.SQL.Add('							               Commentary VARCHAR(500))');
      qGenerica.SQL.Add('');
      qGenerica.SQL.Add('DECLARE @Accounting VARCHAR(10),');
      qGenerica.SQL.Add('        @IdPagoCli INTEGER,');
      qGenerica.SQL.Add('        @Pagado FLOAT,');
      qGenerica.SQL.Add('	       @Diferencia FLOAT,');
      qGenerica.SQL.Add('		     @TieneConciliacion INTEGER,');
      qGenerica.SQL.Add('		     @DescCorta VARCHAR(5),');
      qGenerica.SQL.Add('	       @Delegacion VARCHAR(20),');
      qGenerica.SQL.Add('	       @Department VARCHAR(3),');
      qGenerica.SQL.Add('	       @Division VARCHAR(10),');
      qGenerica.SQL.Add('	       @SAPProvider VARCHAR(20),');
      qGenerica.SQL.Add('		     @DescFormaPago VARCHAR(500)');
      qGenerica.SQL.Add('');
      qGenerica.SQL.Add('SET @Department  = ''COB''');
      qGenerica.SQL.Add('SET @Division    = ''COB CON''');
      qGenerica.SQL.Add('SET @SAPProvider = ''P01003''');
      qGenerica.SQL.Add('');
      qGenerica.SQL.Add('DECLARE Cursor_Pagos CURSOR FOR');
      qGenerica.SQL.Add('	SELECT b.SCCodSCta,');
      qGenerica.SQL.Add('	       p.IdPagoCli,');
      qGenerica.SQL.Add('	       p.Pagado,');
      qGenerica.SQL.Add('		     (p.Pagado - ISNULL(SUM(dp.Pagado), 0)) Diferencia,');
      qGenerica.SQL.Add('		     ISNULL(dp.IdPagoCli, 0) TieneConciliacion,');
      qGenerica.SQL.Add('		     c.DescCorta,');
      qGenerica.SQL.Add('	       c.Delegacion,');
      qGenerica.SQL.Add('		     tp.Descripcio + '' '' +  ISNULL(c.DescCorta, '''') Commentary');
      qGenerica.SQL.Add('   FROM PagoCli p');
      qGenerica.SQL.Add('        LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli');
      qGenerica.SQL.Add('  	     LEFT JOIN Centros c ON c.IdCentro = p.IdCentro');
      qGenerica.SQL.Add('		     LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago');
      qGenerica.SQL.Add('		     LEFT JOIN Bancos b ON b.IdBanco = p.IdBanco');
      qGenerica.SQL.Add('  WHERE p.FecPago = ' + QuotedStr(FormatDateTime('yyyymmdd', iif((srvCfg.fecha = 0), Date, srvCfg.fecha))));
      qGenerica.SQL.Add('    AND p.IdForPago <> 31');
      qGenerica.SQL.Add(' GROUP BY b.SCCodSCta, p.IdPagoCli, p.Pagado, dp.IdPagoCli, c.DescCorta, c.Delegacion, tp.Descripcio, tp.Descripcio + '' '' +  ISNULL(c.DescCorta, '''')');
      qGenerica.SQL.Add('	OPEN Cursor_Pagos');
      qGenerica.SQL.Add('	FETCH NEXT FROM Cursor_Pagos');
      qGenerica.SQL.Add('	INTO @Accounting, @IdPagoCli, @Pagado, @Diferencia, @TieneConciliacion, @DescCorta, @Delegacion, @DescFormaPago');
      qGenerica.SQL.Add('	WHILE @@FETCH_STATUS = 0');
      qGenerica.SQL.Add('	BEGIN');
      qGenerica.SQL.Add('		INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)');
      qGenerica.SQL.Add('		SELECT (CASE WHEN p.IdForPago = 2  THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 5  THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 6  THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 7  THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 14 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 17 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 18 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 19 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 21 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 22 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 32 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 33 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 34 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 35 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 37 THEN ''102001''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 28 THEN ''102003''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 29 THEN ''102003''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 15 THEN ''102004''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 16 THEN ''102004''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 20 THEN ''102004''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 23 THEN ''102004''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 25 THEN ''102004''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 24 THEN ''102005''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 26 THEN ''102005''');
      qGenerica.SQL.Add('                WHEN p.IdForPago = 27 THEN ''102006''');
      qGenerica.SQL.Add('                ELSE ''102001''');
      qGenerica.SQL.Add('           END) Accounting,');
      qGenerica.SQL.Add('           ''H'' Tipo,');
      qGenerica.SQL.Add('           ISNULL(dp.Pagado, p.Pagado) Debit,');
      qGenerica.SQL.Add('           0 Credit,');
      qGenerica.SQL.Add('           '''' Taxgroup,');
      qGenerica.SQL.Add('           ISNULL(c.Delegacion, '''') OfficeGroup,');
      qGenerica.SQL.Add('           ISNULL(g.Descripcio, @Department) Departament,');
      qGenerica.SQL.Add('           ISNULL(omc.Descripcio, @Division) Division,');
      qGenerica.SQL.Add('           ISNULL(c.DescCorta, '''') Branch,');
      qGenerica.SQL.Add('           ISNULL(tc.CodInt, @SAPProvider) SAPProvider,');
      qGenerica.SQL.Add('           tp.Descripcio + '' '' +  ISNULL(c.DescCorta, '''') Commentary');
      qGenerica.SQL.Add('      FROM PagoCli p');
      qGenerica.SQL.Add('           LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli');
      qGenerica.SQL.Add('		   	    LEFT JOIN DeudaCli dc on dc.IdDeudaCli = dp.IdDeudaCli');
      qGenerica.SQL.Add('           LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago');
      qGenerica.SQL.Add('           LEFT JOIN TtosMed tt ON tt.Ident = dc.IdentTM');
      qGenerica.SQL.Add('           LEFT JOIN Centros c ON c.IdCentro = p.IdCentro');
      qGenerica.SQL.Add('           LEFT JOIN TColabos tc ON tt.IdCol = tc.IdCol');
      qGenerica.SQL.Add('           LEFT JOIN Tratamientos_tarifas t ON t.IdTratamientotarifa = tt.IdTto');
      qGenerica.SQL.Add('           LEFT JOIN Tratamientos ttm ON ttm.IdTratamiento = t.IdTratamiento');
      qGenerica.SQL.Add('           LEFT JOIN TGrupos g ON g.IdGrupo = ttm.IdGrupo');
      qGenerica.SQL.Add('           LEFT JOIN TEspecOMC omc ON omc.IdTipoEspec = ttm.IdTipoEspec');
      qGenerica.SQL.Add('     WHERE p.IdPagoCli = @IdPagoCli');
      qGenerica.SQL.Add('	  UNION ALL');
      qGenerica.SQL.Add('   SELECT ''103001'' Accounting,');
      qGenerica.SQL.Add('          ''D'' Tipo,');
      qGenerica.SQL.Add('           0 Debit,');
      qGenerica.SQL.Add('           ISNULL(dp.Pagado, p.Pagado) Credit,');
      qGenerica.SQL.Add('           '''' Taxgroup,');
      qGenerica.SQL.Add('           ISNULL(c.Delegacion, '''') OfficeGroup,');
      qGenerica.SQL.Add('           ISNULL(g.Descripcio, @Department) Departament,');
      qGenerica.SQL.Add('           ISNULL(omc.Descripcio, @Division) Division,');
      qGenerica.SQL.Add('           ISNULL(c.DescCorta, '''') Branch,');
      qGenerica.SQL.Add('           ISNULL(tc.CodInt, @SAPProvider) SAPProvider,');
      qGenerica.SQL.Add('           tp.Descripcio + '' '' + ISNULL(c.DescCorta, '''') Commentary');
      qGenerica.SQL.Add('      FROM PagoCli p');
      qGenerica.SQL.Add('           LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli');
      qGenerica.SQL.Add('			      LEFT JOIN DeudaCli dc on dc.IdDeudaCli = dp.IdDeudaCli');
      qGenerica.SQL.Add('           LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago');
      qGenerica.SQL.Add('           LEFT JOIN TtosMed tt ON tt.Ident = dc.IdentTM');
      qGenerica.SQL.Add('           LEFT JOIN Centros c ON c.IdCentro = p.IdCentro');
      qGenerica.SQL.Add('           LEFT JOIN TColabos tc ON tt.IdCol = tc.IdCol');
      qGenerica.SQL.Add('           LEFT JOIN Tratamientos_tarifas t ON t.IdTratamientotarifa = tt.IdTto');
      qGenerica.SQL.Add('           LEFT JOIN Tratamientos ttm ON ttm.IdTratamiento = t.IdTratamiento');
      qGenerica.SQL.Add('           LEFT JOIN TGrupos g ON g.IdGrupo = ttm.IdGrupo');
      qGenerica.SQL.Add('           LEFT JOIN TEspecOMC omc ON omc.IdTipoEspec = ttm.IdTipoEspec');
      qGenerica.SQL.Add('      WHERE p.IdPagoCli = @IdPagoCli');
      qGenerica.SQL.Add('   ORDER BY Commentary');
      qGenerica.SQL.Add('		IF ((@Diferencia > 0) AND (@TieneConciliacion > 0))');
      qGenerica.SQL.Add('		BEGIN');
      qGenerica.SQL.Add('  			INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)');
      qGenerica.SQL.Add('		  	VALUES (''102001'', ''H'', @Diferencia, 0, '''', @Delegacion, @Department, @Division, @DescCorta, @SAPProvider, @DescFormaPago)');
      qGenerica.SQL.Add('  			INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)');
      qGenerica.SQL.Add('			  VALUES (''103001'', ''D'', 0, @Diferencia, '''', @Delegacion, @Department, @Division, @DescCorta, @SAPProvider, @DescFormaPago)');
      qGenerica.SQL.Add('		END');
      qGenerica.SQL.Add('		FETCH NEXT FROM Cursor_Pagos');
      qGenerica.SQL.Add('		INTO @Accounting, @IdPagoCli, @Pagado, @Diferencia, @TieneConciliacion, @DescCorta, @Delegacion, @DescFormaPago');
      qGenerica.SQL.Add('	END');
      qGenerica.SQL.Add('CLOSE Cursor_Pagos');
      qGenerica.SQL.Add('DEALLOCATE Cursor_Pagos');
      qGenerica.SQL.Add('');
      qGenerica.SQL.Add('SELECT Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary');
      qGenerica.SQL.Add('  FROM @PaymentTemp');
      qGenerica.SQL.Add('ORDER BY Branch, Commentary');
      qGenerica.Open;

      {$region 'SQL SERVER'}
        {DECLARE @PaymentTemp TABLE (Accounting VARCHAR(10),
                                    Tipo VARCHAR(1),
                                    Debit FLOAT,
                                    Credit FLOAT,
                                    TaxGroup VARCHAR(3),
                                    OfficeGroup VARCHAR(20),
                                    Department VARCHAR(30),
                                    Division VARCHAR(30),
                                    Branch VARCHAR(4),
                                    SAPProvider VARCHAR(30),
                                    Commentary VARCHAR(500))
        DECLARE @Accounting VARCHAR(10),
                @IdPagoCli INTEGER,
                @Pagado FLOAT,
                @Diferencia FLOAT,
                @TieneConciliacion INTEGER,
                @DescCorta VARCHAR(5),
                @Delegacion VARCHAR(20),
                @DescFormaPago VARCHAR(500)

        DECLARE Cursor_Pagos CURSOR FOR
          SELECT b.SCCodSCta,
                 p.IdPagoCli,
                 p.Pagado,
                 (p.Pagado - ISNULL(SUM(dp.Pagado), 0)) Diferencia,
                 ISNULL(dp.IdPagoCli, 0) TieneConciliacion,
                 c.DescCorta,
                 c.Delegacion,
                 tp.Descripcio + ' ' +  ISNULL(c.DescCorta, '') Commentary
             FROM PagoCli p
                  LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli
                  LEFT JOIN Centros c ON c.IdCentro = p.IdCentro
                  LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago
                  LEFT JOIN Bancos b ON b.IdBanco = p.IdBanco
            WHERE p.FecPago = '20230329'
            AND p.IdForPago <> 31
            GROUP BY b.SCCodSCta, p.IdPagoCli, p.Pagado, dp.IdPagoCli, c.DescCorta, c.Delegacion, tp.Descripcio, tp.Descripcio + ' ' +  ISNULL(c.DescCorta, '')

          OPEN Cursor_Pagos
          FETCH NEXT FROM Cursor_Pagos
          INTO @Accounting, @IdPagoCli, @Pagado, @Diferencia, @TieneConciliacion, @DescCorta, @Delegacion, @DescFormaPago

          WHILE @@FETCH_STATUS = 0
          BEGIN
            INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)
            SELECT (CASE WHEN p.IdForPago = 2  THEN '102001'
                         WHEN p.IdForPago = 5  THEN '102001'
                         WHEN p.IdForPago = 6  THEN '102001'
                         WHEN p.IdForPago = 7  THEN '102001'
                         WHEN p.IdForPago = 14 THEN '102001'
                         WHEN p.IdForPago = 17 THEN '102001'
                         WHEN p.IdForPago = 18 THEN '102001'
                         WHEN p.IdForPago = 19 THEN '102001'
                         WHEN p.IdForPago = 21 THEN '102001'
                         WHEN p.IdForPago = 22 THEN '102001'
                         WHEN p.IdForPago = 32 THEN '102001'
                         WHEN p.IdForPago = 28 THEN '102003'
                         WHEN p.IdForPago = 29 THEN '102003'
                         WHEN p.IdForPago = 15 THEN '102004'
                         WHEN p.IdForPago = 16 THEN '102004'
                         WHEN p.IdForPago = 20 THEN '102004'
                         WHEN p.IdForPago = 23 THEN '102004'
                         WHEN p.IdForPago = 25 THEN '102004'
                         WHEN p.IdForPago = 24 THEN '102005'
                         WHEN p.IdForPago = 26 THEN '102005'
                         WHEN p.IdForPago = 27 THEN '102006'
                         ELSE '102001'
                   END) Accounting,
                  'H' Tipo,
                   ISNULL(dp.Pagado, p.Pagado) Debit,
                   0 Credit,
                   '' Taxgroup,
                   ISNULL(c.Delegacion, '') OfficeGroup,
                   ISNULL(g.Descripcio, 'COB') Departament,
                   ISNULL(omc.Descripcio, 'COBCON') Division,
                   ISNULL(c.DescCorta, '') Branch,
                   ISNULL(tc.CodInt, 'P01003') SAPProvider,
                   tp.Descripcio + ' ' +  ISNULL(c.DescCorta, '') Commentary
              FROM PagoCli p
                   LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli
                   LEFT JOIN DeudaCli dc on dc.IdDeudaCli = dp.IdDeudaCli
                   LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago
                   LEFT JOIN TtosMed tt ON tt.Ident = dc.IdentTM
                   LEFT JOIN Centros c ON c.IdCentro = p.IdCentro
                   LEFT JOIN TColabos tc ON tt.IdCol = tc.IdCol
                   LEFT JOIN Tratamientos_tarifas t ON t.IdTratamientotarifa = tt.IdTto
                   LEFT JOIN Tratamientos ttm ON ttm.IdTratamiento = t.IdTratamiento
                   LEFT JOIN TGrupos g ON g.IdGrupo = ttm.IdGrupo
                   LEFT JOIN TEspecOMC omc ON omc.IdTipoEspec = ttm.IdTipoEspec
             WHERE p.IdPagoCli = @IdPagoCli
            UNION ALL
            SELECT '103001' Accounting,
                   'D' Tipo,
                   0 Debit,
                   ISNULL(dp.Pagado, p.Pagado) Credit,
                   '' Taxgroup,
                   ISNULL(c.Delegacion, '') OfficeGroup,
                   ISNULL(g.Descripcio, 'COB') Departament,
                   ISNULL(omc.Descripcio, 'COBCON') Division,
                   ISNULL(c.DescCorta, '') Branch,
                   ISNULL(tc.CodInt, 'P01003') SAPProvider,
                   tp.Descripcio + ' ' + ISNULL(c.DescCorta, '') Commentary
              FROM PagoCli p
                   LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli
                   LEFT JOIN DeudaCli dc on dc.IdDeudaCli = dp.IdDeudaCli
                   LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago
                   LEFT JOIN TtosMed tt ON tt.Ident = dc.IdentTM
                   LEFT JOIN Centros c ON c.IdCentro = p.IdCentro
                   LEFT JOIN TColabos tc ON tt.IdCol = tc.IdCol
                   LEFT JOIN Tratamientos_tarifas t ON t.IdTratamientotarifa = tt.IdTto
                   LEFT JOIN Tratamientos ttm ON ttm.IdTratamiento = t.IdTratamiento
                   LEFT JOIN TGrupos g ON g.IdGrupo = ttm.IdGrupo
                   LEFT JOIN TEspecOMC omc ON omc.IdTipoEspec = ttm.IdTipoEspec
             WHERE p.IdPagoCli = @IdPagoCli
            ORDER BY Commentary

            IF ((@Diferencia > 0) AND (@TieneConciliacion > 0))
            BEGIN
                INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)
              VALUES ('102001', 'H', @Diferencia, 0, '', @DescCorta,'COB', 'COBCON', @DescCorta, 'P01003', @DescFormaPago)
                INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)
              VALUES ('103001', 'D', 0, @Diferencia, '', @DescCorta, 'COB', 'COBCON', @DescCorta, 'P01003', @DescFormaPago)
            END

            FETCH NEXT FROM Cursor_Pagos
            INTO @Accounting, @IdPagoCli, @Pagado, @Diferencia, @TieneConciliacion, @DescCorta, @Delegacion, @DescFormaPago
          END
        CLOSE Cursor_Pagos
        DEALLOCATE Cursor_Pagos

        SELECT Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary
         FROM @PaymentTemp
        ORDER BY Branch, Commentary}
      {$endregion}
      numlinea := 0; numRegistros := 0;
      hayRegistros := not qGenerica.IsEmpty;
      while not qGenerica.Eof do begin
        Clear;

        Inc(numRegistros);
        accounting    := qGenerica.FieldByName('Accounting').AsString;
        // Si es un abono la transacción es a la inversa
        if qGenerica.FieldByName('Tipo').AsString = 'H' then begin
          if qGenerica.FieldByName('Debit').AsFloat < 0 then begin
            credit       := Abs(qGenerica.FieldByName('Debit').AsFloat);
            debit        := 0;
          end else begin
            debit      := qGenerica.FieldByName('Debit').AsFloat;
          end;
        end;
        if qGenerica.FieldByName('Tipo').AsString = 'D' then begin
          if qGenerica.FieldByName('Credit').AsFloat < 0 then begin
            debit      := Abs(qGenerica.FieldByName('Credit').AsFloat);
            credit     := 0;
          end else begin
            credit       := qGenerica.FieldByName('Credit').AsFloat;
          end;
        end;
        taxGroup      := qGenerica.FieldByName('TaxGroup').AsString;
        officeGroup   := qGenerica.FieldByName('OfficeGroup').AsString;
        departament   := qGenerica.FieldByName('Department').AsString;
        division      := qGenerica.FieldByName('Division').AsString;
        branch        := qGenerica.FieldByName('Branch').AsString;
        sapProvider   := qGenerica.FieldByName('SAPProvider').AsString;
        commentary    := qGenerica.FieldByName('Commentary').AsString;
        contraAccount := ACCOUNT_DEFAULT;
        AddSAPInfo;

        qGenerica.Next;
        Inc(numLinea);
      end;
      if numRegistros > 0 then
        srvCfg.GrabarLog('Exportando ' + numRegistros.ToString + ' cobros a SAP (' + FormatDateTime('dd/mm/yyyy', srvCfg.fecha) + ')');
    except
      on E: Exception do begin
        raise
      end;
    end;
  finally
    FreeAndNil(qGenerica);
  end;
end;
{$endregion}

{$region 'EXPORTAR COBROS POR BANCO'}
procedure TJournalEntry.ExportarExcelCobrosPorBanco;
var body: TStringList;
begin
  if HayRegistrosParaExportar('PagoCli', 'FecPago') then begin
    body := TStringList.Create;
    CreateJsonObjects;
    try
      try
        ExportarLineasCobrosPorBanco;
        if hayRegistros then
          GeneraEnviaJson(TXT_TITULO_COBROS);
          if not srvCfg.envioOk then begin
            montarMensajeError(body);
            srvCfg.SendEmail('Error Envío Cobros a SAP (' + FormatDateTime('dd/mm/yyyy', srvCfg.fecha) + ')', body);
          end;
          if FileExists(srvCfg.nombreFichero) then
            DeleteFile(srvCfg.nombreFichero);
      except
        on E: Exception do begin
          raise
        end;
      end;
    finally
      body.Free;
      DestroyJsonObjects;
    end;
  end;
end;

procedure TJournalEntry.ExportarLineasCobrosPorBanco;
var qGenerica: TFDQuery; numRegistros: Integer;
begin
  qGenerica := TFDQuery.Create(nil);
  try
    qGenerica.Connection := srvCfg.conexionBaseDatos;
    try
      qGenerica.SQL.Add('DECLARE @PaymentTemp TABLE (Accounting VARCHAR(10),');
      qGenerica.SQL.Add('                            Tipo VARCHAR(1),');
      qGenerica.SQL.Add('							               Debit FLOAT,');
      qGenerica.SQL.Add('				               			 Credit FLOAT,');
      qGenerica.SQL.Add('							               TaxGroup VARCHAR(3),');
      qGenerica.SQL.Add('							               OfficeGroup VARCHAR(20),');
      qGenerica.SQL.Add('							               Department VARCHAR(30),');
      qGenerica.SQL.Add('							               Division VARCHAR(30),');
      qGenerica.SQL.Add('						                 Branch VARCHAR(4),');
      qGenerica.SQL.Add('						               	 SAPProvider VARCHAR(30),');
      qGenerica.SQL.Add('							               Commentary VARCHAR(500))');
      qGenerica.SQL.Add('DECLARE @Accounting VARCHAR(10),');
      qGenerica.SQL.Add('        @IdPagoCli INTEGER,');
      qGenerica.SQL.Add('        @Pagado FLOAT,');
      qGenerica.SQL.Add('	       @Diferencia FLOAT,');
      qGenerica.SQL.Add('		     @TieneConciliacion INTEGER,');
      qGenerica.SQL.Add('		     @DescCorta VARCHAR(5),');
      qGenerica.SQL.Add('	       @Delegacion VARCHAR(20),');
      qGenerica.SQL.Add('	       @Department VARCHAR(3),');
      qGenerica.SQL.Add('	       @Division VARCHAR(10),');
      qGenerica.SQL.Add('	       @SAPProvider VARCHAR(20),');
      qGenerica.SQL.Add('		     @DescFormaPago VARCHAR(500)');
      qGenerica.SQL.Add('');
      qGenerica.SQL.Add('SET @Department  = ''COB''');
      qGenerica.SQL.Add('SET @Division    = ''COB CON''');
      qGenerica.SQL.Add('SET @SAPProvider = ''P01003''');
      qGenerica.SQL.Add('');
      qGenerica.SQL.Add('DECLARE Cursor_Pagos CURSOR FOR');
      qGenerica.SQL.Add('	SELECT b.SCCodSCta,');
      qGenerica.SQL.Add('	       p.IdPagoCli,');
      qGenerica.SQL.Add('	       p.Pagado,');
      qGenerica.SQL.Add('		     (p.Pagado - ISNULL(SUM(dp.Pagado), 0)) Diferencia,');
      qGenerica.SQL.Add('		     ISNULL(dp.IdPagoCli, 0) TieneConciliacion,');
      qGenerica.SQL.Add('		     c.DescCorta,');
      qGenerica.SQL.Add('	       c.Delegacion,');
      qGenerica.SQL.Add('		     tp.Descripcio + '' '' +  ISNULL(c.DescCorta, '''') + '' ('' + ISNULL(b.Descripcio, '''') + '')'' Commentary');
      qGenerica.SQL.Add('   FROM PagoCli p');
      qGenerica.SQL.Add('        LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli');
      qGenerica.SQL.Add('  	     LEFT JOIN Centros c ON c.IdCentro = p.IdCentro');
      qGenerica.SQL.Add('		     LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago');
      qGenerica.SQL.Add('		     LEFT JOIN Bancos b ON b.IdBanco = p.IdBanco');
      qGenerica.SQL.Add('  WHERE p.FecPago = ' + QuotedStr(FormatDateTime('yyyymmdd', iif((srvCfg.fecha = 0), Date, srvCfg.fecha))));
      qGenerica.SQL.Add(' GROUP BY b.SCCodSCta, p.IdPagoCli, p.Pagado, dp.IdPagoCli, c.DescCorta, c.Delegacion, tp.Descripcio, tp.Descripcio + '' '' +  ISNULL(c.DescCorta, '''') + '' ('' + ISNULL(b.Descripcio, '''') + '')''');
      qGenerica.SQL.Add('	OPEN Cursor_Pagos');
      qGenerica.SQL.Add('	FETCH NEXT FROM Cursor_Pagos');
      qGenerica.SQL.Add('	INTO @Accounting, @IdPagoCli, @Pagado, @Diferencia, @TieneConciliacion, @DescCorta, @Delegacion, @DescFormaPago');
      qGenerica.SQL.Add('	WHILE @@FETCH_STATUS = 0');
      qGenerica.SQL.Add('	BEGIN');
      qGenerica.SQL.Add('		INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)');
      qGenerica.SQL.Add('		SELECT @Accounting Accounting,');
      qGenerica.SQL.Add('          ''H'' Tipo,');
      qGenerica.SQL.Add('          ISNULL(dp.Pagado, p.Pagado) Debit,');
      qGenerica.SQL.Add('          0 Credit,');
      qGenerica.SQL.Add('          '''' Taxgroup,');
      qGenerica.SQL.Add('          ISNULL(c.Delegacion, '''') OfficeGroup,');
      qGenerica.SQL.Add('          ISNULL(g.Descripcio, @Department) Departament,');
      qGenerica.SQL.Add('          ISNULL(omc.Descripcio, @Division) Division,');
      qGenerica.SQL.Add('          ISNULL(c.DescCorta, '''') Branch,');
      qGenerica.SQL.Add('          ISNULL(tc.CodInt, @SAPProvider) SAPProvider,');
      qGenerica.SQL.Add('          @DescFormaPago');
      qGenerica.SQL.Add('     FROM PagoCli p');
      qGenerica.SQL.Add('          LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli');
      qGenerica.SQL.Add('		  	    LEFT JOIN DeudaCli dc on dc.IdDeudaCli = dp.IdDeudaCli');
      qGenerica.SQL.Add('          LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago');
      qGenerica.SQL.Add('          LEFT JOIN TtosMed tt ON tt.Ident = dc.IdentTM');
      qGenerica.SQL.Add('          LEFT JOIN Centros c ON c.IdCentro = p.IdCentro');
      qGenerica.SQL.Add('          LEFT JOIN TColabos tc ON tt.IdCol = tc.IdCol');
      qGenerica.SQL.Add('          LEFT JOIN Tratamientos_tarifas t ON t.IdTratamientotarifa = tt.IdTto');
      qGenerica.SQL.Add('          LEFT JOIN Tratamientos ttm ON ttm.IdTratamiento = t.IdTratamiento');
      qGenerica.SQL.Add('          LEFT JOIN TGrupos g ON g.IdGrupo = ttm.IdGrupo');
      qGenerica.SQL.Add('          LEFT JOIN TEspecOMC omc ON omc.IdTipoEspec = ttm.IdTipoEspec');
      qGenerica.SQL.Add('    WHERE p.IdPagoCli = @IdPagoCli');
      qGenerica.SQL.Add('	  UNION ALL');
      qGenerica.SQL.Add('   SELECT ''103001'' Accounting,');
      qGenerica.SQL.Add('          ''D'' Tipo,');
      qGenerica.SQL.Add('          0 Debit,');
      qGenerica.SQL.Add('          ISNULL(dp.Pagado, p.Pagado) Credit,');
      qGenerica.SQL.Add('          '''' Taxgroup,');
      qGenerica.SQL.Add('          ISNULL(c.Delegacion, '''') OfficeGroup,');
      qGenerica.SQL.Add('          ISNULL(g.Descripcio, @Department) Departament,');
      qGenerica.SQL.Add('          ISNULL(omc.Descripcio, @Division) Division,');
      qGenerica.SQL.Add('          ISNULL(c.DescCorta, '''') Branch,');
      qGenerica.SQL.Add('          ISNULL(tc.CodInt, @SAPProvider) SAPProvider,');
      qGenerica.SQL.Add('          @DescFormaPago');
      qGenerica.SQL.Add('     FROM PagoCli p');
      qGenerica.SQL.Add('          LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli');
      qGenerica.SQL.Add('			     LEFT JOIN DeudaCli dc on dc.IdDeudaCli = dp.IdDeudaCli');
      qGenerica.SQL.Add('          LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago');
      qGenerica.SQL.Add('          LEFT JOIN TtosMed tt ON tt.Ident = dc.IdentTM');
      qGenerica.SQL.Add('          LEFT JOIN Centros c ON c.IdCentro = p.IdCentro');
      qGenerica.SQL.Add('          LEFT JOIN TColabos tc ON tt.IdCol = tc.IdCol');
      qGenerica.SQL.Add('          LEFT JOIN Tratamientos_tarifas t ON t.IdTratamientotarifa = tt.IdTto');
      qGenerica.SQL.Add('          LEFT JOIN Tratamientos ttm ON ttm.IdTratamiento = t.IdTratamiento');
      qGenerica.SQL.Add('          LEFT JOIN TGrupos g ON g.IdGrupo = ttm.IdGrupo');
      qGenerica.SQL.Add('          LEFT JOIN TEspecOMC omc ON omc.IdTipoEspec = ttm.IdTipoEspec');
      qGenerica.SQL.Add('    WHERE p.IdPagoCli = @IdPagoCli');
      qGenerica.SQL.Add('');
      qGenerica.SQL.Add('		IF ((@Diferencia > 0) AND (@TieneConciliacion > 0))');
      qGenerica.SQL.Add('		BEGIN');
      qGenerica.SQL.Add('  			INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)');
      qGenerica.SQL.Add('		  	VALUES (''102001'', ''H'', @Diferencia, 0, '''', @Delegacion, @Department, @Division, @DescCorta, @SAPProvider, @DescFormaPago)');
      qGenerica.SQL.Add('  			INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)');
      qGenerica.SQL.Add('			  VALUES (''103001'', ''D'', 0, @Diferencia, '''', @Delegacion, @Department, @Division, @DescCorta, @SAPProvider, @DescFormaPago)');
      qGenerica.SQL.Add('		END');
      qGenerica.SQL.Add('		FETCH NEXT FROM Cursor_Pagos');
      qGenerica.SQL.Add('		INTO @Accounting, @IdPagoCli, @Pagado, @Diferencia, @TieneConciliacion, @DescCorta, @Delegacion, @DescFormaPago');
      qGenerica.SQL.Add('	END');
      qGenerica.SQL.Add('CLOSE Cursor_Pagos');
      qGenerica.SQL.Add('DEALLOCATE Cursor_Pagos');
      qGenerica.SQL.Add('');
      qGenerica.SQL.Add('SELECT Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary');
      qGenerica.SQL.Add('  FROM @PaymentTemp');
      qGenerica.SQL.Add('ORDER BY Branch, Commentary');
      qGenerica.Open;


      {$region 'SQL SERVER'}
        {DECLARE @PaymentTemp TABLE (Accounting VARCHAR(10),
                                    Tipo VARCHAR(1),
                                    Debit FLOAT,
                                    Credit FLOAT,
                                    TaxGroup VARCHAR(3),
                                    OfficeGroup VARCHAR(20),
                                    Department VARCHAR(30),
                                    Division VARCHAR(30),
                                    Branch VARCHAR(4),
                                    SAPProvider VARCHAR(30),
                                    Commentary VARCHAR(500))
        DECLARE @Accounting VARCHAR(10),
                @IdPagoCli INTEGER,
                @Pagado FLOAT,
                @Diferencia FLOAT,
                @TieneConciliacion INTEGER,
                @DescCorta VARCHAR(5),
                @Delegacion VARCHAR(20),
                @DescFormaPago VARCHAR(500)

        DECLARE Cursor_Pagos CURSOR FOR
          SELECT b.SCCodSCta,
                 p.IdPagoCli,
                 p.Pagado,
                 (p.Pagado - ISNULL(SUM(dp.Pagado), 0)) Diferencia,
                 ISNULL(dp.IdPagoCli, 0) TieneConciliacion,
                 c.DescCorta,
                 c.Delegacion,
                 tp.Descripcio + ' ' +  ISNULL(c.DescCorta, '') + ' (' + ISNULL(b.Descripcio, '') + ')' Commentary
             FROM PagoCli p
                  LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli
                  LEFT JOIN Centros c ON c.IdCentro = p.IdCentro
                  LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago
                  LEFT JOIN Bancos b ON b.IdBanco = p.IdBanco
            WHERE p.FecPago = '20230329'
            GROUP BY b.SCCodSCta, p.IdPagoCli, p.Pagado, dp.IdPagoCli, c.DescCorta, c.Delegacion, tp.Descripcio, tp.Descripcio + ' ' +  ISNULL(c.DescCorta, '') + ' (' + ISNULL(b.Descripcio, '') + ')'

          OPEN Cursor_Pagos
          FETCH NEXT FROM Cursor_Pagos
          INTO @Accounting, @IdPagoCli, @Pagado, @Diferencia, @TieneConciliacion, @DescCorta, @Delegacion, @DescFormaPago

          WHILE @@FETCH_STATUS = 0
          BEGIN
            INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)
            SELECT (CASE WHEN p.IdForPago = 2  THEN '102001'
                         WHEN p.IdForPago = 5  THEN '102001'
                         WHEN p.IdForPago = 6  THEN '102001'
                         WHEN p.IdForPago = 7  THEN '102001'
                         WHEN p.IdForPago = 14 THEN '102001'
                         WHEN p.IdForPago = 17 THEN '102001'
                         WHEN p.IdForPago = 18 THEN '102001'
                         WHEN p.IdForPago = 19 THEN '102001'
                         WHEN p.IdForPago = 21 THEN '102001'
                         WHEN p.IdForPago = 22 THEN '102001'
                         WHEN p.IdForPago = 31 THEN '102001'
                         WHEN p.IdForPago = 32 THEN '102001'
                         WHEN p.IdForPago = 28 THEN '102003'
                         WHEN p.IdForPago = 29 THEN '102003'
                         WHEN p.IdForPago = 15 THEN '102004'
                         WHEN p.IdForPago = 16 THEN '102004'
                         WHEN p.IdForPago = 20 THEN '102004'
                         WHEN p.IdForPago = 23 THEN '102004'
                         WHEN p.IdForPago = 25 THEN '102004'
                         WHEN p.IdForPago = 24 THEN '102005'
                         WHEN p.IdForPago = 26 THEN '102005'
                         WHEN p.IdForPago = 27 THEN '102006'
                         ELSE '102001'
                   END) Accounting,
                  'H' Tipo,
                   ISNULL(dp.Pagado, p.Pagado) Debit,
                   0 Credit,
                   '' Taxgroup,
                   ISNULL(c.Delegacion, '') OfficeGroup,
                   ISNULL(g.Descripcio, 'COB') Departament,
                   ISNULL(omc.Descripcio, 'COBCON') Division,
                   ISNULL(c.DescCorta, '') Branch,
                   ISNULL(tc.CodInt, 'P01003') SAPProvider,
                   @DescFormaPago
              FROM PagoCli p
                   LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli
                   LEFT JOIN DeudaCli dc on dc.IdDeudaCli = dp.IdDeudaCli
                   LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago
                   LEFT JOIN TtosMed tt ON tt.Ident = dc.IdentTM
                   LEFT JOIN Centros c ON c.IdCentro = p.IdCentro
                   LEFT JOIN TColabos tc ON tt.IdCol = tc.IdCol
                   LEFT JOIN Tratamientos_tarifas t ON t.IdTratamientotarifa = tt.IdTto
                   LEFT JOIN Tratamientos ttm ON ttm.IdTratamiento = t.IdTratamiento
                   LEFT JOIN TGrupos g ON g.IdGrupo = ttm.IdGrupo
                   LEFT JOIN TEspecOMC omc ON omc.IdTipoEspec = ttm.IdTipoEspec
             WHERE p.IdPagoCli = @IdPagoCli
            UNION ALL
            SELECT '103001' Accounting,
                   'D' Tipo,
                   0 Debit,
                   ISNULL(dp.Pagado, p.Pagado) Credit,
                   '' Taxgroup,
                   ISNULL(c.Delegacion, '') OfficeGroup,
                   ISNULL(g.Descripcio, 'COB') Departament,
                   ISNULL(omc.Descripcio, 'COBCON') Division,
                   ISNULL(c.DescCorta, '') Branch,
                   ISNULL(tc.CodInt, 'P01003') SAPProvider,
                   @DescFormaPago
              FROM PagoCli p
                   LEFT JOIN DeudaPago dp ON dp.IdPagoCli = p.IdPagoCli
                   LEFT JOIN DeudaCli dc on dc.IdDeudaCli = dp.IdDeudaCli
                   LEFT JOIN TForPago tp ON tp.IdForPago = p.IdForPago
                   LEFT JOIN TtosMed tt ON tt.Ident = dc.IdentTM
                   LEFT JOIN Centros c ON c.IdCentro = p.IdCentro
                   LEFT JOIN TColabos tc ON tt.IdCol = tc.IdCol
                   LEFT JOIN Tratamientos_tarifas t ON t.IdTratamientotarifa = tt.IdTto
                   LEFT JOIN Tratamientos ttm ON ttm.IdTratamiento = t.IdTratamiento
                   LEFT JOIN TGrupos g ON g.IdGrupo = ttm.IdGrupo
                   LEFT JOIN TEspecOMC omc ON omc.IdTipoEspec = ttm.IdTipoEspec
             WHERE p.IdPagoCli = @IdPagoCli

            IF ((@Diferencia > 0) AND (@TieneConciliacion > 0))
            BEGIN
                INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)
              VALUES ('102001', 'H', @Diferencia, 0, '', @DescCorta,'COB', 'COBCON', @DescCorta, 'P01003', @DescFormaPago)
                INSERT INTO @PaymentTemp (Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary)
              VALUES ('103001', 'D', 0, @Diferencia, '', @DescCorta, 'COB', 'COBCON', @DescCorta, 'P01003', @DescFormaPago)
            END

            FETCH NEXT FROM Cursor_Pagos
            INTO @Accounting, @IdPagoCli, @Pagado, @Diferencia, @TieneConciliacion, @DescCorta, @Delegacion, @DescFormaPago
          END
        CLOSE Cursor_Pagos
        DEALLOCATE Cursor_Pagos

        SELECT Accounting, Tipo, Debit, Credit, TaxGroup, OfficeGroup, Department, Division, Branch, SAPProvider, Commentary
         FROM @PaymentTemp
        ORDER BY Branch, Commentary}
      {$endregion}
      numlinea := 0; numRegistros := 0;
      hayRegistros := not qGenerica.IsEmpty;
      while not qGenerica.Eof do begin
        Clear;

        Inc(numRegistros);
        accounting    := qGenerica.FieldByName('Accounting').AsString;
        // Si es un abono la transacción es a la inversa
        if qGenerica.FieldByName('Tipo').AsString = 'H' then begin
          if qGenerica.FieldByName('Debit').AsFloat < 0 then begin
            credit       := Abs(qGenerica.FieldByName('Debit').AsFloat);
            debit        := 0;
          end else begin
            debit      := qGenerica.FieldByName('Debit').AsFloat;
          end;
        end;
        if qGenerica.FieldByName('Tipo').AsString = 'D' then begin
          if qGenerica.FieldByName('Credit').AsFloat < 0 then begin
            debit      := Abs(qGenerica.FieldByName('Credit').AsFloat);
            credit     := 0;
          end else begin
            credit       := qGenerica.FieldByName('Credit').AsFloat;
          end;
        end;
        taxGroup      := qGenerica.FieldByName('TaxGroup').AsString;
        officeGroup   := qGenerica.FieldByName('OfficeGroup').AsString;
        departament   := qGenerica.FieldByName('Department').AsString;
        division      := qGenerica.FieldByName('Division').AsString;
        branch        := qGenerica.FieldByName('Branch').AsString;
        sapProvider   := qGenerica.FieldByName('SAPProvider').AsString;
        commentary    := qGenerica.FieldByName('Commentary').AsString;
        contraAccount := ACCOUNT_DEFAULT;
        AddSAPInfo;

        qGenerica.Next;
        Inc(numLinea);
      end;
      if numRegistros > 0 then
        srvCfg.GrabarLog('Exportando ' + numRegistros.ToString + ' cobros a SAP (' + FormatDateTime('dd/mm/yyyy', srvCfg.fecha) + ')');
    except
      on E: Exception do begin
        raise
      end;
    end;
  finally
    FreeAndNil(qGenerica);
  end;
end;
{$endregion}

function TJournalEntry.HayRegistrosParaExportar(nombreTabla, campoFecha: string): Boolean;
var qGenerica: TFDQuery;
begin
  Result := false;
  qGenerica := TFDQuery.Create(nil);
  try
    qGenerica.Connection := srvCfg.conexionBaseDatos;
    qGenerica.SQL.Add('SELECT COUNT(*) NumRegistros');
    qGenerica.SQL.Add('  FROM ' + nombreTabla);
    qGenerica.SQL.Add('  WHERE ' + campoFecha + ' = ' + QuotedStr(FormatDateTime('yyyymmdd', iif((srvCfg.fecha = 0), Date, srvCfg.fecha))));
    qGenerica.Open;
    Result := qGenerica.FieldByName('NumRegistros').Value > 0;
  finally
    qGenerica.Free;
  end;
end;

procedure TJournalEntry.CambiarCuentaSAP;
var qGenerica: TFDQuery;
begin
  qGenerica := TFDQuery.Create(nil);
  try
    qGenerica.Connection := srvCfg.conexionBaseDatos;
    qGenerica.SQL.Add('UPDATE Bancos SET SCCodSCta = CodBIC WHERE CodBIC <> ISNULL(SCCodSCta, '''')');
    qGenerica.ExecSQL;
  finally
    qGenerica.Free;
  end;
end;

procedure TJournalEntry.montarMensajeError(var lista: TStringList);
begin
  lista.Add('************************************************');
  lista.Add('********************  ERROR ********************');
  lista.Add('');
  lista.Add('StatusCode: ' + srvCfg.statusCode.ToString);
  lista.Add('Messsage: ' + srvCfg.error);
  lista.Add('');
  lista.Add('************************************************');
  lista.Add('**************  ATTACHMENT FILE ****************');
  lista.Add('');
  lista.Add(srvCfg.nombreFichero);
  lista.Add('');
  lista.Add('************************************************');
end;

end.


