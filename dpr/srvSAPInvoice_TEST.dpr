program srvSAPInvoice_TEST;

uses
  Vcl.Forms,
  FRM in '..\source\FRM.pas' {frmPrincipal},
  edFuncionesConexion in '..\source\edFuncionesConexion.pas',
  edFuncionesJournal in '..\source\edFuncionesJournal.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.Run;
end.
