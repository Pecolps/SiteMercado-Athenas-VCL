program IntegradorSA;

uses
  Vcl.Forms,
  UPrincipal in 'UPrincipal.pas' {frmPrincipal};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := False;
  Application.ShowMainForm:=       False;
  Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.Run;
end.
