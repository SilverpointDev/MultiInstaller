program MultiInstaller;

uses
  Forms,
  Form.Installer in 'Form.Installer.pas' {FormInstall};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormInstall, FormInstall);
  Application.Run;
end.
