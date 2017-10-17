program DevExpressExample;

uses
  Vcl.Forms,
  MainForm in 'MainForm.pas' {ShuhartForm},
  AboutForm in 'AboutForm.pas' {frAbout},
  WordClass in 'WordClass.pas',
  LoginForm in 'LoginForm.pas' {fLogin};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TShuhartForm, ShuhartForm);
  Application.CreateForm(TfrAbout, frAbout);
  Application.CreateForm(TfLogin, fLogin);
  Application.Run;
end.
