unit LoginForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore, dxSkinCaramel,
  cxGroupBox, Vcl.Menus, Vcl.StdCtrls, cxButtons;

type
  TfLogin = class(TForm)
    cxGroupBox1: TcxGroupBox;
    Label12: TLabel;
    cxPassword: TEdit;
    cxButton1: TcxButton;
    cxButton2: TcxButton;
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fLogin: TfLogin;

implementation

{$R *.dfm}

procedure TfLogin.cxButton1Click(Sender: TObject);
begin
  if (cxPassword.Text <> '15041959') then begin
    //ShowMessage('”казан неверный пароль.');
    MessageDLG('”казан неверный пароль.',mtError,[mbOk],0);
    exit;
  end;

  ModalResult := mrOk;
end;

procedure TfLogin.cxButton2Click(Sender: TObject);
begin
  cxButton2.ModalResult := mrCancel;
end;

procedure TfLogin.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Ord(Key) = VK_ESCAPE then begin
    key:=#0;
    ModalResult := mrCancel;
  end;
  if Ord(Key) = VK_RETURN then begin
    key:=#0;
    cxButton1Click(Sender);
  end;
end;

end.
