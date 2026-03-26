#include "installer_version.iss"

#define MyAppName "CONVERSOR - VEXPER"
#define MyAppPublisher "Vexper Sistemas"
#define MyAppURL "https://vexper.local"
#define MyAppExeName "CONVERSOR - VEXPER.exe"
#define MyAppSourceDir "dist\\CONVERSOR - VEXPER"

[Setup]
AppId={{A2D3AC9C-2B40-4D18-9E44-5B98D3EB0F11}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=Instalador do conversor de banco para Excel da Vexper
VersionInfoTextVersion={#MyAppVersion}
VersionInfoProductName={#MyAppName}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=logo.ico
WizardImageFile=installer_wizard.bmp
WizardSmallImageFile=installer_small.bmp
DisableDirPage=no
DisableProgramGroupPage=yes
ChangesAssociations=no
UsePreviousAppDir=yes
DisableWelcomePage=no
CloseApplications=yes
RestartApplications=no
CloseApplicationsFilter=*.exe
OutputDir=installer
OutputBaseFilename=Instalador CONVERSOR - VEXPER

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Messages]
brazilianportuguese.WelcomeLabel1=Bem-vindo ao instalador do CONVERSOR - VEXPER
brazilianportuguese.WelcomeLabel2=Este assistente instala o conversor da Vexper com suporte a leitura de banco Firebird e exportacao no modelo Excel esperado.
brazilianportuguese.FinishedHeadingLabel=Instalacao concluida
brazilianportuguese.FinishedLabel=O CONVERSOR - VEXPER foi instalado com sucesso neste computador.
brazilianportuguese.SelectDirLabel3=Escolha a pasta onde o sistema sera instalado.
brazilianportuguese.SelectTasksLabel2=Selecione os atalhos adicionais que deseja criar.

[CustomMessages]
brazilianportuguese.AppDescription=Conversor interno Vexper para leitura de banco e exportacao Excel padronizada.

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na area de trabalho"; GroupDescription: "Atalhos adicionais:"; Flags: unchecked

[InstallDelete]
Type: files; Name: "{app}\CONVERSOR - VEXPER atualizado.exe"
Type: files; Name: "{app}\*.dll"
Type: files; Name: "{app}\*.pyd"
Type: files; Name: "{app}\*.pkg"
Type: files; Name: "{app}\base_library.zip"
Type: files; Name: "{app}\python*.dll"
Type: files; Name: "{app}\VCRUNTIME*.dll"

[Files]
Source: "{#MyAppSourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Desinstalar {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
var
	InfoPage: TWizardPage;
	InfoTitle: TNewStaticText;
	InfoText: TNewStaticText;

procedure InitializeWizard;
begin
	InfoPage := CreateCustomPage(wpWelcome, 'Preparacao da instalacao', 'Resumo do que sera instalado neste computador');

	InfoTitle := TNewStaticText.Create(InfoPage);
	InfoTitle.Parent := InfoPage.Surface;
	InfoTitle.Left := ScaleX(0);
	InfoTitle.Top := ScaleY(6);
	InfoTitle.Width := InfoPage.SurfaceWidth;
	InfoTitle.Height := ScaleY(24);
	InfoTitle.Caption := 'CONVERSOR - VEXPER';
	InfoTitle.Font.Style := [fsBold];
	InfoTitle.Font.Size := 12;

	InfoText := TNewStaticText.Create(InfoPage);
	InfoText.Parent := InfoPage.Surface;
	InfoText.Left := ScaleX(0);
	InfoText.Top := ScaleY(38);
	InfoText.Width := InfoPage.SurfaceWidth;
	InfoText.Height := ScaleY(170);
	InfoText.AutoSize := False;
	InfoText.WordWrap := True;
	InfoText.Caption :=
		'Este instalador vai configurar o sistema completo da Vexper para leitura de banco Firebird e exportacao Excel padronizada.' + #13#10 + #13#10 +
		'- instala a versao atual do sistema' + #13#10 +
		'- substitui a versao anterior quando ela existir' + #13#10 +
		'- cria atalhos do aplicativo' + #13#10 +
		'- mantem suporte a atualizacao automatica' + #13#10 + #13#10 +
		'Ao finalizar, o sistema ja podera ser aberto normalmente. Em futuras atualizacoes, a nova versao podera ser instalada por cima da antiga.';
end;