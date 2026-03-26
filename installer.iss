#include "installer_version.iss"

#define MyAppName "CONVERSOR - VEXPER"
#define MyAppPublisher "Vexper Sistemas"
#define MyAppURL "https://vexper.local"
#define MyAppExeName "CONVERSOR - VEXPER atualizado.exe"
#define MyAppSourceExe "dist\\CONVERSOR - VEXPER atualizado.exe"

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

[Files]
Source: "{#MyAppSourceExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Desinstalar {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent