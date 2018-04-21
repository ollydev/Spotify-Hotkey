program SpotifyHotkey;

{$mode objfpc}{$H+}

uses
  Interfaces, Forms,
  main;

{$R *.res}

begin
  RequireDerivedFormResource := True;

  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.

