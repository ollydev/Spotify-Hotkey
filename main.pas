unit main;

// To debug exceptions in a console compile without win32-gui-application (project options > config and target)

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs,
  Messages, JwaWindows, Audio;

type
  TMainForm = class(TForm)
  published
    procedure WMHotKey(var Message: TMessage); message WM_HOTKEY;
  public
    procedure CustomException(Sender: TObject; E: Exception);

    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  MainForm: TMainForm;

implementation

// acquired with Spy++
const
  SPOTIFY_CLASS = 'Chrome_WidgetWin_0';
  SPOTIFY_NEXT_WPARAM = $18000073;
  SPOTIFY_NEXT_LPARAM = $000E045E;
  SPOTIFY_BACK_WPARAM = $18000074;
  SPOTIFY_BACK_LPARAM = $000A0BD2;
  SPOTIFY_PAUSE_PLAY_WPARAM = $18000072;
  SPOTIFY_PAUSE_PLAY_LPARAM = $00130664;

const
  HOTKEY_NEXT = 1;
  HOTKEY_BACK = 2;
  HOTKEY_PAUSE_PLAY = 3;
  HOTKEY_VOLUME_UP = 5;
  HOTKEY_VOLUME_DOWN = 6;

function SendCommand(Handle: HWND; Hotkey: LPARAM): LongBool; stdcall;

  function GetWindowClass: WideString;
  begin
    SetLength(Result, 1024);
    SetLength(Result, GetClassNameW(Handle, @Result[1], Length(Result)));
  end;

  function IsSpotify(PID: UInt32): Boolean;
  var
    Handle: THandle;
    Path: array[0..1024] of Char;
  begin
    Handle := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ, False, PID);

    if (Handle > 0) then
    try
      if GetModuleFileNameEx(Handle, 0, Path, Length(Path)) > 0 then
        Exit(ExtractFileName(Path) = 'Spotify.exe');
    finally
      CloseHandle(Handle);
    end;

    Exit(False);
  end;

var
  PID: UInt32;
begin
  if (GetWindowClass() = SPOTIFY_CLASS) then
  begin
    GetWindowThreadProcessId(Handle, @PID);

    if IsSpotify(PID) then
    begin
      case Hotkey of
        HOTKEY_VOLUME_UP:
          begin
            ChangeVolume(PID, 0.033);

            Exit(False);
          end;

        HOTKEY_VOLUME_DOWN:
          begin
            ChangeVolume(PID, -0.033);

            Exit(False);
          end;

        HOTKEY_NEXT:
          SendMessage(Handle, WM_COMMAND, SPOTIFY_NEXT_WPARAM, SPOTIFY_NEXT_LPARAM);

        HOTKEY_BACK:
          SendMessage(Handle, WM_COMMAND, SPOTIFY_BACK_WPARAM, SPOTIFY_BACK_LPARAM);

        HOTKEY_PAUSE_PLAY:
          SendMessage(Handle, WM_COMMAND, SPOTIFY_PAUSE_PLAY_WPARAM, SPOTIFY_PAUSE_PLAY_LPARAM);
      end;
    end;
  end;

  Exit(True);
end;

procedure TMainForm.WMHotKey(var Message: TMessage);
begin
  EnumWindows(@SendCommand, Message.wParam);
end;

procedure TMainForm.CustomException(Sender: TObject; E: Exception);
begin
  { nothing }
end;

constructor TMainForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  RegisterHotKey(Handle, HOTKEY_NEXT, MOD_WIN or MOD_ALT, VK_RIGHT);
  RegisterHotKey(Handle, HOTKEY_BACK, MOD_WIN or MOD_ALT, VK_LEFT);
  RegisterHotKey(Handle, HOTKEY_PAUSE_PLAY, MOD_WIN or MOD_ALT, VK_NUMPAD0);
  RegisterHotKey(Handle, HOTKEY_VOLUME_UP, MOD_WIN or MOD_ALT, VK_UP);
  RegisterHotKey(Handle, HOTKEY_VOLUME_DOWN, MOD_WIN or MOD_ALT, VK_DOWN);
end;

destructor TMainForm.Destroy;
begin
  UnregisterHotKey(Handle, HOTKEY_NEXT);
  UnregisterHotKey(Handle, HOTKEY_BACK);
  UnregisterHotKey(Handle, HOTKEY_PAUSE_PLAY);
  UnregisterHotKey(Handle, HOTKEY_VOLUME_UP);
  UnregisterHotKey(Handle, HOTKEY_VOLUME_DOWN);

  inherited Destroy();
end;

{$R *.lfm}

initialization
  Application.ShowMainForm := False;
  Application.OnException := @MainForm.CustomException;

end.

