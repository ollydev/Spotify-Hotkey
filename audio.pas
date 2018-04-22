unit audio;

interface

uses
  Classes, SysUtils,
  Windows, ActiveX, ComObj;

type
  EDataFlow = type LongWord;
  ERole = EDataFlow;

const
  eRender = $00000000;
  eCapture = $00000001;
  eAll = $00000002;
  eConsole = $00000000;
  eMultimedia = $00000001;
  eCommunications = $00000002;

type
  IMMDevice = interface(IUnknown)
    ['{D666063F-1587-4E43-81F1-B948E807363F}']
    function Activate(const iid: TGUID; dwClsCtx: UINT; pActivationParams: PPropVariant; out ppInterface: IUnknown): HRESULT; stdcall;
    function GetId(ppstrId: PWChar): HRESULT; stdcall;
    function GetState(var pdwState: UINT): HRESULT; stdcall;
  end;

  IMMDeviceCollection = interface(IUnknown)
    ['{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}']
    function GetCount(var pcDevices: UINT): HRESULT; stdcall;
    function Item(nDevice: UINT; out ppDevice: IMMDevice): HRESULT; stdcall;
  end;

  IMMNotificationClient = interface(IUnknown)
    ['{7991EEC9-7E89-4D85-8390-6C703CEC60C0}']
  end;

const
  IMMDeviceEnumerator_Create: TGUID = '{BCDE0395-E52F-467C-8E3D-C4579291692E}';

type
  IMMDeviceEnumerator = interface(IUnknown)
    ['{A95664D2-9614-4F35-A746-DE8DB63617E6}']
    function EnumAudioEndpoints(dataFlow: EDataFlow; dwStateMask: DWORD; out ppDevices: IMMDeviceCollection): HRESULT; stdcall;
    function GetDefaultAudioEndpoint(dataFlow: EDataFlow; role: ERole; out ppEndpoint: IMMDevice): HRESULT; stdcall;
    function GetDevice(pwstrId: PWChar; out ppDevice: IMMDevice): HRESULT; stdcall;
    function RegisterEndpointNotificationCallback(var pClient: IMMNotificationClient): HRESULT; stdcall;
    function UnregisterEndpointNotificationCallback(var pClient: IMMNotificationClient): HRESULT; stdcall;
  end;

  IRemoteAudioSession = interface(IUnknown)
    ['{33969B1D-D06F-4281-B837-7EAAFD21A9C0}']
    function A: HRESULT; stdcall;
    function B: HRESULT; stdcall;
    function C: HRESULT; stdcall;
    function D: HRESULT; stdcall;
    function E: HRESULT; stdcall;
    function F: HRESULT; stdcall;
    function G: HRESULT; stdcall;
    function H: HRESULT; stdcall;
    function I: HRESULT; stdcall;
    function J: HRESULT; stdcall;
    function K: HRESULT; stdcall;
    function GetProcessID(out PID: UINT): HRESULT; stdcall;
  end;

  IAudioSessionQuerier = interface(IUnknown)
    ['{94BE9D30-53AC-4802-829C-F13E5AD34776}']
    function GetNumSessions(out NumSessions: UINT): HRESULT; stdcall;
    function QuerySession(Num: UINT; out Session: IUnknown): HRESULT; stdcall;
  end;

  IAudioSessionQuery = interface(IUnknown)
    ['{94BE9D30-53AC-4802-829C-F13E5AD34775}']
    function GetQueryInterface(out AudioQuerier: IAudioSessionQuerier): HRESULT; stdcall;
  end;

  IAudioSessionEvents = interface(IUnknown)
    ['{24918ACC-64B3-37C1-8CA9-74A66E9957A8}']
    function OnDisplayNameChanged(NewDisplayName: LPCWSTR; EventContext: PGuid): HRESULT; stdcall;
    function OnIconPathChanged(NewIconPath: LPCWSTR; EventContext: PGuid): HRESULT; stdcall;
    function OnSimpleVolumeChanged(NewVolume: Single; NewMute: LongBool;EventContext: PGuid): HRESULT; stdcall;
    function OnChannelVolumeChanged(ChannelCount: UINT; NewChannelArray: pSingle; ChangedChannel: UINT; EventContext: PGuid): HRESULT; stdcall;
    function OnGroupingParamChanged(NewGroupingParam, EventContext: PGuid): HRESULT; stdcall;
    function OnStateChanged(NewState: UINT): HRESULT; stdcall;
    function OnSessionDisconnected(DisconnectReason: UINT): HRESULT; stdcall;
  end;

  IAudioSessionControl = interface(IUnknown)
    ['{F4B1A599-7266-4319-A8CA-E70ACB11E8CD}']
    function GetState(out pRetVal: UINT): HRESULT; stdcall;
    function GetDisplayName(out Value: LPWSTR): HRESULT; stdcall;
    function SetDisplayName(Value: LPCWSTR; EventContext: PGuid): HRESULT; stdcall;
    function GetIconPath(out Value: LPWSTR): HRESULT; stdcall;
    function SetIconPath(Value: LPCWSTR; EventContext: PGuid): HRESULT; stdcall;
    function GetGroupingParam(Value: PGuid): HRESULT; stdcall;
    function SetGroupingParam(OverrideValue, EventContext: PGuid) : HRESULT; stdcall;
    function RegisterAudioSessionNotification(const NewNotifications: IAudioSessionEvents): HRESULT; stdcall;
    function UnregisterAudioSessionNotification(const NewNotifications: IAudioSessionEvents): HRESULT; stdcall;
  end;

  ISimpleAudioVolume = interface(IUnknown)
    ['{87CE5498-68D6-44E5-9215-6DA47EF883D8}']
    function SetMasterVolume(Level: Single; EventContext: PGUID): HRESULT; stdcall;
    function GetMasterVolume(out Level: Single): HRESULT; stdcall;
    function SetMute(Mute: LongBool; EventContext: PGuid): HRESULT; stdcall;
    function GetMute(out Mute: LongBool): HRESULT; stdcall;
  end;

  IAudioSessionManager = interface(IUnknown)
    ['{BFA971F1-4D5E-40BB-935E-967039BFBEE4}']
    function GetAudioSessionControl(AudioSessionGuid: PGuid; StreamFlag: UINT; out SessionControl: IAudioSessionControl): HRESULT; stdcall;
    function GetSimpleAudioVolume(AudioSessionGuid: PGuid; StreamFlag: UINT; out AudioVolume: ISimpleAudioVolume): HRESULT; stdcall;
  end;

  IAudioSessionEnumerator = interface(IUnknown)
    ['{E2F5BB11-0570-40CA-ACDD-3AA01277DEE8}']
    function GetCount(out SessionCount: Int32): HResult; stdcall;
    function GetSession(const SessionCount: Int32; out Session: IAudioSessionControl): HResult; stdcall;
  end;

  IAudioSessionManager2 = interface(IAudioSessionManager)
    ['{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}']
    function GetSessionEnumerator(out Enumerator: IAudioSessionEnumerator): HResult; stdcall;
  end;

  IAudioSessionControl2 = interface(IAudioSessionControl)
    ['{bfb7ff88-7239-4fc9-8fa2-07c950be9c6d}']
    function GetSessionIdentifier(out Value: LPWSTR): HResult; stdcall;
    function GetSessionInstanceIdentifier(out Value: LPWSTR): Hresult; stdcall;
    function GetProcessId(out PID: UInt32): HResult; stdcall;
    function IsSystemSoundsSession: HResult; stdcall;
    function SetDuckingPreference(const Value: BOOL): HResult; stdcall;
  end;

  IAudioMeterInformation = interface(IUnknown)
    ['{C02216F6-8C67-4B5B-9D00-D008E73E0064}']
    function GetPeakValue(out Peak: Single): HResult; stdcall;
    function GetMeteringChannelCount(out ChannelCount: UINT): HResult; stdcall;
    function GetChannelsPeakValues(const ChannelCount: UINT; out Values: array of Single): HResult; stdcall;
    function QueryHardwareSupport(out HardwareSupportMask: PDWORD): HResult; stdcall;
  end;

function GetSessionControl(PID: UInt32): IAudioSessionControl;
function GetSimpleAudio(PID: UInt32): ISimpleAudioVolume;
function ChangeVolume(PID: UInt32; Delta: Single): Boolean;

implementation

function GetSessionControl(PID: UInt32): IAudioSessionControl;
var
  device: IMMDevice;
  deviceEnumerator: IMMDeviceEnumerator;
  audioSessionManager2: IAudioSessionManager2;
  autoSessionEnumerator: IAudioSessionEnumerator;
  session, sessionCount: Int32;
  sessionControl: IAudioSessionControl;
  audioSessionControl2: IAudioSessionControl2;
  ID: UInt32;
begin
  if FAILED(CoCreateInstance(IMMDeviceEnumerator_Create, nil, CLSCTX_INPROC_SERVER, IMMDeviceEnumerator, deviceEnumerator)) then
    raise Exception.Create('FAILED getting deviceEnumerator');

  if FAILED(deviceEnumerator.GetDefaultAudioEndpoint(eRender, eMultimedia, device)) then
    raise Exception.Create('FAILED getting device');

  if FAILED(device.Activate(IAudioSessionManager2, CLSCTX_ALL, nil, IUnknown(audioSessionManager2))) then
    raise Exception.Create('FAILED getting audioSessionManager2');

  if FAILED(audioSessionManager2.GetSessionEnumerator(autoSessionEnumerator)) then
    raise Exception.Create('FAILED getting autoSessionEnumerator');

  if FAILED(autoSessionEnumerator.GetCount(sessionCount)) then
    raise Exception.Create('FAILED getting sessionCount');

  for session := 0 to sessionCount - 1 do
  begin
    if FAILED(autoSessionEnumerator.GetSession(session, sessionControl)) then
      raise Exception.Create('FAILED getting sessionControl[' + IntToStr(session) + ']');

    if FAILED(sessionControl.QueryInterface(IAudioSessionControl2, audioSessionControl2)) then
      raise Exception.Create('FAILED getting audioSessionControl2[' + IntToStr(session) + ']');

    if FAILED(audioSessionControl2.GetProcessID(ID)) then
      raise Exception.Create('FAILED getting sessionPID[' + IntToStr(session) + ']');;

    if (ID = PID) then
      Exit(audioSessionControl2);
  end;

  Exit(nil);
end;

function GetSimpleAudio(PID: UInt32): ISimpleAudioVolume;
var
  sessionControl: IAudioSessionControl;
  simpleAudio: ISimpleAudioVolume;
begin
  sessionControl := GetSessionControl(PID);

  if (sessionControl <> nil) then
  begin
    if FAILED(sessionControl.QueryInterface(ISimpleAudioVolume, simpleAudio)) then
      raise Exception.Create('FAILED getting simpleAudio');

    Exit(simpleAudio);
  end;

  Exit(nil);
end;

function ChangeVolume(PID: UInt32; Delta: Single): Boolean;
var
  simpleAudio: ISimpleAudioVolume;
  volumeLevel: Single;
begin
  try
    simpleAudio := GetSimpleAudio(PID);

    if (simpleAudio <> nil) then
    begin
      if FAILED(simpleAudio.GetMasterVolume(volumeLevel)) then
        raise Exception.Create('FAILED getting volumeLevel');

      if (volumeLevel + Delta < 0) then
      begin
        if (volumeLevel = 0) then
          Exit(True);

        volumeLevel := 0;
      end else
      if (volumeLevel + Delta > 1) then
      begin
        if (volumeLevel = 1) then
          Exit(True);

        volumeLevel := 1;
      end else
        volumeLevel += Delta;

      if FAILED(simpleAudio.SetMasterVolume(volumeLevel, nil)) then
        raise Exception.Create('FAILED setting volumeLevel');

      Exit(True);
    end;
  except
  end;

  Exit(False);
end;

initialization
  CoInitialize(nil);

finalization
  CoUninitialize();

end.
