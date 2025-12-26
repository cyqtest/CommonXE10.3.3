unit CYQWinService;

interface
uses
  Windows, Messages, SysUtils, Classes, Controls,StdCtrls, ShellAPI, Forms, WinSvc;

  function ServiceGetStatus(sMachine, sService: string ): DWord;
  function ServiceExists(sMachine, sService : string) : Boolean;
  function ServiceRunning(sMachine, sService : string ) : Boolean;
  function ServiceStopped(sMachine, sService : string ) : Boolean;

  function InstallService(ServiceName, DisplayName,Description: pchar; FileName:
    string): Boolean;
  function RunService(svr:String):Boolean;
  function StopService(AServName: string): Boolean;
  function UninstallService(ServiceName: PChar): Boolean;

implementation

//返回指定服务状态
function ServiceGetStatus(sMachine, sService: string ): DWord;
var
  //service control
  //manager handle
  schm,
  //service handle
  schs: SC_Handle;
  //service status
  ss: TServiceStatus;
  //current service status
  dwStat : DWord;
begin
  dwStat := 0;
  //connect to the service
  //control manager
  schm := OpenSCManager(PChar(sMachine), Nil, SC_MANAGER_CONNECT);
  //if successful...
  if(schm  > 0)then
  begin
    //open a handle to
    //the specified service
    schs := OpenService(schm, PChar(sService), SERVICE_QUERY_STATUS);
    //if successful...
    if(schs  > 0)then
    begin
      //retrieve the current status
      //of the specified service
      if(QueryServiceStatus(schs, ss))then
      begin
        dwStat := ss.dwCurrentState;
      end;
      //close service handle
      CloseServiceHandle(schs);
    end;

    // close service control
    // manager handle
    CloseServiceHandle(schm);
  end;

  Result := dwStat;
end;

//判断服务是否安装
function ServiceExists(sMachine, sService : string) : Boolean;
begin
  Result := 0 <> ServiceGetStatus(sMachine, sService);
end;

//服务是否运行
function ServiceRunning(sMachine, sService : string ) : Boolean;
begin
  Result := SERVICE_RUNNING = ServiceGetStatus(sMachine, sService );
end;

//服务是否停止
function ServiceStopped(sMachine, sService : string ) : Boolean;
begin
  Result := SERVICE_STOPPED = ServiceGetStatus(sMachine, sService );
end;

//安装服务
function InstallService(ServiceName, DisplayName,Description: pchar; FileName:
string): Boolean;
  function SetServiceDescription(SH: THandle; Desc: PChar): Bool;
  const
    SERVICE_CONFIG_DESCRIPTION: DWord = 1;
  var
    OSVersionInfo: TOSVersionInfo;
    ChangeServiceConfig2: function(hService: SC_HANDLE; dwInfoLevel: DWORD;
      lpInfo: Pointer): Bool; StdCall;
    LH: THandle;
  begin
    Result :=false;
    OSVersionInfo.dwOSVersionInfoSize := SizeOf(OSVersionInfo);
    GetVersionEx(OSVersionInfo);
    if (OSVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT) and //NT? 环境判断 ，可以去掉
      (OSVersionInfo.dwMajorVersion >= 5) then
    begin
      LH := GetModuleHandle(advapi32);
      Result := LH <> 0;
      if not Result then
        Exit;
      ChangeServiceConfig2 := GetProcAddress(LH, 'ChangeServiceConfig2A');
      Result := @ChangeServiceConfig2 <> nil;
      if not Result then
        Exit;
      Result := ChangeServiceConfig2(SH, SERVICE_CONFIG_DESCRIPTION, @Desc);
      {if Result then
        FreeLibrary(LH); }
    end;
  end;

const
  SERVICE_CONFIG_DESCRIPTION: DWord = 1;
var
  SCManager: SC_HANDLE;
  Service: SC_HANDLE;
  Args: pchar;
begin
  SCManager := OpenSCManager(nil, nil, SC_MANAGER_ALL_ACCESS);
  if SCManager = 0 then
      Exit;
  try
    Service := CreateService(SCManager, ServiceName, DisplayName,
    SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS or
    SERVICE_INTERACTIVE_PROCESS, SERVICE_AUTO_START, SERVICE_ERROR_IGNORE,
    pchar(FileName), nil, nil, nil, nil, nil);
    try
      SetServiceDescription(Service, Description);
    except

    end;
    Args := nil;
    StartService(Service, 0, Args);
    CloseServiceHandle(Service);
  finally
    CloseServiceHandle(SCManager);
  end;
end;

//启动某个服务；
function RunService(svr:String):Boolean;
var
  schService:SC_HANDLE;
  schSCManager:SC_HANDLE;
  ssStatus:TServiceStatus;
  Argv:PChar;
begin
  schSCManager:=OpenSCManager(nil,nil,SC_MANAGER_ALL_ACCESS);
  schService:=OpenService(schSCManager,Pchar(svr),SERVICE_ALL_ACCESS);
  result := True;
  try
    if StartService(schService,0,Argv) then
    begin
      while (QueryServiceStatus(schService,ssStatus)) do
      begin
        Sleep(500);
        Application.ProcessMessages;
        if ssStatus.dwCurrentState=SERVICE_START_PENDING then Sleep(500)
        else Break;
      end;//while
      if ssStatus.dwCurrentState=SERVICE_RUNNING then result := True
      else result := False;
    end
    else
    result := False;
  finally
    CloseServiceHandle(schService);
    CloseServiceHandle(schSCManager);
  end;
end;

function StopService(AServName: string): Boolean;
var
  SCManager, hService: SC_HANDLE;
  SvcStatus: TServiceStatus;
begin
  SCManager := OpenSCManager(nil, nil, SC_MANAGER_ALL_ACCESS);
  Result := SCManager <> 0;
  if Result then
  try
    hService := OpenService(SCManager, PChar(AServName), SERVICE_ALL_ACCESS);
    Result := hService <> 0;
    if Result then
    try
      //停止并卸载服务;
      Result := ControlService(hService, SERVICE_CONTROL_STOP, SvcStatus);
      //删除服务;
//      DeleteService(hService);
    finally
      CloseServiceHandle(hService);
    end;
  finally
    CloseServiceHandle(SCManager);
  end;
end;

//卸载服务
function UninstallService(ServiceName: PChar): Boolean;
var
  SCManager: SC_HANDLE;
  Service: SC_HANDLE;
  Status: TServiceStatus;
begin
  Result := False;
  SCManager := OpenSCManager(nil, nil, SC_MANAGER_ALL_ACCESS);
  if SCManager = 0 then
      Exit;
  try
      Service := OpenService(SCManager, ServiceName, SERVICE_ALL_ACCESS);
      ControlService(Service, SERVICE_CONTROL_STOP, Status);
      DeleteService(Service);
      CloseServiceHandle(Service);
      Result := True;
  finally
      CloseServiceHandle(SCManager);
  end;
end;

end.
 