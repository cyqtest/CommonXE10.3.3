unit USBKey;

interface
uses
  System.SysUtils, System.Classes,Windows, Messages, Forms, Registry;

type
 TFun = function(Ppsw: PAnsiChar; PFileName: PAnsiChar): boolean; stdcall;
   TAR_GetUSBSerialNumber = function(SerialNumber : PAnsiChar): boolean; stdcall;
   TAR_GetVendorSerialNumber = function(SerialNumber : PAnsiChar): boolean; stdcall;
   TAR_GetAlcorUFDCount = function(): integer; stdcall;
   TAR_OpenDiskByID = function(nID: integer): boolean; stdcall;
   TAR_GetFlashDiskDrive2K = function(CheckName1: PAnsiChar;CheckName2: PAnsiChar; KeyString: PAnsiChar; pDrive: PAnsiChar): boolean; stdcall;
   TAR_CheckLicenseCode = function(nCodeSize: integer; pLicenseCode: PAnsiChar): boolean; stdcall;
   TAR_CloseHandle = function(): boolean; stdcall;

   function getUsbDrive: string;
   function getUsbKey: string;

var
  USB_Code, USB_Msg: string;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

const
  NO_DLL_FILE: string = '缺少系统文件，请与软件商联系！';
  NO_USB_DRIVE: string = '请插入开票金税盘，谢谢！';

function getUsbDrive: string;
var
    buf:array [0..MAX_PATH-1] of char;
    m_Result:Integer;
    i:Integer;
    str_temp:string;
begin
  Result := '';
  m_Result:=GetLogicalDriveStrings(MAX_PATH,buf);
  for i:=0 to (m_Result div 4) do
  begin
    str_temp:=string(buf[i*4]+buf[i*4+1]+buf[i*4+2]);
    if GetDriveType(pchar(str_temp)) = DRIVE_REMOVABLE then
    begin
      Result := str_temp;
      if FileExists(str_temp+'AutoKey.dll') then
      begin
        Result := Copy(str_temp,1,1);
        Break;
      end;
    end;
  end;
end;

function getUsbKey: string;
var
  hInst: HWND;
  AR_GetUSBSerialNumber: TAR_GetUSBSerialNumber;
  AR_GetVendorSerialNumber: TAR_GetVendorSerialNumber;
  AR_GetAlcorUFDCount: TAR_GetAlcorUFDCount;
  AR_OpenDiskByID: TAR_OpenDiskByID;
  AR_CheckLicenseCode: TAR_CheckLicenseCode;
  AR_CloseHandle: TAR_CloseHandle;
  AR_GetFlashDiskDrive2K: TAR_GetFlashDiskDrive2K;
  iUSBCount, iPathKey, i: integer;
  arDisk, arKey :array[0..7] of AnsiChar;
  strPath: string;
  strKey: string;
begin
  USB_Msg := '';

  strPath := getUsbDrive;
  if strPath = '' then
  begin
    USB_Msg := NO_USB_DRIVE;
    result := '';
    Exit;
  end;
  // := LoadLibrary(PChar(ExtractFilePath(ParamStr(0))+'AutoKey.dll'));
  hInst := LoadLibrary(PChar(strPath+':\AutoKey.dll'));
  try
    if hInst = 0 then
    begin
      USB_Msg := NO_DLL_FILE;
      result := '';
      Exit;
    end;
    @AR_GetAlcorUFDCount := GetProcAddress(hInst,'AR_GetAlcorUFDCount');
    if not( @AR_GetAlcorUFDCount = nil ) then
    begin
        iUSBCount := AR_GetAlcorUFDCount();
        if iUSBCount = 0 then
        begin
          USB_Msg := NO_USB_DRIVE;
          result := '';
          Exit;
        end;
        @AR_GetFlashDiskDrive2K := GetProcAddress(hInst,'AR_GetFlashDiskDrive2K');
        if not (@AR_GetFlashDiskDrive2K = nil) then
          AR_GetFlashDiskDrive2K('usb', 'disk', '_8.0',PAnsiChar(@arDisk));
        for i := 0 to iUSBCount-1 do
        begin
          if arDisk[i] = strPath then
            iPathKey := i;
        end;
        @AR_OpenDiskByID := GetProcAddress(hInst,'AR_OpenDiskByID');
        if not( @AR_OpenDiskByID = nil ) then
        begin
          if not AR_OpenDiskByID(iPathKey) then
          begin
            USB_Msg := NO_USB_DRIVE;
            result := '';
            Exit;
          end;
        end
        else
        begin
          result := '';
          Exit;
        end;
        @AR_CheckLicenseCode := GetProcAddress(hInst,'AR_CheckLicenseCode');
        if not( @AR_CheckLicenseCode = nil ) then
        begin
          if not AR_CheckLicenseCode(0, nil) then
          begin
            USB_Msg := NO_USB_DRIVE;
            result := '';
            Exit;
          end;
        end
        else
        begin
          result := '';
          Exit;
        end;
        @AR_GetUSBSerialNumber := GetProcAddress(hInst,'AR_GetUSBSerialNumber');
        if not( @AR_GetUSBSerialNumber = nil ) then
        begin
          if not AR_GetUSBSerialNumber(PAnsiChar(@arKey)) then
          begin
            USB_Msg := NO_USB_DRIVE;
            result := '';
            Exit;
          end
          else
          begin
            @AR_CloseHandle := GetProcAddress(hInst,'AR_CloseHandle');
            AR_CloseHandle();
            sleep(200);
            strKey := arKey;
            result := trim(strKey);
          end;
        end;
    end
    else
    begin
      result := '';
    end;
  finally
    FreeLibrary(hInst);
  end;
end;

end.
