{*********************************************************************}
{                                                                     }
{     CYQCommonUnit v1.0  Create By cyq 2013年5月10日11:03:18         }
{                                                                     }
{                                                                     }
{     单元功能：常用公共函数、事件                                    }
{                                                                     }
{*********************************************************************}
unit CYQCommonUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Graphics, Dialogs, Variants,
  StdCtrls, Forms, IniFiles, DB, GridsEh, DBGridEh, PrnDbgeh, Printers, Grids,
  DBGrids, Excel2000, ComObj, ShellAPI, System.Contnrs, System.StrUtils, ShlObj,
  Masks, FtPing, ADODB, DateUtils, WinSock, TypInfo, DBGridEhImpExp, ExtCtrls,
  TlHelp32, ColorGrd, DBgridEhToExcel, FireDAC.Comp.DataSet, FireDAC.Comp.Client,
  System.ZLib, Winapi.Nb30, System.RegularExpressions, System.Zip;
  //GridsEh, DBGridEh, Grids, DBGrids位置不能调换
{$IF CompilerVersion >= 21.0} //Delphi2010
  {$Define Delphi2010_Up}
{$IFEND}

const
  NoVersionInfoStr = 'Not Found Version Information!';
  DIALOGFILE_EXT_XLS = 'Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx;|All files (*.*)|*.*';

  //以下参数用于流导出Excel
var
  arXlsBegin: array[0..5] of Word = ($809, 8, 0, $10, 0, 0);
  arXlsEnd: array[0..1] of Word = ($0A, 00);
  arXlsString: array[0..5] of Word = ($204, 0, 0, 0, 0, 0);
  arXlsNumber: array[0..4] of Word = ($203, 14, 0, 0, 0);
  arXlsInteger: array[0..4] of Word = ($27E, 10, 0, 0, 0);
  arXlsBlank: array[0..4] of Word = ($201, 6, 0, 0, $17);
  oldMode: Cardinal;

  //以下参数用于UCommonMsg
var
  MyMsg_IconIndex: string;
  MyMsg_ButtonResult: Boolean;

  //以下参数用于OpenWithShellExecute 打开指定文件
type
  TOpenFileType = (vFTFileName, vFTURL, vFTDirectory, vFTPrint);

  //获得当前程序目录Ini文件名
  function GetIniFile: string; overload;

  function GetIniFile(FileName: string): string; overload;

    //创建Ini文件
  function CreateIniFile(var IniFile: TIniFile; FileIni: string): Boolean;

  //获取用户当前登录设备的id
  function GetCurrentDevicesID: string;
    //去除StringList的重复项目

  procedure RemoveDuplicatesList(const aStrList: TStringList);

  procedure GetGrdEhColList(DBGridEh: TDBGridEh; TitleList: array of string; var FList: TStringList);
    //通用DBGridEh导出

  procedure DBGridEhToExcel(DBGridEh: TWinControl; TitleList: array of string; SaveToFile: Boolean; DocName: string = '重命名');

  procedure GridToExcel(Grid: TWinControl; TitleList: array of string; SaveToFile: Boolean = False; UseTree: Boolean = True); overload;

  procedure DataSetEhToList(qry: TDataSet; FieldList: array of string; var xList, StrColList: TStringList; DBGridEh: TWinControl = nil); overload;

  procedure DataSetEhToList(qry: TDataSet; FieldList: array of string; var xList: TStringList; DBGridEh: TWinControl = nil); overload;

  procedure CheckValidDBGridEh(DBGridEh: TWinControl);

  function IsValidDBGridEh(DBGridEh: TWinControl): Boolean;
    //从FTitleList中解析出多表头到字符串，以#9分隔列，以#13#10分隔行

  function GetExcelTitleStr(var FTitleList: TStringList; UseMultiTitle: Boolean): string;

  procedure GetDBGridCheckBoxList(DBGrid: TWinControl; var CheckBoxColList: TStringList);

  function IsCheckBoxCol(FCol: TCollection; FieldName: string; CheckBoxColList: TStringList): Boolean;

    //把tsList导出到 Excel 文件（不需要系统安装Excel）
    //如果不指定文件名 AFileName，则自动弹出另存为的对话框。
  procedure WriteListToExcelFile(xList: TStringList; AFileName: string = '');
    //把tsList导出到 Excel

  procedure WriteListToExcel(tsList: TStringList; StrColList: TStrings = nil);

    //支持多表头导出DBgridEh数据;
  procedure ExportExcelEh(const CurDS: TDBGridEh; vSum: Boolean = False; Title: string = '');
    //导出DbGrid数据

  procedure ExportExcelDbGrid(ADbgrid: TDBGrid);
    //以流的方式导出Excel，适合大批量数据导出

  procedure ExportExcelByStream(DBGridEh: TDBGridEh; bWriteTitle: Boolean; Title: string = ''; bWriteFirst: Boolean = False; TitleFirst: string = '');
    //以流的方式导出Excel，适合大批量数据导出(支持多表头)

  procedure ExportExcelByStreamMut(DBGridEh: TDBGridEh; Title: string = ''; ShowPross: Boolean = True);

  procedure ExportDBGridEhToCSV(DBGrid: TDBGridEh; const FileName: string);

    //打印DBgridEh表格数据
  procedure PrintDBGridEhData(DBGridEh: TWinControl; Title: string);
    //检查DBGridEh某一列是否存在重复值

  function CheckGridFieldDuplicate(DBGridEh: TWinControl; FieldName: array of string; EMsg: string = ''): Boolean;
    //关闭窗体DataSet，在CloseForm时调用

  procedure CloseAllDataSet(frm: TForm);
    //向ComboBox里添加值

  procedure AddValueToCombobox(qryTemp: TADOQuery; strInputSQL: string; Combox: TCustomComboBox; bShowAll: Boolean = False); overload;

  procedure AddValueToCombobox(qryTemp: TFDQuery; strInputSQL: string; Combox: TCustomComboBox; bShowAll: Boolean = False); overload;

  procedure AddValueToComboboxByID(qryTemp: TADOQuery; strInputSQL: string; Combox: TCustomComboBox; bShowAll: Boolean = False); overload;

  procedure AddValueToComboboxByID(qryTemp: TFDQuery; strInputSQL: string; Combox: TCustomComboBox; bShowAll: Boolean = False); overload;

  function GetComboboxID(Combox: TCustomComboBox): string;

  function LocateComboboxByID(Combox: TCustomComboBox; ID: string): Boolean;
    //小写数转大写数字

  function CashSmallToBig(small: Real): string;
    //打开一个窗体，非子窗体

  procedure OpenNormalForm(AForm: TCustomForm; const AFormClass: TFormClass);

  procedure OpenNormalForm_zhs(AForm: TCustomForm; const AFormClass: TFormClass);
    //打开一个子窗体：FormClass:窗体类 Fm:窗体名 AOwner:所有者(SELF)

  procedure OpenChildForm(FormClass: TFormClass; var Fm; AOwner: TComponent); overload

  procedure DontCareIme(Parent: TWinControl);

  procedure OpenChildForm(FormClass: TCustomFormClass; var Form; bDontCareIme: Boolean = True); overload
    //执行指定路径Exe

  function ExcuteExe(Exe_Dir: string): Boolean;

    //查找指定进程
  function FindProcess(AFileName: string): boolean;
    //结束指定进程

  procedure KillProcess(AFileName: string);
    //根据进程名称获取程序句柄

  function GetProcessHandleByName(Name: string): THandle;

    //弹出一个提示框，提示框包括成功、错误、询问、提示
  function ShowMsbInfo(AMsg: string; IType: string = 'info'): Boolean;
    //弹出一个提示框，提示框包括成功、错误、询问、提示

  function ShowMessageBoxEx(AMsg: string; IType: string = 'info'): Boolean; overload

  function ShowMessageBoxEx(AMsg: string; Flags: Longint): Boolean; overload

  procedure HideTaskbar; //隐藏任务栏

  procedure ShowTaskbar; //显示任务栏

  procedure DisableTask(Key: Boolean); //禁止任务管理器
    //显示数字在一个框里传入画布，显示框大小，显示字符，显示颜色

  procedure ShowDigiInRect(Canvas: TCanvas; mRect: TRect; str: string; Color: TColor);
    //打开指定数据集

  function OpenDataSet(qry: TADOQuery; sqlTxt: string): Boolean; overload

  function OpenDataSet(qry: TFDQuery; sqlTxt: string): Boolean; overload

  //异步打开指定数据集
  function OpenDataSetSync(qry: TADOQuery; sqlTxt: string): Boolean; overload
  function OpenDataSetSync(qry: TFDQuery; sqlTxt: string): Boolean; overload

    //执行指定SQL语句

  function ExecuteSQL(qry: TADOQuery; sqlTxt: string): Boolean; overload

  function ExecuteSQL(qry: TFDQuery; sqlTxt: string): Boolean; overload

  function ExecuteSQLNoTran(qry: TADOQuery; sqlTxt: string): Boolean; overload

  function ExecuteSQLNoTran(qry: TFDQuery; sqlTxt: string): Boolean; overload

    //批量执行SQL

  function ExecuteSQLBatch(qry: TADOQuery; lvListSQL: TStringList; execOne: Boolean = False): Boolean; overload

  function ExecuteSQLBatch(qry: TFDQuery; lvListSQL: TStringList; execOne: Boolean = False): Boolean; overload

    //初始化日志

  procedure IniSystemLog(const logPath: string = '');
    /// <summary>
    /// logType : 0 运行日志; 1：错误日志；2：提示日志；3：警告日志
    /// </summary>

  procedure WriteSystemLog(sLogText: string; logType: Integer = 0);

  function StartDate_DD(Date: TDateTime): string;  //获得日期开始时间 精确到Date

  function EndDate_DD(Date: TDateTime): string;    //获得日期结束时间 精确到Date

  function StartDate_ZZZ(Date: TDateTime): string; //获得日期开始时间  精确到千分之一

  function EndDate_ZZZ(Date: TDateTime): string;   //获得日期结束时间  精确到千分之一

  function StartDate_SS(Date: TDateTime): string; //获得日期开始时间  精确到秒

  function EndDate_SS(Date: TDateTime): string; //获得日期结束时间  精确到秒

  function GetServerDateTime(qry: TADOQuery): TDateTime; overload; //获得服务器时间  精确到千分之三,即数据库时间

  function GetServerDateTime(qry: TFDQuery): TDateTime; overload; //获得服务器时间  精确到千分之三,即数据库时间

  function GetMonthOfDate(ReturnType: Integer; incYear: Integer = 1): TDateTime; overload;//获得月份相关的天数

  function GetMonthOfDate(InputDate: TDateTime; ReturnType: Integer): TDateTime; overload;

  function myDateTimeToUnix(ConvDate: TDateTime; len: Integer = 10): Int64;  //当前时间转为时间戳 考虑到本地时间和格林尼治时间的差

  function myUnixToDateTime(USec: Int64): TDateTime;      //时间戳转为当前时间 考虑到本地时间和格林尼治时间的差

  function ConvertStrToDateTime(DateTimeStr: string; var ADateTime: TDateTime): boolean;//把时间字符串yyyyMMddhhmmss转成时间

  procedure RefreshQuery(qry: TADOQuery); //刷新ADO数据集

  procedure RefreshDataSet(DataSet: TDataSet);//刷新数据集，刷新后返回到以前状态(记录位置)

  procedure RefreshDataSetEX(DataSet: TDataSet; bDisableControls: Boolean);

  procedure RefreshMemTable(DataSet: TDataSet);//刷新数据集MemTable(EH)，刷新后返回到以前状态(记录位置)

  procedure SetSysDateFormat;                   //设定本程序使用时间格式

  procedure IntegerOnly(var Key: Char);  //限制输入整数

  procedure FloatOnly(var Key: Char; iText: string);   //限制输入浮点数

  procedure NumberOrLetterOnly(var Key: Char);

  function IsNumber(str: string): Boolean; //判断字符是否数字

  function IsFloat(str: string): Boolean; //判断字符是否浮点数

  function IsNumeric(Value: string; const AllowFloat: Boolean; const TrimWhiteSpace: Boolean = True): Boolean;  //判断字符是否浮点数

  function FormatFloatStr(value: string; DisplayFormat: string = '#0.00'):string;

    //检查控件上面CheckBox勾选，用于查询

  function IsCheckNull(Container: TWinControl; DefaultCheckBox: TCheckBox = nil): Boolean;
    //设置控件上面Edit控件Enable属性

  function EditEnable(Container: TWinControl; Able: Boolean; DefaultEdit: TEdit = nil): Boolean;

    //字符转成Utf8编码
  function DecodeUtf8Str(const S: UTF8String): WideString;
    //字符转换成十六进制

  function myStrToHex(s: string): string;
    //十六进制字符转普通字符

  function myHexToStr(s: string): string;
    //查找某一字符串在源字符串中出现的次数

  function SubStrCount(MainStr, SubStr: string): Integer;
    //通过分隔符获取若干子串

  function SplitStr(const Source, Split: string): TStringList;
    //将字符串中的指定子串替换成特定子串，返回一个新字符串

  function SubStrReplace(SourceStr, OldStr, NewStr: string): string;
    //将字符串中的指定子串替换成特定子串，返回一个新字符串（秒杀Memo的回车，换行，空格特殊字符）

  procedure ReplaceMeoToStr(var s: string; const SourceChar: pchar; const RChar: pchar);
    //删除StringList里面重复的项目

  procedure DelStrListDuplicates(const AStringList: TStringList);
    //获得TstringList中没有重复的字符串

  procedure GetUniqueStringList(const AStringList: TStringList);

  // 将数组按指定大小分组
  // arrTotal: 输入的字符串数组
  // groupSize: 每组的大小
  // 返回: TArray<TArray<string>>，包含分组后的数组
  function SplitStringArray(const arrTotal: TArray<string>; groupSize: Integer): TArray<TArray<string>>;
    //获取SQL错误中文信息
  function GetSQLErrorChineseInfo(ErrorInfo: string): string;
    //合计DataSet的指定字段

  function FieldSumValue(DataSet: TDataSet; FieldName: string): Extended; overload

  function FieldSumValue(DataSet: TDataSet; arrField: TArray<string>): TArray<Extended>; overload
    /// <summary>
    /// 指定字段，计算指定条件
    /// </summary>

  function FieldSumValueEx(DataSet: TDataSet; FieldName: string; ValuedName, Value: string; calc: char): Extended; overload

  function FieldSumValueEx(DataSet: TDataSet; FieldName: string; Value: string; calc: Char): Extended; overload
    //显示表格指定（非数字型）列的合计值

  function FieldSumValueFooter(grd: TDBGridEh; FieldName: string; DisplayFormat: string = '#0.00'): string; overload

  function FieldSumValueFooter(grd: TDBGridEh; arrField: TArray<string>; DisplayFormat: string = '#0.00'): string; overload
    //往指定字段Footer赋值

  function FieldFooterValue(grd: TDBGridEh; FieldName: string; FieldValue: string = '无'; DisplayFormat: string = '#0.00'): string;
    //设置指定GrdFooter列值

  function SetGrdFooterValue(grd: TDBGridEh; FieldName: string; FieldValue: string = ''): Boolean;
    //把控件默认值保存到ini文件

  procedure SaveIni(FormName: string; Cmb: TWinControl);

  procedure LoadIni(FormName: string; Cmb: TWinControl);

  function GetLocalComputerIP: string;   //获得本机IP

  function GetLocalComputerName: string; //获得本机名

  function GetWinLogonName: string;      //获得Windows登录用户名

  function GetMacAddress: string;        //获取电脑的mac地址

    //过程功能：显示加载信息

  procedure ShowLodingMessage(Mes: string; ShowTime: Boolean = False; BackgroundColorB: TColor = clWhite; BackgroundColorE: TColor = clWhite; aFontColor: TColor = $00444444);
    //过程功能：隐藏加载信息

  procedure HideLodingMessage;

  function GetWindowsVertion: string;

  function GetFileVersion(FileName: PChar): string;

  function GetFilesTime(sFilename: string; Timetype: Integer): TDateTime;

    //返回XXXX年XX月XX日 星期X格式
  function ReturnToday(CurrTime: TDateTime; bDay: Boolean = True): string;
    //返回硬盘信息

  function GetDisksInfo(strL: TStringList): TStringList;
    //判断操作系统是否win64
    {$IFDEF Delphi2010_Up}

  function IsWin64: Bool;
    {$ENDIF}
    //数字转大写

  function UpperCaseMoney(Decimal: string): string;
    //一段汉字转化成拼音

  function ChineseToLetter(const S: widestring): widestring;

  function IsEnCase(strTemp: string): Boolean; //判断字符是否英文
    //把Excel样式的英文字母序数转换为整数，26进制数且Z = A0

  function AlphaToInt(Value: string): Integer;
    //把非负整数转换成Excel样式的英文字母序数

  function IntToAlpha(Value: Integer): string;

    //遍历指定目录下文件。返回TStringList
  procedure EnumFileInQueue(path: PChar; fileExt: string; fileList: TStringList);
    //转换文件时间

  function CovFileDate(Fd: _FileTime): TDateTime;
    //获取文件时间，Tf表示目标文件路径和名称

  function GetFileTime(const Tf: string): string;
    //删除指定文件夹

  function DeleteFolder(const FolderPath: string): Boolean;

  function DeleteDirectory(NowPath: string): Boolean;
    //复制文件夹到指定新目标

  function CopyAndMoveDirectory(SourceDirName: string; TargetDirName: string): Boolean;
    //查找指定目录下的文件夹

  function GetSubDirectoryList(pPath: string): TStringList;

  function GetDirectoryFileList(pPath: string): TStringList;
    //获取指定文件夹目录下所有文件名。

  procedure GetFileListEx(FilePath, ExtMask: string; FileList: TStrings; SubDirectory: Boolean = True; addPath: Boolean = True);
    //获取指定文件夹目录下所有文件名

  function MakeFileList(Path: string; FileExt: string = ''): TStringList;
    //获取文件名，不带扩展名

  function ExtractFileNameNoExt(FileString: string): string;
    //弹出一个选择文件夹路径的对话框，可新建文件夹

  function SelectFolderDialog(const Handle: integer; const Caption: string; var SelectedFolder: string; const InitFolder: WideString = ' '): boolean;

  //打开指定的文件或位置、打印指定文件
  function OpenWithShellExecute(const targetPath: string; openType: TOpenFileType = vFTFileName): Boolean;
    //字符串加密
  function EncryptionStr(Src: string; Key: string = 'WhComTec'): string;
    //解密字符串

  function DecryptStr(Src: string; Key: string = 'WhComTec'): string;

    //压缩一个指定的文件
  function ZipAFile(sOldFile, sNewFile: string): Boolean;
    //解压一个指定的文件

  function UnZipAFile(sOldFile, sNewFile: string): Boolean;

  //压缩多个文件
  function CompressFiles(const ZipFileName: string; const Files: TArray<string>): Boolean;

    //提升程序权限
  procedure SetPrivilege;

    //快速Ping指定主机
  function PingRemoteMachine(aRemoteIP: string): Boolean;

    //等待函数(保持界面正常响应) MSecs为毫秒
  procedure WaitTime(MSecs: Integer);

  procedure WaitTimeEx(MaxWaitTime: Cardinal);

type
  TCharSegmentSet = set of 0{1}..7;

  TCharSegment = TCharSegmentSet;

var
  sOpenWindowClassStr: string;  //用于记录OpenWindow函数打开的窗体类名，防止一个窗体(Form1)在创建时，
                                //同时打开另一个相同窗体(Form1_1)，导致Form1_1检测权限不成功。

const
  HEAP_ZERO_MEMORY = $00000008;
  SID_REVISION     = 1; // Current revision level

type
  PTokenUser = ^TTokenUser;
  TTokenUser = packed record
    User: TSidAndAttributes;
  end;

implementation

uses
  Registry, Math, LodingForm, UCommonMsg, uLoadstate, qlog, FireDAC.Stan.Option;

//当前程序根目录Ini文件
function GetIniFile: string;
begin
  Result := ExtractFilePath(ParamStr(0));
  if DirectoryExists(Result) then
    Result := Result + 'Setting.ini'
  else
    Result := '';
end;

//当前程序根目录Ini文件
function GetIniFile(FileName: string): string;
begin
  Result := ExtractFilePath(Application.ExeName);
  if DirectoryExists(Result) then
    Result := Result + FileName
  else
    Result := '';
end;

function CreateIniFile(var IniFile: TIniFile; FileIni: string): Boolean;
begin
  Result := Assigned(IniFile);
  if Result then
    Exit;
  Result := (FileIni <> '');
  if Result then
  begin
    try
      IniFile := TIniFile.Create(FileIni);
    except
      Result := False;
    end;
  end;
end;

function ConvertSid(Sid: PSID; pszSidText: PChar; var dwBufferLen: DWORD): BOOL;
var
  psia: PSIDIdentifierAuthority;
  dwSubAuthorities: DWORD;
  dwSidRev: DWORD;
  dwCounter: DWORD;
  dwSidSize: DWORD;
begin
  Result := False;

  dwSidRev := SID_REVISION;

  if not IsValidSid(Sid) then Exit;

  psia := GetSidIdentifierAuthority(Sid);

  dwSubAuthorities := GetSidSubAuthorityCount(Sid)^;

  dwSidSize := (15 + 12 + (12 * dwSubAuthorities) + 1) * SizeOf(Char);

  if (dwBufferLen < dwSidSize) then
  begin
    dwBufferLen := dwSidSize;
    SetLastError(ERROR_INSUFFICIENT_BUFFER);
    Exit;
  end;

  StrFmt(pszSidText, 'S-%u-', [dwSidRev]);

  if (psia.Value[0] <> 0) or (psia.Value[1] <> 0) then
    StrFmt(pszSidText + StrLen(pszSidText),
      '0x%.2x%.2x%.2x%.2x%.2x%.2x',
      [psia.Value[0], psia.Value[1], psia.Value[2],
      psia.Value[3], psia.Value[4], psia.Value[5]])
  else
    StrFmt(pszSidText + StrLen(pszSidText),
      '%u',
      [DWORD(psia.Value[5]) +
      DWORD(psia.Value[4] shl 8) +
      DWORD(psia.Value[3] shl 16) +
      DWORD(psia.Value[2] shl 24)]);

  dwSidSize := StrLen(pszSidText);

  for dwCounter := 0 to (dwSubAuthorities-1) do
  begin
    StrFmt(pszSidText + dwSidSize, '-%u',
      [GetSidSubAuthority(Sid, dwCounter)^]);
    dwSidSize := StrLen(pszSidText);
  end;

  Result := True;
end;

function ObtainTextSid(hToken: THandle; pszSid: PChar;
  var dwBufferLen: DWORD): BOOL;
var
  dwReturnLength: DWORD;
  dwTokenUserLength: DWORD;
  tic: TTokenInformationClass;
  ptu: Pointer;
begin
  Result := False;
  dwReturnLength := 0;
  dwTokenUserLength := 0;
  tic := TokenUser;
  ptu := nil;

  if not GetTokenInformation(hToken, tic, ptu, dwTokenUserLength,
    dwReturnLength)
  then
  begin
    if GetLastError = ERROR_INSUFFICIENT_BUFFER then
    begin
      ptu := HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, dwReturnLength);
      if ptu = nil then Exit;
      dwTokenUserLength := dwReturnLength;
      dwReturnLength    := 0;

      if not GetTokenInformation(hToken, tic, ptu, dwTokenUserLength,
        dwReturnLength) then Exit;
    end else
      Exit;
  end;

  if not ConvertSid((PTokenUser(ptu).User).Sid, pszSid, dwBufferLen) then
    Exit;

  if not HeapFree(GetProcessHeap, 0, ptu) then
    Exit;

  Result := True;
end;


function GetCurrentDevicesID: string;
var
  hAccessToken: THandle;
  bSuccess: BOOL;
  dwBufferLen: DWORD;
  szSid: array[0..260] of Char;
begin
  Result := '';

  bSuccess := OpenThreadToken(GetCurrentThread, TOKEN_QUERY, True,
    hAccessToken);
  if not bSuccess then
  begin
    if GetLastError = ERROR_NO_TOKEN then
      bSuccess := OpenProcessToken(GetCurrentProcess, TOKEN_QUERY,
        hAccessToken);
  end;
  if bSuccess then
  begin
    ZeroMemory(@szSid, SizeOf(szSid));
    dwBufferLen := SizeOf(szSid);

    if ObtainTextSid(hAccessToken, szSid, dwBufferLen) then
      Result := szSid;
    CloseHandle(hAccessToken);
  end;
end;

procedure RemoveDuplicatesList(const aStrList: TStringList);
var
  strLTemp: TStringList;
  i: Integer;
  sTemp: string;
begin
  if not Assigned(aStrList) then
    Exit;

  strLTemp := TStringList.Create;
  try
    strLTemp.BeginUpdate;
    for i := 0 to Pred(aStrList.Count) do
    begin
      sTemp := aStrList.Strings[i];
      if strLTemp.IndexOf(sTemp) < 0 then
        strLTemp.Add(sTemp);
    end;
    strLTemp.EndUpdate;
    aStrList.Assign(strLTemp);
  finally
    FreeandNil(strLTemp);
  end;
end;

procedure CloseAllDataSet(frm: TForm);
var
  i: Integer;
begin
  if not Assigned(frm) then
    Exit;
  with frm do
    for i := 0 to ComponentCount - 1 do
      if (Components[i] is TCustomADODataSet) then
        if (Components[i] as TCustomADODataSet).Active then
          (Components[i] as TCustomADODataSet).Close;
end;

function EditEnable(Container: TWinControl; Able: Boolean; DefaultEdit: TEdit = nil): Boolean;
var
  AControl: TControl;
  I: Integer;
begin
  Result := False;
  with Container do
  begin
    for I := 0 to ControlCount - 1 do
    begin
      AControl := Controls[I];
      if AControl is TEdit then
      begin
        TEdit(AControl).Enabled := Able;
        Result := Result or TEdit(AControl).Enabled;
      end;
    end;
  end;
end;

//导出DBGrid数据// Added by CYQ 2013-08-15 10:04:49
procedure ExportExcelDbGrid(ADbgrid: tdbgrid);
var
  excel: variant;
  WorkBook: variant;
  WorkSheet: variant;
  SaveDialog1: tsavedialog;
  i, j: integer;
  CurDir: string;
  Pnl: TPanel;
begin
  try
{   ADbgrid;
   ADbgrid.DataSource.DataSet;
}
    with ADbgrid.DataSource.DataSet do
      if (Bof and Eof) then
        Exit;
    if (ADbgrid.DataSource.DataSet.state = dsedit) or (ADbgrid.DataSource.DataSet.state = dsinsert) then
    begin
      ShowMessage('表格处理编辑或者新增状态！');
      Exit;
    end;
    try
      excel := CreateOleObject('Excel.Application');
    except
      ShowMessage('Excel启动失败或者该电脑没有安装Excel，请检查！');
      Abort;
      Exit;
    end;

    SaveDialog1 := TSaveDialog.Create(nil);
    SaveDialog1.DefaulText := 'xls';
    SaveDialog1.Filter := DIALOGFILE_EXT_XLS;
    GetDir(0, CurDir);
    SaveDialog1.InitialDir := CurDir;

    if SaveDialog1.Execute then
    begin
      if FileExists(SaveDialog1.FileName) then
      begin
        if MessageDlg('Excel已经存在，是否覆盖?', MtConfirmation, [MBNO, MBYES], 0) = MRNO then
          Exit;
        excel.WorkBooks.Open(SaveDialog1.Filename);
      end
      else
        excel.WorkBooks.Add(1);
    end
    else
    begin
      ShowMessage('请指定要保存的文件名！');
      Exit;
    end;
    WorkBook := excel.Application.WorkBooks[1];
    WorkSheet := WorkBook.WorkSheets[1];

    //显示Panel
    Pnl := TPanel.Create(nil);
    Pnl.Parent := ADbgrid;
    Pnl.Caption := '数据正在导出，请稍候。。。';
    Pnl.Width := 220;
    Pnl.Alignment := taCenter;
    Pnl.Left := ADbgrid.Left + Floor(ADbgrid.Width / 2) - Floor(Pnl.Width / 4);
    Pnl.Top := ADbgrid.Top + Floor(ADbgrid.Height / 4) - Pnl.Height;
    Pnl.Font.Name := '黑体';
    Pnl.Font.Color := clRed;
    Pnl.Color := $00B871FF;
    Pnl.Show;

    for i := 0 to ADbgrid.columns.count - 1 do
      WorkSheet.Cells.Item[1, i + 1] := ADbgrid.Columns[i].Title.Caption;

    j := 2;
    with ADbgrid.DataSource.DataSet do
    begin
      DisableControls;
      First;
      while not eof do
      begin
        for i := 0 to ADbgrid.columns.count - 1 do
        begin
  //             ShowMessage(ADbgrid.columns[i].Fieldname);
          WorkSheet.Cells.Item[j, i + 1] := Trim(Fieldbyname(ADbgrid.Columns[i].Fieldname).AsString);
        end;
        next;
        j := j + 1;
      end;
      EnableControls;
    end;
    WorkBook.Saveas(SaveDialog1.Filename);
    excel.application.quit;
    //ShowMessage('成功导出到文件 : ' + SaveDialog1.Filename);
    Pnl.Caption := '导出完成！';

    SaveDialog1.Free;
  except
    SaveDialog1.Free;
    ADbgrid.DataSource.DataSet.EnableControls;
    excel.Application.Quit;
    WorkSheet.Free;
    WorkBook.Free;
    excel.Free;
    ShowMessage('保存失败，请检查文件是否在打开状态，关闭后重试!');
  end;
end;

procedure ExportExcelByStream(DBGridEh: TDBGridEh; bWriteTitle: Boolean; Title: string = ''; bWriteFirst: Boolean = False; TitleFirst: string = '');
var
  i, j: integer;
  Col, row: word;
  ABookMark: TBookMark;
  aFileStream: TFileStream;
  SaveDialog: TSaveDialog;
  FileName: string;

  procedure incColRow; //增加行列号
  begin
    if Col = DBGridEh.FieldCount - 1 then
    begin
      Inc(row);
      Col := 0;
    end
    else
      Inc(Col);
  end;

  procedure WriteStringCell(AValue: AnsiString);//写字符串数据
  var
    L: Word;
  begin
    L := Length(AValue);
    arXlsString[1] := 8 + L;
    arXlsString[2] := row;
    arXlsString[3] := Col;
    arXlsString[5] := L;
    aFileStream.WriteBuffer(arXlsString, SizeOf(arXlsString));
    aFileStream.WriteBuffer(Pointer(AValue)^, L);
    IncColRow;
  end;

  procedure WriteIntegerCell(AValue: integer);//写整数
  var
    V: Integer;
  begin
    arXlsInteger[2] := row;
    arXlsInteger[3] := Col;
    aFileStream.WriteBuffer(arXlsInteger, SizeOf(arXlsInteger));
    V := (AValue shl 2) or 2;
    aFileStream.WriteBuffer(V, 4);
    IncColRow;
  end;

  procedure WriteFloatCell(AValue: double);//写浮点数
  begin
    arXlsNumber[2] := row;
    arXlsNumber[3] := Col;
    aFileStream.WriteBuffer(arXlsNumber, SizeOf(arXlsNumber));
    aFileStream.WriteBuffer(AValue, 8);
    IncColRow;
  end;

begin
  DBGridEh.datasource.DataSet.DisableControls;
  SaveDialog := TSaveDialog.Create(nil);
  try
    SaveDialog.Filter := DIALOGFILE_EXT_XLS;
    SaveDialog.Filename := Title + '.xls';
    if SaveDialog.Execute then
    begin
      FileName := SaveDialog.Filename;
      if FileName = '' then
        FileName := '导出Excel';
    end
    else
    begin
      DBGridEh.datasource.DataSet.EnableControls;
      Exit;
    end;
  finally
    SaveDialog.Free;
  end;

  if FileExists(FileName) then
    DeleteFile(FileName); //文件存在，先删除
  aFileStream := TFileStream.Create(FileName, fmCreate);
  try    //写文件头 　
    ShowLodingMessage('正在导出数据到Excel...', True);
    aFileStream.WriteBuffer(arXlsBegin, SizeOf(arXlsBegin));   //写列头　

    if bWriteFirst then
    begin
      Col := 0;
      row := 0;
      WriteStringCell(AnsiString(TitleFirst));
      Col := 0;
      row := 1;
    end
    else
    begin
      Col := 0;
      row := 0;
    end;

    if bWriteTitle then
    begin
      for i := 0 to DBGridEh.FieldCount - 1 do
        WriteStringCell(AnsiString(DBGridEh.Columns[i].Title.Caption));
    end;
    //写数据集中的数据
    ABookMark := DBGridEh.datasource.DataSet.GetBookmark;
    DBGridEh.datasource.DataSet.First;
    while not DBGridEh.datasource.DataSet.Eof do
    begin
      for i := 0 to DBGridEh.FieldCount - 1 do
        case DBGridEh.Fields[i].DataType of
          ftSmallint, ftInteger, ftWord, ftAutoInc, ftBytes:
            WriteIntegerCell(DBGridEh.Fields[i].AsInteger);
          ftFloat, ftCurrency, ftBCD:
            WriteFloatCell(DBGridEh.Fields[i].AsFloat);
        else
          WriteStringCell(AnsiString(DBGridEh.Fields[i].AsString));
        end;
      DBGridEh.datasource.DataSet.Next;
    end;
    //写文件尾 　
    aFileStream.WriteBuffer(arXlsEnd, SizeOf(arXlsEnd));
    if DBGridEh.datasource.DataSet.BookmarkValid(ABookMark) then
      DBGridEh.datasource.DataSet.GotoBookmark(ABookMark);

    ShellExecute(0, 'Open', PChar(FileName), nil, nil, SW_SHOW); //打开Excel文件
  finally
    HideLodingMessage;
    aFileStream.Free;
    DBGridEh.datasource.DataSet.EnableControls;
  end;
end;

procedure ExportExcelByStreamMut(DBGridEh: TDBGridEh; Title: string = ''; ShowPross: Boolean = True);
var
  DBGridEhToExcel: TDBGridEhToExcel;
begin
  DBGridEhToExcel := TDBGridEhToExcel.Create(nil);
  try
    DBGridEhToExcel.TitleName := Title;
    DBGridEhToExcel.DBGridEh := DBGridEh;
    DBGridEhToExcel.ShowProgress := ShowPross;
    DBGridEhToExcel.ExportToExcel;
  finally
    DBGridEhToExcel.Free;
  end;
end;

procedure ExportDBGridEhToCSV(DBGrid: TDBGridEh; const FileName: string);
const
  UTF8_BOM: array[0..2] of Byte = ($EF, $BB, $BF);
var
  FS: TFileStream;
  Buffer: UTF8String;
  Line: string;
  i: Integer;
  Field: TField;
begin
  if not Assigned(DBGrid) or not Assigned(DBGrid.DataSource) or not Assigned(DBGrid.DataSource.DataSet) then
    Exit;

  var DataSet := DBGrid.DataSource.DataSet;

  if not DataSet.Active then Exit;

  FS := TFileStream.Create(FileName, fmCreate);
  // 写入 UTF-8 BOM（EF BB BF）
  FS.WriteBuffer(UTF8_BOM, SizeOf(UTF8_BOM));
  try
    DataSet.DisableControls;
    DataSet.First;

    // 写标题（列的Title.Caption）
    Line := '';
    for i := 0 to DBGrid.Columns.Count - 1 do
    begin
      if DBGrid.Columns[i].Visible then
      begin
        Line := Line + '"' + DBGrid.Columns[i].Title.Caption + '"';
        if i < DBGrid.Columns.Count - 1 then
          Line := Line + ',';
      end;
    end;
    Line := Line + sLineBreak;
    Buffer := UTF8String(Line);
    FS.WriteBuffer(Pointer(Buffer)^, Length(Buffer));

    // 写数据
    while not DataSet.Eof do
    begin
      Line := '';
      for i := 0 to DBGrid.Columns.Count - 1 do
      begin
        if DBGrid.Columns[i].Visible then
        begin
          Field := DBGrid.Columns[i].Field;
          if Assigned(Field) then
            Line := Line + '"' + StringReplace(Field.AsString, '"', '""', [rfReplaceAll]) + '"';
          if i < DBGrid.Columns.Count - 1 then
            Line := Line + ',';
        end;
      end;
      Line := Line + sLineBreak;
      Buffer := UTF8String(Line);
      FS.WriteBuffer(Pointer(Buffer)^, Length(Buffer));
      DataSet.Next;
    end;

  finally
    DataSet.EnableControls;
    FS.Free;
  end;
end;


procedure PrintDBGridEhData(DBGridEh: TWinControl; Title: string);
var
  PrintDialog: TPrintDialog;
  RowHeight, Temp_X, Temp_Y, PageEdgeX, PageEdgeY, PixelsPerInchX, PixelsPerInchY, PageCount: integer;
  TempStr: string;
  Scale: Double;
  Rect: TRect;
  //1.输出标题

  procedure Print_Title;
  begin
    Rect := Bounds(0, 0, Printer.PageWidth, PageEdgeY);
    Printer.Canvas.Font.Name := '楷体_GB2312';
    Printer.Canvas.Font.Style := Printer.Canvas.Font.Style + [fsBold];
    Printer.Canvas.Font.Size := 20;
    DrawText(Printer.Canvas.Handle, PChar(Title), Length(Title), Rect, DT_CENTER or DT_VCENTER or DT_SINGLELINE);
    Printer.Canvas.Rectangle(PageEdgeX, PageEdgeY, Printer.PageWidth - PageEdgeX, PageEdgeY + 1);
  end;
  //2.输出列头

  procedure Print_Column;
  var
    j: integer;
  begin
    Printer.Canvas.Font.Name := '黑体';
    Printer.Canvas.Font.Size := 9;
    Temp_X := PageEdgeX;
    Temp_Y := PageEdgeY;

    for j := 1 to (DBGridEh as TDBGridEh).Columns.Count do
    begin
      if not (DBGridEh as TDBGridEh).Columns[j - 1].Visible then
        Continue;

      TempStr := (DBGridEh as TDBGridEh).Columns[j - 1].Title.Caption;
      Rect := Bounds(Temp_X, Temp_Y, Trunc((DBGridEh as TDBGridEh).Columns[j - 1].Width * Scale), RowHeight);
      case (DBGridEh as TDBGridEh).Columns[j - 1].Field.Alignment of
        taCenter:
          DrawText(Printer.Canvas.Handle, PChar(TempStr), Length(TempStr), Rect, DT_CENTER or DT_VCENTER or DT_SINGLELINE); //case.1.居中
        taLeftJustify:
          DrawText(Printer.Canvas.Handle, PChar(TempStr), Length(TempStr), Rect, DT_LEFT or DT_VCENTER or DT_SINGLELINE); //case.2.居左
        taRightJustify:  //case.3.居右
          if Rect.Right - Rect.Left >= Printer.Canvas.TextWidth(TempStr) then
            DrawText(Printer.Canvas.Handle, PChar(TempStr), Length(TempStr), Rect, DT_RIGHT or DT_VCENTER or DT_SINGLELINE)
          else
            DrawText(Printer.Canvas.Handle, PChar(TempStr), Length(TempStr), Rect, DT_LEFT or DT_VCENTER or DT_SINGLELINE);
      end;
      Temp_X := Temp_X + Trunc((DBGridEh as TDBGridEh).Columns[j - 1].Width * Scale);
    end;
    Temp_Y := Temp_Y + RowHeight;
  end;
  //3.输出DBGrid内容

  procedure Print_Cells;
  var
    j: integer;
  begin
    Printer.Canvas.Font.Name := '宋体';
    Printer.Canvas.Font.Style := Printer.Canvas.Font.Style - [fsBold];
    while Temp_Y < Printer.PageHeight - PageEdgeY do
    begin
      Temp_X := PageEdgeX;
      for j := 1 to (DBGridEh as TDBGridEh).Columns.Count do
      begin
        if not (DBGridEh as TDBGridEh).Columns[j - 1].Visible then
          Continue;

        Rect := Bounds(Temp_X, Temp_Y, Trunc((DBGridEh as TDBGridEh).Columns[j - 1].Width * Scale), RowHeight);
        if ((DBGridEh as TDBGridEh).Columns[j - 1].Field is TCurrencyField) or ((DBGridEh as TDBGridEh).Columns[j - 1].Field is TLargeIntField) or ((DBGridEh as TDBGridEh).Columns[j - 1].Field is TSmallIntField) or ((DBGridEh as TDBGridEh).Columns[j - 1].Field is TIntegerField) or ((DBGridEh as TDBGridEh).Columns[j - 1].Field is TFloatField) or ((DBGridEh as TDBGridEh).Columns[j - 1].Field is TWordField) then
          TempStr := FormatFloat(',##0.00', (DBGridEh as TDBGridEh).Columns[j - 1].Field.AsFloat)
        else
          TempStr := (DBGridEh as TDBGridEh).Columns[j - 1].Field.AsString;
        case (DBGridEh as TDBGridEh).Columns[j - 1].Field.Alignment of
          taCenter:
            DrawText(Printer.Canvas.Handle, PChar(TempStr), Length(TempStr), Rect, DT_CENTER or DT_VCENTER or DT_SINGLELINE); //case.1.居中
          taLeftJustify:
            DrawText(Printer.Canvas.Handle, PChar(TempStr), Length(TempStr), Rect, DT_LEFT or DT_VCENTER or DT_SINGLELINE); //case.2.居左
          taRightJustify:
            DrawText(Printer.Canvas.Handle, PChar(TempStr), Length(TempStr), Rect, DT_RIGHT or DT_VCENTER or DT_SINGLELINE); //case.3.居右
        end;
        Temp_X := Temp_X + Trunc((DBGridEh as TDBGridEh).Columns[j - 1].Width * Scale);
      end;
      Temp_Y := Temp_Y + RowHeight;
      (DBGridEh as TDBGridEh).DataSource.DataSet.Next;
      if (DBGridEh as TDBGridEh).DataSource.DataSet.Eof then
        Exit;
    end;
  end;
  //4.输出页脚

  procedure Print_Footer;
  begin
    Temp_Y := Printer.PageHeight - PageEdgeY + RowHeight;
    Printer.Canvas.Rectangle(PageEdgeX, Temp_Y, Printer.PageWidth - PageEdgeX, Temp_Y + 1); //4.0.输出横线
    Rect := Bounds(PageEdgeX, Temp_Y, Printer.PageWidth - PageEdgeX, RowHeight); //4.1.输出日期
    DrawText(Printer.Canvas.Handle, PChar(DateTimeToStr(Now)), Length(DateTimeToStr(Now)), Rect, DT_LEFT or DT_VCENTER or DT_SINGLELINE);
    //4.2.输出页号
    Rect := Bounds(PageEdgeX, Temp_Y, Printer.PageWidth - PageEdgeX * 2, RowHeight);
    DrawText(Printer.Canvas.Handle, PChar('第' + IntToStr(Printer.PageNumber) + '页'), Length('第' + IntToStr(Printer.PageNumber) + '页'), Rect, DT_RIGHT or DT_VCENTER or DT_SINGLELINE);
  end;

begin
  CheckValidDBGridEh(DBGridEh);

  (DBGridEh as TDBGridEh).DataSource.DataSet.DisableControls;

  PrintDialog := TPrintDialog.Create(DBGridEh);
  try
    if PrintDialog.Execute then
    begin
      //0.取当前打印机X,Y方向每英寸像素
      PixelsPerInchX := GetDeviceCaps(Printer.Handle, LOGPIXELSX);
      PixelsPerInchY := GetDeviceCaps(Printer.Handle, LOGPIXELSY);
      //1.变量初始化
      PageEdgeX := PixelsPerInchX div 6;
      PageEdgeY := PixelsPerInchY div 5 * 4;
      RowHeight := Trunc(1.5 * 9 * PixelsPerInchY / 72);
      Scale := PixelsPerInchX / Screen.PixelsPerInch;

      try
        Printer.BeginDoc;
        ShowLodingMessage('正在打印数据,请稍候。。。');
        (DBGridEh as TDBGridEh).DataSource.DataSet.First;
        PageCount := 0;
        while not (DBGridEh as TDBGridEh).DataSource.DataSet.Eof do
        begin
          Print_Title;
          Print_Column;
          Print_Cells;
          Print_Footer;
          Printer.NewPage;
        end;
        if not Printer.Aborted then
          Printer.EndDoc;
      except
        on E: EPrinter do
        begin
          MessageBox(Application.Handle, PChar('打印机没有准备好！'), '提示！', MB_OK + MB_ICONINFORMATION);
          Printer.Abort;
          Exit;
        end;
      end;
    end;
  finally
    (DBGridEh as TDBGridEh).DataSource.DataSet.EnableControls;
    PrintDialog.Free;
    HideLodingMessage;
  end;
end;

//检查DBGridEh指定列的数值是否有重复
function CheckGridFieldDuplicate(DBGridEh: TWinControl; FieldName: array of string; EMsg: string = ''): Boolean;
var
  i: Integer;
  strlTemp: TStringList;
  strTemp: string;
begin
  Result := False;
  CheckValidDBGridEh(DBGridEh);
  if (DBGridEh as TDBGridEh).DataSource.DataSet.IsEmpty then
    Exit;
  (DBGridEh as TDBGridEh).DataSource.DataSet.DisableControls;
  strlTemp := TStringList.Create;
  try
    (DBGridEh as TDBGridEh).DataSource.DataSet.First;
    while not (DBGridEh as TDBGridEh).DataSource.DataSet.eof do
    begin
      strTemp := '';
      for i := Low(FieldName) to High(FieldName) do
        strTemp := strTemp + #1 + (DBGridEh as TDBGridEh).DataSource.DataSet.Fieldbyname(FieldName[i]).AsString;
      if strlTemp.IndexOf(strTemp) <> -1 then
      begin
        Result := True;
        if EMsg <> '' then
          ShowMessageBoxEx(EMsg, 'info');
        Exit;
      end
      else
        strlTemp.Add(strTemp);
      (DBGridEh as TDBGridEh).DataSource.DataSet.Next;
    end;
  finally
    FreeAndNil(strlTemp);
    (DBGridEh as TDBGridEh).DataSource.DataSet.EnableControls;
  end;
end;

//获得DBGrdeh的FieldList
procedure GetGrdEhColList(DBGridEh: TDBGridEh; TitleList: array of string; var FList: TStringList);
var
  i, j: Integer;
  FCol, FTitle: TObject;
  sFieldName, sCaption: string;
  FWidth: LongInt;
  FVisible: Boolean;
begin
  FCol := GetObjectProp(DBGridEh, 'Columns');
  if Length(TitleList) = 0 then
  begin
    for i := 0 to TCollection(FCol).Count - 1 do
    begin
      sFieldName := GetStrProp(TCollection(FCol).Items[i], 'FieldName');
      if (sFieldName <> '') then
        FList.Add(sFieldName);
    end;
  end
  else
  begin
    for i := 0 to Length(TitleList) - 1 do
      for j := 0 to TCollection(FCol).Count - 1 do
      begin
        sFieldName := GetStrProp(TCollection(FCol).Items[j], 'FieldName');
        FTitle := GetObjectProp(TCollection(FCol).Items[j], 'Title');
        sCaption := GetStrProp(FTitle, 'Caption');
        if (sFieldName <> '') and SameText(sCaption, TitleList[i]) then
        begin
          FList.Add(sFieldName);
          break;
        end;
      end;
  end;
end;

type
  TGetFooterValueFunc = function(row, Col: Integer): string of object;

//功能：导出 DBGridEh 指定标题列 TitleList 的字段到 Excel
procedure DBGridEhToExcel(DBGridEh: TWinControl; TitleList: array of string; SaveToFile: Boolean; DocName: string = '重命名');
  //功能：通过 TitleList 分解 获取字段。

  procedure GetDBGridEhFieldList(DBGridEh: TWinControl; TitleList: array of string; var FList, FColList: TStringList);
  var
    i, j: Integer;
    FCol, FTitle: TObject;
    sFieldName, sCaption: string;
    FWidth: LongInt;
    FVisible: Boolean;
  begin
    FCol := GetObjectProp(DBGridEh, 'Columns');
    //如果不指定列，则只导出可视的列
    if Length(TitleList) = 0 then
    begin
      for i := 0 to TCollection(FCol).Count - 1 do
      begin
        sFieldName := GetStrProp(TCollection(FCol).Items[i], 'FieldName');
        FWidth := GetOrdProp(TCollection(FCol).Items[i], 'Width');
        FVisible := GetPropValue(TCollection(FCol).Items[i], 'Visible');
        if FVisible and (FWidth > 0) and (sFieldName <> '') then
        begin
          FList.Add(sFieldName);
          FColList.Add(IntToStr(i));
        end;
      end;
    end
    else
    begin
      for i := 0 to Length(TitleList) - 1 do
      begin
        for j := 0 to TCollection(FCol).Count - 1 do
        begin
          sFieldName := GetStrProp(TCollection(FCol).Items[j], 'FieldName');
          FTitle := GetObjectProp(TCollection(FCol).Items[j], 'Title');
          sCaption := GetStrProp(FTitle, 'Caption');
          if (sFieldName <> '') and SameText(sCaption, TitleList[i]) then
          begin
            FList.Add(sFieldName);
            FColList.Add(IntToStr(j));
            break;
          end;
        end; //for j
      end; //for i
    end; //if
  end;

var
  FieldList: array of string;
  FList, FColList: TStringList;
  i, j: Integer;
  FDS: TObject;
  FCanExport: Boolean;
  xList, StrColList: TStringList;
  FFooterRowCount: Integer;
  FGetValue: TGetFooterValueFunc;
  s, lsFile: string;
  ExpClass: TDBGridEhExportClass;
  SaveDialog: TSaveDialog;
begin
  CheckValidDBGridEh(DBGridEh);

  FDS := GetObjectProp(DBGridEh, 'DataSource');
  if not Assigned(FDS) then
  begin
    Application.MessageBox('DBGridEh尚未指定DataSource！', '导出', MB_ICONERROR);
    Exit;
  end;
  if not Assigned(TDataSource(FDS).DataSet) then
  begin
    Application.MessageBox('DBGridEh尚未指定DataSet数据集！', '导出', MB_ICONERROR);
    Exit;
  end;
  if not TDataSource(FDS).DataSet.Active then
  begin
    Application.MessageBox('DBGridEh的数据集尚未打开！', '导出', MB_ICONERROR);
    Exit;
  end;

  if Assigned(GetPropInfo(DBGridEh, 'CanExportExcel')) then
  begin
    FCanExport := GetPropValue(DBGridEh, 'CanExportExcel');
    if not FCanExport then
    begin
      Application.MessageBox('系统禁止该表格导出数据！', '提示', MB_ICONWARNING);
      Exit;
    end;
  end;

  (DBGridEh as TDBGridEh).datasource.DataSet.DisableControls;
  FColList := TStringList.Create;
  try
    FList := TStringList.Create;
    try
      GetDBGridEhFieldList(DBGridEh, TitleList, FList, FColList);
      SetLength(FieldList, FList.Count);
      for i := 0 to FList.Count - 1 do
        FieldList[i] := FList[i];
    finally
      FList.Free;
    end;

    //导出到列表
    xList := TStringList.Create;
    StrColList := TStringList.Create;
    try
      DataSetEhToList(TDataSource(FDS).DataSet, FieldList, xList, StrColList, DBGridEh);
      //读取页脚值
      if Assigned(GetPropInfo(DBGridEh, 'FooterRowCount')) and Assigned(GetPropInfo(DBGridEh, 'OnGetFooterValue')) then
      begin
        FFooterRowCount := GetOrdProp(DBGridEh, 'FooterRowCount');
        FGetValue := TGetFooterValueFunc(GetMethodProp(DBGridEh, 'OnGetFooterValue'));
        for i := 0 to FFooterRowCount - 1 do
        begin
          s := '';
          for j := 0 to FColList.Count - 1 do
            s := s + FGetValue(i, StrToInt(FColList[j])) + #9;
          if StringReplace(s, #9, '', [rfReplaceAll]) <> '' then
          begin
            s := Copy(s, 1, Length(s) - Length(#9));
            xList.Add(s);
          end;
        end; //for i
      end;
      //导出到Excel
      ShowLodingMessage('正在导出数据到Excel。。。');
      ExpClass := TDBGridEhExportAsXLS;
      if SaveToFile then
      begin
        SaveDialog := TSaveDialog.Create(nil);
        try
          SaveDialog.Filter := DIALOGFILE_EXT_XLS;
          SaveDialog.Filename := DocName + '.xls';
          if SaveDialog.Execute then
          begin
            lsFile := SaveDialog.Filename;
            if lsFile = '' then
              lsFile := '转Excel';
          end
          else
          begin
            (DBGridEh as TDBGridEh).datasource.DataSet.EnableControls;
            Exit;
          end;
        finally
          SaveDialog.Free;
        end;
          //WriteListToExcelFile(xList);
        SaveDBGridEhToExportFile(ExpClass, (DBGridEh as TDBGridEh), lsFile, True);
      end;
      //ExportExcelEh(DBGridEh as TDBGridEh);
//            else
//                WriteListToExcel(xList, StrColList);
    finally
      FreeAndNil(xList);
      FreeAndNil(StrColList);
    end;
  finally
    HideLodingMessage;
    FreeAndNil(FColList);
    (DBGridEh as TDBGridEh).datasource.DataSet.EnableControls;
  end;
end;
//导出表格数据的统一函数(表格包括DBGrid、StringGrid和DBGridEh）

procedure GridToExcel(Grid: TWinControl; TitleList: array of string; SaveToFile: Boolean = False; UseTree: Boolean = True);
var
  bHadTree: Boolean;
  FTreeCount: Integer;
begin
  if Assigned(Grid) then
  begin
    bHadTree := False;
    FTreeCount := 0;
    if (not UseTree) and ((Grid is TDBGrid) or (Grid is TStringGrid) or IsValidDBGridEh(Grid)) then
    begin
      FTreeCount := 0;
      bHadTree := Assigned(GetPropInfo(Grid, 'TreeLayerCount'));
      if bHadTree then
      begin
        FTreeCount := GetOrdProp(Grid, 'TreeLayerCount');       //得到表格树状层数
        SetOrdProp(Grid, 'TreeLayerCount', 0);                  //不启用树, 则设定树层数为零
      end;
    end;

    try
      if IsValidDBGridEh(Grid) then
      begin
        DBGridEhToExcel(Grid, TitleList, SaveToFile);
      end;
    finally
      if (not UseTree) and bHadTree and (FTreeCount <> 0) then    //恢复树层数
        SetOrdProp(Grid, 'TreeLayerCount', FTreeCount);
    end;
  end;
end;
//检测组件是否是合法的TDBGridEh、TJDBGridEh类

procedure CheckValidDBGridEh(DBGridEh: TWinControl);
begin
  if DBGridEh = nil then
    raise Exception.Create('DBGridEh does not Exist!');
  if not IsValidDBGridEh(DBGridEh) then
    raise Exception.Create(Format('%s is not a valid DBGridEh!', [DBGridEh.Name]));
end;

function IsValidDBGridEh(DBGridEh: TWinControl): Boolean;
begin
  Result := False;
  if DBGridEh = nil then
    Exit;
  Result := Pos('$' + UpperCase(DBGridEh.ClassName) + '$', '$TDBGRIDEH$TDBGRID$') > 0;
end;

//替换 Tab 键值 和 换行键值
function ValidExcelCell(txt: string): string;
var
  sTemp: string;
begin
  sTemp := txt;
  if Pos(#9, sTemp) > 0 then
    sTemp := StringReplace(sTemp, #9, ' ', [rfReplaceAll]);
  if Pos(#13, sTemp) > 0 then
    sTemp := StringReplace(sTemp, #13, ' ', [rfReplaceAll]);
  if Pos(#10, sTemp) > 0 then
    sTemp := StringReplace(sTemp, #10, ' ', [rfReplaceAll]);
  Result := sTemp;
end;

//解拆多表头的标题行
function GetExcelTitleStr(var FTitleList: TStringList; UseMultiTitle: Boolean): string;

  function GetFirstItem(var Str: string; Splitter: string): string;
  var
    p: Integer;
  begin
    Str := TrimLeft(Str);
    p := Pos(Splitter, Str);
    if p = 0 then
    begin
      Result := Trim(Str);
      Str := '';
    end
    else
    begin
      Result := TrimLeft(Copy(Str, 1, p - 1));
      Delete(Str, 1, p + Length(Splitter) - 1);
      Str := TrimLeft(Str);
    end;
  end;

var
  s, sRow, sPre, sCur, sTemp: string;
  i: Integer;
  bEnd: Boolean;
  FList: TStringList;
begin
  bEnd := False;
  s := '';
  if UseMultiTitle then
  begin
    FList := TStringList.Create;
    FList.Assign(FTitleList);
    try
      while not bEnd do
      begin
        sPre := '';
        sRow := '';
        bEnd := True;
        for i := 0 to FList.Count - 1 do
        begin
          if bEnd then
            bEnd := Pos('|', FList[i]) = 0;
          sTemp := FList[i];
          sCur := GetFirstItem(sTemp, '|');
          FList[i] := sTemp;
          sTemp := sCur;
          if FList[i] <> '' then
          begin
            if sCur + '|' = sPre then
              sTemp := '';
            sPre := sCur + '|';
          end
          else
            sPre := '';
          sRow := sRow + ValidExcelCell(sTemp) + #9;
        end;  //for
        if sRow <> '' then
          sRow := Copy(sRow, 1, Length(sRow) - Length(#9));

        //换行
        s := s + sRow + #13#10;
      end;
      if s <> '' then
        s := Copy(s, 1, Length(s) - Length(#13#10));
    finally
      FreeAndNil(FList);
    end;
  end
  else
  begin
    for i := 0 to FTitleList.Count - 1 do
      s := s + ValidExcelCell(FTitleList[i]) + #9;
    if s <> '' then
      s := Copy(s, 1, Length(s) - Length(#9));
  end;

  Result := s;
end;

//得到DBGrid中显示为复选框的列集
procedure GetDBGridCheckBoxList(DBGrid: TWinControl; var CheckBoxColList: TStringList);
var
  i: Integer;
  FCol: TObject;
  B: Boolean;
begin
  if DBGrid = nil then
    Exit;
  //得到哪些列显示CheckBox图标
  if Assigned(GetPropInfo(DBGrid, 'CheckBoxColList')) then   //JDBGrid的情况
    CheckBoxColList.Assign(TStringList(GetObjectProp(DBGrid, 'CheckBoxColList')))
  else if Assigned(GetPropInfo(DBGrid, 'Columns')) then           //DBGridEh的情况
  begin
    FCol := GetObjectProp(DBGrid, 'Columns');
    for i := 0 to TCollection(FCol).Count - 1 do
    begin
      if not Assigned(GetPropInfo(TCollection(FCol).Items[i], 'CheckBoxes')) then
        Break;
      B := GetPropValue(TCollection(FCol).Items[i], 'CheckBoxes');
      if B then
        CheckBoxColList.Add(IntToStr(TCollection(FCol).Items[i].ID - 1));
    end;
  end;
end;

//功能：导出 TDataSet 的 FieldList 字段内容到xList中。
procedure DataSetEhToList(qry: TDataSet; FieldList: array of string; var xList, StrColList: TStringList; DBGridEh: TWinControl);
var
  i, j: Integer;
  FList, FSrcList, FTitleList, FCheckBoxColList: TStringList; //FieldName List
  s, sTemp: string;
  BM: TBookMark;
  FTreeCount: Integer;
  FTreeValues: TStringList;
  {bNotGetTree, }  bUseMultiTitle, bDrawMemoText, bTemp: Boolean;
  //eTemp: Extended;
  Cur: TCursor;
  FCol, FTitle: TObject;
  sFieldName: string;
  FVisible: Boolean;
begin
  if DBGridEh <> nil then
    CheckValidDBGridEh(DBGridEh);
  FTreeCount := 0;
  //bNotGetTree := True;
  Cur := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  FTreeValues := TStringList.Create;
  try
    FCheckBoxColList := TStringList.Create;
    FList := TStringList.Create;
    try
      FSrcList := TStringList.Create;
      FTitleList := TStringList.Create;
      try
          //如果没有传入导出字段，则读取所有字段
        if Length(FieldList) = 0 then
        begin
          for i := 0 to qry.FieldCount - 1 do
            FSrcList.Add(qry.Fields[i].FieldName);
        end
        else   //否则读取所有导出字段
        begin
          for i := 0 to Length(FieldList) - 1 do
            FSrcList.Add(FieldList[i]);
        end;

        s := '';
          //DBGridEh 6.0后已经隐藏多表头属性,直接赋值
          //bUseMultiTitle := (DBGridEh <> nil) And GetPropValue(DBGridEh as TDBGridEh, 'UseMultiTitle');
        bUseMultiTitle := (DBGridEh as TDBGridEh).UseMultiTitle;
        bDrawMemoText := (DBGridEh <> nil) and GetPropValue(DBGridEh, 'DrawMemoText');

          //如果存在DBGridEh，则仅读取要求导出并且DBGridEh的列所关联的字段，以及导出标题内容 s
        if DBGridEh <> nil then
        begin
          if Assigned(GetPropInfo(DBGridEh, 'TreeLayerCount')) then
            FTreeCount := GetOrdProp(DBGridEh, 'TreeLayerCount');     //得到表格树状层数

          FCol := GetObjectProp(DBGridEh, 'Columns');
          for j := 0 to TCollection(FCol).Count - 1 do
          begin
            FVisible := GetPropValue(TCollection(FCol).Items[j], 'Visible');
            if not FVisible then
              Continue;
                  (* add over *)
            sFieldName := GetStrProp(TCollection(FCol).Items[j], 'FieldName');
            if sFieldName <> '' then
            begin
              for i := 0 to FSrcList.Count - 1 do
              begin
                if SameText(sFieldName, FSrcList[i]) then
                begin
                  FList.Add(FSrcList[i]);
                  FTitle := GetObjectProp(TCollection(FCol).Items[j], 'Title');
                  FTitleList.Add(GetStrProp(FTitle, 'Caption'));
                              {注释:由于现在树层显示层数表示可视列的层数, 因此以下这段不再需要.
                              if (j >= FTreeCount - 1) And bNotGetTree then
                              begin
                                  //得到指定字段的树状层数
                                  if j > FTreeCount - 1 then
                                      FTreeCount := FList.Count - 1
                                  else
                                      FTreeCount := FList.Count;
                                  bNotGetTree := False;
                              end;
                              }
                  break;
                end; //if
              end; //for i
            end; //if sFieldName <> ''
          end; //for j

              //导出标题内容 s
          s := GetExcelTitleStr(FTitleList, bUseMultiTitle);
          if s <> '' then
          begin
                  //s := Copy(s, 1, Length(s) - Length(#9));
            xList.Add(s);
          end;
        end
        else  //否则，读取所有导出字段
          FList.Assign(FSrcList);
      finally
        FreeAndNil(FTitleList);
        FreeAndNil(FSrcList);
      end;

      //如果字段类型为字符串、日期时间型，则导出后的Excel单元格设置为文本格式。
      StrColList.Clear;
      for i := 0 to FList.Count - 1 do
      begin
        if qry.FieldByName(FList[i]).DataType in [ftString, ftDate, ftTime, ftDateTime] then
          StrColList.Add(IntToStr(i));
      end;

      //得到要显示为复选框的列
      GetDBGridCheckBoxList(DBGridEh, FCheckBoxColList);

      //读取数据内容
      BM := qry.GetBookmark;
      try
        qry.DisableControls;
        //qry.DisableConstraints;
        qry.First;
        while not qry.Eof do
        begin
          s := '';
          for i := 0 to FList.Count - 1 do
          begin
            if (qry.FieldByName(FList[i]).DataType = ftMemo) and bDrawMemoText then
              sTemp := qry.FieldByName(FList[i]).AsString
            else
            (* add over *)
              sTemp := qry.FieldByName(FList[i]).DisplayText;
            sTemp := ValidExcelCell(sTemp);

            //处理树状层次显示
            if i < FTreeCount then
            begin
              if i >= FTreeValues.Count then
                FTreeValues.Add(sTemp)
              else if FTreeValues[i] = sTemp then
                sTemp := ''
              else
              begin
                for j := FTreeValues.Count - 1 downto i + 1 do
                  FTreeValues.Delete(j);
                FTreeValues[i] := sTemp;
              end;
            end;

            //处理复选框显示
            if (sTemp <> '') and (DBGridEh <> nil) then
            begin
              FCol := GetObjectProp(DBGridEh, 'Columns');
              if IsCheckBoxCol(TCollection(FCol), FList[i], FCheckBoxColList) then
              begin
                if tryStrToBool(qry.FieldByName(FList[i]).AsString, bTemp) and bTemp then
                  sTemp := '√'
                else
                  sTemp := '';
              end;
            end;

            {//注释:已使用强制单元格格式为文本格式的办法替代。
            //防止字符串型的数值, 导入Excel后变为数值型丢失前面的零(如: '001' ===> 1).
            if (sTemp <> '') And (qry.FieldByName(FList[i]).DataType = ftString)
                    And TryStrToFloat(sTemp, eTemp) then
            begin
                if FloatToStr(eTemp) <> sTemp then  //如果不等, 才转换
                    sTemp := '=Trim("' + sTemp + '")';
            end;}

            s := s + sTemp + #9;
            Application.ProcessMessages;
          end;
          if StringReplace(s, #9, '', [rfReplaceAll]) <> '' then
          begin
            s := Copy(s, 1, Length(s) - Length(#9));
            xList.Add(s);
          end;
          qry.Next;
        end;
      finally
        qry.GotoBookmark(BM);
        //qry.EnableConstraints;
        qry.EnableControls;
      end;
    finally
      FreeAndNil(FList);
      FreeAndNil(FCheckBoxColList);
    end;
  finally
    Screen.Cursor := Cur;
    FreeAndNil(FTreeValues);
  end;
end;

//同上
procedure DataSetEhToList(qry: TDataSet; FieldList: array of string; var xList: TStringList; DBGridEh: TWinControl);
var
  StrColList: TStringList;
begin
  StrColList := TStringList.Create;
  try
    DataSetEhToList(qry, FieldList, xList, StrColList, DBGridEh);
  finally
    FreeAndNil(StrColList);
  end;
end;

//列ColIndex是否为复选框列
function IsCheckBoxCol(FCol: TCollection; FieldName: string; CheckBoxColList: TStringList): Boolean;
  //找到列最小的ID

  function GetMinColID: Integer;
  var
    i: Integer;
  begin
    Result := FCol.Items[0].ID;
    for i := 1 to FCol.Count - 1 do
    begin
      if FCol.Items[i].ID < Result then
        Result := FCol.Items[i].ID;
    end;
  end;

var
  i, ColIndex: Integer;
  Offset: Integer;  //列Column.ID的偏移值。因为如果一开始时没有绑定列，则缺省状态下动态创建了一个列，
                    //打开数据集后，该列被删除，再建立新列，新列的ID从2开始。如果关闭数据集再重新打开，新列的ID值更大。
begin
  if TDBGridColumns(FCol).State = csDefault then
    Offset := GetMinColID - 1
  else
    Offset := 0;
  ColIndex := -1;
  for i := 0 to FCol.Count - 1 do
  begin
    if SameText(FieldName, GetPropValue(TCollection(FCol).Items[i], 'FieldName')) then
    begin
      ColIndex := i;
      Break;
    end;
  end;
  Result := (ColIndex <> -1) and (CheckBoxColList.Count > 0) and (CheckBoxColList.IndexOf(IntToStr(FCol.Items[ColIndex].ID - Offset - 1)) <> -1);
end;

//把tsList导出到 Excel 文件（不需要系统安装Excel）
procedure WriteListToExcelFile(xList: TStringList; AFileName: string = '');
begin
  WriteListToExcelFile(xList, AFileName);
end;

//功能：把List的内容导入到Excel中
procedure WriteListToExcel(tsList: TStringList; StrColList: TStrings);
begin
  WriteListToExcel(tsList, StrColList);
end;


//Added by CYQ 2013-08-09
//输入DBGridEh名称(必传)，是否输出底部合计行(默认False)，输出标题(默认为'')。
procedure ExportExcelEh(const CurDS: TDBGridEh; vSum: Boolean = False; Title: string = '');

  function IntToChr(i: integer): string;
  var
    c: string;
  begin
    if i < 26 then
      c := ''
    else
      c := Chr((i div 26) - 1 + 65);
    result := UpperCase(c + Chr((i mod 26) + 65));
  end;

var
  nowhangi, i, j, k, ii, LCID, liRecNu, liFieldNu: integer; //nowhangi当前 execl的所在行，jilushu当前的记录位置
  lsFile, cols, lsstr: string;
  ExcelApplication1: TExcelApplication;
  ExcelWorkbook1: TExcelWorkbook;
  ExcelWorkSheet1: TExcelWorksheet;
  xl: OleVariant;
  SaveDialog1: TSaveDialog;
  isGroups: Boolean;
  MultiArr: OleVariant;
begin
  if not Assigned(CurDS) then
  begin
    ShowMessageBoxEx('请点击选择要导出数据的表格！', 'Info');
    Exit;
  end;
  CurDS.datasource.DataSet.DisableControls;
  CurDS.datasource.DataSet.Last;   //先提取全部数据

  if not CurDS.datasource.DataSet.Active then
    Exit;

  liRecNu := CurDS.datasource.DataSet.RecordCount; //取记录数
  liFieldNu := CurDS.Columns.Count; //取字段数
  ii := 0;
  for i := 1 to liFieldNu do
  begin
    if not CurDS.Columns.Items[i - 1].Visible then
      Continue;
    Inc(ii);
  end;
  cols := IntToChr(ii - 1);  //最大列号

  if (liRecNu < 1) or (ii < 1) then
  begin
    Application.MessageBox('导入的数据为空！', '警告', MB_ICONHAND);
    Exit;
  end;

  //取要存为的文件名及路径
  SaveDialog1 := TSaveDialog.Create(nil);
  try
    SaveDialog1.Filter := DIALOGFILE_EXT_XLS;
    SaveDialog1.Filename := Title + '.xls';
    if SaveDialog1.Execute then
    begin
      lsFile := SaveDialog1.Filename;
      if lsFile = '' then
        lsFile := '转Excel';
    end
    else
    begin
      CurDS.datasource.DataSet.EnableControls;
      Exit;
    end;
  finally
    SaveDialog1.Free;
  end;

  try
    ShowLodingMessage('数据正在导出,请稍候...');
    if Pos('.', lsFile) = 0 then
      lsFile := lsFile + '.xls';

    //调用Excel
    try
      ExcelApplication1 := TExcelApplication.Create(ExcelApplication1);
      ExcelWorkbook1 := TExcelWorkbook.Create(ExcelWorkbook1);
      ExcelWorkSheet1 := TExcelWorksheet.Create(ExcelWorkSheet1);
    except
      Application.MessageBox('对不起，您的操作系统没有安装Excel或者Excel出现严重问题！', '警告', MB_ICONHAND);
      CurDS.datasource.DataSet.EnableControls;
      Abort;
      Exit;
    end;
    Screen.Cursor := crHourGlass;
    LCID := LOCALE_USER_DEFAULT;

    if FileExists(lsFile) then //删除动态生成的临时temp.xls文件
    try
      DeleteFile(lsFile);
    except

    end;

    try
      ExcelApplication1.Connect;
      //  ExcelApplication1.Workbooks.Add(null,0);
      ExcelApplication1.Workbooks.Add(1, 0);
      //如果调用一个模板 ，改动该句如：
      //  ExcelApplication1.Workbooks.Add(lsFile, 0);
      ExcelWorkbook1.ConnectTo(ExcelApplication1.Workbooks[1]);
      ExcelWorkSheet1.ConnectTo(ExcelWorkbook1.Sheets[1] as _WorkSheet);

      nowhangi := 1;

      //初始化Excel单元格，字体等
      ExcelApplication1.Range['A' + IntToStr(nowhangi), cols + IntToStr(nowhangi)].Merge(xl); //合并execl单元格
      ExcelWorkSheet1.Cells.Item[nowhangi, 1].Font.Size := 24;
      ExcelWorkSheet1.Cells.Item[nowhangi, 1].Font.Name := '宋体';
      ExcelWorkSheet1.cells.Item[nowhangi, 1].font.fontStyle := fsBold;
      ExcelWorkSheet1.Cells.Item[nowhangi, 1].RowHeight := 24;
      ExcelWorkSheet1.Cells.Item[nowhangi, 1] := Title;
      ExcelWorkSheet1.Cells.Item[nowhangi, 1].HorizontalAlignment := 3;
      ExcelWorkSheet1.cells.Item[nowhangi, 1].font.Color := clred;

      //判断是否有多表头
      isGroups := False;
      lsstr := '';
      k := 1;
      if CurDS.UseMultiTitle then
        for i := 0 to liFieldNu - 1 do
          if CurDS.Columns.Items[i].Visible then
          begin
            if Pos('|', CurDS.Columns.Items[i].Title.Caption) > 0 then
            begin
              lsstr := Trim(Copy(CurDS.Columns.Items[i].Title.Caption, 1, Pos('|', CurDS.Columns.Items[i].Title.Caption) - 1));
              j := k;
              isGroups := TRUE;
              Break;
            end;
            Inc(k);
          end;

      //获得多表头的列
      if isGroups then
      begin
        k := 1;
        MultiArr := VarArrayCreate([0, ii - 1], VarVariant);
        for i := 0 to liFieldNu - 1 do
        begin
          if CurDS.Columns.Items[i].Visible then
          begin
            if Pos('|', CurDS.Columns.Items[i].Title.Caption) > 0 then
            begin

              if Trim(Copy(CurDS.Columns.Items[i].Title.Caption, 1, Pos('|', CurDS.Columns.Items[i].Title.Caption) - 1)) <> lsstr then
              begin
                j := k;
                lsstr := Copy(CurDS.Columns.Items[i].Title.Caption, 1, Pos('|', CurDS.Columns.Items[i].Title.Caption) - 1);
              end;
              MultiArr[k - 1] := VarArrayOf([j, k]); //多表头
            end
            else
              MultiArr[k - 1] := VarArrayOf([0, 0]); //一个表头
            Inc(k);
          end;
        end;
      end;

      //处理表格标题
      if isGroups then
      begin
        nowhangi := nowhangi + 1;
        ii := 1;
        for i := 0 to liFieldNu - 1 do
        begin
          if CurDS.Columns.Items[i].Visible then
          begin
            if MultiArr[ii - 1][0] = MultiArr[ii - 1][1] then
            begin
              if MultiArr[ii - 1][0] > 0 then //多表头显示部分
              begin
                ExcelWorkSheet1.Cells.Item[nowhangi, ii] := Copy(CurDS.Columns.Items[i].Title.Caption, 1, Pos('|', CurDS.Columns.Items[i].Title.Caption) - 1);
                ExcelWorkSheet1.Cells.Item[nowhangi + 1, ii] := Trim(Copy(CurDS.Columns.Items[i].Title.Caption, Pos('|', CurDS.Columns.Items[i].Title.Caption) + 1, 20));
              end
              else
              begin
                ExcelWorkSheet1.Cells.Item[nowhangi, ii] := CurDS.Columns.Items[i].Title.Caption;
                ExcelApplication1.Range[IntToChr(ii - 1) + '2', IntToChr(ii - 1) + '3'].Merge(xl); //合并单维列头
              end;
            end
            else //单表头显示部分
            begin
              ExcelWorkSheet1.Cells.Item[nowhangi + 1, ii] := Trim(Copy(CurDS.Columns.Items[i].Title.Caption, Pos('|', CurDS.Columns.Items[i].Title.Caption) + 1, 20));
              ExcelApplication1.Range[IntToChr(MultiArr[ii - 1][0] - 1) + '2', IntToChr(MultiArr[ii - 1][1] - 1) + '2'].Merge(xl); //多维列头合并
            end;
            ExcelWorkSheet1.Cells.Item[nowhangi, ii].HorizontalAlignment := 3; //居中
            ExcelWorkSheet1.Cells.Item[nowhangi + 1, ii].HorizontalAlignment := 3;
            ExcelWorkSheet1.Cells.Item[nowhangi, ii].Font.Color := CurDS.Columns[i].Font.Color;
            Inc(ii);
          end;
        end;
        nowhangi := nowhangi + 1;
      end
      else //无多维列头
      begin
        nowhangi := nowhangi + 1;
        ii := 1;
        for i := 0 to liFieldNu - 1 do
        begin
          if CurDS.Columns.Items[i].Visible then
          begin
            ExcelWorkSheet1.Cells.Item[nowhangi, ii] := CurDS.Columns.Items[i].Title.Caption;
            ExcelWorkSheet1.Cells.Item[nowhangi, ii].Font.Color := CurDS.Columns[i].Font.Color;
            ExcelWorkSheet1.Cells.Item[nowhangi, ii].HorizontalAlignment := 3;
            Inc(ii);
          end;
        end;
      end;

      //定义列的数据类型
      ii := 1;
      for i := 0 to liFieldNu - 1 do
      begin
        if CurDS.Columns.Items[i].Visible then
        begin
          if CurDS.Columns[i].Field.DataType in [ftString, ftFixedChar, ftwidestring, ftMemo] then
            ExcelApplication1.Range[IntToChr(ii - 1) + IntToStr(nowhangi), IntToChr(ii - 1) + IntToStr(liRecNu + nowhangi)].NumberFormatLocal := '@'  //字符型
          else if CurDS.Columns[i].Field.DataType = ftDateTime then
            ExcelApplication1.Range[IntToChr(ii - 1) + IntToStr(nowhangi), IntToChr(ii - 1) + IntToStr(liRecNu + nowhangi)].NumberFormatLocal := 'yyyy-MM-dd HH:mm:ss'  //日期型
          else if CurDS.Columns[i].Field.DataType = ftDate then
            ExcelApplication1.Range[IntToChr(ii - 1) + IntToStr(nowhangi), IntToChr(ii - 1) + IntToStr(liRecNu + nowhangi)].NumberFormatLocal := 'yyyy-MM-dd'
          else if CurDS.Columns[i].Field.DataType = ftTime then
            ExcelApplication1.Range[IntToChr(ii - 1) + IntToStr(nowhangi), IntToChr(ii - 1) + IntToStr(liRecNu + nowhangi)].NumberFormatLocal := 'HH:mm:ss'; //字符型

          inc(ii);
        end;
      end;

      //输出查询数据
      with CurDS do
      begin
        datasource.DataSet.First;
        while not datasource.DataSet.Eof do
        begin
          nowhangi := nowhangi + 1;
          ii := 1;
          for i := 0 to liFieldNu - 1 do
            if Columns.Items[i].Visible then
            begin
              ExcelWorkSheet1.Cells.Item[nowhangi, ii] := Fields[i].Text;
              ExcelWorkSheet1.Cells.Item[nowhangi, ii].Font.Color := CurDS.Columns[i].Font.Color;
              //ExcelWorkSheet1.Cells.Item[nowhangi, ii].Interior.style := CurDS.Columns[i].Font.Style;
              ExcelWorkSheet1.Cells.Item[nowhangi, ii].Interior.Pattern := 1;
//              if CurDS.Columns[i].Color = clBlack then
//              ExcelWorkSheet1.Cells.Item[nowhangi, ii].Interior.Color := clWhite else
//              ExcelWorkSheet1.Cells.Item[nowhangi, ii].Interior.Color := CurDS.Columns[i].Color;
              Inc(ii);
            end;
          datasource.DataSet.next;
        end;
        datasource.DataSet.First;
      end;

      //需要输出合计行
      if vSum then
      begin
        ii := 0;
        if isGroups then
          k := 4
        else
          k := 3;
        nowhangi := nowhangi + 1;
        for i := 0 to liFieldNu - 1 do
          if CurDS.Columns.Items[i].Visible then
          begin
            if (CurDS.Columns[i].Field.DataType in [ftSmallint, ftInteger, ftFloat, ftCurrency, ftLargeint, ftFMTBcd, ftBCD]) then
              ExcelWorkSheet1.Cells.Item[nowhangi, ii + 1] := '=Sum(' + IntToChr(ii) + IntToStr(k) + ':' + IntToChr(ii) + Trim(IntToStr(nowhangi - 1)) + ')';
            Inc(ii);
          end;
        ii := 0;
        for i := 0 to liFieldNu - 1 do
          if CurDS.Columns.Items[i].Visible then
          begin
            if (CurDS.Columns[i].Field.DataType in [ftString, ftwidestring, ftMemo]) then
            begin
              ExcelWorkSheet1.Cells.Item[nowhangi, ii + 1] := '合计';
              ExcelWorkSheet1.Cells.Item[nowhangi, ii + 1].HorizontalAlignment := 3;
              Break;
            end;
            Inc(ii);
          end;
      end;
      //画单元格边框
      if isGroups then  //计算总行数
        k := 3
      else
        k := 2;
      if vSum then
        ii := liRecNu + k + 1
      else
        ii := liRecNu + k;
      with ExcelApplication1.Range['A' + IntToStr(2), cols + IntToStr(ii)].Borders do //画单元格边框
      begin
        LineStyle := xlContinuous;
        Weight := xlThin;
        ColorIndex := xlAutomatic;
      end;

      ExcelWorkbook1.SaveCopyAs(lsFile);
      Application.MessageBox(pchar('数据已成功导出到' + lsFile + '！'), '提示', MB_ICONINFORMATION);
    except
      Application.MessageBox('导出到EXCEL失败！', '警告', MB_ICONHAND);
      CurDS.datasource.DataSet.EnableControls;
    end;
  finally
    CurDS.datasource.DataSet.EnableControls;
    ExcelApplication1.DisplayAlerts[LCID] := False;
    ExcelApplication1.Quit;
    ExcelWorkSheet1.Disconnect;
    ExcelWorkbook1.Disconnect;
    ExcelApplication1.Disconnect;
    HideLodingMessage;

    if not VarIsNull(MultiArr) then
      VarClear(MultiArr);
    Screen.Cursor := crDefault;
  end;
end;

procedure AddValueToCombobox(qryTemp: TADOQuery; strInputSQL: string; Combox: TCustomComboBox; bShowAll: Boolean = False); overload;
begin
  if strInputSQL = '' then
    Exit;
  Combox.Items.Clear;

  with qryTemp do
  begin
    Close;
    SQL.Text := strInputSQL;
    try
      Open;
      DisableControls;
      First;
      if RecordCount > 0 then
        while not Eof do
        begin
          Combox.Items.Add(Trim(PChar(Fields[0].AsString)));
          Next;
        end;
      //Combox.ItemIndex := 0;
      EnableControls;

      if bShowAll then  //添加全部选项
      begin
        Combox.Items.Add('全部');
        Combox.ItemIndex := Pred(Combox.Items.Count);
      end;
    except
    end;
  end;
end;

procedure AddValueToCombobox(qryTemp: TFDQuery; strInputSQL: string; Combox: TCustomComboBox; bShowAll: Boolean = False); overload;
begin
  if strInputSQL = '' then
    Exit;
  Combox.Items.Clear;

  with qryTemp do
  begin
    Close;
    SQL.Text := strInputSQL;
    try
      Open;
      DisableControls;
      First;
      if RecordCount > 0 then
        while not Eof do
        begin
          Combox.Items.Add(Trim(PChar(Fields[0].AsString)));
          Next;
        end;
      //Combox.ItemIndex := 0;
      EnableControls;

      if bShowAll then  //添加全部选项
      begin
        Combox.Items.Add('全部');
        Combox.ItemIndex := Pred(Combox.Items.Count);
      end;
    except
    end;
  end;
end;

//一定要按照先ID，再Name的顺序传入
procedure AddValueToComboboxByID(qryTemp: TADOQuery; strInputSQL: string; Combox: TCustomComboBox; bShowAll: Boolean = False);
var
  i: Integer;
  sID: PString;
begin
  if strInputSQL = '' then
    Exit;
  Combox.Items.Clear;

  with qryTemp do
  begin
    Close;
    SQL.Text := strInputSQL;
    try
      Open;
      DisableControls;
      First;
      if RecordCount > 0 then
        while not Eof do
        begin
          New(sID);
          sID^ := Fields[0].AsString;
          Combox.Items.AddObject(Trim(Fields[1].AsString), TObject(sID^)); //Ö¸Ïò¶ÔÏóµØÖ·
          Next;
        end;
      Combox.ItemIndex := 0;
      EnableControls;

      if bShowAll then  //添加全部选项
      begin
        Combox.Items.Add('全部');
        Combox.ItemIndex := Pred(Combox.Items.Count);
      end;
    except
      EnableControls;
    end;
  end;
end;

procedure AddValueToComboboxByID(qryTemp: TFDQuery; strInputSQL: string; Combox: TCustomComboBox; bShowAll: Boolean = False);
var
  i: Integer;
  sID: PString;
begin
  if strInputSQL = '' then
    Exit;
  Combox.Items.Clear;

  with qryTemp do
  begin
    Close;
    SQL.Text := strInputSQL;
    try
      Open;
      DisableControls;
      First;
      if RecordCount > 0 then
        while not Eof do
        begin
          New(sID);
          sID^ := Trim(Fields[0].AsString);
          Combox.Items.AddObject(Trim(Fields[1].AsString), TObject(sID^)); //添加Combox的KeyValue KeyName
        //Dispose(sID);
          Next;
        end;
      EnableControls;

      if bShowAll then  //添加全部选项
      begin
        Combox.Items.Add('全部');
        Combox.ItemIndex := Pred(Combox.Items.Count)
      end
      else
        Combox.ItemIndex := 0;
    except
      EnableControls;
    end;
  end;
end;

function GetComboboxID(Combox: TCustomComboBox): string;
begin
  if Combox.ItemIndex >= 0 then
    Result := string(Combox.Items.Objects[Combox.ItemIndex])
  else
    Result := '';
  //Result := PChar(string(Combox.Items.Objects[Combox.ItemIndex]));
end;

function LocateComboboxByID(Combox: TCustomComboBox; ID: string): Boolean;
var
  i: Integer;
begin
  Result := False;
  Combox.ItemIndex := -1;
  if Combox.Items.Count >= 0 then
    for i := 0 to Pred(Combox.Items.Count) do
    begin
      if SameText(string(Combox.Items.Objects[i]), ID) then
      begin
        Combox.ItemIndex := i;
        Result := True;
        Exit;
      //Combox.SelText := Combox.Items.Strings[i];
      end;
    end;
  //Result := PChar(string(Combox.Items.Objects[Combox.ItemIndex]));
end;

function CashSmallToBig(small: Real): string;
//小写金额转换为大写金额
var
  SmallMonth, BigMonth: string;
  wei1, qianwei1: string[2];
  wei, qianwei, dianweizhi, qian: integer;
begin
  Result := '';
  {------- 修改参数令值更精确 -------}
  {小数点后的位数,需要的话也可以改动该值}
  qianwei := -2;
  {转换成货币形式,需要的话小数点后加多几个零}
  SmallMonth := formatfloat('0.00', small);
  BigMonth := '';
  {---------------------------------}
  dianweizhi := pos('.', SmallMonth);{小数点的位置}
  {循环小写货币的每一位,从小写的右边位置到左边}
  for qian := length(SmallMonth) downto 1 do
  begin
    {如果读到的不是小数点就继续}
    if qian <> dianweizhi then
    begin
      {位置上的数转换成大写}
      case strtoint(copy(SmallMonth, qian, 1)) of
        1:
          wei1 := '壹';
        2:
          wei1 := '贰';
        3:
          wei1 := '叁';
        4:
          wei1 := '肆';
        5:
          wei1 := '伍';
        6:
          wei1 := '陆';
        7:
          wei1 := '柒';
        8:
          wei1 := '捌';
        9:
          wei1 := '玖';
        0:
          wei1 := '零';
      end;
      {判断大写位置,可以继续增大到real类型的最大值}
      case qianwei of
        -3:
          qianwei1 := '厘';
        -2:
          qianwei1 := '分';
        -1:
          qianwei1 := '角';
        0:
          qianwei1 := '元';
        1:
          qianwei1 := '拾';
        2:
          qianwei1 := '佰';
        3:
          qianwei1 := '千';
        4:
          qianwei1 := '万';
        5:
          qianwei1 := '拾';
        6:
          qianwei1 := '佰';
        7:
          qianwei1 := '千';
        8:
          qianwei1 := '亿';
        9:
          qianwei1 := '十';
        10:
          qianwei1 := '佰';
        11:
          qianwei1 := '千';
      end;
      inc(qianwei);
      BigMonth := wei1 + qianwei1 + BigMonth;{组合成大写金额}
    end;
  end;
  Result := BigMonth;
end;

procedure OpenNormalForm(AForm: TCustomForm; const AFormClass: TFormClass);

  function IsExistForm: Boolean;
  var
    i: Integer;
  begin
    Result := False;
    for i := 0 to Screen.FormCount - 1 do
    begin
      if Screen.Forms[i].ClassType = AFormClass then
      begin
        Result := True;
        AForm := TForm(Screen.Forms[i]);
        Break;
      end;
    end;
  end;

begin
  try
    if not IsExistForm then
    begin
      AForm := AFormClass.Create(nil);
      AForm.ShowModal;
    end
    else
    begin
      ShowWindow(AForm.Handle, SW_NORMAL);
      AForm.BringToFront;
      AForm.ShowModal;
    end;
  finally
    AForm.Free;
  end;
end;

procedure OpenNormalForm_zhs(AForm: TCustomForm; const AFormClass: TFormClass);

  function IsExistForm: Boolean;
  var
    i: Integer;
  begin
    Result := False;
    for i := 0 to Screen.FormCount - 1 do
    begin
      if Screen.Forms[i].ClassType = AFormClass then
      begin
        Result := True;
        AForm := TForm(Screen.Forms[i]);
        Break;
      end;
    end;
  end;

begin
  if not IsExistForm then
  begin
    AForm := AFormClass.Create(nil);
    AForm.Show;
  end
  else
  begin
    ShowWindow(AForm.Handle, SW_NORMAL);
    AForm.BringToFront;
    AForm.Show;
  end;
end;

procedure OpenChildForm(FormClass: TFormClass; var Fm; AOwner: TComponent);
var
  i: integer;
  Child: TForm;
begin
  for i := 0 to Screen.FormCount - 1 do
    if Screen.Forms[i].ClassType = FormClass then
    begin
      Child := Screen.Forms[i];
      if Child.WindowState = wsMinimized then
        ShowWindow(Child.handle, SW_SHOWMAXIMIZED)
      else
        ShowWindow(Child.handle, SW_SHOWMAXIMIZED);
      if (not Child.Visible) then
        Child.Visible := True;
      Child.BringToFront;
      Child.Setfocus;
      TForm(Fm) := Child;
      exit;
    end;
  Child := TForm(FormClass.NewInstance);
  TForm(Fm) := Child;
  Child.Create(AOwner);
  Child.WindowState := wsMaximized;
end;


//把某容器Parent里的所有控件设置输入法为imDontCare, 不关心.
procedure DontCareIme(Parent: TWinControl);
var
  Ctrl: TControl;
  i: Integer;
begin
  if Assigned(GetPropInfo(Parent, 'ImeMode')) then
  begin
    SetStrProp(Parent, 'ImeMode', 'imDontCare');
    SetStrProp(Parent, 'ImeName', '');
  end;
  for i := 0 to Parent.ControlCount - 1 do
  begin
    Ctrl := Parent.Controls[i];
    if Ctrl is TWinControl then
      DontCareIme(TWinControl(Ctrl));
  end;
end;

procedure OpenChildForm(FormClass: TCustomFormClass; var Form; bDontCareIme: Boolean);
var
  Data: TDataModule;
  i, j: Integer;
  Comp, MesComp: TComponent;
  SaveCursor: TCursor;
begin
  MesComp := nil;
  Application.ProcessMessages;
  if Pos('$' + FormClass.ClassName + '$', '$' + sOpenWindowClassStr + '$') > 0 then
    Exit; //cyw 09.1.15
  sOpenWindowClassStr := sOpenWindowClassStr + FormClass.ClassName + '$';                 //cyw 09.1.15
  SaveCursor := Screen.Cursor;                                                            //cyw 09.1.15
  try
        //Screen.Cursor := crHourGlass;                                                       //cyw 09.1.15
    if Assigned(Application.MainForm) then
      LockWindowUpdate(Application.MainForm.ClientHandle);

    if Assigned(TCustomForm(Form)) then
    begin
      if TCustomForm(Form).WindowState = wsMinimized then
        TCustomForm(Form).Perform(WM_SYSCOMMAND, SC_RESTORE, 0);
    end
    else
    begin
      TCustomForm(Form) := FormClass.Create(Application);
      TCustomForm(Form).WindowState := wsMaximized;//默认最大化
      if bDontCareIme then
        DontCareIme(TCustomForm(Form));
    end;
    TCustomForm(Form).BringToFront;
  finally
    LockWindowUpdate(0);
    Screen.Cursor := SaveCursor;
    sOpenWindowClassStr := StringReplace(sOpenWindowClassStr, FormClass.ClassName + '$', '', [rfIgnoreCase]); //cyw 09.1.15
  end;
end;

function ExcuteExe(Exe_Dir: string): Boolean;
var
  ScommandLine: string;
  LpStartUpInfo: TStartupInfo;
  LpProcessInformation: TProcessInformation;
begin
  try
    ScommandLine := Exe_Dir;
    Fillchar(LpStartUpInfo, sizeof(TStartupInfo), #0);
    LpStartUpInfo.cb := SizeOf(TStartupInfo);
    LpStartUpInfo.dwflags := STARTF_USESHOWWINDOW;
    LpStartUpInfo.wshowwindow := SW_NORMAL;
    CreateProcess(nil, Pchar(ScommandLine), nil, nil, True, CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil, nil, LpStartUpInfo, LpProcessInformation);
    Result := True;
  except
    Result := False;
  end;
end;

function FindProcess(AFileName: string): boolean;
var
  hSnapshot: THandle; //用于获得进程列表
  lppe: TProcessEntry32; //用于查找进程
  Found: Boolean; //用于判断进程遍历是否完成
begin
  Result := False;
  hSnapshot := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);//获得系统进程列表
  lppe.dwSize := SizeOf(TProcessEntry32);//在调用Process32First   API之前，需要初始化lppe记录的大小
  Found := Process32First(hSnapshot, lppe);//将进程列表的第一个进程信息读入ppe记录中
  while Found do
  begin
    if ((UpperCase(ExtractFileName(lppe.szExeFile)) = UpperCase(AFileName)) or (UpperCase(lppe.szExeFile) = UpperCase(AFileName))) then
    begin
      Result := True;
    end;
    Found := Process32Next(hSnapshot, lppe);//将进程列表的下一个进程信息读入lppe记录中
  end;
end;

procedure KillProcess(AFileName: string);
const
  PROCESS_TERMINATE = $0001;
var
  ContinueLoop: BOOL;
  FSnapShotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  FSnapShotHandle := CreateToolhelp32SnapShot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
  ContinueLoop := Process32First(FSnapShotHandle, FProcessEntry32);
  while integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = UpperCase(AFileName)) or (UpperCase(FProcessEntry32.szExeFile) = UpperCase(AFileName))) then
      TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0), FProcessEntry32.th32ProcessID), 0);
    ContinueLoop := Process32Next(FSnapShotHandle, FProcessEntry32);
  end;
end;

function GetProcessHandleByName(Name: string): THandle;
var
  Hd, Hs: THandle;
  dExit: Cardinal;
  Tmp, Tmp1: string;
  Lp: TProcessEntry32;
begin
  Result := 0;
  Lp.dwSize := sizeof(TProcessEntry32);
  Hd := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if Process32First(Hd, Lp) then
    repeat
      Tmp := UpperCase(Trim(Name));
      Tmp1 := Trim(UpperCase(Lp.szExeFile));
      if AnsiPos(Tmp, Tmp1) > 0 then
      begin
        Result := OpenProcess(PROCESS_ALL_ACCESS, true, Lp.th32ProcessID);
        break;
      end
    until Process32Next(Hd, Lp) = False;
end;

function ShowMsbInfo(AMsg: string; IType: string = 'info'): Boolean;
begin
  Result := False;
  if UpperCase(IType) = 'OK' then  //成功
    Application.MessageBox(Pchar(AMsg), '完成', MB_ICONINFORMATION + MB_OK)
  else if UpperCase(IType) = 'ERROR' then //错误
    Application.MessageBox(Pchar(AMsg), '错误', MB_ICONSTOP + MB_OK)
  else if UpperCase(IType) = 'ASK' then  //询问
    if Application.MessageBox(Pchar(AMsg), '确认', MB_ICONQUESTION + MB_YESNO) = IDYES then
      result := true
    else
      result := false
  else if UpperCase(IType) = 'INFO' then  //提示
    Application.MessageBox(Pchar(AMsg), '提示', MB_ICONINFORMATION + MB_OK)
  else if UpperCase(IType) = 'WARN' then  //警告
    Application.MessageBox(Pchar(AMsg), '警告', MB_ICONWARNING + MB_OK);
end;

function ShowMessageBoxEx(AMsg: string; IType: string = 'info'): Boolean;
begin
  Result := False;
  HideLodingMessage;
  MyMsg_IconIndex := UpperCase(IType);
  if not Assigned(frmCommonMsg) then
    frmCommonMsg := TfrmCommonMsg.Create(Application);
  try
    frmCommonMsg.pnlMsg.Caption := AMsg;
    frmCommonMsg.Update;
    frmCommonMsg.ShowModal;
  finally
    Result := MyMsg_ButtonResult;
    FreeAndNil(frmCommonMsg);
  end;
end;

function ShowMessageBoxEx(AMsg: string; Flags: Longint): Boolean;
begin
  Result := False;
  HideLodingMessage;
  case Flags of
    MB_OK:
      MyMsg_IconIndex := 'OK';
    MB_YESNO:
      MyMsg_IconIndex := 'ASK';
    MB_ICONERROR:
      MyMsg_IconIndex := 'ERROR';
    MB_ICONWARNING:
      MyMsg_IconIndex := 'WARN';
  else
    MyMsg_IconIndex := 'INFO';
  end;

  if not Assigned(frmCommonMsg) then
    frmCommonMsg := TfrmCommonMsg.Create(Application);
  try
    frmCommonMsg.pnlMsg.Caption := AMsg;
    frmCommonMsg.Update;
    frmCommonMsg.ShowModal;
  finally
    Result := MyMsg_ButtonResult;
    FreeAndNil(frmCommonMsg);
  end;
end;

procedure HideTaskbar; //隐藏
var
  wndHandle: THandle;
  wndClass: array[0..50] of Char;
begin
  StrPCopy(@wndClass[0], 'Shell_TrayWnd');
  wndHandle := FindWindow(@wndClass[0], nil);
  ShowWindow(wndHandle, SW_HIDE);
end;

procedure ShowTaskbar; //显示
var
  wndHandle: THandle;
  wndClass: array[0..50] of Char;
begin
  StrPCopy(@wndClass[0], 'Shell_TrayWnd');
  wndHandle := FindWindow(@wndClass[0], nil);
  ShowWindow(wndHandle, SW_RESTORE);
end;

procedure DisableTask(Key: Boolean);  //禁止
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create;
  Reg.RootKey := HKEY_CURRENT_USER;
  if Reg.OpenKey('\Software\MicroSoft\Windows\CurrentVersion\Policies\System', True) then
  begin
    if Key then
      Reg.WriteString('DisableTaskMgr', '1')
    else
      Reg.WriteInteger('DisableTaskMgr', 0);
    Reg.CloseKey;
  end;
  Reg.Free;
end;

//**********显示数字于一个框内start**********//
procedure DrawSegment(Canvas: TCanvas; dSegNum: Integer; const SegmentRect: TRect);
const
  cBorderGap = 1;  //边界宽度
  cSegmentThickness = 2; //每个线条宽度
  cHorzMargine = 1; //水平线条边界缩进
  cVertMargine = 1; //竖直线条边界缩进
var
  Ht, Lt, Rt, Tp, Bt, VertCentre, SegHalf: Integer;
  fPoints: array[1..6] of TPoint;
begin
  Ht := SegmentRect.Bottom - SegmentRect.Top;
  Lt := SegmentRect.Left + cBorderGap;
  Rt := SegmentRect.Right - cBorderGap - 1 - (cSegmentThickness * 2); //move in a seg and a half
  Tp := SegmentRect.Top + cBorderGap;
  Bt := Ht - cBorderGap - 1 + SegmentRect.Top;

  VertCentre := ((Bt - Tp) div 2);
  SegHalf := (cSegmentThickness div 2);

  case dSegNum of
    0:
      begin
        fPoints[1].x := Rt + cSegmentThickness{SegHalf};
        fPoints[1].y := Bt - (VertCentre div 2);// - (cSegmentThickness * 2);
        fPoints[2].x := Rt + SegHalf + cSegmentThickness * 2;// - (SegHalf + cSegmentThickness);
        fPoints[2].y := Bt;
        Canvas.Ellipse(fPoints[1].x, fPoints[1].y, fPoints[2].x, fPoints[2].y);
      end;
    3:
      begin
        fPoints[1].x := Lt + cHorzMargine;
        fPoints[1].y := Tp;

        fPoints[2].x := Rt - cHorzMargine;
        fPoints[2].y := Tp;
        fPoints[3].x := fPoints[2].x - cSegmentThickness;
        fPoints[3].y := fPoints[2].y + cSegmentThickness;
        fPoints[4].x := fPoints[1].x + cSegmentThickness;
        fPoints[4].y := fPoints[1].y + cSegmentThickness;
        Canvas.Polygon(Slice(fPoints, 4));
      end;
    4:
      begin
        fPoints[1].x := Lt;
        fPoints[1].y := Tp + cVertMargine;
        fPoints[2].x := Lt;
        fPoints[2].y := Tp + VertCentre - cVertMargine;
        fPoints[3].x := fPoints[2].x + cSegmentThickness;
        fPoints[3].y := fPoints[2].y - cSegmentThickness;
        fPoints[4].x := fPoints[1].x + cSegmentThickness;
        fPoints[4].y := fPoints[1].y + cSegmentThickness;
        Canvas.Polygon(Slice(fPoints, 4));
      end;
    5:
      begin
        fPoints[1].x := Lt;
        fPoints[1].y := Tp + cVertMargine + VertCentre;
        fPoints[2].x := Lt;
        fPoints[2].y := Tp + VertCentre - cVertMargine + VertCentre;
        fPoints[3].x := fPoints[2].x + cSegmentThickness;
        fPoints[3].y := fPoints[2].y - cSegmentThickness;
        fPoints[4].x := fPoints[1].x + cSegmentThickness;
        fPoints[4].y := fPoints[1].y + cSegmentThickness;
        Canvas.Polygon(Slice(fPoints, 4));
      end;
    6:
      begin
        fPoints[1].x := Lt + cHorzMargine;
        fPoints[1].y := Tp + VertCentre + VertCentre;
        fPoints[2].x := Rt - cHorzMargine;
        fPoints[2].y := fPoints[1].y;
        fPoints[3].x := fPoints[2].x - cSegmentThickness;
        fPoints[3].y := fPoints[2].y - cSegmentThickness;
        fPoints[4].x := fPoints[1].x + cSegmentThickness;
        fPoints[4].y := fPoints[1].y - cSegmentThickness;
        Canvas.Polygon(Slice(fPoints, 4));
      end;
    2:
      begin

        fPoints[1].x := Rt;
        fPoints[1].y := Tp + cVertMargine;
        fPoints[2].x := fPoints[1].x;
        fPoints[2].y := Tp + VertCentre - cVertMargine;
        fPoints[3].x := fPoints[2].x - cSegmentThickness;
        fPoints[3].y := fPoints[2].y - cSegmentThickness;
        fPoints[4].x := fPoints[1].x - cSegmentThickness;
        fPoints[4].y := fPoints[1].y + cSegmentThickness;
        Canvas.Polygon(Slice(fPoints, 4));
      end;
    7:
      begin
        fPoints[1].x := Rt;
        fPoints[1].y := Tp + cVertMargine + VertCentre;
        fPoints[2].x := fPoints[1].x;
        fPoints[2].y := Tp + VertCentre - cVertMargine + VertCentre;
        fPoints[3].x := fPoints[2].x - cSegmentThickness;
        fPoints[3].y := fPoints[2].y - cSegmentThickness;
        fPoints[4].x := fPoints[1].x - cSegmentThickness;
        fPoints[4].y := fPoints[1].y + cSegmentThickness;
        Canvas.Polygon(Slice(fPoints, 4));
      end;
    1:
      begin
        fPoints[1].x := Lt + cHorzMargine;
        fPoints[1].y := Tp + VertCentre;
        fPoints[2].x := fPoints[1].x + SegHalf;
        fPoints[2].y := Tp + VertCentre - SegHalf; // 1 is Pen size
        fPoints[3].x := Rt - cHorzMargine - SegHalf;
        fPoints[3].y := fPoints[2].y;
        fPoints[4].x := Rt - cHorzMargine;
        fPoints[4].y := fPoints[1].y;
        fPoints[5].x := Rt - cHorzMargine - SegHalf;
        fPoints[5].y := Tp + VertCentre + SegHalf;
        fPoints[6].x := fPoints[2].x;
        fPoints[6].y := fPoints[5].y;
        Canvas.Polygon(fPoints);
      end;
  end;
end;

procedure ShowDigiInRect(Canvas: TCanvas; mRect: TRect; str: string; Color: TColor); //显示数字在一个框里

  function Padr(s: string; numPad: Integer): string;
  var
    i, l: Integer;
  begin
    Result := '';
    l := numPad - Length(s);
    for i := 1 to l do
      Result := Result + '';
    Result := Result + s;
  end;

  procedure MakeSegments(dChar: Char; var dSegment: TCharSegment);
  begin
    dSegment := [];
    case dChar of
      '1':
        dSegment := [2, 7];
      '2':
        dSegment := [3, 2, 1, 5, 6];
      '3':
        dSegment := [3, 2, 1, 7, 6];
      '4':
        dSegment := [4, 1, 2, 7];
      '5':
        dSegment := [3, 4, 1, 7, 6];
      '6':
        dSegment := [3, 4, 5, 6, 7, 1];
      '7':
        dSegment := [3, 2, 7];
      '8':
        dSegment := [3, 4, 5, 6, 7, 2, 1];
      '9':
        dSegment := [3, 4, 1, 2, 7];
      '0':
        dSegment := [3, 4, 5, 6, 7, 2];
      '-':
        dSegment := [1];
      '.':
        dSegment := [0];
    end;
  end;

  procedure DrawSegments(Canvas: TCanvas; dSegment: TCharSegment; SegmentRect: TRect);
  var
    i: Byte;
  begin
    for i := 0 to 7 do
      if (byte(i) in dSegment) then
        DrawSegment(Canvas, i, SegmentRect);
  end;

const
  cNumChars = 2;
var
  MySeg: TCharSegment;
  xPos, i: Integer;
  MyRect: TRect;
  csize: Integer;
  s: string;
  clFront, clBack: TColor;
begin
  s := Padr(str, cNumChars);
  csize := (mRect.Right - mRect.Left) div Length(s);
  xPos := 0;
  clFront := Canvas.Pen.Color;
  clBack := Canvas.Brush.Color;
  Canvas.Brush.Color := Color;//clPurple; // 添充色
  Canvas.Pen.Color := Color;//clPurple;//clLime; //前景色

  for i := 1 to Length(s) do
  begin
    MakeSegments(s[i], MySeg);
    MyRect.Top := mRect.Top;
    MyRect.Bottom := mRect.Bottom;
    MyRect.Left := mRect.Left + (xPos) * csize;
    MyRect.Right := MyRect.Left + csize;
    DrawSegments(Canvas, MySeg, MyRect);
    Inc(xPos);
  end;
  Canvas.Pen.Color := clBlack;//clFront;
  Canvas.Brush.Color := clBack;
end;

function OpenDataSet(qry: TADOQuery; sqlTxt: string): Boolean;
var
  FErrorInfo: string;
begin
  Result := False;
  if Trim(sqlTxt) = '' then
  begin
    ShowMessageBoxEx('无SQL查询语句!', 'info');
    Exit;
  end;

  with qry do
  try
    Close;
    SQL.Clear;
    SQL.Add(sqlTxt);
    Open;
    Result := True;
  except
    on e: Exception do
    begin
      Result := False;
      FErrorInfo := e.Message;
      FErrorInfo := GetSQLErrorChineseInfo(FErrorInfo);
      WriteSystemLog('OpenDataSet:' + sqlTxt + #13#10 + FErrorInfo);
      ShowMessageBoxEx(FErrorInfo, 'info');
    end;
  end;
end;


function OpenDataSet(qry: TFDQuery; sqlTxt: string): Boolean;
var
  FErrorInfo: string;
begin
  Result := False;

  if Trim(sqlTxt) = '' then
  begin
    ShowMessageBoxEx('无SQL查询语句!', 'info');
    Exit;
  end;

  with qry do
  try
    Close;
    //FetchOptions.RecordCountMode := cmVisible;
    SQL.Clear;
    SQL.Add(sqlTxt);
    Open;
    Result := True;
  except
    on e: Exception do
    begin
      qry.Close;
      FErrorInfo := e.Message;
      WriteSystemLog('OpenDataSet:' + sqlTxt + #13#10 + FErrorInfo);
      ShowMessageBoxEx(FErrorInfo, 'info');
    end;
  end;
end;

function OpenDataSetSync(qry: TADOQuery; sqlTxt: string): Boolean;
var
  FErrorInfo: string;
begin
  Result := False;
  if Trim(sqlTxt) = '' then
  begin
    Exit;
  end;

  with qry do
  try
    Close;
    SQL.Clear;
    SQL.Add(sqlTxt);
    Open;
    Result := True;
  except
    on e: Exception do
    begin
      Result := False;
      FErrorInfo := e.Message;
      FErrorInfo := GetSQLErrorChineseInfo(FErrorInfo);
      WriteSystemLog('OpenDataSet:' + sqlTxt + #13#10 + FErrorInfo);

      TThread.Synchronize(nil,
        procedure
        begin
          ShowMessageBoxEx(FErrorInfo,MB_ICONINFORMATION);;
        end);
    end;
  end;
end;

function OpenDataSetSync(qry: TFDQuery; sqlTxt: string): Boolean;
var
  FErrorInfo: string;
begin
  Result := False;

  if Trim(sqlTxt) = '' then
  begin
    Exit;
  end;

  with qry do
  try
    Close;
    SQL.Clear;
    SQL.Add(sqlTxt);
    Open;
    Result := True;
  except
    on e: Exception do
    begin
      qry.Close;
      FErrorInfo := e.Message;
      WriteSystemLog('OpenDataSet:' + sqlTxt + #13#10 + FErrorInfo);
      TThread.Synchronize(nil,
        procedure
        begin
          ShowMessageBoxEx(FErrorInfo,MB_ICONINFORMATION);;
        end);
    end;
  end;
end;

function ExecuteSQL(qry: TADOQuery; sqlTxt: string): Boolean;
var
  FErrorInfo: string;
begin
  Result := True;
  if Trim(sqlTxt) = '' then
  begin
    Result := False;
    ShowMessageBoxEx('无SQL语句!', 'INFO');
    Exit;
  end;
  //更新语句，使用事务
  qry.Connection.BeginTrans;
  with qry do
  try
    Close;
    SQL.Clear;
    SQL.Add(sqlTxt);
    ExecSQL;
    qry.Connection.CommitTrans;

  except
    on e: Exception do
    begin
      Result := False;
      qry.Connection.RollbackTrans;
      FErrorInfo := e.Message;
      FErrorInfo := GetSQLErrorChineseInfo(FErrorInfo);
      WriteSystemLog('ExecuteSQL:' + sqlTxt + #13#10 + FErrorInfo);
      ShowMessageBoxEx(FErrorInfo, 'info');
    end;
  end;
end;

function ExecuteSQL(qry: TFDQuery; sqlTxt: string): Boolean;
var
  FErrorInfo: string;
begin
  Result := True;
  if Trim(sqlTxt) = '' then
  begin
    Result := False;
    ShowMessageBoxEx('无SQL语句!', 'INFO');
    Exit;
  end;

  //更新语句，使用事务
  qry.Connection.StartTransaction;
  with qry do
  try
    Close;
    SQL.Clear;
    SQL.Add(sqlTxt);
    ExecSQL;
    qry.Connection.Commit;
  except
    on e: Exception do
    begin
      Result := False;
      qry.Connection.Rollback;
      FErrorInfo := e.Message;
      FErrorInfo := GetSQLErrorChineseInfo(FErrorInfo);
      WriteSystemLog('ExecuteSQL:' + sqlTxt + #13#10 + FErrorInfo);
      ShowMessageBoxEx(FErrorInfo, 'info');
    end;
  end;
end;

function ExecuteSQLNoTran(qry: TADOQuery; sqlTxt: string): Boolean;
var
  FErrorInfo: string;
begin
  Result := True;
  if Trim(sqlTxt) = '' then
  begin
    Result := False;
    ShowMessageBoxEx('无SQL语句!', 'INFO');
    Exit;
  end;

  with qry do
  try
    Close;
    SQL.Clear;
    SQL.Add(sqlTxt);
    ExecSQL;
  except
    on e: Exception do
    begin
      Result := False;
      FErrorInfo := e.Message;
      FErrorInfo := GetSQLErrorChineseInfo(FErrorInfo);
      WriteSystemLog('ExecuteSQLNoTran:' + sqlTxt + #13#10 + FErrorInfo);
      ShowMessageBoxEx(FErrorInfo, 'info');
    end;
  end;
end;

function ExecuteSQLNoTran(qry: TFDQuery; sqlTxt: string): Boolean;
var
  FErrorInfo: string;
begin
  Result := True;

  if Trim(sqlTxt) = '' then
  begin
    Result := False;
    ShowMessageBoxEx('无SQL语句!', 'INFO');
    Exit;
  end;

  try
    with qry do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sqlTxt);
      ExecSQL;
    end;
  except
    on e: Exception do
    begin
      Result := False;
      qry.Close;
      FErrorInfo := e.Message;
      FErrorInfo := GetSQLErrorChineseInfo(FErrorInfo);
      WriteSystemLog('ExecuteSQLNoTran:' + sqlTxt + #13#10 + FErrorInfo);
      ShowMessageBoxEx(FErrorInfo, 'info');
    end;
  end;
end;

function ExecuteSQLBatch(qry: TADOQuery; lvListSQL: TStringList; execOne: Boolean = False): Boolean;
var
  FErrorInfo: string;
  I: Integer;
begin
  Result := True;

  if not Assigned(lvListSQL) then
    Exit;

  if Trim(lvListSQL.Text) = '' then
    Exit;

  try
    with qry do
    begin
      Close;
      if not qry.Connection.InTransaction then
        qry.Connection.BeginTrans;
      if execOne then
      begin
        SQL.Clear;
        SQL.Add(lvListSQL.Text);
        ExecSQL;
      end
      else
      begin
        for I := 0 to Pred(lvListSQL.Count) do
        begin
          SQL.Clear;
          SQL.Add(lvListSQL.Strings[I]);
          ExecSQL;
        end;
      end;
      qry.Connection.CommitTrans;
    end;
  except
    on e: Exception do
    begin
      Result := False;
      qry.Connection.RollbackTrans;
      FErrorInfo := e.Message;
      FErrorInfo := GetSQLErrorChineseInfo(FErrorInfo);
      WriteSystemLog('ExecuteSQLBatch:' + lvListSQL.Text + #13#10 + FErrorInfo);
      ShowMessageBoxEx(FErrorInfo, 'info');
    end;
  end;
end;

function ExecuteSQLBatch(qry: TFDQuery; lvListSQL: TStringList; execOne: Boolean = False): Boolean;
var
  FErrorInfo: string;
  I: Integer;
begin
  Result := True;

  if not Assigned(lvListSQL) then
    Exit;

  if Trim(lvListSQL.Text) = '' then
    Exit;

  try
    with qry do
    begin
      Close;
      if not qry.Connection.InTransaction then
        qry.Connection.StartTransaction;
      if execOne then
      begin
        SQL.Clear;
        SQL.Add(lvListSQL.Text);
        ExecSQL;
      end
      else
      begin
        for I := 0 to Pred(lvListSQL.Count) do
        begin
          SQL.Clear;
          SQL.Add(lvListSQL.Strings[I]);
          ExecSQL;
        end;
      end;

      qry.Connection.Commit;
    end;
  except
    on e: Exception do
    begin
      Result := False;
      qry.Connection.Rollback;
      FErrorInfo := e.Message;
      FErrorInfo := GetSQLErrorChineseInfo(FErrorInfo);
      writeSystemLog('ExecuteSQLBatch:' + lvListSQL.Text + #13#10 + FErrorInfo);
      ShowMessageBoxEx(FErrorInfo, 'info');
    end;
  end;
end;

procedure IniSystemLog(const logPath: string = '');
begin
  if Trim(logPath) = '' then
  begin
    SetDefaultLogFile(ExtractFilePath(Application.ExeName) + Format('log\runLog\%s.log', [FormatDateTime('yyyy-MM-dd', Now)]), 2097152 * 5, True, False).MaxLogHistories := 5; // 最大保留两个历史日志+1个当前日志文件
  end else
  begin
    SetDefaultLogFile(ExtractFilePath(logPath) + Format('\%s.log', [FormatDateTime('yyyy-MM-dd', Now)]), 2097152 * 5, True, False).MaxLogHistories := 5; // 最大保留两个历史日志+1个当前日志文件
  end;

end;

procedure writeSystemLog(sLogText: string; logType: Integer = 0);
begin
  try
    case logType of
      1:
        PostLog(llError, 'llError：' + sLogText);         //error
      2:
        PostLog(llHint, 'llHint：' + sLogText);          //llHint
      3:
        PostLog(llWarning, 'llWarning：' + sLogText);       //llWarning
    else
      PostLog(llMessage, 'llMessage：' + sLogText);
    end;

  except

  end;
end;

function StartDate_DD(Date: TDateTime): string;
begin
  Result := FormatDateTime('yyyy-MM-dd', StartOfTheDay(Date));
end;

function StartDate_ZZZ(Date: TDateTime): string;
begin
  Result := FormatDateTime('yyyy-MM-dd hh:mm:ss:zzz', StartOfTheDay(Date));
end;

function StartDate_SS(Date: TDateTime): string;
begin
  Result := FormatDateTime('yyyy-MM-dd hh:mm:ss', StartOfTheDay(Date));
end;

function EndDate_DD(Date: TDateTime): string;
begin
  Result := FormatDateTime('yyyy-MM-dd', EndOfTheDay(Date));
end;

function EndDate_ZZZ(Date: TDateTime): string;
begin
  Result := FormatDateTime('yyyy-MM-dd hh:mm:ss:zzz', EndOfTheDay(Date));
end;

function EndDate_SS(Date: TDateTime): string;
begin
  Result := FormatDateTime('yyyy-MM-dd hh:mm:ss', EndOfTheDay(Date));
end;

procedure RefreshQuery(qry: TADOQuery); //刷新ADO数据集
begin
  qry.Close;
  qry.Open;
end;

//刷新数据集，刷新后返回到以前状态(记录位置)
procedure RefreshDataSet(DataSet: TDataSet);
var
  BM: TBookmark;
begin
  with DataSet do
  begin
    DisableControls;
    try
      if Active then
        BM := GetBookmark
      else
        BM := nil;
      try
        Close;
        Open;
        if (BM <> nil) and not (Bof and Eof) and BookmarkValid(BM) then
        try
          GotoBookmark(BM);
        except
        end;
      finally
        if BM <> nil then
          FreeBookmark(BM);
      end;
    finally
      EnableControls;
    end;
  end;
end;

procedure RefreshDataSetEX(DataSet: TDataSet; bDisableControls: Boolean);
var
  BM: TBookMark;
begin
  if bDisableControls then
    DataSet.DisableControls;
  try
    BM := DataSet.GetBookMark;
    try
      DataSet.Close;
      DataSet.Open;
      if (not DataSet.IsEmpty) and Assigned(BM) then
      try
        if DataSet.BookmarkValid(BM) then  //cyw 08.4.29
          DataSet.GotoBookMark(BM)
        else
          DataSet.Last;
      except
        DataSet.First;
      end;
    finally
      DataSet.FreeBookMark(BM);
    end;
  finally
    if bDisableControls then   //cyw 08.8.13
      DataSet.EnableControls;
  end;
end;

//刷新数据集MebTableEh，刷新后返回到以前状态(记录位置)
procedure RefreshMemTable(DataSet: TDataSet);
var
  BM: TBookmark;
begin
  with DataSet do
  begin
    if DataSet.RecordCount > 0 then
      BM := GetBookmark
    else
      BM := nil;
    try
      Close;
      Active := True;
      if (BM <> nil) and not (Bof and Eof) and BookmarkValid(BM) then
      try
        GotoBookmark(BM);
      except
      end;
    finally
      if BM <> nil then
        FreeBookmark(BM);
    end;
  end;
end;

procedure SetSysDateFormat;
begin
  // 设定程序本身所使用的日期时间格式
  with FormatSettings do
  begin
    ShortDateFormat := 'yyyy-mm-dd';
    LongDateFormat := 'yyyy-mm-dd';
    ShortTimeFormat := 'hh:nn:ss';
    LongTimeFormat := 'hh:nn:ss';
    DateSeparator := '-';
    TimeSeparator := ':';
  end;
end;

procedure IntegerOnly(var Key: Char);
begin
  if not (Key in ['0'..'9', ^C, ^V, Chr(VK_BACK), Chr(VK_RETURN)]) then
    Key := #0;
end;

procedure NumberOrLetterOnly(var Key: Char);
begin
  if not (Key in ['0'..'9', 'a'..'z', 'A'..'Z', #8, #13, ^C, ^V, Chr(VK_BACK), Chr(VK_RETURN)]) then
    Key := #0;
end;

procedure FloatOnly(var Key: Char; iText: string);
begin
  if not (Key in ['0'..'9', '.', '-', ^C, ^V, Chr(VK_BACK), Chr(VK_RETURN)]) then
    Key := #0;

  if (Key = '.') and (Pos('.', iText) > 0) then
    Key := #0;
end;

function IsNumber(str: string): Boolean;
var
  i: Integer;
begin
  Result := False;
  if Trim(str) = '' then
    Exit;

  Result := True;
  for i := 1 to Length(str) do
    if not (str[i] in ['0'..'9']) then
    begin
      Result := False;
      Exit;
    end;
end;

function IsFloat(str: string): Boolean;
var
  i, j, k: Integer;
  tmpStr: string;
begin
  Result := True;
  if Trim(str) = '' then
  begin
    result := False;
    Exit;
  end;
  j := 0;
  k := 0;
  //只能有一个.号或-号
  for i := 1 to Length(str) do
  begin
    if not (str[i] in ['0'..'9', '.', '-']) then
    begin
      Result := False;
      Exit;
    end;
    if (str[i] = '.') then
      Inc(j);
    if (str[i] = '-') then
      Inc(k);
  end;
  if j > 1 then
  begin
    Result := False;
    Exit;
  end;
  if k > 1 then
  begin
    Result := False;
    Exit;
  end;

  //-号只能在首位
  if Pos('-', str) > 1 then
  begin
    Result := False;
    Exit;
  end;

  //首尾不能为.
  if (str[1] = '.') or (str[Length(str)] = '.') then
  begin
    Result := False;
    Exit;
  end;

  //整数最高位不能连续为零;
  tmpStr := Copy(str, 1, Pos('.', str) - 1);
  j := 0;
  if Length(str) > 1 then
    for i := 1 to Length(tmpStr) do
      if (tmpStr[1] = '0') and (tmpStr[2] = '0') then
      begin
        Result := False;
        Exit;
      end;

  //如果最高位为零，则必须为个位数的小数
  tmpStr := Copy(str, 1, Pos('.', str) - 1);
  if Length(tmpStr) > 1 then
    for i := 1 to Length(tmpStr) do
      if tmpStr[1] = '0' then
      begin
        Result := False;
        Break;
        Exit;
      end;
end;

function IsNumeric(Value: string; const AllowFloat: Boolean; const TrimWhiteSpace: Boolean = True): Boolean;
var
  ValueInt: Int64;
  ValueFloat: Extended;
begin
  if TrimWhiteSpace then
    Value := SysUtils.Trim(Value);
  // Check for valid integer
  Result := SysUtils.TryStrToInt64(Value, ValueInt);
  if not Result and AllowFloat then
    // Wasn't valid as integer, try float
    Result := SysUtils.TryStrToFloat(Value, ValueFloat);
end;

function FormatFloatStr(value: string; DisplayFormat: string = '#0.00'):string;
begin
  Result := FormatFloat(DisplayFormat, StrToFloatDef(value, 0))
end;

function GetServerDateTime(qry: TADOQuery): TDateTime;
begin
  with qry do
  begin
    qry.Close;
    SQL.Text := 'Select GetDate() As ServerDateTime';
    Open;
    Result := FieldByName('ServerDateTime').AsDateTime;
    Close;
  end;
end;

function GetServerDateTime(qry: TFDQuery): TDateTime;
begin
  with qry do
  begin
    qry.Close;
    SQL.Text := 'Select GetDate() As ServerDateTime';
    Open;
    Result := FieldByName('ServerDateTime').AsDateTime;
    Close;
  end;
end;

function GetMonthOfDate(ReturnType: Integer; incYear: Integer = 1): TDateTime; overload;

  function getMonth(iMonth: Integer): Integer;
  begin
    if not (iMonth in [1..12]) then
      Result := 12
    else
      Result := iMonth;
  end;

var
  Year, Month, Day1, Day2: Word;
  NewDate: TDateTime;
begin
  DecodeDate(Now, Year, Month, Day1);
  case ReturnType of
    0:
      NewDate := StartOfTheDay(StartOfTheMonth(Now));                 // This Month First Day
    1:
      NewDate := EndOfTheDay(EndOfTheMonth(Now));                   // This Month Last Day
    3:                                                  // Last Month First Day
      begin
        DecodeDate(Now - Day1, Year, Month, Day2);
        NewDate := StartOfTheDay(StrToDate(FormatDateTime('yyyy-MM-dd', Now - Day1 - Day2 + 1)));
      end;
    4:
      NewDate := StartOfTheDay(EncodeDate(Year, getMonth(Month - 1), Day1));      // Last Month Today
    5:
      NewDate := EndOfTheDay(EncodeDate(Year, Month, 1) - 1);            // Last Month Last Day
    6:
      NewDate := StartOfTheDay(EncodeDate(Year - 1, Month, 1));        // Last Year This Day
    7:
      NewDate := EndOfTheDay(EncodeDate(Year - 1, getMonth(Month - 1), Day1)); // Last Year Last Month Today
    8:
      NewDate := EndOfTheDay(EncodeDate(Year - 1, Month, 1) - 1);      // Last Year Last Month Last Day
    9:
      NewDate := StartOfTheDay(EncodeDate(Year, Month, Day1));         // This Month This Day
    10:
      NewDate := StartOfTheDay(EncodeDate(Year, Month, Day1) - 1);     // Yesterday
    11:
      NewDate := StartOfTheDay(EncodeDate(Year, 1, 1));               // This Year First Day
    12:
      NewDate := EndOfTheDay(EncodeDate(Year, 12, 31));             // This Year Last Day
  end;
  Result := NewDate;
end;

function GetMonthOfDate(InputDate: TDateTime; ReturnType: Integer): TDateTime; overload;
var
  Year, Month, Day1, Day2: Word;
  NewDate: TDateTime;
begin
  DecodeDate(InputDate, Year, Month, Day1);
  case ReturnType of
    0:
      begin
        NewDate := StartOfTheDay(EncodeDate(Year, Month, 1));           // This InputDate Month First Day
      end;
    1:
      begin
        DecodeDate(IncMonth(InputDate), Year, Month, Day1);
        NewDate := EndOfTheDay(EncodeDate(Year, Month, 1) - 1);         // This InputDate Month Last Day
      end;
    3:                                                  // Last InputDate Month First Day
      begin
        DecodeDate(Now - Day1, Year, Month, Day2);
        NewDate := StartOfTheDay(StrToDate(FormatDateTime('yyyy-MM-dd', Now - Day1 - Day2 + 1)));
      end;
    4:
      begin
        DecodeDate(IncMonth(InputDate, -1), Year, Month, Day1);
        NewDate := StartOfTheDay(EncodeDate(Year, Month, Day1));            // Last InputDate Month Today
      end;
    5:
      NewDate := EndOfTheDay(EncodeDate(Year, Month, 1) - 1);            // Last InputDate Month Last Day
    6:
      NewDate := StartOfTheDay(EncodeDate(Year - 1, Month, 1));        // Last InputDate Year This Day
    7:
      begin
        DecodeDate(IncMonth(InputDate, -13), Year, Month, Day1);
        NewDate := EndOfTheDay(EncodeDate(Year, Month, Day1)); // Last InputDate Year Last Month Today
      end;
    8:
      begin
        DecodeDate(IncMonth(InputDate, -12), Year, Month, Day1);
        NewDate := EndOfTheDay(EncodeDate(Year, Month, 1) - 1);      // Last InputDate Year Last Month Last Day
      end;
    9:
      NewDate := StartOfTheDay(EncodeDate(Year, Month, Day1));         // This InputDate Month This Day
    10:
      NewDate := StartOfTheDay(EncodeDate(Year, Month, Day1) - 1);     //InputDate Yesterday
    11:
      NewDate := StartOfTheDay(EncodeDate(Year, 1, 1));               // InputDate Year First Day
    12:
      NewDate := EndOfTheDay(EncodeDate(Year, 12, 31));             // InputDate Year Last Day
  end;
  Result := NewDate;
end;

function myDateTimeToUnix(ConvDate: TDateTime; len: Integer = 10): Int64;
var
  zi: TTimeZoneInformation;
  LDate: TDateTime;
begin
  case len of
    10:
      Result := DateTimeToUnix(ConvDate) - 8 * 60 * 60;
  else
    LDate := TTimeZone.local.ToUniversalTime(ConvDate);

    Result := MilliSecondsBetween(UnixDateDelta, LDate);
  end;

end;

function myUnixToDateTime(USec: Int64): TDateTime;
const
  cUnixStartDate: TDateTime = 25569.0; // 1970/01/01
var
  zi: TTimeZoneInformation;
begin
  GetTimeZoneInformation(zi);
  Result := ((USec - 60 * zi.Bias) / 86400) + cUnixStartDate;
end;

function ConvertStrToDateTime(DateTimeStr: string; var ADateTime: TDateTime): boolean;
  //判断是否是日期时间

  function StrIsValidDateTime(DateTimeStr: string; var Y, M, D, H, N, S: Integer): boolean;
  begin
    Result := False;
    if TryStrToInt(Copy(DateTimeStr, 1, 4), Y) then
    begin
      if TryStrToInt(Copy(DateTimeStr, 5, 2), M) then
      begin
        if TryStrToInt(Copy(DateTimeStr, 7, 2), D) then
        begin
          if TryStrToInt(Copy(DateTimeStr, 9, 2), H) then
          begin
            if TryStrToInt(Copy(DateTimeStr, 11, 2), N) then
            begin
              if TryStrToInt(Copy(DateTimeStr, 13, 2), S) then
              begin
                Result := IsValidDateTime(Y, M, D, H, N, S, 0);
              end;
            end;
          end;
        end;
      end;
    end;
  end;

var
  Y, M, D, H, N, S: Integer;
begin
  Result := False;
  if StrIsValidDateTime(DateTimeStr, Y, M, D, H, N, S) then
  begin
    ADateTime := EncodeDatetime(Y, M, D, H, N, S, 0);
    Result := True;
  end;
end;

function IsCheckNull(Container: TWinControl; DefaultCheckBox: TCheckBox = nil): Boolean;
var
  AControl: TControl;
  I: Integer;
begin
  Result := False;
  with Container do
  begin
    for I := 0 to ControlCount - 1 do
    begin
      AControl := Controls[I];
      if AControl is TCheckBox then
        Result := Result or TCheckBox(AControl).Checked;//等于 Result := TCheckBox(AControl).Checked；
      if Result then
        Exit;
    end;
  end;
  if AsSigned(DefaultCheckBox) then
    DefaultCheckBox.Checked := True;
  raise Exception.Create('为了提高查询速度，请至少选择一个查询条件！');
end;

//字符转成UTF-8编码
function DecodeUtf8Str(const S: UTF8String): WideString;
var
  lenSrc, lenDst: Integer;
begin
  lenSrc := Length(S);
  if (lenSrc = 0) then
    Exit;
  lenDst := MultiByteToWideChar(CP_UTF8, 0, Pointer(S), lenSrc, nil, 0);
  SetLength(Result, lenDst);
  MultiByteToWideChar(CP_UTF8, 0, Pointer(S), lenSrc, Pointer(Result), lenDst);
end;

//字符转换成十六进制
function myStrToHex(s: string): string;
var
  TmpStr: string;
  i: integer;
begin
  TmpStr := '';
  for i := 1 to Length(s) do
  begin
    TmpStr := TmpStr + IntToHex(ord(s[i]), 2);
  end;
  Result := TmpStr;
end;

//十六进制字符转普通字符
function myHexToStr(s: string): string;
var
  HexS, TmpStr: string;
  i: Integer;
  a: Byte;
begin
  HexS := s;
  if Length(HexS) mod 2 = 1 then
  begin
    HexS := HexS + '0';
  end;
  TmpStr := '';
  for i := 1 to (Length(HexS) div 2) do
  begin
    a := StrToInt('$' + HexS[2 * i - 1] + HexS[2 * i]);
    TmpStr := TmpStr + Chr(a);
  end;
  Result := TmpStr;
end;

//查找某一字条串在源串中出现的次数
function SubStrCount(MainStr, SubStr: string): Integer;
var
  iCount, i: integer;
  Str: string;
begin
  Result := 0;
  if Length(MainStr) = 0 then
    Exit;
  iCount := 0;
  Str := trim(MainStr);
  for i := 0 to StrLen(Pchar(Str)) do
  begin
    if Copy(Str, i, 1) = SubStr then
      iCount := iCount + 1;
  end;
  Result := iCount;
end;

//通过分隔符获取若干子串
function SplitStr(const Source, Split: string): TStringList;
var
  Temp: string;
  i: Integer;
begin
  Result := TStringList.Create;
  if Source = '' then
    exit;
  Temp := Source;
  i := pos(Split, Source);
  while i <> 0 do
  begin
    Result.add(copy(Temp, 0, i - 1));
    Delete(Temp, 1, i + Length(Split) - 1);
    i := Pos(Split, Temp);
  end;
  Result.Add(Temp);
end;

//将字符串中的指定子串替换成特定子串，返回一个新字符串
function SubStrReplace(SourceStr, OldStr, NewStr: string): string;
var
  Instr, SubLen, MainLen: Integer;
  Templ, Tempr: string;
begin
  SubLen := length(OldStr);
  MainLen := length(SourceStr);
  if (SubLen = 0) or (SubLen > MainLen) or (OldStr = NewStr) then
  begin
    Result := SourceStr;
    Exit;
  end;
  while pos(OldStr, SourceStr) <> 0 do
  begin
    Instr := pos(OldStr, SourceStr);
    Templ := Copy(SourceStr, 1, Instr - 1);
    Tempr := Copy(SourceStr, Instr + SubLen, MainLen);
    SourceStr := Templ + NewStr + Tempr;
  end;
  Result := SourceStr;
end;

//将字符串中的指定子串替换成特定子串，返回一个新字符串（秒杀Memo的回车，换行，空格特殊字符）
procedure ReplaceMeoToStr(var s: string; const SourceChar: pchar; const RChar: pchar);
//传入字符串，要替换的字符，替换后字符
var
  ta, i, j: integer;
  m, n, pn, sn: integer;
  SLen, SCLen, RCLen: integer; //SLen传入字符串的长度，SCLen表示要替换字符串的长度，RCLen表示替换后字符串的长度
  IsSame: integer;
  newp: array of char; //保存替换后的字符数组
begin
  SLen := strlen(pchar(s));
  SCLen := strlen(SourceChar);
  RCLen := strlen(RChar);
  j := pos(string(SourceChar), s);
  s := s + chr(0);
  ta := 0;
  i := j;
  while s[i] <> chr(0) do   //这个循环用ta统计模式串在原串中出现的次数
  begin
    n := 0;
    IsSame := 1;
    for m := i to i + SCLen - 1 do
    begin
      if m > SLen then
      begin
        IsSame := 0;
        break;
      end;
      if s[m] <> SourceChar[n] then
      begin
        IsSame := 0;
        break;
      end;
      n := n + 1;
    end;
    if IsSame = 1 then
    begin
      ta := ta + 1;
      i := m;
    end
    else
      i := i + 1;
  end;
  if j > 0 then
  begin
    pn := 0;
    sn := 1;
    setlength(newp, SLen - ta * SCLen + ta * RCLen + 1);//分配newp的长度，+1表示后面还有一个#0结束符
    while s[sn] <> chr(0) do //主要循环，开始替换
    begin
      n := 0;
      IsSame := 1;
      for m := sn to sn + SCLen - 1 do //比较子串是否和模式串相同
      begin
        if m > SLen then
        begin
          IsSame := 0;
          break;
        end;
        if s[m] <> SourceChar[n] then
        begin
          IsSame := 0;
          break;
        end;
        n := n + 1;
      end;
      if IsSame = 1 then//相同
      begin
        for m := 0 to RCLen - 1 do
        begin
          newp[pn] := RChar[m];
          pn := pn + 1;
        end;
        sn := sn + SCLen;
      end
      else
      begin //不同
        newp[pn] := s[sn];
        pn := pn + 1;
        sn := sn + 1;
      end;
    end;
    newp[pn] := #0;
    s := string(newp); //重置s，替换完成！
  end;
end;

//删除StringList里面重复的项目
procedure DelStrListDuplicates(const AStringList: TStringList);
var
  buffer: TStringList;
  cnt: Integer;
begin
  AStringList.Sort;
  buffer := TStringList.Create;
  try
    buffer.Sorted := True;
    buffer.Duplicates := dupIgnore;
    buffer.BeginUpdate;
    for cnt := 0 to Pred(AStringList.Count) do
      buffer.Add(AStringList[cnt]);
    buffer.EndUpdate;
    AStringList.Assign(buffer);
  finally
    FreeandNil(buffer);
  end;
end;

//获得TstringList中没有重复的字符串
procedure GetUniqueStringList(const AStringList: TStringList);
var
  CurrentString: string;
  LastString: string;
  Count: Integer;
  Index: Integer;
  UniqueList: TStringList;
begin
  Count := 1;
  LastString := '';

  try
    UniqueList := TStringList.Create;
    AStringList.Sorted := True;
    for Index := 0 to Pred(AStringList.Count) do
    begin
      CurrentString := AStringList[Index];
      if CurrentString = LastString then
        Count := Count + 1
      else
      begin
        if LastString <> '' then
        begin
          if Count = 1 then
            UniqueList.Add(LastString)
          else
            Count := 1;
        end;
      end;
      LastString := CurrentString;
    end;
    if Count = 1 then
      UniqueList.Add(LastString);
    AStringList.Assign(UniqueList);
  finally
    UniqueList.Free;
  end;
end;

// 将数组按指定大小分组
// arrTotal: 输入的字符串数组
// groupSize: 每组的大小
// 返回: TArray<TArray<string>>，包含分组后的数组
function SplitStringArray(const arrTotal: TArray<string>; groupSize: Integer): TArray<TArray<string>>;
var
  i, groupCount, currentGroup: Integer;
  resultArray: TArray<TArray<string>>;
begin
  // 参数校验
  if (groupSize <= 0) or (Length(arrTotal) = 0) then
  begin
    SetLength(resultArray, 0);
    Exit(resultArray);
  end;

  // 计算需要的分组数
  groupCount := (Length(arrTotal) + groupSize - 1) div groupSize;
  SetLength(resultArray, groupCount);

  // 按组分配元素
  currentGroup := 0;
  for i := 0 to Length(arrTotal) - 1 do
  begin
    // 为当前组分配空间
    if (i mod groupSize = 0) then
    begin
      SetLength(resultArray[currentGroup], groupSize);
    end;

    // 将元素放入当前组
    resultArray[currentGroup][i mod groupSize] := arrTotal[i];

    // 移动到下一组
    if (i mod groupSize = groupSize - 1) then
      Inc(currentGroup);
  end;

  // 处理最后一组可能不满的情况
  if (Length(arrTotal) mod groupSize <> 0) then
  begin
    SetLength(resultArray[groupCount - 1], Length(arrTotal) mod groupSize);
  end;

  Result := resultArray;
end;

//遍历指定目录下文件。返回TStringList
procedure EnumFileInQueue(path: PChar; fileExt: string; fileList: TStringList);
var
  searchRec: TSearchRec;
  found: Integer;
  tmpStr: string;
  curDir: string;
  dirs: TQueue;
  pszDir: PChar;
begin
  dirs := TQueue.Create; //创建目录队列
  dirs.Push(path); //将起始搜索路径入队
  pszDir := dirs.Pop;
  curDir := StrPas(pszDir);
   //开始遍历,直至队列为空(即没有目录需要遍历)
  while (True) do
  begin
      //加上搜索后缀,得到类似'c:\*.*' 、'c:\windows\*.*'的搜索路径
    tmpStr := curDir + '\*.*';
      //在当前目录查找第一个文件、子目录
    found := FindFirst(tmpStr, faAnyFile, searchRec);
    while found = 0 do //找到了一个文件或目录后
    begin
          //如果找到的是个目录
      if (searchRec.Attr and faDirectory) <> 0 then
      begin
          {在搜索非根目录(C:\、D:\)下的子目录时会出现'.','..'的"虚拟目录"
          大概是表示上层目录和下层目录吧。。。要过滤掉才可以}
        if (searchRec.Name <> '.') and (searchRec.Name <> '..') then
        begin
               {由于查找到的子目录只有个目录名，所以要添上上层目录的路径
                searchRec.Name = 'Windows';
                tmpStr:='c:\Windows';
                加个断点就一清二楚了
               }
          tmpStr := curDir + '\' + searchRec.Name;
               {将搜索到的目录入队。让它先晾着。
                因为TQueue里面的数据只能是指针,所以要把string转换为PChar
                同时使用StrNew函数重新申请一个空间存入数据，否则会使已经进
                入队列的指针指向不存在或不正确的数据(tmpStr是局部变量)。}
          dirs.Push(StrNew(PChar(tmpStr)));
        end;
      end
      else //如果找到的是个文件
      begin
        if fileExt = '.*' then
          fileList.Add(curDir + '\' + searchRec.Name)
        else
        begin
          if SameText(RightStr(curDir + '\' + searchRec.Name, Length(fileExt)), fileExt) then
            fileList.Add(curDir + '\' + searchRec.Name);
        end;
      end;
          //查找下一个文件或目录
      found := FindNext(searchRec);
    end;
      {当前目录找到后，如果队列中没有数据，则表示全部找到了；
        否则就是还有子目录未查找，取一个出来继续查找。}
    if dirs.Count > 0 then
    begin
      pszDir := dirs.Pop;
      curDir := StrPas(pszDir);
      StrDispose(pszDir);
    end
    else
      break;
  end;
   //释放资源
  dirs.Free;
  FindClose(searchRec);
end;

//转换文件的时间格式
function CovFileDate(Fd: _FileTime): TDateTime;
var
  SystemTime: _SystemTime;
  TempFileTime: _FileTime;
begin
  FileTimeToLocalFileTime(Fd, TempFileTime);
  FileTimeToSystemTime(TempFileTime, SystemTime);
  CovFileDate := SystemTimeToDateTime(SystemTime);
end;

//获取文件时间，Tf表示目标文件路径和名称
function GetFileTime(const Tf: string): string;
const
  TimeFormat = 'yyyy-MM-dd hh:mm:ss';
var
  Tp: TSearchRec;
  T1, T2, T3: string;
begin
  Result := '';
  FindFirst(Tf, faAnyFile, Tp); //查找目标文件
  //返回文件的创建时间
  T1 := FormatDateTime(TimeFormat, CovFileDate(Tp.FindData.ftCreationTime));
  //返回文件的修改时间
  T2 := FormatDateTime(TimeFormat, CovFileDate(Tp.FindData.ftLastWriteTime));
  //返回文件的当前访问时间
  T3 := FormatDateTime(TimeFormat, Now);

  FindClose(Tp);
  Result := T2;
end;

//获取SQL的中文出错信息
function GetSQLErrorChineseInfo(ErrorInfo: string): string;
var
  sTemp: string;
begin
  sTemp := ErrorInfo;
  sTemp := StringReplace(sTemp, '[Microsoft][ODBC SQL Server Driver][SQL Server]', '', [rfReplaceAll]);
  sTemp := StringReplace(sTemp, '[Microsoft][ODBC SQL Server Driver][SQL'#13#10'Server]', '', [rfReplaceAll]);
  sTemp := StringReplace(sTemp, 'General SQL error.'#13#10, '', [rfReplaceAll]);
  sTemp := StringReplace(sTemp, 'Key violation.'#13#10, '', [rfReplaceAll]);
  Result := sTemp;
end;

//合计DataSet指定字段
function FieldSumValue(DataSet: TDataSet; FieldName: string): Extended;
var
  BM: TBookmark;
  fn: TField;
begin
  Result := 0.00;
  with DataSet do
  begin
    BM := GetBookmark;
    fn := FindField(FieldName);
    if not Assigned(fn) then
      Exit;

    fn := FieldByName(FieldName);  //获取字段对应的列的数据列表
    DisableControls;
    try
      First;
      while not Eof do
      begin
        Result := Result + StrToFloatDef(fn.AsString, 0);
        Next;
      end;
    finally
      GotoBookmark(BM);
      //Bookmark := BM;
      EnableControls;
      FreeBookmark(BM);
    end;
  end;
end;

function FieldSumValue(DataSet: TDataSet; arrField: TArray<string>): TArray<Extended>;
var
  i: Integer;
  BM: TBookmark;
  fn: TField;
begin
  SetLength(result, Length(arrField));
  for i := Low(result) to High(result) do
    Result[i] := 0.00;

  with DataSet do
  begin
    BM := GetBookmark;

    DisableControls;
    First;
    try
      while not Eof do
      begin
        for i := Low(arrField) to High(arrField) do
        begin
          fn := FindField(arrField[i]);  //获取字段对应的列的数据列表
          if Assigned(fn) then
            Result[i] := Result[i] + StrToFloatDef(fn.AsString, 0);
        end;

        Next;
      end;
    finally
      GotoBookmark(BM);
      EnableControls;
      FreeBookmark(BM);
    end;
  end;

end;

//eg:FieldSumValueEx(ADataset,'defaultMoney','water_bill_date',FieldByName('water_bill_date').AsString)
function FieldSumValueEx(DataSet: TDataSet; FieldName: string; ValuedName, Value: string; calc: char): Extended;
var
  BM: TBookmark;
  fn, tmpFN: TField;
begin
  Result := 0;
  with DataSet do
  begin
    BM := GetBookmark;
    fn := DataSet.Fields.FindField(FieldName);
    tmpFN := DataSet.Fields.FindField(ValuedName);
    //fn := FieldByName(FieldName);
    if Assigned(fn) then
    begin
      DisableControls;
      try
        First;
        while not Eof do
        begin
          if Assigned(tmpFN) then
          begin
            if calc = '=' then
            begin
              if SameText(tmpFN.AsString, Value) then
                Result := Result + StrToFloatDef(fn.AsString, 0);
            end
            else
            begin
              if not SameText(tmpFN.AsString, Value) then
                Result := Result + StrToFloatDef(fn.AsString, 0);
            end;
          end;
          Next;
        end;
      finally
        GotoBookmark(BM);
        //Bookmark := BM;
        EnableControls;
        FreeBookmark(BM);
      end;
    end;
  end;
end;

function FieldSumValueEx(DataSet: TDataSet; FieldName: string; Value: string; calc: char): Extended;
var
  BM: TBookmark;
  fn: TField;
begin
  Result := 0;
  with DataSet do
  begin
    BM := GetBookmark;
    fn := DataSet.Fields.FindField(FieldName);
    //fn := FieldByName(FieldName);
    if Assigned(fn) then
    begin
      DisableControls;
      try
        First;
        while not Eof do
        begin
          if calc = '=' then
          begin
            if SameText(fn.AsString, Value) then
              Result := Result + StrToFloatDef(fn.AsString, 0);
          end
          else
          begin
            if not SameText(fn.AsString, Value) then
              Result := Result + StrToFloatDef(fn.AsString, 0);
          end;

          Next;
        end;
      finally
        GotoBookmark(BM);
        //Bookmark := BM;
        EnableControls;
        FreeBookmark(BM);
      end;
    end;
  end;
end;
//显示表格指定（非数字型）列的合计值
//数字型字段请直接使用对应表格的Footer属性

function FieldSumValueFooter(grd: TDBGridEh; FieldName: string; DisplayFormat: string = '#0.00'): string;
var
  cf: TColumnFooterEh;
begin
  Result := '0.00';
  with grd do
  try
    if not Assigned(FindFieldColumn(FieldName)) then
      Exit;

    grd.DataSource.DataSet.DisableControls;

    cf := FieldColumns[FieldName].Footer;  //获取字段对应的表脚单元格
    cf.ValueType := fvtStaticText;         //以静态文本的方式显示
    cf.Alignment := taRightJustify;        //居右显示
    cf.Value := FormatFloat(DisplayFormat, //格式化
      FieldSumValue(DataSource.DataSet, FieldName));
  finally
    grd.DataSource.DataSet.EnableControls;
  end;
  Result := cf.Value;
end;

function FieldSumValueFooter(grd: TDBGridEh; arrField: TArray<string>; DisplayFormat: string = '#0.00'): string;
var
  cf: TColumnFooterEh;
  i: Integer;
  arrFloat: TArray<Extended>;
begin
  Result := '0.00';
  with grd do
  try
    SetLength(arrFloat, Length(arrField));
    arrFloat := FieldSumValue(DataSource.DataSet, arrField);
    grd.DataSource.DataSet.DisableControls;

    for i := Low(arrField) to High(arrField) do
    begin
      if not Assigned(FindFieldColumn(arrField[i])) then
        Continue;

      cf := FieldColumns[arrField[i]].Footer;  //获取字段对应的表脚单元格
      cf.ValueType := fvtStaticText;         //以静态文本的方式显示
      cf.Alignment := taRightJustify;        //居右显示
      cf.Value := FormatFloat(DisplayFormat, //格式化
        arrFloat[i]);
    end;
  finally
    grd.DataSource.DataSet.EnableControls;
  end;
  //Result := cf.Value;
end;

function FieldFooterValue(grd: TDBGridEh; FieldName: string; FieldValue: string = '无'; DisplayFormat: string = '#0.00'): string;
//  function IsFloatValid(Value: string): Boolean;
//  var
//    i: Integer;
//    TempValue: Boolean;
//  begin
//    TempValue:=True;
//    for i:=1 to Length(Value) do
//    if not (Value[i] in ['-','.','1','2','3','4','5','6','7','8','9','0']) then
//    begin
//      TempValue:=False;
//      Break;
//    end;
//    Result:=TempValue;
//  end;
var
  cf: TColumnFooterEh;
begin
  with grd do
  begin
    cf := FieldColumns[FieldName].Footer;  //获取字段对应的单元表格
    cf.ValueType := fvtStaticText;         //以静态文本的方式显示
    cf.Alignment := taRightJustify;        //居右显示
    if IsNumeric(FieldValue, True) then
      cf.Value := FormatFloat(DisplayFormat, StrToFloat(FieldValue))
    else
      cf.Value := FieldValue;
  end;
  Result := cf.Value;
end;

function SetGrdFooterValue(grd: TDBGridEh; FieldName: string; FieldValue: string = ''): Boolean;
var
  cf: TColumnFooterEh;
begin
  Result := False;
  with grd do
  try
    if not Assigned(FindFieldColumn(FieldName)) then
      Exit;

    cf := FieldColumns[FieldName].Footer;  //获取字段对应的表脚单元格
    cf.ValueType := fvtStaticText;         //以静态文本的方式显示
    cf.Alignment := taRightJustify;        //居右显示
    cf.Value := FieldValue;
    Result := True;
  except

  end;
end;

//保存控件值，支持ComboBox，Edit，RadioButton,CheckBox 如果需要再按以下格式手动修改代码
procedure SaveIni(FormName: string; Cmb: TWinControl);
var
  IniFile: TIniFile;
  Value: string;
  bValue: Boolean;
begin
  if CreateIniFile(IniFile, GetIniFile) then
  try
    if not (Cmb is TCheckBox) then
    begin
      if (Cmb is TComboBox) then
        Value := (Cmb as TComboBox).Text;
      if Cmb is TEdit then
        Value := Trim((Cmb as TEdit).Text);
      IniFile.WriteString(FormName, Cmb.Name, Value);
      if (Cmb is TRadioButton) then
      begin
        bValue := (Cmb as TRadioButton).Checked;
        IniFile.WriteBool(FormName, Cmb.Name, bValue);
      end;
    end
    else
    begin
      bValue := (Cmb as TCheckBox).Checked;
      IniFile.WriteBool(FormName, Cmb.Name, bValue);
    end;
  finally
    IniFile.Free;
  end;
end;

procedure LoadIni(FormName: string; Cmb: TWinControl);
var
  IniFile: TIniFile;
begin
  if CreateIniFile(IniFile, GetIniFile) then
  try
    if not (Cmb is TCheckBox) then
    begin
      if (Cmb is TComboBox) then
        (Cmb as TComboBox).Text := IniFile.ReadString(FormName, Cmb.Name, '');
      if Cmb is TEdit then
        (Cmb as TEdit).Text := IniFile.ReadString(FormName, Cmb.Name, '');
      if (Cmb is TRadioButton) then
        (Cmb as TRadioButton).Checked := IniFile.ReadBool(FormName, Cmb.Name, False);
    end
    else
      (Cmb as TCheckBox).Checked := IniFile.ReadBool(FormName, Cmb.Name, False);
  finally
    IniFile.Free;
  end;
end;

function GetLocalComputerIP: string;
var
  ch: array[1..32] of char;
  wsData: TWSAData;
  myHost: PHostEnt;
  i: integer;
begin
  Result := '';
  if WSAstartup(2, wsData) <> 0 then
    Exit; // can’t start winsock
  try
    if GetHostName(@ch[1], 32) <> 0 then
      Exit; // getHostName failed
  except
    Exit;
  end;
  myHost := GetHostByName(@ch[1]); // GetHostName error
  if myHost = nil then
    exit;
  for i := 1 to 4 do
  begin
    Result := Result + IntToStr(Ord(myHost.h_addr^[i - 1]));
    if i < 4 then
      Result := Result + '.';
  end;
end;

function GetLocalComputerName: string;
var
  FStr: PChar;
  FSize: Cardinal;
begin
  FSize := 255;
  GetMem(FStr, FSize);
  Windows.GetComputerName(FStr, FSize);
  Result := FStr;
  FreeMem(FStr);
end;

function GetWinLogonName: string;
var
  FStr: PChar;
  FSize: Cardinal;
begin
  FSize := 255;
  GetMem(FStr, FSize);
  GetUserName(FStr, FSize);
  Result := FStr;
  FreeMem(FStr);
end;

//客户端机器的MAC地址
function GetMacAddress: string;
var
  NCB: PNCB;
  Adapter: PAdapterStatus;
  URetCode: AnsiChar;
  RetCode: AnsiChar;
  I: integer;
  Lenum: PlanaEnum;
  _SystemID: AnsiString;
  TMPSTR: string;
begin
  Result := '';
  _SystemID := '';
  Getmem(NCB, SizeOf(TNCB));
  Fillchar(NCB^, SizeOf(TNCB), 0);

  Getmem(Lenum, SizeOf(TLanaEnum));
  Fillchar(Lenum^, SizeOf(TLanaEnum), 0);

  Getmem(Adapter, SizeOf(TAdapterStatus));
  Fillchar(Adapter^, SizeOf(TAdapterStatus), 0);

  Lenum.Length := chr(0);
  NCB.ncb_command := chr(NCBENUM);
  NCB.ncb_buffer := Pointer(Lenum);
  NCB.ncb_length := SizeOf(Lenum);
  RetCode := Netbios(NCB);

  I := 0;
  repeat
    Fillchar(NCB^, SizeOf(TNCB), 0);
    NCB.ncb_command := chr(NCBRESET);
    NCB.ncb_lana_num := Lenum.lana[I];
    RetCode := Netbios(NCB);

    Fillchar(NCB^, SizeOf(TNCB), 0);
    NCB.ncb_command := chr(NCBASTAT);
    NCB.ncb_lana_num := Lenum.lana[I];
    // Must be 16
    NCB.ncb_callname := ('*');

    NCB.ncb_buffer := Pointer(Adapter);

    NCB.ncb_length := SizeOf(TAdapterStatus);
    RetCode := Netbios(NCB);
    //---- calc _systemId from mac-address[2-5] XOR mac-address[1]...
    if (RetCode = chr(0)) or (RetCode = chr(6)) then
    begin
      _SystemID := IntToHex(Ord(Adapter.adapter_address[0]), 2) + '-' + IntToHex(Ord(Adapter.adapter_address[1]), 2) + '-' + IntToHex(Ord(Adapter.adapter_address[2]), 2) + '-' + IntToHex(Ord(Adapter.adapter_address[3]), 2) + '-' + IntToHex(Ord(Adapter.adapter_address[4]), 2) + '-' + IntToHex(Ord(Adapter.adapter_address[5]), 2);
    end;
    Inc(I);
  until (I >= Ord(Lenum.Length)) or (_SystemID <> '00-00-00-00-00-00');
  FreeMem(NCB);
  FreeMem(Adapter);
  FreeMem(Lenum);
  Result := _SystemID
end;

//显示加载信息
procedure ShowLodingMessage(Mes: string; ShowTime: Boolean = False; BackgroundColorB: TColor = clWhite; BackgroundColorE: TColor = clWhite; aFontColor: TColor = $00444444);
begin
  //$00FF8000
  if not Assigned(frmLoding) then
    frmLoding := TfrmLoding.Create(Application);

  with frmLoding do
  begin
    pnlShowMessage.GradientColorStart := BackgroundColorB;
    pnlShowMessage.GradientColorStop := BackgroundColorE;

    timeVisible := ShowTime;
    if ShowTime then
    begin
      lblTime.color := BackgroundColorE;
    end
    else
      lblTime.Visible := False;

    if Mes <> '' then
      pnlMsg.Caption := Mes
    else
      pnlMsg.Caption := '系统正在加载数据，请稍候……';
    pnlMsg.Font.Color := aFontColor;
    //frmLoding.AutoFormSize;
    frmLoding.Show;
    pnlShowMessage.Repaint;
    pnlMsg.Repaint;
  end;
end;

//隐藏加载信息
procedure HideLodingMessage;
begin
  if Assigned(frmLoding) then
    FreeAndNil(frmLoding);
end;

function ReturnToday(CurrTime: TDateTime; bDay: Boolean = True): string;
var
  Year, Month, Day: Word;
const
  Days: array[1..7] of string = ('星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六');
begin
  DecodeDate(CurrTime, Year, Month, Day);
  if bDay then
    Result := IntToStr(Year) + '年' + IntToStr(Month) + '月' + IntToStr(Day) + '日' + ' ' + Days[DayOfWeek(CurrTime)]
  else
    Result := IntToStr(Year) + '年' + IntToStr(Month) + '月' + IntToStr(Day) + '日'
end;

/// <summary>
/// 小写转大写
/// </summary>
/// <param name="mmje"></param>
/// <returns></returns>
function UpperCaseMoney(Decimal: string): string;
var
  s, t: TStringList;
  i, ti: Integer;
  tmps: string;
  IsDecimal: Integer;
begin
  IsDecimal := 0;
  ti := 0;
  for i := 1 to Length(Decimal) do
  begin
    if Decimal[i] = '.' then
    begin
      Inc(IsDecimal);
      ti := i;
    end;
    if not (Decimal[i] in ['0'..'9', '.']) or (IsDecimal > 1) then
    begin
      Result := Decimal;
      Exit;
    end;
  end;
  if ti = 0 then                        //小数点的位置  Begin    tmps := Decimal + '.';
    ti := Length(tmps);
  tmps := Decimal;
  s := TStringList.Create;
  t := TStringList.Create;
  try
    s.Delimiter := ',';
    t.Delimiter := ',';
    s.CommaText := '0=零,1=壹,2=贰,3=叁,4=肆,5=伍,6=陆,7=柒,8=捌,9=玖';
    t.CommaText := '-1=分,0=角,1=元,2=拾,3=佰,4=仟,5=万,6=拾,7=佰,8=仟,9=亿,10=拾,11=佰,12=仟,13=兆,14=拾,15=佰,16=仟';
    ti := Length(tmps) - ti;
    tmps := StringReplace(tmps, '.', '', [rfReplaceAll]);
    for i := 1 to Length(tmps) do
    begin
      if (tmps[Length(tmps) - i + 1] = '0') and (i mod 4 = 1) then
        Result := t.Values[IntToStr(i - ti)] + Result
      else if (tmps[Length(tmps) - i + 1] = '0') and (i mod 4 <> 1) then
        Result := s.Values[tmps[Length(tmps) - i + 1]] + Result
      else
        Result := s.Values[tmps[Length(tmps) - i + 1]] + t.Values[IntToStr(i - ti)] + Result;
    end;
    while Pos('零零', Result) > 0 do
      Result := StringReplace(Result, '零零', '零', [rfReplaceAll]);
    if Pos('零兆', Result) > 0 then
      Result := StringReplace(Result, '零兆', '兆', [rfReplaceAll]);
    if Pos('零亿', Result) > 0 then
      Result := StringReplace(Result, '零亿', '亿', [rfReplaceAll]);
    if Pos('零万', Result) > 0 then
      Result := StringReplace(Result, '零万', '万', [rfReplaceAll]);
    if Pos('零元', Result) > 0 then
      Result := StringReplace(Result, '零元', '元', [rfReplaceAll]);
    if Pos('兆亿', Result) > 0 then
      Result := StringReplace(Result, '兆亿', '兆', [rfReplaceAll]);
    if Pos('亿万', Result) > 0 then
      Result := StringReplace(Result, '亿万', '亿', [rfReplaceAll]);
    if Pos('兆万', Result) > 0 then
      Result := StringReplace(Result, '兆万', '兆', [rfReplaceAll]);
  finally
    t.Free;
    s.Free;
  end;
end;

function ChineseToLetter(const S: widestring): widestring;

  function GetLetter(const w: AnsiString): WideString;
  var
    i: Integer;
  begin
    case WORD(w[1]) shl 8 + WORD(w[2]) of
      $B0A1..$B0C4:
        Result := 'A';
      $B0C5..$B2C0:
        Result := 'B';
      $B2C1..$B4ED:
        Result := 'C';
      $B4EE..$B6E9:
        Result := 'D';
      $B6EA..$B7A1:
        Result := 'E';
      $B7A2..$B8C0:
        Result := 'F';
      $B8C1..$B9FD:
        Result := 'G';
      $B9FE..$BBF6:
        Result := 'H';
      $BBF7..$BFA5:
        Result := 'J';
      $BFA6..$C0AB:
        Result := 'K';
      $C0AC..$C2E7:
        Result := 'L';
      $C2E8..$C4C2:
        Result := 'M';
      $C4C3..$C5B5:
        Result := 'N';
      $C5B6..$C5BD:
        Result := 'O';
      $C5BE..$C6D9:
        Result := 'P';
      $C6DA..$C8BA:
        Result := 'Q';
      $C8BB..$C8F5:
        Result := 'R';
      $C8F6..$CBF9:
        Result := 'S';
      $CBFA..$CDD9:
        Result := 'T';
      $CDDA..$CEF3:
        Result := 'W';
      $CEF4..$D188:
        Result := 'X';
      $D1B9..$D4D0:
        Result := 'Y';
      $D4D1..$D7F9:
        Result := 'Z';
    else
      result := w;
    end;
  end;

var
  i: integer;
begin
  result := '';
  for i := 1 to Length(S) do
    Result := Result + GetLetter(S[i]);
end;

function IsEnCase(strTemp: string): Boolean;
var
  i: integer;
begin
  if strTemp = '' then
  begin
    Result := False;
    Exit;
  end;
  Result := True;
  strTemp := Trim(strTemp);
  for i := 1 to Length(strTemp) do
    if (strTemp[i] in ['A'..'Z']) or (strTemp[i] in ['a'..'z']) then
      Result := True
    else
    begin
      Result := False;
      Exit;
    end;
end;

function AlphaToInt(Value: string): Integer;
//把Excel样式的英文字母序数转换为整数，26进制数且Z = A0

  function IntegerType(Value: Char; Index: Integer): Integer;
  begin
    Value := UpCase(Value);
    if (Value >= 'A') or (Value <= 'Z') then
    begin
      Result := (Ord(Value) - 64);
      if Index > 0 then
        Result := Result * Trunc(IntPower(26, Index));
    end
    else
      Result := 0;
  end;

var
  i, Len: Integer;
  Temp: Integer;
begin
  Result := 0;
  if (Value = '0') or (Length(Value) < 1) then
    Exit;
  Len := Length(Value);
  if (Len > 7) or ((Len = 7) and (Value > 'FXSHRXW')) then
    raise EConvertError.Create('Overflow while converting to Integer type');

  for i := 1 to Len do
  begin
    Temp := IntegerType(Value[i], Len - i);
    if Temp > 0 then
      Result := Result + Temp
    else
      raise EConvertError.Create('Cannot convert to Integer type');
  end;
end;

function IntToAlpha(Value: Integer): string;
//把非负整数转换成Excel样式的英文字母序数

  function Alphabetic(Value: Integer): string;
  var
    CharMod: Integer;
  begin
    CharMod := Value mod 26;
    Value := Value div 26;

    if CharMod = 0 then         //26
    begin
      Result := 'Z';            //A0
      Value := Value - 1;       //末尾位由0变成26则前一位应少乘一个26
    end
    else
      Result := Char(64 + CharMod);

    if Value <= 0 then         //0 ~ 26
      Exit
    else if Value <= 26 then   //26 + 1 ~ 26 * 26，优化处理即[Value = Value mod 26]
      Result := Char(64 + Value) + Result
    else                       // 26 * 26 + 1 ~
      Result := Alphabetic(Value) + Result;
  end;

begin
  if Value <= 0 then
    Result := '0'
  else
    Result := Alphabetic(Value);
end;

//获得用户操作系统信息，最新支持Win8
function GetWindowsVertion: string;

  function GetWindowsVersionString: string;
  var
    oSVersion: TOSVersionInfoA;
  begin
    Result := '';
    oSVersion.dwOSVersionInfoSize := SizeOf(TOSVersionInfoA);
    if GetVersionExA(oSVersion) then
      with oSVersion do
        Result := Trim(Format('%s', [szCSDVersion]));
  end;

var
  AWin32Version: Extended;
  sWin, SNo: string;
begin
  sWin := 'Windows';
  SNo := Format('%d.%d', [Win32MajorVersion, Win32MinorVersion]);
  AWin32Version := StrToFloat(SNo);
  case Win32Platform of
    VER_PLATFORM_WIN32s:
      Result := sWin + '32';
    VER_PLATFORM_WIN32_WINDOWS:
      begin
        if AWin32Version = 4.0 then
          Result := sWin + '95'
        else if AWin32Version = 4.1 then
          Result := sWin + '98'
        else if AWin32Version = 4.9 then
          Result := sWin + 'Me'
        else
          Result := sWin + '9x';
      end;
    VER_PLATFORM_WIN32_NT:
      begin
        if AWin32Version = 3.51 then
          Result := sWin + 'NT 3.51'
        else if AWin32Version = 4.0 then
          Result := sWin + 'NT 4.0'
        else if AWin32Version = 5.0 then
          Result := sWin + '2000'
        else if AWin32Version = 5.1 then
          Result := sWin + 'XP'
        else if AWin32Version = 5.2 then
          Result := sWin + '2003'
        else if AWin32Version = 6.0 then
          Result := sWin + 'Vista'
        else if AWin32Version = 6.1 then
          Result := sWin + '7'
        else if AWin32Version = 6.2 then
          Result := sWin + '8'
        else if AWin32Version = 6.2 then
          Result := sWin + 'RT'
        else if AWin32Version = 6.3 then
          Result := sWin + '8.1'
        else
          Result := sWin;
      end
  else
    Result := sWin;
  end;
  Result := Result + '  ' + GetWindowsVersionString + '(' + SNo + ')';
end;

//获得执行程序版本号，如果没版本号则返回'Not Found Version Information'
function GetFileVersion(FileName: PChar): string;
var
  VerInfo: Pointer;
  VerInfoSize: DWORD;
  VerValueSize: DWORD;
  VerValue: PVSFixedFileInfo;
  Dummy: DWORD;
  V1, V2, V3, V4: Word;
begin
  VerInfoSize := GetFileVersionInfoSize(FileName, Dummy);
  if VerInfoSize = 0 then
  begin
    Result := NoVersionInfoStr;
    Exit;
  end;
  GetMem(VerInfo, VerInfoSize);
  try
    GetFileVersionInfo(FileName, 0, VerInfoSize, VerInfo);
    VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
    with VerValue^ do
    begin
      V1 := dwFileVersionMS shr 16;
      V2 := dwFileVersionMS and $FFFF;
      V3 := dwFileVersionLS shr 16;
      V4 := dwFileVersionLS and $FFFF;
    end;
    Result := Format('%d.%d.%d.%d', [V1, V2, V3, V4]);
  finally
    FreeMem(VerInfo, VerInfoSize);
  end;
end;

//获得执行程序创建时间，修改时候，最后访问时间；
function GetFilesTime(sFilename: string; Timetype: Integer): TDateTime;
var
  ffd: TWin32FindData;
  dft: DWord;
  lft, Time: TFileTime;
  sHandle: THandle;
begin
  sHandle := Windows.FindFirstFile(PChar(sFilename), ffd);
  if (sHandle <> INVALID_HANDLE_VALUE) then
  begin
    case Timetype of
      0:
        Time := ffd.ftCreationTime;    //创建时间
      1:
        Time := ffd.ftLastWriteTime;   //修改时间
      2:
        Time := ffd.ftLastAccessTime;  //访问时间
    end;
    Windows.FindClose(sHandle);
    FileTimeToLocalFileTime(Time, lft);
    FileTimeToDosDateTime(lft, LongRec(dft).HI, LongRec(dft).Lo);
    Result := FileDateToDateTime(dft);
  end
  else
    Result := 0;
end;

function GetDisksInfo(strL: TStringList): TStringList;
var
  str: string;
  Drivers: Integer;
  driver: char;
  i, temp, disksizee: integer;
  d1, d2, d3, d4: DWORD;
begin
  strL := TStringList.Create;
  Drivers := GetLogicalDrives;
  temp := (1 and Drivers);
  for i := 0 to 26 do
  begin
    if temp = 1 then
    begin
      driver := char(i + integer('A'));
      str := driver + ':';
      if (driver <> '') and (getdrivetype(pchar(str)) <> drive_cdrom) and (getdrivetype(pchar(str)) <> DRIVE_REMOVABLE) then
      begin
        GetDiskFreeSpace(pchar(str), d1, d2, d3, d4);
        strL.Add(str + Format('总空间:%f GB', [d4 / 1024 / 1024 / 1024 * d2 * d1]) + Format(' 剩余空间:%f GB', [d3 / 1024 / 1024 / 1024 * d2 * d1]));
      end;
    end;
    Drivers := (Drivers shr 1);
    temp := (1 and Drivers);
  end;
  Result := strL;
  SetErrorMode(oldMode);
end;

//字符串加密
function EncryptionStr(Src: string; Key: string = 'WhComTec'): string;
var
  idx: integer;
  KeyLen: Integer;
  KeyPos: Integer;
  offset: Integer;
  dest: string;
  SrcPos: Integer;
  SrcAsc: Integer;
  TmpSrcAsc: Integer;
  Range: Integer;
begin
  KeyLen := Length(Key);
  if KeyLen = 0 then
    Key := 'Think Space';
  KeyPos := 0;
  SrcPos := 0;
  SrcAsc := 0;
  Range := 256;

  Randomize;
  offset := Random(Range);
  dest := format('%1.2x', [offset]);
  for SrcPos := 1 to Length(Src) do
  begin
    SrcAsc := (Ord(Src[SrcPos]) + offset) mod 255;
    if KeyPos < KeyLen then
      KeyPos := KeyPos + 1
    else
      KeyPos := 1;
    SrcAsc := SrcAsc xor Ord(Key[KeyPos]);
    dest := dest + format('%1.2x', [SrcAsc]);
    offset := SrcAsc;
  end;
  Result := dest;
end;

//解密函数
function DecryptStr(Src: string; Key: string = 'WhComTec'): string;
var
  KeyLen: Integer;
  KeyPos: Integer;
  offset: Integer;
  dest: string;
  SrcPos: Integer;
  SrcAsc: Integer;
  TmpSrcAsc: Integer;
  Range: Integer;
begin
  KeyLen := Length(Key);
  if KeyLen = 0 then
    Key := 'Think Space';
  KeyPos := 0;
  SrcPos := 0;
  SrcAsc := 0;
  Range := 256;
  offset := StrToInt('$' + copy(Src, 1, 2));
  SrcPos := 3;
  repeat
    SrcAsc := StrToInt('$' + copy(Src, SrcPos, 2));
    if KeyPos < KeyLen then
      KeyPos := KeyPos + 1
    else
      KeyPos := 1;
    TmpSrcAsc := SrcAsc xor Ord(Key[KeyPos]);
    if TmpSrcAsc <= offset then
      TmpSrcAsc := 255 + TmpSrcAsc - offset
    else
      TmpSrcAsc := TmpSrcAsc - offset;
    dest := dest + chr(TmpSrcAsc);
    offset := SrcAsc;
    SrcPos := SrcPos + 2;
  until SrcPos >= Length(Src);
  Result := dest;
end;

function zipAFile(sOldFile, sNewFile: string): Boolean;
var
  cs: TCompressionStream;
  fs, ms: TMemoryStream;
  num: Integer;
begin
  Result := False;
  if not FileExists(sOldFile) then
    Exit;

  try
    fs := TMemoryStream.Create;
    fs.LoadFromFile(sOldFile);
    num := fs.Size;

    ms := TMemoryStream.Create;
    ms.Write(num, SizeOf(num));

    cs := TCompressionStream.Create(clMax, ms);
    fs.SaveToStream(cs);
    cs.Free;

    ForceDirectories(ExtractFilePath(sNewFile));
    ms.SaveToFile(sNewFile);

    ms.Free;
    fs.Free;

    Result := True;
  except
    on e: Exception do
    begin
      ShowMessageBoxEx('压缩文件出错，消息：' + e.Message, 'info');
    end;
  end;
end;

function unZipAFile(sOldFile, sNewFile: string): Boolean;
var
  ds: TDecompressionStream;
  fs, ms: TMemoryStream;
  num: Integer;
begin
  Result := False;
  if not FileExists(sOldFile) then
    Exit;

  try
    fs := TMemoryStream.Create;
    fs.LoadFromFile(sOldFile);

    fs.Position := 0;
    fs.ReadBuffer(num, SizeOf(num));

    ms := TMemoryStream.Create;
    ms.SetSize(num);

    ds := TDecompressionStream.Create(fs);

    ds.Read(ms.Memory^, num);

    ForceDirectories(ExtractFilePath(sNewFile));
    ms.SaveToFile(sNewFile);

    ds.Free;
    ms.Free;
    fs.Free;
    Result := True;
  except
    on e: Exception do
    begin
      ShowMessageBoxEx('解压文件出错，消息：' + e.Message, 'info');
    end;

  end;

end;

function CompressFiles(const ZipFileName: string; const Files: TArray<string>): Boolean;
var
  ZipFile: TZipFile;
  I: Integer;
begin
  result := False;
  if Length(Files) = 0 then
  begin
    ShowMessageBoxEx('传入目标文件不能为空！',MB_ICONINFORMATION);
    Exit;
  end;

  ZipFile := TZipFile.Create;
  try
    try
      // 创建或覆盖 ZIP 文件
      ZipFile.Open(ZipFileName, zmWrite);

      // 添加所有文件到压缩包
      for I := 0 to High(Files) do
      begin
        // 在 ZIP 中只保留文件名（不含路径）
        if FileExists(files[i]) then
          ZipFile.Add(Files[I], ExtractFileName(Files[I]));
      end;

      // 保存并关闭 ZIP 文件
      ZipFile.Close;

      Result := True;
    except
      on e:Exception do
      begin
        ShowMessageBoxEx('生成压缩文件失败，消息：' + e.message,MB_ICONINFORMATION);
      end;

    end;
  finally
    ZipFile.Free;
  end;
end;

function DeleteFolder(const FolderPath: string): Boolean;
var
  FileOp: TSHFileOpStruct;
begin
  FillChar(FileOp, SizeOf(FileOp), 0);
  FileOp.Wnd := 0;
  FileOp.wFunc := FO_DELETE;
  FileOp.pFrom := PChar(FolderPath + #0);
  FileOp.pTo := nil;
  FileOp.fFlags := FOF_SILENT or FOF_NOCONFIRMATION or FOF_NOERRORUI or FOF_WANTNUKEWARNING;

  Result := (SHFileOperation(FileOp) = 0);
end;

function DeleteDirectory(NowPath: string): Boolean; // 删除整个目录
var
  search: TSearchRec;
  ret: integer;
  key: string;
begin
  if NowPath[Length(NowPath)] <> '\' then
    NowPath := NowPath + '\';
  key := NowPath + '*.*';
  ret := findFirst(key, faanyfile, search);
  while ret = 0 do
  begin
    if ((search.Attr and fadirectory) = fadirectory) then
    begin
      if (search.Name <> '.') and (search.name <> '..') then
        DeleteDirectory(NowPath + search.name);
    end
    else
    begin
      if ((search.Attr and fadirectory) <> fadirectory) then
      begin
        deletefile(NowPath + search.name);
      end;
    end;
    ret := FindNext(search);
  end;
  findClose(search);
  //removedir(NowPath); 如果需要删除文件夹则添加
  result := True;
end;

function CopyAndMoveDirectory(SourceDirName: string; TargetDirName: string): Boolean;
var
  hFindFile: Cardinal;
  t, tfile: string;
  sCurDir: string[255];
  FindFileData: WIN32_FIND_DATA;
begin
  //记录当前目录
  sCurDir := GetCurrentDir;
  ChDir(SourceDirName);
  hFindFile := FindFirstFile('*.*', FindFileData);
  if hFindFile <> INVALID_HANDLE_VALUE then
  begin
    if not DirectoryExists(TargetDirName) then
      ForceDirectories(TargetDirName);
    repeat
      tfile := FindFileData.cFileName;
      if (tfile = '.') or (tfile = '..') then
        Continue;
      if FindFileData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY then
      begin
        t := TargetDirName + '\' + tfile;
        if not DirectoryExists(t) then
          ForceDirectories(t);
        if SourceDirName[Length(SourceDirName)] <> '\' then
          CopyAndMoveDirectory(SourceDirName + '\' + tfile, t)
        else
          CopyAndMoveDirectory(SourceDirName + tfile, TargetDirName + tfile);
      end
      else
      begin
        t := TargetDirName + '\' + tfile;
        CopyFile(PChar(tfile), PChar(t), True);
      end;
    until FindNextFile(hFindFile, FindFileData) = false;
    ///  FindClose(hFindFile);
  end
  else
  begin
    ChDir(sCurDir);
    result := false;
    exit;
  end;
  //回到当前目录
  ChDir(sCurDir);
  Result := true;
end;

{$IFDEF Delphi2010_Up}
function IsWin64: Bool;
begin
  Result := False;
  if GetProcAddress(GetModuleHandle(kernel32), 'IsWow64Process') <> nil then
    IsWow64Process(GetCurrentProcess, Result);
end;
{$ENDIF}

function GetSubDirectoryList(pPath: string): TStringList;
var
  SearchRec: TSearchRec;
  i: Integer;
  lStringList: TStringList;
  TempPath: string;
begin
  lStringList := TStringList.Create;
  TempPath := pPath + '\*.*';
  i := FindFirst(TempPath, faDirectory, SearchRec);
  while i = 0 do
  begin
    if SearchRec.Attr = 16 then
    begin
      if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') then
        lStringList.Add(SearchRec.Name);
    end;
    i := FindNext(SearchRec);
  end;
  FindClose(SearchRec);
  Result := lStringList;
end;

function GetDirectoryFileList(pPath: string): TStringList;
var
  SearchRec: TSearchRec;
  found: integer;
  lStringList: TStringList;
  TempPath: string;
begin
  lStringList := TStringList.Create;
  TempPath := pPath + '\*.*';
  found := FindFirst(TempPath, faAnyFile, SearchRec);
  while found = 0 do
  begin
    if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') and (SearchRec.Attr <> faDirectory) then
      lStringList.Add(SearchRec.Name);
    found := FindNext(SearchRec);
  end;
  FindClose(SearchRec);
  Result := lStringList;
end;

// procedure TForm1.Button1Click(Sender: TObject);
// begin
// Memo1.Lines.Clear;
// GetFileListEx('Z:\', '*.*', Memo1.Lines, False);
// Caption:=IntToStr(Memo1.Lines.count);
// end;
//
// procedure TForm1.Button2Click(Sender: TObject);
// begin
//
// Memo1.Lines.Clear;
// GetFileListEx('Z:\', '*.cs',  Memo1.Lines, true);
/// /GetFileListEx('Z:\', '*.cs;*.txt',  Memo1.Lines, true);
// Caption:=IntToStr(Memo1.Lines.count);
// end;

// 遍历目录及子目录
procedure GetFileListEx(FilePath, ExtMask: string; FileList: TStrings; SubDirectory: Boolean = True; addPath: Boolean = True);

  function Match(FileName: string; MaskList: TStrings): Boolean;
  var
    i: integer;
  begin
    Result := False;
    for i := 0 to MaskList.Count - 1 do
    begin
      if MatchesMask(FileName, MaskList[i]) then
      begin
        Result := True;
        break;
      end;
    end;
  end;

var
  FileRec: TSearchRec;
  MaskList: TStringList;
begin
  if DirectoryExists(FilePath) then
  begin
    if FilePath[Length(FilePath)] <> '\' then
      FilePath := FilePath + '\';
    if FindFirst(FilePath + '*.*', faAnyFile, FileRec) = 0 then
    begin
      MaskList := TStringList.Create;
      try
        ExtractStrings([';'], [], PChar(ExtMask), MaskList);
        FileList.BeginUpdate;
        repeat
          if ((FileRec.Attr and faDirectory) <> 0) and SubDirectory then
          begin
            if (FileRec.Name <> '.') and (FileRec.Name <> '..') then
              GetFileListEx(FilePath + FileRec.Name + '\', ExtMask, FileList);
          end
          else
          begin
            if Match(FilePath + FileRec.Name, MaskList) then
            begin
              if addPath then
                FileList.Add(FilePath + FileRec.Name)
              else
                FileList.Add( { FilePath + } FileRec.Name);
            end;
          end;
        until FindNext(FileRec) <> 0;
        FileList.EndUpdate;
      finally
        MaskList.Free;
      end;
    end;
    FindClose(FileRec);
  end;
end;

function MakeFileList(Path: string; FileExt: string = ''): TStringList;
var
  sch: TSearchrec;
begin
  Result := TStringlist.Create;

  if rightStr(trim(Path), 1) <> '\' then
    Path := trim(Path) + '\'
  else
    Path := trim(Path);

  if Trim(FileExt) = '' then
    FileExt := '.*';

  if not DirectoryExists(Path) then
  begin
    Result.Clear;
    exit;
  end;

  if FindFirst(Path + '*', faAnyfile, sch) = 0 then
  begin
    repeat
      Application.ProcessMessages;
      if ((sch.Name = '.') or (sch.Name = '..')) then
        Continue;
      if DirectoryExists(Path + sch.Name) then
      begin
        Result.AddStrings(MakeFileList(Path + sch.Name, FileExt));
      end
      else
      begin
        if (UpperCase(extractfileext(Path + sch.Name)) = UpperCase(FileExt)) or (FileExt = '.*') then
          Result.Add(Path + sch.Name);
      end;
    until FindNext(sch) <> 0;
    FindClose(sch);
  end;
end;

function ExtractFileNameNoExt(FileString: string): string;
var
  FileWithExtString: string;
  FileExtString: string;
  LenExt: Integer;
  LenNameWithExt: Integer;
begin
  FileWithExtString := ExtractFileName(FileString);
  LenNameWithExt := Length(FileWithExtString);
  FileExtString := ExtractFileExt(FileString);
  LenExt := Length(FileExtString);
  if LenExt = 0 then
  begin
    Result := FileWithExtString;
  end
  else
  begin
    Result := Copy(FileWithExtString, 1, (LenNameWithExt - LenExt));
  end;
end;

//Handle：窗体句柄；Caption：弹框提示；SelectedFolder： 返回选择路径；
//InitFolder：默认打开路径，空则为桌面
function SelectFolderDialog(const Handle: integer; const Caption: string; var SelectedFolder: string; const InitFolder: WideString = ' '): boolean;
var
  BInfo: _browseinfo;
  Buffer: array[0..MAX_PATH] of Char;
  ID: IShellFolder;
  Eaten, Attribute: Cardinal;
  ItemID: PItemidlist;
begin
  Result := False;
  BInfo.HwndOwner := Handle;
  BInfo.lpfn := nil;
  BInfo.lpszTitle := Pchar(Caption);
  BInfo.ulFlags := BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE;
  SHGetDesktopFolder(ID);
  ID.ParseDisplayName(0, nil, PWideChar(InitFolder), Eaten, ItemID, Attribute);
  BInfo.pidlRoot := ItemID;
  GetMem(BInfo.pszDisplayName, MAX_PATH);
  try
    if SHGetPathFromIDList(SHBrowseForFolder(BInfo), Buffer) then
    begin
      SelectedFolder := Buffer;
      if Length(SelectedFolder) <> 3 then
        SelectedFolder := SelectedFolder + '\';
      result := True;
    end
    else
    begin
      SelectedFolder := '';
      Result := False;
    end;
  finally
    FreeMem(BInfo.pszDisplayName);
  end;
end;

function OpenWithShellExecute(const targetPath: string; openType: TOpenFileType = vFTFileName): Boolean;
var
  ResultCode: NativeInt;
begin
  case openType of
    vFTFileName, vFTURL: ResultCode := ShellExecute(0, 'open', PChar(targetPath), nil, nil, SW_SHOWNORMAL);
    vFTDirectory: ResultCode := ShellExecute(0, 'explorer', PChar('/select,'+targetPath), nil, nil, SW_SHOWNORMAL);
    vFTPrint: ResultCode := ShellExecute(0, 'print', PChar(targetPath), nil, nil, SW_HIDE);
  end;

  Result := (LongInt(ResultCode) > 32); // 检查返回值是否表明操作成功
end;

procedure SetPrivilege;
var
  OldTokenPrivileges, TokenPrivileges: TTokenPrivileges;
  ReturnLength: dword;
  hToken: THandle;
  Luid: int64;
begin
  OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVILEGES, hToken);
  LookupPrivilegeValue(nil, 'SeDebugPrivilege', Luid);
  TokenPrivileges.Privileges[0].Luid := Luid;
  TokenPrivileges.PrivilegeCount := 1;
  TokenPrivileges.Privileges[0].Attributes := 0;
  AdjustTokenPrivileges(hToken, False, TokenPrivileges, SizeOf(TTokenPrivileges), OldTokenPrivileges, ReturnLength);
  OldTokenPrivileges.Privileges[0].Luid := Luid;
  OldTokenPrivileges.PrivilegeCount := 1;
  OldTokenPrivileges.Privileges[0].Attributes := TokenPrivileges.Privileges[0].Attributes or SE_PRIVILEGE_ENABLED;
  AdjustTokenPrivileges(hToken, False, OldTokenPrivileges, ReturnLength, PTokenPrivileges(nil)^, ReturnLength);
end;

function PingRemoteMachine(aRemoteIP: string): Boolean;
var
  FtPing: TFtPing;
  aReply: string;
begin
  Result := False;
  FtPing := TFtPing.Create(nil);
  try
    FtPing.RemoteIP := aRemoteIP;
    if FtPing.Ping(aReply) then
      Result := True;
  finally
    FtPing.Free;
  end;
end;

procedure WaitTime(MSecs: Integer);
var
  FirstTickCount, Now: Longint;
begin
  FirstTickCount := GetTickCount();
  repeat
    Application.ProcessMessages;
    Sleep(1);  // 避免CPU占用率高

    Now := GetTickCount();
  until (Now - FirstTickCount >= MSecs) or (Now < FirstTickCount);
end;

procedure WaitTimeEx(MaxWaitTime: Cardinal);
var
  i: Integer;
  WaitedTime: Cardinal;
begin
  WaitedTime := 0;
  while (WaitedTime < MaxWaitTime) do
  begin
    Application.ProcessMessages;
    SleepEx(100, False);
    Inc(WaitedTime, 100);

  end
end;

end.

