unit uClosingEffect;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics;

const
  WM_REDIRCREATE = WM_USER + 382;
  WM_REDIRDESTROY = WM_USER + 383;
  WM_REDIRCLOSING = WM_USER + 384;
  REDIR_CHECK_INTERVAL = 50;
  REDIR_CHECK_STEP = 4;

function BeginGrayscaleClosingEffect(Wnd: HWND; FadingTime: Cardinal = 1000): Cardinal;

procedure EndGrayscaleClosingEffect(ID: Cardinal; IsRestoring: Boolean = False);


var
  AlphaColorTable : array [0..255, 0..255] of Byte;
  GrayColorTable : array [0..2, 0..255] of Integer;
  CRCTable : array [0..255] of DWORD =
  ( $00000000, $77073096, $EE0E612C, $990951BA,
    $076DC419, $706AF48F, $E963A535, $9E6495A3,
    $0EDB8832, $79DCB8A4, $E0D5E91E, $97D2D988,
    $09B64C2B, $7EB17CBD, $E7B82D07, $90BF1D91,
    $1DB71064, $6AB020F2, $F3B97148, $84BE41DE,
    $1ADAD47D, $6DDDE4EB, $F4D4B551, $83D385C7,
    $136C9856, $646BA8C0, $FD62F97A, $8A65C9EC,
    $14015C4F, $63066CD9, $FA0F3D63, $8D080DF5,
    $3B6E20C8, $4C69105E, $D56041E4, $A2677172,
    $3C03E4D1, $4B04D447, $D20D85FD, $A50AB56B,
    $35B5A8FA, $42B2986C, $DBBBC9D6, $ACBCF940,
    $32D86CE3, $45DF5C75, $DCD60DCF, $ABD13D59,
    $26D930AC, $51DE003A, $C8D75180, $BFD06116,
    $21B4F4B5, $56B3C423, $CFBA9599, $B8BDA50F,
    $2802B89E, $5F058808, $C60CD9B2, $B10BE924,
    $2F6F7C87, $58684C11, $C1611DAB, $B6662D3D,
    $76DC4190, $01DB7106, $98D220BC, $EFD5102A,
    $71B18589, $06B6B51F, $9FBFE4A5, $E8B8D433,
    $7807C9A2, $0F00F934, $9609A88E, $E10E9818,
    $7F6A0DBB, $086D3D2D, $91646C97, $E6635C01,
    $6B6B51F4, $1C6C6162, $856530D8, $F262004E,
    $6C0695ED, $1B01A57B, $8208F4C1, $F50FC457,
    $65B0D9C6, $12B7E950, $8BBEB8EA, $FCB9887C,
    $62DD1DDF, $15DA2D49, $8CD37CF3, $FBD44C65,
    $4DB26158, $3AB551CE, $A3BC0074, $D4BB30E2,
    $4ADFA541, $3DD895D7, $A4D1C46D, $D3D6F4FB,
    $4369E96A, $346ED9FC, $AD678846, $DA60B8D0,
    $44042D73, $33031DE5, $AA0A4C5F, $DD0D7CC9,
    $5005713C, $270241AA, $BE0B1010, $C90C2086,
    $5768B525, $206F85B3, $B966D409, $CE61E49F,
    $5EDEF90E, $29D9C998, $B0D09822, $C7D7A8B4,
    $59B33D17, $2EB40D81, $B7BD5C3B, $C0BA6CAD,
    $EDB88320, $9ABFB3B6, $03B6E20C, $74B1D29A,
    $EAD54739, $9DD277AF, $04DB2615, $73DC1683,
    $E3630B12, $94643B84, $0D6D6A3E, $7A6A5AA8,
    $E40ECF0B, $9309FF9D, $0A00AE27, $7D079EB1,
    $F00F9344, $8708A3D2, $1E01F268, $6906C2FE,
    $F762575D, $806567CB, $196C3671, $6E6B06E7,
    $FED41B76, $89D32BE0, $10DA7A5A, $67DD4ACC,
    $F9B9DF6F, $8EBEEFF9, $17B7BE43, $60B08ED5,
    $D6D6A3E8, $A1D1937E, $38D8C2C4, $4FDFF252,
    $D1BB67F1, $A6BC5767, $3FB506DD, $48B2364B,
    $D80D2BDA, $AF0A1B4C, $36034AF6, $41047A60,
    $DF60EFC3, $A867DF55, $316E8EEF, $4669BE79,
    $CB61B38C, $BC66831A, $256FD2A0, $5268E236,
    $CC0C7795, $BB0B4703, $220216B9, $5505262F,
    $C5BA3BBE, $B2BD0B28, $2BB45A92, $5CB36A04,
    $C2D7FFA7, $B5D0CF31, $2CD99E8B, $5BDEAE1D,
    $9B64C2B0, $EC63F226, $756AA39C, $026D930A,
    $9C0906A9, $EB0E363F, $72076785, $05005713,
    $95BF4A82, $E2B87A14, $7BB12BAE, $0CB61B38,
    $92D28E9B, $E5D5BE0D, $7CDCEFB7, $0BDBDF21,
    $86D3D2D4, $F1D4E242, $68DDB3F8, $1FDA836E,
    $81BE16CD, $F6B9265B, $6FB077E1, $18B74777,
    $88085AE6, $FF0F6A70, $66063BCA, $11010B5C,
    $8F659EFF, $F862AE69, $616BFFD3, $166CCF45,
    $A00AE278, $D70DD2EE, $4E048354, $3903B3C2,
    $A7672661, $D06016F7, $4969474D, $3E6E77DB,
    $AED16A4A, $D9D65ADC, $40DF0B66, $37D83BF0,
    $A9BCAE53, $DEBB9EC5, $47B2CF7F, $30B5FFE9,
    $BDBDF21C, $CABAC28A, $53B39330, $24B4A3A6,
    $BAD03605, $CDD70693, $54DE5729, $23D967BF,
    $B3667A2E, $C4614AB8, $5D681B02, $2A6F2B94,
    $B40BBE37, $C30C8EA1, $5A05DF1B, $2D02EF8D );


implementation
// download by http://www.codefans.net
type
  TClosingEffectWnd = class
  private
    { surface wnd }
    FRedirClass: Cardinal;
    FRedirWnd: HWND;
    FDC: HDC;
    FCache: HBITMAP;
    FCacheW, FCacheH: Integer;

    { target window info }
    FWnd: HWND;       // target window
    FOldProc: Pointer;
    FAtt: Boolean;   // target window has WS_EX_LAYERED style
    FAttC: Cardinal; // target window's layered attribute color
    FAttF: Cardinal; // target window's layered attribute flag
    FAttA: Byte;     // target window's layered attribute alpha
    FAeroGlass: LongBool;  // vista/win7 with aero theme enabled

    { states }
    FFading: Boolean;
    FAlphaEnd: Integer;
    FAlphaStp: Integer;
    FCurrAlpha: Integer;
    FClosing: Boolean;
    FChg: Boolean;        // surface changed
    FRedraw: Boolean;     // redraw surface
    FDigests: TList;      // scan line digests
    FBlend: TBlendFunction;

    procedure InitWndInfos;
    procedure InitRedirWnd;
    procedure FinalRedirWnd;
    procedure CreateRedirWnd;
    procedure DestroyRedirWnd;
    procedure ClosingRedirWnd;
    procedure RestoringRedirWnd;

    procedure HookedWndProc(var Message: TMessage);
    procedure RedirWndProc(var Message: TMessage);

    procedure Render;
    function CheckDigests(Bits: Pointer; w, h: Integer): Boolean;
    procedure AdjustFading;
    procedure DoGrayScale(src, dst: Pointer; w, h: Integer);
  public
    constructor Create(Wnd: HWND; Time: Cardinal);
    destructor Destroy; override;

    procedure Closing;
    procedure Restoring;
  end;

{ global function implementation }
function BeginGrayscaleClosingEffect(Wnd: HWND; FadingTime: Cardinal = 1000): Cardinal;
begin
  if not IsWindowVisible(Wnd) or IsIconic(Wnd) then Result := 0
  else
    Result := Cardinal(TClosingEffectWnd.Create(Wnd, FadingTime));
end;

procedure EndGrayscaleClosingEffect(ID: Cardinal; IsRestoring: Boolean = False);
begin
  if ID <> 0 then
  begin
    with TClosingEffectWnd(ID) do
    begin
      if IsRestoring then
        Restoring
      else Closing;
      Free;
    end;
  end;
end;


{ CRC32 }
function CRC32(Buf: Pointer; Cnt: Integer; Step: Integer; Seed: Cardinal): Cardinal;
var
  i: Integer;
  p: PByte;
begin
  p := PByte(Buf);
  Result := not Seed;
  for i := 1 to Cnt do
  begin
    Result := ((Result shr 8) and $00FFFFFF) xor CRCTable[p^ xor Byte(Result)];
    Inc(p, Step);
  end;
  Result := not Result;
end;

{ Used APIs }
function GetLayeredWindowAttributes(hwnd: HWND; var pcrKey: Cardinal; var phAlpha: Byte; var pdwFlags: Cardinal): BOOL stdcall; external 'user32.dll' name 'GetLayeredWindowAttributes';

{ DWM APIs }
type
  PDWM_BLURBEHIND = ^DWM_BLURBEHIND;
  DWM_BLURBEHIND = packed record
    dwFlags: DWORD;
    fEnable: BOOL;
    hRgnBlur: HRGN;
    fTransitionOnMaximized: BOOL;
  end;
  _DWM_BLURBEHIND = DWM_BLURBEHIND;
  TDWMBlurBehind = DWM_BLURBEHIND;
  PDWMBlurBehind = ^TDWMBlurBehind;

var
  hDWMAPI: THandle;
  _DwmIsCompositionEnabled : function (out pfEnabled: BOOL): HResult; stdcall;
  _DwmGetWindowAttribute: function (hwnd: HWND; dwAttribute: DWORD;
                              pvAttribute: Pointer; cbAttribute: DWORD): HResult;
  _DwmSetWindowAttribute: function (hwnd: HWND; dwAttribute: DWORD;
                              pvAttribute: Pointer; cbAttribute: DWORD): HResult; stdcall;
  _DwmEnableBlurBehindWindow: function (hWnd: HWND; const pBlurBehind: TDWMBlurBehind): HResult; stdcall;


function DwmIsCompositionEnabled(out pfEnabled: BOOL): HResult;
begin
  if Assigned(_DwmIsCompositionEnabled) then
    Result := _DwmIsCompositionEnabled(pfEnabled)
  else
  begin
    if hDWMAPI = 0 then
      hDWMAPI := LoadLibrary('DWMAPI.DLL');

    Result := E_NOTIMPL;
    if hDWMAPI > 0 then
    begin
      _DwmIsCompositionEnabled := GetProcAddress(hDWMAPI, 'DwmIsCompositionEnabled'); // Do not localize
      if Assigned(_DwmIsCompositionEnabled) then
        Result := _DwmIsCompositionEnabled(pfEnabled);
    end;
  end;
end;

function DwmGetWindowAttribute(hwnd: HWND; dwAttribute: DWORD;
  pvAttribute: Pointer; cbAttribute: DWORD): HResult;
begin
  if Assigned(_DwmGetWindowAttribute) then
    Result := _DwmGetWindowAttribute(hwnd, dwAttribute, pvAttribute, cbAttribute)
  else
  begin
    if hDWMAPI = 0 then
      hDWMAPI := LoadLibrary('DWMAPI.DLL');

    Result := E_NOTIMPL;
    if hDWMAPI > 0 then
    begin
      _DwmGetWindowAttribute := GetProcAddress(hDWMAPI, 'DwmGetWindowAttribute'); // Do not localize
      if Assigned(_DwmGetWindowAttribute) then
        Result := _DwmGetWindowAttribute(hwnd, dwAttribute, pvAttribute,
          cbAttribute);
    end;
  end;
end;

function DwmSetWindowAttribute(hwnd: HWND; dwAttribute: DWORD;
  pvAttribute: Pointer; cbAttribute: DWORD): HResult;
begin
  if Assigned(_DwmSetWindowAttribute) then
    Result := _DwmSetWindowAttribute(hwnd, dwAttribute, pvAttribute, cbAttribute)
  else
  begin
    if hDWMAPI = 0 then
      hDWMAPI := LoadLibrary('DWMAPI.DLL');

    Result := E_NOTIMPL;
    if hDWMAPI > 0 then
    begin
      _DwmSetWindowAttribute := GetProcAddress(hDWMAPI, 'DwmSetWindowAttribute'); // Do not localize
      if Assigned(_DwmSetWindowAttribute) then
        Result := _DwmSetWindowAttribute(hwnd, dwAttribute, pvAttribute,
          cbAttribute);
    end;
  end;
end;

function DwmEnableBlurBehindWindow(hWnd: HWND; const pBlurBehind: TDWMBlurBehind): HResult;
begin
  if Assigned(_DwmEnableBlurBehindWindow) then
    Result := _DwmEnableBlurBehindWindow(hWnd, pBlurBehind)
  else
  begin
    if hDWMAPI = 0 then
      hDWMAPI := LoadLibrary('DWMAPI.DLL');

    Result := E_NOTIMPL;
    if hDWMAPI > 0 then
    begin
      _DwmEnableBlurBehindWindow := GetProcAddress(hDWMAPI, 'DwmEnableBlurBehindWindow'); // Do not localize
      if Assigned(_DwmEnableBlurBehindWindow) then
        Result := _DwmEnableBlurBehindWindow(hWnd, pBlurBehind);
    end;
  end;
end;


{ TClosingEffectWnd }

const
  ShowingFlags: array [Boolean] of Cardinal = (
    SWP_NOSIZE or SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE or SWP_HIDEWINDOW,
    SWP_NOSIZE or SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE or SWP_SHOWWINDOW
  );

constructor TClosingEffectWnd.Create(Wnd: HWND; Time: Cardinal);
var
  n: Integer;
label
  NoFade;
begin
  FDigests := TList.Create;
  FWnd := Wnd;
  InitWndInfos;
  FBlend.BlendOp := AC_SRC_OVER;
  FBlend.BlendFlags := 0;
  FBlend.AlphaFormat := AC_SRC_ALPHA;
  FCurrAlpha := FAttA;
  FAlphaEnd := 1;
  n := Time div REDIR_CHECK_INTERVAL;
  if n = 0 then
  begin
    FAlphaStp := 0;
  end
  else begin
    FAlphaStp := (FAlphaEnd - FCurrAlpha) div n;
    if FAlphaStp = 0 then FAlphaStp := -1;
  end;
  FBlend.SourceConstantAlpha := FAttA;
  FFading := True;
  InitRedirWnd;
end;

destructor TClosingEffectWnd.Destroy;
begin
  FinalRedirWnd;
  if FCache <> 0 then
    DeleteObject(FCache);
  if FDC <> 0 then
    DeleteDC(FDC);
  FDigests.Free;

  inherited;
end;

procedure TClosingEffectWnd.InitWndInfos;
begin
  FAtt := GetWindowLong(FWnd, GWL_EXSTYLE) and WS_EX_LAYERED <> 0;
  if FAtt then
    GetLayeredWindowAttributes(FWnd, FAttC, FAttA, FAttF)
  else
    FAttF := 0;
  if FAttF and LWA_ALPHA = 0 then
    FAttA := 255;
    
  { checking Vista/win7 Aero }
  FAeroGlass := Win32MajorVersion >= 6;
  if FAeroGlass then
  begin
    FAeroGlass := (DwmIsCompositionEnabled(FAeroGlass) = S_OK) and FAeroGlass;
    if FAeroGlass then
      FAeroGlass := (DwmGetWindowAttribute(FWnd, 1, @FAeroGlass, sizeof(FAeroGlass)) = S_OK) and FAeroGlass;
  end;
end;

procedure TClosingEffectWnd.InitRedirWnd;
begin
  FOldProc := Pointer(SetWindowLong(FWnd, GWL_WNDPROC, Integer(MakeObjectInstance(HookedWndProc))));
  SendMessage(FWnd, WM_REDIRCREATE, 0, 0);
end;

procedure TClosingEffectWnd.FinalRedirWnd;
begin
  if FRedirWnd <> 0 then
    SendMessage(FWnd, WM_REDIRDESTROY, 0, 0);
  FreeObjectInstance(Pointer(SetWindowLong(FWnd, GWL_WNDPROC, Integer(FOldProc))));
end;

procedure TClosingEffectWnd.CreateRedirWnd;
const
  nameChars: string = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ-_0123456789';
var
  wndClass: TWndClass;
  style, exstyle: Cardinal;
  rgn: HRGN;
  r, r2: TRect;
  name: array [0..80] of Char;
  bkatt: Integer;

  function RandomName: PChar;
  var
    i: Integer;
  begin
    for i := 0 to 7 do
      name[i] := nameChars[Random(Length(nameChars))+1];
    name[8] := #0;
    Result := name;
  end;


begin
  ZeroMemory(@wndClass, sizeof(wndClass));
  style := GetWindowLong(FWnd, GWL_STYLE);
  exstyle := GetWindowLong(FWnd, GWL_EXSTYLE);

  SetWindowLong(FWnd, GWL_EXSTYLE, exstyle or WS_EX_LAYERED);

  GetWindowRect(FWnd, r);
  rgn := CreateRectRgn(0,0,1,1);
  case GetWindowRgn(FWnd, rgn) of
    SIMPLEREGION:
    begin
      GetRgnBox(rgn, r2);
      OffsetRect(r2, r.Left, r.Top);
      if EqualRect(r, r2) then
      begin
        DeleteObject(rgn);
        rgn := 0;
      end;
    end;
    NULLREGION, ERROR:
    begin
      DeleteObject(rgn);
      rgn := 0;
    end;
  end;
  with WndClass do
  begin
    Style := GetClassLong(FWnd, GCL_STYLE);
    lpfnWndProc := @DefWindowProc;
    hInstance := SysInit.HInstance;
    hIcon := GetClassLong(FWnd, GCL_HICON);
    hCursor := GetClassLong(FWnd, GCL_HCURSOR);
    hbrBackground := GetStockObject(BLACK_BRUSH);
    repeat
      lpszClassName := PChar(RandomName);
      FRedirClass := Windows.RegisterClass(WndClass);
    until FRedirClass <> 0;
  end;
  GetWindowText(FWnd, name, 80);
  exstyle := exstyle and not WS_EX_APPWINDOW;
  exstyle := exstyle or WS_EX_LAYERED OR WS_EX_NOACTIVATE;

  FRedirWnd := CreateWindowEx(exstyle, PChar(FRedirClass),
                   name, style and not WS_VISIBLE, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, 0, 0, HInstance, nil);
  SetWindowLong(FRedirWnd, GWL_WNDPROC, Integer(MakeObjectInstance(RedirWndProc)));
  if rgn <> 0 then
    SetWindowRgn(FRedirWnd, rgn, True);

  SetWindowPos(FRedirWnd, FWnd, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, SWP_NOACTIVATE);

  if FAeroGlass then
  begin
    // since we have no way to copy Aero-glass border image, we have to turn aero-glass off
    bkatt := 1;  // disabled
    DwmSetWindowAttribute(FWnd, 2, @bkatt, sizeof(bkatt));
  end;

  //SetWindowPos(FRedirWnd, 0, 0, 0, 0, 0, ShowFlags[True]);
  SetTimer(FRedirWnd, 1, REDIR_CHECK_INTERVAL, nil);

end;

procedure TClosingEffectWnd.DestroyRedirWnd;
var
  bkatt: Integer;
begin
  KillTimer(FRedirWnd, 1);
  FreeObjectInstance(Pointer(SetWindowLong(FRedirWnd, GWL_WNDPROC, Integer(@DefWindowProc))));
  DestroyWindow(FRedirWnd);
  FRedirWnd := 0;
  Windows.UnregisterClass(PChar(FRedirClass), HInstance);
  FRedirClass := 0;
  if FClosing then
    ShowWindow(FWnd, SW_HIDE);
  if FAtt then
    SetLayeredWindowAttributes(FWnd, FAttC, FAttA, FAttF)
  else
    SetWindowLong(FWnd, GWL_EXSTYLE, GetWindowLong(FWnd, GWL_EXSTYLE) and not WS_EX_LAYERED);

  if FAeroGlass then
  begin
    bkatt := 2;  // enabled
    DwmSetWindowAttribute(FWnd, 2, @bkatt, sizeof(bkatt));
  end;

end;

procedure TClosingEffectWnd.Closing;
begin
  SendMessage(FWnd, WM_REDIRCLOSING, 1, 0);
end;

procedure TClosingEffectWnd.Restoring;
begin
  SendMessage(FWnd, WM_REDIRCLOSING, 0, 0);
end;

procedure TClosingEffectWnd.ClosingRedirWnd;
var
  Msg: TMsg;
begin
  FFading := True;
  FClosing := True;
  FAlphaEnd := 1;
  while FFading do
  begin
    while PeekMessage(Msg, 0, 0, 0, PM_REMOVE) do
    begin
      if Msg.message = WM_QUIT then Exit;
      TranslateMessage(Msg);
      DispatchMessage(Msg);
    end;
    if FFading then
      WaitMessage;
  end;
  DestroyRedirWnd;
end;

procedure TClosingEffectWnd.RestoringRedirWnd;
var
  Msg: TMsg;
begin
  FFading := True;
  FClosing := False;
  FAlphaEnd := FAttA;
  FAlphaStp := -FAlphaStp;
  while FFading do
  begin
    while PeekMessage(Msg, 0, 0, 0, PM_REMOVE) do
    begin
      if Msg.message = WM_QUIT then Exit;
      TranslateMessage(Msg);
      DispatchMessage(Msg);
    end;
    if FFading then
      WaitMessage;
  end;
  DestroyRedirWnd;
end;

procedure TClosingEffectWnd.RedirWndProc(var Message: TMessage);
begin
  with Message do
  begin
    if Msg = WM_TIMER then
      Render;
    Result := DefWindowProc(FRedirWnd, Msg, WParam, LParam);
  end;
end;

procedure TClosingEffectWnd.HookedWndProc(var Message: TMessage);
begin
  with Message do
  begin
    Result := 0;
    case Msg of
      WM_REDIRCREATE:
      begin
        CreateRedirWnd;
        Exit;
      end;
      WM_REDIRDESTROY:
      begin
        DestroyRedirWnd;
        Exit;
      end;
      WM_REDIRCLOSING:
      begin
        if WParam = 0 then
          RestoringRedirWnd
        else
          ClosingRedirWnd;
        Exit;
      end;
    end;
    Result := CallWindowProc(FOldProc, FWnd, Msg, WParam, Lparam);
  end;
end;

procedure TClosingEffectWnd.Render;
var
  hdr: TBitmapInfo;
  dc, mdc: HDC;
  bm: HBITMAP;
  bits, bits2: Pointer;
  r, r2: TRect;
  pt: TPoint;
  sz: TSize;
  rgn: HRGN;
begin
  KillTimer(FRedirWnd, 1);
  HideCaret(0);
  GetWindowRect(FWnd, r);
  GetWindowRect(FRedirWnd, r2);
  sz.cx := r.Right - r.Left;
  sz.cy := r.Bottom - r.Top;
  if not EqualRect(r, r2) or (sz.cx <> FCacheW) or (sz.cy <> FCacheH) then
  begin   // moved or resized
    if (sz.cx <> FCacheW) or (sz.cy <> FCacheH) then // resized
    begin
      // rebuild form region
      rgn := CreateRectRgn(0, 0, 1, 1);
      case GetWindowRgn(FWnd, rgn) of
        SIMPLEREGION:
        begin
          GetRgnBox(rgn, r2);
          OffsetRect(r2, r.Left, r.Top);
          if EqualRect(r, r2) then
          begin
            DeleteObject(rgn);
            rgn := 0;
          end;
        end;
        NULLREGION, ERROR:
        begin
          DeleteObject(rgn);
          rgn := 0;
        end;
      end;
      if rgn <> 0 then
        SetWindowRgn(FRedirWnd, rgn, False);

      FChg := True;
    end;

    SetWindowPos(FRedirWnd, FWnd, r.Left, r.Top, sz.cx, sz.cy, SWP_NOACTIVATE);
  end;

  // create DIB
  ZeroMemory(@hdr, sizeof(hdr));
  with hdr.bmiHeader do
  begin
    biSize := SizeOf(TBitmapInfoHeader);
    biWidth := sz.cx;
    biHeight := sz.cy;
    biPlanes := 1;
    biBitCount := 24;
    biCompression := BI_RGB;
  end;
  bm := CreateDIBSection(0, hdr, DIB_RGB_COLORS, bits, 0, 0);
  mdc := CreateCompatibleDC(0);
  DeleteObject(SelectObject(mdc, bm));

  // copy window surface image
  dc := GetWindowDC(FWnd);
  BitBlt(mdc, 0, 0, sz.cx, sz.cy, dc, 0, 0, SRCCOPY);
  ReleaseDC(FWnd, dc);

  // check if changed
  if CheckDigests(bits, sz.cx, sz.cy) then
    FChg := True;

  if FChg then  // rebuild surface
  begin
    FChg := False;
    if FCache <> 0 then
      DeleteObject(FCache);
    with hdr.bmiHeader do
      biBitCount := 32;
    FCache := CreateDIBSection(0, hdr, DIB_RGB_COLORS, bits2, 0, 0);
    if FDC <> 0 then
      SelectObject(FDC, FCache)
    else begin
      FDC := CreateCompatibleDC(0);
      DeleteObject(SelectObject(FDC, FCache));
    end;

    DoGrayScale(bits, bits2, sz.cx, sz.cy);

    FRedraw := True;
  end;

  AdjustFading;

  if FRedraw then  // update layered window
  begin
    pt := Point(0, 0);
    UpdateLayeredWindow(FRedirWnd, 0, nil, @sz, FDC, @pt, 0, @FBlend, ULW_ALPHA);
    SetWindowPos(FRedirWnd, FWnd, r.Left, r.Top, sz.cx, sz.cy, SWP_NOACTIVATE);

    {if FAeroGlass then
    begin
      blur.dwFlags := 7;
      blur.fEnable := True;
      blur.hRgnBlur := 0;
      blur.fTransitionOnMaximized := True;
      DwmEnableBlurBehindWindow(FRedirWnd, blur);
    end;
     }
    FRedraw := False;
  end;

  DeleteObject(bm);
  DeleteDC(mdc);
  ShowCaret(0);

  if IsWindowVisible(FWnd) <> IsWindowVisible(FRedirWnd) then
    SetWindowPos(FRedirWnd, 0, 0, 0, 0, 0, ShowingFlags[IsWindowVisible(FWnd)]);

  SetTimer(FRedirWnd, 1, REDIR_CHECK_INTERVAL, nil);
end;

function TClosingEffectWnd.CheckDigests(Bits: Pointer; w, h: Integer): Boolean;
var
  cnt: Integer;
  n, p, ln, stp, y: Integer;
  v: Cardinal;
begin
  Result := False;
  cnt := h div REDIR_CHECK_STEP;
  FDigests.Count := cnt;
  ln := BytesPerScanLine(w, 24, 32);
  stp := ln * REDIR_CHECK_STEP;
  y := REDIR_CHECK_STEP;
  n := 0;
  p := Integer(Bits) + (h-y-1)*ln;
  while y < h do
  begin
    v := CRC32(Pointer(p), ln, 1, 0);
    if v <> Cardinal(FDigests.Items[n]) then
    begin
      FDigests.Items[n] := Pointer(v);
      Result := True;
    end;
    Inc(n);
    Inc(y, REDIR_CHECK_STEP);
    Dec(p, stp);
  end;
end;

procedure TClosingEffectWnd.DoGrayScale(src, dst: Pointer; w, h: Integer);
var
  x, y, dln, sln, dpx, spx, dw, sw: Integer;
  f: Boolean;
  t: Cardinal;
  v: Byte;
begin
  dw := BytesPerScanLine(w, 32, 32);
  sw := BytesPerScanLine(w, 24, 32);
  dln := Integer(dst);
  sln := Integer(src);
  f := FAtt and (FAttF and LWA_COLORKEY <> 0);
  t := FAttC and $FFFFFF;
  for y := 1 to h do
  begin
    dpx := dln;
    spx := sln;
    for x := 1 to w do
    begin
      if f and (t = (PCardinal(spx)^ and $FFFFFF)) then
        PInteger(dpx)^ := 0
      else begin
        v := (GrayColorTable[0, PByte(spx)^] + GrayColorTable[1, PByte(spx+1)^] + GrayColorTable[2, PByte(spx+2)^]) shr 16;
        PByte(dpx)^ := v;
        PByte(dpx+1)^ := v;
        PByte(dpx+2)^ := v;
        PByte(dpx+3)^ := $FF;
      end;
      Inc(dpx, 4);
      Inc(spx, 3);
    end;
    Inc(dln, dw);
    Inc(sln, sw);
  end;
end;

procedure TClosingEffectWnd.AdjustFading;
var
  s: Cardinal;
begin
  if FFading then
  begin
    if FCurrAlpha = FAlphaEnd then
    begin
      if not FClosing then
      begin
        FFading := False;
        Exit;
      end;
    end
    else begin
      if ((FAlphaStp < 0) and (FCurrAlpha + FAlphaStp > FAlphaEnd)) or
           ((FAlphaStp > 0) and (FCurrAlpha + FAlphaStp < FAlphaEnd)) then
        Inc(FCurrAlpha, FAlphaStp)
      else
        FCurrAlpha := FAlphaEnd;
      SetLayeredWindowAttributes(FWnd, FAttC, FCurrAlpha, FAttF or LWA_ALPHA);
      //UpdateWindow(FWnd);
    end;
    if FClosing then
    begin
      FRedraw := True;
      if (FAlphaStp < 0) and (Integer(FBlend.SourceConstantAlpha) + FAlphaStp > 0) then
        Inc(FBlend.SourceConstantAlpha, FAlphaStp)
      else begin
        FBlend.SourceConstantAlpha := 0;
        FFading := False;
      end;
    end;
  end;
end;



var
  I, J: Integer;
initialization
  for I := 0 to 255 do
    for J := 0 to 255 do
      AlphaColorTable[I, J] := I * J div 255;
  for I := 0 to 255 do
  begin
    GrayColorTable[0, I] := I * 9437;   // 0.14
    GrayColorTable[1, I] := I * 36503;  // 0.56
    GrayColorTable[2, I] := I * 19595;  // 0.30
  end;

end.
