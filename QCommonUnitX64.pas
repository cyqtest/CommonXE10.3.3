{*******************************************************}
{                                                       }
{     QCommonUnit v1.0  by QX 2008.07.16               }
{                                                       }
{                                                       }
{     单元功能：控件闪动/映射表/事件链表                }
{                                                       }
{*******************************************************}

unit QCommonUnitX64;

interface

uses
  SysUtils, ExtCtrls, Controls, Graphics, Windows, Forms, Classes, Menus;

type

  //辅助类，用于将TWinControl的Color属性公开
  //
  TColorCtrl = class(TWinControl)
  public
    property Color;
  end;

  TFlashObject = class
  private
    FCtrl: TWinControl;
    FPrimaryColor: TColor;
    FChangeColor: TColor;
    FCount: Integer;
    FNext: TFlashObject;
  public
    constructor Create;
    destructor Destroy; override;
    procedure InitCount(const Value: Integer);
    function EndCount: Boolean;
    property Ctrl: TWinControl read FCtrl write FCtrl;
    property PrimaryColor: TColor read FPrimaryColor write FPrimaryColor;
    property ChangeColor: TColor read FChangeColor write FChangeColor;
    property NextObject: TFlashObject read FNext write FNext;
  end;

  TFlashCtrls = class
  private
    FTimer: TTimer;
    FFlashCount: Integer;

    FObject: TFlashObject;
    procedure StartFlash;
    procedure EndFlash;
    function DelCtrl(AObject: TFlashObject): TFlashObject;
    function GetPriorObject(AObject: TFlashObject): TFlashObject;
    procedure DoTimer(Sender: TObject);
    procedure FreeAll;
    function FindCtrl(const Ctrl: TWinControl): TFlashObject;
  public
    constructor Create;
    destructor Destroy; override;

    function AddCtrl(const Ctrl: TWinControl; const Color: TColor = clRed;
      const FlashCount: Integer = 6): TFlashObject;
  end;
  
  PValueItem = ^TValueItem;
  TValueItem = record
    FName: array[0..47] of Char;
    FValue: Integer;
    FFloatValue: Extended;
    FIndex: Integer;
  end;

  TItemFreeEvent = procedure(Item: TValueItem);

  TValueMap = class
  private
    Values: TList;
    FOnBeforeFree: TItemFreeEvent;
    procedure DoBeforeFree(Item: TValueItem);
    function GetValue(Name: string): Integer;
    procedure SetValue(Name: string; const AValue: Integer);
    function GetFloatValue(Name: string): Extended;
    procedure SetFloatValue(Name: string; const eValue: Extended);
  public
    constructor Create;
    destructor Destroy; override;
    procedure Clear; virtual;
    function ExistsValue(const Name: string): Boolean;
    procedure DeleteValue(const Name: string);
    function FindValue(const Name: string): PValueItem;

    function AddValue(const Name: string; const iValue: Integer): Boolean; overload;
    function AddValue(const Name: string; const eValue: Extended): Boolean; overload;
    function AddValue(const Name: string): Boolean; overload;
    property Value[Name: string]: Integer read GetValue write SetValue;
    property FloatValue[Name: string]: Extended read GetFloatValue write SetFloatValue;
    property OnBeforeFree: TItemFreeEvent read FOnBeforeFree write FOnBeforeFree;
  end;

  PMethod = ^TObjectMethod;
  TObjectMethod = record
    MethodAddr: Pointer;
    MethodOwner: Pointer;
    MethodName: array[0..23] of Char;
  end;

  TObjectMethodList = class
  private
    FMethodList: TList;
    function GetMethodCount: Integer;
    function GetMethod(AIndex: Integer): TMethod;
  public
    procedure AddMethod(const MethodName: string; const AMethod: TMethod);
    procedure DelMethod(const AMethod: TMethod);
    //function GetMethod(const MethodName: string): TMethod;
    procedure Clear;
    procedure ExcuteMethod;
    property MethodCount: Integer read GetMethodCount;
    property MethodList[AIndex: Integer]: TMethod read GetMethod; default;
  end;

  procedure SwapValue(var A, B: Integer);
  procedure HintAtCtrl(const Ctrl: TWinControl; const Focused: Boolean = True;
    const Color: TColor = clRed; const FlashCount: Integer = 6);
  procedure SetMenuVisible(PopupMenu: TPopupMenu; Flag: Boolean);

var
  FlashCtrls: TFlashCtrls;
  
implementation

procedure SetMenuVisible(PopupMenu: TPopupMenu; Flag: Boolean);
var
  I: Integer;
begin
  with PopupMenu do
  for I := 0 to Items.Count - 1 do
    Items[I].Visible := not Flag;
end;

procedure HintAtCtrl(const Ctrl: TWinControl; const Focused: Boolean;
  const Color: TColor; const FlashCount: Integer);
begin
  FlashCtrls.AddCtrl(Ctrl, Color, FlashCount);
  if Focused and Ctrl.Enabled and Ctrl.Visible then
  try
    Ctrl.SetFocus;
  except
    MessageBox(GetParentForm(Ctrl).Handle, '无法设置焦点！', '提示信息！', MB_ICONWARNING);
  end;
end;

procedure InitUnit;
begin
  FlashCtrls := TFlashCtrls.Create;
end;

procedure ReleaseUnit;
begin
  FreeAndNil(FlashCtrls);
end;

{Functions}
procedure SwapValue(var A, B: Integer);
asm
  PUSH [EAX]
  MOV  ECX, [EDX]
  MOV  [EAX], ECX
  POP  ECX
  MOV  [EDX], ECX
  RET
end;

{ TFlashObject }

constructor TFlashObject.Create;
begin
  FCount := 0;
  FCtrl := nil;
end;

destructor TFlashObject.Destroy;
begin  
  inherited;
end;

function TFlashObject.EndCount: Boolean;
begin
  Dec(FCount);
  Result := FCount <= 0;
end;

procedure TFlashObject.InitCount(const Value: Integer);
begin
  if (FCount <> Value) and (Value > 0) then
    FCount := Value;
end;


{ TFlashCtrls }

function TFlashCtrls.FindCtrl(const Ctrl: TWinControl): TFlashObject;
begin
  if Assigned(Ctrl) then
  begin
    Result := FObject;
    while Assigned(Result) and (Result.Ctrl <> Ctrl) do
      Result := Result.NextObject;
  end
  else Result := nil;
end;

function TFlashCtrls.AddCtrl(const Ctrl: TWinControl;
  const Color: TColor = clRed; const FlashCount: Integer = 6): TFlashObject;
begin
  if Assigned(Ctrl) then
  begin
    Result := FindCtrl(Ctrl);
    if Assigned(Result) then
    begin
      Result.InitCount(FlashCount);
    end
    else
    begin
      Result := TFlashObject.Create;
      Result.InitCount(FlashCount);
      Result.Ctrl := Ctrl;
      Result.PrimaryColor := TColorCtrl(Ctrl).Color;
      Result.ChangeColor := Color;
      Result.NextObject := FObject;
      FObject := Result;
      StartFlash;
    end;
  end
  else raise Exception.Create('无效控件指针！');
end;

function TFlashCtrls.GetPriorObject(AObject: TFlashObject): TFlashObject;
begin
  if Assigned(AObject) then
  begin
    Result := FObject;
    while Assigned(Result) and (Result.NextObject <> AObject) do
      Result := Result.NextObject;
  end
  else Result := nil;
end;

function TFlashCtrls.DelCtrl(AObject: TFlashObject): TFlashObject;
var
  Instance: TFlashObject;
begin
  if AObject = FObject then
  begin
    FObject := AObject.NextObject;
  end
  else
  begin
    Instance := GetPriorObject(AObject);
    if Assigned(Instance) then
      Instance.NextObject := AObject.NextObject
    else
      raise Exception.Create('无法搜索到对应的指针！');
  end;
  Result := AObject.NextObject;
  FreeAndNil(AObject);
  EndFlash;
end;

constructor TFlashCtrls.Create;
begin
  inherited Create;
  FTimer := TTimer.Create(nil);
  FTimer.Interval := 100;
  FTimer.Enabled := False;
  FTimer.OnTimer := DoTimer;
end;

destructor TFlashCtrls.Destroy;
begin
  FreeAndNil(FTimer);
  FreeAll;
  inherited;
end;

procedure TFlashCtrls.FreeAll;
var
  Instance: TFlashObject;
begin
  while Assigned(FObject) do
  begin
    Instance := FObject.NextObject;
    FreeAndNil(FObject);
    FObject := Instance;
  end;
end;

procedure TFlashCtrls.EndFlash;
begin
  Dec(FFlashCount);
  if FFlashCount = 0 then
  begin
    FTimer.Enabled := False;
    FreeAll;
  end;
end;

procedure TFlashCtrls.StartFlash;
begin
  Inc(FFlashCount);
  if not FTimer.Enabled then FTimer.Enabled := True;
end;

procedure TFlashCtrls.DoTimer(Sender: TObject);
var
  Instance: TFlashObject;
  clTemp: TColor;
begin
  Instance := FObject;
  while Assigned(Instance) do
  begin
    try
      if Instance.EndCount then
      begin
        TColorCtrl(Instance.Ctrl).Color := Instance.PrimaryColor;
        Instance := DelCtrl(Instance);
        Continue;
      end
      else
      begin
        clTemp := TColorCtrl(Instance.Ctrl).Color;
        TColorCtrl(Instance.Ctrl).Color := Instance.ChangeColor;
        Instance.ChangeColor := clTemp;
      end;
    except
      Instance := DelCtrl(Instance);
      Continue;
    end;
    if Assigned(Instance) then
      Instance := Instance.NextObject;
  end;
  if FObject = nil then FTimer.Enabled := False;
end;

{ TValueMap }

function TValueMap.AddValue(const Name: string; const iValue: Integer): Boolean;
var
  Value: PValueItem;
begin
  Result := False;
  if Assigned(FindValue(Name)) then
    Exit;
  New(Value);
  try
    StrPCopy(Value^.FName, Name);
    Value^.FValue := iValue;
    Values.Add(Value);
    Result := True;
  except
    Dispose(Value);
    raise Exception.Create('内存分配错误！');
  end;
end;

function TValueMap.AddValue(const Name: string): Boolean;
begin
  Result := AddValue(Name, 0);
end;

constructor TValueMap.Create;
begin
  inherited Create;
  Values := TList.Create;
end;

procedure TValueMap.DeleteValue(const Name: string);
var
  Value: PValueItem;
begin
  Value := FindValue(Name);
  if not Assigned(Value) then
  begin
    raise Exception.Create('不存在此变量名！');
  end;
  DoBeforeFree(Value^);
  Values.Remove(Value);
  //Values.Delete(Value.FIndex);
  Dispose(Value);
end;

destructor TValueMap.Destroy;
begin
  Clear;
  Values.Free;
  inherited;
end;

procedure TValueMap.DoBeforeFree(Item: TValueItem);
begin
  if Assigned(FOnBeforeFree) then FOnBeforeFree(Item);
end;

procedure TValueMap.Clear;
var
  vi: PValueItem;
begin
  while Values.Count > 0 do
  begin
    vi := Values.Last;
    DoBeforeFree(vi^);
    Values.Remove(vi);
    Dispose(vi);
  end;
end;

function TValueMap.FindValue(const Name: string): PValueItem;
var
  i: Integer;
begin
  for i := Values.Count - 1 downto 0 do
  begin
    if SameText(Name, TValueItem(Values[i]^).FName) then
    begin
      Result := Values[i];
      Result.FIndex := i;
      Exit;
    end;
  end;
  Result := nil;
end;

function TValueMap.GetValue(Name: string): Integer;
var
  Value: PValueItem;
begin
  Value := FindValue(Name);
  if not Assigned(Value) then
    raise Exception.Create('不存在此变量名！');
  Result := Value^.FValue;
end;

procedure TValueMap.SetValue(Name: string; const AValue: Integer);
var
  Value: PValueItem;
begin
  Value := FindValue(Name);
  if not Assigned(Value) then
  begin
    raise Exception.Create('不存在此变量名！');
  end;
  Value^.FValue := AValue;
end;

function TValueMap.ExistsValue(const Name: string): Boolean;
var
  I: Integer;
begin
  for I := 0 to Values.Count - 1 do
  begin
    Result := SameText(TValueItem(Values[I]^).FName, Name);
    if Result then Exit;
  end;
  Result := False;
end;

function TValueMap.AddValue(const Name: string;
  const eValue: Extended): Boolean;
var
  Value: PValueItem;
begin
  Result := AddValue(Name, 0);
  if Result then
  begin
    Value := FindValue(Name);
    Value^.FFloatValue := eValue;
  end;
end;

function TValueMap.GetFloatValue(Name: string): Extended;
var
  Value: PValueItem;
begin
  Value := FindValue(Name);
  if not Assigned(Value) then
    raise Exception.Create('不存在此变量名！');
  Result := Value^.FFloatValue;
end;

procedure TValueMap.SetFloatValue(Name: string; const eValue: Extended);
var
  Value: PValueItem;
begin
  Value := FindValue(Name);
  if not Assigned(Value) then
  begin
    raise Exception.Create('不存在此变量名！');
  end;
  Value^.FFloatValue := eValue;
end;

{ TObjectMethodList }

procedure TObjectMethodList.AddMethod(const MethodName: string; const AMethod: TMethod);
var
  p: PMethod;
begin
  if Assigned(AMethod.Code) and Assigned(AMethod.Data) then
  begin
    if FMethodList = nil then FMethodList := TList.Create;
    New(p);
    try
      p^.MethodAddr := AMethod.Code;
      p^.MethodOwner := AMethod.Data;
      MoveMemory(@p^.MethodName, @MethodName[1], Length(p^.MethodName));
      FMethodList.Add(p);
    except
      Dispose(p);
    end;
  end;
end;

procedure TObjectMethodList.DelMethod(const AMethod: TMethod);
var
  p: PMethod;
  I: Integer;
begin
  if FMethodList = nil then Exit;
  for I := 0 to FMethodList.Count - 1 do
  begin
    P := FMethodList[I];
    if (p.MethodAddr = AMethod.Code)
      and (p.MethodOwner = AMethod.Data) then
    begin
      Dispose(p);
      FMethodList.Delete(I);
      Break;
    end;
  end;
  if FMethodList.Count = 0 then FreeAndNil(FMethodList);
end;

//function TObjectMethodList.GetMethod(const MethodName: string): TMethod;
//var
//  I: Integer;
//  P: PMethod;
//begin
//  FillChar(Result, SizeOf(TMethod), 0);
//  if Assigned(FMethodList) then
//    for I := 0 to FMethodList.Count - 1 do
//    begin
//      P := FMethodList[I];
//      if SameText(StrPas(@P.MethodName), MethodName) then
//      begin
//        Result.Code := P.MethodAddr;
//        Result.Data := P.MethodOwner;
//        Exit;
//      end;
//    end;
//end;

procedure TObjectMethodList.Clear;
var
  p: PMethod;
  I: Integer;
begin
  if FMethodList <> nil then
  begin
    for I := FMethodList.Count - 1 downto 0 do
    begin
      P := FMethodList[I];
      Dispose(p);
      FMethodList.Delete(I);
    end;
    FreeAndNil(FMethodList);
  end;
end;

procedure TObjectMethodList.ExcuteMethod;
var
  p: PMethod;
  I: Integer;
begin
  if FMethodList <> nil then
  begin
    for I := 0 to FMethodList.Count - 1 do
    begin
      p := FMethodList[I];
      if Assigned(P.MethodAddr) and Assigned(P.MethodOwner) then
      asm
        MOV  EDX, P
        MOV  EAX, [EDX].LongInt[4]
        CALL [EDX].POINTER
      end;
    end;
  end;
end;

function TObjectMethodList.GetMethodCount: Integer;
begin
  if Assigned(FMethodList) then
    Result := FMethodList.Count
  else
    Result := -1;
end;

function TObjectMethodList.GetMethod(AIndex: Integer): TMethod;
var
  P: PMethod;
begin
  FillChar(Result, SizeOf(Result), 0);
  if Assigned(FMethodList) then
  begin
    if AIndex < FMethodList.Count then
    begin
      P := FMethodList[AIndex];
      Result.Code := P.MethodAddr;
      Result.Data := P.MethodOwner;
    end;
  end;
end;

initialization
  InitUnit;

finalization
  ReleaseUnit;

end.
