unit AutoFitScreen;
{实现窗体自适应调整尺寸以适应不同屏幕分辩率的显示问题。scg，2012年3月5日
}

interface
uses
  SysUtils, Windows, Classes, Graphics, Controls, Forms, Dialogs, Math,
  TypInfo;

var
  CurrWidth, CurrHeight: SmallInt;

const //记录设计时的屏幕分辨率
  OriWidth = 1366;
  OriHeight = 768;

type

  TfmForm = class(TForm) //实现窗体屏幕分辨率的自动调整
  private
    fScrResolutionRateW: Double;
    fScrResolutionRateH: Double;
    fIsFitDeviceDone: Boolean;
    procedure FitDeviceResolution;
  protected
    property IsFitDeviceDone: Boolean read fIsFitDeviceDone;
    property ScrResolutionRateH: Double read fScrResolutionRateH;
    property ScrResolutionRateW: Double read fScrResolutionRateW;
  public
    constructor Create(AOwner: TComponent); override;
  end;

  TResolutionForm = class(TfmForm) //增加对话框窗体的修改确认
  protected
    fIsDlgChange: Boolean;
  public
    constructor Create(AOwner: TComponent); override;
    property IsDlgChange: Boolean read fIsDlgChange default false;
  end;

implementation

function PropertyExists(const AObject: TObject; const APropName: string): Boolean;
 //判断一个属性是否存在
var
  PropInfo: PPropInfo;
begin
  PropInfo := GetPropInfo(AObject.ClassInfo, APropName);
  Result := Assigned(PropInfo);
end;

function GetObjectProperty(
  const AObject: TObject;
  const APropName: string
  ): TObject;
var
  PropInfo: PPropInfo;
begin
  Result := nil;
  PropInfo := GetPropInfo(AObject.ClassInfo, APropName);
  if Assigned(PropInfo) and
    (PropInfo^.PropType^.Kind = tkClass) then
    Result := GetObjectProp(AObject, PropInfo);
end;


constructor TfmForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  //if CurrWidth = 0  then
    CurrWidth := Screen.Width;
  //if CurrHeight = 0 then
    CurrHeight := Screen.Height;

  fScrResolutionRateH := 1;
  fScrResolutionRateW := 1;

  try
    if not fIsFitDeviceDone then
    begin
      FitDeviceResolution;
      fIsFitDeviceDone := True;
    end;
  except
    fIsFitDeviceDone := False;
  end;
end;

procedure TfmForm.FitDeviceResolution;
var
  LocList: TList;
  LocFontRate: Double;
  LocFontSize: Integer;
  LocFont: TFont;
  locK: Integer;
{计算尺度调整的基本参数}
  procedure CalBasicScalePars;
  begin
    try
      Self.Scaled := False;
      fScrResolutionRateH := screen.height / CurrHeight;
      fScrResolutionRateW := screen.Width / CurrWidth;
      LocFontRate := Min(fScrResolutionRateH, fScrResolutionRateW);
    except
      raise;
    end;
  end;

{保存原有坐标位置：利用递归法遍历各级容器里的控件，直到最后一级}
  procedure ControlsPostoList(vCtl: TControl; vList: TList);
  var
    locPRect: ^TRect;
    i: Integer;
    locCtl: TControl;
    locFontp: ^Integer;
  begin
    try
      New(locPRect);
      locPRect^ := vCtl.BoundsRect;
      vList.Add(locPRect);
      if PropertyExists(vCtl, 'FONT') then
      begin
        LocFont := TFont(GetObjectProperty(vCtl, 'FONT'));
        New(locFontp);
        locFontP^ := LocFont.Size;
        vList.Add(locFontP);
//        ShowMessage(vCtl.Name+'Ori:='+InttoStr(LocFont.Size));
      end;
      if vCtl is TWinControl then
        for i := 0 to TWinControl(vCtl).ControlCount - 1 do
        begin
          locCtl := TWinControl(vCtl).Controls[i];
          ControlsPosToList(locCtl, vList);
        end;
    except
      raise;
    end;
  end;

{计算新的坐标位置：利用递归法遍历各级容器里的控件，直到最后一层。
 计算坐标时先计算顶级容器级的，然后逐级递进}
  procedure AdjustControlsScale(vCtl: TControl; vList: TList; var vK: Integer);
  var
    locOriRect, LocNewRect: TRect;
    i: Integer;
    locCtl: TControl;
  begin
    try
      if vCtl.Align <> alClient then
      begin
        locOriRect := TRect(vList.Items[vK]^);
        with locNewRect do
        begin
          Left := Round(locOriRect.Left * fScrResolutionRateW);
          Right := Round(locOriRect.Right * fScrResolutionRateW);
          Top := Round(locOriRect.Top * fScrResolutionRateH);
          Bottom := Round(locOriRect.Bottom * fScrResolutionRateH);
          vCtl.SetBounds(Left, Top, Right - Left, Bottom - Top);
        end;
      end;
      if PropertyExists(vCtl, 'FONT') then
      begin
        Inc(vK);
        LocFont := TFont(GetObjectProperty(vCtl, 'FONT'));
        locFontSize := Integer(vList.Items[vK]^);
        LocFont.Size := Round(LocFontRate * locFontSize);
//        ShowMessage(vCtl.Name+'New:='+InttoStr(LocFont.Size));
      end;
      Inc(vK);
      if vCtl is TWinControl then
        for i := 0 to TwinControl(vCtl).ControlCount - 1 do
        begin
          locCtl := TWinControl(vCtl).Controls[i];
          AdjustControlsScale(locCtl, vList, vK);
        end;
    except
      raise;
    end;
  end;

{释放坐标位置指针和列表对象}
  procedure FreeListItem(vList: TList);
  var
    i: Integer;
  begin
    for i := 0 to vList.Count - 1 do
      Dispose(vList.Items[i]);
    vList.Free;
  end;

begin
  LocList := TList.Create;
  try
    try
      if (Screen.width <> CurrWidth) or (Screen.Height <> CurrHeight) then
      begin
        CalBasicScalePars;
//        AdjustComponentFont(Self);
        ControlsPostoList(Self, locList);
        locK := 0;
        AdjustControlsScale(Self, locList, locK);

      end;
    except on E: Exception do
        raise Exception.Create('进行屏幕分辨率自适应调整时出现错误' + E.Message);
    end;
  finally
    FreeListItem(locList);
  end;
end;


{ TResolutionForm }

constructor TResolutionForm.Create(AOwner: TComponent);
begin
  inherited;
  fIsDlgChange := False;
end;

end.

