{*********************************************************************}
{                                                                     }
{     DBGridEhToExcelEX v1.0  Modify By QX,CYQ 2018年8月24日9:12:49   }
{                                                                     }
{                                                                     }
{     单元功能：支持DBgridEh多表头导出到XLS                           }
{               支持Office2003-office2013格式                         }
{     注意：    该单元需要XLSReadWriteII5控件。                       }
{     Frome：   修改网上不知名的哪位大侠的。原来用的是XLSReadWriteII2 }
{               修改多处Bug                                           }
{                                                                     }
{*********************************************************************}
unit DBGridEhToExcelEX;

interface

uses
  Windows, Messages, SysUtils, Classes, DBGridEh, XLSReadWriteII5, Xc12DataStyleSheet5,
  Forms, Controls, Gauges, ShellApi, DB, ExtCtrls, StdCtrls, Graphics;

type
  TGridHeaderItem = class(TCollectionItem)
  private
    FCaption: string;
    FFontName: string;
    FFontSize: integer;
    FBold: Boolean;
    //FItalic: Boolean;
    //FUnderline: Boolean;
    //FStrikeOut: Boolean;
    FMerged: Boolean;
    FAlignment: integer;
  public
    property Caption: string read FCaption write FCaption;
    property FontName: string read FFontName write FFontName;
    property FontSize: Integer read FFontSize write FFontSize;
    property Bold: Boolean read FBold write FBold;
    property Merged: Boolean read FMerged write FMerged;
    property Alignment: integer read FAlignment write FAlignment;
  end;

  TGridHeader = class(TCollection)
  private
    function GetItem(Index: integer): TGridHeaderItem;
    procedure SetItem(Index: Integer; Value: TGridHeaderItem);
  public
    property Items[Index: Integer]: TGridHeaderItem read GetItem write SetItem;
    function Add: TGridHeaderItem;
  end;

  TDBGridEhToExcelEx = class(TComponent)
  private
    { Private declarations }
    FGauge: TGauge; {进度条}
    FProgressForm: TForm; {进度窗体}
    FShowProgress: Boolean; {是否显示进度窗体}
    FDBGridEh: TDBGridEh; //对应DBGridEh
    FFileName: string; {保存文件名}
    FTitleRowCount: integer;
    FStartRow: integer;
    FColumnCount: integer;
    FGridHeader: TGridHeader;
    procedure SetDBGridEh(const Value: TDBGridEh);
    procedure SetGridHeader(const Value: TGridHeader);
    procedure SetFileName(const Value: string);
    procedure SetTitleRowCount(const Value: integer);
    procedure WriteGridHeader(XLS: TXLSReadWriteII5);
    procedure WriteTitleData(XLS: TXLSReadWriteII5);
    procedure WriteDataCell(XLS: TXLSReadWriteII5);
    procedure WriteFooter(XLS: TXLSReadWriteII5);
    procedure FormatDataCell(XLS: TXLSReadWriteII5);
    function GetColumnCount: integer;
    function GetTitleHeightFull: integer;
    function GetRowHeight: integer;
    procedure SetShowProgress(const Value: Boolean);
    procedure CreateProcessForm(AOwner: TComponent);
  protected
    { Protected declarations }
    property TitleRowCount: integer read FTitleRowCount write SetTitleRowCount;
    property ColumnCount: integer read FColumnCount;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ExportToExcel(OpenFile: Boolean); {输出Excel文件}
  published
    { Published declarations }
    property DBGridEh: TDBGridEh read FDBGridEh write SetDBGridEh;
    property FileName: string read FFileName write SetFileName;
    property ShowProgress: Boolean read FShowProgress write SetShowProgress;
    property GridHeader: TGridHeader read FGridHeader write SetGridHeader;
    procedure AddGridHeader(aCaption, aFontName: string; aFontSize: Integer; aBold, aMerged: Boolean; aAlignment: integer);
  end;

implementation

{ TGridHeader }

function TGridHeader.Add: TGridHeaderItem;
begin
  Result := TGridHeaderItem(inherited Add);
end;

function TGridHeader.GetItem(Index: Integer): TGridHeaderItem;
begin
  Result := TGridHeaderItem(inherited GetItem(Index));
end;

procedure TGridHeader.SetItem(Index: Integer; Value: TGridHeaderItem);
begin
  inherited SetItem(Index, Value);
end;
{ TDBGridEhToExcelEx }

procedure TDBGridEhToExcelEx.AddGridHeader(aCaption, aFontName: string; aFontSize: Integer; aBold, aMerged: Boolean; aAlignment: integer);
begin
  with FGridHeader.Add do
  begin
    Caption := ACaption;
    FontName := AFontName;
    FontSize := AFontSize;
    Bold := ABold;
    Merged := AMerged;
    Alignment := AAlignment;
  end;
end;

constructor TDBGridEhToExcelEx.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FGridHeader := TGridHeader.Create(TGridHeaderItem);
  FShowProgress := True;
  FStartRow := 0;
end;

function TDBGridEhToExcelEx.GetTitleHeightFull: integer;
var
  I, J, K, Interlinear, TitleHeightFull: integer;
  tm: TTEXTMETRIC;
  Canvas: TCanvas;
begin
  TitleHeightFull := 0;
  if dgTitles in DBGridEh.Options then
  begin
    if DBGridEh.Flat then Interlinear := 2 else Interlinear := 4;
    Canvas := TCanvas.Create;
    Canvas.Handle := GetDC(0);
    try
      if DBGridEh.UseMultiTitle = True then
        TitleHeightFull := Canvas.TextHeight('Wg') + DBGridEh.VTitleMargin * 2
      else begin
        K := 0;
        for I := 0 to DBGridEh.Columns.Count - 1 do
        begin
          Canvas.Font := DBGridEh.Columns[I].Title.Font;
          J := Canvas.TextHeight('Wg') + Interlinear;
          if J > K then K := J;
        end;
        if K = 0 then
        begin
          Canvas.Font := DBGridEh.TitleFont;
          K := Canvas.TextHeight('Wg') + Interlinear;
        end;
        TitleHeightFull := K;
        if (DBGridEh.TitleHeight <> 0) or (DBGridEh.TitleLines <> 0) then
        begin
          K := 0;
          for I := 0 to DBGridEh.Columns.Count - 1 do
          begin
            Canvas.Font := DBGridEh.Columns[I].Title.Font;
            J := Canvas.TextHeight('Wg') + Interlinear;
            if J > K then
            begin
              K := J;
              GetTextMetrics(Canvas.Handle, tm);
            end;
          end;
          if K = 0 then
          begin
            Canvas.Font := DBGridEh.TitleFont;
            GetTextMetrics(Canvas.Handle, tm);
          end;
          TitleHeightFull := tm.tmExternalLeading + tm.tmHeight * DBGridEh.TitleLines + 2 + DBGridEh.TitleHeight;
          if dgRowLines in DBGridEh.Options then TitleHeightFull := TitleHeightFull + 1;
        end;
      end;
    finally
      ReleaseDC(0, Canvas.Handle);
      Canvas.Handle := 0;
      Canvas.Free;
    end;
  end;
  Result := TitleHeightFull;
end;

function TDBGridEhToExcelEx.GetRowHeight: integer;
var
  I, J, K, GridLineWidth, DefaultRowHeight: Integer;
  tm: TTEXTMETRIC;
  Canvas: TCanvas;
begin
  GridLineWidth := 1;
  Canvas := TCanvas.Create;
  Canvas.Handle := GetDC(0);
  try
    Canvas.Font := DBGridEh.Font;
    if DBGridEh.Flat
      then J := 1
    else J := 3;
    if dgRowLines in DBGridEh.Options then
      Inc(J, GridLineWidth);
    K := Canvas.TextHeight('Wg');
    GetTextMetrics(Canvas.Handle, tm);
    if (DBGridEh.RowHeight > 0) or (DBGridEh.RowLines > 0)
      then DefaultRowHeight := DBGridEh.RowHeight + (tm.tmExternalLeading + tm.tmHeight) * DBGridEh.RowLines
    else DefaultRowHeight := K + J;
    if (dghFitRowHeightToText in DBGridEh.OptionsEh) then
    begin
      I := (DefaultRowHeight - J) mod K;
      if (I > K div 2) or ((DefaultRowHeight - J) div K = 0)
        then DefaultRowHeight := ((DefaultRowHeight - J) div K + 1) * K + J
      else DefaultRowHeight := (DefaultRowHeight - J) div K * K + J;
      DBGridEh.RowLines := (DefaultRowHeight - J) div K;
      DBGridEh.RowHeight := J;
    end;
    //if (tm.tmExternalLeading + tm.tmHeight + tm.tmInternalLeading + FInterlinear < DefaultRowHeight)
      //then DBGridEh.AllowWordWrap := True
      //else DBGridEh.AllowWordWrap := False;
  finally
    ReleaseDC(0, Canvas.Handle);
    Canvas.Handle := 0;
    Canvas.Free;
  end;
  Result := DefaultRowHeight;
end;

function TDBGridEhToExcelEx.GetColumnCount: integer;
var
  i, ColumnCount: integer;
begin
  ColumnCount := 0;
  for i := 0 to FDBGridEh.VisibleColumns.Count - 1 do
  begin
    if FDBGridEh.VisibleColumns[i].Visible then
      Inc(ColumnCount);
  end;
  Result := ColumnCount;
end;

procedure TDBGridEhToExcelEx.SetGridHeader(const Value: TGridHeader);
begin
  FGridHeader.Assign(Value);
end;

procedure TDBGridEhToExcelEx.SetDBGridEh(const Value: TDBGridEh);
begin
  FDBGridEh := Value;
end;

procedure TDBGridEhToExcelEx.SetFileName(const Value: string);
begin
  FFileName := Value;
end;

procedure TDBGridEhToExcelEx.SetTitleRowCount(const Value: integer);
begin
  FTitleRowCount := Value;
end;

procedure TDBGridEhToExcelEx.WriteGridHeader(XLS: TXLSReadWriteII5);
var
  i: integer;
begin
  FColumnCount := GetColumnCount;
  for i := 0 to FGridHeader.Count - 1 do
  begin
    XLS.Sheets[0].AsString[0, i] := FGridHeader.Items[i].Caption;
    XLS.Sheets[0].Cell[0, i].FontName := FGridHeader.Items[i].FontName;
    XLS.Sheets[0].Cell[0, i].FontSize := FGridHeader.Items[i].FontSize;
    if FGridHeader.Items[i].Bold then
      XLS.Sheets[0].Cell[0, i].FontStyle := [xfsBold];
    case FGridHeader.Items[i].Alignment of
      -1: XLS.Sheets[0].Cell[0, i].HorizAlignment := chaLeft;
      0: XLS.Sheets[0].Cell[0, i].HorizAlignment := chaCenter;
      1: XLS.Sheets[0].Cell[0, i].HorizAlignment := chaRight;
    end;
    if FGridHeader.Items[i].Merged then
      XLS.Sheets[0].MergedCells.Add(0, i, FColumnCount - 1, i);
      //XLS.Sheets[0].Range.Items[0,i,FColumnCount-1,i].Merged:=True;
  end;
  FStartRow := FGridHeader.Count;
end;

procedure TDBGridEhToExcelEx.WriteTitleData(XLS: TXLSReadWriteII5);

  function GetSubStr(SepChar: Char; var aString: string): string;
  var
    MyStr: string;
    SepCharPos: Integer;
  begin
    SepCharPos := Pos(SepChar, aString);
    MyStr := Copy(aString, 1, SepCharPos - 1);
    Delete(aString, 1, SepCharPos);
    GetSubStr := MyStr;
  end;

var
  TitleCell: array of array of string;
  TitleCellJB, AvailableJB: array of integer;
  ParentTitleCell: array of string;
  I, J, M, N, P, FJB, Width, Height, Col, Row, StartPos, MaxJB: Integer;
  Caption: string;
  FitColWidths: Boolean;
begin
  if FDBGridEh.UseMultiTitle then
  begin
    //下面代码计算并分割标题行数据
    FTitleRowCount := 0;
    SetLength(TitleCell, FColumnCount);
    SetLength(TitleCellJB, FColumnCount);
    for i := 0 to FColumnCount - 1 do
    begin
      FJB := 0;
      Caption := FDBGridEh.{Columns}VisibleColumns[i].Title.Caption + '|';
      while Caption <> '' do
      begin
        inc(FJB);
        SetLength(TitleCell[i], FJB);
        TitleCell[i, FJB - 1] := GetSubStr('|', Caption);
        if FTitleRowCount < FJB then FTitleRowCount := FJB; //计算总级别
      end;
      TitleCellJB[i] := FJB;
    end;
    //下面导出标题行数据
    StartPos := 0;
    SetLength(AvailableJB, FColumnCount);
    SetLength(ParentTitleCell, FColumnCount);
    for i := 0 to FColumnCount - 1 do
      AvailableJB[i] := FTitleRowCount;

    for i := 0 to FTitleRowCount - 1 do //每一个级别
    begin
      for j := 0 to FColumnCount - 1 do
      begin
        if TitleCellJB[j] >= i + 1 then
        begin
          StartPos := j;
          Break;
        end;
      end;

      for j := 0 to FColumnCount - 1 do
      begin
        if i = 0 then ParentTitleCell[j] := ''
        else
          if TitleCellJB[j] - 1 > i then ParentTitleCell[j] := TitleCell[j, i - 1];
      end;

      while StartPos <= FColumnCount - 1 do
      begin
        Width := 1;
        Caption := TitleCell[StartPos, i];
        for m := StartPos + 1 to FColumnCount do
        begin
          if TitleCellJB[m] < i + 1 then continue;
          if (TitleCell[m, i] = Caption) and (ParentTitleCell[StartPos] = ParentTitleCell[m]) then inc(Width)
          else break;
        end;
        MaxJB := 0;
        for n := StartPos to StartPos + Width - 1 do
          if TitleCellJB[n] - i > MaxJB then MaxJB := TitleCellJB[n] - i;
        Height := AvailableJB[StartPos] - MaxJB + 1;
        Col := StartPos;
        Row := FStartRow + FTitleRowCount - AvailableJB[StartPos];
        XLS.Sheets[0].AsString[Col, Row] := Caption;
        if (Height <> 1) or (Width <> 1) then
          XLS.Sheets[0].MergedCells.Add(Col, Row, Col + Width - 1, Row + Height - 1);
          //XLS.Sheets[0].Range.Items[Col,Row,Col+Width-1,Row+Height-1].Merged:=True;
        for p := StartPos to StartPos + Width - 1 do
          AvailableJB[p] := AvailableJB[p] - Height;

        //Inc(StartPos);
        StartPos := m;
      end;
    end;
  end else
  begin
    FTitleRowCount := 1;
    for I := 0 to FColumnCount - 1 do
      XLS.Sheets[0].AsString[I, FStartRow] := FDBGridEh.VisibleColumns[I].Title.Caption;
  end;
  Height := FStartRow + FTitleRowCount - 1;
  Width := FColumnCount - 1;
  //cancel by cyq 2021年5月24日 11:45:32
  with XLS.Sheets[0].Range.Items[0, FStartRow, Width, Height] do
  begin
    BorderOutlineStyle := cbsThin;
    BorderInsideVertStyle := cbsThin;
    BorderInsideHorizStyle := cbsThin;

    HorizAlignment := chaCenter;
    VertAlignment := cvaCenter;
    FontCharset := FDBGridEh.TitleFont.Charset;
    FontName := FDBGridEh.TitleFont.Name;
    FontSize := FDBGridEh.TitleFont.Size;

    FontColor := FDBGridEh.TitleFont.Color;
  end;

  //下面代码设置列宽,导出数据时应将AutoFitColWidths属性值设置为False
  XLS.Sheets[0].Columns.AddIfNone(0, FColumnCount);
  FitColWidths := FDBGridEh.AutoFitColWidths;
  FDBGridEh.AutoFitColWidths := False;
  for i := 0 to FColumnCount - 1 do
    XLS.Sheets[0].Columns[i].PixelWidth := FDBGridEh.VisibleColumns[i].Width;
  FDBGridEh.AutoFitColWidths := FitColWidths;
  //下面代码设置标题行行高
  //XLS.Sheets[0].Rows.AddIfNone(FStartRow,FTitleRowCount);
  for i := FStartRow to FStartRow + FTitleRowCount - 1 do
    XLS.Sheets[0].Rows[i].PixelHeight := GetTitleHeightFull;
end;

procedure TDBGridEhToExcelEx.WriteDataCell(XLS: TXLSReadWriteII5);
var
  i, Row: integer;
  Bookmark: TBookmark;
begin
  //下面导出数据行
  if FDBGridEh.DataSource.DataSet.Active then
  begin
    FDBGridEh.DataSource.DataSet.First;
    Bookmark := FDBGridEh.DataSource.DataSet.GetBookmark;
    Row := FStartRow + FTitleRowCount;
    FDBGridEh.DataSource.DataSet.DisableControls;
    while not FDBGridEh.DataSource.DataSet.Eof do
    begin
      for i := 0 to FColumnCount - 1 do
      begin
        if (FDBGridEh.VisibleColumns[i].Field = nil) or (FDBGridEh.VisibleColumns[i].Field.IsNull) then
          XLS.Sheets[0].AsBlank[i, Row] := true
        else begin
          case FDBGridEh.VisibleColumns[i].Field.DataType of //根据字段类型填写单元格
            ftUnknown: XLS.Sheets[0].AsVariant[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsVariant;
            ftBCD : XLS.Sheets[0].AsFloat[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsFloat;
            ftString: XLS.Sheets[0].AsString[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsString;
            ftInteger, ftLargeint: XLS.Sheets[0].AsInteger[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsInteger;
            ftSmallint: XLS.Sheets[0].AsInteger[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsInteger;
            ftFloat: XLS.Sheets[0].AsFloat[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsFloat;
            ftCurrency:XLS.Sheets[0].AsFloat[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsFloat;
            ftBoolean: XLS.Sheets[0].AsBoolean[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsBoolean;
            ftDateTime: XLS.Sheets[0].AsDateTime[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsDateTime;
            ftVariant: XLS.Sheets[0].AsVariant[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsVariant;
            //{$IF CompilerVersion < 18.5}
            ftWideString: XLS.Sheets[0].AsString[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsString;
            //{$IFEND}
            // Added by cyq 2020-10-16 00:07:25
            ftGuid, ftTimeStamp, ftFMTBcd, // 32..37
            ftFixedWideChar, ftWideMemo, ftOraTimeStamp, ftOraInterval, // 38..41
            ftLongWord, ftShortint, ftByte, ftExtended:
            XLS.Sheets[0].AsString[i, Row] := FDBGridEh.VisibleColumns[i].Field.AsString;
          end;
        end;
        {
        case FDBGridEh.VisibleColumns[i].Alignment of  //设置单元格横向对齐方式
          taLeftJustify: XLS.Sheets[0].Cell[i,Row].HorizAlignment:=chaLeft;
          taRightJustify: XLS.Sheets[0].Cell[i,Row].HorizAlignment:=chaRight;
          taCenter: XLS.Sheets[0].Cell[i,Row].HorizAlignment:=chaCenter;
        end;
        XLS.Sheets[0].Cell[i,Row].NumberFormat:=FDBGridEh.VisibleColumns[i].DisplayFormat;  //设置单元格显示格式
        }
      end;
      //显示进度条进度过程
      if ShowProgress then
      begin
        FGauge.Progress := FDBGridEh.DataSource.DataSet.RecNo;
        FGauge.Refresh;
      end;
      FDBGridEh.DataSource.DataSet.Next;
      inc(Row);
    end;
    FDBGridEh.DataSource.DataSet.GotoBookmark(Bookmark);
    FDBGridEh.DataSource.DataSet.EnableControls;
  end;
end;

procedure TDBGridEhToExcelEx.WriteFooter(XLS: TXLSReadWriteII5); {输出DBGridEh表脚}
var
  i, j, Row: integer;
  fCount: Integer;
begin
  Row := FStartRow + FTitleRowCount + FDBGridEh.DataSource.DataSet.RecordCount {- 1};//有footer时不需要-1
  for i := 0 to FDBGridEh.FooterRowCount - 1 do
  begin
    for j := 0 to FColumnCount - 1 do
    begin
      if i > FDBGridEh.FooterRowCount - 1 then
        XLS.Sheets[0].AsBlank[j, Row] := true
      else begin
        case FDBGridEh.VisibleColumns[j].Footer.ValueType of
          fvtNon: XLS.Sheets[0].AsBlank[j, Row] := true;
          fvtSum, fvtAvg, fvtCount: XLS.Sheets[0].AsVariant[j, Row] := FDBGridEh.VisibleColumns[j].Footer.SumValue;
          fvtFieldValue: XLS.Sheets[0].AsVariant[j, Row] := FDBGridEh.DataSource.DataSet.FieldByName(FDBGridEh.VisibleColumns[j].Footer.FieldName).AsVariant;
          fvtStaticText: XLS.Sheets[0].AsString[j, Row] := FDBGridEh.VisibleColumns[j].Footer.Value;
        end;
        //¸ñÊ½»¯Footer
        case FDBGridEh.VisibleColumns[j].Footer.Alignment of //设置单元格对齐方式
          taLeftJustify: XLS.Sheets[0].Cell[j, Row].HorizAlignment := chaLeft;
          taRightJustify: XLS.Sheets[0].Cell[j, Row].HorizAlignment := chaRight;
          taCenter: XLS.Sheets[0].Cell[j, Row].HorizAlignment := chaCenter;
        end;
        XLS.Sheets[0].Cell[j, Row].NumberFormat := FDBGridEh.VisibleColumns[j].Footer.DisplayFormat; //设置单元格显示
      end;
    end;
    inc(Row);
  end;
end;

procedure TDBGridEhToExcelEx.FormatDataCell(XLS: TXLSReadWriteII5); {格式化}
var
  i, Height, Row: integer;
begin
  Row := FStartRow + FTitleRowCount;
  Height := Row + FDBGridEh.DataSource.DataSet.RecordCount + FDBGridEh.FooterRowCount - 1;

  //小于15000行时才进行格式化
  if FDBGridEh.DataSource.DataSet.RecordCount < 15000 then
  begin
    //cancel by cyq 2021年5月24日 11:45:32
    with XLS.Sheets[0].Range.Items[0, Row, FColumnCount - 1, Height] do
    begin
      BorderOutlineStyle := cbsThin;
      BorderInsideVertStyle := cbsThin;
      BorderInsideHorizStyle := cbsThin;

      VertAlignment := cvaCenter;
      FontCharset := FDBGridEh.Font.Charset;
      FontName := FDBGridEh.Font.Name;
      FontSize := FDBGridEh.Font.Size;

      FontColor := FDBGridEh.Font.Color;
    end;
  end;
    //下面代码设置数据行行高
    //XLS.Sheets[0].Rows.AddIfNone(Row,Height);
  for i := Row to Height do
    XLS.Sheets[0].Rows[i].PixelHeight := GetRowHeight;
end;

procedure TDBGridEhToExcelEx.CreateProcessForm(AOwner: TComponent);
var
  Panel: TPanel;
  Prompt: TLabel; {提示的标签}
begin
  if Assigned(FProgressForm) then
    exit;
  FProgressForm := TForm.Create(AOwner);
  with FProgressForm do
  begin
    try
      Font.Name := '宋体'; {设置字体}
      Font.Size := 9;
      BorderStyle := bsNone;
      Width := 350;
      Height := 100;
      BorderWidth := 1;
      Color := clWhite;
      Position := poScreenCenter;
      Panel := TPanel.Create(FProgressForm);
      with Panel do
      begin
        Parent := FProgressForm;
        Color := clWhite;
        Align := alClient;
        BevelInner := bvNone;
        //BevelOuter := bvRaised;
        BevelOuter := bvNone;
        BevelKind := bkFlat;
        Caption := '';
      end;
      Prompt := TLabel.Create(Panel);
      with Prompt do
      begin
        Parent := Panel;
        AutoSize := True;
        Left := 20;
        Top := 25;
        Caption := ' 正在导出数据，请稍候......';
        Font.Style := [fsBold];
        Font.Color := $00CA8715;
      end;
      FGauge := TGauge.Create(Panel);
      with FGauge do
      begin
        Parent := Panel;
        BorderStyle := bsNone;
        ForeColor := $004F9D00;
        Left := 20;
        Top := 45;
        Height := 20;
        Width := 320;
        MinValue := 0;
        MaxValue := FDBGridEh.DataSource.DataSet.RecordCount;
      end;
    except
    end;
  end;
  FProgressForm.Show;
  FProgressForm.Update;
end;

procedure TDBGridEhToExcelEx.SetShowProgress(const Value: Boolean);
begin
  FShowProgress := Value;
end;

procedure TDBGridEhToExcelEx.ExportToExcel(OpenFile: Boolean);
var
  XLS: TXLSReadWriteII5;
  Msg: string;
begin
  //如果数据集为空或没有打开则退出
  if (FDBGridEh.DataSource.DataSet.IsEmpty) or (not FDBGridEh.DataSource.DataSet.Active) then exit;
  //如果保存的文件名为空则退出
  if Trim(FileName) = '' then exit;
  Screen.Cursor := crHourGlass;
  try
    try
      if FileExists(FileName) then
      begin
        Msg := '文件（' + FileName + ' ）已经存在，是否覆盖？';
        if Application.MessageBox(PChar(Msg), '提示信息', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) = IDYES then
          DeleteFile(FileName) //删除文件
        else
          exit;
      end;
      if ShowProgress then
        CreateProcessForm(nil); //显示进度窗体

      XLS := TXLSReadWriteII5.Create(nil);
      try
        XLS.Filename := FileName;
        WriteGridHeader(XLS);
        WriteTitleData(XLS);
        WriteDataCell(XLS); {输出数据集内容}
        WriteFooter(XLS);
        FormatDataCell(XLS);
        XLS.Write;
      finally
        XLS.Free;
      end;
      if OpenFile then
        ShellExecute(0, 'Open', PChar(FileName), nil, nil, SW_SHOW); //打开Excel文件
    except
    end;
  finally
    if ShowProgress then FreeAndNil(FProgressForm);
    Screen.Cursor := crDefault;
  end;
end;

destructor TDBGridEhToExcelEx.Destroy;
begin
  FGridHeader.Free;
  inherited Destroy;
end;

end.


