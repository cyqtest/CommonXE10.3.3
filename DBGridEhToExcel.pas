unit DBGridEhToExcel;

interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ComCtrls, ExtCtrls, StdCtrls, Gauges, DBGridEh, ShellApi;

type
  TTitleCell = array of array of String;

  //分解DBGridEh的标题
  TDBGridEhTitle = class
  private
    FDBGridEh: TDBGridEh;  //对应DBGridEh
    FColumnCount: integer; //DBGridEh列数(指visible为True的列数)
    FRowCount: integer;    //DBGridEh多表头层数(没有多表头则层数为1)
    procedure SetDBGridEh(const Value: TDBGridEh);
    function GetTitleRow: integer;    //获取DBGridEh多表头层数
    function GetTitleColumn: integer; //获取DBGridEh列数
  public
    //分解DBGridEh标题，由TitleCell二维动态数组返回
    procedure GetTitleData(var TitleCell: TTitleCell);
  published
    property DBGridEh: TDBGridEh read FDBGridEh write SetDBGridEh;
    property ColumnCount: integer read FColumnCount;
    property RowCount: integer read FRowCount;
  end;

  TDBGridEhToExcel = class(TComponent)
  private
    FCol: integer;
    FRow: integer;
    FProgressForm: TForm;                                  {进度窗体}
    FGauge: TGauge;                                        {进度条}
    Stream: TStream;                                       {输出文件流}
    FBookMark: TBookmark;                                  
    FShowProgress: Boolean;                                {是否显示进度窗体}
    FDBGridEh: TDBGridEh;
    FBeginDate: TCaption;                                  {开始日期}
    FTitleName: TCaption;                                  {Excel文件标题}
    FEndDate: TCaption;                                    {结束日期}
    FUserName: TCaption;                                   {制表人}
    FFileName: String;                                     {保存文件名}
    procedure SetShowProgress(const Value: Boolean);
    procedure SetDBGridEh(const Value: TDBGridEh);
    procedure SetBeginDate(const Value: TCaption);
    procedure SetEndDate(const Value: TCaption);
    procedure SetTitleName(const Value: TCaption);
    procedure SetUserName(const Value: TCaption);
    procedure SetFileName(const Value: String);   

    procedure IncColRow;
    procedure WriteBlankCell;                              {写空单元格}
    {写数字单元格}
    procedure WriteFloatCell(const AValue: Double; const IncStatus: Boolean=True);
    {写整型单元格}
    procedure WriteIntegerCell(const AValue: Integer; const IncStatus: Boolean=True);
    {写字符单元格}
    procedure WriteStringCell(const AValue: string; const IncStatus: Boolean=True);
    procedure WritePrefix;
    procedure WriteSuffix;
    procedure WriteHeader;                                 {输出Excel标题}
    procedure WriteTitle;                                  {输出Excel列标题}
    procedure WriteDataCell;                               {输出数据集内容}
    procedure WriteFooter;                                 {输出DBGridEh表脚}
    procedure SaveStream(aStream: TStream);
    procedure CreateProcessForm(AOwner: TComponent);       {生成进度窗体}
    {根据表格修改数据集字段顺序及字段中文标题}
    procedure SetDataSetCrossIndexDBGridEh;
    {获取表格的字段数}
    function GetDBGridColumnCount(IsIncludeInvisibleColumn: Boolean): Integer;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ExportToExcel; {输出Excel文件}
  published
    property DBGridEh: TDBGridEh read FDBGridEh write SetDBGridEh;
    property ShowProgress: Boolean read FShowProgress write SetShowProgress;
    property TitleName: TCaption read FTitleName write SetTitleName;
    property BeginDate: TCaption read FBeginDate write SetBeginDate;
    property EndDate: TCaption read FEndDate write SetEndDate;
    property UserName: TCaption read FUserName write SetUserName;
    property FileName: String read FFileName write SetFileName;
  end;

var
  CXlsBof: array[0..5] of Word = ($809, 8, 0, $10, 0, 0);
  CXlsEof: array[0..1] of Word = ($0A, 00);
  CXlsLabel: array[0..5] of Word = ($204, 0, 0, 0, 0, 0);
  CXlsNumber: array[0..4] of Word = ($203, 14, 0, 0, 0);
  CXlsRk: array[0..4] of Word = ($27E, 10, 0, 0, 0);
  CXlsBlank: array[0..4] of Word = ($201, 6, 0, 0, $17);

implementation
{ TDBGridEhTitle }

function TDBGridEhTitle.GetTitleColumn: integer;
var
  i, ColumnCount: integer;
begin
  ColumnCount := 0;
  for i := 0 to DBGridEh.Columns.Count - 1 do
  begin
    if DBGridEh.Columns[i].Visible then
      Inc(ColumnCount);
  end;

  Result := ColumnCount;
end;

procedure TDBGridEhTitle.GetTitleData(var TitleCell: TTitleCell);
var
  i, Row, Col: integer;
  Caption: String;
begin
  FColumnCount := GetTitleColumn;
  FRowCount := GetTitleRow;
  SetLength(TitleCell,FColumnCount,FRowCount);
  Row := 0;
  for i := 0 to DBGridEh.Columns.Count - 1 do
  begin
    if DBGridEh.Columns[i].Visible then
    begin
      Col := 0;
      Caption := DBGridEh.Columns[i].Title.Caption;
      while POS('|', Caption) > 0 do
      begin
        TitleCell[Row,Col] := Copy(Caption, 1, Pos('|',Caption)-1);
        Caption := Copy(Caption,Pos('|', Caption)+1, Length(Caption));
        Inc(Col);
      end;
      TitleCell[Row, Col] := Caption;
      Inc(Row);
    end;
  end;
end;

function TDBGridEhTitle.GetTitleRow: integer;
var
  i, j: integer;
  MaxRow, Row: integer;
begin
  MaxRow := 1;
  for i := 0 to DBGridEh.Columns.Count - 1 do
  begin
    Row := 1;
    for j := 0 to Length(DBGridEh.Columns[i].Title.Caption) do
    begin
      if DBGridEh.Columns[i].Title.Caption[j] = '|' then
        Inc(Row);
    end;

    if MaxRow < Row then
      MaxRow :=  Row;
  end;

  Result := MaxRow;
end;

procedure TDBGridEhTitle.SetDBGridEh(const Value: TDBGridEh);
begin
  FDBGridEh := Value;
end;

{ TDBGridEhToExcel }

constructor TDBGridEhToExcel.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FShowProgress := True;
end;

procedure TDBGridEhToExcel.SetShowProgress(const Value: Boolean);
begin
  FShowProgress := Value;
end;

procedure TDBGridEhToExcel.SetDBGridEh(const Value: TDBGridEh);
begin
  FDBGridEh := Value;
end;

procedure TDBGridEhToExcel.SetBeginDate(const Value: TCaption);
begin
  FBeginDate := Value;
end;

procedure TDBGridEhToExcel.SetEndDate(const Value: TCaption);
begin
  FEndDate := Value;
end;

procedure TDBGridEhToExcel.SetTitleName(const Value: TCaption);
begin
  FTitleName := Value;
end;

procedure TDBGridEhToExcel.SetUserName(const Value: TCaption);
begin
  FUserName := Value;
end;

procedure TDBGridEhToExcel.SetFileName(const Value: String);
begin
  FFileName := Value;
end;

procedure TDBGridEhToExcel.IncColRow;
begin
  //if FCol = DBGridEh.DataSource.DataSet.FieldCount - 1 then
  if FCol = GetDBGridColumnCount(False)-1 then
  begin
    Inc(FRow);
    FCol := 0;
  end
  else
    Inc(FCol);
end;

procedure TDBGridEhToExcel.WriteBlankCell;
begin
  CXlsBlank[2] := FRow;
  CXlsBlank[3] := FCol;
  Stream.WriteBuffer(CXlsBlank, SizeOf(CXlsBlank));
  IncColRow;
end;

procedure TDBGridEhToExcel.WriteFloatCell(const AValue: Double; const IncStatus: Boolean=True);
begin
  CXlsNumber[2] := FRow;
  CXlsNumber[3] := FCol;
  Stream.WriteBuffer(CXlsNumber, SizeOf(CXlsNumber));
  Stream.WriteBuffer(AValue, 8);

  if IncStatus then
    IncColRow;
end;

procedure TDBGridEhToExcel.WriteIntegerCell(const AValue: Integer; const IncStatus: Boolean=True);
var
  V: Integer;
begin
  CXlsRk[2] := FRow;
  CXlsRk[3] := FCol;
  Stream.WriteBuffer(CXlsRk, SizeOf(CXlsRk));
  V := (AValue Shl 2) Or 2;
  Stream.WriteBuffer(V, 4);

  if IncStatus then
    IncColRow;
end;

procedure TDBGridEhToExcel.WriteStringCell(const AValue: string; const IncStatus: Boolean=True);
var
  L: integer;
begin
  L := Length(AValue);
  CXlsLabel[1] := 8 + L;
  CXlsLabel[2] := FRow;
  CXlsLabel[3] := FCol;
  CXlsLabel[5] := L;
  Stream.WriteBuffer(CXlsLabel, SizeOf(CXlsLabel));
  Stream.WriteBuffer(Pointer(AValue)^, L);

  if IncStatus then
    IncColRow;
end;

procedure TDBGridEhToExcel.WritePrefix;
begin
  Stream.WriteBuffer(CXlsBof, SizeOf(CXlsBof));
end;

procedure TDBGridEhToExcel.WriteSuffix;
begin
  Stream.WriteBuffer(CXlsEof, SizeOf(CXlsEof));
end;

procedure TDBGridEhToExcel.WriteHeader;
var
  OpName, OpDate: String; 
begin
  //标题
  FCol := 3;
  WriteStringCell(TitleName,False);
  FCol := 0;

  Inc(FRow);

  if Trim(BeginDate) <> '' then
  begin
    //开始日期
    FCol := 0;
    WriteStringCell(BeginDate,False);
    FCol := 0
  end;

  if Trim(EndDate) <> '' then
  begin
    //结束日期
    FCol := 5;
    WriteStringCell(EndDate,False);
    FCol := 0;
  end;

  if (Trim(BeginDate) <> '') or (Trim(EndDate) <> '') then
    Inc(FRow);

  //制表人
  OpName := '制表人：' + UserName;
  FCol := 0;
  WriteStringCell(OpName,False);
  FCol := 0;

  //制表时间
  OpDate := '制表时间：' + DateTimeToStr(Now);
  FCol := 5;
  WriteStringCell(OpDate,False);
  FCol := 0;

  Inc(FRow);  
end;

procedure TDBGridEhToExcel.WriteTitle;
var
  i, j: integer;
  DBGridEhTitle: TDBGridEhTitle;
  TitleCell: TTitleCell;
begin
  DBGridEhTitle := TDBGridEhTitle.Create;
  try
    DBGridEhTitle.DBGridEh := FDBGridEh;
    DBGridEhTitle.GetTitleData(TitleCell);

    try
      for i := 0 to DBGridEhTitle.RowCount - 1 do
      begin
        for j := 0 to DBGridEhTitle.ColumnCount - 1 do
        begin
          FCol := j;
          WriteStringCell(TitleCell[j,i],False);
        end;
        Inc(FRow);
      end;
      FCol := 0;
    except

    end;
  finally
    DBGridEhTitle.Free;
  end;
end;

procedure TDBGridEhToExcel.WriteDataCell;
var
  i: integer;
begin
  DBGridEh.DataSource.DataSet.DisableControls;
  FBookMark := DBGridEh.DataSource.DataSet.GetBookmark;
  try
    DBGridEh.DataSource.DataSet.First;
    while not DBGridEh.DataSource.DataSet.Eof do
    begin
//      for i := 0 to DBGridEh.DataSource.DataSet.FieldCount - 1 do
//      begin
//        if DBGridEh.DataSource.DataSet.Fields[i].IsNull or (not DBGridEh.DataSource.DataSet.Fields[i].Visible) then
//          WriteBlankCell
//        else
//        begin
//          case DBGridEh.DataSource.DataSet.Fields[i].DataType of
//            ftSmallint, ftInteger, ftWord, ftAutoInc, ftBytes:
//              WriteIntegerCell(DBGridEh.DataSource.DataSet.Fields[i].AsInteger);
//            ftFloat, ftCurrency, ftBCD:
//              WriteFloatCell(DBGridEh.DataSource.DataSet.Fields[i].AsFloat);
//          else
//            if DBGridEh.DataSource.DataSet.Fields[i] Is TBlobfield then  // 此类型的字段(图像等)暂无法读取显示
//              WriteStringCell('')
//            else
//              WriteStringCell(DBGridEh.DataSource.DataSet.Fields[i].AsString);
//          end;
//        end;
//      end;

      for i := 0 to DBGridEh.Columns.Count - 1 do
      begin
        if not DBGridEh.Columns.Items[i].Visible then
          Continue;
        if DBGridEh.Columns.Items[i].Field.IsNull then
          WriteBlankCell
        else
        begin
          case DBGridEh.Columns.Items[i].Field.DataType of
            ftSmallint, ftInteger, ftWord, ftAutoInc, ftBytes:
              begin
                if DBGridEh.Columns.Items[i].PickList.Count=0 then
                  WriteIntegerCell(DBGridEh.DataSource.DataSet.fieldbyname(DBGridEh.Columns.Items[i].FieldName).AsInteger)
                else
                begin
                  WriteStringCell(DBGridEh.Columns.Items[i].PickList.Strings[DBGridEh.DataSource.DataSet.fieldbyname(DBGridEh.Columns.Items[i].FieldName).AsInteger-1]);
                end;
              end;
            ftFloat, ftCurrency, ftBCD:
              begin
                if DBGridEh.Columns.Items[i].PickList.Count=0 then
                  WriteFloatCell(DBGridEh.DataSource.DataSet.fieldbyname(DBGridEh.Columns.Items[i].FieldName).AsFloat)
                else
                begin
                  WriteStringCell(DBGridEh.Columns.Items[i].PickList.Strings[DBGridEh.DataSource.DataSet.fieldbyname(DBGridEh.Columns.Items[i].FieldName).AsInteger-1]);
                end;
              end;
          else
            if DBGridEh.Columns.Items[i].Field Is TBlobfield then  // 此类型的字段(图像等)暂无法读取显示
              WriteStringCell('')
            else
            begin
              if DBGridEh.Columns.Items[i].PickList.Count=0 then
                WriteStringCell(DBGridEh.DataSource.DataSet.fieldbyname(DBGridEh.Columns.Items[i].FieldName).AsString)
              else
              begin
                WriteStringCell(DBGridEh.Columns.Items[i].PickList.Strings[DBGridEh.DataSource.DataSet.fieldbyname(DBGridEh.Columns.Items[i].FieldName).AsInteger-1]);
              end;
            end;
          end;
        end;
      end;

      //显示进度条进度过程
      if ShowProgress then
      begin
        FGauge.Progress := DBGridEh.DataSource.DataSet.RecNo;
        FGauge.Refresh;
      end;

      DBGridEh.DataSource.DataSet.Next;
    end;

  finally
    if DBGridEh.DataSource.DataSet.BookmarkValid(FBookMark) then
    DBGridEh.DataSource.DataSet.GotoBookmark(FBookMark);

    DBGridEh.DataSource.DataSet.EnableControls;
  end;
end;

procedure TDBGridEhToExcel.WriteFooter;
var
  i, j: integer;
begin
  if DBGridEh.FooterRowCount = 0 then exit;

  FCol := 0;
  if DBGridEh.FooterRowCount = 1 then
  begin
    for i := 0 to DBGridEh.Columns.Count - 1 do
    begin
      if DBGridEh.Columns[i].Visible then
      begin
        WriteStringCell(DBGridEh.Columns[i].Footer.Value,False);
        Inc(FCol);
      end;
    end;
  end
  else if DBGridEh.FooterRowCount > 1 then
  begin
    for i := 0 to DBGridEh.Columns.Count - 1 do
    begin
      if DBGridEh.Columns[i].Visible then
      begin
        for j := 0 to DBGridEh.Columns[i].Footers.Count - 1 do
        begin
          WriteStringCell(DBGridEh.Columns[i].Footers[j].Value ,False);
          Inc(FRow);
        end;
        Inc(FCol);
        FRow := FRow - DBGridEh.Columns[i].Footers.Count;
      end;
    end;
  end;
  FCol := 0;
end;

procedure TDBGridEhToExcel.SaveStream(aStream: TStream);
begin
  FCol := 0;
  FRow := 0;
  Stream := aStream;

  //输出前缀
  WritePrefix;

  //输出表格标题
  //WriteHeader;

  //输出列标题
  WriteTitle;

  //输出数据集内容
  WriteDataCell;

  //输出DBGridEh表脚
  WriteFooter;

  //输出后缀
  WriteSuffix;
end;

procedure TDBGridEhToExcel.ExportToExcel;
var
  FileStream: TFileStream;
  Msg: String;
  SaveDialog: TSaveDialog;
begin
  //如果数据集为空或没有打开则退出
  if (DBGridEh.DataSource.DataSet.IsEmpty) or (not DBGridEh.DataSource.DataSet.Active) then
    exit;

  SaveDialog := TSaveDialog.Create(nil);
  try
    SaveDialog.Filter := 'Excel files (*.xls)|*.xls|All files (*.*)|*.*';
    SaveDialog.Filename := TitleName+'.xls';
    if SaveDialog.Execute then
    begin
      FileName := SaveDialog.Filename;
      if FileName='' then
        FileName:='转Excel';
    end else begin
      DBGridEh.DataSource.DataSet.EnableControls;
      Exit;
    end;
  finally
    SaveDialog.Free;
  end;
    
  //根据表格修改数据集字段顺序及字段中文标题
  SetDataSetCrossIndexDBGridEh;

  Screen.Cursor := crHourGlass;
  try
    try
      if FileExists(FileName) then
      begin
        Msg := '已存在文件（' + FileName + '），是否覆盖？';
        if Application.MessageBox(PChar(Msg),'提示',MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON2) = IDYES then
        begin
          //删除文件
          DeleteFile(FileName)
        end
        else
          exit;
      end;

      //显示进度窗体
      if ShowProgress then
        CreateProcessForm(nil);
        
      FileStream := TFileStream.Create(FileName, fmCreate);
      try
        //输出文件
        SaveStream(FileStream);
      finally
        FileStream.Free;
      end;
      
      //打开Excel文件
      ShellExecute(0, 'Open', PChar(FileName), nil, nil, SW_SHOW);
    except

    end;
  finally
    if ShowProgress then
      FreeAndNil(FProgressForm);
    Screen.Cursor := crDefault;
  end;
end;

destructor TDBGridEhToExcel.Destroy;
begin
  inherited Destroy;
end;

procedure TDBGridEhToExcel.CreateProcessForm(AOwner: TComponent);
var
  Panel: TPanel;
  Prompt: TLabel;                                           {提示的标签}
begin
  if Assigned(FProgressForm) then
    exit;

  FProgressForm := TForm.Create(AOwner);
  with FProgressForm do
  begin
    try
      Font.Name := '宋体';                                  {设置字体}
      Font.Size := 9;
      BorderStyle := bsNone;
      Width := 300;
      Height := 100;
      BorderWidth := 1;
      Color := clBlack;
      Position := poScreenCenter;

      Panel := TPanel.Create(FProgressForm);
      with Panel do
      begin
        Parent := FProgressForm;
        Align := alClient;
        BevelInner := bvNone;
        BevelOuter := bvRaised;
        Caption := '';
      end;

      Prompt := TLabel.Create(Panel);
      with Prompt do
      begin
        Parent := Panel;
        AutoSize := True;
        Left := 25;
        Top := 25;
        Caption := '正在导出数据，请稍候......';
        Font.Style := [fsBold];
      end;

      FGauge := TGauge.Create(Panel);
      with FGauge do
      begin
        Parent := Panel;
        ForeColor := clBlue;
        Left := 20;
        Top := 50;
        Height := 13;
        Width := 260;
        MinValue := 0;
        MaxValue := DBGridEh.DataSource.DataSet.RecordCount;
      end;
    except

    end;
  end;

  FProgressForm.Show;
  FProgressForm.Update;
end;

procedure TDBGridEhToExcel.SetDataSetCrossIndexDBGridEh;
var
  i: integer;
begin
  for i := 0 to DBGridEh.Columns.Count - 1 do
  begin
    DBGridEh.DataSource.DataSet.FieldByName(DBGridEh.Columns.Items[i].FieldName).Index := i;
    DBGridEh.DataSource.DataSet.FieldByName(DBGridEh.Columns.Items[i].FieldName).DisplayLabel
      := DBGridEh.Columns.Items[i].Title.Caption;
    DBGridEh.DataSource.DataSet.FieldByName(DBGridEh.Columns.Items[i].FieldName).Visible :=
      DBGridEh.Columns.Items[i].Visible;
  end;

  for i := 0 to DBGridEh.DataSource.DataSet.FieldCount - 1 do
  begin
    if POS('*****',DBGridEh.DataSource.DataSet.Fields[i].DisplayLabel) > 0 then
      DBGridEh.DataSource.DataSet.Fields[i].Visible := False;
  end;
end;

function TDBGridEhToExcel.GetDBGridColumnCount(
  IsIncludeInvisibleColumn: Boolean): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i:=0 to DBGridEh.Columns.Count-1 do
  begin
    if IsIncludeInvisibleColumn then
    begin
      inc(Result);
    end
    else
    begin
      if DBGridEh.Columns.Items[i].Visible then
        inc(Result);
    end;
  end;
end;

end.
