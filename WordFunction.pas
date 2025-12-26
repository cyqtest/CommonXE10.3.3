{*********************************************************************}
{                                                                     }
{     WordFunction  //Modified by  By CYQ                             }
{                                                                     }
{                                                                     }
{     修改以前的Word操作单元                                          }
{                                                                     }
{*********************************************************************}
unit WordFunction;

interface

uses
  Windows, Forms, Classes, DB, Comobj, ExtCtrls, Chart, DBGridEh, OleServer, Word2000,
  Clipbrd, SysUtils, ComCtrls, DBGrids,
  Controls, JExcelProc, TypInfo, Grids, Dialogs;

//开始操作Word
procedure StartOperationWord(bHide: Boolean = True);  //bHide是否显示Word的启动界面
//结束Word操作，释放变量
procedure EndOperationWord(bSave: Boolean = True);   //bSave是否自动保存文档


//新建Word
function AddNewDocument(AFileName: String = ''): Boolean;
//添加标题
procedure AddTitleText(TitleText: String; FontSize: Integer = 16; FontName: String = '宋体';
      //0居中  1左对齐  2右对齐
      Alignment: Integer = 0; bBold: Boolean = True);


//打开模版
function CopyAndOpenWordModel(ScrFileName, DestFileName: String; bOpen: Boolean = True): Boolean;
//复制文件
function CopyWordFile(ScrFile, DestFile: String; bDelete: Boolean = True): Boolean;
//替换文档中的字符
function WordReplaceText(OldText, NewText: String): Boolean;
//删除指定区域Range中的字符数
procedure DeleteBlank(n: Integer);
//在指定标签后插入文本，标签为空时，在文档的最后面插入
function WordInsertText(txt: String; bFormat: Boolean; sBookMark: String = ''; bNewLine: Boolean = True): Boolean;
//在指定标签后插入指定列、行的表格，标签为空时，在文档的最后面插入
function WordAddTable(Col, Row: Integer; sBookMark: String = ''; bNewLine: Boolean = True): Boolean;
//设置表格的边框显示与隐藏 True为显示，False为隐藏
procedure SetTableBorderVisible(wTable: Variant; B: Boolean);
{DataSet}
//==============================================================================
{DBGrid}
//导出 TDataSet 的 FieldList 的字段内容到 xList 中，并把需要导出的列数保存到 Col 中。
//procedure DataSetToList(qry: TDataSet; FieldList: Array of String; TitleList: Array of String;
//        var xList: TStringList; var Col: Integer; DBGrid: TDBGrid = nil);

//导出 TDataSet 的 FieldList 的字段内容到 Word 指定标签的表格中。
//例：全部字段：DataSetToWord(Table1, [])
//    指定字段：DataSetToWord(Table1, ['Field1', 'Field2', 'Field3'])
procedure DataSetToWord(qry: TDataSet; FieldList: Array of String; TitleList: Array of String;
        sBookMark: String = ''; DBGrid: TDBGrid = nil; bHint: Boolean = True);
//导出 DBGrid 指定标题列 TitleList 的字段内容到 Word
//例：全部列：DBGridToWord(DBGrid1, [])
//    指定列：DBGridToWord(DBGrid1, ['列1', '列2', '列3']
procedure DBGridToWord(DBGrid: TDBGrid; TitleList: Array of String; sBookMark: String = '');
{DBGridEh}
//导出 TDataSet 的 FieldList 的字段内容到 xList 中，并把需要导出的列数保存到 Col 中。
procedure DataSetEhToList(qry: TDataSet; FieldList: Array of String;
        var xList: TStringLIst; var Col: Integer; DBGridEh: TComponent = nil);
//导出 TDataSet 的 FieldList 的字段内容到 Word 指定标签的表格中。
//例：全部字段：DataSetEhToWord(Table1, [])
//    指定字段：DataSetEhToWord(Table1, ['Field1', 'Field2', 'Field3'])
procedure DataSetEhToWord(qry: TDataSet; FieldList: Array of String;
        DBGridEh: TComponent = nil; sBookMark: String = '');
//导出 DBGridEh 指定标题列 TitleList 的字段内容到 Word
//例：全部列：DBGridEhToWord(DBGridEh1, [])
//    指定列：DBGridEhToWord(DBGridEh1, ['列1', '列2', '列3']
procedure DBGridEhToWord(DBGridEh: TComponent; TitleList: Array of String; sBookMark: String = '');
//==============================================================================
{StringGrid}
//导出 StringGrid 指定列 ColList 的内容到 Word
//例：全部列：StrGridToWordCol(StrGrid1, [])
//    指定列：StrGridToWordCol(StrGrid1, [1, 2, 3]
//procedure StrGridToWordCol(StrGrid: TStringGrid; ColList: Array of Integer;
//        sBookMark: String = '');

//导出 StringGrid 指定标题列 TitleList 的内容到 Excel
//例：全部列：StrGridToWord(StrGrid1, [])
//    指定列：StrGridToWord(StrGrid1, ['列1', '列2', '列3']
//procedure StrGridToWord(StrGrid: TStringGrid; TitleList: Array of String;
//        sBookMark: String = '');
//==============================================================================

//导出表格数据的统一函数(表格包括DBGrid、StringGrid和DBGridEh）
procedure GridToWord(Grid: TWinControl; TitleList: Array of String;
        sBookMark: String = ''; UseTree: Boolean = True);
//把tsList导出到 Word 指定标签的表格中。
procedure WriteListToWord(tsList: TStringList; Col: Integer; sBookMark: String = ''; bHint: Boolean = True);
//==============================================================================
var
  wDoc, wApp: Variant;

implementation

//开始操作Word
procedure StartOperationWord(bHide: Boolean);
begin
    try
        try
            wApp := GetActiveOleObject('Word.Application');
        except
            wApp := CreateOleObject('Word.Application');
        end;
        wApp.Visible := Not bHide;
    except
        Application.MessageBox('未安装Microsoft Word', '错误', MB_ICONERROR);
        Exit;
    end;
end;

//结束Word操作，释放变量
procedure EndOperationWord(bSave: Boolean);
begin
    try
        wDoc.Close(bSave);
        wApp.Quit;
    except
        Application.MessageBox('遇到未知错误，关闭Word失败', '错误', MB_ICONERROR);
        Exit;
    end;
end;

//新建Word
function AddNewDocument(AFileName: String = ''): Boolean;
begin
    Result := False;
    
    try
        wDoc := wApp.Documents.Add();
        if AFileName <> '' then
        begin
            wDoc.SaveAs(AFileName);
        end;
    except
        Application.MessageBox('新建Word文档失败', '错误', MB_ICONERROR);
        Exit;
    end;

    Result := True;
end;

//添加标题 
procedure AddTitleText(TitleText: String; FontSize: Integer; FontName: String; Alignment: Integer; bBold: Boolean);
begin
    //定位到首行的第一个格
    wApp.ActiveDocument.Range(0, 0).InsertAfter(#13);
    wApp.ActiveDocument.Range(0, 0).InsertAfter(TitleText);
    if Alignment = 0 then
        wApp.ActiveDocument.Range.ParagraphFormat.Alignment := wdAlignParagraphCenter
    else
    if Alignment = 1 then
        wApp.ActiveDocument.Range.ParagraphFormat.Alignment := wdAlignParagraphLeft
    else
        wApp.ActiveDocument.Range.ParagraphFormat.Alignment := wdAlignParagraphRight;

    wApp.ActiveDocument.Range.Font.Name := FontName;
    wApp.ActiveDocument.Range.Font.Size := FontSize;
    wApp.ActiveDocument.Range.Bold := bBold;
end;

//打开模版
function CopyAndOpenWordModel(ScrFileName, DestFileName: String; bOpen: Boolean = True): Boolean;
var
  AFileName: String;
  sd: TSaveDialog;
begin
    Result := False;

    sd := TSaveDialog.Create(nil);
    try
        sd.InitialDir := ExtractFilePath(Application.ExeName);
        sd.FileName := DestFileName;
        sd.Filter := 'Word 文件|*.doc';
        sd.DefaultExt := 'doc';
        if Not sd.Execute then Exit;
        if sd.FileName = '' then Exit;
        AFileName := sd.FileName;
    finally
        FreeAndNil(sd);
    end;
    //复制文件
    if Not CopyWordFile(ScrFileName, AFileName) then Exit;
    //打开文件
    if bOpen then
    begin
        try
            wDoc := wApp.Documents.Open(AFileName);
        except
            Application.MessageBox('打开文件失败', '错误', MB_ICONERROR);
            Exit;
        end;
    end;

    Result := True;
end;

//复制文件
function CopyWordFile(ScrFile, DestFile: String; bDelete: Boolean = True): Boolean;
begin
    Result := False;

    if Not FileExists(ScrFile) then
    begin
        Application.MessageBox('源文件不存在，不能复制。', '错误', MB_ICONERROR);
        Exit;
    end;

    if ScrFile = DestFile then
    begin
        Application.MessageBox('源文件和目标文件相同，不能复制。', '错误', MB_ICONERROR);
        Exit;
    end;

    if FileExists(DestFile) then
    begin
        if Not bDelete then
        begin
            Application.MessageBox('目标文件已经存在，不能复制。', '错误', MB_ICONERROR);
            Exit;
        end;
        //if Not FcDeleteFile(PChar(DestFile)) then Exit;
    end;

    if Not CopyFile(PChar(ScrFile), PChar(DestFile), False) then
    begin
        Application.MessageBox('发生未知的错误，复制文件失败。', '错误', MB_ICONERROR);
        Exit;
    end;
    //目标文件去掉只读属性
    FileSetAttr(DestFile, FileGetAttr(DestFile) And Not $00000001);

    Result := True;
end;

//替换文档中的字符
function WordReplaceText(OldText, NewText: String): Boolean;
begin
    Result := False;
    //简单处理，直接执行替换操作
    try
        //清除查找内容和替换内容的格式并进行赋值。
        wApp.Selection.Find.ClearFormatting;
        wApp.Selection.Find.Replacement.ClearFormatting;
        wApp.Selection.Find.Text := OldText;
        wApp.Selection.Find.Replacement.Text := NewText;
        //向下搜索
        wApp.Selection.Find.Forward := True;
        //查找的以后继续查找下一个
        wApp.Selection.Find.Wrap := wdFindContinue;
        //不限定格式
        wApp.Selection.Find.Format := False;
        //不区分大小写
        wApp.Selection.Find.MatchCase := False;
        //全字匹配
        wApp.Selection.Find.MatchWholeWord := True;
        //区分全/半角
        wApp.Selection.Find.MatchByte := True;
        //不使用通配符
        wApp.Selection.Find.MatchWildcards := False;
        wApp.Selection.Find.MatchSoundsLike := False;
        wApp.Selection.Find.MatchAllWordForms := False;

        //关闭拼音查找和语法查找，以便提高程序运行的效率
        wApp.Options.CheckSpellingAsYouType := False;
        wApp.Options.CheckGrammarAsYouType := False;
        
        //执行替换所有的操作
        wApp.Selection.Find.Execute(Replace := wdReplaceAll);
    except
        Application.MessageBox('替换失败', '错误', MB_ICONERROR);
        Exit;
    end;
    Result := True;
end;

//删除指定区域Range中的字符数
procedure DeleteBlank(n: Integer);
var
  i: Integer;
begin
    for i := 1 to n do
        wApp.Selection.TypeBackspace;
end;

//在指定标签后插入文本，标签为空时，在文档的最后面插入
function WordInsertText(txt: String; bFormat: Boolean; sBookMark: String = ''; bNewLine: Boolean = True): Boolean;
var
  wRange: Variant;
  iRangeEnd: Integer;
begin
    Result := False;

    try
        if sBookMark = '' then
        begin
            //在文档末尾
            iRangeEnd := wDoc.Range.End - 1;
            if iRangeEnd < 0 then iRangeEnd := 0;

            wRange:= wDoc.Range(iRangeEnd, iRangeEnd);
        end
        else
        begin
            //在书签处
            try
                //定位书签
                if wDoc.BookMarks.Exists(sBookMark) then
                begin
                    wRange := wDoc.Bookmarks.Item(sBookMark).Range;
                end
                else
                //找不到书签，跳过
                begin
                    Result := True;
                    Exit;
                end;
            except
                Application.MessageBox('出现异常，请与开发人员联系！', '错误', MB_ICONERROR);
                Exit;
            end;
        end;
        //换行插入
        if bNewLine then
            wRange.InsertAfter(#13);

        wRange.InsertAfter(txt);

        //删除文字后文字相应长度的空格
        if bFormat then
        begin
            wRange := wDoc.Range(wRange.End + Length(txt), wRange.End + Length(txt));
            wRange.Select;
            DeleteBlank(Length(txt));
        end;
    except
        Exit;
    end;
    Result := True;
end;

//在指定标签后插入指定列、行的表格，标签为空时，在文档的最后面插入
function WordAddTable(Col, Row: Integer; sBookMark: String = ''; bNewLine: Boolean = True): Boolean;
var
  wRange, wTable: Variant;
  iRangeEnd: Integer;
begin
    Result := False;
    try
        if sBookMark = '' then
        begin
            //在文档末尾
            iRangeEnd := wDoc.Range.End - 1;
            if iRangeEnd < 0 then iRangeEnd := 0;

            wRange:= wDoc.Range(iRangeEnd, iRangeEnd);
        end
        else
        begin
            //在书签处
            try
                //定位书签
                if wDoc.BookMarks.Exists(sBookMark) then
                begin
                    wRange := wDoc.Bookmarks.Item(sBookMark).Range;
                end
                else
                //找不到书签，跳过
                begin
                    Result := True;
                    Exit;
                end;
            except
                Application.MessageBox('出现异常，请与开发人员联系！', '错误', MB_ICONERROR);
                Exit;
            end;
        end;
        //换行插入
        if bNewLine then
            wRange.InsertAfter(#13);
        //插入表格
        wTable := wDoc.Tables.Add(wRange, Row, Col);
        //设置表格边框显示
        SetTableBorderVisible(wTable, True);
        //改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。
        wTable.Columns.AutoFit;
    except
        Exit;
    end;
    Result := True;
end;

//设置表格的边框显示与隐藏 True为显示，False为隐藏
procedure SetTableBorderVisible(wTable: Variant; B: Boolean);
begin
    if B then
    begin
        wTable.Borders.Item(wdborderLeft).LineStyle := wdLineStyleSingle;
        wTable.Borders.Item(wdBorderRight).LineStyle := wdLineStyleSingle;
        wTable.Borders.Item(wdBorderTop).LineStyle := wdLineStyleSingle;
        wTable.Borders.Item(wdBorderBottom).LineStyle := wdLineStyleSingle;
        wTable.Borders.Item(wdBorderHorizontal).LineStyle := wdLineStyleSingle;
        wTable.Borders.Item(wdBorderVertical).LineStyle := wdLineStyleSingle;
        wTable.Borders.Shadow := False;
    end
    else
    begin
        wTable.Borders.Item(wdborderLeft).LineStyle := wdLineStyleNone;
        wTable.Borders.Item(wdBorderRight).LineStyle := wdLineStyleNone;
        wTable.Borders.Item(wdBorderTop).LineStyle := wdLineStyleNone;
        wTable.Borders.Item(wdBorderBottom).LineStyle := wdLineStyleNone;
        wTable.Borders.Item(wdBorderHorizontal).LineStyle := wdLineStyleNone;
        wTable.Borders.Item(wdBorderVertical).LineStyle := wdLineStyleNone;
        wTable.Borders.Shadow := False;
    end;
end;

{DataSet}
//==============================================================================
{DBGrid}
//导出 TDataSet 的 FieldList 的字段内容到 xList 中，并把需要导出的列数保存到 Col 中。
//procedure DataSetToList(qry: TDataSet; FieldList: Array of String; TitleList: Array of String;
//        var xList: TStringList; var Col: Integer; DBGrid: TDBGrid = nil);
//var
//  i, j: Integer;
//  FList, FTitleList: TStringList; //FieldName List
//  s, sTemp: String;
//  BM: TBookMark;
//  FTreeCount: Integer;
//  FTreeValues: TStringList;
//  bNotGetTree, bUseMultiTitle: Boolean;
//  eTemp: Extended;
//  Cur: TCursor;
//  pCol: TColumn;
//begin
//    FTreeCount := 0;
//    bNotGetTree := True;
//    Cur := Screen.Cursor;
//    Screen.Cursor := crHourGlass;
//    FTreeValues := TStringList.Create;
//    try
//        FList := TStringList.Create;
//        try
//            FTitleList := TStringList.Create;
//            try
//                s := '';
//                bUseMultiTitle := (DBGrid <> nil) And (DBGrid is TCustomDBGrid)
//                    And (DBGrid as TCustomDBGrid).UseMultiTitle;
//                //如果没有传入导出字段，则读取所有字段
//                if Length(FieldList) = 0 then
//                begin
//                    //如果存在DBGrid，则仅读取DBGrid的列所关联的字段，以及导出标题内容 s
//                    if DBGrid <> nil then
//                    begin
//                        for i := 0 to DBGrid.Columns.Count - 1 do
//                        begin
//                            if DBGrid.Columns[i].FieldName <> '' then
//                            begin
//                                FList.Add(DBGrid.Columns[i].FieldName);
//                                //处理树状多层表头
//                                sTemp := DBGrid.Columns[i].Title.Caption;
//                                pCol := DBGrid.Columns[i].ParentColumn;
//                                while pCol <> nil do
//                                begin
//                                    sTemp := pCol.Title.Caption + '|' + sTemp;
//                                    pCol := pCol.ParentColumn;
//                                    bUseMultiTitle := True;
//                                end;
//                                FTitleList.Add(sTemp);
//                                //FTitleList.Add(DBGrid.Columns[i].Title.Caption);
//                                //s := s + DBGrid.Columns[i].Title.Caption + #9;
//                            end; //if
//                        end; //for
//                        if DBGrid is TJCustomDBGrid then
//                            FTreeCount := (DBGrid as TJCustomDBGrid).TreeLayerCount;    //得到表格树状层数
//                    end
//                    else //否则，读取所有字段
//                    begin
//                        for i := 0 to qry.FieldCount - 1 do
//                            FList.Add(qry.Fields[i].FieldName);
//                    end;
//                end
//                else //传入了导出字段
//                begin
//                    //如果存在DBGrid，则仅读取要求导出并且DBGrid的列所关联的字段，以及导出标题内容 s
//                    if DBGrid <> nil then
//                    begin
//                        if DBGrid is TJCustomDBGrid then
//                            FTreeCount := (DBGrid as TJCustomDBGrid).TreeLayerCount;    //得到表格树状层数
//
//                        for i := 0 to Length(FieldList) - 1 do
//                        begin
//                            for j := 0 to DBGrid.Columns.Count - 1 do
//                            begin
//                                if (DBGrid.Columns[j].FieldName <> '')
//                                    And (CompareText(DBGrid.Columns[j].FieldName, FieldList[i]) = 0) then
//                                begin
//                                    FList.Add(FieldList[i]);
//                                    //处理树状多层表头
//                                    sTemp := DBGrid.Columns[j].Title.Caption;
//                                    pCol := DBGrid.Columns[j].ParentColumn;
//                                    while pCol <> nil do
//                                    begin
//                                        sTemp := pCol.Title.Caption + '|' + sTemp;
//                                        pCol := pCol.ParentColumn;
//                                        bUseMultiTitle := True;
//                                    end;
//                                    FTitleList.Add(sTemp);
//                                    //FTitleList.Add(DBGrid.Columns[j].Title.Caption);
//                                    //s := s + DBGrid.Columns[j].Title.Caption + #9;
//                                    if (j >= FTreeCount - 1) And bNotGetTree then
//                                    begin
//                                        //得到指定字段的树状层数
//                                        if j > FTreeCount - 1 then
//                                            FTreeCount := FList.Count - 1
//                                        else
//                                            FTreeCount := FList.Count;
//                                        bNotGetTree := False;
//                                    end;
//                                    break;
//                                end; //if
//                            end; //for
//                        end; //for
//                    end
//                    else  //否则，读取所有导出字段
//                    begin
//                        for i := 0 to Length(FieldList) - 1 do
//                            FList.Add(FieldList[i]);
//
//                        //如果传入标题，则按照传入的标题显示
//                        if Length(TitleList) > 0 then
//                        begin
//                            for i := 0 to Length(TitleList) - 1 do
//                            begin
//                                FTitleList.Add(TitleList[i]);
//                            end;
//                        end;
//                    end;
//                end;
//                s := GetExcelTitleStr(FTitleList, bUseMultiTitle);
//                if s <> '' then
//                begin
//                    //s := Copy(s, 1, Length(s) - Length(#9));
//                    xList.Add(s);  //导出标题内容 s
//                end;
//            finally
//                FreeAndNil(FTitleList);
//            end;
//
//            //获取数据集的列数
//            Col := FList.Count;
//
//            //读取数据内容
//            BM := qry.GetBookmark;
//            try
//                qry.DisableControls;
//                //qry.DisableConstraints;
//                qry.First;
//                while not qry.Eof do
//                begin
//                    s := '';
//                    for i := 0 to FList.Count - 1 do
//                    begin
//                        sTemp := ValidExcelCell(qry.FieldByName(FList[i]).DisplayText);
//
//                        //处理树状层次显示
//                        if i < FTreeCount then
//                        begin
//                            if i >= FTreeValues.Count then
//                                FTreeValues.Add(sTemp)
//                            else
//                            if FTreeValues[i] = sTemp then
//                                sTemp := ''
//                            else
//                            begin
//                                for j := FTreeValues.Count - 1 downto i + 1 do
//                                    FTreeValues.Delete(j);
//                                FTreeValues[i] := sTemp;
//                            end;
//                        end;
//
//                       { //防止字符串型的数值, 导入Excel后变为数值型丢失前面的零(如: '001' ===> 1).
//                        if (sTemp <> '') And (qry.FieldByName(FList[i]).DataType = ftString)
//                                And TryStrToFloat(sTemp, eTemp) then
//                        begin
//                            if (FloatToStr(eTemp) <> sTemp) or (eTemp > 1E14) then //如果不等或eTemp超大(如身份证), 才转换
//                                sTemp := '=Trim("' + sTemp + '")';
//                        end;   }
//
//                        s := s + sTemp + #9;
//                        Application.ProcessMessages;
//                    end;
//                    if StringReplace(s, #9, '', [rfReplaceAll]) <> '' then
//                    begin
//                        s := Copy(s, 1, Length(s) - Length(#9));
//                        xList.Add(s);
//                    end;
//                    qry.Next;
//                end;
//            finally
//                qry.GotoBookmark(BM);
//                qry.FreeBookmark(BM);
//                //qry.EnableConstraints;
//                qry.EnableControls;
//            end;
//        finally
//            FList.Free;
//        end;
//    finally
//        Screen.Cursor := Cur;
//        FTreeValues.Free;
//    end;
//end;


//导出 TDataSet 的 FieldList 的字段内容到 Word 指定标签的表格中。
//例：全部字段：DataSetToWord(Table1, [])
//    指定字段：DataSetToWord(Table1, ['Field1', 'Field2', 'Field3'])
procedure DataSetToWord(qry: TDataSet; FieldList: Array of String; TitleList: Array of String;
        sBookMark: String = ''; DBGrid: TDBGrid = nil; bHint: Boolean = True);
var
  xList: TStringList;
  iCol: Integer;
begin
    xList := TStringList.Create;
    try
        //DataSetToList(qry, FieldList, TitleList, xList, iCol, DBGrid);

        WriteListToWord(xList, iCol, sBookMark, bHint);
    finally
        FreeAndNil(xList);
    end;
end;

//功能：导出 DBGrid 指定标题列 TitleList 的字段到 Word
//例：全部列：DBGridToWord(DBGrid1, [])
//    指定列：DBGridToWord(DBGrid1, ['列1', '列2', '列3']
procedure DBGridToWord(DBGrid: TDBGrid; TitleList: Array of String; sBookMark: String = '');
    //功能：通过 TitleList 分解 获取字段。
    procedure GetDBGridFieldList(DBGrid: TDBGrid; TitleList: Array of String;
            var FList: TStringList);
    var
      i, j: Integer;
    begin
        if Length(TitleList) = 0 then
        begin
            for i := 0 to DBGrid.Columns.Count - 1 do
            begin
                if DBGrid.Columns[i].FieldName <> '' then
                    FList.Add(DBGrid.Columns[i].FieldName);
            end;
        end
        else
        begin
            for i := 0 to Length(TitleList) - 1 do
            begin
                for j := 0 to DBGrid.Columns.Count - 1 do
                begin
                    if (DBGrid.Columns[j].FieldName <> '')
                        And (CompareText(DBGrid.Columns[j].Title.Caption, TitleList[i]) = 0) then
                    begin
                        FList.Add(DBGrid.Columns[j].FieldName);
                        break;
                    end;
                end; //for j
            end; //for i
        end; //if
    end; 
var
  FieldList: Array of String;
  FList: TStringList;
  i: Integer;
begin
    if Not Assigned(DBGrid.DataSource) then
    begin
        Application.MessageBox('DBGrid尚未指定DataSource！', '导出', MB_ICONERROR);
        Exit;
    end;
    if Not Assigned(DBGrid.DataSource.DataSet) then
    begin
        Application.MessageBox('DBGrid尚未指定DataSet数据集！', '导出', MB_ICONERROR);
        Exit;
    end;
    if Not DBGrid.DataSource.DataSet.Active then
    begin
        Application.MessageBox('DBGrid的数据集尚未打开！', '导出', MB_ICONERROR);
        Exit;
    end;

    FList := TStringList.Create;
    try
        //如果不指定列，则只导出可视的列
        if Length(TitleList) = 0 then
        begin
            for i := 0 to DBGrid.Columns.Count - 1 do
            begin
                if DBGrid.Columns[i].Visible And (DBGrid.Columns[i].Width > 0) then
                    FList.Add(DBGrid.Columns[i].FieldName);
            end;
        end
        else
            GetDBGridFieldList(DBGrid, TitleList, FList);
        SetLength(FieldList, FList.Count);
        for i := 0 to FList.Count - 1 do
            FieldList[i] := FList[i];
    finally
        FList.Free;
    end;
    DataSetToWord(DBGrid.DataSource.DataSet, FieldList, TitleList, sBookMark, DBGrid);
end;

//检测组件是否TDBGridEh、TJDBGridEh类
function IsValidDBGridEh(DBGridEh: TComponent): Boolean;
begin
    Result := False;
    if DBGridEh = nil then Exit;
    Result := Pos('$' + UpperCase(DBGridEh.ClassName) + '$', '$TDBGRIDEH$TJDBGRIDEH$') > 0;
end;

//检测组件是否是合法的TDBGridEh、TJDBGridEh类
procedure CheckValidDBGridEh(DBGridEh: TComponent);
begin
    if DBGridEh = nil then
        raise Exception.Create('DBGridEh does not Exist!');
    if Not IsValidDBGridEh(DBGridEh) then
        raise Exception.Create(Format('%s is not a valid DBGridEh!', [DBGridEh.Name]));
end;

{DBGridEh}
//导出 TDataSet 的 FieldList 的字段内容到 xList 中，并把需要导出的列数保存到 Col 中。
procedure DataSetEhToList(qry: TDataSet; FieldList: Array of String;
        var xList: TStringLIst; var Col: Integer; DBGridEh: TComponent = nil);
var
  i, j: Integer;
  FList, FTitleList: TStringList; //FieldName List
  s, sTemp: String;
  BM: TBookMark;
  FTreeCount: Integer;
  FTreeValues: TStringList;
  bNotGetTree, bUseMultiTitle: Boolean;
  eTemp: Extended;
  Cur: TCursor;
  FCol, FTitle: LongInt;
  sFieldName: String;
begin
    if DBGridEh <> nil then
      CheckValidDBGridEh(DBGridEh);
    FTreeCount := 0;
    bNotGetTree := True;
    Cur := Screen.Cursor;
    Screen.Cursor := crHourGlass;
    FTreeValues := TStringList.Create;
    try
        FList := TStringList.Create;
        try
            FTitleList := TStringList.Create;
            try
                s := '';
                bUseMultiTitle := (DBGridEh <> nil) And GetPropValue(DBGridEh, 'UseMultiTitle');
                //如果没有传入导出字段，则读取所有字段
                if Length(FieldList) = 0 then
                begin
                    //如果存在DBGridEh，则仅读取DBGridEh的列所关联的字段，以及导出标题内容 s
                    if DBGridEh <> nil then
                    begin
                        FCol := GetOrdProp(DBGridEh, 'Columns');
                        for i := 0 to TCollection(FCol).Count - 1 do
                        begin
                            sFieldName := GetStrProp(TCollection(FCol).Items[i], 'FieldName');
                            if sFieldName <> '' then
                            begin
                                FList.Add(sFieldName);
                                FTitle := GetOrdProp(TCollection(FCol).Items[i], 'Title');
                                FTitleList.Add(GetStrProp(TPersistent(FTitle), 'Caption'));
                            end; //if
                        end; //for
                        if Assigned(GetPropInfo(DBGridEh, 'TreeLayerCount')) then
                            FTreeCount := GetOrdProp(DBGridEh,'TreeLayerCount');      //得到表格树状层数
                    end
                    else //否则，读取所有字段
                        for i := 0 to qry.FieldCount - 1 do
                            FList.Add(qry.Fields[i].FieldName);
                end
                else //传入了导出字段
                begin
                    //如果存在DBGridEh，则仅读取要求导出并且DBGridEh的列所关联的字段，以及导出标题内容 s
                    if DBGridEh <> nil then
                    begin
                        if Assigned(GetPropInfo(DBGridEh, 'TreeLayerCount')) then
                            FTreeCount := GetOrdProp(DBGridEh, 'TreeLayerCount');     //得到表格树状层数

                        FCol := GetOrdProp(DBGridEh, 'Columns');
                        for i := 0 to Length(FieldList) - 1 do
                        begin
                            for j := 0 to TCollection(FCol).Count - 1 do
                            begin
                                sFieldName := GetStrProp(TCollection(FCol).Items[j], 'FieldName');
                                if (sFieldName <> '') And (CompareText(sFieldName, FieldList[i]) = 0) then
                                begin
                                    FList.Add(FieldList[i]);
                                    FTitle := GetOrdProp(TCollection(FCol).Items[j], 'Title');
                                    FTitleList.Add(GetStrProp(TPersistent(FTitle), 'Caption'));
                                    if (j >= FTreeCount - 1) And bNotGetTree then
                                    begin
                                        //得到指定字段的树状层数
                                        if j > FTreeCount - 1 then
                                            FTreeCount := FList.Count - 1
                                        else
                                            FTreeCount := FList.Count;
                                        bNotGetTree := False;
                                    end;
                                    break;
                                end; //if
                            end; //for
                        end; //for
                    end
                    else  //否则，读取所有导出字段
                        for i := 0 to Length(FieldList) - 1 do
                            FList.Add(FieldList[i]);
                end;
                s := GetExcelTitleStr(FTitleList, bUseMultiTitle);
                if s <> '' then
                begin
                    //s := Copy(s, 1, Length(s) - Length(#9));
                    xList.Add(s);  //导出标题内容 s
                end;
            finally
                FreeAndNil(FTitleList);
            end;

            //获取数据集的列数
            Col := FList.Count;

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
                        sTemp := ValidExcelCell(qry.FieldByName(FList[i]).DisplayText);

                        //处理树状层次显示
                        if i < FTreeCount then
                        begin
                            if i >= FTreeValues.Count then
                                FTreeValues.Add(sTemp)
                            else
                            if FTreeValues[i] = sTemp then
                                sTemp := ''
                            else
                            begin
                                for j := FTreeValues.Count - 1 downto i + 1 do
                                    FTreeValues.Delete(j);
                                FTreeValues[i] := sTemp;
                            end;
                        end;

                        {//防止字符串型的数值, 导入Excel后变为数值型丢失前面的零(如: '001' ===> 1).
                        if (sTemp <> '') And (qry.FieldByName(FList[i]).DataType = ftString)
                                And TryStrToFloat(sTemp, eTemp) then
			                  begin
                            if FloatToStr(eTemp) <> sTemp then  //如果不等, 才转换
                                sTemp := '=Trim("' + sTemp + '")';
			                  end;    }

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
            FList.Free;
        end;
    finally
        Screen.Cursor := Cur;
        FTreeValues.Free;
    end;
end;

//导出 TDataSet 的 FieldList 的字段内容到 Word 指定标签的表格中。
//例：全部字段：DataSetEhToWord(Table1, [])
//    指定字段：DataSetEhToWord(Table1, ['Field1', 'Field2', 'Field3'])
procedure DataSetEhToWord(qry: TDataSet; FieldList: Array of String;
        DBGridEh: TComponent = nil; sBookMark: String = '');
var
  xList: TStringList;
  iCol: Integer;
begin
    xList := TStringList.Create;
    try
        DataSetEhToList(qry, FieldList, xList, iCol, DBGridEh);

        WriteListToWord(xList, iCol, sBookMark);
    finally
        FreeAndNil(xList);
    end;
end;

//导出 DBGridEh 指定标题列 TitleList 的字段内容到 Word
//例：全部列：DBGridEhToWord(DBGridEh1, [])
//    指定列：DBGridEhToWord(DBGridEh1, ['列1', '列2', '列3']
procedure DBGridEhToWord(DBGridEh: TComponent; TitleList: Array of String; sBookMark: String = '');
    //功能：通过 TitleList 分解 获取字段。
    procedure GetDBGridEhFieldList(DBGridEh: TComponent; TitleList: Array of String;
            var FList: TStringList);
    var
      i, j: Integer;
      FCol, FTitle: LongInt;
      sFieldName, sCaption: String;
    begin
        FCol := GetOrdProp(DBGridEh, 'Columns');
        if Length(TitleList) = 0 then
        begin
            for i := 0 to TCollection(FCol).Count - 1 do
            begin
                sFieldName := GetStrProp(TCollection(FCol).Items[i], 'FieldName');
                if sFieldName <> '' then
                    FList.Add(sFieldName);
            end;
        end
        else
        begin
            for i := 0 to Length(TitleList) - 1 do
            begin
                for j := 0 to TCollection(FCol).Count - 1 do
                begin
                    sFieldName := GetStrProp(TCollection(FCol).Items[j], 'FieldName');
                    FTitle := GetOrdProp(TCollection(FCol).Items[j], 'Title');
                    sCaption := GetStrProp(TPersistent(FTitle), 'Caption');
                    if (sFieldName <> '') And (CompareText(sCaption, TitleList[i]) = 0) then
                    begin
                        FList.Add(sFieldName);
                        break;
                    end;
                end; //for j
            end; //for i
        end; //if
    end; 
var
  FieldList: Array of String;
  FList: TStringList;
  i: Integer;
  FDS, FCol, FWidth: LongInt;
  FVisible: Boolean;
begin
    CheckValidDBGridEh(DBGridEh);

    FDS := GetOrdProp(DBGridEh, 'DataSource');
    if FDS = 0 then
    begin
        Application.MessageBox('DBGridEh尚未指定DataSource！', '导出', MB_ICONERROR);
        Exit;
    end;
    if Not Assigned(TDataSource(FDS).DataSet) then
    begin
        Application.MessageBox('DBGridEh尚未指定DataSet数据集！', '导出', MB_ICONERROR);
        Exit;
    end;
    if Not TDataSource(FDS).DataSet.Active then
    begin
        Application.MessageBox('DBGridEh的数据集尚未打开！', '导出', MB_ICONERROR);
        Exit;
    end;

    FList := TStringList.Create;
    try
        //如果不指定列，则只导出可视的列
        if Length(TitleList) = 0 then
        begin
            FCol := GetOrdProp(DBGridEh, 'Columns');
            for i := 0 to TCollection(FCol).Count - 1 do
            begin
                FWidth := GetOrdProp(TCollection(FCol).Items[i], 'Width');
                FVisible := GetPropValue(TCollection(FCol).Items[i], 'Visible');
                if FVisible And (FWidth > 0) then
                    FList.Add(GetStrProp(TCollection(FCol).Items[i], 'FieldName'));
            end;
        end
        else
            GetDBGridEhFieldList(DBGridEh, TitleList, FList);
        SetLength(FieldList, FList.Count);
        for i := 0 to FList.Count - 1 do
            FieldList[i] := FList[i];
    finally
        FList.Free;
    end;
    DataSetEhToWord(TDataSource(FDS).DataSet, FieldList, DBGridEh, sBookMark);
end;

////功能：导出 StringGrid 指定列 ColList 的内容到 Word
//procedure StrGridToWordCol(StrGrid: TStringGrid; ColList: Array of Integer; sBookMark: String = '');
//var
//  i, j, k, m: Integer;
//  s, sTemp: String;
//  xList: TStringList;
//  FTreeCount: Integer;
//  FTreeValues: TStringList;
//  FList: TStringList;
//  bNotGetTree: Boolean;
//  Cur: TCursor;
//  Col: Integer;
//begin
//    FTreeCount := 0;
//    bNotGetTree := True;
//    if StrGrid is TStrGrid then
//        FTreeCount := (StrGrid as TJStrGrid).TreeLayerCount;   //得到表格树状层数
//
//    Cur := Screen.Cursor;
//    Screen.Cursor := crHourGlass;
//    xList := TStringList.Create;
//    FTreeValues := TStringList.Create;
//    FList := TStringList.Create;
//    try
//        //得到导出列FList和导出列的树状层数FTreeCount
//        if Length(ColList) = 0 then
//        begin
//            for i := StrGrid.FixedCols to StrGrid.ColCount - 1 do
//            begin
//                if StrGrid.ColWidths[i] > 0 then    //如果不指定列，则只显示列宽大于零的列
//                    FList.Add(IntToStr(i));
//            end;
//        end
//        else
//        begin
//            for i := 0 to Length(ColList) - 1 do
//                for j := StrGrid.FixedCols to StrGrid.ColCount - 1 do
//                begin
//                    if ColList[i] = j then
//                    begin
//                        FList.Add(IntToStr(j));
//                        if (j >= FTreeCount - 1) And bNotGetTree then
//                        begin
//                            //得到指定字段的树状层数
//                            if j > FTreeCount - 1 then
//                                FTreeCount := FList.Count - 1
//                            else
//                                FTreeCount := FList.Count;
//                            bNotGetTree := False;
//                        end;
//                        Break;
//                    end;
//                end;
//        end;
//
//        //获取数据集的列数
//        Col := FList.Count;
//
//        //导出数据
//        for i := 0 to StrGrid.RowCount - 1 do
//        begin
//            s := '';
//            for j := 0 to FList.Count - 1 do
//            begin
//                k := StrToInt(FList[j]);
//                sTemp := ValidExcelCell(StrGrid.Cells[k, i]);
//
//                //处理树状层次显示
//                if j < FTreeCount then
//                begin
//                    if j >= FTreeValues.Count then
//                        FTreeValues.Add(sTemp)
//                    else
//                    if FTreeValues[j] = sTemp then
//                        sTemp := ''
//                    else
//                    begin
//                        for m := FTreeValues.Count - 1 downto j + 1 do
//                            FTreeValues.Delete(m);
//                        FTreeValues[j] := sTemp;
//                    end;
//                end;
//
//                s := s + sTemp + #9;
//            end; //for j
//            if StringReplace(s, #9, '', [rfReplaceAll]) <> '' then
//                s := Copy(s, 1, Length(s) - Length(#9));
//            xList.Add(s);
//        end; //for i
//
//        WriteListToWord(xList, Col, sBookMark);
//    finally
//        xList.Free;
//        FTreeValues.Free;
//        FList.Free;
//        Screen.Cursor := Cur;
//    end;
//end;

//导出 StringGrid 指定标题列 TitleList 的内容到 Excel
//例：全部列：StrGridToWord(StrGrid1, [])
////    指定列：StrGridToWord(StrGrid1, ['列1', '列2', '列3']
//procedure StrGridToWord(StrGrid: TStringGrid; TitleList: Array of String;
//        sBookMark: String = '');
//var
//  iRow, iCol, k, m: Integer;
//  ColList: Array of Integer;
//begin
//    if (Length(TitleList) = 0) or (StrGrid.FixedRows <= 0) then
//        StrGridToWordCol(StrGrid, [], sBookMark)
//    else
//    begin
//        iRow := StrGrid.FixedRows - 1;  //如果存在多个固定行，以最后一个固定行为标题行
//        SetLength(ColList, Length(TitleList));
//        m := 0;
//        //查找指定的目标列
//        for k := 0 to Length(TitleList) - 1 do
//        begin
//            for iCol := StrGrid.FixedCols to StrGrid.ColCount - 1 do
//            begin
//                if CompareText(TitleList[k], StrGrid.Cells[iCol, iRow]) = 0 then
//                begin
//                    ColList[m] := iCol;
//                    m := m + 1;
//                    break;
//                end;
//            end; //for k
//        end; //for iCol
//        StrGridToWordCol(StrGrid, ColList, sBookMark);
//    end;
//end;

//导出表格数据的统一函数(表格包括DBGrid、StringGrid和DBGridEh）
procedure GridToWord(Grid: TWinControl; TitleList: Array of String;
        sBookMark: String = ''; UseTree: Boolean = True);
var
  bHadTree: Boolean;
  FTreeCount: Integer;
begin
  if Assigned(Grid) then
  begin
    bHadTree := False;
    FTreeCount := 0;
    if (Not UseTree) And ((Grid is TDBGrid) or (Grid is TStringGrid) or IsValidDBGridEh(Grid)) then
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
          DBGridEhToWord(Grid, TitleList, sBookMark);
    finally
      if (Not UseTree) And bHadTree And (FTreeCount <> 0) then    //恢复树层数
          SetOrdProp(Grid, 'TreeLayerCount', FTreeCount);
    end;
  end;
end;

//把tsList导出到 Word 指定标签的表格中。
procedure WriteListToWord(tsList: TStringList; Col: Integer; sBookMark: String = ''; bHint: Boolean = True);
var
  wRange, wTable: Variant;
  iRangeEnd, Row, i: Integer;
  Cur: TCursor;
  Clipboard1: TClipboard;
begin
  if tsList.Count = 0 then
  begin
    if bHint then
      Application.MessageBox('没有数据可以导出！', '提示', MB_ICONWARNING);
    Exit;
  end;

  if tsList.Count > 1000 then
  begin
    if Application.MessageBox(PChar('将要导出 ' + IntToStr(tsList.Count) + ' 行数据，可能要花费较长时间。'
          + #13#10 + #13#10 + '要继续吗？'), '提示', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) = ID_NO then
      Exit;
  end;

  Row := tsList.Count;

  Cur := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  try
    if sBookMark = '' then
    begin
      //在文档末尾
      iRangeEnd := wDoc.Range.End - 1;
      if iRangeEnd < 0 then iRangeEnd := 0;

      wRange:= wDoc.Range(iRangeEnd, iRangeEnd);
    end
    else
    begin
      //在书签处
      try
        //定位书签
        if wDoc.BookMarks.Exists(sBookMark) then
        begin
          wRange := wDoc.Bookmarks.Item(sBookMark).Range;
        end
        else
        //找不到书签，跳过
        begin
          Exit;
        end;
      except
        Application.MessageBox('出现异常，请与开发人员联系！', '错误', MB_ICONERROR);
        Exit;
      end;
    end;
    //插入表格之前换行
    wRange.InsertAfter(#13);
    //插入表格
    wTable := wDoc.Tables.Add(wRange, Row, Col);
    //设置表格边框显示
    SetTableBorderVisible(wTable, True);
    //改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。
    wTable.Columns.AutoFit;

    //==剪切板== 需要uses : Clipbrd
    try
      if Clipboard1 = nil then
      begin
        Clipboard1 := TClipboard.Create;
      end;
      Clipboard1.AsText := tsList.Text;
      wTable.Range.Paste;
      Clipboard1.Clear;
  //            for i := 1 to Col do
  //              wTable.Columns.Item(i).SetWidth(50, wdAdjustNone);
    finally
      FreeAndNil(Clipboard1);
    end;
  finally
    Screen.Cursor := Cur;
  end;
end;
end.
