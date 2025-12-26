{*********************************************************************}
{                                                                     }
{     CYQTreeUnitEx v1.0  Create By cyq                                 }
{                                                                     }
{                                                                     }
{     单元功能：常用树操作，控件为RZTreeView                          }
{     数据库结构：                                                    }
{               TreeID varchar(20), TreeName varchar(50),             }
{               PrtID varchar(20), TreLevel int                       }
{依次表示：树节点ID，树节点名称，父节点ID，当前树节点级：0级最高      }
{*********************************************************************}
{树节点编码规则：唯一性，建议设置主键}
//下次修改：RzCheckTree与RzTreeView函数不使用重载。减少代码量
//2014-01-05 增加数据库操作函数

unit CYQTreeUnitEx;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ExtCtrls, DB, StdCtrls, DBSumLst, Spin, RzTreeVw,
  FireDAC.UI.Intf, FireDAC.VCLUI.Wait, FireDAC.Stan.ExprFuncs,
  FireDAC.Phys.SQLiteDef, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.Phys.Intf, FireDAC.Stan.Def, FireDAC.Stan.Pool,
  FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Comp.Client, FireDAC.Phys.SQLite,
  FireDAC.Comp.UI, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, System.Generics.Collections;

type
  TTreeNodeDate = record    //定义一个记录类型树结构
    //BH:string;            //节点项目编号
    prtID: integer;         //父节点
    TreeID: Integer;        //树节点编码
    NodeName: string;       //树节点名称
    TreLevel: Integer;       //节点等级
    ImgIndex: Integer;      //树图标Index 建议外部处理，如果要传该参数则初始化树时传入。比较累赘
    FolderType: Boolean;    //记录的类型,true 为节点,即他的下面还有节点或记录,false 为记录,即他的下面已经没有数据,也就是最后的一层
  end;

  NodeData = ^TTreeNodeDate;

var
  Node: TTreeNode;
  PNode: NodeData;
  //初始化树结构

procedure IniTreeData(Tree: TRzCheckTree; SourceSQL: string; qryTemp: TFDQuery); overload

procedure IniTreeData(Tree: TRzTreeView; SourceSQL: string; qryTemp: TFDQuery); overload

procedure IniTreeData(Tree: TRzTreeView; qryTemp: TFDQuery); overload

procedure IniTreeDataFast(Tree: TRzTreeView; SourceSQL: string; qryTemp: TDataSet);
  //树操作

function CheckNullTree(Tree: TRzCheckTree): Boolean; overload

procedure CheckNoTree(Tree: TRzCheckTree); //取消选中任何树节点

function GetMaxTreeID(ATreeView: TRzTreeView): Integer;  //获取树最大编号

function GetCurTreeLevel(ATreeView: TRzTreeView): Integer;//获取当前节点级别

function GetCurTreeID(ATreeView: TRzTreeView): Integer; overload  //获取当前树节点ID

function GetCurTreeID(ATreeView: TRzCheckTree): Integer; overload  //获取当前树节点ID

function SelectedTree(ATreeView: TRzTreeView): Boolean;   //检查树是否有选中

function IsExistsChildNode(ATreeView: TRzTreeView): Boolean; overload

function IsExistsChildNode(ATreeNode: TTreeNode): Boolean; overload
  //增加树同级节点

function AddTreeNode(NodeName: string; ATreeView: TRzTreeView; ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;
  //增加子节点

function AddTreeChildNode(NodeName: string; ATreeView: TRzTreeView; ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;
  //删除选中节点（支持有子节点）

function DeleteTree(ATreeView: TRzTreeView; ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;
  //修改节点名称

function ModifyNodeName(NewNodeName: string; ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;
  //检查是否存在同级节点

function CheckSameTreeName(NodeName: string; ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;

function CheckChildTreeName(NodeName: string; ANode: TTreeNode): Boolean;

  //获得全部根节点
function GetRootNodes(ATreeView: TRzTreeView): TStringList;
  //获得某节点全部子节点

procedure GetChildNodes(ATreeNode: TTreeNode; AStringList: TStringList);
  //获得某节点全部子节点

procedure GetAllChildNodes(ATreeNode: TTreeNode; AStringList: TStringList);

implementation

uses
  CYQCommonUnit;

procedure IniTreeData(Tree: TRzCheckTree; SourceSQL: string; qryTemp: TFDQuery);

  procedure FillOneNode(qry: TFDQuery; TreeName: TRzCheckTree; ParentNode: TTreeNode);
  begin
    with qry do
    begin
      New(PNode);
      //PNode.BH :=Trim(FieldByName('jcxmbh').AsString);
      PNode.TreLevel := FieldByName('TreLevel').asInteger;
      PNode.TreeID := FieldByName('TreeID').asInteger;
      PNode.NodeName := Trim(FieldByName('TreeName').AsString);
      PNode.ImgIndex := FieldByName('TreLevel').asInteger;
//      if FieldByName('TreLevel').asInteger=3 then
//        PNode.FolderType := false
//      else
//        PNode.FolderType := True;
      Node := TreeName.Items.AddChildObject(ParentNode, PNode.NodeName, PNode);
      Node.ImageIndex := PNode.ImgIndex;
      Node.SelectedIndex := 0;
    end;
  end;

  function FindNode(TreeName: TRzCheckTree; TreeID: string): TTreeNode;

    function FindChildNode(TreeName: TRzCheckTree; TreeID: string; CurrNode: TTreeNode): TTreeNode;
    var
      Node: TTreeNode;
    begin
      Result := nil;
      Node := CurrNode;
      if Assigned(Node.Data) and SameText(IntToStr(NodeData(Node.Data).TreeID), TreeID) then
      begin
        Result := Node;
        Exit;
      end;
      Node := Node.getFirstChild;
      while Assigned(Node) do
      begin
        //递归找下级节点
        Result := FindChildNode(TreeName, TreeID, Node);
        if Assigned(Result) then
          Exit;
        Node := Node.getNextSibling;
      end;
    end;

  var
    RootNode: TTreeNode;
  begin
    RootNode := TreeName.Items.GetFirstNode;
    while Assigned(RootNode) do
    begin
      Result := FindChildNode(TreeName, TreeID, RootNode);
      if Assigned(Result) then
        Exit;
      RootNode := RootNode.getNextSibling;
    end;
  end;

var
  iLevel: Integer;
begin
  //加载树
  Tree.Items.BeginUpdate;
  try
    Tree.Items.Clear;
    with qryTemp do
    begin
      Close;
      SQL.Text := SourceSQL;//'Exec sp_ReturnXMLXTree2';
      Open;
      First;
      iLevel := 0;
      while Locate('TreLevel', iLevel, []) do
      begin
        while not Eof do
        begin
          if FieldByName('TreLevel').AsInteger = iLevel then
          begin
            if iLevel = 0 then
              FillOneNode(qryTemp, Tree, nil)
            else
              FillOneNode(qryTemp, Tree, FindNode(Tree, FieldByName('PrtID').AsString));
          end;
          Next;
        end;
        Inc(iLevel);
      end;
      Close;
    end;
  finally
    Tree.Items.EndUpdate;
  end;
//  if Tree.Items.Count > 0 then
//    Tree.Items[0].Expanded := True;   //展开外部处理
    //Tree.Items[0].ImageIndex := 1;   图标外部处理
end;

procedure IniTreeData(Tree: TRzTreeView; SourceSQL: string; qryTemp: TFDQuery);

  procedure FillOneNode(qry: TFDQuery; TreeName: TRzTreeView; ParentNode: TTreeNode);
  begin
    with qry do
    begin
      New(PNode);
      //PNode.BH :=Trim(FieldByName('jcxmbh').AsString);
      PNode.TreLevel := FieldByName('TreLevel').asInteger;
      PNode.TreeID := FieldByName('TreeID').AsInteger;
      PNode.NodeName := Trim(FieldByName('TreeName').AsString);
      PNode.ImgIndex := FieldByName('TreLevel').asInteger;
//      if FieldByName('TreLevel').asInteger=3 then  //如果展开，默认展开三级
//        PNode.FolderType := false
//      else
//        PNode.FolderType := True;
      Node := TreeName.Items.AddChildObject(ParentNode, PNode.NodeName, PNode);
      Node.ImageIndex := PNode.ImgIndex;
      Node.SelectedIndex := 0;
    end;
  end;

  function FindNode(TreeName: TRzTreeView; TreeID: string): TTreeNode;

    function FindChildNode(TreeName: TRzTreeView; TreeID: string; CurrNode: TTreeNode): TTreeNode;
    var
      Node: TTreeNode;
    begin
      Result := nil;
      Node := CurrNode;
      if Assigned(Node.Data) and SameText(IntToStr(NodeData(Node.Data).TreeID), TreeID) then
      begin
        Result := Node;
        Exit;
      end;
      Node := Node.getFirstChild;
      while Assigned(Node) do
      begin
        //递归找下级节点
        Result := FindChildNode(TreeName, TreeID, Node);
        if Assigned(Result) then
          Exit;
        Node := Node.getNextSibling;
      end;
    end;

  var
    RootNode: TTreeNode;
  begin
    RootNode := TreeName.Items.GetFirstNode;
    while Assigned(RootNode) do
    begin
      Result := FindChildNode(TreeName, TreeID, RootNode);
      if Assigned(Result) then
        Exit;
      RootNode := RootNode.getNextSibling;
    end;
  end;

var
  iLevel: Integer;
begin
  //加载树
  Tree.Items.BeginUpdate;
  try
    Tree.Items.Clear;
    with qryTemp do
    begin
      Close;
      SQL.Text := SourceSQL;//'Exec sp_ReturnXMLXTree2';
      Open;
      First;
      iLevel := 0;
      while Locate('TreLevel', iLevel, []) do
      begin
        while not Eof do
        begin
          if FieldByName('TreLevel').AsInteger = iLevel then
          begin
            if iLevel = 0 then
              FillOneNode(qryTemp, Tree, nil)
            else
              FillOneNode(qryTemp, Tree, FindNode(Tree, FieldByName('PrtID').AsString));
          end;
          Next;
        end;
        Inc(iLevel);
      end;
      Close;
    end;
  finally
    Tree.Items.EndUpdate;
  end;
end;

procedure IniTreeData(Tree: TRzTreeView; qryTemp: TFDQuery);

  procedure FillOneNode(qry: TFDQuery; TreeName: TRzTreeView; ParentNode: TTreeNode);
  begin
    with qry do
    begin
      New(PNode);
      //PNode.BH :=Trim(FieldByName('jcxmbh').AsString);
      PNode.TreLevel := FieldByName('TreLevel').asInteger;
      PNode.TreeID := FieldByName('TreeID').asInteger;
      PNode.NodeName := Trim(FieldByName('TreeName').AsString);
      PNode.ImgIndex := FieldByName('TreLevel').asInteger;
//      if FieldByName('TreLevel').asInteger=3 then
//        PNode.FolderType := false
//      else
//        PNode.FolderType := True;
      Node := TreeName.Items.AddChildObject(ParentNode, PNode.NodeName, PNode);
      Node.ImageIndex := PNode.ImgIndex;
      Node.SelectedIndex := 0;
    end;
  end;

  function FindNode(TreeName: TRzTreeView; TreeID: string): TTreeNode;

    function FindChildNode(TreeName: TRzTreeView; TreeID: string; CurrNode: TTreeNode): TTreeNode;
    var
      Node: TTreeNode;
    begin
      Result := nil;
      Node := CurrNode;
      if Assigned(Node.Data) and SameText(IntToStr(NodeData(Node.Data).TreeID), TreeID) then
      begin
        Result := Node;
        Exit;
      end;
      Node := Node.getFirstChild;
      while Assigned(Node) do
      begin
        //递归找下级节点
        Result := FindChildNode(TreeName, TreeID, Node);
        if Assigned(Result) then
          Exit;
        Node := Node.getNextSibling;
      end;
    end;

  var
    RootNode: TTreeNode;
  begin
    RootNode := TreeName.Items.GetFirstNode;
    while Assigned(RootNode) do
    begin
      Result := FindChildNode(TreeName, TreeID, RootNode);
      if Assigned(Result) then
        Exit;
      RootNode := RootNode.getNextSibling;
    end;
  end;

begin
  //加载树
  Tree.Items.BeginUpdate;
  try
    Tree.Items.Clear;
    with qryTemp do
    begin
      First;
      while not Eof do
      begin
        if FieldByName('TreLevel').AsInteger = 0 then
          FillOneNode(qryTemp, Tree, nil)
        else
          FillOneNode(qryTemp, Tree, FindNode(Tree, FieldByName('PrtID').AsString));
        Next;
      end;
      Close;
    end;
  finally
    Tree.Items.EndUpdate;
  end;
end;

procedure IniTreeDataFast(Tree: TRzTreeView; SourceSQL: string; qryTemp: TDataSet);
var
  NodeMap: TDictionary<Integer, TTreeNode>; // TreeID => TreeNode
  DataList: TList<NodeData>;
  i: Integer;
  P: NodeData;
  Node: TTreeNode;
begin
  Tree.Items.BeginUpdate;
  try
    Tree.Items.Clear;
    NodeMap := TDictionary<Integer, TTreeNode>.Create;
    DataList := TList<NodeData>.Create;
    try
      // 第一步：加载所有数据进内存
      if qryTemp is TFDQuery then
      begin
        OpenDataSet(TFDQuery(qryTemp), SourceSQL)
      end;

      while not qryTemp.Eof do
      begin
        New(P);
        P^.TreeID := qryTemp.FieldByName('TreeID').AsInteger;
        P^.PrtID := qryTemp.FieldByName('PrtID').AsInteger;
        P^.TreLevel := qryTemp.FieldByName('TreLevel').AsInteger;
        P^.NodeName := Trim(qryTemp.FieldByName('TreeName').AsString);
        P^.ImgIndex := P^.TreLevel;
        DataList.Add(P);
        qryTemp.Next;
      end;

      // 第二步：构造树结构
      for i := 0 to DataList.Count - 1 do
      begin
        P := DataList[i];
        if (P^.PrtID = 0) and (P^.TreLevel = 0) then // 顶级节点
          Node := Tree.Items.AddChildObject(nil, P^.NodeName, P)
        else if NodeMap.TryGetValue(P^.PrtID, Node) then
          Node := Tree.Items.AddChildObject(Node, P^.NodeName, P)
        else
          Continue; // 找不到父节点，跳过或记录日志

        Node.ImageIndex := P^.ImgIndex;
        Node.SelectedIndex := 0;
        NodeMap.Add(P^.TreeID, Node);
      end;
    finally
      NodeMap.Free;
      DataList.Free; // 注意：如果要管理内存，还需逐个 Dispose(P)
    end;
  finally
    Tree.Items.EndUpdate;
  end;
end;

function CheckNullTree(Tree: TRzCheckTree): Boolean;
var
  iNode: Integer;
begin
  Result := False;
  if Assigned(Tree) then
    if Tree.Items.Count = 0 then
      Exit;
  for iNode := 0 to Tree.Items.Count - 1 do
    if Tree.ItemState[iNode] = csChecked then
    begin
      Result := True;
      Exit;
    end;
end;

procedure CheckNoTree(Tree: TRzCheckTree);
var
  iNode: Integer;
begin
  if Assigned(Tree) then
    if Tree.Items.Count = 0 then
      Exit;
  for iNode := 0 to Tree.Items.Count - 1 do
    Tree.ItemState[iNode] := csUnchecked;
  //Tree.ItemState[iNode] := csUnknown;
  //Tree.ItemState[iNode] := csPartiallyChecked;
end;

function GetMaxTreeID(ATreeView: TRzTreeView): Integer;
var
  i: integer;
  TempNode: NodeData;
begin
  Result := 0;
  for i := 0 to ATreeView.Items.Count - 1 do
  begin
    TempNode := NodeData(ATreeView.Items.Item[i].Data);
    if Result <= TempNode.TreeID then
      Result := TempNode.TreeID;
  end;
end;

function GetCurTreeLevel(ATreeView: TRzTreeView): Integer;
var
  TempNode: NodeData;
begin
  TempNode := NodeData(ATreeView.Selected.Data);
  Result := TempNode.TreLevel;
end;

function GetCurTreeID(ATreeView: TRzTreeView): Integer;
var
  TempNode: NodeData;
begin
  TempNode := NodeData(ATreeView.Selected.Data);
  Result := TempNode.TreeID;
end;

function GetCurTreeID(ATreeView: TRzCheckTree): Integer;
var
  TempNode: NodeData;
begin
  TempNode := NodeData(ATreeView.Selected.Data);
  Result := TempNode.TreeID;
end;

function SelectedTree(ATreeView: TRzTreeView): Boolean;
begin
  Result := True;
  if not Assigned(ATreeView.Selected) then
  begin
    Result := False;
    raise Exception.Create('请选择相应的树节点！');
  end;
end;

// Added by CYQ 2014-01-08 15:47:35
//增加同级树节点
function AddTreeNode(NodeName: string; ATreeView: TRzTreeView; ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;
var
  NewNode: TTreeNode;
  strSQL: string;
begin
  Result := False;
  if NodeName = '' then
  begin
    ShowMessageBoxEx('请输入树节点名称！', 'info');
    Exit;
  end;
  if TableName = '' then
  begin
    ShowMessageBoxEx('无法获取表名！', 'info');
    Exit;
  end;

  if not SelectedTree(ATreeView) then
    Exit;
  try
    try
      ATreeView.Items.BeginUpdate;

      Result := True;
      New(PNode);
      PNode.TreeID := GetMaxTreeID(ATreeView) + 1;
      PNode.NodeName := NodeName;
      PNode.TreLevel := GetCurTreeLevel(ATreeView);

      strSQL := 'Insert Into ' + TableName + ' (TreeID, PrtID, TreeName, TreLevel)' + ' values(' + IntToStr(PNode.TreeID) + ',';
      if ANode = nil then
        ANode := ATreeView.Items.Item[0];
      if ANode.Level = 0 then
        strSQL := strSQL + 'null,' + QuotedStr(NodeName) + ', 0)'
      else
        strSQL := strSQL + IntToStr(NodeData(ANode.Parent.Data).TreeID) + ',' + QuotedStr(NodeName) + ',' + IntToStr(PNode.TreLevel) + ')';

      if not CheckSameTreeName(NodeName, ANode, TableName, qryTemp) then
        Exit;

      if not CheckChildTreeName(NodeName, ANode) then
        if ExecuteSQL(qryTemp, strSQL) then
        begin
          if Node.Level = 0 then
            NewNode := ATreeView.Items.AddChildObject(ANode, NodeName, PNode)
          else
            NewNode := ATreeView.Items.AddObject(ANode, NodeName, PNode);
          NewNode.ImageIndex := 0;
          NewNode.SelectedIndex := 2;   //外部处理
        end;
    finally
      ATreeView.Items.EndUpdate;
    end;
  except
    Result := False;
  end;
  ATreeView.Refresh;
end;

//增加指定节点子节点
function AddTreeChildNode(NodeName: string; ATreeView: TRzTreeView; ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;
var
  NewNode: TTreeNode;
  strSQL: string;
begin
  Result := False;
  if NodeName = '' then
  begin
    ShowMessageBoxEx('请输入树节点名称！', 'info');
    Exit;
  end;
  if TableName = '' then
  begin
    ShowMessageBoxEx('无法获取表名！', 'info');
    Exit;
  end;
  if not SelectedTree(ATreeView) then
    Exit;
  try
    try
      ATreeView.Items.BeginUpdate;

      Result := True;
      New(PNode);
      PNode.TreeID := GetMaxTreeID(ATreeView) + 1;
      PNode.NodeName := NodeName;
      PNode.TreLevel := GetCurTreeLevel(ATreeView) + 1;

      strSQL := 'Insert Into ' + TableName + ' (TreeID, PrtID, TreeName, TreLevel)' + ' Values(' + IntToStr(PNode.TreeID) + ',';
      if ANode = nil then
        ANode := ATreeView.Items.Item[0];
      strSQL := strSQL + IntToStr(NodeData(ANode.Data).TreeID) + ',' + QuotedStr(NodeName) + ',' + IntToStr(PNode.TreLevel) + ')';
      if not CheckChildTreeName(NodeName, ANode) then
        if ExecuteSQL(qryTemp, strSQL) then
          NewNode := ATreeView.Items.AddChildObject(ANode, NodeName, PNode);
      NewNode.ImageIndex := 1;
      NewNode.SelectedIndex := 2;   //外部处理
    finally
      ATreeView.Items.EndUpdate;
    end;
  except
    Result := False;
  end;
  ATreeView.Refresh;
end;

//修改树节点名称
function ModifyNodeName(NewNodeName: string; ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;
var
  NewNode: TTreeNode;
  strSQL: string;
begin
  strSQL := ' Update ' + TableName + ' Set TreeName = ' + QuotedStr(NewNodeName) + ' Where TreeID = ' + IntToStr(NodeData(ANode.Data).TreeID);
  if ExecuteSQL(qryTemp, strSQL) then
  begin
    ANode.Text := NewNodeName;
    NodeData(ANode.Data).NodeName := NewNodeName;
  end;
end;

function IsExistsChildNode(ATreeView: TRzTreeView): Boolean; overload;
var
  Node: TTreeNode;
begin
  Result := False;
  if not SelectedTree(ATreeView) then
    Exit;
  Node := ATreeView.Selected;
  Result := Node.HasChildren;
end;

function IsExistsChildNode(ATreeNode: TTreeNode): Boolean; overload;
begin
  Result := ATreeNode.HasChildren;
end;

//删除选中树节点（包括其下子节点）
function DeleteTree(ATreeView: TRzTreeView; ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;

  function DelTreeDataByID(TreeID: integer): boolean;
  var
    strSQL: string;
  begin
    Result := False;
    strSQL := 'Delete From ' + TableName + ' Where TreeID = ' + IntToStr(TreeID);
    if ExecuteSQL(qryTemp, strSQL) then
      Result := True;
  end;

  function DelTreeNode(ParentID: integer): Boolean;
  var
    qryExecSQL: TFDQuery;
    strSQL, FErrorInfo: string;
    TreeID: Integer;
  begin
    Result := False;
    try
      qryExecSQL := TFDQuery.Create(nil);
      qryExecSQL.Connection := qryTemp.Connection;
      strSQL := 'Select * From ' + TableName + ' Where  PrtID = ' + IntToStr(ParentID);
      OpenDataSet(qryExecSQL, strSQL);
      with qryExecSQL do
        if RecordCount > 0 then
        begin
          First;
          while not Eof do
          begin
            TreeID := FieldByName('TreeID').AsInteger;
            DelTreeNode(TreeID);
            Result := DelTreeDataByID(TreeID);
            Next;
          end;
        end;
      qryExecSQL.Free;
    except
      on e: Exception do
      begin
        Result := False;
        FErrorInfo := e.Message;
        FErrorInfo := GetSQLErrorChineseInfo(FErrorInfo);
        Application.MessageBox(PChar(FErrorInfo), '错误', MB_ICONERROR);
        Exit;
      end;
    end;
    Result := DelTreeDataByID(ParentID);
  end;

begin
  Result := False;

  if ANode.AbsoluteIndex = 0 then
  begin
    raise Exception.Create('禁止删除最高根节点');
    Exit;
  end;

//  if IsExistsChildNode(ATreeView) then
//    if not ShowMessageBoxEx('该节点存在子节点，确定删除？', 'ask') then Exit;

  if DelTreeNode(NodeData(ANode.Data).TreeID) then
  begin
    ATreeView.Items.Delete(ANode);
    Result := True;
  end;
end;

//判断同同级节点是否有重复
function CheckSameTreeName(NodeName: string; ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;
var
  strSQL: string;
begin
  Result := True;
  strSQL := 'Select * From ' + TableName + ' Where TreLevel = ' + IntToStr(NodeData(ANode.Data).TreLevel) + ' and TreeName = ' + QuotedStr(NodeName);
  with qryTemp do
  begin
    OpenDataSet(qryTemp, strSQL);
    if RecordCount > 0 then
    begin
      ShowMessageBoxEx('该级已经存在相同的节点名称，请检查！', 'info');
      Result := False;
      Exit;
    end;
  end;
end;

//检查子节点是否有重复。（包括子节点的子节点）
//function CheckChildTreeName(NodeName: string;ANode: TTreeNode; TableName: string; qryTemp: TFDQuery): Boolean;
//begin
////
//end;

function CheckChildTreeName(NodeName: string; ANode: TTreeNode): Boolean;
var
  i: integer;
  strTemp, strNodeName: string;
begin
  Result := False;
  strNodeName := NodeName;
  for i := 0 to ANode.Count - 1 do
  begin
    strTemp := trim(Trim(ANode[i].Text));
    if SameText(UpperCase(Trim(strNodeName)), UpperCase(Trim(strTemp))) then
    begin
      Result := True;
      ShowMessageBoxEx('已有相同子节点，请检查！', 'info');
      Exit;
    end;
    if ANode[i].HasChildren then
      Result := CheckChildTreeName(strNodeName, ANode[i]);
  end;
end;

function GetRootNodes(ATreeView: TRzTreeView): TStringList;
var
  TempNode: TTreeNode;
begin
  Result := TStringList.Create;
  TempNode := ATreeView.Items.GetFirstNode;
  if Assigned(TempNode) then
    while TempNode <> nil do
    begin
      Result.Add(IntToStr(NodeData(TempNode.Data).TreeID));
      TempNode := TempNode.getNextSibling;
    end
  else
    Result.Text := '';
end;

procedure GetChildNodes(ATreeNode: TTreeNode; AStringList: TStringList);
var
  TempNode: TTreeNode;
begin
  TempNode := ATreeNode.getFirstChild;
  if Assigned(TempNode) then
    while TempNode <> nil do
    begin
      if not TempNode.HasChildren then
        AStringList.Add(IntToStr(NodeData(TempNode.Data).TreeID));
      GetChildNodes(TempNode, AStringList);
      TempNode := TempNode.GetNextChild(TempNode);
    end;
  AStringList.Sort;
end;

procedure GetAllChildNodes(ATreeNode: TTreeNode; AStringList: TStringList);
var
  TempNode: TTreeNode;
begin
  TempNode := ATreeNode.getFirstChild;
  if Assigned(TempNode) then
    while TempNode <> nil do
    begin
      AStringList.Add(IntToStr(NodeData(TempNode.Data).TreeID));
      GetAllChildNodes(TempNode, AStringList);
      TempNode := TempNode.GetNextChild(TempNode);
    end;
  AStringList.Sort;
end;

end.

