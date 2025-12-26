{******************************************************************************}
{   Create by CYQ 2013-09-16 14:41:19                                          }
{   增加导出Word文档功能时，发现列数太多而显示在Word文档上面的表格很丑陋       }
{   故增加该功能提供给用户自定义导出所需要的数据列内容。                       }
{   该功能对导出Excel与Word通用                                                }
{******************************************************************************}

unit FormColSelect;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGridEh, CheckLst, Buttons, ComCtrls,
  ExtCtrls, Math, IniFiles;

type
  TFrmColSelect = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    CLB: TCheckListBox;
    Panel3: TPanel;
    btnExit: TBitBtn;
    btnSave: TBitBtn;
    UpDownFix: TUpDown;
    EdtFix: TEdit;
    Label1: TLabel;
    ChkMultiTitle: TCheckBox;
    Panel4: TPanel;
    ChkSort: TCheckBox;
    ChkVisiable: TCheckBox;
    ChkReadOnly: TCheckBox;
    Label2: TLabel;
    CbxSumType: TComboBox;
    CmbFieldname: TComboBox;
    EditSumValue: TLabeledEdit;
    Label3: TLabel;
    CheckBox2: TCheckBox;
    Label4: TLabel;
    Label5: TLabel;
    EdtFieldName: TEdit;
    EdtFieldTitle: TEdit;
    btnLast: TSpeedButton;
    btnNext: TSpeedButton;
    btnPrior: TSpeedButton;
    btnFirst: TSpeedButton;
    ChkCount: TCheckBox;
    chkSaveSettint: TCheckBox;
    chbSearchAll: TCheckBox;
    chbSearchCol: TCheckBox;
    chbShowRowRecNo: TCheckBox;
    procedure Button2Click(Sender: TObject);
    procedure CLBDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure chkSaveSettintClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure CLBClick(Sender: TObject);
    procedure CbxSumTypeChange(Sender: TObject);
    procedure CmbFieldnameChange(Sender: TObject);
    procedure EditSumValueChange(Sender: TObject);
    procedure ChkReadOnlyClick(Sender: TObject);
    procedure EdtFieldTitleChange(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure frBtnNextClick(Sender: TObject);
    procedure frBtnPrevClick(Sender: TObject);
    procedure btnLastClick(Sender: TObject);
    procedure btnNextClick(Sender: TObject);
    procedure btnPriorClick(Sender: TObject);
    procedure btnFirstClick(Sender: TObject);
  private
    OldIndex: Integer;
    { Private declarations }
  public
    Grid: TDBGridEH;
    procedure LoadData;
    procedure SaveData;
    procedure SaveFieldValue;
    { Public declarations }
  end;

procedure ShowGridColEditor(Grid: TDBGridEH);

var
  FrmColSelect: TFrmColSelect;

implementation

uses
  CYQCommonUnit;

{$R *.dfm}

procedure ShowGridColEditor(Grid: TDBGridEH);
begin
  //显示GRIDEh的编辑窗口
  IsValidDBGridEh(Grid);
  if not Assigned(Grid) or not Assigned(Grid.DataSource) or not
    Assigned(Grid.DataSource.DataSet) then
    Exit;

  if not Assigned(FrmColSelect) then
    FrmColSelect := TFrmColSelect.Create(Application);
  try
    FrmColSelect.Grid := Grid;
    FrmColSelect.ShowModal;
  finally
    FreeAndNil(FrmColSelect);
  end;
end;

procedure TFrmColSelect.Button2Click(Sender: TObject);
begin
  SaveData;
  ModalResult := mrOk;
end;

procedure TFrmColSelect.CLBDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  if Sender = Source then
  begin
    Accept := true;
    OldIndex := TCheckListBox(sender).ItemIndex;
  end
  else
    OldIndex := -1;
end;

procedure TFrmColSelect.SpeedButton1Click(Sender: TObject);
var
  i: integer;
begin
  if CLB.ItemIndex > 0 then
  begin
    i := clb.ItemIndex - 1;
    clb.Items.Move(clb.ItemIndex, i);
    CLB.ItemIndex := i;
    CLB.Selected[i];
  end;
end;

procedure TFrmColSelect.SpeedButton2Click(Sender: TObject);
var
  i: integer;
begin
  if CLB.ItemIndex < CLB.Count - 1 then
  begin
    i := clb.ItemIndex + 1;
    clb.Items.Move(clb.ItemIndex, i);
    CLB.ItemIndex := i;
    CLB.Selected[i];
  end;
end;

procedure TFrmColSelect.chkSaveSettintClick(Sender: TObject);
//此处代码用于保存界面的设置到INI文件
//该功能会导致ini文件有大量冗余信息，作废该功能  by cyq
//var
//  Ini: TIniFile;
begin
//    if CheckBox1.Checked then
//      Grid.SaveGridLayoutIni(GetIniFile,
//        Grid.Owner.ClassName + '$' + Grid.Name, true)
//    else
//    begin
//      Ini := TIniFile.Create(ExtractFilePath(Application.ExeName) + '\Setting.ini');
//      try
//        Ini.EraseSection(Grid.Owner.ClassName + '$' + Grid.Name);
//      finally
//        Ini.Free;
//      end;
//    end;
end;

procedure TFrmColSelect.FormShow(Sender: TObject);
begin
  LoadData;

  btnExit.SetFocus;
end;

procedure TFrmColSelect.LoadData;
var
  i: integer;
begin
  //装入GRID的数据到界面 by cyq
  //第一步,取全局参数
  Grid.DataSource.DataSet.DisableControls;
  ChkCount.Checked := Grid.FooterRowCount > 0;
  ChkMultiTitle.Checked := Grid.UseMultiTitle;
  EdtFix.Text := inttostr(Grid.FrozenCols);
  UpDownFix.Max := Grid.Columns.Count - 1;
  UpDownFix.Position := Grid.FrozenCols;
  chbSearchAll.Checked := Grid.SearchPanel.Enabled;
  chbSearchCol.Checked := Grid.STFilter.Visible;
  if gioShowRecNoEh in Grid.IndicatorOptions then
    chbShowRowRecNo.Checked := True;

  //从GRID中装入COLUMNS信息
  CLB.Clear;
  for i := 0 to Grid.Columns.Count - 1 do
  begin
    CLB.Items.AddObject(Grid.Columns.Items[i].Title.Caption,
      Grid.Columns.Items[i]);
    clb.Checked[i] := Grid.Columns.Items[i].Visible;
    //填充cmbFieldName的字段选项
    CmbFieldname.Items.Add(Grid.Columns.Items[i].FieldName);
  end;
  CLB.Selected[0] := True;
  Grid.DataSource.DataSet.EnableControls;
end;

procedure TFrmColSelect.SaveData;
var
  i: integer;
begin
  //从界面中的设置保存到GRID中 by cyq
  //保存全局GRID的参数
  Grid.FooterRowCount := IfThen(ChkCount.Checked, 1, 0);
  Grid.UseMultiTitle := ChkMultiTitle.Checked;
  Grid.FrozenCols := UpDownFix.Position;
  Grid.SearchPanel.Enabled := chbSearchAll.Checked;
  Grid.SearchPanel.CaseSensitive := True;
  Grid.STFilter.Visible := chbSearchCol.Checked;
  if chbSearchCol.Checked then
  begin
    Grid.STFilter.InstantApply := True;
    Grid.STFilter.Local := True;
  end;

  if chbShowRowRecNo.Checked then
    Grid.IndicatorOptions := Grid.IndicatorOptions +[gioShowRecNoEh]
  else
    Grid.IndicatorOptions := Grid.IndicatorOptions - [gioShowRecNoEh];

  for i := 0 to CLB.Count - 1 do
  begin
    TColumnEh(CLB.Items.Objects[i]).Visible := CLB.Checked[i];
    TColumnEh(CLB.Items.Objects[i]).Index := i;
  end;
end;

procedure TFrmColSelect.SpeedButton6Click(Sender: TObject);
begin
  if Clb.ItemIndex >= 0 then
  begin
    clb.Items.Move(clb.ItemIndex, 0);
    CLB.Selected[0];
  end;
end;

procedure TFrmColSelect.CLBClick(Sender: TObject);
var
  i: integer;
begin
  if CLB.ItemIndex >= 0 then
  begin
    i := clb.ItemIndex;

    case TColumnEh(CLB.Items.Objects[i]).Footer.ValueType of
      fvtNon: CbxSumType.ItemIndex := 0;
      fvtStaticText: CbxSumType.ItemIndex := 1;
      fvtFieldValue: CbxSumType.ItemIndex := 2;
      fvtAvg: CbxSumType.ItemIndex := 3;
      fvtCount: CbxSumType.ItemIndex := 4;
      fvtSum: CbxSumType.ItemIndex := 5;
    end;
    EditSumValue.Text:=TColumnEh(CLB.Items.Objects[i]).Footer.Value;

    EdtFieldName.Text := TColumnEh(CLB.Items.Objects[i]).FieldName;
    EdtFieldTitle.Text := TColumnEh(CLB.Items.Objects[i]).Title.Caption;
    TColumnEh(CLB.Items.Objects[i]).Visible := CLB.Checked[i];
    ChkVisiable.Checked := TColumnEh(CLB.Items.Objects[i]).Visible;
    ChkSort.Checked := TColumnEh(CLB.Items.Objects[i]).Title.TitleButton;
    ChkReadOnly.Checked := TColumnEh(CLB.Items.Objects[i]).ReadOnly;
    case TColumnEh(CLB.Items.Objects[i]).Footer.ValueType of
      fvtNon: CbxSumType.ItemIndex := 0;
      fvtStaticText: CbxSumType.ItemIndex := 1;
      fvtFieldValue: CbxSumType.ItemIndex := 2;
      fvtAvg: CbxSumType.ItemIndex := 3;
      fvtCount: CbxSumType.ItemIndex := 4;
      fvtSum: CbxSumType.ItemIndex := 5;
    else
      CbxSumType.ItemIndex := -1;
    end;
  end;
end;

procedure TFrmColSelect.CbxSumTypeChange(Sender: TObject);
var
  bField, bValue: Boolean; //用这两个变量来标识字段选择及文字框是否可用
begin
  Grid.DataSource.DataSet.DisableControls;
  case CbxSumType.ItemIndex of
    0:begin
        bField := False;
        bValue := False;
      end;
    1:begin
        bField := False;
        bValue := True;
      end;
    2:begin
        bField := True;
        bValue := False;
      end;
    3:begin
        bField := True;
        bValue := False;
      end;
    4:begin
        bField := True;
        bValue := False;
      end;
    5:
      begin
        bField := True;
        bValue := False;
      end;
  else
    begin
      bField := False;
      bValue := False;
    end;
  end;
  if bField then
    CmbFieldname.ItemIndex :=
      CmbFieldname.Items.IndexOf(TColumnEh(CLB.Items.Objects[CLB.ItemIndex]).FieldName);
  CmbFieldname.Enabled := bField;

  if bValue then
    EditSumValue.Text :=
      TColumnEh(CLB.Items.Objects[CLB.ItemIndex]).Footer.Value;
  EditSumValue.Enabled := bValue;

  SaveFieldValue;
  Grid.DataSource.DataSet.EnableControls;
end;

procedure TFrmColSelect.SaveFieldValue;
var
  i: integer;
  v: TFooterValueType;
begin
  //保存Field部分的设置
  try
    Grid.DataSource.DataSet.DisableControls;
    i := CLB.ItemIndex;
    if i >= 0 then
    begin
      case CbxSumType.ItemIndex of
        0: v := fvtNon;
        1: v := fvtStaticText;
        2: v := fvtFieldValue;
        3: v := fvtAvg;
        4: v := fvtCount;
        5: v := fvtSum;
        -1: v := fvtNon;
      end;
      TColumnEh(CLB.Items.Objects[i]).Footer.ValueType := v;
      TColumnEh(CLB.Items.Objects[i]).Footer.FieldName := CmbFieldname.Text;
      TColumnEh(CLB.Items.Objects[i]).Footer.Value := EditSumValue.Text;
    end;
    TColumnEh(CLB.Items.Objects[i]).Visible := ChkVisiable.Checked;
    TColumnEh(CLB.Items.Objects[i]).ReadOnly := ChkReadOnly.Checked;
    TColumnEh(CLB.Items.Objects[i]).Title.TitleButton := ChkSort.Checked;
    TColumnEh(CLB.Items.Objects[i]).Title.Caption := EdtFieldTitle.Text;
  finally
    Grid.DataSource.DataSet.EnableControls;
  end;
end;

procedure TFrmColSelect.CmbFieldnameChange(Sender: TObject);
begin
  SaveFieldValue;
end;

procedure TFrmColSelect.EditSumValueChange(Sender: TObject);
begin
  SaveFieldValue;
end;

procedure TFrmColSelect.ChkReadOnlyClick(Sender: TObject);
begin
  SaveFieldValue;
  CLB.Checked[CLB.ItemIndex] := ChkVisiable.Checked;
end;

procedure TFrmColSelect.EdtFieldTitleChange(Sender: TObject);
begin
  SaveFieldValue;
end;

procedure TFrmColSelect.CheckBox2Click(Sender: TObject);
var
  i: Integer;
begin
  if Clb.ItemIndex >= 0 then
  for i := 0 to CLB.Items.Count - 1 do
  begin
    CLB.Checked[i] := CheckBox2.Checked;
    if CheckBox2.Checked then
      CheckBox2.Caption := '选择全部' else CheckBox2.Caption := '取消全部';
  end;
end;

procedure TFrmColSelect.frBtnNextClick(Sender: TObject);
var
  i: integer;
begin
  if CLB.ItemIndex < CLB.Count - 1 then
  begin
    i := clb.ItemIndex + 1;
    clb.Items.Move(clb.ItemIndex, i);
    CLB.ItemIndex := i;
    CLB.Selected[i];
  end;
end;

procedure TFrmColSelect.frBtnPrevClick(Sender: TObject);
var
  i: integer;
begin
  if CLB.ItemIndex > 0 then
  begin
    i := clb.ItemIndex - 1;
    clb.Items.Move(clb.ItemIndex, i);
    CLB.ItemIndex := i;
    CLB.Selected[i];
  end;
end;

procedure TFrmColSelect.btnLastClick(Sender: TObject);
begin
  if Clb.ItemIndex >= 0 then
  begin
    CLB.Items.Move(clb.ItemIndex, CLB.Count - 1);
    CLB.Selected[CLB.Count - 1] := True;
  end;
end;

procedure TFrmColSelect.btnNextClick(Sender: TObject);
var
  i: integer;
begin
  if CLB.ItemIndex < CLB.Count - 1 then
  begin
    i := clb.ItemIndex + 1;
    clb.Items.Move(clb.ItemIndex, i);
    CLB.ItemIndex := i;
    CLB.Selected[i];
  end;
end;

procedure TFrmColSelect.btnPriorClick(Sender: TObject);
var
  i: integer;
begin
  if CLB.ItemIndex > 0 then
  begin
    i := clb.ItemIndex - 1;
    clb.Items.Move(clb.ItemIndex, i);
    CLB.ItemIndex := i;
    CLB.Selected[i];
  end;
end;

procedure TFrmColSelect.btnFirstClick(Sender: TObject);
begin
  if Clb.ItemIndex >= 0 then
  begin
    clb.Items.Move(clb.ItemIndex, 0);
    CLB.Selected[0] := True;
  end;
end;

end.

