{*********************************************************************}
{                                                                     }
{     LodingForm  Create By cyq 2013年7月10日                         }
{                                                                     }
{     显示一个查找，更新，导入，导出等处理数据的加载窗口              }
{     调用：引用CYQCommonUnit                                         }
{           ShowLodingMessage(显示内容,颜色【可为空】)显示加载窗口    }
{           HideLodingMessage隐藏加载窗口                             }
{*********************************************************************}
unit LodingForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, RzPanel, RzBckgnd, RzAnimtr,
  ImgList, pngimage, System.ImageList;

type
  //显示进度时钟的线程
  TShowClockTime = Class(TThread)
  private
    BeginTime: TDateTime;
  protected
    procedure Execute; override;
  public
    constructor Create;
  end;

  TfrmLoding = class(TForm)
    pnlShowMessage: TRzPanel;
    imlLoading: TImageList;
    imgLoading: TImage;
    pnlMsg: TRzPanel;
    lblTime: TLabel;
    btnClose: TSpeedButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure pnlShowMessageMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }
    myClock: TShowClockTime;
  public
    { Public declarations }
    timeVisible: Boolean;
  end;

var
  frmLoding: TfrmLoding;

implementation

{$R *.dfm}

{TShowClockTime}

//显示时钟进度的线程
constructor TShowClockTime.Create;
begin
  BeginTime := Now;
  FreeOnTerminate := True;
  inherited Create(True);
end;

procedure TShowClockTime.Execute;
begin
  Repeat
    Sleep(1000);
    if Assigned(frmLoding) And Assigned(frmLoding.lblTime) And (Not Terminated) then
    begin
      //Synchronize
      if frmLoding.timeVisible and Assigned(frmLoding) then
      begin
        frmLoding.lblTime.Caption := FormatDateTime('h:mm:ss', Now - BeginTime);
        frmLoding.lblTime.Repaint;
      end;
    end;
  Until
    (Not Assigned(frmLoding)) or (Not Assigned(frmLoding.lblTime)) or Terminated;
end;

{TfrmLoding}

procedure TfrmLoding.FormCreate(Sender: TObject);
var
  img: TBitmap;
begin
  imlLoading.GetIcon(1,imgLoading.Picture.Icon);
//  img := TBitmap.Create;
//  try
//    imlLoading.GetBitmap(0, img);
//    if Assigned(img) then
//    begin
//      imgLoading.Picture.Bitmap.Assign(img);
//    end;
//    imgLoading.Repaint;
//  finally
//    FreeAndNil(img);
//  end;
  myClock := TShowClockTime.Create;
end;


procedure TfrmLoding.FormShow(Sender: TObject);
begin
  SetWindowPos(frmLoding.Handle,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE or SWP_NOSIZE );
  //激活时间进度线程
  myClock.Resume;
end;

procedure TfrmLoding.pnlShowMessageMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  ReleaseCapture;
  PostMessage(Self.Handle, WM_SYSCOMMAND, SC_MOVE + 1, 0);
end;

procedure TfrmLoding.FormDestroy(Sender: TObject);
begin
  myClock.Terminate;    //终止时间进度线程
end;

procedure TfrmLoding.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
  frmLoding := nil;
end;

procedure TfrmLoding.btnCloseClick(Sender: TObject);
begin
  Close;
end;

end.
