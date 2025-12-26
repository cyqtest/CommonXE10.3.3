(*
 Delphi版雪花算法
 作者：不得闲 QQ：75492895
 用于生成Int64位的唯一值ID，WorkerID用于区分工作站，
 ID会随着时间增加位数，每毫秒可生成4096个ID

 用法：
 创建全局变量：snow: TDxSnowflake;
 创建对象：snow := TDxSnowflake.Create; // 不要忘了在退出时释放snow.Free;
 调用：
 snow.WorkerID:=100;
 mmo1.Lines.Add( FormatFloat('#0',snow.Generate));
*)
unit uDxSnowflake;

interface

uses System.SysUtils, System.SyncObjs, System.Generics.Collections,
  System.DateUtils;

type
  TWorkerID = 0 .. 1023;

  TDxSnowflake = class
  private
    FWorkerID: TWorkerID;
    FLocker: TCriticalSection;
    fTime: Int64;
    fstep: Int64;
  public
    constructor Create;
    destructor Destroy; override;
    property WorkerID: TWorkerID read FWorkerID write FWorkerID;
    function getNewID: Int64;
  end;

implementation

const
  Epoch: Int64 = 1539615188000; // 北京时间2018-10-15号
  // 工作站的节点位数
  WorkerNodeBits: Byte = 10;
  // 序列号的节点数
  StepBits: Byte = 12;
  timeShift: Byte = 22;
  nodeShift: Byte = 12;

var
  WorkerNodeMax: Int64;
  nodeMask: Int64;

  stepMask: Int64;

procedure InitNodeInfo;
begin
  WorkerNodeMax := -1 xor (-1 shl WorkerNodeBits);
  nodeMask := WorkerNodeMax shl StepBits;
  stepMask := -1 xor (-1 shl StepBits);
end;
{ TDxSnowflake }

constructor TDxSnowflake.Create;
begin
  FLocker := TCriticalSection.Create;
end;

destructor TDxSnowflake.Destroy;
begin
  FLocker.Free;
  inherited;
end;

function TDxSnowflake.getNewID: Int64;
var
  curtime: Int64;
begin
  FLocker.Acquire;
  try
    curtime := DateTimeToUnix(Now) * 1000;
    if curtime = fTime then
    begin
      fstep := (fstep + 1) and stepMask;
      if fstep = 0 then
      begin
        while curtime <= fTime do
          curtime := DateTimeToUnix(Now) * 1000;
      end;
    end
    else
      fstep := 0;
    fTime := curtime;
    Result := (curtime - Epoch) shl timeShift or
      FWorkerID shl nodeShift or fstep;
  finally
    FLocker.Release;
  end;
end;

initialization

InitNodeInfo;

end.
