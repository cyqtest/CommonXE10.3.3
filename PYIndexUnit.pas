unit PYIndexUnit;

{*******************************************************}
{              汉字-->拼音声母处理 v4.0                 }
{           Developed by cyw(QQ: 26890954)              }
{           2013.9.13 All Right Reserved.               }
{*******************************************************}


{
一、特点：
1、支持多音字，支持所有中国汉字(共20902个)，包括难字、生僻字、广东白话字、繁体字等所有计算机能显示的全部汉字。
   例如：镕、啱、揾、叻、嘅、咁、門、長、發財、車、冇乜嘢。
2、同时支持Unicode和ANSI字符集。
3、支持Delphi6至DelphiXE系列的所有Delphi版本(最后一次测试是XE4)。

二、实现原理(Unicode)：
1、根据汉字的编码规则，汉字在Unicode的编码范围为$4E00至$9FA5(即:19968至40869, 共20902个汉字)。
2、对中国汉字字库中的每个汉字，扫描从互联网下载的较完整的<GBK汉字各种编码汇总.xls>拼音字库，找到每个汉字的全拼拼音
   (包括多音，一个汉字最多4种读音)，并获取每个读音的声母(首字母), 生成汉字声母字典文件。
3、把声母字典文件中每个汉字的声母，按汉字字库的Unicode编码规则全部串在一起，由于一个汉字最多3种读音(有4种以上读音的
   仅有"竓qfzygs, 竕sfzyg, 竡ygsb"三个生僻字, 仅取前三个读音)，为每个汉字预留3个字节空间，生成Unicode汉字声母表,
   如本单元所定义的PartUnicode常量。
4、查找原理：要查找一个汉字在声母表中的对应位置，应使用以下公式计算得：
       汉字声母索引 = (汉字Unicode编码 - 19968) * 3 + 1;  //其中19968是指Unicode字符集中的第一个汉字位置; 3表示每个汉字预留的3个字节空间

三、实现原理(ANSI)：
1、根据汉字的编码规则，第一字节为区码(纵向)，第二字节为位码（横向），
   第一字节的范围是129-254，共126种，其中161-169段存放各种标点符号，254段为无效汉字，真正有效段为129-160,170-253，共116种。
   第二字节的范围是64-254，其中127段为分隔符，不存在任何内容，真正有效段为64-126,128-254，共190种。
   排除无效字符及标点符号后，有效的汉字库分为三部分：
   1)第一部分(区:129-160 位:64-254[除127]) 	                    32*190 =6080
   2)第二部分(区:170-253 位:64-160[除127], 区253位156-160留空)  84*96-5=8059
   3)第三部分(区:176-247 位:161-254, 区215位250-254留空)        72*94-5=6763
                                                                  合计: 20902个汉字
2、按以上规则可生成三个部分的中国汉字字库，每个区码一行，每个位码一列。
3、同二(2)。
4、同二(3)，如本单元所定义的Part1..Part3常量。
5、查找原理：要查找一个汉字在声母表中的对应位置，应使用以下公式计算得：
       汉字声母索引 = ((汉字区码 - 段落起始区码) * 位码长度 + 汉字位码 - 段落起始位码) * 3 + 1;  //其中3表示每个汉字预留的3个字节空间
   如汉字“啊”，其区位码分别是(176,161)，根据上述第1点所定义，属于第三部分汉字，理应在常量Part3中查找，
   由于第三部分的位码范围是161-254，横向长度是94，即一行(一个区)有94个汉字，所以“啊”在Part3中的位置应该是：
   ((176-176) * 94 + 161 - 161) * 3 + 1 = 1，即Part3[1] = 'a';

四、更新历史：
1、10.9.8 由于全拼输入法的拼音库太过全面，有部分多音字的第一个拼音却不是常用的拼音，导致汉字转声母有时不准确，
          决定抛弃全拼输入法的拼音库，改为采用从互联网下载的较完整的<GBK汉字各种编码汇总.xls>拼音字库，重新生成汉字声母表。
2、10.12.27 上一次更新时是在Excel中编辑<GBK汉字各种编码汇总.xls>文件的，在Excel内替换右括号")"为空时，"Jun1)"替换后会变成"1-Jun"，
             Excel的奇怪现象，导致"军"字的声母变成数字"1"，类似的情况还有很多。本次更新全面修正此类问题，并且经检测发现99.9%以上的
             多音字最多只有三个读音，仅"竓qfzygs, 竕sfzyg, 竡ygsb"三个生僻字有4种以上读音, 故声母表全部由4个字节改为预留3个字节，
             上述三个僻字仅截取前三个读音的声母。
3、12.2.14 之前的版本按常用程度分为五个级别，但却漏了(区:170-175 位:64-160)一段汉字，合计漏了6*96=576个汉字未处理。今改为按汉字编码
           分布规则，划分为三个部分（三段），共计20902个汉字，是迄今为止史上最完整的汉字字库。据此生成的Part1、Part2、Part3三个声母
           常量也是最完整的汉字声母表。另外，从该版本起，开始支持Delphi2007-DelphiXE系列。
4、13.9.13 针对Unicode的编码规则，生成全新的Unicode原生汉字声母表，放弃以前把String转换为AnsiString的方法，由于少了转换步骤，所以理论上
           在Delphi2009(或更高版本)时算法速度更优更快。同时使用编译指令兼容Delphi6-2007的非UniCode开发环境，这部分保留以前的算法。
}

interface

uses
  SysUtils;

//--------------------------- 声明段 -------------------------------------------

//获取指定汉字的拼音索引字母，如：汉 = h。多音字返回多个声母，如：行 = hx
function GetPYIndex(const OneWord: String; const OnlyOne: Boolean = False): String;
//获取指定字符串的拼音索引字母（多音字总是取第一个音的声母)。如：中山大学 = zsdx
function GetStrPYIndex(const Str: String): String;
//获取字符串Str中与拼音声母字符串PYStr匹配的起始位置。如：PosPY('zsdx'，'我来自中山大学')，在Unicode环境中返回4，非Unicode环境返回7
function PosPY(const PYStr, Str: String): Integer;
//判断字符串Str的拼音声母是否全部与PYStr相同。如：SamePY('zsdx', '中山大学') = True
function SamePY(const PYStr, Str: String): Boolean;
//获取字符串Str中拼音声母串与字符串SubStr拼音声母串匹配的起始位置。如：PosTextPY('纵声大笑', '我来自中山大学')，在Unicode环境中返回4，非Unicode环境返回7
function PosTextPY(const SubStr, Str: String): Integer;
//判断字符串Str的拼音声母是否全部与字符串SubStr的拼音声母相同。如：SameTextPY('纵声大笑', '中山大学') = True
function SameTextPY(const SubStr, Str: String): Boolean;

//---------------------------- 实现段 ------------------------------------------
implementation

const
  MaxPYLength = 3; //最大的拼音声母个数

{$IFDEF UNICODE}
  //Unicode汉字声母表 $4E00至$9FA5(即:19968至40869, 共20902个)
  PartUnicode =
     'y  dz k  q  s  x  n  wm z  s  s  x  j  bf y  m  g  c  c  z  qj p  s  s  q  b  y  c  d  s  c  d  q  l  d  y  l  y  b  s  s  j  g  y  qp z  j  j  f  g  c  c  l  z  zd n  w  d  w  z  jd l  j  p  f  y  ya n  n  j  j  t  my y  n  z  w  z  h  f  ly z  p  p  '
    +'q  h  g  cs cs y  y  n  mn j  q  y  x  x  g  d  n  n  s  n  s  j  n  j  n  s  n  n  m  l  n  r  x  y  f  s  n  gq n  n  n  n  qg z  gjqg  l  l  y  j  l  n  y  z  s  s  e  c  y  k  y  y  h  q  w  j  s  s  g  g  y  x  y  q  y  jq t  w  k  t  j  h  y  c  '
    +'h  m  n  x  j  t  l  h  j  y  q  b  y  x  d  l  d  w  r  r  j  n  w  y  s  r  l  d  z  j  p  cq b  z  j  j  b  r  c  f  s  l  n  c  z  s  t  z  f  x  x  c  h  t  r  q  g  gy d  d  l  y  c  c  s  s  y  m  m  r  j  c  y  q  z  p  w  w  j  j  y  f  c  r  '
    +'w  f  d  f  z  q  p  y  d  d  w  y  x  k  y  j  a  w  j  f  f  x  j  b  c  f  t  z  y  h  hk y  cz y  s  w  cz c  y  x  s  ct l  c  x  x  w  z  c  x  n  b  g  n  n  x  b  x  l  z  s  q  sc b  s  jgqp  y  s  a  z  dt h  m  d  z  b  q  b  s  c  w  d  z  '
    +'z  y  y  tb z  h  b  t  s  yt yd f  z  g  n  t  n  xs j  y  w  q  n  k  n  p  h  h  l  x  g  y  b  f  m  j  ne b  j  h  h  g  q  t  j  c  y  s  x  s  t  k  z  gk l  y  c  k  g  l  y  s  m  z  x  y  a  l  m  e  l  dt c  c  x  g  z  y  r  j  x  j  z  l  '
    +'n  jy z  cz q  k  c  n  n  j  w  h  j  c  z  c  c  q  l  j  s  t  s  t  b  n  h  bp t  y  x  c  e  q  x  k  k  w  jz y  f  l  z  q  l  y  h  j  x  s  p  s  f  x  l  m  p  b  ys sq x  xs x  y  t  c  c  n  y  l  l  l  s  j  x  f  h  j  x  p  j  b  tc f  '
    +'f  y  a  b  yz x  b  c  c  z  b  z  y  c  l  w  l  c  z  g  g  b  t  s  s  m  d  t  j  c  x  p  tc h  y  q  t  g  jl j  x  c  j  f  z  k  j  z  j  q  n  l  z  w  l  s  l  h  d  z  b  w  j  n  c  j  z  y  z  s  q  n  y  c  q  y  n  z  c  j  j  w  y  b  '
    +'r  t  w  p  y  f  t  w  e  xj c  s  k  d  z  c  t  b  y  h  y  z  q  y  y  j  x  z  c  f  b  z  z  m  j  y  x  x  sc d  cz z  o  t  t  b  z  l  j  w  f  c  kg s  c  s  x  f  y  r  l  r  y  g  m  b  d  t  h  j  x  s  q  j  c  c  s  b  x  y  y  t  s  w  '
    +'f  b  j  d  z  t  n  b  c  n  z  l  c  y  z  z  p  s  a  cz y  z  z  s  q  q  c  s  h  z  q  y  d  x  l  b  p  j  l  l  m  q  x  q  y  d  z  x  s  q  j  t  z  p  x  l  c  g  l  q  d  c  w  z  f  h  c  t  d  jy j  s  f  x  j  e  j  j  tz l  b  g  x  s  '
    +'x  j  m  y  j  j  q  p  f  z  a  s  y  j  n  s  y  d  j  x  k  j  c  d  j  s  z  n  b  a  r  t  c  c  l  n  j  q  m  w  n  q  n  c  l  l  l  k  b  d  b  z  z  s  y  h  q  c  l  t  w  l  c  c  r  s  h  l  l  z  n  t  y  l  n  er w  y  z  y  x  c  z  x  '
    +'x  g  d  k  d  mw t  cz e  d  e  j  t  s  y  y  s  sk d  q  d  f  m  x  d  bk j  l  k  r  w  n  q  l  ys b  g  l  x  n  l  g  t  g  x  b  qj j  d  zc n  y  j  s  j  y  j  c  j  m  r  n  y  m  g  r  c  j  c  z  g  j  m  z  m  g  x  m  m  r  y  x  k  j  '
    +'n  y  m  s  g  m  z  z  y  m  k  n  f  x  m  b  d  t  g  fp b  h  c  j  h  k  y  l  p  f  m  d  x  l  x  j  j  s  m  s  q  g  z  s  j  l  qj d  l  d  g  j  y  c  a  l  c  m  z  c  s  d  j  l  l  n  x  d  j  f  f  f  f  j  c  n  f  n  n  f  f  p  f  k  '
    +'h  k  g  d  p  qk x  k  tg aw c  j  d  h  h  z  d  d  d  r  r  c  f  q  y  j  k  q  c  c  w  j  d  x  h  w  j  l  y  l  l  z  g  c  f  c  q  j  s  m  l  z  p  b  j  j  pb l  s  b  c  j  g  g  d  c  k  k  d  e  z  s  qx sc c  k  j  g  c  g  k  d  j  t  '
    +'j  l  l  z  y  c  x  k  l  q  c  c  g  j  c  l  t  f  p  c  q  c  z  g  w  b  j  d  q  ys z  j  j  b  y  j  h  s  j  d  d  w  g  f  s  j  g  z  k  c  j  c  tz l  l  f  s  p  k  jc g  q  j  h  z  z  l  j  p  l  g  j  g  j  j  t  h  j  j  y  j  z  c  c  '
    +'m  l  z  l  y  q  b  g  j  w  m  l  j  k  x  z  d  z  n  j  q  s  y  z  m  l  j  l  l  j  k  y  w  x  m  k  j  l  h  s  k  j  h  b  m  c  l  y  y  m  k  x  j  q  l  b  m  l  l  k  m  d  x  x  k  w  y  x  w  s  l  m  l  p  s  j  q  q  jc q  x  y  q  f  '
    +'j  t  j  d  x  m  x  x  l  l  c  r  q  b  sb y  j  b  g  w  y  n  n  g  g  b  c  n  x  p  j  t  g  p  a  p  f  g  d  j  q  b  h  b  n  cs fx j  y  z  j  k  j  k  h  x  q  f  g  q  cz k  f  h  y  g  k  h  d  k  l  l  s  d  j  q  x  py qo y  a  y  b  n  '
    +'qo s  x  q  n  s  z  s  w  h  b  s  x  w  h  x  w  b  zc z  x  dcsm  n  dc j  b  sl b  k  b  b  z  kq l  y  l  x  g  w  x  j  j  w  ay q  z  m  y  w  s  j  q  l  s  jq x  x  j  q  w  j  e  q  x  n  chah  e  t  l  z  ac l  y  y  y  s  z  z  p  n  k  y  '
    +'z  cs p  t  l  s  h  t  z  c  f  y  cs y  x  y  l  j  xs d  c  j  qj a  g  y  s  l  ca l  l  y  y  y  s  s  l  q  q  q  n  l  d  x  z  s  cs cs s  cs a  d  y  c  j  y  s  f  s  g  b  f  r  s  s  z  q  s  b  x  j  p  s  j  y  s  d  r  c  k  g  jg l  g  '
    +'td k  z  j  zs b  d  k  t  c  s  y  q  p  yx h  s  t  c  l  d  j  n  h  m  yx m  cj g  x  y  z  hg j  d  c  t  m  h  l  t  x  z  xh y  l  a  m  o  x  y  j  c  l  t  y  f  bp q  q  j  p  fp b  d  f  e  h  ty kh s  q  h  z  y  w  w  c  n  c  x  c  d  w  '
    +'h  oh w  g  y  j  l  e  g  m  da q  c  w  g  f  j  h  c  s  n  t  m  y  dt o  l  b  y  g  n  q  w  e  s  q  p  w  n  m  l  r  y  d  z  s  z  z  l  y  q  p  z  gw cz w  x  hkan  x  p  y  x  s  h  m  d  q  jz g  z  t  d  p  p  b  f  bf h  z  h  h  j  y  '
    +'f  d  z  w  k  g  kg z  b  l  d  n  z  s  x  h  q  e  e  g  z  xd y  l  z  m  m  z  y  j  z  kgls  z  x  kh h  k  h  t  x  e  x  x  gk y  l  y  a  p  s  t  h  xh d  w  hk z  y  d  p  x  a  g  k  y  d  x  b  hy n  hy n  k  d  n  j  n  m  y  h  y  l  p  '
    +'m  g  oewc  s  l  n  z  h  k  x  x  l  b  z  z  l  b  m  l  s  f  b  h  h  g  s  g  y  y  g  g  b  h  s  c  y  a  j  t  x  g  l  x  t  z  m  c  wn z  y  d  q  d  q  m  n  g  d  n  l  l  s  z  h  n  g  j  z  w  f  y  hx q  s  w  s  c  e  l  q  a  j  y  '
    +'n  y  t  l  s  x  t  h  a  z  n  k  z  z  s  d  h  l  a  x  x  t  w  w  c  j  h  q  q  t  d  d  w  z  b  c  c  h  y  q  z  f  l  x  p  s  l  z  q  g  p  z  s  z  n  g  l  y  d  q  tc b  d  l  x  n  t  c  t  a  j  d  k  y  w  n  s  y  z  l  j  h  h  dz '
    +'z  c  k  rn y  y  z  y  wo m  h  y  c  h  h  h  x  h  j  k  z  w  s  x  h  d  n  x  l  y  s  c  q  y  dcsp  s  l  y  zc w  m  y  p  n  k  x  y  j  l  k  c  h  t  y  h  as x  q  s  y  s  h  x  ay s  m  c  h  k  dt s  c  rn s  g  j  p  w  q  s  g  z  j  '
    +'l  w  w  s  c  h  s  jl h  s  q  n  h  n  s  n  g  n  d  a  n  t  b  a  a  l  c  z  m  s  sz t  d  q  j  c  j  k  t  s  c  j  n  x  p  l  g  gj x  h  h  g  o  x  z  c  xs p  d  m  m  h  l  d  g  t  y  b  y  n  j  m  x  hy m  r  zc p  l  x  j  z  cz k  '
    +'z  x  s  h  fm l  q  x  x  c  d  hm x  w  z  p  c  k  c  z  c  d  y  t  c  j  y  x  q  h  l  xs d  h  y  p  j  q  x  n  l  s  y  y  d  z  o  z  j  n  hy h  q  e  z  y  s  j  y  a  y  xj k  y  p  d  g  x  d  d  n  s  p  p  y  z  n  d  h  t  ho r  h  xh '
    +'y  d  p  c  j  j  h  t  c  n  n  c  t  l  h  b  y  n  y  h  m  h  z  l  l  n  n  xa m  y  l  l  l  m  d  c  p  p  x  h  m  x  d  k  y  c  y  r  d  l  t  x  j  c  h  h  z  n  xa c  l  c  c  l  y  l  n  z  s  x  n  j  z  z  l  n  n  n  n  w  h  y  q  s  '
    +'n  jn h  x  y  n  t  t  dt k  y  j  p  y  c  h  h  y  e  g  k  c  w  t  w  l  g  j  r  l  g  g  t  g  t  y  g  y  h  p  y  h  y  l  qj y  q  g  c  w  y  qj k  p  y  y  e  tsgt  t  t  l  h  y  yh l  l  t  y  t  t  s  y  l  n  y  z  wx g  y  w  g  p  y  '
    +'d  q  q  z  zs d  q  n  n  k  c  qy n  m  j  j  z  z  b  x  t  q  f  jy k  d  f  f  b  t  k  hp z  k  b  x  d  j  j  k  d  j  j  t  l  b  w  f  z  p  p  t  k  q  t  z  t  g  p  d  w  n  t  p  j  y  f  a  l  q  m  k  g  x  b  cd c  l  z  f  h  z  c  l  '
    +'l  l  l  a  n  p  n  x  td j  k  l  c  c  l  g  y  h  d  z  f  g  y  d  d  g  c  y  y  f  g  y  d  x  k  s  s  e  n  d  h  y  k  d  n  d  k  n  n  a  x  x  hy b  p  b  y  y  h  x  c  c  g  a  p  f  q  y  j  j  d  m  l  x  c  s  j  z  l  l  pb c  n  b  '
    +'s  x  g  j  y  n  d  y  b  wy j  s  pb c  w  j  l  z  k  z  d  d  t  a  c  s  b  k  z  d  y  z  y  p  j  z  q  s  j  n  kj k  t  k  n  j  dz j  g  y  ew pb g  t  l  n  y  q  n  a  c  dh n  t  c  y  h  b  l  g  d  z  h  bp b  y  dt m  j  r  ea g  k  z  '
    +'y  h  e  y  y  b  j  m  c  d  t  a  f  z  j  z  h  g  c  j  n  l  g  h  l  d  w  x  j  j  k  y  t  c  y  k  s  s  s  m  td w  c  t  t  q  z  l  p  b  s  z  d  t  w  c  x  g  z  a  g  y  k  t  y  w  x  l  n  l  c  p  b  c  l  l  o  q  m  m  z  s  s  l  '
    +'c  m  b  j  c  s  d  z  k  y  d  c  z  j  g  q  j  d  s  m  c  y  td z  q  q  l  n  z  q  z  x  s  s  b  p  k  d  f  q  m  d  d  z  z  s  dh d  t  dy m  f  h  t  d  y  c  n  a  q  j  q  k  y  p  b  d  j  y  y  x  t  l  j  hp d  r  q  x  l  m  h  k  y  '
    +'d  h  r  n  l  k  l  y  t  w  hp l  l  r  l  l  r  c  x  y  l  b  n  s  r  n  z  z  s  y  m  kq z  z  h  h  k  y  h  x  k  s  m  z  s  y  z  g  c  x  f  b  n  b  s  q  l  f  z  xj x  n  n  x  k  x  w  y  m  s  d  d  y  q  n  g  g  q  m  m  y  h  c  d  '
    +'z  t  t  f  g  y  y  hb g  s  bt t  t  y  b  y  k  jg d  n  k  y  j  b  em l  h  d  y  p  j  qj n  f  x  f  d  n  k  z  h  qx k  s  b  y  j  t  z  b  x  h  f  d  s  b  d  a  s  w  p  a  w  a  j  l  d  y  j  s  f  h  b  l  c  nr n  d  n  q  j  t  j  n  '
    +'c  h  x  f  j  s  r  f  w  h  z  f  m  d  r  f  j  y  x  w  z  p  d  j  k  z  y  j  y  m  p  c  y  z  n  y  n  x  f  b  y  t  f  y  f  w  y  g  d  b  n  z  z  z  d  n  y  t  x  z  e  m  m  q  b  s  q  e  h  x  f  z  n  b  m  f  l  z  z  s  r  s  y  q  '
    +'j  g  s  x  w  z  j  s  p  r  y  t  j  s  j  g  s  k  j  j  g  lm j  j  y  n  z  j  j  x  h  g  j  k  y  m  l  p  y  y  y  c  x  y  c  g  q  z  s  w  h  w  l  y  r  j  l  p  x  s  l  c  x  m  n  s  m  w  k  l  c  d  n  k  n  y  n  p  s  j  sczz  h  d  '
    +'z  e  p  t  x  mw w  y  w  x  y  y  s  y  w  l  x  j  q  c  q  x  z  d  c  l  a  e  e  l  m  c  p  j  p  c  l  w  b  x  s  q  h  f  w  r  t  f  n  j  t  n  q  j  h  j  q  d  x  h  w  l  b  y  cz c  f  j  l  ya l  k  y  y  n  l  d  x  n  h  y  c  s  t  '
    +'y  y  w  n  c  j  t  x  y  w  t  r  m  d  r  q  n  w  q  c  m  f  j  d  x  z  m  h  m  a  y  x  n  w  m  y  z  q  ts x  t  l  m  r  s  p  w  w  j  h  a  n  b  x  t  g  z  y  p  x  y  y  r  r  c  l  m  p  a  m  g  k  q  j  s  z  y  c  y  m  y  j  s  n  '
    +'x  ty p  l  n  b  a  p  p  y  p  y  l  x  m  y  z  k  y  n  l  d  g  y  j  z  c  c  hr n  l  m  z  h  h  a  n  q  n  b  g  w  q  t  z  m  x  x  m  l  l  h  g  d  z  x  n  h  x  h  r  x  y  c  j  m  f  f  x  y  w  c  f  s  b  s  s  q  l  hqxn  n  d  y  '
    +'c  a  n  n  m  t  c  j  c  y  p  n  x  n  y  t  y  c  n  n  y  m  n  m  s  x  n  d  l  y  l  y  s  l  j  n  l  x  y  s  s  q  m  l  l  y  z  l  z  j  j  j  k  y  z  z  c  s  f  b  z  x  x  m  s  t  b  j  g  n  x  n  c  h  l  s  n  m  c  j  s  c  y  z  '
    +'n  f  z  l  x  b  r  n  n  n  y  l  m  n  r  t  g  z  q  y  s  a  t  s  w  r  y  h  y  j  z  m  g  d  h  z  g  z  d  wy y  b  s  s  c  s  k  x  s  y  h  y  t  s  x  g  c  q  g  x  z  z  b  h  y  x  jg s  c  r  h  m  k  k  b  sx cs z  j  y  j  y  m  k  '
    +'q  h  z  j  f  n  b  h  m  q  h  y  s  n  j  n  z  y  b  k  n  q  m  c  j  g  q  h  w  l  s  n  z  s  w  x  k  h  l  j  h  y  y  b  q  c  b  f  c  d  s  x  d  l  d  s  p  f  z  f  sy k  jq jq z  wy z  x  s  d  d  x  j  s  e  e  e  g  j  s  c  s  s  y  '
    +'g  c  l  x  x  n  wy wy y  l  l  y  mp w  w  w  g  y  d  k  z  j  g  g  g  g  g  g  s  y  c  k  n  j  wy ns j  p  c  x  b  j  j  t  q  t  j  w  d  s  s  pb j  x  c  x  n  z  x  n  m  e  l  p  t  f  sz x  t  l  l  x  c  l  j  x  j  j  l  j  sz x  c  tz '
    +'n  s  w  x  l  e  n  n  l  yg q  r  w  h  s  y  c  s  q  n  y  b  y  a  y  w  j  e  j  q  f  w  q  c  q  q  c  j  q  g  x  a  l  d  b  z  z  y  j  g  k  g  x  b  l  t  q  y  f  x  j  l  t  p  y  d  k  y  q  h  p  m  a  t  l  c  n  d  n  k  x  m  t  x  '
    +'y  n  h  k  l  e  f  x  d  l  l  e  g  q  td y  m  s  a  w  h  zs m  l  j  t  k  y  n  x  l  y  jq z  l  j  e  e  y  y  b  q  q  f  f  n  l  y  x  h  d  s  c  t  g  j  h  x  y  wy l  k  l  l  x  q  k  c  c  t  n  h  j  l  q  m  k  k  z  g  c  y  ya g  '
    +'l  l  l  j  d  c  g  y  d  h  z  w  yx p  y  s  j  b  z  j  d  z  g  y  z  z  h  y  w  y  f  q  d  t  y  z  s  z  y  e  z  k  l  y  m  g  j  j  h  t  s  m  qk w  y  z  l  jh y  y  w  z  c  s  r  k  q  y  q  l  t  d  x  w  c  d  r  j  a  l  w  s  q  z  '
    +'w  b  d  c  q  y  n  c  j  n  n  s  z  j  l  n  c  d  c  d  td l  z  z  z  a  c  q  q  z  z  d  d  x  y  b  l  x  c  b  q  j  y  l  z  l  l  l  j  d  d  z  jq g  y  q  y  j  z  y  x  n  y  y  y  e  x  j  x  k  s  d  a  z  n  y  r  d  l  z  y  y  y  n  '
    +'j  l  sx l  l  d  y  x  j  c  y  k  y  w  n  q  c  c  l  d  d  n  y  y  y  n  y  c  g  c  z  h  j  x  c  c  l  g  z  q  j  g  n  w  n  n  c  q  q  j  y  s  b  z  z  xh y  j  x  j  n  x  j  f  z  b  s  b  d  s  f  n  s  f  p  z  x  h  d  w  z  t  d  m  '
    +'p  p  t  f  l  z  z  b  z  d  m  y  y  p  q  j  r  s  d  z  s  q  z  s  q  x  b  d  g  cd p  z  s  w  d  w  c  s  q  z  g  m  d  h  z  x  m  w  w  f  y  b  p  n  g  p  h  dt m  j  t  h  z  s  m  m  b  g  z  m  b  z  j  c  f  z  h  f  cz b  b  n  m  q  '
    +'d  f  m  b  cd m  c  j  x  l  g  p  n  j  b  b  x  g  y  h  y  y  j  ga p  t  z  ga z  m  q  b  q  d  c  g  y  b  j  x  l  w  n  k  y  d  p  d  y  m  g  c  f  t  p  f  x  y  z  t  z  x  d  z  x  t  g  k  pm t  y  b  b  c  l  b  j  a  s  k  y  t  s  s  '
    +'q  y  y  m  s  c  x  f  j  h  h  l  xs l  l  s  j  p  q  j  j  qj a  k  l  y  l  d  l  y  c  c  t  s  x  m  ca w  f  ga n  g  b  q  x  l  l  l  l  n  y  x  t  y  l  t  y  x  y  t  d  p  j  h  n  h  g  n  k  b  y  q  nl f  j  y  y  z  b  y  y  e  s  s  '
    +'e  s  s  g  d  y  h  f  h  w  t  c  q  b  s  d  z  j  t  f  d  m  x  h  c  n  j  z  y  m  q  w  s  r  x  j  d  z  j  qj b  dt qj b  b  s  dt j  g  n  f  b  k  n  b  x  dt k  qj h  m  k  w  j  j  j  g  d  l  l  t  h  z  h  h  y  y  y  y  h  h  s  x  z  '
    +'t  y  y  y  c  c  b  d  b  p  y  p  z  y  c  c  z  t  j  pf z  y  w  c  b  d  l  f  w  z  c  w  j  d  x  x  h  y  h  l  h  w  c  z  x  j  t  c  n  l  c  d  p  x  n  q  c  z  cz z  l  y  x  j  j  cs j  b  h  f  x  w  p  y  w  x  z  p  c  d  z  z  b  d  '
    +'c  c  j  w  j  h  m  l  x  b  q  x  x  b  y  l  r  d  d  g  j  r  r  c  t  t  t  g  q  d  c  z  w  m  x  f  y  t  m  m  z  c  w  j  w  x  y  y  w  z  zs k  y  b  z  c  c  c  t  t  q  n  h  x  n  w  x  x  k  h  k  f  h  t  s  w  o  c  c  j  y  b  c  m  '
    +'p  z  z  y  k  b  n  n  z  p  b  t  h  h  j  d  l  s  z  d  d  y  t  y  f  j  p  x  y  n  g  f  x  b  y  q  x  z  b  h  x  c  p  x  x  t  n  z  d  n  n  y  c  nr x  s  x  l  h  k  m  z  x  l  t  h  d  h  k  g  h  x  x  s  s  h  q  y  h  h  c  j  y  x  '
    +'g  l  h  z  x  c  x  n  h  e  kq dt t  g  q  x  q  y  p  k  d  h  ew n  t  y  k  c  n  y  m  y  y  j  m  k  q  y  y  y  j  x  zq l  t  h  h  q  t  b  y  q  h  x  b  m  y  h  s  q  c  kl w  w  y  l  l  h  c  y  y  l  n  n  ew q  x  q  w  m  c  f  b  d  '
    +'c  c  m  l  j  g  g  x  d  q  k  t  l  x  k  k  n  q  c  d  g  c  j  w  y  j  j  l  y  h  h  q  y  t  t  n  w  c  h  h  x  c  x  w  h  ew r  z  j  y  d  j  c  c  d  b  q  c  d  g  d  n  y  x  z  d  h  c  q  r  x  c  b  h  z  t  q  c  b  x  w  g  q  w  '
    +'y  y  b  x  h  m  b  y  m  y  kq d  y  e  c  m  q  k  y  a  q  y  n  g  y  z  s  l  f  n  k  k  q  g  y  s  s  q  y  s  h  n  g  j  c  t  x  k  z  y  c  s  s  b  k  y  x  h  y  y  l  s  t  y  c  x  q  t  h  y  s  m  n  s  c  p  m  m  g  c  c  c  c  c  '
    +'m  nt z  t  a  s  m  g  q  z  j  h  k  l  o  s  j  y  l  s  w  t  m  q  z  y  q  k  d  z  l  j  q  q  y  p  l  z  y  c  z  t  c  q  q  p  b  b  c  j  z  c  l  p  k  h  q  c  y  y  x  x  d  t  d  d  d  s  j  c  x  f  f  l  l  c  h  q  x  m  j  l  w  c  '
    +'j  c  x  t  s  p  y  c  x  n  d  t  j  s  h  j  w  h  d  q  q  q  c  k  x  y  a  m  y  l  s  j  h  m  l  a  l  y  g  x  c  y  y  d  m  a  m  d  q  m  l  m  c  z  n  n  y  y  b  z  x  k  y  f  l  m  c  n  c  m  l  h  x  r  c  j  j  h  sz y  l  n  m  t  '
    +'j  g  gz z  g  y  w  j  xq s  r  xh c  w  j  g  j  q  h  q  z  d  q  j  d  c  j  j  sc k  j  k  g  d  z  cq g  j  j  y  j  y  l  xh z  x  xh c  d  q  h  h  h  e  s  l  m  h  l  f  s  bp d  j  s  y  y  s  h  f  y  s  s  c  z  q  l  p  bp d  r  fb n  z  '
    +'t  z  d  kg y  k  h  s  c  c  g  k  w  t  q  z  c  k  m  s  y  n  b  c  r  x  q  bp j  y  f  a  x  p  z  z  e  d  z  c  j  y  k  b  c  j  w  h  y  j  b  q  z  z  y  w  n  y  s  z  p  t  d  k  zs p  f  p  b  a  z  t  k  l  q  n  h  b  b  z  p  t  p  p  '
    +'t  y  z  z  y  b  h  n  y  d  c  p  z  m  m  cs y  c  q  m  c  j  fb z  z  d  c  m  n  l  f  p  b  p  l  n  g  q  j  t  b  t  t  anyj  z  p  z  b  b  d  n  j  k  l  j  q  y  l  n  b  z  q  h  kg s  j  z  n  g  g  q  s  t  z  k  c  x  c  h  p  zy s  n  '
    +'b  c  g  z  k  d  d  z  q  a  nl z  g  j  k  d  rn t  l  z  l  d  w  j  n  j  z  l  y  wz t  xj n  d  j  z  j  h  x  n  a  t  n  c  b  g  t  z  c  s  s  k  m  ln j  p  j  y  t  s  n  w  xj c  f  j  w  j  j  t  k  h  t  z  p  l  b  h  s  n  j  s  s  y  '
    +'j  b  h  b  j  y  z  l  s  t  l  s  b  j  h  d  n  w  q  p  s  l  m  m  f  b  j  d  w  a  j  y  z  c  c  j  t  b  n  n  n  z  w  xq x  c  d  s  l  q  g  d  s  d  p  d  z  g  j  t  q  q  p  s  q  l  y  y  j  z  l  g  y  h  s  d  l  c  t  c  b  j  t  k  '
    +'t  y  c  z  j  t  q  k  b  sp j  l  g  n  n  z  ds n  cs s  g  p  y  n  j  z  j  j  y  y  k  n  h  r  p  w  s  z  x  m  tdsn  c  s  z  z  y  s  h  b  y  h  y  z  a  xj y  w  k  c  j  t  l  l  c  k  j  j  t  j  h  g  c  s  sd x  y  q  y  c  z  b  y  n  '
    +'n  l  w  q  c  g  l  z  g  j  g  q  y  q  c  c  z  s  s  b  c  r  bp c  s  k  y  d  z  n  x  j  s  q  g  x  s  s  j  m  e  c  n  s  t  j  t  p  b  d  l  t  h  z  w  x  q  w  qc c  z  e  x  n  q  c  z  g  w  e  s  g  s  s  b  y  b  s  t  s  c  s  l  c  '
    +'c  g  b  f  s  d  q  n  z  l  c  c  g  l  l  l  z  gx h  z  c  t  h  c  n  m  j  g  y  z  a  z  c  m  s  k  s  s  t  z  m  m  z  cs k  b  j  y  g  q  l  j  y  j  p  p  l  d  x  r  k  z  y  x  c  c  s  n  h  s  h  h  d  z  n  l  z  h  zc j  j  c  d  ds '
    +'c  b  c  j  x  l  b  f  q  b  cz z  z  t  p  q  d  n  n  x  l  j  c  t  h  q  wz j  g  y  l  k  l  s  z  z  p  c  j  d  s  c  q  j  h  j  q  k  d  x  g  p  b  a  j  y  n  n  s  m  j  t  z  d  x  l  c  j  y  r  y  y  n  h  j  b  n  g  z  j  k  m  j  x  '
    +'l  t  b  s  l  l  r  z  p  y  l  s  s  zc n  x  j  h  l  l  h  y  l  l  q  q  zc q  l  s  y  m  r  c  n  c  x  s  l  j  m  l  zc l  t  z  l  d  w  d  jg j  l  l  n  z  g  g  q  x  p  p  s  k  y  g  g  g  g  b  f  z  b  d  k  m  w  g  g  c  x  m  c  g  '
    +'d  x  j  m  c  j  s  d  y  c  a  b  x  j  d  l  n  b  ct d  d  y  g  s  k  y  d  q  d  x  d  j  j  y  x  hm s  a  q  a  z  d  z  f  s  l  q  x  x  j  n  q  z  yd l  b  l  x  x  w  x  qj qj z  b  j  z  l  f  b  b  l  y  l  w  d  s  l  j  h  x  j  y  z  '
    +'j  w  t  d  j  c  y  f  q  z  qj z  z  d  z  s  x  z  z  q  l  z  c  d  z  f  x  h  wy s  p  y  n  pb q  z  m  l  p  p  l  f  f  x  j  j  n  z  z  y  l  s  j  n  y  q  z  f  p  f  z  g  s  y  wm j  j  j  h  r  d  j  z  z  x  t  x  x  g  l  g  h  t  d  '
    +'x  c  s  k  y  s  w  m  m  t  c  w  y  b  a  z  b  j  k  s  h  f  h  g  c  x  m  h  f  q  h  y  x  x  y  z  f  t  s  j  y  z  b  x  y  x  p  z  l  c  h  m  z  m  b  x  h  z  z  s  s  y  f  d  m  n  c  w  d  a  b  a  z  l  x  k  t  cz s  h  h  x  k  xk '
    +'j  j  z  j  s  t  h  y  g  x  s  x  y  y  h  h  h  j  w  x  k  z  x  sc s  b  z  z  w  h  h  h  c  w  t  z  z  z  p  jy x  s  y  x  q  q  j  g  z  y  z  a  w  l  l  c  w  x  z  n  x  g  y  x  y  h  f  m  k  h  y  d  w  s  q  m  n  j  n  a  y  c  y  s  '
    +'p  m  j  k  g  w  c  q  h  y  l  a  j  g  m  z  x  h  m  m  c  n  z  h  bp h  x  c  l  x  d  j  p  l  t  x  y  j  k  d  y  y  l  t  t  x  f  q  z  h  y  x  x  s  j  b  j  n  a  y  r  s  m  x  y  pb l  c  k  d  n  y  h  l  x  r  l  n  l  l  s  t  y  c  '
    +'y  y  q  yz g  z  h  h  s  c  c  s  m  cz zc t  z  c  x  hk y  q  f  pb y  y  r  p  b  f  l  f  q  n  n  t  s  z  l  l  jz m  h  w  t  cz jq qj y  z  w  t  l  n  m  l  m  d  w  m  b  z  z  sz n  z  r  b  p  d  d  d  l  q  j  j  b  x  t  c  s  n  z  q  '
    +'q  y  g  w  c  s  x  f  w  z  l  x  c  c  r  sb z  d  z  m  c  y  g  g  d  y  q  s  g  t  n  n  n  l  s  m  y  m  m  s  y  h  f  b  j  d  g  ym x  c  c  pb s  h  x  cn z  c  s  b  s  j  y  y  g  j  m  p  b  w  a  f  f  y  f  n  x  h  y  d  x  z  y  l  '
    +'r  e  m  z  g  zq zc y  n  d  s  z  n  l  l  j  c  s  q  f  n  x  x  k  p  t  x  z  g  x  j  j  gj bf m  y  y  s  s  n  b  td y  l  b  n  l  h  b  fb z  d  c  y  f  b  m  g  q  r  r  m  z  s  z  x  y  s  gj t  z  n  n  y  d  z  z  cz d  g  n  j  a  fb '
    +'j  b  d  k  n  z  b  l  c  z  s  z  p  s  g  cz y  c  j  s  z  l  m  n  r  zs n  b  z  z  l  d  l  n  ly l  y  s  x  s  q  z  qx l  cy x  z  l  s  gk k  b  r  xj b  r  b  z  c  y  c  x  z  j  z  e  e  y  f  g  k  l  z  l  n  y  h  g  z  c  g  z  l  f  '
    +'j  h  g  t  g  w  k  r  a  a  j  y  z  k  z  q  t  s  s  h  j  j  x  d  z  y  z  n  y  n  n  z  y  r  z  d  q  q  h  gh j  z  x  s  s  z  b  t  k  j  b  b  f  r  t  j  x  l  l  f  q  w  j  g  c  l  q  t  y  m  b  l  p  z  d  x  t  z  a  g  b  d  h  z  '
    +'z  r  b  g  j  h  w  n  j  t  j  x  l  k  s  c  f  s  m  w  l  l  d  c  y  s  j  t  x  k  z  s  c  f  w  j  l  b  n  n  t  z  l  l  j  z  l  l  q  b  l  c  q  m  q  q  c  g  c  d  f  p  b  pb h  z  c  z  j  l  p  y  y  g  h  d  t  g  w  d  x  f  c  z  '
    +'q  y  y  y  q  y  s  r  c  l  qx z  f  k  l  z  z  z  g  f  f  c  q  n  w  g  l  h  j  y  c  j  j  cs z  l  q  z  zc y  j  bp j  z  z  b  p  d  c  s  n  n  j  g  x  d  q  n  k  n  l  z  n  n  n  n  p  s  n  t  s  d  y  f  w  w  d  j  z  j  y  sz x  y  '
    +'y  c  z  c  yj h  zc w  b  b  y  h  x  r  y  l  y  b  h  k  j  k  s  f  x  t  j  j  m  m  c  hk h  l  l  t  n  y  y  m  s  x  x  y  z  p  d  j  j  y  c  sd y  c  w  m  d  j  j  kj q  y  r  h  l  ly n  g  p  n  g  t  l  y  y  c  l  j  n  n  n  x  j  y  '
    +'z  f  n  m  l  r  g  j  j  t  y  z  b  s  y  z  m  s  j  y  j  h  gh f  z  q  m  s  y  x  r  s  z  c  y  t  l  r  t  q  z  s  s  t  k  x  g  q  k  g  s  p  t  g  x  dz n  j  s  g  c  q  c  q  h  m  x  g  g  z  t  q  y  d  j  jk z  n  l  b  z  n  x  q  '
    +'l  h  y  q  g  g  g  t  h  qc s  c  b  y  h  j  h  h  k  y  y  g  k  g  g  c  m  j  d  z  ly l  c  cz l  x  q  s  f  t  g  j  s  l  l  l  m  l  c  s  k  cs t  b  l  j  s  z  s  z  m  m  n  y  t  p  z  s  x  q  h  j  c  n  n  q  n  y  e  x  z  q  z  c  '
    +'p  s  h  k  z  z  y  z  x  x  d  f  g  m  w  qc rn l  l  q  x  r  f  z  t  l  y  s  t  c  t  m  j  c  s  j  j  t  h  j  n  x  t  n  r  z  tc z  f  q  r  h  c  gk l  l  g  c  n  n  n  n  j  d  n  l  n  n  y  t  s  j  t  l  n  y  x  s  s  z  x  c  g  j  '
    +'z  y  q  p  y  l  f  h  d  j  s  b  b  d  c  z  ghkj  j  j  q  z  j  q  d  y  b  s  s  l  ly c  m  y  t  t  m  q  n  b  h  j  q  mn n  y  g  jk y  e  q  y  q  m  z  gj c  j  k  p  d  c  n  m  y  z  g  q  l  l  s  l  n  c  l  m  h  o  l  zj g  d  y  ly '
    +'f  z  s  l  n  c  n  z  l  y  l  z  c  j  e  s  h  n  y  l  l  n  x  n  j  x  l  y  j  y  y  y  x  n  b  c  l  j  s  s  w  c  q  q  n  n  y  l  l  z  l  d  j  n  l  l  z  l  lj b  n  y  l  n  q  c  h  x  y  y  q  o  x  c  c  q  k  y  j  x  x  x  y  k  '
    +'l  k  s  x  ea yq q  h  c  q  k  k  k  k  c  s  n  y  x  x  y  q  x  y  g  w  t  j  o  h  t  h  x  p  x  x  h  sx n  l  c  y  k  y  c  h  z  z  c  b  w  q  b  b  w  j  q  c  s  c  s  z  s  s  l  c  y  l  g  d  d  s  j  z  m  m  y  m  c  y  t  s  d  s  '
    +'x  x  s  c  j  p  q  q  s  q  y  l  y  f  zs y  c  h  d  j  y  n  y  w  cx b  t  j  s  y  d  c  h  c  y  d  d  j  l  b  d  j  j  s  o  d  z  y  q  y  s  kq kq y  x  d  h  h  g  q  j  y  o  h  d  y  x  w  g  m  m  m  a  z  d  y  b  b  b  p  p  b  c  m  '
    +'n  n  p  n  j  z  s  m  t  x  e  r  x  j  m  h  q  d  n  t  p  j  d  c  b  s  n  m  s  s  y  t  h  j  t  s  l  m  l  t  r  c  p  l  z  s  z  m  l  q  d  sz d  m  j  m  q  p  n  q  d  x  c  f  r  n  n  f  s  d  q  q  y  x  h  y  a  y  k  q  y  d  d  l  '
    +'q  y  y  y  s  s  z  b  y  d  s  l  n  t  f  g  t  z  q  b  z  m  c  h  d  h  c  z  c  w  f  d  x  t  m  q  q  s  p  h  q  w  w  x  s  r  gh j  c  w  n  n  ts z  c  q  m  g  w  q  j  r  j  h  t  q  j  b  b  g  w  z  f  x  j  h  n  q  f  x  x  qs y  w  '
    +'y  y  h  y  sc c  d  y  dz h  h  q  m  r  m  td m  w  c  t  b  s  z  p  p  z  z  g  l  m  z  f  o  l  l  c  f  w  h  m  m  s  j  z  t  t  d  h  l  m  y  f  f  y  t  z  z  g  z  y  s  k  j  j  x  q  y  j  z  q  bp h  mb b  z  c  l  y  g  h  g  f  m  s  '
    +'h  p  c  f  z  s  n  c  l  p  bp qx s  n  j  y  z  s  l  x  x  f  p  m  t  y  j  y  g  b  x  ls l  d  lp x  p  z  j  y  z  j  y  h  h  z  c  y  w  h  j  y  l  s  j  e  x  f  s  z  z  y  w  x  k  z  j  l  n  a  d  y  m  l  y  m  q  j  p  w  x  x  h  x  '
    +'s  k  t  qj j  e  z  r  p  x  x  z  gh h  m  h  w  qx p  w  q  l  y  j  j  qj j  j  z  s  z  c  n  hk j  l  c  h  hx n  x  j  l  q  w  z  jx h  b  m  z  y  x  b  d  hw h  y  p  y  l  h  l  h  l  g  f  w  l  c  f  y  y  t  l  h  j  j  c  w  m  s  c  p  '
    +'x  s  t  k  p  n  h  j  x  s  n  t  yc x  x  t  e  s  t  j  c  t  l  s  s  l  s  t  d  l  l  l  w  wg y  h  d  n  r  j  z  s  f  g  x  s  s  y  c  z  y  k  w  h  t  d  h  w  j  g  l  h  t  z  d  q  d  j  z  x  xy q  g  g  h  l  t  z  p  h  c  s  q  f  '
    +'c  l  n  j  t  c  l  z  p  f  s  t  p  d  y  n  y  l  g  m  j  l  l  y  c  q  h  y  n  s  ph c  h  y  l  h  q  y  qj t  m  z  y  m  b  y  w  r  f  q  y  k  j  s  y  s  l  z  d  n  j  ms p  x  y  y  s  s  r  h  z  j  n  y  q  t  q  d  f  z  b  w  wg d  '
    +'w  w  r  x  c  w  g  g  y  h  x  m  k  m  y  y  y  h  m  x  m  z  h  n  k  s  e  p  m  l  q  q  m  t  c  w  c  t  m  x  m  x  j  p  j  j  h  f  x  y  y  z  s  j  z  h  t  y  b  m  s  t  s  yc j  z  n  qj j  n  y  ts l  h  y  n  b  y  q  c  l  c  y  c  '
    +'n  z  w  s  m  y  l  kh n  j  x  l  g  g  n  n  p  j  g  t  y  s  y  l  y  m  z  s  k  t  t  w  l  g  s  m  z  s  y  l  m  p  w  l  c  w  xq w  q  c  s  s  y  zq s  y  x  y  r  h  s  s  n  t  s  r  w  p  c  c  p  w  c  m  h  d  h  h  x  z  d  z  yx f  '
    +'j  hg g  z  t  t  s  b  j  h  g  y  g  ls z  y  s  m  y  c  l  l  l  x  b  t  y  x  h  b  b  z  j  k  s  s  d  m  a  l  hx h  y  c  f  y  g  m  q  y  p  j  y  j  qx x  j  l  l  l  j  g  c  l  z  g  q  l  y  c  j  c  c  t  o  t  y  x  m  t  m  s  h  l  '
    +'l  w  l  q  f  x  y  m  z  m  k  lt p  s  z  z  c  x  h  kh j  y  cs l  c  t  y  j  c  y  h  x  s  g  y  x  n  n  x  l  z  w  p  y  j  p  x  h  j  w  p  j  p  w  xs q  q  x  l  x  s  d  h  m  r  s  l  z  z  y  d  w  d  t  c  x  kh n  s  t  z  s  h  b  '
    +'s  c  c  s  t  p  l  w  s  s  cd j  cd h  j  l  c  g  c  h  s  s  p  h  y  l  h  f  h  h  x  j  s  x  a  l  l  n  y  l  ms z  d  h  z  x  y  l  s  x  l  m  z  y  kh c  l  d  y  a  h  l  c  m  d  dt y  s  p  j  t  q  j  z  l  n  g  j  f  s  j  s  h  c  '
    +'t  s  d  s  z  l  b  l  m  s  s  m  n  y  yx m  j  q  b  j  h  r  zc w  t  y  y  d  c  h  j  l  j  a  p  z  w  b  g  q  x  b  k  f  n  b  j  dw l  l  l  l  y  y  l  s  j  y  d  w  h  x  p  s  b  c  m  l  j  pb s  c  g  bp h  x  l  q  h  y  r  l  j  x  '
    +'y  s  w  x  h  h  z  ls l  d  f  h  l  n  n  y  m  j  l  j  y  f  l  y  j  y  c  d  r  j  l  f  s  y  z  f  s  l  l  c  q  y  q  f  g  q  y  h  n  s  z  l  y  l  m  d  t  d  j  c  n  h  b  z  l  l  n  w  l  q  x  y  g  y  y  h  b  m  g  d  h  x  x  n  '
    +'h  l  z  z  j  z  x  c  z  z  z  c  y  q  z  f  n  gj w  p  y  l  c  p  k  p  y  k  p  m  c  l  qg k  d  g  x  z  g  x  w  q  b  d  x  z  z  k  z  f  b  d  d  l  z  x  j  t  pb j  p  t  t  b  y  td h  z  z  d  w  s  l  c  p  n  h  s  l  t  j  x  x  q  '
    +'l  h  y  x  x  x  y  w  z  y  s  w  t  t  z  k  h  l  x  z  x  z  p  y  h  g  z  h  k  n  f  s  y  h  n  t  j  r  n  x  f  j  c  p  tj z  t  w  h  p  l  s  h  f  c  r  h  n  s  l  x  x  j  x  x  y  h  z  qj dt x  q  w  n  n  h  y  h  m  j  dt b  f  l  '
    +'k  h  c  x  c  wm h  j  f  y  j  c  f  p  q  c  x  q  x  zc y  y  y  j  y  g  r  p  y  n  s  c  s  n  n  n  n  c  h  k  z  d  y  h  f  l  x  x  h  j  j  b  y  z  w  t  t  x  nx n  c  y  j  j  y  m  s  w  y  x  q  r  m  h  x  z  w  f  q  s  y  l  z  n  '
    +'g  g  b  h  y  x  n  n  b  w  t  t  c  s  y  b  h  x  x  w  x  y  h  h  x  y  x  n  k  n  y  x  m  l  y  w  r  n  n  q  l  x  b  b  c  l  j  s  y  l  f  s  y  t  j  z  y  h  y  z  a  w  l  h  o  r  j  m  n  s  c  z  j  x  x  x  y  x  c  h  c  y  q  r  '
    +'y  x  q  z  d  d  s  j  f  s  l  y  l  t  s  f  f  y  x  l  m  t  y  j  m  n  n  y  y  y  x  l  t  z  c  s  x  q  c  l  h  z  x  l  w  y  x  z  h  n  n  l  r  x  k  x  j  c  dt y  h  l  b  r  l  m  b  r  d  l  a  x  k  s  n  l  l  j  l  y  x  x  l  y  '
    +'n  rs y  l  c  j  t  g  n  c  m  t  l  z  l  l  c  y  z  l  p  z  p  z  y  a  w  n  j  j  f  y  b  d  y  y  z  s  e  pqbc  k  z  z  q  d  q  p  b  p  s  j  p  d  y  t  t  b  d  b  b  b  y  n  d  y  c  n  c  p  j  m  t  m  l  r  m  f  m  m  r  w  y  f  '
    +'b  s  j  g  y  g  s  m  d  q  q  q  z  t  x  m  k  q  w  g  x  l  l  p  j  g  z  b  q  r  d  j  j  jq f  p  k  j  kd c  x  b  l  j  m  s  w  l  d  t  s  j  x  l  d  l  p  p  b  x  c  w  k  c  q  q  b  f  q  b  c  c  ah j  z  g  m  y  k  b  h  y  h  h  '
    +'z  y  k  n  d  q  z  y  b  p  j  n  s  p  x  t  h  l  f  p  n  s  y  g  y  j  d  b  g  x  n  h  h  j  hm z  j  h  s  t  r  s  t  l  d  x  s  k  z  y  s  y  b  m  x  j  l  x  y  s  l  b  z  y  s  l  z  x  j  h  f  y  bp q  n  b  y  l  l  j  q  k  y  g  '
    +'z  m  c  y  z  z  y  m  c  c  s  l  n  l  h  z  h  w  f  w  y  x  z  m  w  y  x  t  y  n  xh j  h  b  y  y  m  c  y  s  b  m  h  y  s  m  y  d  y  s  h  n  y  z  c  h  m  j  j  m  z  c  a  a  h  c  b  j  b  b  h  b  l  y  t  y  l  s  x  s  n  x  g  j  '
    +'d  h  k  x  x  t  x  x  n  b  h  n  m  l  n  g  s  l  t  x  m  r  h  n  l  x  q  q  x  m  z  ls l  y  s  w  q  g  d  l  b  j  h  d  c  g  j  y  q  y  yc m  h  w  f  mw j  y  bf b  b  y  j  y  j  w  j  m  d  p  w  h  x  q  l  d  y  a  p  d  f  x  x  b  '
    +'c  g  j  s  p  c  k  r  s  s  y  z  j  m  s  l  b  z  z  j  f  l  j  j  j  l  g  x  z  g  y  x  y  x  l  s  z  q  k  x  b  e  x  y  x  h  g  c  x  b  p  n  d  y  h  w  e  c  td w  w  c  j  m  b  t  x  c  h  x  y  q  q  l  l  x  f  l  y  x  l  l  j  l  '
    +'s  s  n  w  d  b  z  c  m  y  j  c  l  w  s  w  d  c  z  p  c  h  q  e  k  c  q  b  w  l  c  g  y  d  b  l  q  p  pb q  z  q  f  n  q  d  j  h  y  m  m  c  x  t  x  d  r  m  z  w  r  h  x  c  j  z  y  l  q  x  d  y  y  n  h  y  y  h  r  s  l  n  r  s  '
    +'y  w  w  j  j  y  m  t  l  t  l  l  g  t  q  c  j  z  y  a  b  t  c  k  z  c  j  y  c  c  q  l  y  s  q  x  a  l  m  z  y  n  y  w  l  w  d  n  z  x  q  d  l  l  q  s  h  g  p  j  f  j  l  j  n  j  a  b  c  q  z  d  j  g  t  h  h  s  s  t  n  y  j  f  '
    +'b  s  w  z  l  x  j  x  r  h  g  l  d  l  z  r  l  z  q  z  g  s  l  l  l  l  z  l  y  m  x  x  g  d  z  h  g  b  d  p  h  z  p  b  r  l  w  n  jh qw b  p  f  d  w  o  n  n  n  h  l  y  p  c  n  j  c  c  n  d  m  b  c  p  b  z  z  n  c  y  q  x  l  d  '
    +'o  m  z  b  l  z  w  p  d  w  y  y  g  d  s  t  t  h  c  s  q  s  c  c  r  s  s  s  y  s  l  f  y  b  n  n  t  y  j  s  z  d  f  n  d  p  dt h  td q  z  m  b  q  l  x  l  c  m  y  f  f  g  t  j  j  q  w  f  t  m  n  p  j  w  d  n  l  b  z  xc m  m  c  '
    +'n  g  b  d  z  l  qx l  p  n  fp h  y  y  m  j  y  l  s  d  c  h  d  c  j  w  j  c  c  t  l  j  c  l  d  t  l  j  j  c  b  d  d  psyn  d  s  s  z  y  c  n  d  b  j  l  g  g  j  z  x  s  x  ny l  y  c  y  b  j  j  x  x  c  b  y  l  z  c  f  z  p  p  g  '
    +'k  c  x  q  d  z  f  z  t  j  j  f  j  d  j  x  z  b  n  z  y  j  q  c  t  y  j  w  h  d  y  c  z  h  y  m  d  j  x  t  t  m  p  x  s  p  l  z  c  d  w  s  l  s  h  x  y  p  z  g  t  f  m  l  c  j  t  a  c  b  b  m  g  d  e  w  y  c  y  z  x  d  s  z  '
    +'jz y  h  f  l  y  s  t  y  g  w  h  k  jx y  y  l  s  j  c  x  g  y  w  j  c  b  l  l  c  s  ny d  d  b  t  z  b  s  c  l  y  z  c  z  z  s  s  q  d  l  l  m  j  y  y  h  f  l  l  q  l  l  x  f  d  y  h  ay b  x  g  g  n  y  w  y  y  p  l  lj s  d  l  '
    +'d  l  l  b  j  c  y  x  j  z  n  l  h  l  j  d  x  y  y  q  y  t  d  l  l  l  b  n  g  p  f  d  f  b  b  q  b  z  z  m  d  p  j  h  g  c  l  g  m  j  j  p  g  a  e  h  h  b  wh c  q  x  a  j  h  h  h  z  c  h  x  y  p  h  j  a  x  h  l  p  h  j  p  g  '
    +'p  z  j  q  c  q  z  g  j  j  z  z  g  z  d  m  q  y  y  b  z  z  p  h  y  h  y  b  w  h  a  z  y  j  h  y  k  f  g  d  p  f  q  sc d  l  z  m  l  j  x  j  p  g  a  l  x  z  d  a  g  l  m  d  g  x  m  m  z  q  w  t  x  d  x  x  p  f  ds m  m  sx s  y  '
    +'m  p  f  m  d  m  m  k  x  k  s  y  z  y  s  h  d  z  k  j  s  y  s  m  m  z  z  z  m  d  y  d  y  z  z  c  z  x  b  m  l  s  t  m  d  y  e  m  x  c  k  j  m  z  t  y  y  m  z  m  z  z  m  s  s  h  h  d  c  c  j  e  w  x  x  k  l  j  s  t  h  w  l  s  '
    +'q  l  ya j  z  l  l  s  j  s  s  d  p  pb m  h  n  l  g  j  c  z  y  h  m  x  x  h  g  n  c  j  m  d  h  x  t  k  g  r  m  x  f  w  m  c  k  m  w  k  d  c  k  s  x  q  m  m  m  s  z  z  y  d  k  m  s  c  l  c  m  p  c  j  m  h  r  p  x  qy p  z  d  s  '
    +'s  l  c  x  k  y  x  t  w  l  k  j  y  a  h  z  j  g  z  jq w  c  j  n  x  y  h  m  m  bp m  l  g  j  x  m  h  l  m  l  g  m  x  c  t  k  z  m  jqgl  y  s  c  j  s  y  s  z  h  s  y  j  z  j  c  d  a  j  z  h  bp sd d  q  j  z  g  w  t  k  q  x  f  k  '
    +'d  m  s  d  j  l  f  m  n  hx k  z  q  k  j  f  e  y  p  z  y  s  z  c  d  p  y  n  f  f  m  z  q  y  k  t  t  d  z  z  e  f  m  z  l  b  n  p  p  l  p  l  p  b  p  s  z  a  l  l  t  n  l  k  c  k  q  z  k  gh e  n  j  l  wk a  lg k  x  y  d  p  x  n  '
    +'h  s  x  q  n  w  q  n  k  x  q  c  l  h  y  x  x  m  l  n  c  c  w  l  y  m  q  y  c  k  y  n  n  l  c  j  n  s  z  k  p  y  z  k  c  q  z  q  l  j  b  d  m  d  j  h  l  a  s  q  l  b  y  d  w  q  l  w  d  g  b  q  c  r  y  d  d  z  t  j  y  b  k  b  '
    +'w  s  z  d  x  d  t  n  p  j  d  t  c  n  q  n  f  x  q  q  m  g  n  s  e  c  l  s  t  bp h  p  w  s  l  c  t  x  x  l  p  w  y  d  z  k  l  n  q  g  z  c  q  a  p  l  l  k  q  c  y  l  b  q  m  q  c  z  q  c  n  j  s  l  q  z  d  j  x  l  d  d  h  p  '
    +'z  q  d  l  j  j  x  z  q  d  j  y  z  h  h  z  l  k  c  j  q  d  wy j  p  p  y  p  q  a  k  j  y  r  m  p  z  b  n  m  c  x  kg l  l  z  l  l  f  q  p  y  l  l  l  m  b  s  g  l  z  y  s  s  l  r  s  y  s  q  t  m  x  y  x  zq q  z  b  s  c  n  y  s  '
    +'y  z  t  f  f  m  z  z  s  m  z  q  h  z  s  s  c  c  mn l  y  x  w  t  p  z  g  x  z  jz g  z  g  s  j  z  gj k  d  d  h  t  q  g  g  z  l  l  b  j  d  z  l  sc b  z  h  y  x  y  z  h  z  f  y  w  x  y  t  y  m  s  d  n  z  z  y  j  g  t  c  m  t  n  '
    +'x  q  y  x  j  s  cs x  h  s  l  n  nm d  l  r  y  t  z  l  r  y  y  l  x  q  h  t  x  s  r  t  z  c  g  y  x  b  n  q  q  zc f  h  y  k  m  z  j  b  z  y  mblk  b  p  n  l  y  z  p  b  l  m  c  n  q  y  z  z  z  s  j  z  t  j  c  t  z  h  h  y  z  z  '
    +'j  r  d  y  z  h  n  f  x  k  l  f  z  s  l  k  g  j  t  c  t  s  s  y  l  l  g  z  r  z  b  b  jq z  z  k  l  p  k  b  c  z  y  s  n  n  y  x  b  j  f  b  n  j  zc z  x  c  d  w  l  z  y  j  x  z  z  d  j  jq g  g  g  r  s  n  j  k  m  s  m  z  j  l  '
    +'sw j  y  w  q  s  n  y  h  q  j  s  x  p  j  z  t  n  l  s  n  s  h  r  n  y  n  j  t  w  c  h  g  l  b  n  r  j  l  z  x  w  j  q  x  q  k  y  s  j  y  c  z  t  l  q  z  y  b  b  y  b  y  z  j  q  d  w  g  y  z  c  y  t  j  c  j  x  c  k  c  w  d  k  '
    +'k  z  x  s  n  k  d  n  yx w  w  y  y  j  q  y  y  t  l  y  t  d  j  l  x  w  k  c  j  n  k  l  c  c  p  z  c  q  q  d  z  z  q  l  c  sg f  qgsc  h  qj qfzgs sfzs  m  j  z  z  l  l  b  j  j  z  ygssq j  h  t  s  j  d  y  s  j  q  j  p  d  s  z  c  d  '
    +'c  h  j  k  j  z  z  l  p  y  c  g  m  z  n  d  j  x  b  s  j  z  z  s  y  z  y  h  g  x  c  p  b  j  y  d  s  s  x  d  z  n  c  g  l  q  m  b  t  s  f  c  b  f  d  z  d  l  z  n  f  g  f  j  g  f  s  m  p  n  j  q  l  n  b  l  g  q  c  y  y  x  b  q  '
    +'g  d  j  j  q  s  r  f  k  z  t  j  d  h  c  z  k  l  b  s  d  z  c  f  yj t  p  l  l  j  g  j  h  t  x  z  c  s  s  z  z  x  s  t  jc y  g  k  g  c  k  g  y  n  q  x  j  p  l  z  b  b  b  g  c  g  y  j  z  g  c  z  q  s  z  l  b  j  l  s  j  f  z  g  '
    +'k  q  q  j  c  g  y  c  j  b  z  q  t  l  d  x  r  j  n  b  s  x  x  p  z  s  h  sy z  y  c  f  w  d  s  j  j  h  x  m  f  c  z  p  f  z  h  q  h  q  m  q  n  k  n  l  y  h  t  y  c  g  f  r  z  g  n  q  x  c  g  p  d  l  b  z  c  s  c  z  q  l  l  j  '
    +'b  l  h  b  d  c  y  p  s  c  p  p  d  y  m  z  zk x  g  y  h  c  k  c  p  z  j  g  s  l  z  l  n  s  c  n  s  l  d  l  x  b  m  s  d  l  d  d  f  j  m  k  d  q  d  h  s  l  z  x  l  s  z  n  p  q  p  g  j  d  l  y  b  d  s  k  g  q  l  b  z  l  n  l  '
    +'k  y  y  h  z  t  t  m  c  j  n  q  t  z  z  f  s  z  q  k  t  l  l  j  t  y  y  l  l  n  l  l  q  y  z  q  l  b  d  z  l  s  l  y  y  z  yx m  d  f  s  z  s  n  x  l  x  z  n  c  z  q  n  b  b  w  s  k  r  f  b  cs y  l  c  t  n  b  l  g  j  p  m  c  '
    +'zn z  l  s  t  l  x  s  h  t  zc c  y  zy l  z  b  n  f  m  q  n  l  x  f  l  c  jg l  y  l  j  q  c  b  c  l  z  j  g  n  s  s  t  b  r  m  h  x  z  h  j  z  c  l  x  f  n  b  gy x  g  t  q  n  c  z  t  m  s  f  z  k  j  m  s  s  n  c  l  j  k  b  h  '
    +'s  z  j  n  t  n  l  z  d  n  t  l  m  m  j  xj g  z  j  y  j  c  z  x  y  hg y  hg w  r  w  w  q  n  z  t  n  f  j  s  c  p  y  s  h  z  j  f  y  r  d  j  s  f  s  c  j  z  b  j  f  z  q  z  c  h  z  l  x  f  x  s  b  z  q  l  z  s  g  y  f  t  z  d  '
    +'c  s  z  x  z  j  b  j  p  s  z  k  j  r  h  x  j  z  c  g  b  j  k  h  c  g  g  t  xj k  j  q  g  l  x  bp x  f  gj t  r  t  y  l  x  q  x  h  d  t  s  j  x  h  j  z  j  j  c  m  z  l  c  q  s  b  t  x  w  q  g  x  td x  x  h  x  f  t  s  d  k  f  j  '
    +'h  z  y  j  f  j  x  n  z  l  d  l  l  l  c  q  s  q  q  z  q  w  q  x  s  w  q  t  w  g  w  b  z  c  g  c  lg l  q  z  b  c  l  m  q  j  t  z  g  z  y  z  x  l  j  f  r  m  y  z  f  l  x  n  s  n  x  x  j  k  x  r  m  jq d  z  d  m  m  y  x  b  s  q  '
    +'b  h  g  z  m  w  f  w  y  g  m  j  l  z  bp y  y  t  g  z  y  c  c  d  j  y  z  x  s  n  g  n  y  j  y  z  n  b  g  p  z  j  c  q  s  y  x  s  x  r  t  f  y  z  g  r  h  z  t  x  s  z  z  t  h  c  b  f  c  l  s  y  x  z  l  z  q  m  z  l  m  p  l  m  '
    +'x  z  j  s  s  fp s  b  y  s  m  zy q  h  x  xj n  x  r  x  h  q  z  z  z  s  s  l  y  f  l  c  z  j  r  c  r  x  h  h  z  x  q  n  d  s  h  xj s  j  j  h  qz c  j  jz b  c  y  n  s  y  s  x  j  b  q  jk p  x  z  q  p  l  m  l  x  z  k  y  x  l  x  c  '
    +'n  l  c  y  c  x  x  z  z  l  x  d  l  l  l  m  j  y  hg z  xq hg y  j  w  k  j  r  w  y  h  c  p  s  g  n  r  z  lg f  z  w  f  z  z  n  s  x  g  x  f  l  z  s  x  z  z  z  b  f  c  s  y  j  d  b  r  j  k  r  d  h  h  gj x  j  l  j  j  t  g  x  j  x  '
    +'x  s  t  j  td j  x  l  y  x  q  f  c  s  g  s  w  m  s  b  c  t  l  q  z  z  w  l  z  z  k  x  j  m  l  t  m  jq y  h  s  d  d  bp x  g  z  h  d  l  b  m  y  j  f  r  z  f  c  g  c  l  y  j  b  p  m  l  y  s  m  s  x  l  s  z  j  q  qszh  jz z  f  x  '
    +'g  f  q  f  q  b  p  h  n  g  y  y  q  x  g  z  t  n  q  w  y  l  t  l  g  w  g  w  w  h  n  l  f  m  f  g  z  j  m  g  m  g  b  g  t  j  f  ls y  z  z  g  z  y  z  a  f  l  s  s  p  m  l  bp f  l  c  w  b  j  z  t  l  j  j  m  z  l  pb j  j  l  y  m  '
    +'q  t  m  y  y  y  f  b  g  y  g  q  z  g  l  y  z  d  x  q  y  x  r  q  q  q  h  s  x  y  y  q  x  y  g  j  t  y  x  f  s  f  s  l  l  g  n  q  c  y  g  y  c  w  f  h  c  c  c  fp x  p  y  l  y  p  l  l  z  q  x  x  x  x  x  q  q  h  h  s  s  h  j  dz '
    +'c  f  t  s  c  z  j  x  s  p  z  w  h  h  h  h  h  a  p  y  l  q  n  l  p  q  a  f  y  h  x  d  y  l  n  k  m  z  q  g  g  g  d  d  e  s  r  e  n  dz l  t  z  g  c  h  y  bp p  c  s  q  j  j  h  c  l  l  j  t  o  l  n  j  p  z  l  j  l  h  y  m  h  e  '
    +'z  d  y  d  s  q  y  c  d  d  h  g  z  n  n  d  z  c  l  z  y  w  l  l  z  n  t  e  y  d  g  n  l  h  s  l  p  j  j  b  d  g  w  x  p  c  n  n  t  y  c  k  l  k  c  l  w  k  l  l  c  a  s  s  t  k  n  z  d  n  n  j  t  t  l  y  y  z  s  s  y  s  s  z  '
    +'z  r  y  l  j  q  k  c  g  d  h  h  y  r  x  r  z  y  d  g  r  g  c  w  c  g  z  q  f  f  b  p  p  j  f  z  y  n  a  k  r  g  y  w  y  j  p  q  x  x  f  k  j  t  s  z  z  x  s  w  z  d  d  f  b  b  q  t  b  g  t  z  k  z  n  p  z  f  p  z  x  z  p  j  '
    +'s  z  b  m  q  h  k  y  y  x  y  l  d  k  l  j  n  y  p  k  y  g  hg g  d  z  j  x  x  e  a  x  p  n  z  n  c  t  z  c  m  x  c  x  m  m  j  x  n  k  s  z  q  n  m  n  l  w  b  w  w  q  j  j  y  h  c  l  s  t  m  c  s  x  n  j  c  x  x  t  p  c  n  fp '
    +'d  t  n  n  p  g  l  l  l  z  c  j  l  s  p  b  l  p  g  j  cz d  t  n  j  j  lx y  ay r  s  cd f  f  j  f  q  w  d  p  g  z  d  w  m  r  zd z  c  gl o  d  a  x  n  s  s  n  y  z  r  e  s  t  y  j  w  j  y  j  d  b  c  f  x  n  m  w  t  t  bp q  l  w  '
    +'s  t  s  z  g  y  b  l  j  p  x  g  l  b  n  c  l  g  p  c  b  j  f  t  m  x  z  l  j  y  l  z  x  c  l  t  p  n  c  l  c  g  x  t  f  z  j  s  h  c  r  x  s  f  y  s  g  d  k  n  t  l  b  y  j  c  y  j  lx l  s  t  g  q  c  b  x  nr h  z  x  b  x  k  '
    +'lx y  l  h  z  l  q  z  l  n  z  c  q  w  gw z  l  g  z  j  n  cx j  g  cx m  n  z  z  g  j  d  z  x  t  z  j  x  y  c  y  y  c  x  x  j  y  y  x  j  j  x  s  s  s  j  s  t  s  s  t  d  p  p  g  h  t  c  s  x  w  z  d  c  s  y  n  p  t  f  b  c  h  f  '
    +'bp b  l  z  j  c  l  z  z  dt b  x  g  c  jx l  h  p  x  n  f  z  f  l  s  y  l  t  n  w  b  m  n  j  h  s  z  b  m  d  n  b  c  y  s  c  c  l  d  n  y  ct n  d  q  l  y  j  j  h  m  q  l  l  c  s  g  l  j  j  s  y  f  p  y  y  c  c  y  l  t  j  ay n  '
    +'t  j  j  p  w  y  c  m  m  g  q  y  y  s  x  d  x  q  mw z  h  s  z  x  bp f  t  w  w  z  q  s  w  q  r  f  k  jg l  z  j  q  q  y  f  b  r  x  j  h  h  f  w  j  g  z  y  q  a  c  m  yd f  r  h  c  y  y  b  y  n  w  l  p  e  x  c  c  zn s  yt y  r  l  '
    +'t  ts d  m  q  l  r  k  m  p  b  g  m  y  y  jq p  r  k  zn n  b  b  s  q  y  x  b  h  y  z  d  j  d  n  g  h  p  m  f  s  g  b  w  f  z  m  f  qj m  m  b  zc m  z  d  c  g  j  l  n  n  n  x  y  q  g  m  l  r  y  g  qx c  c  y  x  z  l  w  d  k  c  j  '
    +'c  g  g  m  cz j  j  f  y  z  z  j  h  y  c  f  r  r  c  m  t  z  n  z  x  h  k  q  g  d  j  x  c  c  j  e  a  s  c  r  j  ty h  p  l  j  lp r  z  d  j  r  b  c  q  h  jq d  n  r  hx yx l  y  qx j  s  y  m  h  z  y  d  w  c  d  f  r  y  h  b  b  y  d  '
    +'t  s  s  c  c  w  b  x  g  l  p  z  m  l  z  j  d  q  s  s  c  f  j  m  m  x  j  c  xs j  y  t  y  c  gw h  y  c  j  w  y  nc s  x  l  fp e  m  w  j  n  m  k  l  l  s  wz t  x  h  y  y  y  n  c  m  m  c  yw j  d  q  d  j  z  g  l  l  j  w  j  n  k  h  '
    +'p  z  g  g  f  l  c  c  s  c  z  m  c  b  l  t  b  h  b  q  j  x  q  d  j  p  d  j  z  t  g  h  g  l  f  j  a  w  b  z  y  z  j  l  t  s  t  d  h  jq h  c  t  c  b  c  h  f  l  q  m  p  w  d  s  h  y  y  t  qj w  c  n  n  t  j  t  l  n  n  m  n  n  d  '
    +'y  y  y  x  s  q  k  x  wm y  y  f  l  x  x  n  z  w  c  x  y  pb m  a  e  l  y  h  g  j  w  z  z  j  b  r  x  yxsa  q  j  f  l  l  p  f  h  h  h  y  t  z  z  x  sr g  q  j  m  h  s  p  g  d  z  q  w  b  w  p  j  h  z  j  d  y  j  c  q  w  x  k  t  hx '
    +'x  s  q  l  z  y  y  m  y  s  d  z  g  n  q  c  k  k  n  j  l  w  p  n  s  y  s  c  s  y  z  l  n  m  h  q  s  y  l  j  x  bp c  x  t  l  h  z  q  z  p  c  y  c  y  k  p  p  p  n  s  x  f  y  z  j  j  r  c  e  m  h  s  z  m  n  x  l  x  g  l  r  w  gh '
    +'c  s  t  l  r  s  x  b  y  g  b  z  g  n  x  c  n  l  n  j  l  c  l  y  n  y  m  d  x  w  t  z  p  a  l  c  x  p  qo j  c  j  w  t  c  y  y  j  l  b  l  xs b  z  l  q  m  y  l  j  bp g  h  d  s  l  s  s  d  m  x  mw b  d  c  z  s  x  wy h  a  m  l  c  '
    +'z  c  p  j  m  c  n  h  j  y  j  n  s  y  k  c  h  s  k  q  m  cz z  q  d  l  l  k  a  b  l  w  j  qx s  fb m  o  c  d  xj jq r  r  l  y  q  c  h  j  m  y  b  y  q  l  r  h  e  t  f  j  z  f  r  f  k  s  r  y  x  f  j  d  w  t  s  x  x  y  w  s  q  j  '
    +'y  s  l  y  x  w  j  h  s  d  l  x  y  y  x  h  b  h  a  w  h  w  j  c  x  l  m  y  l  j  c  sc q  l  k  y  d  t  t  x  b  z  s  lx f  d  x  g  x  s  j  k  h  s  x  x  y  b  s  s  x  d  p  w  n  c  m  r  p  t  qj z  c  z  e  n  y  g  c  x  q  f  j  x  '
    +'k  j  b  d  m  l  j  cz m  q  q  xl n  o  x  s  l  y  x  x  l  y  l  l  j  d  z  p  t  y  m  h  b  f  s  t  t  q  q  w  l  h  s  g  y  n  l  sz z  a  l  z  x  c  l  h  t  we r  r  q  h  l  s  t  m  y  p  y  x  j  j  x  m  n  s  j  n  n  bn r  y  xl y  '
    +'j  l  l  y  q  y  l  t  w  y  l  q  y  f  m  l  k  l  j  d  n  l  l  t  f  z  w  k  z  h  l  j  m  l  h  l  j  n  l  j  n  n  l  q  x  y  l  m  b  h  h  l  n  l  z  x  q  c  h  x  c  f  x  x  l  h  y  h  j  j  g  b  y  z  z  k  b  x  s  ch q  d  j  q  '
    +'d  s  n  d  z  sy y  g  z  hj h  m  g  s  xh c  s  y  m  x  f  e  p  c  q  w  w  r  b  p  y  y  j  q  r  y  qc ct y  j  h  q  q  z  y  h  m  w  f  f  h  g  z  f  r  j  f  c  d  b  x  n  t  q  y  z  p  c  y  hk h  j  l  f  r  z  g  p  b  x  z  d  b  b  '
    +'g  r  qj sy t  l  g  d  g  y  l  c  q  m  g  c  h  h  m  f  y  w  l  z  y  x  k  j  l  y  p  j  gh s  y  w  m  q  q  g  g  z  m  n  z  j  n  s  q  x  l  q  sx y  j  ts c  b  ey h  s  x  f  s  z  f  x  z  w  f  l  l  b  c  y  y  j  d  y  t  d  t  h  w  '
    +'z  s  f  j  m  q  q  y  j  lz m  q  s  x  l  l  d  t  t  k  g  h  y  b  f  p  w  d  y  y  s  q  q  r  n  q  w  l  g  w  d  e  b  z  w  c  y  y  g  c  n  lz k  j  xh t  m  x  m  y  j  s  x  h  y  b  r  w  f  y  m  w  f  r  x  y  yq m  xh y  s  c  t  z  '
    +'z  t  f  y  k  m  l  d  h  q  d  lx wg y  j  n  l  c  r  y  j  t  l  p  s  x  x  x  y  w  l  s  b  r  r  j  w  x  h  q  y  b  h  t  y  d  n  h  h  x  m  m  y  w  y  t  y  c  n  n  m  n  s  sz c  c  d  a  l  w  z  t  c  p  q  p  y  j  l  l  q  z  y  j  '
    +'s  w  j  w  z  z  m  m  g  l  m  x  c  l  m  x  n  z  m  x  m  z  s  q  t  z  p  p  j  q  b  l  p  g  x  j  z  h  f  l  j  j  h  y  c  j  s  rn x  w  c  x  s  n  c  d  l  x  s  y  j  d  c  q  z  x  s  l  q  y  c  l  z  x  l  z  z  x  m  x  q  r  j  m  '
    +'h  r  h  z  j  b  h  m  f  l  j  lz m  l  c  l  q  n  l  d  x  z  l  l  l  f  y  p  r  g  j  y  n  x  c  q  q  d  c  m  q  j  z  z  x  h  n  p  n  x  z  m  e  m  m  s  xh y  k  y  n  l  x  sz x  t  l  j  x  y  h  w  d  c  w  d  z  h  q  y  y  b  g  y  '
    +'b  c  y  s  c  f  g  f  sc j  n  z  d  r  z  z  x  zq x  r  z  r  q  j  j  y  m  c  a  n  h  r  j  t  l  d  b  p  y  z  b  s  t  j  h  x  x  z  y  p  bf d  w  f  g  z  z  r  p  y  m  n  n  k  x  c  q  bp y  x  n  b  n  f  y  c  k  r  j  jq c  m  j  e  '
    +'g  r  z  g  y  c  l  n  n  z  d  n  k  k  n  s  j  k  c  l  j  s  p  g  y  y  c  l  q  q  j  y  b  z  s  s  q  l  l  l  k  j  f  t  bp g  t  y  l  c  c  c  d  b  l  sc p  p  f  y  l  g  y  d  tx z  j  q  j  z  g  k  n  t  s  f  c  x  b  d  kh d  x  x  '
    +'h  y  b  b  f  y  t  y  h  b  cz l  n  n  y  t  g  d  h  r  y  r  n  j  s  b  t  c  s  n  y  j  q  h  k  l  l  l  zx s  l  y  d  x  x  w  b  c  j  q  s  b  x  n  p  j  z  j  z  j  d  z  f  b  x  x  b  r  m  l  a  d  h  c  s  n  c  l  b  j  d  s  t  b  '
    +'l  p  r  z  n  s  w  s  b  x  b  c  l  l  x  x  l  z  d  n  z  s  j  p  y  n  y  x  x  y  f  tq n  n  f  b  h  j  j  jx g  b  y  g  j  p  m  m  m  m  s  s  zc l  j  m  t  l  y  z  j  x  s  w  x  t  y  l  e  d  q  p  j  m  p  g  q  z  j  g  d  j  l  q  '
    +'j  w  j  q  l  l  s  d  g  jx g  y  g  m  s  c  l  j  j  x  d  t  y  g  j  q  j  j  j  c  j  z  c  j  g  d  z  d  s  h  q  g  z  jx g  g  c  jx h  q  x  s  n  j  l  z  z  b  x  h  s  g  z  x  c  x  y  l  j  x  y  x  y  y  d  f  q  q  j  h  j  f  x  d  '
    +'h  c  t  x  j  y  r  x  y  s  q  t  j  x  y  e  f  y  y  s  s  y  x  j  x  n  c  y  z  x  f  x  c  s  x  s  z  xh y  y  s  c  h  s  h  x  z  z  z  g  z  z  z  g  f  j  d  l  d  y  l  n  p  z  g  s  j  a  z  t  y  q  z  p  b  x  cq b  d  z  t  z  c  z  '
    +'y  x  x  y  h  h  s  c  j  s  h  c  g  g  q  h  j  h  g  x  h  s  c  t  m  z  m  e  h  y  x  g  e  b  td c  l  z  k  k  w  y  t  j  z  r  s  l  e  k  e  s  t  d  b  c  y  h  q  q  s  a  y  x  c  j  x  w  w  g  s  p  h  j  sy z  sy d  n  c  s  j  k  q  '
    +'c  x  s  w  x  f  c  t  y  n  y  dt p  c  c  c  z  j  q  t  c  w  j  q  j  z  z  z  q  z  l  j  z  h  l  s  b  h  p  y  d  x  p  s  x  s  h  h  e  z  d  x  f  p  t  j  q  y  z  c  x  h  y  a  x  n  c  f  z  y  y  h  x  g  n  q  m  y  w  n  t  z  s  j  '
    +'b  n  h  h  g  y  m  xs m  x  q  c  nx s  s  b  c  q  sy j  y  x  x  t  y  y  h  y  b  c  q  l  m  m  s  z  m  j  z  z  l  l  c  o  g  x  z  a  a  j  z  y  h  j  m  c  h  h  c  x  z  s  x  s  d  z  n  l  e  y  j  j  z  j  b  h  z  w  jz z  sz q  t  z  '
    +'p  s  x  z  z  d  s  q  j  j  j  l  n  y  a  z  p  h  h  y  y  s  r  n  q  z  t  h  z  h  n  y  j  y  j  h  d  z  x  z  l  s  w  c  l  y  b  z  y  e  c  w  c  y  c  r  y  l  c  h  z  h  z  y  d  z  y  d  y  j  d  f  r  j  j  h  t  r  s  q  t  x  y  x  '
    +'j  r  j  h  o  j  y  n  x  e  l  x  s  f  s  f  j  z  g  h  p  z  sz x  z  s  z  d  z  c  q  z  b  y  y  k  l  s  g  s  j  h  c  z  s  h  d  g  q  g  x  y  z  g  x  c  h  x  z  j  w  y  q  w  g  y  h  k  sy s  ea q  z  z  n  d  z  f  k  w  y  s  s  dt '
    +'c  l  z  s  t  s  y  m  c  d  h  j  x  x  y  w  e  y  x  c  z  a  y  d  m  p  x  m  d  s  x  y  b  s  q  m  j  m  z  j  m  t  jz q  l  p  j  y  q  z  c  gy q  h  y  j  h  h  h  qx x  h  l  h  d  l  d  j  qk c  fl d  w  b  s  x  f  z  z  y  y  s  c  h  '
    +'t  y  t  y  j  b  h  e  c  x  h  j  k  g  q  f  x  b  h  y  z  j  f  x  h  w  h  b  d  z  f  y  z  b  c  ha p  n  p  g  d  y  d  m  s  x  h  k  h  hm m  a  m  l  n  b  y  j  t  m  p  x  e  j  m  c  t  h  q  b  z  y  f  c  g  t  y  h  w  p  h  f  t  g  '
    +'z  z  e  z  s  b  z  e  g  p  b  m  d  s  k  f  t  y  c  m  h  b  l  lr h  g  p  z  jg x  z  j  g  z  j  y  x  z  s  b  b  q  s  c  z  z  l  z  cs c  s  t  p  g  x  m  j  s  f  d  c  c  z  j  z  d  j  x  s  y  b  z  l  f  c  j  s  a  z  f  g  s  z  l  '
    +'w  b  c  z  z  z  b  y  z  t  z  y  n  s  w  y  j  g  x  z  b  d  s  y  n  x  l  g  z  b  z  f  y  g  c  z  x  b  z  h  z  f  t  p  b  g  z  g  e  j  b  s  t  g  k  d  m  f  h  y  z  z  jg h  z  l  l  z  z  g  j  q  z  l  s  f  d  j  s  s  c  b  z  g  '
    +'p  d  l  f  z  f  z  s  z  y  z  y  z  s  y  g  c  x  s  n  x  x  c  h  c  z  x  t  z  z  l  j  f  z  g  q  s  q  q  x  c  j  q  c  c  c  c  d  jq c  d  s  z  z  y  q  j  c  c  g  x  z  t  d  l  g  s  c  x  z  s  y  j  j  q  t  c  c  l  qc d  q  z  t  '
    +'qc c  h  q  q  y  z  y  n  z  z  zj p  b  k  h  d  j  f  c  j  f  z  t  y  p  q  y  q  t  t  y  n  l  m  b  d  k  t  j  c  p  q  z  j  d  z  f  p  j  s  b  n  j  l  g  y  j  d  x  j  d  c  q  k  zs g  q  k  x  c  l  b  z  j  t  c  j  d  q  b  x  d  j  '
    +'j  j  s  t  c  x  n  x  b  x  q  m  s  l  y  j  c  x  n  t  j  q  w  w  c  j  j  n  j  j  l  l  l  h  j  c  w  q  t  b  z  q  q  c  z  c  z  p  z  z  d  z  y  d  d  c  y  z  d  z  c  c  j  g  t  j  f  z  d  p  r  n  t  c  t  j  d  c  xq t  q  z  d  t  '
    +'j  n  p  l  z  b  c  l  l  c  t  d  s  x  k  j  z  q  dz mp z  l  b  z  n  b  t  j  d  c  x  f  c  z  dc b  c  z  j  j  l  t  q  q  p  l  d  c  k  z  t  b  b  z  j  c  q  d  c  j  w  y  n  l  l  z  l  z  c  c  d  w  l  l  x  w  z  l  x  r  x  n  t  q  '
    +'j  c  z  x  k  j  l  sj g  d  n  q  t  d  d  g  l  n  l  a  j  j  t  n  n  y  n  k  q  l  l  d  z  n  t  d  n  y  cj zygg  j  w  y  x  d  x  f  r  s  q  s  t  c  d  e  n  q  m  r  r  q  z  h  h  q  h  d  l  d  a  z  f  k  a  p  b  g  g  p  z  rf e  b  '
    +'z  z  y  k  y  q  s  p  e  q  j  j  g  l  k  q  z  z  z  j  l  y  s  y  h  y  z  w  f  q  z  n  l  z  z  l  z  h  w  c  g  k  y  p  q  g  n  p  g  b  l  p  l  r  r  j  y  x  c  c  c  g  y  h  s  f  z  f  w  b  z  y  w  t  g  z  x  y  l  j  c  z  w  h  '
    +'n  zc j  z  p  l  f  f  l  g  s  k  h  y  j  d  e  y  x  h  l  p  l  l  l  l  cj zygg  x  d  r  z  e  l  r  h  g  k  l  z  z  y  h  z  l  y  q  s  z  z  j  z  q  l  j  z  f  l  n  b  h  g  w  l  c  z  c  f  j  w  s  p  y  x  z  l  z  l  x  g  c  c  bp '
    +'z  b  l  l  c  x  b  b  b  b  n  b  b  c  b  b  c  r  n  n  z  c  c  n  r  b  b  n  n  l  d  c  g  q  y  y  q  x  y  g  m  q  z  w  n  z  y  t  y  j  h  y  f  w  t  e  h  z  n  j  y  w  l  c  c  n  t  z  y  j  j  c  d  e  d  p  w  d  z  t  s  t  n  j  '
    +'h  t  y  m  b  j  n  y  j  z  l  x  t  s  sk t  p  h  n  d  j  x  x  b  y  x  q  t  z  q  d  d  t  j  t  d  y  z  t  g  w  s  c  s  z  q  f  l  s  h  l  n  z  b  c  j  b  h  d  l  y  z  j  y  c  k  w  t  y  d  y  l  b  n  y  d  s  d  s  y  c  c  t  y  '
    +'s  z  y  y  e  bp g  e  x  h  q  d  d  w  n  yw g  y  c  l  x  t  d  c  y  s  t  q  n  y  g  z  a  sk c  s  s  z  z  d  z  l  c  c  l  z  r  q  x  y  yw w  l  j  s  b  y  m  x  s  hx z  d  e  m  b  b  l  l  y  y  l  l  y  t  d  q  y  s  h  y  m  r  q  '
    +'n  k  f  k  b  f  x  n  n  s  b  y  c  h  xy b  w  j  y  h  t  q  b  p  b  s  b  w  d  z  y  l  k  g  z  s  k  y  g  h  q  z  j  xh h  x  j  x  g  n  l  j  k  z  l  y  y  c  d  x  l  f  w  f  g  h  l  j  g  j  y  b  x  b  l  y  b  x  q  p  q  g  n  t  '
    +'z  p  l  n  c  y  b  x  d  j  y  q  y  d  y  m  r  b  e  y  j  y  y  h  k  x  x  s  t  m  x  r  c  z  z  j  w  x  y  h  y  b  m  c  f  l  y  z  h  q  y  z  f  w  x  d  b  x  b  c  w  z  m  s  l  p  d  m  y  c  k  f  m  z  k  l  z  c  y  q  y  c  c  l  '
    +'h  x  f  z  l  y  d  q  z  p  z  y  g  y  j  y  z  m  dz x  t  z  f  n  n  y  t  t  q  t  zc h  g  s  fp c  d  m  l  c  c  y  t  z  x  j  c  y  t  jx m  k  s  l  p  z  h  y  s  n  w  l  l  y  t  p  z  c  t  z  c  c  k  t  x  d  h  x  x  t  q  c  y  fp '
    +'k  s  m  q  c  c  y  y  a  z  h  t  j  p  l  y  l  z  l  y  j  b  j  x  t  fp n  y  l  j  y  y  n  r  x  c  y  l  m  m  n  x  j  s  m  y  b  c  s  y  s  s  l  zc y  l  lx j  j  g  y  l  d  z  d  l  q  h  f  z  z  b  l  f  n  d  s  q  k  c  z  f  y  h  '
    +'h  g  q  m  j  d  s  x  y  c  t  t  x  n  q  n  j  p  y  y  bp f  c  j  t  y  y  f  b  n  x  e  j  d  g  y  q  b  j  r  c  n  f  y  y  q  p  g  h  y  j  s  y  z  n  g  r  h  t  k  n  l  n  n  d  z  n  t  s  m  g  k  l  b  y  g  b  p  y  s  z  b  y  dt '
    +'j  z  s  s  t  j  z  t  st x  z  b  h  b  s  c  s  b  z  c  z  p  t  q  f  z  ml qy f  l  y  p  y  b  b  j  g  s  z  m  n  x  d  j  m  t  s  y  s  k  k  b  j  t  x  h  j  c  e  g  b  s  m  jh y  j  z  c  s  t  m  l  j  y  x  r  c  z  q  s  c  x  x  q  '
    +'p  y  z  h  m  k  ydtx  x  x  j  c  l  j  y  r  m  y  y  g  a  d  y  s  k  q  l  n  a  d  h  r  s  k  q  x  z  x  z  t  c  g  g  z  d  l  m  l  w  x  y  b  w  s  y  c  td b  h  j  h  c  f  c  w  z  s  x  w  w  t  g  z  l  x  q  s  h  n  y  c  z  j  x  '
    +'e  m  p  l  s  r  c  g  l  t  n  z  n  t  l  z  j  c  y  j  g  d  t  c  l  g  l  b  l  l  q  p  j  m  z  p  a  p  x  y  z  l  ae k  t  k  d  w  c  z  z  b  n  c  c  t  d  q  q  z  q  y  j  g  m  c  d  x  l  t  g  c  s  z  l  m  l  h  b  g  l  k  z  n  '
    +'n  w  z  n  d  x  n  h  l  n  m  k  y  d  l  g  x  d  t  w  c  f  r  j  e  r  c  t  z  h  y  d  x  y  k  x  h  w  f  z  c  q  s  h  k  n  m  q  q  h  z  h  h  y  m  j  d  j  s  k  h  x  z  j  z  b  z  z  x  y  m  p  a  j  n  m  c  t  b  x  l  s  x  l  '
    +'z  y  n  w  r  t  s  q  g  s  c  b  p  t  b  s  g  z  w  y  h  t  l  k  s  s  s  w  gh z  z  l  y  y  t  n  x  j  g  m  j  r  n  s  n  n  n  n  l  s  k  zc t  x  g  x  l  s  a  m  m  l  b  w  l  d  q  h  y  l  a  k  q  c  q  c  t  m  y  c  f  j  b  s  '
    +'l  x  c  l  z  j  c  l  x  x  k  n  b  n  n  z  l  h  j  p  h  q  p  l  s  x  s  c  k  s  l  n  h  p  s  f  q  c  y  t  x  j  j  z  l  j  l  d  tcxz  j  j  z  d  l  y  d  j  n  t  p  t  n  n  d  s  k  j  f  s  l  j  h  y  l  z  q  q  z  l  b  t  h  y  '
    +'d  g  dc j  f  d  b  y  a  dz x  dy z  h  z  j  n  t  h  q  b  y  k  n  x  j  j  q  c  z  m  l  l  j  z  k  s  p  l  d  s  c  l  b  b  l  n  n  l  e  l  x  j  l  b  j  y  c  x  j  x  g  c  n  l  c  q  p  l  z  l  z  n  j  t  s  l  j  g  y  z  d  z  p  '
    +'l  t  q  c  s  s  f  d  m  n  y  c  x  g  b  t  j  d  c  z  n  b  g  b  q  y  q  j  w  g  k  f  h  t  n  bp y  q  z  q  g  b  k  p  b  b  y  z  m  t  j  dt y  t  b  l  s  qy m  b  s  x  ts b  n  p  d  x  k  l  e  m  y  y  c  j  y  n  z  dc t  l  d  y  '
    +'k  z  z  x  td d  x  h  q  s  h  ydtg  m  z  s  j  y  c  c  t  a  y  r  z  l  p  w  l  t  l  k  x  s  l  z  c  g  g  e  x  c  l  f  x  l  k  j  r  t  l  q  j  a  q  z  n  c  m  b  q  d  k  k  c  x  g  l  c  z  j  z  x  j  h  p  t  d  j  j  m  z  q  y  '
    +'k  q  s  e  c  q  z  d  s  h  h  a  d  m  l  z  f  m  m  z  b  g  n  t  j  n  n  l  gh b  y  j  b  r  b  t  m  l  b  y  j  d  z  x  l  cxtj  l  p  l  d  l  p  c  q  d  h  l  h  z  l  y  c  b  l  c  x  cz cz j  a  d  q  l  m  cz m  m  s  s  h  m  y  b  '
    +'h  b  n  k  k  b  h  r  s  x  x  j  m  x  m  d  z  n  n  p  k  l  b  b  r  h  g  g  hg f  c  h  g  m  n  k  l  l  t  s  y  y  y  c  q  l  c  s  k  y  mw y  ey h  y  w  x  n  x  q  y  w  b  a  w  y  k  q  l  ds n  n  t  n  dt k  h  q  c  g  d  q  k  t  '
    +'g  p  kh x  h  c  bp d  h  t  w  n  m  s  s  y  h  b  w  c  r  w  x  h  j  m  k  m  z  n  g  w  t  m  l  k  f  g  h  k  j  y  l  ds y  y  c  x  w  h  y  ye c  l  q  h  k  q  h  t  dt q  kh h  f  fy l  d  xs q  w  y  t  y  y  d  e  s  b  p  k  y  r  z  '
    +'p  j  f  y  y  z  j  c  e  q  d  z  z  dy l  ae t  t  pb b  f  j  l  l  c  x  d  l  m  jx s  d  x  e  g  w  g  s  j  q  x  c  f  b  s  sx s  z  p  d  y  z  c  x  z  n  y  x  p  p  z  y  d  l  y  j  c  c  p  l  ty x  l  n  x  y  z  y  r  s  c  y  y  y  '
    +'t  y  l  w  w  n  d  s  a  h  j  s  y  g  y  h  g  y  w  wk a  x  t  j  z  d  a  x  y  s  r  l  t  d  p  s  s  y  x  f  n  e  j  d  x  y  z  h  l  x  l  l  l  zc h  z  sz j  n  y  q  y  q  y  x  y  j  g  h  z  g  j  c  y  j  c  h  z  l  y  c  d  s  h  '
    +'h  s  g  c  z  y  j  s  c  l  l  n  x  z  j  j  y  y  x  n  f  s  m  w  f  p  y  l  l  y  l  l  a  b  m  d  d  hy w  z  x  j  m  c  x  z  t  z  p  m  l  q  c  h  s  f  w  z  y  n  c  t  l  n  d  yj w  l  s  l  x  h  y  m  m  y  l  m  b  w  w  k  y  x  '
    +'y  a  d  d  xs y  l  l  d  j  p  y  b  p  w  n  x  j  m  m  m  l  l  hs a  f  d  l  l  a  f  l  b  n  h  h  b  q  q  lj q  z  j  c  q  j  jl d  j  t  f  f  k  m  m  m  p  y  tm h  y  gj d  r  j  r  d  d  w  r  q  j  x  n  b  y  s  r  m  z  d  b  y  y  '
    +'t  bt j  h  p  y  m  y  j  t  j  x  a  a  h  g  g  d  q  t  mw y  s  t  qs x  k  b  t  z  b  k  j  l  x  r  b  ye q  q  h  x  m  j  j  b  d  j  n  t  g  t  b  x  p  g  b  k  t  l  g  q  x  j  j  j  c  d  h  x  q  d  w  j  l  w  r  f  m  j  g  w  q  h  '
    +'c  n  r  x  s  w  g  b  t  g  y  g  b  w  h  s  w  d  w  r  f  h  w  y  t  j  j  x  x  x  j  y  z  y  s  l  p  h  y  y  p  ya y  x  h  y  d  q  p  x  s  h  x  y  x  g  s  k  q  h  y  w  b  d  d  d  p  p  l  c  jg l  h  q  e  e  w  jx g  s  y  y  k  d  '
    +'p  p  l  f  j  t  h  k  j  l  t  c  y  j  h  h  jg t  t  p  l  t  z  z  c  d  l  y  h  q  k  c  j  q  y  s  t  e  e  y  h  k  y  z  y  x  x  y  y  s  d  d  j  k  l  l  p  y  m  q  y  h  q  g  x  q  h  cz r  h  b  x  p  l  l  n  q  y  d  q  h  x  s  x  '
    +'x  w  g  d  q  b  s  h  y  l  l  p  jg jx j  t  hg y  j  k  y  p  h  t  h  y  y  k  t  y  e  z  y  e  n  m  d  s  h  l  cz r  p  q  f  b  n  f  x  z  b  s  f  t  l  g  x  s  j  b  s  w  y  y  s  k  s  f  l  x  l  p  p  l  b  b  b  l  n  s  f  b  f  y  '
    +'z  b  s  j  s  s  y  l  p  b  b  f  f  f  f  sy s  c  j  d  s  t  j  s  x  t  r  y  j  c  y  f  f  s  y  z  y  z  b  j  t  l  c  t  s  b  s  d  h  r  t  j  j  b  y  t  c  x  y  y  e  y  l  y  c  b  n  e  b  j  d  s  y  s  y  h  g  s  j  z  b  x  b  y  '
    +'t  f  z  w  g  e  n  h  h  h  t  h  j  h  h  xt f  w  g  c  s  tx b  g  x  k  l  s  t  y  y  m  t  m  b  y  x  j  s  k  z  s  c  d  y  j  r  c  y  t  h  x  z  f  h  m  y  m  c  x  l  z  n  s  d  j  t  xt t  x  r  y  c  f  y  j  s  b  s  d  y  e  r  x  '
    +'h  l  j  x  b  b  d  e  y  n  j  g  h  x  g  c  k  g  s  c  y  m  b  l  x  j  m  s  z  n  s  k  g  x  f  b  n  b  b  t  h  f  j  y  a  f  x  w  x  f  b  x  m  y  fp h  d  td t  c  x  z  z  p  x  r  s  y  w  z  d  l  y  b  b  k  td y  q  w  q  j  b  z  '
    +'y  p  z  j  z  n  j  p  z  j  l  z  t  f  y  s  b  t  t  s  l  m  p  t  z  r  t  d  x  q  s  j  e  hx b  n  y  l  n  d  x  l  j  s  q  m  l  h  t  x  t  j  e  c  x  a  l  z  z  s  p  k  t  l  z  k  qj q  y  f  s  y  jg y  w  p  c  p  q  f  h  j  h  y  '
    +'t  q  x  z  k  r  s  g  t  k  s  q  c  z  l  p  t  x  c  d  y  y  z  s  s  l  z  s  l  x  l  z  m  a  c  pb c  q  b  z  y  x  h  b  s  x  l  z  d  l  t  z  t  j  t  y  l  z  j  y  y  t  b  z  y  p  l  t  x  j  s  j  x  h  l  b  m  y  td x  c  q  r  b  '
    +'l  z  s  s  f  j  z  z  t  n  j  y  td x  m  y  j  h  l  h  p  b  l  c  y  x  q  j  q  q  k  z  z  s  c  p  z  k  s  w  a  l  q  s  pb l  c  zc z  j  s  x  g  w  w  w  y  g  y  a  ts jx b  b  c  t  d  k  h  q  h  k  g  t  g  p  b  k  q  y  s  l  b  x  '
    +'b  b  c  k  b  m  l  l  n  d  z  s  t  b  k  l  g  g  q  k  q  l  z  bs k  k  t  f  x  r  m  d  k  b  f  t  p  z  f  r  t  p  p  m  f  e  r  q  n  x  g  j  p  z  s  s  t  l  b  z  t  p  s  z  q  z  s  j  d  h  l  j  q  l  z  b  p  m  s  m  m  s  x  l  '
    +'q  q  n  h  k  n  b  l  r  d  d  n  h  x  d  k  d  d  j  c  y  y  lg j  f  q  g  z  l  g  s  y  g  m  j  q  j  k  h  b  ptbm  x  y  x  l  y  t  q  w  l  w  j  c  p  b  m  j  x  c  y  z  y  d  r  j  b  h  t  d  j  y  e  q  s  h  t  m  g  s  f  y  p  l  '
    +'w  h  l  z  f  f  n  y  n  n  hg x  q  h  p  l  t  b  q  p  f  b  j  w  j  d  b  y  g  p  n  x  t  b  f  z  j  g  n  n  n  t  j  s  h  x  e  a  w  t  z  y  l  l  t  y  q  b  w  j  p  gx x  g  h  n  n  k  n  d  j  t  m  s  z  s  q  y  n  z  g  g  n  w  '
    +'q  t  f  h  c  l  s  s  g  m  n  n  n  n  y  n  z  q  qz x  n  c  j  d  q  g  z  d  l  f  n  y  k  l  j  c  j  l  l  z  l  m  z  z  n  n  n  n  s  s  h  t  h  x  j  l  z  j  b  b  h  q  j  w  w  y  c  r  d  h  l  y  q  q  j  b  e  y  f  sx j  hx t  h  '
    +'n  r  n  w  j  h  w  p  s  l  m  s  s  g  z  t  t  y  g  r  q  q  w  r  n  l  a  l  h  m  j  t  q  j  s  m  x  q  b  j  j  z  j  q  z  y  z  k  x  b  j  q  x  b  j  x  s  h  z  s  s  f  g  l  x  m  x  n  x  f  g  h  k  z  s  z  g  g  s  l  c  n  n  a  '
    +'r  j  x  h  n  l  l  l  m  z  x  e  l  g  l  x  y  d  j  y  t  l  f  b  k  b  p  n  l  y  z  f  b  b  h  p  t  gx j  k  w  e  t  z  h  k  j  j  x  z  x  x  g  l  l  j  l  s  t  g  s  h  j  j  y  qz l  q  z  f  k  c  g  n  n  d  j  s  s  z  f  d  b  c  '
    +'t  w  w  s  e  q  f  h  q  j  b  s  a  q  t  g  y  p  j  l  b  x  b  m  m  y  w  x  g  s  l  z  h  g  l  s  g  n  y  f  l  j  b  y  f  d  j  f  n  g  s  f  m  b  y  z  h  q  f  f  w  j  s  y  f  y  j  j  p  h  z  b  y  y  z  f  f  w  o  t  j  n  l  m  '
    +'f  t  w  l  b  z  g  y  z  q  x  cz d  j  y  g  z  y  y  r  y  z  y  n  y  z  w  e  g  a  z  y  h  j  j  l  z  r  t  h  l  r  m  g  r  j  x  z  c  l  n  n  n  l  j  j  y  h  t  b  w  j  y  b  x  x  b  x  j  j  t  j  t  e  e  k  hg w  s  l  n  n  l  b  '
    +'s  f  a  z  p  q  q  b  d  l  q  j  j  t  y  y  q  l  y  z  k  d  k  s  q  j  n  e  j  z  l  d  q  c  g  j  q  n  n  j  s  n  c  m  r  f  q  t  h  t  e  j  m  f  c  t  y  h  y  pb y  m  h  y  d  mw j  n  c  f  g  y  y  x  w  s  h  c  t  x  r  l  j  gh '
    +'j  z  h  z  c  yn y  y  j  lp t  k  t  t  n  t  m  j  l  z  c  l  z  z  a  y  y  o  c  z  l  r  l  b  s  z  y  w  j  y  t  s  j  y  h  b  y  s  h  f  j  l  y  k  j  x  x  t  m  z  y  y  l  t  x  x  y  p  s  l  q  y  j  z  y  z  y  y  p  n  h  m  y  m  '
    +'d  y  y  l  b  l  h  l  s  y  y  g  q  l  l  nd j  j  y  m  s  o  y  c  b  z  g  d  l  y  x  y  l  c  q  y  x  t  s  z  e  g  x  h  z  g  l  h  w  b  l  j  hg e  y  x  t  w  q  m  a  k  b  p  q  c  g  y  s  h  h  e  gh q  c  m  w  y  y  w  l  j  y  j  '
    +'h  y  y  z  l  l  j  j  y  l  h  z  y  h  m  g  s  l  j  l  j  x  c  j  j  y  c  l  y  c  j  b  c  p  z  jq z  jq m  m  w  l  c  jq l  n  q  l  j  jq j  l  x  y  j  m  l  s  z  l  j  q  l  y  c  m  m  g  c  f  m  m  f  p  q  q  m  f  x  l  q  m  cl f  '
    +'f  q  m  m  m  m  h  n  z  n  f  h  h  j  g  t  t  h  x  k  h  s  l  n  cl h  h  y  q  z  x  t  m  m  q  d  c  y  d  y  x  y  q  m  y  qj y  l  d  d  c  y  a  y  t  a  z  d  c  y  m  d  y  d  l  z  f  f  f  m  m  y  c  q  c  w  z  z  m  a  b  t  b  y  '
    +'c  t  d  m  n  d  z  gh g  d  f  t  y  p  c  g  q  y  t  t  s  s  f  f  w  b  d  t  t  q  s  s  y  s  t  w  n  j  h  j  y  t  s  x  x  y  l  b  y  yq h  h  w  h  x  g  z  x  w  z  n  n  qz z  j  zj j  j  qj j  c  c  c  h  y  k  x  b  z  s  z  c  n  j  '
    +'t  l  l  c  q  x  yk n  j  n  c  k  y  c  y  n  c  c  q  n  x  y  e  w  y  c  z  d  c  j  y  c  c  h  y  j  l  b  t  z  yk y  c  q  w  l  p  g  p  y  l  l  g  k  t  l  t  l  g  k  gjqq  b  gjqy  c  h  j  x  y  ';

{$ELSE}

  //第一段汉字声母表 (区:129-160 位:64-254[除127]) 	32*190=6080
  Part1 =
     'k  s  x  n  m  c  s  d  q  l  y  b  j  j  j  g  c  z  n  jd f  y  ya n  j  n  z  h  cs y  n  g  d  n  n  n  s  n  j  n  s  n  n  n  x  y  f  s  n  gq n  n  n  n  z  gjqg  l  l  y  j  n  s  y  s  s  g  y  q  y  t  n  h  j  y  y  d  l  d  w  j  n  w  b  '
    +'b  f  s  n  x  c  h  g  d  c  c  s  r  j  c  q  w  y  f  c  w  d  z  p  y  d  d  w  y  x  j  a  j  b  c  f  t  cz y  c  x  x  x  c  x  n  n  x  x  z  q  b  p  y  s  a  z  h  m  z  b  q  b  s  c  y  z  b  xs j  n  n  h  h  x  g  f  m  b  h  h  g  q  c  '
    +'x  s  t  gk l  y  g  y  m  x  a  l  e  l  c  c  x  z  r  j  j  z  n  n  j  j  c  z  c  c  l  j  s  t  s  t  b  n  h  t  y  x  x  k  k  w  y  f  l  h  j  x  s  p  x  m  x  x  t  c  n  l  s  h  x  j  b  tc f  y  b  yz x  c  c  z  b  z  y  c  l  w  l  c  '
    +'z  g  t  s  m  t  c  x  p  q  g  jl x  j  f  z  z  l  s  l  h  d  z  b  w  j  n  c  j  y  s  n  y  c  q  n  z  c  w  y  b  t  y  f  t  w  e  c  s  k  d  c  b  y  h  y  z  q  y  y  x  z  c  f  b  z  m  j  y  x  x  sc d  cz z  t  b  z  j  w  s  c  s  x  '
    +'y  r  r  y  g  m  d  t  h  j  x  s  q  j  c  c  s  b  x  y  y  t  s  w  f  b  j  z  n  z  l  y  z  z  p  s  cz y  z  z  s  q  q  h  z  q  y  d  x  l  b  p  j  l  l  m  q  x  q  y  d  z  s  q  j  t  z  p  l  c  g  q  d  c  w  z  f  h  c  t  d  jy f  x  '
    +'j  e  l  b  g  x  x  m  y  j  j  q  f  z  a  s  y  j  n  s  y  d  k  j  c  j  s  z  n  b  a  t  c  c  l  n  j  q  m  w  n  q  n  c  l  l  k  b  d  b  z  z  s  y  h  q  c  l  t  w  l  c  c  r  s  h  l  l  z  n  t  y  l  n  z  x  d  d  t  cz e  e  j  y  '
    +'s  sk q  f  m  x  d  bk l  k  w  n  l  ys n  t  n  j  y  m  r  y  m  c  j  g  j  m  z  g  x  y  k  y  m  s  m  z  k  n  f  x  m  t  g  h  p  f  m  x  j  s  m  s  g  j  qj d  g  y  a  l  c  m  z  c  s  d  j  l  x  d  f  f  j  c  n  f  n  n  f  f  k  g  '
    +'p  k  h  r  c  j  q  c  j  d  w  j  l  f  q  j  m  l  z  b  j  j  s  c  g  c  k  d  e  j  c  j  l  l  z  y  c  k  c  c  c  l  f  c  q  c  z  g  b  d  q  z  j  j  y  h  d  d  w  g  s  j  z  k  c  j  c  tz l  l  f  s  k  g  h  z  z  l  j  l  g  j  g  j  '
    +'j  t  j  j  j  z  c  c  m  l  z  y  j  k  x  z  y  z  m  l  j  k  y  w  x  m  k  j  l  k  j  h  m  c  l  y  k  j  q  l  b  l  k  m  d  x  w  y  x  w  s  l  l  p  s  j  q  jc q  x  y  q  f  j  t  j  d  m  x  x  l  l  c  r  q  y  j  b  n  n  g  g  n  p  '
    +'j  t  g  a  g  d  j  q  n  j  y  j  k  h  q  f  q  cz k  h  y  g  h  d  k  l  l  s  d  j  q  x  a  y  qo x  n  z  s  x  w  w  x  dc j  sl k  x  x  j  ay q  s  s  x  q  w  j  e  x  n  h  z  ac y  z  z  p  n  k  y  z  p  t  s  t  z  f  cs y  l  d  qj a  '
    +'g  y  l  ca l  l  y  y  y  s  l  q  q  n  l  d  z  cs s  cs a  d  s  g  b  r  s  z  j  s  j  y  r  c  g  q  c  n  h  m  m  x  z  c  y  o  x  y  c  q  j  p  e  h  w  n  c  d  oh g  j  l  g  m  q  w  j  h  c  s  t  n  e  s  q  p  w  m  r  d  s  z  z  l  '
    +'y  q  p  z  x  y  x  d  q  g  t  p  b  f  bf h  h  d  w  z  b  z  x  h  q  e  e  xd z  s  x  k  h  t  e  x  gk l  t  xh y  d  n  n  n  h  y  l  p  m  z  h  x  l  b  l  b  m  l  s  f  h  s  y  g  g  b  h  y  j  t  x  g  l  m  c  z  y  d  q  d  q  n  g  '
    +'d  n  l  n  g  z  w  y  q  e  q  a  j  n  y  t  l  x  a  n  z  d  h  x  x  t  w  w  j  h  q  q  d  w  z  b  c  h  y  q  z  l  x  q  g  z  g  l  y  d  q  tc n  t  a  d  y  z  z  c  k  y  z  y  m  h  y  h  h  j  z  w  s  x  h  d  n  l  y  s  c  q  y  dcs'
    +'p  s  y  w  y  n  x  j  c  h  t  y  h  q  s  y  h  x  m  c  h  rn g  j  p  w  q  z  l  w  jl s  n  s  g  n  n  b  a  l  c  z  m  t  j  c  j  k  t  s  n  x  l  x  h  h  g  o  x  z  c  p  d  m  h  g  t  y  n  j  x  hy m  r  l  x  j  k  x  h  fm q  x  c  '
    +'d  x  w  z  p  c  k  c  z  c  d  y  t  x  q  h  l  xs h  y  x  n  s  y  y  d  z  z  n  hy h  j  y  a  y  k  y  p  d  x  d  d  n  p  y  z  n  d  h  t  ho xh y  d  p  c  j  j  c  n  n  t  l  h  b  y  n  y  h  m  h  z  l  l  n  n  m  y  l  l  l  m  d  c  '
    +'p  p  x  m  x  d  k  y  c  y  d  l  t  x  c  h  h  z  n  xa c  l  c  c  l  y  l  z  s  x  n  j  z  z  l  n  n  n  h  y  n  t  k  y  j  p  y  c  h  e  g  w  t  g  j  r  g  t  g  y  h  h  l  g  c  w  y  qj k  p  y  y  e  tsgt  t  t  l  h  y  l  l  y  t  '
    +'t  y  l  n  y  y  q  q  zs d  q  n  n  n  m  j  z  x  t  q  f  k  d  f  b  x  d  j  j  d  j  p  t  q  z  w  t  j  y  f  l  q  g  x  b  l  z  f  h  z  a  n  p  n  j  k  l  c  c  y  h  d  z  f  g  y  d  g  c  y  x  s  n  h  d  n  d  n  a  x  x  b  p  b  '
    +'y  y  h  x  c  c  p  f  q  y  j  j  d  l  x  j  z  l  n  b  j  y  d  wy j  s  c  w  j  l  z  k  z  d  t  c  s  b  k  d  y  z  z  q  j  n  k  k  n  j  g  y  ew g  l  n  y  n  a  c  n  t  c  h  b  l  g  z  b  y  m  j  r  ea g  z  y  h  e  y  b  j  m  c  '
    +'t  a  f  z  j  z  h  g  c  j  n  l  g  h  d  w  x  j  j  k  y  y  k  s  s  m  w  c  t  q  z  l  p  b  z  d  t  w  c  x  z  a  g  y  k  w  x  l  n  l  c  p  b  c  l  l  o  q  m  m  z  s  l  b  c  d  z  k  d  c  z  j  g  q  j  d  c  y  td z  q  n  z  q  '
    +'s  s  b  p  k  d  f  q  d  z  z  s  dh d  t  dy m  f  h  t  d  y  c  n  a  q  q  k  y  p  d  j  y  x  t  l  j  hp d  r  q  x  l  m  k  y  d  r  n  l  k  l  y  t  w  hp l  l  r  l  l  c  x  y  l  b  n  n  z  y  m  z  z  h  k  h  x  k  s  m  z  s  y  g  '
    +'x  f  n  b  s  q  l  z  x  n  n  x  y  m  d  q  n  g  q  m  m  c  z  g  g  bt t  y  b  n  y  j  b  em d  p  j  x  d  n  h  k  s  b  h  f  d  s  b  a  w  p  w  a  j  l  d  y  j  s  f  h  b  l  c  d  q  j  n  c  x  f  j  h  d  f  j  x  w  z  p  k  z  y  '
    +'p  c  y  z  y  n  x  f  f  y  d  b  z  z  n  y  t  x  z  e  m  b  s  e  h  x  f  z  n  b  f  l  z  r  s  y  q  z  j  r  t  j  g  s  k  j  g  j  j  n  z  j  x  h  g  j  k  y  m  l  p  y  y  y  x  c  g  q  s  h  x  s  c  x  m  n  m  k  l  c  d  k  n  y  '
    +'p  s  z  h  z  p  t  x  w  y  w  x  y  y  s  l  j  c  q  x  z  d  l  a  e  l  m  c  p  j  c  l  x  s  q  h  f  w  r  t  f  n  t  n  q  j  j  q  d  x  h  w  l  y  cz c  f  l  ya k  y  y  n  l  d  x  n  h  y  y  n  c  j  t  y  w  t  r  m  d  r  q  n  w  '
    +'q  c  m  f  j  d  x  z  h  m  a  y  x  n  w  z  q  ts x  t  l  m  r  s  p  w  w  j  h  n  b  x  t  g  z  y  y  r  r  l  m  p  a  m  k  q  z  y  c  y  m  y  s  n  ty p  l  n  b  p  y  y  x  m  y  k  y  n  g  j  z  c  hr l  z  h  h  a  n  n  b  g  w  q  '
    +'t  z  m  x  x  m  l  l  h  g  d  z  x  n  h  x  h  r  y  c  j  m  f  f  x  y  w  c  f  s  s  q  l  hqxn  n  d  y  c  a  n  n  m  t  c  j  c  y  p  n  x  n  y  t  y  c  n  m  n  s  x  n  d  l  y  l  y  l  j  n  l  x  y  s  s  q  m  l  l  y  z  l  j  z  '
    +'x  s  t  n  c  s  m  c  j  y  n  z  l  x  b  n  n  y  l  r  q  t  r  y  y  j  z  g  h  s  y  t  s  c  q  g  x  z  b  s  h  m  k  cs z  y  q  h  z  j  n  b  h  q  s  n  j  n  z  y  b  k  n  j  q  h  s  n  s  w  x  k  h  j  y  y  b  q  c  b  f  l  d  p  '
    +'z  f  k  jq z  x  s  d  d  j  e  e  s  s  y  l  x  x  n  wy l  mp w  w  w  y  d  k  z  g  g  g  g  g  x  b  j  t  d  s  c  x  n  x  n  m  l  p  t  f  x  l  c  x  j  j  l  j  sz x  n  w  x  l  e  n  n  l  r  w  h  s  c  n  y  b  y  a  w  j  e  f  w  q  '
    +'q  j  z  z  y  j  g  x  b  l  t  q  f  t  p  y  k  q  p  t  l  c  n  n  x  t  x  n  h  k  l  e  f  d  l  l  e  g  q  y  m  s  a  w  h  m  l  j  t  k  y  n  l  y  j  e  y  b  q  f  n  l  y  x  h  d  s  c  t  g  h  x  y  wy l  k  x  q  c  t  n  h  j  l  '
    +'m  k  k  z  g  y  g  l  l  l  d  c  y  d  h  z  w  p  j  z  j  d  y  z  z  h  y  y  f  q  d  t  y  z  s  e  z  k  l  y  m  g  j  h  t  w  y  z  l  jh y  y  w  z  c  s  k  q  q  t  d  x  w  c  d  r  j  a  l  w  q  b  d  q  y  n  c  n  n  s  z  j  l  n  '
    +'c  d  c  d  td l  z  z  a  c  q  q  z  z  d  d  x  y  b  l  x  c  b  q  j  y  l  z  l  l  j  d  z  jq g  y  q  y  j  z  y  x  n  y  y  y  e  x  j  x  k  s  d  a  z  n  r  d  l  z  y  y  y  n  j  l  sx l  l  y  x  j  c  y  k  y  n  q  c  c  l  d  d  n  '
    +'y  y  y  n  y  g  h  j  c  l  n  n  n  q  z  z  y  j  x  j  n  f  z  d  n  f  p  z  h  d  d  m  f  m  y  y  p  q  j  r  s  s  q  z  s  q  d  g  p  z  s  w  d  s  q  d  h  z  x  w  y  b  p  n  g  p  dt m  j  t  h  z  s  b  g  z  m  b  j  c  z  h  b  b  '
    +'n  m  q  d  f  m  b  cd m  c  j  x  l  j  b  g  j  t  z  ga m  q  d  y  b  j  n  y  c  t  y  z  t  z  x  t  g  k  pm y  b  b  c  l  j  s  s  q  y  m  s  c  x  f  j  h  h  l  xs s  j  p  q  j  j  l  y  d  l  y  c  t  s  x  m  ca w  f  ga n  g  b  q  l  '
    +'l  l  n  y  x  t  y  l  t  x  d  p  h  n  h  f  j  y  z  y  e  s  e  s  d  h  w  t  q  b  s  j  t  f  d  c  z  y  q  s  x  j  d  z  j  qj b  b  s  dt j  n  f  b  k  n  b  x  dt k  qj h  m  k  w  j  j  l  h  h  y  y  y  h  h  z  y  y  c  d  y  p  c  z  '
    +'t  j  z  w  d  l  f  w  x  w  c  z  j  c  n  c  n  q  z  cz z  l  x  j  j  cs j  b  f  w  p  x  z  p  c  d  z  b  c  c  j  w  m  l  x  b  q  l  r  d  g  r  g  q  d  x  f  y  m  m  w  j  x  y  w  z  y  c  c  t  t  q  h  x  w  x  k  j  y  b  m  p  k  b  '
    +'n  t  h  j  z  d  t  y  f  x  b  y  x  z  b  h  x  p  x  x  t  n  n  n  c  x  l  h  k  m  z  x  t  d  h  g  x  x  s  h  q  h  c  y  x  g  l  c  x  q  y  p  d  h  n  m  y  j  m  y  y  y  j  zq l  h  q  t  b  q  x  m  y  h  c  w  l  l  c  y  l  n  ew w  '
    +'c  d  c  m  l  g  g  q  k  t  l  x  k  k  n  d  g  c  y  j  j  l  y  h  q  y  t  n  c  h  x  ew r  z  j  y  d  n  y  d  c  q  c  b  h  z  t  b  x  w  g  w  b  x  h  m  y  kq d  y  c  m  q  k  a  q  y  n  z  s  n  q  g  y  s  q  y  s  h  n  g  j  c  t  '
    +'x  k  z  y  c  s  s  b  k  x  h  y  y  l  s  t  y  x  t  y  m  n  c  p  m  g  c  c  c  c  c  m  z  t  a  s  g  q  z  j  l  o  s  j  y  l  s  t  m  q  z  q  d  z  l  j  q  q  y  p  l  z  y  c  z  t  c  q  q  p  b  c  j  c  l  p  k  h  c  y  y  x  x  d  '
    +'t  d  d  s  j  c  x  f  f  l  l  x  m  l  w  c  j  c  x  t  s  p  y  x  n  d  t  j  s  j  w  h  q  q  q  c  k  y  y  l  s  j  h  a  y  g  x  c  y  y  d  m  a  m  d  q  m  l  m  c  z  n  y  y  b  z  x  k  y  f  l  m  c  n  c  l  h  x  r  c  j  j  h  sz '
    +'l  n  m  t  j  g  z  y  g  j  d  d  c  j  sc k  z  cq g  j  y  y  xh z  x  xh q  h  h  e  s  l  m  d  s  y  y  y  q  l  fb n  z  z  d  y  k  h  s  c  c  g  w  t  c  x  q  j  y  a  x  z  z  d  y  k  b  j  w  h  y  q  z  z  y  w  n  p  p  b  a  z  n  b  '
    +'z  t  p  y  z  z  y  h  n  y  c  p  z  m  q  c  j  z  p  n  q  t  b  d  j  k  q  q  h  n  g  t  c  x  c  h  b  z  k  d  d  nl z  g  j  d  rn l  z  l  d  j  n  j  y  x  n  t  b  g  t  c  s  k  m  ln j  p  j  s  n  xj c  f  j  j  h  t  p  s  s  y  j  h  '
    +'b  j  y  z  l  s  t  b  n  w  q  s  m  m  f  d  w  j  y  z  c  t  b  n  z  w  xq c  s  l  q  g  d  g  q  s  q  l  y  z  l  g  y  h  s  d  c  b  t  z  k  sp j  n  n  n  s  p  n  j  z  j  j  y  n  h  p  w  s  n  s  z  z  s  h  b  y  h  y  z  xj k  j  t  '
    +'l  l  c  j  t  h  g  c  s  x  q  c  b  y  n  n  w  z  g  g  q  y  q  c  s  s  r  bp k  y  d  z  n  x  q  x  j  m  e  c  s  t  j  p  l  t  h  z  w  x  w  qc c  z  e  n  q  z  w  s  g  s  s  c  s  l  c  c  g  f  d  q  n  l  c  c  g  l  l  z  gx h  z  c  '
    +'t  h  n  j  g  y  a  z  c  m  s  k  s  s  t  z  cs k  b  j  y  g  q  j  p  l  d  x  r  k  z  y  x  c  n  h  h  d  n  l  z  h  j  j  c  d  ds b  c  j  x  b  f  z  t  p  q  d  n  j  t  q  wz j  y  k  l  z  z  p  c  j  d  s  j  j  k  d  x  g  a  j  y  n  '
    +'n  m  j  t  d  l  j  y  r  y  y  n  h  j  b  n  g  z  j  k  m  j  x  l  t  b  s  l  l  r  z  y  l  s  s  zc n  x  j  l  l  h  y  l  l  q  q  q  l  s  y  m  c  n  c  x  s  l  j  m  l  zc l  t  l  d  w  d  jg l  l  g  g  q  x  k  g  g  b  b  d  k  m  w  '
    +'g  c  c  g  x  j  c  s  d  y  b  x  d  n  d  d  y  k  y  q  d  x  d  y  x  hm a  a  d  z  s  l  q  x  x  j  n  q  z  yd l  b  l  x  x  x  qj qj j  z  l  b  y  l  w  s  j  j  y  j  t  d  y  q  qj z  z  z  z  q  l  z  c  d  z  x  h  p  y  n  q  p  l  f  '
    +'f  j  z  y  s  j  n  z  f  p  f  z  g  s  y  j  j  h  x  t  t  d  x  c  y  s  m  m  t  c  w  b  b  j  s  h  f  g  x  f  q  h  y  z  f  s  j  y  b  x  p  z  l  h  m  b  x  h  z  s  f  d  w  d  a  b  a  l  k  t  s  h  x  k  xk j  j  z  t  h  g  x  h  j  '
    +'x  k  z  x  s  z  z  h  h  w  t  z  z  z  s  y  x  q  z  y  a  w  l  c  w  x  n  g  y  y  h  f  m  h  y  d  w  q  m  n  j  y  c  y  s  p  j  k  g  w  c  q  h  y  l  g  m  z  x  h  m  c  n  z  h  h  x  c  l  d  j  p  l  x  y  j  k  d  y  y  l  t  t  x  '
    +'f  q  z  h  y  x  x  s  j  b  j  n  a  y  r  m  l  c  k  d  n  y  h  l  r  l  l  l  s  t  y  c  y  z  h  s  c  s  cz c  x  hk y  q  f  pb b  l  f  n  n  t  l  jz m  h  t  jq y  z  w  t  l  n  l  d  z  n  r  b  d  d  l  q  j  b  t  c  n  q  y  w  x  f  '
    +'z  r  z  d  m  c  y  g  g  y  n  n  n  s  m  s  y  f  d  x  c  s  h  x  cn c  s  y  y  j  m  b  a  f  y  f  n  h  y  d  x  y  e  z  y  n  d  n  l  s  n  x  t  x  g  bf m  y  y  s  s  n  b  y  n  l  h  b  fb z  d  c  y  f  m  z  s  s  n  y  z  z  d  g  '
    +'n  a  fb b  d  b  z  s  z  p  s  g  cz y  j  l  m  n  r  n  n  y  x  q  z  cy x  z  l  s  k  b  r  b  r  b  z  c  y  c  j  z  e  e  y  f  g  z  l  n  z  c  f  r  y  k  q  s  x  d  z  y  z  n  y  n  n  z  x  s  z  b  t  k  j  b  b  r  x  l  l  f  q  w  '
    +'j  g  c  q  y  l  p  z  d  x  t  z  a  b  d  h  z  r  b  j  h  w  n  j  t  j  x  l  k  c  f  s  l  d  c  y  j  k  z  c  w  j  l  b  n  n  t  z  l  l  z  l  q  b  l  c  q  q  c  c  d  f  p  pb h  c  z  l  y  y  g  h  d  g  w  x  f  c  z  q  y  y  y  q  '
    +'y  r  qx z  f  l  z  z  f  c  q  n  w  l  h  j  c  j  j  cs z  q  y  bp z  z  b  p  d  c  s  n  n  j  g  x  n  k  n  z  n  n  n  n  p  s  n  s  d  f  w  w  j  z  j  y  x  y  y  z  yj h  w  b  b  y  h  x  r  y  l  y  b  h  k  j  k  s  f  t  j  m  m  hk '
    +'h  t  y  y  s  x  y  z  p  d  j  y  sd y  w  m  d  j  j  r  h  ly n  g  n  t  y  n  n  n  j  y  z  f  n  m  g  j  t  y  s  y  z  m  s  j  y  j  h  gh q  m  y  r  s  z  y  t  r  t  z  s  s  k  x  g  q  k  s  p  t  g  x  dz n  j  g  q  q  m  x  g  z  t  '
    +'q  y  d  j  z  n  l  n  x  q  l  h  y  q  g  g  g  t  h  s  c  b  y  j  h  h  k  y  y  g  k  g  g  m  d  z  ly l  c  cz l  x  q  s  t  g  j  s  l  l  l  m  l  c  s  k  b  l  j  s  z  s  m  n  y  t  p  z  s  q  j  c  n  n  n  e  x  z  z  c  p  s  h  k  '
    +'z  z  x  x  d  f  m  w  rn l  l  q  x  r  f  z  l  y  s  t  c  t  m  s  j  t  h  j  n  t  n  r  tc z  f  q  r  h  c  gk l  l  g  n  n  n  n  j  d  n  n  n  t  s  j  l  n  y  s  s  z  x  c  g  j  z  y  p  f  h  d  j  s  b  d  c  z  ghkj  j  j  z  j  q  '
    +'d  y  b  s  s  ly y  t  t  m  q  n  b  h  j  q  mn n  y  g  jk y  e  q  y  q  m  z  gj c  j  k  p  d  c  n  m  y  z  g  q  l  l  s  l  n  c  l  m  h  o  l  zj g  d  y  ly f  z  s  l  n  c  n  z  l  y  l  z  c  j  e  s  h  n  y  l  l  n  x  n  j  x  l  '
    +'y  j  y  y  y  x  n  b  c  l  j  s  s  w  c  q  q  n  n  y  l  l  z  l  d  j  n  l  l  z  l  lj b  n  y  l  n  y  q  x  c  c  q  k  y  j  x  x  x  k  l  k  s  ea h  c  q  k  k  k  c  n  y  y  x  y  w  t  j  o  h  t  h  x  p  x  x  h  n  l  c  y  k  y  '
    +'c  h  b  b  j  q  c  s  c  s  z  s  s  l  c  y  l  g  d  z  m  m  y  s  x  j  q  q  s  q  y  f  y  c  h  y  n  y  w  cx t  j  s  y  d  c  h  c  d  d  j  l  b  d  j  j  z  q  y  s  kq kq y  x  h  q  j  o  h  d  y  x  g  m  a  z  p  c  n  n  p  n  j  s  '
    +'m  t  x  e  r  x  j  q  d  n  p  j  d  b  n  m  s  y  t  h  t  s  l  m  l  t  r  z  s  z  m  l  d  j  q  r  n  n  s  q  x  q  y  d  y  y  z  b  d  s  l  n  f  g  b  z  m  d  z  c  w  f  d  t  m  q  p  q  w  w  n  n  z  c  q  q  j  r  j  h  q  b  g  z  '
    +'f  j  h  n  x  x  y  y  d  y  h  h  r  m  w  c  t  b  z  p  z  z  l  z  f  m  z  t  l  m  y  f  t  s  k  j  j  y  j  z  h  b  z  c  y  g  h  p  c  z  s  c  s  j  y  s  x  f  j  g  d  y  z  y  h  z  c  w  h  j  s  e  x  f  z  z  y  w  k  n  a  y  m  l  '
    +'y  m  q  p  x  x  s  k  qj j  p  x  z  gh m  w  q  y  j  s  n  c  j  l  q  w  h  m  y  b  d  h  y  y  l  h  l  g  l  c  f  y  t  l  j  c  w  s  c  p  x  s  k  p  h  j  n  t  x  t  e  t  l  s  s  l  s  d  y  n  x  s  c  z  k  w  t  d  w  j  g  l  h  t  '
    +'q  d  j  g  g  h  z  p  h  c  q  f  c  l  j  l  z  p  t  y  n  l  m  j  l  l  q  y  n  ph h  y  l  q  qj z  y  m  b  y  w  r  f  y  k  j  s  s  n  p  x  y  r  h  j  n  q  t  f  wg d  w  r  c  g  y  h  m  m  y  y  h  x  m  z  h  n  k  s  e  l  q  q  m  '
    +'t  c  w  c  x  m  x  j  j  f  y  y  s  j  z  t  y  b  m  s  t  s  yc j  z  n  j  n  ts l  h  y  n  b  y  q  c  l  c  y  c  n  z  m  y  l  n  l  g  n  n  p  j  g  s  y  l  y  m  z  s  t  t  w  g  s  z  s  l  m  w  c  w  w  q  c  s  y  y  y  s  s  s  w  '
    +'p  c  w  c  m  h  h  h  x  d  z  yx j  g  s  j  h  y  g  ls c  x  y  x  h  b  b  z  j  k  s  s  m  a  l  hx y  c  f  y  g  m  q  y  j  y  j  x  j  l  l  j  g  c  z  g  q  j  c  c  t  o  t  y  x  m  t  s  h  l  w  l  q  f  z  k  p  z  c  x  kh j  y  cs '
    +'l  c  t  j  c  y  h  x  s  g  n  n  x  z  p  y  j  p  x  h  j  w  p  j  w  xs q  x  x  s  d  m  r  s  z  z  y  d  w  d  x  kh n  t  s  h  b  c  s  p  l  w  s  s  cd j  h  j  l  c  h  h  y  l  h  f  h  h  x  j  s  x  a  l  n  y  l  ms d  h  z  x  y  s  '
    +'x  l  m  z  y  kh c  l  d  y  h  l  m  d  y  s  p  j  t  q  z  n  g  j  f  s  j  h  c  t  s  d  s  z  l  m  s  s  m  n  y  yx m  j  q  j  zc w  t  y  y  d  c  h  j  l  j  a  w  b  g  q  x  b  k  f  n  b  j  dw l  l  l  l  y  y  l  s  j  y  d  w  h  x  '
    +'p  s  b  c  m  l  j  s  c  g  bp h  x  l  q  r  l  j  x  y  s  w  h  h  z  ls l  d  f  h  l  n  n  y  m  j  l  j  y  l  y  j  c  d  r  j  l  f  s  y  z  f  s  l  l  c  q  y  q  f  q  y  n  s  z  l  y  l  m  d  t  d  j  c  n  h  z  l  l  n  w  l  q  x  '
    +'y  g  y  y  g  x  x  n  h  z  z  x  c  z  q  z  f  n  w  p  y  p  k  p  k  p  m  l  g  x  g  x  b  d  x  z  z  k  z  f  b  d  d  l  z  p  t  y  td h  w  p  n  j  x  x  q  h  x  x  x  y  w  z  y  s  w  t  t  z  k  x  x  z  p  h  g  z  h  n  n  n  f  j  '
    +'c  p  tj z  t  h  l  s  h  c  r  h  n  s  l  x  x  j  x  x  z  qj dt x  q  n  n  y  j  l  k  h  c  x  c  wm h  j  f  y  c  f  p  q  c  x  q  x  y  j  y  g  p  y  n  s  c  s  n  n  n  n  c  h  k  z  y  h  f  l  x  j  b  y  z  w  t  t  x  nx n  c  y  j  '
    +'j  m  w  y  x  q  r  h  f  q  s  y  l  n  g  g  y  x  n  n  w  t  c  y  b  h  x  x  w  y  h  h  x  y  n  k  n  y  m  l  y  w  n  n  q  b  b  c  l  j  l  f  s  y  t  j  z  h  y  z  w  l  h  o  r  j  n  c  z  j  x  x  y  x  c  h  c  y  q  y  x  q  z  d  '
    +'d  s  j  f  s  y  l  t  s  f  x  l  m  t  y  j  m  n  n  y  y  x  l  t  c  x  q  c  l  h  z  l  w  y  x  z  h  n  n  l  r  k  x  j  c  dt y  h  l  b  r  l  m  r  d  l  a  x  k  s  n  l  l  j  l  y  x  x  l  y  n  rs y  l  c  t  g  n  c  m  t  l  z  l  '
    +'l  y  l  z  p  z  w  n  j  y  z  e  c  k  z  z  q  d  q  p  s  j  y  t  t  b  b  b  b  n  d  c  c  j  t  r  m  f  r  y  f  b  s  y  g  m  d  q  q  z  m  k  q  l  p  g  z  b  q  r  j  f  k  j  c  x  b  l  j  m  s  w  l  d  t  s  x  l  d  l  p  p  b  x  '
    +'c  w  k  c  b  b  c  c  j  k  b  h  y  h  h  z  q  z  y  p  j  n  s  t  l  p  n  s  y  y  d  b  ';

  //第二段汉字声母表 (区:170-253 位:64-160[除127], 区253位156-160留空)  	84*96-5=8059
  Part2 = 
     'g  x  h  hm z  j  h  s  t  t  l  m  x  x  b  z  y  s  z  j  h  f  y  bp q  b  y  l  j  q  k  y  z  z  z  y  l  n  z  h  f  w  y  y  t  y  xh j  b  y  y  s  b  m  s  m  y  d  y  s  h  n  y  z  c  h  m  j  m  c  a  h  c  b  b  b  h  b  l  y  t  y  s  x  '
    +'s  n  x  g  j  d  h  k  x  x  n  b  h  n  m  l  n  g  s  l  t  x  m  r  n  l  x  q  q  m  z  l  s  q  g  l  b  h  d  c  g  y  q  y  yc h  w  f  j  y  b  b  y  j  y  j  j  d  p  q  y  a  p  f  x  c  g  j  s  c  r  s  s  y  z  l  b  z  j  j  j  l  g  x  '
    +'z  y  x  y  x  s  q  k  x  b  x  x  g  c  x  p  n  d  y  w  e  c  td w  w  c  j  m  b  t  x  c  h  x  y  q  x  f  x  l  l  j  s  n  w  d  b  z  c  m  y  l  w  s  w  d  c  p  c  e  c  b  l  g  d  b  q  z  q  f  n  d  j  h  y  m  c  x  t  x  d  r  m  z  '
    +'w  r  h  x  c  j  z  y  l  q  d  y  h  l  n  r  s  y  w  w  j  j  y  m  t  l  l  l  t  q  c  j  z  b  t  c  k  z  c  y  q  y  s  q  a  l  m  y  n  w  w  d  n  z  x  q  d  l  l  q  s  g  j  f  j  l  j  n  j  a  z  d  j  g  t  h  h  s  s  t  n  y  j  f  '
    +'b  s  z  l  x  j  x  r  h  g  l  d  l  z  r  l  z  q  z  g  s  l  l  l  l  l  y  m  x  x  g  d  z  h  b  p  z  l  n  jh qw b  p  f  d  n  n  n  h  y  n  j  c  c  n  d  m  c  p  b  z  n  c  y  q  x  l  d  o  z  l  w  d  w  y  y  t  h  c  q  s  c  c  r  '
    +'s  s  l  n  z  f  p  td q  m  q  l  x  l  m  y  f  g  j  f  t  m  n  j  w  d  n  m  m  c  n  g  b  d  z  l  p  n  h  y  y  m  j  y  l  d  h  d  c  j  c  c  t  l  c  l  d  l  j  c  b  d  d  n  d  s  z  n  b  g  z  x  x  c  j  x  x  c  b  z  x  q  f  z  '
    +'f  j  d  n  c  t  j  w  h  d  c  m  j  x  t  m  x  s  p  d  s  y  p  z  g  m  l  j  t  a  c  b  m  d  e  y  x  s  z  jz y  f  y  s  t  g  w  h  k  y  j  x  g  y  c  l  ny d  z  s  c  y  z  s  s  d  l  l  j  f  l  q  l  l  x  f  d  y  h  x  g  g  n  y  '
    +'w  y  l  lj s  d  l  l  b  j  c  y  j  z  n  l  h  l  j  x  y  y  y  t  d  l  l  l  b  n  p  f  f  q  b  z  m  p  c  l  m  j  p  g  e  h  b  c  q  a  j  h  h  h  z  c  h  x  y  h  j  a  x  h  l  p  h  j  g  p  q  q  z  g  j  j  z  z  g  z  d  q  y  b  '
    +'z  h  h  b  w  y  f  f  q  d  l  z  l  j  x  j  p  a  l  x  z  d  a  g  l  g  m  q  w  t  x  x  f  m  m  s  y  p  f  m  x  s  y  z  y  s  h  d  z  k  j  s  m  m  z  z  d  d  y  z  c  b  m  l  t  d  y  e  m  x  z  y  m  z  m  m  s  h  h  c  c  j  e  w  '
    +'x  x  k  s  t  h  w  l  s  q  l  z  l  l  s  j  p  h  l  g  z  y  h  m  x  x  h  g  n  c  j  m  h  x  t  x  f  w  k  m  w  k  d  c  q  m  s  z  z  y  d  k  m  s  c  l  c  m  j  h  r  x  p  s  l  c  x  y  x  w  k  j  y  a  h  j  z  w  c  j  n  x  y  h  '
    +'m  m  bp m  l  g  x  m  h  l  m  l  g  m  x  t  k  z  l  y  s  c  j  s  h  y  z  j  z  h  bp d  q  z  w  t  k  q  d  j  l  f  m  n  k  z  j  f  e  z  s  p  n  m  q  y  k  t  e  m  z  n  p  l  p  l  b  p  n  k  k  q  z  k  e  j  l  wk a  k  y  p  n  h  '
    +'n  w  q  n  k  q  c  l  h  y  x  x  m  l  n  c  c  y  c  k  y  n  n  l  c  n  s  z  k  y  z  k  c  q  z  q  l  j  b  m  j  h  q  l  y  w  q  l  g  r  y  d  t  y  k  w  s  z  d  x  d  t  n  n  q  n  f  q  q  m  g  s  e  l  s  t  h  p  w  t  x  x  l  w  '
    +'y  d  l  n  q  z  c  q  a  p  l  l  k  q  c  y  l  b  q  q  c  z  c  n  j  s  l  z  j  x  d  d  p  z  q  d  l  j  x  z  q  j  y  z  h  h  z  l  k  c  j  q  d  wy p  p  y  p  q  a  k  j  y  r  p  z  b  n  m  x  kg l  l  z  l  l  f  q  p  y  l  l  l  m  '
    +'s  g  l  z  y  r  y  t  m  x  y  zq b  s  c  n  y  s  y  z  t  f  m  s  m  c  l  y  w  z  g  x  z  g  g  s  j  z  gj k  d  t  g  g  z  l  l  d  z  b  z  h  y  y  z  h  z  y  w  x  y  t  y  m  s  d  n  z  y  j  g  t  c  m  t  n  x  q  y  j  s  cs x  h  '
    +'s  l  n  nm d  l  y  t  z  l  r  x  t  r  z  c  y  x  n  q  f  h  y  z  j  z  y  k  b  p  n  l  y  p  b  l  n  z  z  j  t  t  z  h  h  y  z  z  r  d  z  n  f  k  z  s  k  g  j  t  t  y  l  l  g  z  z  b  jq z  k  l  p  b  z  y  n  n  y  x  b  j  f  b  '
    +'n  j  zc z  x  c  d  l  z  y  x  z  g  g  r  s  n  j  k  s  m  z  j  l  sw j  y  w  q  n  y  h  q  j  x  p  j  z  t  n  l  s  n  s  h  r  n  y  n  j  t  w  c  h  g  l  b  n  j  l  z  w  y  s  j  y  l  y  b  b  y  z  j  d  w  g  c  y  x  c  k  d  z  x  '
    +'s  n  n  w  w  y  q  y  y  t  l  t  d  j  l  x  w  k  c  j  n  k  c  c  p  z  c  q  q  d  z  z  q  c  sg f  qgsc  h  qj qfzgs sfzm  j  z  l  l  b  ygssq h  j  d  y  s  j  q  p  s  z  c  d  c  h  j  k  j  l  p  c  m  z  n  x  s  z  y  z  y  g  c  p  y  '
    +'d  s  x  z  n  g  q  m  b  c  f  l  n  f  f  g  f  s  m  n  q  n  l  g  q  c  y  y  b  g  j  s  r  j  h  z  l  d  c  f  t  l  l  j  g  j  h  t  x  z  c  s  z  z  t  jc g  g  k  y  n  x  p  l  z  b  b  g  c  g  y  j  z  c  z  s  z  l  l  j  f  z  g  q  '
    +'j  c  c  j  b  n  b  s  x  p  s  h  sy y  c  f  w  d  s  j  j  x  m  f  c  f  z  h  q  q  m  q  n  n  y  h  t  y  c  r  z  n  q  x  g  p  d  z  c  s  z  q  l  j  b  h  b  d  c  y  s  c  p  p  y  z  zk x  g  y  h  c  k  p  z  j  l  z  n  s  c  n  s  l  '
    +'l  x  b  m  s  d  l  d  f  j  m  k  d  q  s  l  x  l  s  z  n  p  q  p  g  j  d  l  y  d  s  k  g  q  l  n  l  k  y  y  h  z  t  t  m  c  n  q  t  z  z  f  s  z  q  k  t  l  l  j  t  y  y  l  l  n  l  l  q  y  z  q  l  b  d  z  l  s  l  y  y  z  yx f  '
    +'s  z  s  n  x  n  c  z  q  n  b  b  w  s  k  r  b  cs y  l  c  t  n  g  j  m  z  l  s  h  t  c  l  z  b  n  m  q  n  x  f  l  y  l  j  q  b  c  g  n  s  t  b  m  z  j  l  x  f  n  b  gy x  n  z  t  s  f  j  m  s  s  n  c  l  k  b  h  s  z  j  t  n  l  '
    +'z  d  n  t  l  m  j  g  z  j  y  j  c  z  x  y  hg y  hg w  r  w  q  n  z  t  n  f  j  s  c  p  y  s  h  z  j  f  y  r  d  j  f  c  j  z  b  f  z  q  c  h  z  x  f  x  s  b  z  q  l  z  s  g  y  f  t  z  d  c  s  z  x  z  j  b  j  p  s  z  k  j  r  h  '
    +'x  j  z  c  g  b  j  k  h  c  g  g  t  xj k  j  q  g  l  x  bp x  f  gj t  r  t  y  l  x  q  h  d  t  s  j  x  h  j  j  j  c  m  z  l  c  q  s  b  t  x  w  q  g  x  td x  x  h  x  f  t  s  d  k  f  j  h  z  y  j  f  j  x  n  z  l  d  l  l  l  c  q  s  '
    +'q  z  q  w  q  x  s  w  t  w  g  w  b  z  c  g  c  lg l  q  z  b  c  l  m  q  j  t  z  g  z  y  z  x  l  j  f  r  m  y  z  f  l  x  n  s  n  x  x  j  k  x  r  m  jq d  z  d  m  m  y  x  b  s  q  b  h  g  z  m  w  f  w  y  g  m  j  l  z  bp y  y  t  g  '
    +'z  y  c  c  d  j  y  z  x  s  n  g  n  y  j  y  z  n  b  g  p  z  j  c  q  s  y  x  s  x  r  t  f  y  z  g  r  h  z  t  x  s  z  z  t  h  c  b  f  c  l  s  y  x  z  l  z  q  m  z  l  m  p  l  x  z  j  s  s  s  b  y  s  m  q  h  x  xj n  x  r  x  h  q  '
    +'z  z  z  s  s  l  y  f  l  c  z  j  r  c  r  x  h  h  z  x  q  n  d  s  h  xj s  j  j  h  qz c  j  jz b  c  y  n  s  y  s  x  j  b  q  jk p  x  q  p  l  m  l  x  z  k  y  x  l  x  c  n  l  c  y  c  x  x  z  z  l  x  l  l  l  h  r  z  z  d  x  y  t  y  '
    +'x  c  x  f  f  q  b  p  h  n  g  y  g  z  t  n  q  w  y  l  t  l  g  w  n  m  g  z  j  m  m  g  t  j  f  ls g  z  y  a  f  s  m  l  bp f  c  w  b  j  t  l  j  m  z  l  pb j  j  l  m  t  y  y  y  f  b  y  g  q  z  g  y  z  y  r  q  q  h  x  y  y  q  x  '
    +'y  g  y  f  s  f  s  l  n  q  g  c  f  h  c  c  fp x  p  l  p  l  z  q  x  x  x  q  h  h  s  s  h  j  t  s  c  x  s  z  w  h  h  h  p  l  q  n  l  p  q  a  h  x  d  n  g  g  g  d  r  e  dz t  p  c  q  j  j  c  l  j  z  l  j  l  h  y  h  z  q  y  d  h  '
    +'z  n  z  c  z  y  w  t  e  y  d  n  h  s  l  j  b  d  g  w  x  p  c  n  n  t  y  c  k  l  l  w  k  l  l  c  s  s  t  k  n  z  d  n  n  j  t  t  l  z  s  s  z  y  q  k  c  g  d  h  y  r  r  y  c  q  b  p  p  y  r  w  j  p  x  x  k  d  f  b  b  q  t  z  '
    +'k  z  n  p  z  f  x  m  q  k  y  y  p  z  x  e  x  z  n  c  t  c  m  x  x  m  x  n  w  w  q  j  y  h  l  s  t  m  c  s  x  n  j  c  x  x  t  c  n  d  n  p  g  l  z  c  j  l  s  b  l  p  g  j  cz d  n  j  r  s  cd j  d  p  g  w  m  r  zd z  gl o  d  a  '
    +'n  s  y  z  r  t  w  j  j  d  b  c  q  w  s  t  s  z  y  l  j  p  x  l  n  c  l  g  p  c  j  f  z  l  j  y  l  x  c  t  n  l  c  g  x  t  f  z  j  h  c  r  x  s  f  g  d  k  n  j  c  j  lx l  t  q  c  b  x  nr h  z  x  b  x  k  lx y  l  h  z  l  q  z  '
    +'l  n  z  q  w  gw l  g  j  j  g  cx m  n  g  j  d  z  x  t  x  c  x  y  x  j  j  x  s  j  t  s  d  p  p  g  h  t  x  d  n  p  t  f  c  l  l  jx h  p  n  f  z  f  l  y  l  n  b  n  j  h  z  b  m  d  n  b  c  y  c  l  d  n  y  n  d  q  l  y  j  j  h  q  '
    +'l  l  c  s  j  p  y  y  c  l  t  t  j  p  y  c  m  g  y  x  z  h  s  z  t  w  q  w  r  f  z  h  j  g  a  c  m  r  h  y  y  y  n  p  yt d  l  r  p  y  zn n  b  b  q  y  x  h  z  d  d  n  h  m  s  g  b  w  f  z  d  g  n  n  x  g  l  r  c  c  y  x  z  l  '
    +'w  d  k  c  g  g  j  y  z  f  m  z  n  z  h  k  g  d  c  j  e  a  s  c  p  j  lp r  z  d  c  n  d  w  c  d  f  r  y  b  y  s  c  c  w  b  x  g  z  m  z  j  d  q  s  c  f  m  x  j  c  j  y  c  h  c  j  w  y  nc s  x  w  j  n  m  t  n  m  c  d  q  d  z  '
    +'l  l  w  n  k  p  z  g  g  l  c  c  z  m  l  b  q  j  q  d  j  j  z  t  g  h  j  a  w  b  z  j  l  t  t  d  h  h  c  c  h  f  l  s  y  y  t  qj w  c  n  t  j  n  n  m  n  n  d  q  k  x  wm y  y  f  l  x  n  z  w  c  x  pb m  a  y  h  g  j  w  z  z  j  '
    +'r  x  yxsa  q  j  f  l  l  p  h  h  h  y  t  z  j  m  h  s  g  z  q  w  b  w  j  d  y  j  q  x  hx s  l  z  y  y  m  y  s  z  g  n  q  n  w  n  s  y  s  c  s  y  z  n  h  q  y  l  j  x  c  x  t  l  h  z  q  z  p  c  y  c  y  p  p  n  x  f  y  r  c  m  '
    +'s  m  n  l  x  g  l  gh c  t  l  x  y  g  z  g  n  x  c  n  n  c  l  n  d  x  w  t  z  a  l  c  p  qo j  c  j  w  t  c  y  y  j  b  l  b  z  l  q  m  y  l  j  bp g  h  d  l  s  s  d  x  b  d  c  s  x  h  a  m  l  z  p  j  m  c  n  h  j  y  j  y  k  c  '
    +'h  s  k  q  m  cz z  l  w  j  qx s  m  o  c  d  r  l  y  q  c  h  j  m  y  b  y  l  r  e  t  f  j  f  r  f  k  s  y  x  f  d  w  t  s  x  x  y  w  s  j  s  l  y  x  s  d  x  y  y  x  h  a  h  w  j  c  x  l  m  l  j  c  sc q  l  k  y  d  t  t  z  s  lx '
    +'f  d  x  g  s  j  k  s  x  y  b  d  p  w  n  c  m  p  qj z  c  z  e  n  y  c  x  q  f  j  x  k  b  d  m  l  j  q  q  n  x  s  l  y  x  x  y  l  l  j  d  z  p  y  m  h  b  s  t  t  q  q  w  l  h  s  g  y  n  l  sz z  a  l  x  c  l  t  we r  r  q  l  s  '
    +'t  m  y  p  y  x  j  j  x  n  s  j  n  bn r  y  xl y  j  l  l  y  q  y  l  t  w  y  l  y  m  l  k  l  j  d  n  l  l  t  f  z  w  k  h  l  j  l  h  l  j  n  l  j  n  n  l  q  x  y  l  m  b  z  x  c  h  x  c  f  x  l  h  h  j  j  b  y  z  z  k  b  x  s  '
    +'d  q  d  n  d  z  sy y  g  z  x  f  e  p  c  q  w  y  y  j  q  r  qc y  j  h  q  q  z  m  w  f  f  h  f  r  b  n  t  p  c  j  l  f  z  g  p  b  x  d  b  b  r  t  g  y  g  c  h  m  f  l  y  x  l  y  p  j  s  y  w  m  q  g  g  n  x  l  q  j  ts c  b  h  '
    +'x  z  f  x  l  b  y  y  j  d  d  h  s  j  q  y  j  s  l  l  d  t  k  g  h  b  f  w  d  y  s  q  r  n  w  l  d  e  b  z  w  c  y  n  lz j  t  m  x  m  j  s  x  y  r  w  f  y  m  w  r  x  xh y  s  c  t  z  z  t  y  m  l  d  q  lx wg y  j  y  j  t  s  x  '
    +'x  w  l  b  r  j  w  x  h  y  b  h  y  d  n  h  h  x  m  y  w  y  t  y  c  n  n  n  s  d  l  w  z  c  q  p  y  j  l  q  y  j  w  j  w  z  z  m  g  l  c  l  m  x  n  m  z  s  q  t  z  p  j  q  b  l  g  x  j  z  f  l  j  j  y  c  j  rn x  w  c  x  s  n  '
    +'c  d  l  s  y  j  d  q  z  x  s  q  y  c  l  z  x  z  z  x  m  x  q  r  j  h  z  j  b  h  f  l  j  lz m  l  q  n  l  d  x  z  l  l  l  f  y  p  r  g  y  n  x  c  q  q  c  m  j  z  z  h  n  p  z  m  e  m  m  s  k  y  n  l  x  sz t  l  x  h  w  d  c  w  '
    +'d  z  g  y  y  f  f  j  z  d  r  z  x  zq x  r  z  j  c  h  r  t  l  b  y  z  b  s  j  h  x  z  y  p  bf d  f  g  z  z  r  y  n  n  k  x  c  q  y  n  b  n  y  c  k  r  j  c  m  j  e  r  z  g  y  n  n  n  k  n  s  j  k  l  j  s  g  c  l  j  y  b  z  s  '
    +'q  l  f  t  g  t  y  l  c  c  c  l  p  f  y  y  d  z  q  j  z  k  n  t  s  f  c  x  d  kh d  x  x  y  f  y  t  y  h  n  n  y  t  g  h  r  y  n  j  s  b  s  n  y  j  h  k  l  l  s  l  y  d  x  x  w  b  c  j  s  b  n  p  j  z  j  z  j  d  z  f  b  x  x  '
    +'b  r  m  l  a  d  h  c  s  n  c  l  d  s  t  b  l  p  z  n  s  w  s  b  x  b  c  l  l  x  x  l  z  d  n  z  s  j  y  n  y  x  f  n  n  b  h  j  j  jx g  b  y  g  j  p  m  m  m  m  s  s  zc l  j  m  t  l  y  z  j  x  s  w  x  t  y  l  e  d  q  p  j  m  '
    +'p  g  q  z  j  g  d  j  l  q  j  w  j  q  l  l  s  d  g  y  t  j  j  c  z  c  j  d  d  h  q  g  z  g  jx h  q  x  n  j  l  z  b  x  s  g  z  x  c  x  y  l  j  x  y  x  y  d  f  q  q  j  j  f  x  d  h  c  t  x  j  y  r  x  y  s  q  t  j  x  y  e  f  y  '
    +'y  s  s  y  x  j  x  n  c  y  z  x  f  x  c  s  x  s  z  xh y  y  s  c  h  s  h  x  z  z  z  g  z  z  g  f  j  d  l  d  y  n  p  z  g  s  j  a  z  t  y  q  z  p  b  x  cq b  d  z  t  z  c  z  y  x  x  y  h  h  s  c  j  s  h  c  g  g  q  h  j  h  g  x  '
    +'h  s  c  t  m  m  e  h  y  x  g  e  b  td c  l  z  k  k  w  j  z  r  s  l  e  k  e  t  d  b  c  y  h  q  q  s  a  y  x  c  j  x  w  w  g  s  p  h  j  sy z  sy d  n  c  s  j  k  q  c  x  s  w  x  f  c  t  y  n  y  dt p  c  c  c  z  j  q  t  c  w  j  q  '
    +'j  z  z  z  q  z  l  j  z  h  l  s  b  h  p  y  d  x  p  s  x  s  h  h  e  z  d  x  f  p  t  j  q  y  z  c  x  h  y  a  x  n  c  f  z  y  y  h  x  g  n  q  m  y  w  n  t  z  s  b  n  h  h  g  y  m  xs m  x  q  c  nx s  s  b  c  q  sy j  y  x  x  t  y  '
    +'y  h  y  b  c  l  m  m  s  z  m  j  z  z  l  l  c  o  g  x  z  a  a  j  z  y  h  j  m  c  h  h  c  x  z  s  x  s  d  z  n  l  e  y  j  j  z  j  b  h  z  w  jz z  sz q  t  z  p  s  x  z  z  d  s  q  j  j  l  n  y  a  z  h  h  y  y  s  r  n  q  z  t  h  '
    +'z  h  n  y  j  y  j  h  d  z  x  z  l  s  w  c  l  y  b  z  y  e  c  w  c  y  c  r  y  l  c  h  z  h  z  y  d  z  y  d  t  r  x  x  b  z  s  x  q  h  y  j  h  h  h  qx x  l  h  d  l  qk fl d  b  s  x  f  z  z  y  y  c  h  t  y  y  j  b  h  e  c  j  k  '
    +'g  q  f  x  b  z  j  f  x  h  w  h  d  z  f  y  ha p  n  p  g  d  y  m  s  h  k  h  a  l  n  b  y  j  t  m  x  e  j  c  t  h  q  b  z  y  f  c  g  t  y  h  w  p  h  f  t  g  z  z  e  z  s  b  z  e  g  p  b  m  d  s  k  f  t  y  c  m  h  b  l  lr h  g  '
    +'p  z  jg x  z  j  g  z  j  y  x  z  s  b  b  q  s  c  z  z  l  z  cs c  s  t  p  g  x  m  j  s  f  d  c  c  z  j  z  d  j  x  s  y  b  z  l  f  c  j  s  a  z  f  g  s  z  l  w  b  c  z  z  z  b  y  z  t  z  y  n  s  w  y  j  g  x  z  b  d  s  y  n  x  '
    +'l  g  z  y  b  z  f  y  x  x  x  c  c  x  t  z  l  s  q  q  x  c  j  q  c  c  c  d  d  s  z  z  j  c  c  g  x  t  d  g  s  c  x  z  s  y  j  j  q  c  c  l  d  q  z  t  qc c  h  q  q  y  z  y  n  z  k  h  j  c  f  y  p  q  t  y  n  l  m  k  c  q  z  j  '
    +'d  p  j  n  y  j  d  d  c  q  zs x  c  b  z  j  c  d  j  s  t  c  x  n  x  b  q  m  s  j  x  n  j  q  w  w  j  j  n  j  j  l  l  l  w  q  z  q  q  c  z  p  z  y  d  c  y  d  z  c  g  t  f  z  n  c  t  q  z  d  t  j  l  z  b  c  l  l  t  d  s  x  k  j  '
    +'z  q  dz mp z  l  z  n  t  j  x  z  b  z  j  l  t  q  q  l  d  k  t  b  b  j  q  d  c  j  w  y  n  z  l  z  c  d  w  l  l  x  w  l  r  x  n  t  q  j  c  z  x  k  j  l  d  n  t  d  g  l  n  l  a  j  j  n  n  y  n  k  q  l  l  d  z  n  t  d  n  y  cj zyg'
    +'g  j  y  x  d  x  f  r  s  q  s  t  c  d  e  n  q  m  r  r  q  z  h  h  q  h  d  l  d  a  z  f  k  a  p  b  g  g  p  z  rf e  b  z  z  y  k  y  q  s  p  e  q  j  j  g  l  k  q  z  z  z  j  l  y  s  y  h  y  z  w  f  q  z  n  l  z  z  l  z  h  w  c  g  '
    +'k  y  p  q  g  n  p  g  b  l  p  l  r  r  j  y  x  c  c  c  g  y  h  s  f  z  f  w  b  z  y  w  t  g  z  x  y  l  j  c  z  w  h  n  zc j  z  p  l  f  f  l  g  s  k  h  y  j  d  e  y  x  h  l  p  l  l  l  l  d  z  l  w  c  z  b  l  c  x  b  b  n  b  c  '
    +'b  b  n  n  z  c  n  r  b  n  n  c  g  y  y  q  z  w  n  z  t  h  t  e  n  c  n  z  c  w  t  t  n  h  t  y  n  y  j  l  x  d  j  y  q  d  j  y  z  w  l  s  h  n  z  b  c  j  b  y  z  j  y  c  t  d  y  n  d  s  y  c  t  s  z  y  y  e  g  d  w  n  y  c  '
    +'x  t  d  c  y  s  n  g  z  sk c  s  s  d  z  l  c  c  r  q  x  y  yw w  l  s  m  hx z  d  e  b  b  l  y  y  l  l  t  q  s  y  r  n  f  k  b  n  s  y  c  h  b  j  b  w  k  g  s  k  g  h  j  h  x  x  g  n  l  y  c  d  l  f  w  l  j  g  b  x  b  x  q  q  '
    +'n  z  p  l  n  y  b  x  j  y  q  y  m  r  b  y  y  y  h  k  x  x  s  t  m  x  r  c  z  z  j  w  x  y  h  y  m  c  f  l  z  h  q  f  w  x  d  b  x  b  c  w  z  m  l  d  m  y  c  k  f  m  k  l  z  c  y  q  y  c  c  h  x  z  l  y  y  z  m  dz t  z  n  n  '
    +'t  q  fp c  d  c  y  y  p  z  h  w  l  l  y  t  z  t  c  x  d  t  q  c  y  fp k  s  c  c  y  y  a  z  j  p  l  y  l  z  y  j  t  fp n  y  y  y  n  r  c  y  l  m  m  n  x  j  s  m  y  b  s  s  lx j  g  y  l  d  z  d  l  q  h  z  z  b  l  f  n  d  s  q  '
    +'k  c  z  f  y  h  h  g  q  m  j  d  s  x  y  c  t  t  x  n  q  n  j  p  y  y  bp f  c  j  t  y  y  f  b  n  x  e  j  d  g  y  q  b  j  r  c  n  f  y  y  q  p  g  h  y  j  s  y  z  n  g  r  h  t  k  n  l  n  n  d  z  n  t  s  m  g  k  l  b  y  g  b  p  '
    +'y  s  z  b  y  dt j  z  s  s  t  j  z  t  st x  z  b  h  b  s  c  s  b  z  c  z  p  t  q  f  z  ml qy f  l  y  p  y  b  b  j  g  s  z  m  n  x  d  j  m  t  s  y  s  k  k  b  t  x  h  j  c  e  g  b  s  m  jh y  j  z  c  s  t  m  l  j  y  x  r  c  z  s  '
    +'c  x  x  q  p  y  z  h  m  k  ydtx  x  x  j  c  l  j  y  r  m  y  y  g  a  d  y  s  k  q  n  a  d  h  r  s  k  q  x  z  x  z  t  c  g  g  z  d  l  m  l  w  x  y  b  s  y  c  td b  h  j  h  c  f  c  w  z  s  x  w  w  t  g  z  l  x  q  s  h  n  y  c  z  '
    +'j  x  e  m  p  l  s  r  c  g  l  t  n  z  n  t  l  z  j  c  y  j  g  d  t  c  l  g  l  b  l  l  q  p  j  m  z  p  a  p  x  y  z  l  ae k  t  k  d  w  c  z  z  b  n  c  c  t  d  q  q  z  q  y  j  g  m  c  d  x  l  t  g  c  s  z  l  m  l  h  b  g  l  k  '
    +'z  n  n  w  n  d  x  n  h  l  n  m  k  y  d  l  g  x  d  t  w  c  f  r  j  e  r  c  t  z  h  y  d  x  y  k  x  h  w  f  z  c  q  s  h  k  n  q  q  h  z  h  h  y  m  j  d  j  s  k  h  x  z  j  z  b  z  z  x  y  m  p  a  j  n  m  c  t  b  x  l  s  x  z  '
    +'y  n  w  r  t  s  q  g  s  c  b  p  t  b  s  g  z  w  y  h  t  l  k  s  s  s  w  gh z  z  l  y  y  t  n  x  j  g  m  j  r  n  s  n  n  n  n  l  s  k  zc t  x  g  x  l  s  m  m  l  b  w  l  d  q  h  y  l  k  q  c  q  c  t  m  y  c  f  j  b  s  l  x  c  '
    +'l  z  j  c  l  x  x  k  n  b  n  n  z  l  h  j  p  h  q  p  l  s  x  s  c  k  s  l  n  h  p  s  f  q  c  y  t  x  j  j  z  l  j  l  d  tcxz  j  j  z  d  l  y  d  j  n  t  p  t  n  n  d  s  k  j  f  s  l  j  h  y  l  z  q  q  z  l  b  t  h  y  d  g  dc '
    +'j  f  d  y  a  dz x  dy z  h  z  j  n  t  h  q  b  y  k  n  x  j  j  q  c  z  m  l  l  j  z  k  s  p  l  d  s  c  l  b  b  l  n  n  l  e  l  j  l  b  j  y  c  x  j  x  g  c  n  l  c  q  p  l  z  l  z  n  j  t  s  l  s  y  x  b  x  m  y  z  x  w  q  c  '
    +'z  y  h  z  b  t  r  b  l  h  l  c  cz cz j  a  d  q  l  m  m  m  s  s  h  m  y  b  h  b  n  k  k  b  h  r  s  x  x  j  m  x  m  d  z  n  n  p  k  l  b  b  r  h  g  g  hg f  c  h  g  m  n  k  l  l  t  s  y  y  y  c  q  l  c  s  k  y  mw y  ey h  y  w  '
    +'x  n  x  q  y  w  b  a  w  y  k  q  l  ds n  n  t  n  dt k  h  q  c  g  d  q  k  t  g  p  kh x  h  c  bp d  h  t  w  n  h  k  ds h  dt h  l  xs y  t  y  y  d  e  s  p  k  y  z  p  c  e  q  d  l  t  b  d  s  d  e  g  w  j  q  x  c  f  s  sx s  p  z  x  '
    +'z  p  y  d  l  y  j  c  l  n  x  y  z  y  r  s  c  y  y  y  t  w  n  d  a  j  g  y  h  y  w  t  j  d  a  x  y  s  r  l  t  d  p  s  y  x  f  n  e  j  d  y  z  l  x  l  l  h  z  y  q  y  h  g  j  h  z  y  c  s  h  h  s  g  c  z  y  j  s  l  l  n  x  z  '
    +'j  j  n  f  s  m  f  p  y  l  y  a  m  d  d  hy z  m  c  z  l  q  c  w  z  y  c  t  n  d  yj w  l  l  h  y  m  m  y  l  m  b  w  w  k  x  y  d  d  y  l  d  j  p  y  w  n  x  j  m  m  l  l  hs a  f  d  l  l  a  f  l  b  n  h  h  b  q  q  z  c  q  jl j  '
    +'t  f  m  p  tm h  y  d  r  j  r  d  d  w  r  q  n  y  s  r  m  z  b  y  y  t  bt j  h  p  m  y  j  t  j  a  h  g  g  t  y  s  t  x  k  b  t  z  b  k  l  x  b  ye q  q  h  x  m  j  d  n  t  t  x  p  g  b  k  t  l  g  q  x  j  j  j  c  d  h  x  q  d  w  '
    +'j  l  w  r  f  m  j  g  w  q  h  c  n  r  x  s  w  g  b  t  g  y  g  b  w  h  s  w  d  f  j  x  x  x  j  z  l  p  h  y  y  p  ya y  x  h  y  d  q  p  x  s  h  x  y  x  g  s  k  q  h  y  w  b  d  d  d  p  p  l  c  jg l  h  q  e  e  w  jx g  s  y  y  k  '
    +'d  p  p  l  f  j  t  h  k  j  l  t  c  y  j  h  h  jg t  t  p  l  t  z  z  c  d  l  y  h  q  k  c  j  q  y  s  t  e  e  y  h  k  y  z  y  x  x  y  y  s  d  d  j  k  l  l  p  y  m  q  y  h  q  g  x  q  h  cz r  h  b  x  p  l  l  n  q  t  j  h  y  y  l  '
    +'f  b  n  f  x  z  b  s  f  t  l  g  x  s  j  b  s  w  y  y  s  k  s  f  l  x  l  p  p  l  b  b  b  l  n  s  f  b  y  z  s  y  l  f  f  f  s  c  j  d  s  t  j  t  r  y  j  c  y  f  f  s  y  z  y  z  b  j  t  l  c  t  s  b  s  d  h  r  t  j  j  b  y  t  '
    +'c  x  y  y  e  l  y  b  n  e  b  j  d  s  y  s  y  h  g  s  j  z  b  x  b  y  t  f  z  w  g  e  n  h  h  h  h  j  h  h  xt f  w  g  c  s  tx b  g  x  k  l  s  t  y  y  m  t  m  b  y  x  j  s  k  z  s  c  d  y  j  r  c  h  x  z  f  h  m  y  m  c  x  l  '
    +'z  n  d  t  d  h  l  x  d  j  g  g  y  b  f  b  n  b  b  t  h  f  j  y  a  x  w  f  b  x  m  y  fp h  d  td t  c  x  z  z  p  x  r  s  y  w  z  d  l  y  b  b  k  td y  q  w  q  j  b  z  y  p  z  j  z  n  j  p  z  j  l  z  t  f  y  s  b  t  t  s  l  m  '
    +'p  t  z  r  t  d  x  q  s  j  e  hx b  n  y  l  n  d  x  l  j  s  q  m  l  h  t  x  t  j  e  c  x  a  l  z  z  s  p  k  t  l  z  k  qj q  y  f  s  y  jg y  w  p  c  p  q  f  h  j  h  y  t  q  x  z  k  r  s  g  t  k  s  q  c  z  l  p  t  x  c  d  y  y  '
    +'z  s  s  l  z  s  l  x  l  z  m  a  c  pb c  q  b  z  y  x  h  b  s  x  l  z  d  l  t  z  t  j  t  y  l  z  j  y  y  t  b  z  y  p  l  t  x  j  s  j  x  h  l  b  r  y  b  x  q  z  s  k  s  w  w  w  y  g  y  a  b  b  c  t  q  k  t  g  p  y  s  b  x  b  '
    +'b  c  m  l  l  n  z  t  b  k  l  g  q  k  q  l  z  k  t  f  x  r  d  k  b  f  p  f  t  p  p  m  f  e  r  q  n  g  p  z  s  s  t  l  b  t  p  s  z  z  s  j  d  h  l  q  l  z  p  m  s  m  m  s  x  l  q  q  n  k  n  b  r  d  d  n  h  x  d  k  d  d  j  y  '
    +'y  j  f  q  g  z  l  g  s  m  j  q  j  x  y  t  q  l  j  p  b  j  x  c  y  z  y  d  r  j  b  h  t  d  j  y  e  q  s  h  t  m  g  s  f  y  p  l  w  h  l  z  f  f  n  y  n  n  hg x  q  h  p  l  t  b  q  p  f  b  j  w  j  d  b  y  g  p  n  x  t  b  f  z  '
    +'j  g  n  n  n  t  j  s  h  x  e  a  w  t  z  y  l  l  t  y  q  b  w  j  p  gx x  g  h  n  n  k  n  d  j  t  m  s  z  s  q  y  n  z  g  g  n  w  q  t  f  h  c  l  s  s  g  m  n  n  n  n  y  n  z  q  qz x  n  c  j  d  q  g  z  d  l  f  n  y  k  l  j  c  '
    +'j  l  l  z  l  m  z  z  n  n  n  n  s  s  h  t  h  x  j  l  z  j  b  b  h  q  j  w  w  y  c  r  d  h  l  y  q  q  j  b  e  y  f  sx j  hx t  h  n  r  n  w  j  h  w  p  s  l  m  s  s  g  z  t  t  y  g  r  q  q  w  r  n  l  a  l  h  m  j  t  q  j  s  m  '
    +'x  q  b  j  j  z  j  q  z  y  z  k  x  b  j  q  x  b  j  x  s  h  z  s  s  f  g  l  x  m  x  n  x  f  g  h  k  z  s  z  g  g  s  l  c  n  n  a  r  j  x  h  n  l  l  l  m  z  x  e  l  g  l  x  d  j  t  b  k  y  z  b  p  k  t  z  h  k  z  j  y  q  s  b  '
    +'c  t  w  w  q  j  p  j  w  x  h  g  s  g  n  y  f  l  j  b  y  f  d  j  f  n  g  s  f  m  b  y  z  h  q  f  f  w  j  s  y  f  y  j  j  p  h  z  b  y  y  z  f  f  w  o  t  j  n  l  m  f  t  w  l  b  z  g  y  z  q  x  cz d  j  y  g  z  y  y  r  y  z  y  '
    +'n  y  z  w  e  g  a  z  y  h  j  j  l  z  r  t  h  l  r  m  g  r  j  x  z  c  l  n  n  n  l  j  j  y  h  t  b  w  j  y  b  x  x  b  x  j  j  t  j  t  e  e  k  hg w  s  l  n  n  l  b  s  f  a  z  p  q  q  b  d  l  q  j  j  t  y  y  q  l  y  z  k  d  k  '
    +'s  q  j  n  e  j  z  l  d  q  c  g  j  q  n  n  j  s  n  c  m  r  f  q  t  h  t  e  j  m  f  c  t  y  h  y  pb y  m  h  y  d  mw j  n  c  f  g  y  y  x  w  s  h  c  t  x  r  l  j  gh j  z  h  z  c  yn y  y  j  lp t  k  t  t  n  t  m  j  l  z  c  l  z  '
    +'z  a  y  y  o  c  z  l  r  l  b  s  z  y  w  j  y  t  s  j  y  h  b  y  s  h  f  j  l  y  k  j  x  x  t  m  z  y  y  l  t  x  x  y  p  s  l  q  y  j  z  y  z  y  y  p  n  h  m  y  m  d  y  y  l  b  l  h  l  s  y  y  g  q  l  l  s  c  x  l  x  h  z  w  '
    +'k  q  g  y  s  h  q  y  w  l  j  y  y  h  z  m  s  l  j  l  j  x  c  j  j  y  y  c  b  c  p  z  jq m  w  l  c  jq l  n  j  jq j  l  x  y  j  m  l  z  j  q  l  y  c  m  g  c  f  m  m  f  p  q  q  m  f  x  l  m  cl f  q  m  m  n  z  n  f  h  j  g  t  t  '
    +'x  k  h  cl h  y  z  x  t  m  q  d  y  m  y  y  d  y  a  y  t  z  d  c  y  m  d  y  d  l  f  m  y  c  q  c  w  z  z  m  a  b  t  b  c  m  d  z  g  d  f  y  c  g  q  y  t  t  s  f  w  b  d  t  t  q  s  s  s  t  n  j  h  j  y  t  s  x  l  y  h  w  h  x  '
    +'g  x  w  z  n  n  qz z  j  zj j  j  c  c  c  h  y  k  x  b  z  s  z  c  n  j  t  l  l  c  q  x  yk n  j  n  c  k  y  c  y  n  c  c  q  n  x  y  e  w  y  c  z  d  c  j  y  h  y  l  p  g  p  y  l  l  g  k  t  l  t  gjqq  b  c  h  j  x  y  ';

  //第三段汉字声母表 (区:176-247 位:161-254)         72*94 =6768
  Part3 = 
     'a  ae a  a  a  a  a  a  ay a  a  ay a  a  a  a  a  a  a  a  a  a  a  a  a  a  a  aw a  a  a  a  a  a  a  a  b  b  bp b  b  b  b  b  b  b  b  b  b  bp b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  bp bp b  b  b  b  b  b  b  b  b  b  b  b  b  b  bp b  b  '
    +'bp b  b  b  b  b  b  b  b  b  b  b  b  bp b  b  b  b  bp b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  bp b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  bp b  b  b  b  b  b  b  b  bp bp b  b  b  b  b  bp b  '
    +'b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  b  bp b  b  b  b  b  b  b  b  b  b  b  bp b  b  b  b  b  b  bf b  b  b  b  b  c  c  c  c  c  c  c  c  c  c  c  c  c  cs ct c  c  c  c  c  c  c  c  cz c  c  c  c  c  cs c  cz '
    +'c  c  c  c  c  c  cz c  cz c  c  c  c  c  c  c  c  c  c  cs c  c  c  c  c  c  c  cz c  c  c  c  c  cz c  c  chact c  c  c  c  c  c  cz cz c  c  c  c  cj c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  cs c  c  cd c  c  c  c  c  cj c  c  '
    +'cs c  c  c  c  c  c  c  c  c  c  c  c  c  c  ch c  c  c  c  c  c  c  c  c  cq c  c  c  cx c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  cz c  c  c  c  c  cz c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  '
    +'c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  c  cz c  c  c  c  d  d  d  d  d  d  da d  d  d  d  d  d  d  d  d  d  d  d  d  d  dcsd  ds d  d  d  d  d  d  d  dt d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  '
    +'d  d  d  d  d  d  d  d  d  dt d  d  d  d  d  d  d  dz d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  dt d  d  d  d  d  d  d  d  d  d  d  d  d  d  dt d  d  d  d  d  d  d  dz d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  dt dt d  d  d  d  d  d  '
    +'d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  d  dz d  d  d  d  d  dc d  d  dt d  ds d  d  d  d  d  d  d  d  d  dt d  d  dh ey e  e  e  e  e  e  ew e  e  e  e  e  e  e  er e  e  e  e  e  e  f  f  f  f  f  f  f  f  f  f  fp f  f  '
    +'f  f  fp f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  fp f  f  f  f  f  fp f  f  f  f  f  fb f  f  f  f  f  f  f  f  f  f  f  '
    +'f  f  f  f  f  f  f  fp f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  f  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  gh g  g  g  g  g  g  g  g  g  g  gj g  g  gh g  g  g  g  g  gj '
    +'g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  gh g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  gh g  g  g  gy g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  g  ga g  g  g  g  gh g  gjqg  g  g  g  '
    +'g  g  gj g  g  g  g  g  g  g  g  g  g  g  g  hk h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  hb h  h  h  h  h  h  h  h  h  h  h  hkah  h  h  h  h  h  h  hg h  hm h  h  h  h  h  h  h  hm h  h  h  h  h  h  h  h  h  h  h  '
    +'h  h  hj h  h  h  h  hg h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  hx h  h  h  h  h  hy h  h  hg h  h  h  h  h  h  h  h  hp h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  '
    +'h  h  h  h  h  h  hk h  h  h  h  h  hx h  h  h  h  h  h  h  h  h  h  h  h  h  h  h  j  j  j  j  j  jq j  j  j  j  j  j  j  j  j  j  jq j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  jz j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  jg '
    +'j  jg j  j  j  jg j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  jk j  j  j  jx j  j  j  j  j  j  j  j  j  j  j  j  j  jq j  j  j  j  j  j  j  j  j  jx jq j  j  j  j  j  j  j  j  j  j  j  j  j  jy j  j  '
    +'j  j  jz j  jc j  jx j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  jx j  j  j  jg j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  jg j  j  j  jy jg j  j  j  j  j  j  j  j  j  j  j  '
    +'j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  j  jz j  j  j  j  j  j  j  j  j  j  j  j  jg j  j  j  j  j  j  j  j  jq j  j  j  j  j  j  j  j  j  j  j  jy j  j  j  j  j  jz j  jx j  j  k  kg kq kglk  k  kj k  k  k  k  k  k  '
    +'k  k  k  k  k  kg k  k  k  k  k  k  k  k  k  k  k  k  k  k  kq kh k  k  k  k  k  k  k  k  k  k  k  kh k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  k  kg k  k  kh k  k  k  k  kg '
    +'k  k  k  l  l  l  lz lx l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  lm l  l  l  l  ly l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  '
    +'l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  '
    +'l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  ls l  l  l  l  l  l  l  l  l  l  l  l  l  l  l  lg l  l  l  l  l  l  l  l  l  l  l  l  l  m  '
    +'m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  mw m  m  m  m  mw m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  my m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  mblm  mb m  m  m  m  '
    +'m  m  m  mw m  mw m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  m  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  '
    +'n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  nd ns n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  n  nl n  n  n  nr n  n  ny n  n  n  n  oewo  o  o  o  o  o  o  p  p  p  p  p  pb p  p  p  p  p  p  p  '
    +'p  p  p  p  p  p  p  p  p  pb p  p  p  p  pb pb p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  py p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  p  '
    +'pb p  p  p  p  p  ptbp  p  p  p  p  p  p  p  p  p  pb p  p  p  p  p  pb pb qj q  qx q  q  q  q  qx q  q  qj q  qj q  qx q  q  qj q  q  q  q  q  q  q  q  q  qx q  q  q  q  q  q  qx q  q  q  qx q  q  q  qy q  q  q  q  q  qg q  q  q  q  q  q  qj q  q  qk '
    +'q  q  q  q  q  q  q  q  qj q  qc q  q  q  q  qy q  q  q  qs q  q  q  q  q  q  qj qj q  q  q  q  q  q  q  q  q  q  q  q  qs q  q  q  q  q  q  q  q  q  q  q  q  q  q  q  q  q  q  q  q  q  q  q  q  qo qj q  q  q  q  q  q  q  q  qc q  qj q  q  q  q  q  q  '
    +'q  q  qx q  q  qg q  q  q  q  q  q  q  q  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  r  s  s  s  s  s  s  s  s  s  s  s  '
    +'s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  sc s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  sc s  s  s  s  s  sb s  s  s  s  s  s  s  sy s  s  s  s  sy s  s  s  s  s  s  s  s  sj s  sczs  s  sc s  s  s  s  '
    +'s  s  s  s  s  s  s  s  sx sc s  s  s  s  s  s  s  s  s  s  s  s  sd s  s  s  sy s  s  sz s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  sk s  s  s  s  sz s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  '
    +'s  s  s  s  s  s  s  sz sz s  s  s  s  s  s  s  s  s  s  s  s  s  sc s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  sy s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  sc s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  sx s  s  s  s  '
    +'s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  s  t  t  t  t  td t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  ts t  t  t  t  t  t  t  tc t  t  t  t  t  t  t  t  t  t  t  t  ty t  t  t  '
    +'t  t  t  t  t  t  t  t  tdst  t  t  tb t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  ty t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  tg t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  t  tz t  '
    +'t  t  t  t  t  td t  t  t  t  t  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  wh w  wy w  wm w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  wy w  w  wy w  w  w  w  w  w  w  w  wy w  w  w  w  w  w  '
    +'w  w  w  w  w  w  w  w  w  wz w  wg w  w  w  w  w  w  w  w  w  w  w  w  w  wm w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  w  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  xq x  x  x  x  x  x  x  x  x  x  xj x  xh x  x  xh '
    +'x  x  x  x  x  x  x  x  xs xj xh x  x  x  x  x  xq x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  xh x  x  x  x  x  x  x  x  x  xa x  x  x  xy x  x  x  xj x  x  x  x  x  x  x  xh x  x  xj x  xy '
    +'x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  xs x  x  x  x  x  x  x  x  x  x  xh x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  xq x  x  xs x  x  x  x  x  x  x  x  xc x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  '
    +'x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  x  y  y  y  y  y  y  yd y  y  ya y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  '
    +'y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  yx yz y  y  y  y  y  y  y  y  y  y  y  y  y  yw y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  yg y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  '
    +'y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  yc y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  yt ys y  y  y  y  y  '
    +'y  y  y  y  y  y  y  y  y  y  y  y  y  y  yx y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  y  z  z  z  z  z  z  z  z  z  z  z  zc '
    +'z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  zc z  z  zc z  z  zygz  z  z  zs z  z  z  z  z  z  z  z  z  z  z  z  z  z  zn z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  '
    +'z  z  z  z  z  z  zs z  zs z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  zq z  z  z  z  z  z  z  z  z  z  z  z  zs z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  zs z  z  z  '
    +'z  z  z  z  z  z  z  z  z  z  z  z  z  zc z  zc z  z  z  z  z  z  z  zy z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  zy z  z  z  z  z  z  z  z  z  z  zc z  z  zc z  z  z  z  z  z  z  z  z  '
    +'z  z  z  z  z  z  z  z  z  zc z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  zj zc z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z  z                 c  j  w  g  n  s  p  g  c  lg n  e  s  y  p  b  t  y  y  z  '
    +'d  x  y  k  y  g  t  zd jq n  mn j  q  m  b  s  gj z  s  c  y  j  s  y  y  fx p  g  k  b  z  g  y  d  y  w  j  k  g  k  l  j  ys w  k  p  j  q  h  y  j  w  r  d  z  l  s  gy m  r  y  p  y  w  w  ct c  k  z  n  k  y  yd g  t  t  n  jgqj  ne y  k  k  z  '
    +'y  t  c  j  n  m  c  y  l  q  l  y  p  y  sq f  q  r  p  z  s  l  w  b  t  g  k  j  f  y  xj j  w  z  l  t  b  n  c  x  j  j  j  j  tz x  d  t  t  s  q  z  y  c  d  x  x  h  g  c  k  b  p  h  f  f  s  s  t  y  b  g  m  x  l  p  b  y  l  l  b  h  l  x  '
    +'s  m  z  m  y  j  h  s  o  j  n  g  h  d  z  q  y  k  l  g  j  h  s  g  q  z  h  x  q  g  k  ea z  z  w  y  s  c  s  c  j  x  y  e  y  x  a  d  z  p  m  d  s  s  m  z  j  jz q  j  y  z  c  j  j  fy w  q  j  b  dy z  pb x  g  z  n  z  c  p  w  h  wk x  '
    +'h  q  k  m  w  f  b  p  b  y  d  t  j  z  z  k  q  xh y  l  y  g  x  f  p  t  y  j  y  y  z  p  s  z  l  f  c  h  m  q  s  h  g  m  x  x  s  x  j  y  qk d  c  s  b  b  q  b  e  f  s  j  y  h  wx w  g  z  k  p  y  l  qy b  g  l  d  l  c  cd t  n  m  a  '
    +'y  d  td k  s  s  n  g  y  c  s  g  x  l  hy z  a  y  pb n  pb t  s  d  kj d  y  l  h  g  y  m  y  l  c  x  p  y  c  j  n  d  q  j  w  x  q  x  f  y  y  f  j  l  e  jq bp z  r  x  c  c  q  w  q  q  s  b  zn k  y  m  g  p  l  b  m  j  r  q  c  f  l  n  '
    +'y  m  y  q  m  ts qx ty r  b  zc j  t  h  z  t  q  f  r  x  q  h  x  m  jq j  c  j  l  yx qx g  j  m  s  h  z  k  b  s  wz y  e  m  y  l  t  x  fp s  y  d  xs gw l  y  c  j  jq x  s  j  n  q  b  s  c  t  y  h  b  f  t  d  c  y  z  d  j  yw y  g  h  q  '
    +'f  sr x  w  c  k  q  k  x  e  b  p  t  l  p  x  j  z  s  r  m  e  b  w  h  j  l  bp j  s  l  y  y  s  m  d  x  l  c  l  q  k  xs l  h  xj j  r  z  j  m  fb q  h  x  h  w  y  w  s  b  h  t  r  x  xl g  l  h  q  h  f  n  m  g  y  k  l  d  y  x  z  p  wy '
    +'l  g  g  s  m  t  cs f  p  anyj  j  z  y  l  j  t  y  a  n  j  g  b  j  p  l  q  g  sd z  y  q  y  a  x  b  k  y  s  e  c  j  s  z  n  s  l  y  z  h  z  x  l  z  c  g  h  p  x  z  h  z  n  y  t  d  s  b  c  j  k  td l  z  a  y  f  m  y  dt l  e  bp b  '
    +'g  q  y  z  k  x  gw l  d  n  d  n  y  s  k  j  s  h  d  l  y  x  b  cz g  hy x  y  p  k  d  j  m  m  z  n  g  m  m  c  l  g  wn z  s  z  x  z  j  f  z  rn m  l  z  z  t  h  c  s  y  d  b  d  l  l  s  c  dz dt n  l  k  j  y  k  j  s  y  c  j  l  k  wo '
    +'h  q  a  s  d  k  n  h  c  s  g  as n  h  d  a  ay s  h  t  c  p  l  c  p  q  y  b  sz d  m  p  j  l  p  zc j  o  q  l  c  d  h  j  xj y  s  p  r  c  h  n  w  jn n  l  h  l  y  y  q  y  yh w  z  p  t  cd z  g  w  w  m  z  f  f  j  q  q  q  q  y  x  a  '
    +'c  l  b  h  k  d  j  x  d  g  m  m  y  td jq x  z  l  l  s  y  g  yx g  k  j  r  y  w  z  w  y  c  l  z  m  s  s  j  z  l  d  b  y  d  c  pf c  x  y  h  l  x  c  h  y  z  j  q  s  q  q  ah g  m  n  y  x  p  f  r  k  s  s  b  j  l  y  x  y  s  y  g  l  '
    +'n  s  c  m  h  c  w  w  m  n  z  j  j  l  x  x  h  c  h  s  y  z  s  xt t  x  r  y  c  y  x  b  y  h  c  s  m  x  j  s  z  n  p  w  g  p  x  x  t  a  y  b  g  a  qj c  x  l  y  x  d  c  c  w  z  o  c  w  k  c  c  zs b  n  h  c  p  d  y  z  n  f  c  y  '
    +'y  t  y  c  k  x  kq y  b  s  q  kl k  y  t  q  q  x  f  c  w  c  h  c  y  k  e  l  z  q  b  s  q  y  j  q  c  c  l  m  t  h  s  y  w  h  m  k  t  l  k  j  l  y  c  x  w  h  ye q  q  h  t  q  kh qp pqbq  s  q  s  c  f  y  m  m  dz m  g  b  w  h  w  l  '
    +'g  s  ls l  y  s  d  l  m  lp x  p  t  h  m  j  h  w  l  j  z  y  h  z  j  x  hk t  x  j  l  hx x  r  s  w  l  w  z  j  c  b  x  m  hw z  q  x  s  d  z  p  ms g  f  c  s  g  l  s  x  y  m  qj s  h  x  p  j  x  w  m  y  zq k  s  m  y  p  l  r  t  h  b  '
    +'x  f  t  p  m  h  y  x  l  c  h  lt h  l  z  y  l  x  g  s  s  s  s  t  c  l  s  l  dt c  l  r  p  b  h  z  h  x  y  y  f  h  b  m  g  d  m  y  c  n  q  q  w  l  q  h  j  j  c  y  w  j  z  y  e  j  j  d  h  p  b  l  q  x  t  q  k  w  h  l  c  h  q  x  '
    +'a  g  t  l  x  l  j  x  m  s  l  j  h  t  z  k  z  j  e  c  x  j  c  j  n  m  f  b  y  c  s  f  y  w  y  b  j  z  g  n  y  s  d  z  s  q  y  r  s  l  j  p  c  l  p  w  x  s  d  w  e  j  b  j  c  b  c  n  a  y  t  w  g  m  p  a  p  c  l  y  q  p  c  l  '
    +'z  x  s  b  n  m  s  g  g  f  n  z  j  j  b  z  s  f  z  y  n  td x  h  p  l  q  k  z  c  z  w  a  l  s  pb c  zc j  x  m  y  z  hg w  k  y  p  s  g  x  f  z  f  c  d  k  h  j  g  x  td l  q  f  s  g  d  s  l  q  w  z  k  x  t  m  h  s  bp g  z  m  j  '
    +'z  r  g  l  y  j  b  p  m  l  m  s  x  l  z  j  q  qszh  z  y  j  c  z  y  d  j  w  bf mw j  k  l  d  d  p  m  j  e  g  x  y  h  y  l  x  h  l  q  y  q  h  k  y  c  w  c  j  m  y  y  x  n  a  t  j  h  y  c  c  x  z  p  c  q  l  b  z  w  w  y  t  w  sb '
    +'q  c  m  l  p  m  ym r  j  c  c  zc x  f  pb z  n  z  z  l  j  p  l  x  x  y  z  t  z  l  gj d  ly td c  k  l  y  r  z  z  g  q  t  gk j  h  h  gh j  l  j  a  x  f  g  f  j  z  s  l  c  f  d  q  z  l  c  l  g  j  d  j  c  sz n  zc l  l  j  p  j  q  d  '
    +'c  c  l  c  j  x  m  y  z  f  t  s  x  g  c  g  s  b  r  z  x  j  q  qc cs t  z  h  g  y  q  t  j  q  q  l  z  x  j  y  l  y  l  b  c  y  a  m  c  s  t  y  l  p  d  j  b  y  r  e  g  k  l  z  y  z  h  l  y  s  z  q  l  z  n  w  c  z  c  l  l  w  j  q  '
    +'j  j  j  k  d  g  j  z  o  l  b  b  z  p  p  g  l  g  h  t  g  z  x  y  gj h  z  m  y  c  n  q  sc y  cz y  h  b  h  g  x  k  a  m  t  x  y  x  n  b  s  k  y  z  z  g  j  z  l  q  j  d  f  c  j  x  d  y  g  j  q  j  j  p  m  g  w  g  j  j  jq p  kd q  '
    +'s  b  g  b  m  m  c  j  s  s  c  l  p  q  p  d  x  c  d  y  y  k  y  p  c  j  d  d  y  y  g  y  w  r  h  j  r  t  g  z  n  y  q  l  d  k  l  j  s  z  z  g  z  q  z  j  g  d  y  k  s  hg p  z  m  t  l  c  p  w  n  j  ay f  y  z  d  j  c  n  m  w  e  s  '
    +'c  y  g  l  b  t  z  c  g  m  s  s  l  l  y  x  yq s  x  sx b  s  j  s  b  b  s  g  g  h  f  j  l  wy p  m  z  j  n  l  y  y  w  d  q  s  h  z  x  t  y  y  w  h  m  zc y  h  y  w  d  b  x  b  t  l  m  s  y  y  y  f  s  x  j  c  b  dt x  x  l  h  j  h  '
    +'f  s  s  x  z  q  h  f  z  mn z  c  z  t  q  sc x  z  x  r  t  t  d  j  h  n  nr y  z  q  q  m  nt q  d  m  m  gz y  y  td x  m  j  g  d  hx c  d  y  z  b  f  f  a  l  l  z  t  d  l  t  f  x  m  x  q  z  d  n  lg w  q  d  b  d  c  z  j  d  x  b  z  g  '
    +'s  q  q  d  d  j  c  m  b  k  z  f  f  x  m  k  d  m  d  s  y  y  s  z  c  m  l  j  d  s  ya n  s  pb r  s  k  m  k  m  p  c  k  l  g  dt b  q  t  f  z  s  w  t  f  g  g  l  y  p  l  l  j  z  h  g  j  j  g  y  p  z  l  t  c  s  m  c  n  b  t  j  b  q  '
    +'f  k  t  h  bp y  z  g  k  p  b  b  y  m  t  dt s  s  x  ts b  n  p  d  k  l  e  y  c  j  n  y  dc d  y  k  z  td d  h  q  h  s  ydtz  s  c  t  a  r  l  l  t  k  z  l  g  e  c  l  l  k  j  l  q  j  a  q  n  b  d  k  k  g  h  p  j  t  z  q  k  s  e  c  '
    +'s  h  a  l  q  f  m  m  g  j  n  l  y  j  b  b  t  m  l  y  z  x  d  cxtj  p  l  d  l  p  c  q  d  h  z  y  c  b  z  s  c  z  b  z  m  s  l  j  f  l  k  r  z  j  s  n  f  r  g  j  h  x  p  d  h  y  j  y  b  z  g  d  l  q  c  s  e  z  g  x  l  b  l  hg '
    +'y  x  t  w  m  a  b  c  h  e  c  m  w  y  j  y  z  l  l  j  j  y  h  l  g  c  d  j  l  s  l  y  g  k  d  z  p  z  x  j  y  y  z  l  w  c  x  s  z  f  g  w  y  y  d  l  y  h  c  l  jx s  c  m  b  j  h  b  l  y  z  l  y  c  b  l  y  d  p  d  q  y  s  x  '
    +'q  z  b  y  t  d  k  y  yx j  y  y  c  n  r  j  m  p  d  jq g  k  l  c  l  j  b  cz tx bp j  d  d  b  b  l  b  l  c  zx q  r  p  psyx  j  c  jqgl  z  c  s  h  l  t  o  l  j  n  m  d  d  d  l  n  g  k  a  tq h  q  h  jx hg y  k  h  e  z  n  m  s  h  r  '
    +'p  h  q  q  j  c  h  g  m  f  p  r  x  h  j  g  d  y  c  h  g  hk l  y  r  z  q  l  c  y  q  j  n  z  s  q  t  k  q  j  y  m  s  z  sx w  l  c  f  q  q  q  x  y  f  g  g  y  p  t  q  w  l  m  c  r  n  f  k  k  f  s  y  y  l  yq b  m  q  a  m  m  m  y  '
    +'x  c  t  p  sz h  c  p  t  x  x  z  z  s  m  p  h  p  s  h  m  c  l  m  l  d  q  f  y  q  x  s  z  y  j  d  j  j  z  z  h  q  p  d  s  z  g  l  s  t  j  b  c  k  b  x  y  q  z  yj s  g  p  s  x  q  z  q  z  r  q  t  b  d  k  y  x  z  k  h  h  g  f  l  '
    +'b  c  s  m  d  l  d  g  d  z  d  b  l  z  y  y  c  x  n  n  c  s  y  b  z  b  f  g  l  z  z  x  s  w  m  s  c  ct m  q  n  j  q  s  b  d  q  s  j  t  x  x  m  b  l  t  x  zc c  l  z  s  h  z  c  x  r  q  j  g  j  y  l  x  z  f  j  p  h  y  m  z  q  q  '
    +'zy d  f  q  j  jq l  z  z  n  z  j  c  d  g  z  y  g  zc t  x  m  z  y  s  c  t  l  k  p  h  t  x  h  t  l  b  j  x  j  l  x  s  c  d  q  x  c  b  b  t  j  f  q  z  f  s  l  t  j  b  t  k  q  b  x  x  j  j  l  j  c  h  c  z  d  b  z  j  d  c  z  j  d  '
    +'c  p  r  n  p  xq c  j  p  f  c  z  l  c  l  z  x  z  d  m  x  m  p  h  j  s  g  z  g  s  z  z  q  jl y  l  w  t  j  p  f  s  y  a  xs m  c  j  b  t  z  yk y  c  w  m  y  t  zc sz j  j  l  jq c  q  l  w  z  m  a  l  b  x  y  f  b  p  n  l  s  f  h  t  '
    +'gx j  w  e  j  j  x  x  g  l  l  j  s  t  g  s  h  j  qz l  z  f  k  c  g  n  n  d  s  z  f  d  e  q  f  h  b  s  a  q  t  g  y  l  b  x  m  m  y  g  s  z  l  d  y  d  q  mw j  j  r  g  b  jx ts k  gh d  h  g  k  b  l  q  k  b  d  m  b  y  l  x  w  c  '
    +'x  y  t  t  y  bs k  m  r  t  j  z  x  q  j  b  h  l  m  h  m  j  jq z  m  q  a  s  l  d  c  y  x  y  q  d  l  qj c  a  f  y  w  y  x  yq h  z  ';
{$ENDIF}

//判断AStr的第AIndex个字符是否为一个合法汉字
function IsChineseWordChar(const AStr: string; const AIndex: Integer): Boolean;
var
{$IFDEF UNICODE}
  iCode: Integer;
{$ELSE}
  SecCode: Integer; //汉字区码
  PosCode: Integer; //汉字位码
{$ENDIF}
begin
  {$IFDEF UNICODE}
  iCode := Ord(AStr[AIndex]);
  Result := (iCode >= $4E00) and (iCode <= $9FA5);  //$4E00至$9FA5(即:19968至40869, 共20902个汉字)
  {$ELSE}
  Result := AStr[AIndex] in LeadBytes;
  if Result then
  begin
    SecCode := Ord(AStr[AIndex]);
    PosCode := Ord(AStr[AIndex + 1]);                                                   //下面{ and (PosCode <> 127)}出于效率考虑，跳过之。
    Result := ((SecCode >= 129) and (SecCode <= 160) and (PosCode >= 64) and (PosCode <= 254){ and (PosCode <> 127)}) //第一部分(区:129-160 位:64-254[除127])
           or ((SecCode >= 170) and (SecCode <= 253) and (PosCode >= 64) and (PosCode <= 160){ and (PosCode <> 127)}) //第二部分(区:170-253 位:64-160[除127], 区253位156-160留空)
           //     and not ((SecCode = 253) and (PosCode >= 156) and (PosCode <= 160)))  //由于这部分发生的概率极低，出于效率考虑，跳过之。
           or ((SecCode >= 176) and (SecCode <= 247) and (PosCode >= 161) and (PosCode <= 254));                    //第三部分(区:176-247 位:161-254, 区215位250-254留空)
           //     and not ((SecCode = 215) and (PosCode >= 250) and (PosCode <= 254))); //由于这部分发生的概率极低，出于效率考虑，跳过之。
  end;
  {$ENDIF}
end;

//获取指定汉字的拼音索引字母，如：“汉”的索引字母是“H”，多音字返回多个声母，如：“行”返回“hx”。
//参数OnlyOne表示多音字是否仅返回第一个声母。缺省False。
function GetPYIndex(const OneWord: String; const OnlyOne: Boolean = False): String;
{$IFDEF UNICODE}
  function GetPart(const AOffset: Integer): String;
  var
    i: Integer;
  begin
    Result := PartUnicode[AOffset];
    if not OnlyOne then
    begin
      i := 1;
      while (i < MaxPYLength) and (PartUnicode[AOffset + i] <> ' ') do
      begin
        Result := Result + PartUnicode[AOffset + i];
        Inc(i);
      end;
    end;
  end;

{$ELSE}
  function GetPart1(const AOffset: Integer): string;
  var
    i: Integer;
  begin
    Result := Part1[AOffset];
    if not OnlyOne then
    begin
      i := 1;
      while (i < MaxPYLength) and (Part1[AOffset + i] <> ' ') do
      begin
        Result := Result + Part1[AOffset + i];
        Inc(i);
      end;
    end;
  end;

  function GetPart2(const AOffset: Integer): string;
  var
    i: Integer;
  begin
    Result := Part2[AOffset];
    if not OnlyOne then
    begin
      i := 1;
      while (i < MaxPYLength) and (Part2[AOffset + i] <> ' ') do
      begin
        Result := Result + Part2[AOffset + i];
        Inc(i);
      end;
    end;
  end;

  function GetPart3(const AOffset: Integer): string;
  var
    i: Integer;
  begin
    Result := Part3[AOffset];
    if not OnlyOne then
    begin
      i := 1;
      while (i < MaxPYLength) and (Part3[AOffset + i] <> ' ') do
      begin
        Result := Result + Part3[AOffset + i];
        Inc(i);
      end;
    end;
  end;
{$ENDIF}
var
{$IFDEF UNICODE}
  iCode, iIndex: Integer;
{$ELSE}
  SecCode: Integer; //汉字区码
  PosCode: Integer; //汉字位码
  Offset: Integer; //字库偏移量
{$ENDIF}
begin
  if OneWord = '' then
  begin
    Result := '';
    Exit;
  end;
  {$IFDEF UNICODE}
    iCode := Ord(OneWord[1]);
    if (iCode >= 19968) and (iCode <= 40869) then
      Result := GetPart((iCode - 19968) * MaxPYLength + 1)
    else
      Result := OneWord;
  {$ELSE}
    SecCode := Ord(OneWord[1]);
    if SecCode < 129 then    //非汉字
    begin
      Result := OneWord;
      Exit;
    end;
    PosCode := Ord(OneWord[2]);
    //第一段汉字 (区:129-160 位:64-254[除127]) 	32*190=6080
    if (SecCode >= 129) and (SecCode <= 160)
      and (PosCode >= 64) and (PosCode <= 254) then
    begin
      if PosCode > 127 then Dec(PosCode, 1);
      Offset := ((SecCode - 129) * 190 + PosCode - 64) * MaxPYLength + 1;
      Result := GetPart1(Offset);
    end
    else
    //第二段汉字 (区:170-253 位:64-160[除127])  	84*96 =8064
    if (SecCode >= 170) and (SecCode <= 253)
      and (PosCode >= 64) and (PosCode <= 160) then
    begin
      if PosCode > 127 then Dec(PosCode, 1);
      Offset := ((SecCode - 170) * 96 + PosCode - 64) * MaxPYLength + 1;
      Result := GetPart2(Offset);
    end
    else
    //第三段汉字 (区:176-247 位:161-254)         72*94 =6768
    if (SecCode >= 176) and (SecCode <= 247)
      and (PosCode >= 161) and (PosCode <= 254) then
    begin
      Offset := ((SecCode - 176) * 94 + PosCode - 161) * MaxPYLength + 1;
      Result := GetPart3(Offset);
    end
    //其他
    else
      Result := OneWord;
  {$ENDIF}
end;

//获取指定字符串的拼音索引字母（多音字总是取第一个音的声母)
function GetStrPYIndex(const Str: String): String;
var
  i: Integer;
  Len: Integer;
begin
  Result := '';
  Len := Length(Str);
  i := 1;
  while i <= Len do
  begin
    if Ord(Str[i]) < 128 then //非汉字
      Result := Result + Str[i]
    else
    begin
      {$IFDEF UNICODE}
      Result := Result + GetPYIndex(Str[i], True);
      {$ELSE}
      Result := Result + GetPYIndex(Str[i] + Str[i + 1], True);
      Inc(i);
      {$ENDIF}
    end;
    Inc(i);
  end;
end;

//把s中的非双字节转为小写
procedure LowerText(var s: String);
var
  i: Integer;
begin
  i := 1;
  while i <= Length(s) do
  begin
    {$IFNDEF UNICODE}
    if s[i] in LeadBytes then
    begin
      Inc(i, 2);
      Continue;
    end;
    {$ENDIF}
    if (Ord(s[i]) >= Ord('A')) and (Ord(s[i]) <= Ord('Z')) then
    begin
      {$WARN UNSAFE_CODE OFF}
      s[i] := Char(Ord(s[i]) + 32);
      {$WARN UNSAFE_CODE ON}
    end;
    Inc(i);
  end;
end;

//获取字符串Str中与拼音声母字符串PYStr匹配的起始位置 (内部函数)
//参数: IsSame如果Str的全部拼音声母等于PYStr, 则输出True.
function InternalPosPY(const PYStr, Str: String; var IsSame: Boolean): Integer;
var
  i, j, cMatch: Integer;
  sPY, LowerPYStr: String;
  PYL, k: Integer;
  PYMatch: Boolean;
begin
  if PYStr = Str then
  begin
    Result := 1;
    IsSame := True;
    Exit;
  end;
  IsSame := False;
  Result := Pos(PYStr, Str);
  if Result > 0 then Exit;

  LowerPYStr := PYStr;
  LowerText(LowerPYStr);
  cMatch := 0;
  i := 0;
  while i < Length(Str) do
  begin
    for j := 1 to Length(PYStr) do
    begin
      if Str[i + cMatch + j] = PYStr[j] then Continue;
      if IsChineseWordChar(Str, i + cMatch + j)
        and (Ord(LowerPYStr[j]) >= Ord('a'))
        and (Ord(LowerPYStr[j]) <= Ord('z')) then
      begin
        PYMatch := False;
        sPY := GetPYIndex(Str[i + cMatch + j]{$IFNDEF UNICODE} + Str[i + cMatch + j + 1]{$ENDIF});
        PYL := Length(sPY);
        k := 1;
        while k <= PYL do
        begin
          if sPY[k] = LowerPYStr[j] then
          begin
            PYMatch := True;
            Break;
          end;
          Inc(k);
        end;
        if PYMatch then
        begin
          {$IFNDEF UNICODE}
          Inc(cMatch);
          {$ENDIF}
          Continue;
        end;
      end;
      {$IFNDEF UNICODE}
      if Str[i + cMatch + j] in LeadBytes then
        Inc(i);
      {$ENDIF}
      cMatch := 0; //cyw 12.8.24
      Break;
    end; //for j
    if j > Length(PYStr) then
    begin
      Result := i + 1;
      IsSame := (Result = 1) and (i + cMatch + j - 1 = Length(Str));
      Exit;
    end;
    i := i + cMatch + 1;
    cMatch := 0;
  end; //for i
end;

//获取字符串Str中与拼音声母字符串PYStr匹配的起始位置
function PosPY(const PYStr, Str: String): Integer;
var
  B: Boolean;
begin
  Result := InternalPosPY(PYStr, Str, B);
end;

//判断字符串Str的拼音声母是否全部与PYStr相同
function SamePY(const PYStr, Str: String): Boolean;
begin
  InternalPosPY(PYStr, Str, Result);
end;

//获取字符串Str中拼音声母串与字符串SubStr拼音声母串匹配的起始位置 (内部函数)
//参数: IsSame如果SubStr的全部拼音声母等于Str的全部拼音声母, 则输出True.
function InternalPosTextPY(const SubStr, Str: String; var IsSame: Boolean): Integer;
var
  i, j, cMatch: Integer;
  sPY, bPY: String;
  PYL, bPYL, k, m: Integer;
  PYMatch: Boolean;
begin
  if SubStr = Str then
  begin
    Result := 1;
    IsSame := True;
    Exit;
  end;
  IsSame := False;
  Result := Pos(SubStr, Str);
  if Result > 0 then Exit;

  cMatch := 0;
  i := 0;
  while i < Length(Str) do
  begin
    j := 1;
    while j <= Length(SubStr) do
    begin
      if Str[i + cMatch + j] = SubStr[j] then
      begin
        if Ord(SubStr[j]) < 128 then
        begin
          Inc(j);
          Continue;
        end;
        if Str[i + cMatch + j + 1] = SubStr[j + 1] then
        begin
          Inc(j, 2);
          Continue;
        end;
      end;
      if IsChineseWordChar(Str, i + cMatch + j) then
      begin
        PYMatch := False;
        sPY := GetPYIndex(Str[i + cMatch + j]{$IFNDEF UNICODE} + Str[i + cMatch + j + 1]{$ENDIF});
        if IsChineseWordChar(SubStr, j) then
          bPY := GetPYIndex(SubStr[j]{$IFNDEF UNICODE} + SubStr[j + 1]{$ENDIF})
        else
          bPY := LowerCase(SubStr[j]);
        PYL := Length(sPY);
        bPYL := Length(bPY);
        k := 1;
        while k <= PYL do
        begin
          m := 1;
          while m <= bPYL do
          begin
            if sPY[k] = bPY[m] then
            begin
              PYMatch := True;
              Break;
            end;
            Inc(m);
          end;
          if PYMatch then Break;
          Inc(k);
        end;
        if PYMatch then
        begin
          {$IFNDEF UNICODE}
          if SubStr[j] in LeadBytes then
            Inc(j, 2)
          else
          {$ENDIF}
          begin
            Inc(j);
            {$IFNDEF UNICODE}
            Inc(cMatch);
            {$ENDIF}
          end;
          Continue;
        end;
      end;
      {$IFNDEF UNICODE}
      if Str[i + cMatch + j] in LeadBytes then
        Inc(i);
      {$ENDIF}
      cMatch := 0;  //cyw 12.8.24
      Break;
    end; //while j
    if j > Length(SubStr) then
    begin
      Result := i + 1;
      IsSame := (Result = 1) and (i + cMatch + j - 1 = Length(Str));
      Exit;
    end;
    i := i + cMatch + 1;
    cMatch := 0;
  end; //for i
end;

//获取字符串Str中拼音声母串与字符串SubStr拼音声母串匹配的起始位置
function PosTextPY(const SubStr, Str: String): Integer;
var
  B: Boolean;
begin
  Result := InternalPosTextPY(SubStr, Str, B);
end;

//判断字符串Str的拼音声母是否全部与字符串SubStr的拼音声母相同
function SameTextPY(const SubStr, Str: String): Boolean;
begin
  InternalPosTextPY(SubStr, Str, Result);
end;

end.

