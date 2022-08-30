unit tables;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,dialogs;

type TimeAndDate = class
private
  Minute:byte;
  Hour  :byte;
  Day   :byte;
  Month :byte;
  Year  :integer;
public
  function  GetTime:string;
  function  GetDate:string;
  procedure ReadTime(text:string);
  procedure ReadDate(text:string);
end;

Readings = class
private
  Temp:real;
public
  Time:TimeAndDate;
  procedure ReadTemp(text:string);
  function GetTempStr:string;
  function GetTempReal:real;
end;

tCoord = record
     stolb,strok:integer;
end;

type tablesInfo = class
  private
    tableAll:boolean;
    tableCrit:boolean;
    chart:boolean;
    count:byte;
  public
    constructor create;
    procedure ReadInfo(s:string);
    function  GetTableAll:boolean;
    function  GetTableCrit:boolean;
    function  GetChart:boolean;
    function  Getcount:byte;
end;

Table = class
private
  StartCoord:tCoord;
  EndCoord:tCoord;
  function CheckEmpty(var e:olevariant; p:tCoord):boolean;
  function tCoordToRange(var e:olevariant; p1,p2:tCoord):string;
  procedure GetCoord(var e:olevariant; CritTable:boolean);
  function findLeftUp(var e:olevariant; p:tCoord):tCoord;
  function findRightDOwn(var e:olevariant; p:tCoord):tCoord;
public
  constructor create;
  procedure CopyTable(var e:olevariant; ListNom:integer; CritTable:boolean);
  procedure PasteTable(var x,e:olevariant; CritTable:boolean; Nomer:byte);
end;

Diag = class
private

public
  procedure CopyDiag(var e:olevariant; ListNom:integer; CritTable:boolean);
  procedure PasteDiag(var x,e:olevariant; CritTable:boolean; Nomer:byte);
end;

Test = class

private
  nom:byte;
  ListName:string;
  procedure GetListName;

public
  Chart:Diag;
  CritTable:Table;
  BaseTable:Table;
  procedure ReadNom(k:integer);
  procedure DoSomeMagic;
end;

implementation

{TimeAndDate}

function  TimeAndDate.GetDate:string;
begin
  GetDate:=IntToStr(Day)+'.'+IntToStr(Month)+'.'+IntToStr(Year);
end;
function  TimeAndDate.GetTime:string;
begin
  GetTime:=IntToStr(Hour)+':'+IntToStr(Minute);
end;
procedure TimeAndDate.ReadTime(text:string);
begin

end;
procedure TimeAndDate.ReadDate(text:string);
begin

end;

{Readings}

procedure Readings.ReadTemp(text:string);
begin

end;
function Readings.GetTempStr:string;
begin

end;
function Readings.GetTempReal:real;
begin

end;

{tablesInfo}

constructor tablesInfo.create;
begin
  inherited;
end;

function GetNumberFromString(s:string; var i:integer; &type:string):integer;
var buf:string;
begin
  buf:='';
  i:= pos(&type,s);
  if i=0 then
    GetNumberFromString:=0
  else
  begin
    while (((s[i]<='0') or (s[i]>='9')) and (i<=length(s))) do
      i:=i+1;
    while (not((s[i]<='0') or (s[i]>='9')) and (i<=length(s))) do
    begin
      buf:=buf+s[i];
      i:=i+1;
    end;
    if buf<>'' then
      GetNumberFromString:=StrToInt(buf)
    else
      showmessage('Ошибка считывания данных ' + &type + ' из тега: (' + s + ')');
  end;
end;

procedure tablesInfo.ReadInfo(s:string);
var i:integer;
begin
  //c_Isp_All=1_Crit=1_Chart=1_Count=7
  if GetNumberFromString(s,i,'All') = 1 then
  begin
    showmessage(s + ': tableAll');
    tableAll:=True
  end
  else
    //tableAll:=False;
    showmessage(s + ': NOTtableAll');
  if GetNumberFromString(s,i,'Crit') = 1 then
    begin
    showmessage(s + ': tableCrit');
    tableCrit:=True
  end
  else
    tableCrit:=False;
  if GetNumberFromString(s,i,'Chart') = 1 then
    begin
    showmessage(s + ': chart');
    chart :=True
  end
  else
    chart:=False;
  count:=GetNumberFromString(s,i,'Count');
  if count=0 then
    showmessage('Ошибка считывания данных из тега: (' + s + ')')
  else
    showmessage(IntToStr(count));
end;

function   tablesInfo.GetTableAll:boolean; 
begin
  GetTableAll:=tableAll;
end;

function   tablesInfo.GetTableCrit:boolean;
begin
  GetTableCrit:=tableCrit;
end;

function   tablesInfo.GetChart:boolean;
begin
  GetChart:=chart;
end;

function   tablesInfo.Getcount:byte;
begin
  Getcount:=count;
end;


{Table}

constructor Table.create;
begin
  inherited
 // tabl.create;
end;

function checkstolbAndstrok(var e:olevariant; p:tCoord):boolean;
begin
  checkstolbAndstrok:=True;
  if p.strok>e.ActiveWorkbook.ActiveSheet.rows.count then
  begin
    showmessage('Выход за пределы строк');
    checkstolbAndstrok:=False;
  end;
  if p.stolb>e.ActiveWorkbook.ActiveSheet.Columns.count then
  begin
    showmessage('Выход за пределы столбцов');
    checkstolbAndstrok:=False;
  end;
  if p.stolb<0 then
  begin
    showmessage('Номер столбца меньше нуля');
    checkstolbAndstrok:=False;
  end;
  if p.strok<0 then
  begin
    showmessage('Номер строки меньше нуля');
    checkstolbAndstrok:=False;
  end
end;

function Table.CheckEmpty(var e:olevariant; p:tCoord):boolean;
var i:integer;
begin
  //showmessage(E.ActiveSheet.Cells[p.strok, p.stolb].text);
  if checkstolbAndstrok(e,p) then
  if E.ActiveSheet.Cells[p.strok, p.stolb].text='' then
    CheckEmpty:=True
  else
    CheckEmpty:=False
end;
function Table.tCoordToRange(var e:olevariant; p1,p2:tCoord):string;
var buf:integer;
begin
  buf:=ord('A')-1;
  if p1.stolb>26 then
    tCoordToRange:='A'+chr(buf+p1.stolb-26)
  else
    tCoordToRange:=chr(buf+p1.stolb);
 // tCoordToRange:=IntToStr(p1.stolb)+',';
  tCoordToRange+=IntToStr(p1.strok);
  tCoordToRange+=':';
  if p2.stolb>26 then
    tCoordToRange+='A'+chr(buf+p2.stolb-26)
  else
    tCoordToRange+=chr(buf+p2.stolb);
 //tCoordToRange+=IntToStr(p1.stolb)+',';
  tCoordToRange+=IntToStr(p2.strok);
end;

function Table.findLeftUp(var e:olevariant; p:tCoord):tCoord;
var i:integer;
begin
  while not(checkEmpty(e,p)) do
  begin
    p.stolb:=p.stolb-1;
  end;
  p.stolb:=p.stolb+1;
  while not(checkEmpty(e,p)) do
  begin
    p.strok:=p.strok-1;
  end;
  p.strok:=p.strok+1;
  findLeftUp:=p;
   //showmessage(tCoordToRange(e,p,p));
end;
function Table.findRightDOwn(var e:olevariant; p:tCoord):tCoord;
var i:integer;
begin
  while not(checkEmpty(e,p)) do
  begin
    p.stolb:=p.stolb+1;
   // showmessage(tCoordToRange(e,p,p));
  end;
  p.stolb:=p.stolb-1;
  while not(checkEmpty(e,p)) do
  begin
    p.strok:=p.strok+1;
   // showmessage(tCoordToRange(e,p,p));
  end;
  p.strok:=p.strok-1;
  findRightDOwn:=p;
  //showmessage(tCoordToRange(e,p,p));
end;
procedure Table.GetCoord(var e:olevariant; CritTable:boolean);
var i,j:integer;
    p:tCoord;
begin
  if not(CritTable) then
  begin
    p.strok:=1;
    p.stolb:=1;
   // showmessage(tCoordToRange(e,p,p));
    while checkEmpty(e,p) do
    begin
      p.strok+=1;
      p.stolb+=1;
      //showmessage(tCoordToRange(e,p,p));
    end;
   // showmessage(tCoordToRange(e,p,p));
    StartCoord:=findLeftUp(e,p);
    EndCoord:=findRightDown(e,p);
   // showmessage(tCoordToRange(e,StartCoord,EndCoord));
    //showmessage(tCoordToRange(e,StartCoord,EndCoord));
  end
  else
  begin
      GetCoord(e,False);
      p.strok:=StartCoord.strok;
      p.stolb:=EndCoord.stolb+1;
     // showmessage(tCoordToRange(e,StartCoord,EndCoord));
      //showmessage(tCoordToRange(e,p,p));
      j:=0;
      while (checkEmpty(e,p)  and (j<20))  do
      begin
        j+=1;
        i:=1;
        p.strok:=StartCoord.strok;
        p.stolb:=EndCoord.stolb+j;
        while (checkEmpty(e,p) and (i<32+EndCoord.strok))  do
        begin
          p.strok+=i*2;
          i+=1;
        end;
      end;
      if j>70 then
        showmessage('Не обнаружена критическая таблица. Рекомендуется закрыть программу во избежания ошибок')
      else
      begin
        StartCoord:=findLeftUp(e,p);
        EndCoord:=findRightDown(e,p);
        StartCoord.strok-=4;
      end;
      //запрашиваем координаты конца основной таблицы
      //поднимаемся до начала основной таблицы
      //отспупаем на одну клетку вправо
      //проверя на пустоту, проходим до низа таблицы, делая шаг через три ячейки, до максимального значения = край таблицы + 32 ячеек
      //отступаем на единицу вправо, и в цикле продолжаем проверять ячейки  максимально до +30, с шагом в одну ячейку направо
      //когда нашли непустую ячейку, поднимаемся до левого верхнего края
      //находим правую нижнюю ячейку
  end;
end;
procedure Table.CopyTable(var e:olevariant; ListNom:integer; CritTable:boolean);
begin
  e.ActiveWorkbook.Sheets.item[3+ListNom].select;
  getcoord(e,CritTable);
  //showmessage(tCoordToRange(e,StartCoord,EndCoord));
 // showmessage(E.ActiveSheet.Cells[StartCoord.strok, StartCoord.stolb].text);
  //showmessage(E.ActiveSheet.Cells[EndCoord.strok, EndCoord.stolb].text);

  e.ActiveWorkbook.ActiveSheet.Range[widestring(tCoordToRange(e,StartCoord,EndCoord))].copy;
end;

procedure Table.PasteTable(var x,e:olevariant; CritTable:boolean; Nomer:byte);
begin

end;



{Diag}

procedure Diag.CopyDiag(var e:olevariant; ListNom:integer; CritTable:boolean);
begin
  e.ActiveWorkbook.Sheets.item[3+ListNom].select;
  showmessage('');
  e.ActiveWorkbook.ActiveSheet.ChartObjects(widestring('Диаграмма 1')).activate;
  e.ActiveWorkbook.ActiveChart.ChartArea.Select;
  e.ActiveWorkbook.ActiveChart.ChartArea.Copy;
  e.ActiveWorkbook.ActiveSheet.Range('AA28').select;
  e.ActiveWorkbook.ActiveSheet.Paste;
end;
procedure Diag.PasteDiag(var x,e:olevariant; CritTable:boolean; Nomer:byte);
begin

end;

{Test}

procedure Test.GetListName;
begin

end;
procedure Test.ReadNom(k:integer);
begin

end;
procedure Test.DoSomeMagic;
begin

end;
end.

