REM Author:Jason
REM DATE:2019-2-27
REM Version 2.2.1

Set objShell = CreateObject("Wscript.Shell")
Set WshShell=WScript.CreateObject("WScript.Shell")
set fso =createobject("scripting.filesystemobject")
path_logs="E:\MySQL_BACK\logs\"         '日志路径
exec_dump="D:\MySQLDump\mysqldump.exe"  '应用程序路径
exec_del_time=7                         '日志保留天数

rem 备份名，目录，IP，端口，账号，密码，数据库名
dim array(2,6)

rem 192.168.30.2_DDZY
array(0,0)="192.168.30.2_DDZY"
array(0,1)="E:\MySQL_BACK\192.168.30.2_DDZY\"
array(0,2)="192.168.15.2"
array(0,3)="3306"
array(0,4)="back"
array(0,5)="abcd1234"
array(0,6)="ddzy"

rem 192.168.20.2_BOSS
array(1,0)="192.168.20.2_BOSS"
array(1,1)="E:\MySQL_BACK\192.168.20.2_BOSS\"
array(1,2)="192.168.16.2"
array(1,3)="3306"
array(1,4)="back"
array(1,5)="asdf1234"
array(1,6)="boss"

rem 192.168.23.25_CHARGE
array(2,0)="192.168.23.25_CHARGE"
array(2,1)="E:\MySQL_BACK\192.168.23.25_CHARGE\"
array(2,2)="192.168.67.25"
array(2,3)="3306"
array(2,4)="back"
array(2,5)="zxcv1234"
array(2,6)="charge"

for a=0 to 3
dates=cdate(date)
times=cdate(time)
For Each i In Split(dates,"/")
If Len(i)<2 Then i1="0"& i Else i1=i
i2=i2+i1
Next
For Each k In Split(times,":")
If Len(k)<2 Then k1="0"& k Else k1=k
k2=k2+k1
Next
set ts=fso.opentextfile(path_logs&i2&".log",8,true)
if a=0 Then
WScript.Sleep 3000
ts.writeline "=================="&date&"==================="
end If
WScript.Sleep 3000
ts.writeline times
WScript.Sleep 3000
ts.writeline "Starting to backup database "&array(a,0)
WScript.Sleep 3000
exec_return=objShell.run("%comspec% /c "&exec_dump&" -h "&array(a,2)&" -P "&array(a,3)&" -u"&array(a,4)&" -p"&array(a,5)&" --skip-lock-tables "&array(a,6)&">"&array(a,1)&""&i2+k2&".bak",0,TRUE)
WScript.Sleep 3000
times=cdate(time)
ts.writeline times
WScript.Sleep 3000
if exec_return=0 Then
ts.writeline "Backup database "&array(a,0)&" completed."
exec_del_ok="1"
elseif exec_return=1 Then
ts.writeline "Backup failed."
elseif exec_return=2 Then
ts.writeline "Account or Password Error."
else
ts.writeline "Unknow error,Error Code is "&exec_return&"."
end If
exec_return=""
i2=""
k2=""

if exec_del_ok="1" Then
dates_add=DateAdd("d", -exec_del_time, dates)
For Each l In Split(dates_add,"/")
If Len(l)<2 Then l1="0"& l Else l1=l
l2=l2+l1
Next
WScript.Sleep 3000
set f=fso.GetFolder(array(a,1))
n=0
for each item in f.files
if lcase(left(item.name,8))=l2 and lcase(right(item.name,4))=".bak" Then
WScript.Sleep 3000
ts.writeline times
ts.writeline "Found backupfile of database "&array(a,0)
WScript.Sleep 3000
objShell.run "%comspec% /c del "&array(a,1)&""&l2&"*.bak"
WScript.Sleep 3000
ts.writeline times
ts.writeline "Deleted backupfile "&l2&"*.bak of database "&array(a,0)
n=n+1
end if
Next
if n=0 Then
WScript.Sleep 3000
ts.writeline times
ts.writeline "Can not find backupfile "&l2&"*.bak of database "&array(a,0)
end If
else
WScript.Sleep 3000
ts.writeline times
ts.writeline "The file of backup was not deleted because the backup was not completed."
end If
exec_del_ok=""
WScript.Sleep 3000
ts.writeline "==================Gorgeous dividing line==================="
ts.close
l2=""
next
