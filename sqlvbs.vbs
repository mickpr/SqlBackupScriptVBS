strComputer = "SERWERPT" 'SQL Server name
const BackupDir = "E:\MSSQL\Backup" 'Folder to back up to. Trailing slash not needed

Set objConnection = CreateObject("ADODB.Connection")

objConnection.Open _
    "Provider=SQLOLEDB;Data Source=" & strComputer & ";" & _
        "Trusted_Connection=Yes;Initial Catalog=Master"

Set objRecordset = objConnection.Execute("Select Name From SysDatabases")

If objRecordset.Recordcount = 0 Then
    Wscript.Echo "No databases could be found."
Else
    Do Until objRecordset.EOF
        if (	objRecordset.Fields("Name")<>"master" AND _
		objRecordset.Fields("Name")<>"model"  AND _ 
		objRecordset.Fields("Name")<>"msdb" AND _ 
		objRecordset.Fields("Name")<>"tempdb") then 
        'if (objRecordset.Fields("Name")="master") then 
           Wscript.Echo objRecordset.Fields("Name")
 	   dbName= objRecordset.Fields("Name")

	   Set objFSO = CreateObject ("Scripting.FileSystemObject")
 	   'get year/month/day into sDate string
  	   backupFileName = dbname & "_" & GetDateTimeString()
  
  	   backupFilePath = BackupDir & "\" & backupFileName & ".bak"
	   'Start new DB command
	   SET cmdbackup = CreateObject("ADODB.Command")
	   cmdbackup.activeconnection = objConnection
	   cmdbackup.CommandTimeout = 3600 ' 3600 sekund
	   'Set command to be executed to generate backup file
	   cmdbackup.commandtext = "backup database " & dbName & " to disk='" & backupFilePath & "'"
	   'Execute DB command to generate file
	   cmdbackup.Execute

        end if
        objRecordset.MoveNext
    Loop
End If



Function GetDateTimeString()
Dim vDate
	vDate = Now()
	GetDateTimeString = Right("0000" & Year(vDate),4) & Right("0" & Month(vDate),2) & Right("0" & Day(vDate),2)
	GetDateTimeString = GetDateTimeString & "-" & Right("00" & Hour(vDate),2)  & Right("00" & Minute(vDate),2)
End Function
