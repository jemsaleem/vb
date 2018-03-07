Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim file, fso, dict, FileName, row, line, a , filedatestr, filedate, filedatetime, f, fc , objFile, sf, strComputerName , strUserName , curDateTime
Dim strConnection, conn, rs, strSQL
Set wshShell = CreateObject( "WScript.Shell" )

strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
Set fso = CreateObject("Scripting.FileSystemObject")

 
strConnection = "Driver={SQL Server};Server=192.168.9.XX;" & _
"Database=XXXXX;Uid=XXXXX;Pwd=XXXXXX;"
 
 
Set conn = Wscript.CreateObject("ADODB.Connection")
conn.Open strConnection

	Set f = fso.GetFolder("C:\ABC\")
	Set sf = f.SubFolders
	For Each subfolder in sf 
		
		Set f=fso.GetFolder("C:\ABC\" +subfolder.name)
		Set fc = f.Files

		 For Each objFile in fc

		 Filename = f + "\" + objfile.name 
		 'wscript.echo FileName

			Set file = fso.OpenTextFile(Filename , ForReading)
			filedatestr = Split(FileName, "_")
			filedate = filedatestr(1)+"-"+filedatestr(2)+"-"+ left(filedatestr(3), 2)
			'Wscript.echo filedate 
			
			conn.BeginTrans
				row = 1
				Do Until file.AtEndofStream
				  line = file.Readline
				  line = Replace (line,"'", "")
				  a=Split(line,vbTab)
				  filedatetime = filedate + " " +a(1)
				  'Wscript.Echo line
				  conn.Execute = "INSERT INTO Log_XX_Files (Server, Rownumber, Log_line,Severity,Time,Module,Service,Text,Pid,Id,Code,Filename,Date) VALUES ('"&subfolder.name&"', '"&row&"' ,'"&line&"','"&a(0)&"','"&filedatetime&"','"&a(2)&"','"&a(3)&"','"&a(4)&"','"&a(5)&"','"&a(6)&"','"&a(7)&"','"& objfile.name &"', '"& filedate &"' )" 
				  row = row + 1
				Loop
			conn.CommitTrans

			curDateTime = year(Date) & "-" & _
				right("0" & month(date), 2) &  "-" & _
				right("0" & day(date), 2) & " " & _
				right("0" & hour(time), 2) &  ":" & _
				right("0" & minute(time), 2) &  ":" & _
				right("0" & second(time), 2)
			conn.Execute = "INSERT INTO Log_Sapphire_File_Import_History (Server, Filename, DateIngest,Message,Computername,Login) VALUES ('"&subfolder.name&"', '"& objfile.name &"' ,'"& curDateTime &"','"& row &"', '"& strComputerName &"', '"& strUserName &"')" 
			file.close 
			'delete text file processed from directory
			Set file = fso.Getfile(Filename)
			file.Delete
			'append to lofile
			Set file = fso.OpenTextFile( "C:\ABC\log_script_XX.txt" , ForAppending, True)
			file.writeline Now & " " & Filename & " inserted to log database and deleted: " & row & " lines processed"
			file.close
			
		 Next
	Next
	'next file to process
	'close recordset, database 
	'rs.close
	Set rs = Nothing
	conn.close
	Set conn = Nothing
