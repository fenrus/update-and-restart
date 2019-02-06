'
' This is a script to copy an access database to a local machine from dfs share.
' The script checks the dfs share each run for a new version of the DB.
' if there is a newer version than we have locally, we'll replace the local one
' with the one from the server.
' The old local version will be backed up to ".bak" (only one generation backwards is kept)
'
' All local script data is saved under "localpath"
'
' David Syk.
' 2018-02-19
'


' Set a copule of variables required for the script to run
Set oShell = CreateObject("WScript.Shell")
strHomeFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")

localpath = strHomeFolder & "\AppData\Local\dbname\"
serverpath = "\\serverpath\goes\here\"

dim localfile, serverfile
localfile = localpath & "database.mdb"
localfilebak = localfile & ".bak"

serverfile = serverpath & "database.mdb"

' 1.  Verify that the folder where we want to store the file exist
'	if no: 	create folder
'				Copy the database there. verify md5. Execute.
'	if yes: go to 2.
' 2. Check if the file we want to execute exist, if no 
'    verify md5 on the local copy of the file and check with file on network share
' if yes: execute the local copy
' if no: copy the file from network to local disk. Verify md5. execute if OK.

' FSO for FileSystemObject handling stuff. (copy, move, remove etc)
Set fso = CreateObject("Scripting.FileSystemObject")

' Do we have access to the serverpath defined above?
dfsexist = fso.FolderExists(serverpath)
if (dfsexist) then
else
	msgBox "Error: Can not access " & serverpath
	WScript.Quit 1
End if

' Does the local folder for storing the database exist ?
exists = fso.FolderExists(localpath)
if (exists) then 

	' Lets check if there is a local copy of the file.
	exists2 = fso.FileExists(localfile)
	if (exists2) then 
		' if there's a file in place, we're not doing anything here. It happens further down in the code.
	Else
		' Since there is no file in place we're copying the one from the server.
		fso.CopyFile serverfile, localfile
	End if

	'Compare the files based on modification date
	Dim filesys,demofile,localdate,serverdate
	Set filesys = CreateObject("Scripting.FileSystemObject")
	' DEBUG: 
	' MsgBox localfile & " exist"
	Set demofile = filesys.GetFile(localfile)
	localdate = demofile.DateLastModified
	Set demofile = filesys.GetFile(serverfile)
	serverdate = demofile.DateLastModified
	
	If DateDiff("s", localdate, serverdate) >= 1 Then
		' serverdate is more recent than localdate, comparison is on second level
		' here we will copy the newer file from the server to the client
		' DEBUG:
		' msgBox serverfile & " modified " & serverdate & " is newer than " & localfile & " modified " & localdate

		' rename the current local copy to	have .bak at the end and blindly delete the old .bak copy.
		if fso.FileExists(localfilebak) then
			fso.DeleteFile(localfilebak)
        end if
		
		' Backup the old local copy to .bak
		fso.MoveFile localfile, localfilebak
		' Copy the latest version from server
		fso.CopyFile serverfile, localfile
		
	else
		' The local copy  of the file have the same modification timestamp as the one on server.
		' We consider it the same. 
		' We can now go ahead and start the application outside the if-then-else statements.
	End If
	
else
	' if the folder under AppData in profile does not exist it needs to be created.
	' let's create it and copy the latest database from the server into it.
	Set objFolder = fso.CreateFolder(localpath)

	' Copy latest version from server
	fso.CopyFile serverfile, localfile

	msgBox localpath & " have been created, And application copied from server. Will start database now."
end if

' DEBUG: msgBox "lets run the db!"
' All different if's and buts should now have led to that we have the latest file in place.
' Then, lets start it.
Dim oShell
Set oShell = WScript.CreateObject ("WSCript.shell")
oShell.run localfile