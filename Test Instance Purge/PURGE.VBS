'TEST INSTANCE PURGE TOOL EXAMPLE
'******
purgeName = "<Email Subject>"
purgeUser = "<QC User for email notification>"
purgeKey = "<QC API Key>"
purgeSecret = "<QC API Secret>"
serverURL = "<QC Server URL>"
domainName = "<QC Domain name>"
projectName = "<QC Project name>"
'******

dim fs
set fs=CreateObject("Scripting.FileSystemObject")
Set stdout = fs.GetStandardStream (1)
Dim tdc
Set tdc = CreateObject("TDApiOle80.TDConnection")
 
 'Connect to server
 tdc.InitConnectionEx serverURL
 
 'Authenticate user
 tdc.Login purgeKey, purgeSecret
 
 ' Connect to project
 tdc.Connect domainName, projectName
stdout.writeline("CONNECTED TO PROJECT!!!")

'QUERY SETUP
  myQuery = WScript.Arguments.Item(0)
set com = tdc.Command
com.CommandText = myQuery
'Execute the query and get the recordset.
    Set RecSet = com.Execute
    ColCnt = RecSet.ColCount
    RecCnt = RecSet.RecordCount
stdout.writeline("we have items: " & RecCnt)
RecSet.First

'LOOP THROUGH QUERY ITEMS AND PURGE
myErr = 0
For i = 1 To RecCnt
stdout.writeline(i & "/" & RecCnt &" working on Test Set : " & RecSet.FieldValue(0) & " and instance : " & RecSet.FieldValue(1))
On Error Resume Next 
'*****PURGE*****
        'Set filter for open test sets.
        Set tsFact = tdc.TestSetFactory
        Set tsFilter = tsFact.Filter
        tsFilter.Filter("CY_CYCLE_ID") = RecSet.FieldValue(0)

        'Set filter for runs that did not pass.
        Set runFact = tdc.RunFactory
        Set runFilter = runFact.Filter
        runFilter.Filter("RN_TESTCYCL_ID") = RecSet.FieldValue(1)

'PURGE SET TO KEEP THE LAST 5 RUNS
'READ MORE HERE: https://admhelp.microfocus.com/alm/api_refs/ota/Content/ota/topic9008.html
        myResult = tdc.PurgeRuns2(tsFilter.Text, runFilter.Text, 5,1,0,False)
 '*****PURGE*****
If err.number <> 0 Then
stdout.writeline("ERROR : " & Err.Description)
	If myErr = 0 Then
		myErr = 1
		RecSet.Next
	End If
Else
 RecSet.Next
End If
Next

If myErr = 1 Then
	tdc.SendMail purgeUser,purgeUser,"ERROR: " & purgeName, "Please run again"
	stdout.writeline("DONE WITH ERRORS!!!")
Else
	tdc.SendMail purgeUser,purgeUser,"COMPLETED: " & purgeName, "Ready for next purge"
	stdout.writeline("DONE!!!")
End If

' Disconnect from the project. 
 tdc.Disconnect
 
 ' Log the user off the server
 tdc.Logout
 
 'Release the TDConnection object.
 tdc.ReleaseConnection
 
 ' Release the object
 Set tdc = Nothing