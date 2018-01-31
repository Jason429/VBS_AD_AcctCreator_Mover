'Option Explicit

' This is only an example.  Change the LDAP OU you are pointed to.

'*********** DECLARE VARIABLES **********
'* Instantiate variables
'* Note that we are not going to be using NetBIOS variables
'* Remove them shortly

Dim objExcel,  strExcelPath, objSheet, namedArguments
Dim cfso, mfso, efso
Dim fileCreated, fileMoved, fileError
Dim objRootLDAP

Dim intRow, strInput



'* String for account
'* Account name, Rank, Last Name, Initials 

Dim strAcct, strRank, strLastName, strInitials
Dim strOU

'* String for Position and Unit (if there is a good place for unit, maybe office)
Dim strPos, strUnit, strPW

'* String for GroupDN iteration
Dim strGroupDN, iter

'* Original strings below


'Dim strLast, strFirst,  strInitials, strPW, intRow, intCol
Dim objUser, objGroup, objContainer
Dim strCN, strNTName,  strContainerDN, strDescription, strDisplName
Dim strOffice
Dim strHomeFolder,  strHomeDrive, objFSO, objShell
Dim intRunError,  strNetBIOSDomain, strDNSDomain
Dim objRootDSE,  objTrans, strLogonScript, strUPN

'* For input arguement filenames
Dim f_input, f_created, f_moved, f_error

'*********** CONSTANT FOR FILE HANDLING **********
Const ForReading = 1, ForWriting = 2, ForAppending = 8


'*********** END DECLARE VARIABLES **********


'*********** FUNCTION STATEMENTS HERE **********
'****** START FUNCTION FuncQuit *****
'   *** DISPLAY USAGE ***
Function FuncQuit(errortype, details)
	 If errortype = "help" Then
	    Call MsgBox ("AccountCreator Usage Guide:" & vbCrLf & _
	    	   "Format AccountCreator /input:<file> /created:<file> /moved:<file> /error:<file>" & vbCrLf & _
	    	   "/input:  -The name of the excel file to be used as input" & vbCrLf & _
		   "/created: - The full filename to create for created accounts" & vbCrLF & _
		   "/moved: - The full filename to create file for all moved accounts" & vbCrLf & _
		   "/error: - The full filename to create for any errors while running the program" & vbCrLF & _
		   vbCrLF & _
		   "Notes:" & _
		   "The excel file will only be read from the first workbook in the excel file you use." & vbCrLf & _
		   "MOVE ALL ACCOUNTS to CREATE to FIRST WORKBOOK!!!!!" & vbCrLf & _
		   vbCrLf & _
		   "Created, moved and error will append to the file if it exists.  If the file specified does" & vbCrLf & _
		   "not exist, the file will be created.  These text files will use tabs to seperate columns so" & vbCrLf &_
		   "they can be easily parsed in Microsoft Excel." , vbOkOnly, "AccountCreator Usage Guide")
	WScript.Quit
	End If


End Function
'****** END FUNCTION FuncQuit ********

'****** START FUNCTION SearchName - SEARCH FOR ACCOUNT ***
Function SearchName(sUser)
  Dim objConnection, objCommand, objRecordSet
  Dim arrPath, strDN
  Dim i, strPath, strLen
  Const ADS_SCOPE_SUBTREE = 2
  
  Set objConnection = CreateObject("ADODB.Connection")
  Set objCommand = CreateObject("ADODB.Command")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open "Active Directory Provider"
  
  Set objCommand.ActiveConnection = objConnection
  
  objCommand.Properties("Page Size") = 1000
  objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE

  objCommand.CommandText = _
  	"SELECT distinguishedName FROM 'LDAP://dc=example,dc=root,dc=com' " & _
  	"WHERE objectCategory='user' " & _
  	"AND sAMAccountName='" + sUser + "'"
  	
  Set objRecordSet = objCommand.Execute
  
  On Error Resume Next

  
  objRecordSet.MoveFirst
  ' IF NOT FOUND, RETURN "Not Found"
  
  If Err.Number <> 0 Then
    On Error Goto 0
    Err.Clear
    SearchName = "Not Found"
    Exit Function
  End If
    
  strPath = ""
  
  Do Until objRecordSet.EOF
  	strDN = objRecordSet.Fields("distinguishedName").Value
  	arrPath = Split(strDN, ",")
  	For i = 1 to Ubound(arrPath)
  		strPath = strPath + arrPath(i) + ","
  	Next
  	strLen = Len(strPath)
  	strPath = Left(strPath, strLen - 1)
  	objRecordSet.MoveNext
  Loop

  SearchName = strPath

End Function
'*********END FUNCTION SearchName************

'*********** END FUNCTION STATEMENTS **********

'*********** MAIN PROGRAM **********

'* Check args here
'* If incorrect, run function


If Wscript.Arguments.Count <> 4 Then
   Call FuncQuit("help","")
Else
   Set namedArguments = WScript.Arguments.Named
   f_input = namedArguments.Item("input")
   f_created = namedArguments.Item("created")
   f_moved = namedArguments.Item("moved")
   f_error = namedArguments.Item("error")
End If

'On Error Resume Next
Set cfso = CreateObject("Scripting.FileSystemObject")
Set fileCreated = cfso.OpenTextFile("" & f_created, 8, True)
If Err.Number <> 0 Then
  'On Error GoTo 0
  WScript.Echo " Error is " & Err.Description
  WScript.Echo "Unable to create/append file " & f_created  & vbCrLf
  WScript.Quit
End If

Set mfso = CreateObject("Scripting.FileSystemObject")   
Set fileMoved = mfso.OpenTextFile("" & f_moved, ForAppending, True)
If Err.Number <> 0 Then
  On Error GoTo 0
  WScript.Echo "Unable to create/append file " & f_error & vbCrLf
  WScript.Quit
End If

Set efso = CreateObject("Scripting.FileSystemObject")
Set fileError = efso.OpenTextFile("" & f_error, ForAppending, True)
If Err.Number <> 0 Then
  On Error GoTo 0
  WScript.Echo "Unable to create/append file " & f_error
  WScript.Quit
End If

'* Open spreadsheet.
Set objExcel =  CreateObject("Excel.Application")

On Error Resume Next
objExcel.Workbooks.Open  f_input
If Err.Number <> 0 Then
  On Error GoTo 0
  Wscript.Echo "Unable to open spreadsheet  " & f_input
  Wscript.Quit
End If

On Error GoTo 0

Set objSheet =  objExcel.ActiveWorkbook.Worksheets(1)

'*********** Original Script **********

  On Error GoTo 0
' Start with row 2 of spreadsheet.
  ' Assume first row has column headings.

intRow = 2

' Read each row of spreadsheet until a blank value
  ' encountered in column 5 (the column for cn).
  ' For each row, create user and set attribute values.

Do While objSheet.Cells(intRow, 7).Value <> ""
    

    Dim OU_Exists, AcctCreated
    OU_Exists = True
    AcctCreated = True
    ' Read values from spreadsheet  for this user.

    strAcct = Trim(objSheet.Cells(intRow,  7).Value)
    strRank = Trim(objSheet.Cells(intRow,  4).Value)
    strLastName = Trim(objSheet.Cells(intRow,  5).Value)
    strInitials = Trim(objSheet.Cells(intRow, 6).Value)
    strOU = Trim(objSheet.Cells(intRow,  8).Value)
    strPos = Trim(objSheet.Cells(intRow, 2).Value)
    strUnit = Trim(objSheet.Cells(intRow,  1).Value)
    strPW = Trim(objSheet.Cells(intRow,  10).Value)
    strGroupDn = Trim(objSheet.Cells(intRow, 9).Value) '*Split by ;
    strGroupDn = Split(strGroupDn,";")
    

        
    ' Bind to container where users to be created.
    'On Error Resume Next

    Set objContainer =  GetObject("LDAP://" & strOu & "," & "dc=example,dc=root,dc=com")

    If Err.Number <>  0 Then
        On Error GoTo 0
        fileError.Write Now & vbTab & "Unable to bind to OU: " & _
            vbTab & strOU & vbCrLf
        WScript.Echo "Unable to bind to OU: " & strOU
        OU_Exists = False
        Err.Clear
    End If

    searchResult = SearchName(strAcct)
    'wscript.echo searchResult

    If searchResult <> "Not Found" and OU_Exists = True Then

        objContainer.MoveHere "LDAP://cn=" & strAcct & "," & searchResult, vbNullString
        fileMoved.write Now & vbTab & strAcct & vbTab & " moved to " & vbTab & strOu & vbCrLf

    ElseIf OU_Exists = True Then
	' Create user object.

	On Error Resume Next
	Set objUser =  objContainer.Create("user", "cn=" & strAcct)
	If Err.Number <> 0 Then
	    On Error GoTo 0
    	    fileError.Write Now & vbTab & "Unable to create user with cn: " & _
    	        vbTab & strAcct & vbCrLf
    	    WScript.Echo "Unable to create user with cn: " & strAcct
	    AcctCreated = False
            Err.Clear
	Else

            On Error GoTo 0
	    ' Assign mandatory attributes  and save user object.

            objUser.Put "sAMAccountName", strAcct

	    objUser.SetInfo
	    On  Error Resume Next
	    
	    If Err.Number <> 0 Then
	        On  Error GoTo 0
    	        fileError.Write Now & vbTab & "Unable to create user with NT name: " & _
    	            vbTab & strAcct & vbCrLf
    	        WScript.Echo "Unable to create user with NT name: " & strAcct
    	        AcctCreated = False
    	        Err.Clear
            End If
        End If
    End If

    If AcctCreated = True and OU_Exists = True And searchResult = "Not Found" Then
        ' Set  password for user.

	objUser.SetPassword strPW
	If Err.Number <> 0 Then
	    On Error GoTo 0
	    fileError.Write Now & vbTab & "Unable to set password for user: " & _
	        vbTab & strAcct & vbCrLf
	    Wscript.Echo "Unable to set password  for user " & strAcct
	    Err.Clear
	End If

	On Error GoTo 0
	'  Disable the user account.

	objUser.AccountDisabled = True
	
	' Assign values to remaining attributes.

	
	If strLast <> "" Then
	  objUser.sn = strLast
	End If
	
	objUser.Description = strLastName & " " & strRank & _
	    " " & strInitials & " (" & strUnit & ")"

        objUser.displayName = strAcct

	objUser.physicalDeliveryOfficeName = strUnit
	
        ' Set password expired. Must be changed on next logon.

	objUser.pwdLastSet = 0

        ' Save changes.

        On Error Resume Next
	objUser.SetInfo

	If Err.Number <> 0 Then
	    On Error GoTo 0
	    fileError.Write Now & vbTab & "Unable to set attributes for user with NT name: " & _
	        vbTab & strAcct & vbCrLf
	    Wscript.Echo "Unable to set attributes  for user with NT name: " & _
	        strAcct
	    Err.Clear
	End If

	fileCreated.write Now & vbTab & strAcct & vbTab & "created. Check error logs for usergroup errors." & vbCrLf
	
	On Error GoTo 0
	For iter = 0 to ubound(strGroupDn)
            'msgbox("Into For Loop " & iter)
	 
    	    '***TODO Create function to iterate through group membership
	    ' strGroupDN = Trim(objSheet.Cells(intRow,  intCol).Value)
	    On Error Resume Next
	    Set objGroup =  GetObject("LDAP://" & strGroupDn(iter))

	    If Err.Number <> 0 Then
	        On Error GoTo 0
	        fileError.Write Now & vbTab & "Unable to bind to group: " & strGroupDN(iter) & vbCrLf
	        Wscript.Echo "Unable to bind to group  " & strGroupDN(iter)
	        Err.Clear
	    Else
	    
	        objGroup.Add objUser.AdsPath
	        If Err.Number <> 0 Then
	            On Error GoTo 0
	            fileError.Write Now & vbTab & "Unable to add user to " & strGroupDn(iter) & _
	                " group.  User is " & strAcct & vbCrLf
	            Wscript.Echo "Unable to add user to " & strGroupDn(iter) & _
	                " group.  User is " & strAcct 
	        End If
            End If
	 Next
    End If
    '  Increment to next user.
    intRow = intRow + 1
Loop

Wscript.Echo  "Done"

' Clean up.

  objExcel.ActiveWorkbook.Close
  objExcel.Application.Quit
  Set objUser = Nothing
  Set objGroup = Nothing
  Set objContainer =  Nothing
  Set objSheet = Nothing
  Set objExcel = Nothing
  Set objFSO = Nothing
  Set objShell = Nothing
  Set objTrans = Nothing
  Set objRootDSE =  Nothing
