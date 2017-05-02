'------------------------------------------------------------------------------------------------------------------------------------------------------------------'
'SCRIPT WILL ACCEPT COST CENTRE REFERANCE FILE AND ATTEMPT TO VERIFY THE COST CENTRES AGAINST A MASTER/REVIEW FILE
'AUTHOR: JEFFERY KOZERA
'VERSION 4.1
'LAST UPDATE: APRIL 28TH 2017
'------------------------------------------------------------------------------------------------------------------------------------------------------------------'

'BOOLEANS--------------------------------------------------------------------------------
Const delHead = false 'SHOULD THE HEADER 'IF IT EXISTS' BE DELETED?
Const isAuto = true 'IS THIS SCRIPT BEING RUN  BY A USER WHO NEEDS FEEDBACK? OR THROUGH AN AUTOMATED PROCESS? (FALSE = USER, TRUE = AUTOMATED)
Const igSpecial = TRUE 'ENABLE THIS TO IGNORE SPECIAL CHARACTERS AND SPACES DURING COST CENTRE VERIFICATION
Const cleanCC = FALSE 'ENABLE THIS TO REMOVE SPECIAL CHARACTERS AND SPACES WHEN RETURNING/WRITING THE COST CENTRE TO THE RESULT FILE

'READ/WRITE LOCATIONS--------------------------------------------------------------------
Const CCENTRE_FILE_LOCATION = "E:\PP_Processes\All\COST_CENTRE_MASTER\cc.csv" 'FULL PATH TO CSV FILE THAT HAS ALL OF THE COST CENTRES
Dim REVIEW_FILE_LOCATION: REVIEW_FILE_LOCATION = Watch.GetJobFileName 'FULL CSV PATH TO THE FILE THAT HAS COST CENTERS REQUIRING VALIDATION
Dim RESULT_FILE_LOCATION: RESULT_FILE_LOCATION = Watch.GetJobFileName 'FULL CSV PATH FOR RESULTING FILE (THIS CAN BE THE SAME AS THE REVIEW FILE IF DESIRED)

'Const CCENTRE_FILE_LOCATION = "C:\Users\kozeraje\Desktop\Script test\cc\cc.csv" 'FULL PATH TO CSV FILE THAT HAS ALL OF THE COST CENTRES
'Dim REVIEW_FILE_LOCATION: REVIEW_FILE_LOCATION = "C:\Users\kozeraje\Desktop\Script test\IN\IN.CSV" 'FULL CSV PATH TO THE FILE THAT HAS COST CENTERS REQUIRING VALIDATION
'Dim RESULT_FILE_LOCATION: RESULT_FILE_LOCATION = "C:\Users\kozeraje\Desktop\Script test\OUT\OUT.CSV" 'FULL CSV PATH FOR RESULTING FILE (THIS CAN BE THE SAME AS THE REVIEW FILE IF DESIRED)

'STRINGS---------------------------------------------------------------------------------
Const ccDel = "," 'DELIMITER USED IN COST CENTRE FILE
Const rfDel = "," 'DELIMITER USED IN REVIEW FILE
Const invalidStr = "NO" 'STRING VALUE TO USE WHEN INVALID
Const validStr = "YES" 'STRING VALUE TO USE WHEN VALID

'REGEX-----------------------------------------------------------------------------------
Set re = New RegExp: re.Pattern = """[^""]*,[^""]*""": re.Global = True 'REGEX TO DETECT COMMAS BETWEEN DOUBLE QUOTES (USE THIS TO PREVENT SPLITTING ON VALUES WITHIN QUOTES)

'INTEGERS--------------------------------------------------------------------------------
'NOTE: COLUMNS IN THIS SCRIPT START AT '0', SAME CONCEPT AS ARRAYS
Const ccColumn = 1 'COLUMN WHERE COST CENTRES ARE STORED IN THE COST CENTRE FILE
Const rcColumn = 17  'COLUMN WHERE COST CENTRES ARE STORED IN THE REVIEW FILE
Const rrColumn = 18  'COLUMN WHERE RESULT OF VERIFICATION WILL BE WRITTEN
Const rrTotalColums = 20   'TOTAL AMOUNT OF COLUMNS REQUIRED IN REVIEW FILE

'COUNTERS FOR USER FEEDBACK
Dim tValidcc: tValidcc = 0 'VALID
Dim tInvalidcc: tInvalidcc = 0 'INVALID

'USE THESE FOR FSO READ WRITE
Const FORAPPENDING = 8
Const FORREADING = 1
Const FORWRITING = 2

'ARRAYS----------------------------------------------------------------------------------
Dim costArr 'ARRAY WIL BE USED TO STORE VALID COST CENTRES
Dim outArr 'ARRAY WILL BE USED TO STORE RESULTING FILE FOR OUTPUT (APPEND TO THIS WHEN DONE PROCESSING INPUT LINES)

'OBJECTS--------------------------------------------------------------------------------
Dim objProgressMsg 'OBJECT FOR PROGRESS WINDOW
set fso = CreateObject("Scripting.FileSystemObject")



'CALL  STARTUP  SUBROUTINE
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
CALL Start_Check()
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



'===================================================================================================================================================='
Sub Start_Check()
'===================================================================================================================================================='
	
	
'___________________________________________________________________________________________'
'CHECK IF COST CENTRE FILE EXISTS
	If not fso.FileExists(CCENTRE_FILE_LOCATION) Then
		
		If isAuto = false then
			MsgBox CCENTRE_FILE_LOCATION & " NOT FOUND! SCRIPT HAS TERMINATED",VBExclamation, "COST CENTRE CSV"
		else
			Call Err.Raise(vbObjectError + 10, CCENTRE_FILE_LOCATION & " NOT FOUND! SCRIPT HAS TERMINATED")
		end if
		
		
'CHECK IF REVIEW FILE EXISTS
	elseif not fso.FileExists(REVIEW_FILE_LOCATION) Then
		
		If isAuto = false then
			MsgBox REVIEW_FILE_LOCATION & " NOT FOUND! SCRIPT HAS TERMINATED",VBExclamation, "REVIEW CSV"
		else
			Call Err.Raise(vbObjectError + 10, REVIEW_FILE_LOCATION & " NOT FOUND! SCRIPT HAS TERMINATED")
		end if
		
	else
		
'run main sub to create new fsu file
		If isAuto = false then: ProgressMsg"VALIDATING COST CENTRES, PLEASE WAIT...", "CC VALIDATOR V4": end if 'show processing window
		
		Call Main()
		
		If isAuto = false then
			ProgressMsg"", "CC VALIDATOR V4"  'CLOSE PROCESSING WINDOW
			MsgBox("CC VALIDATOR HAS COMPLETED SUCCESSFULLY :D" & vbCrLf & vbCrLf  & "TOTAL PROCESSED: " &  tValidcc + tInvalidcc & vbCrLf & "VALID: " & tValidcc &  vbCrLf & "INVALID: " &  tInvalidcc) 'SHOW RESULTS TO USER
		end if
		
	End If
'___________________________________________________________________________________________'
	
'===================================================================================================================================================='
end Sub 'end sub Start_check'


'================================================================================'
Sub Main
'================================================================================'
	
	Dim ss 'TEMP VAR TO HOLD CURRENT LINE STRING FOR POST PROCESSING
	Dim sa 'TEMP VAR TO HOLD LINE ARR WHEN SPLIT
	Dim cc 'TEMP VAR TO HOLD CONST CENTRE FOR PROCESSING
	
'FILL VALID COST CENTRE ARRAY USING INIT_COSTARR FUNCTION WITH THE PROVIDED PARAMS
	costArr = Init_CostArr( CCENTRE_FILE_LOCATION , ccColumn, ccDel, igSpecial)
	
'OPEN REVIEW FILE FOR 'READING ONLY'
	Set sr = fso.OpenTextFile(REVIEW_FILE_LOCATION, FORREADING)
	
'LOOP THROUGH REVIEW FILE AND CHECK COST CENTRE. WRITE RESULTING LINE TO OUT ARRAY
'___________________________________________________________________________________________'
	Do Until sr.AtEndOfStream
		
'SPLIT CURRENT LINE ON COMMA TO ARR (READ LINE FIRST BEFORE HEADER SKIP DETECTION)
		ss = sr.ReadLine		
		ss = re.Replace(ss, GetRef("ReplaceCallback")) 'REMOVE ALL COMMAS BETWEEN DOUBLE QUOTES
		sa = Split(ss, rfDel) 'SPLIT ON GIVEN REVIEW FILE DELIMITER		
		
'REDIM SPLIT INPUT LINE TO SIZE REQUIRED FOR OUTPUT (IF REQUIRED)
		if  UBound(sa) < rrTotalColums then
			
'DETETMINE THE REQUIRED REDIM SIZE TO ADD MISSING COLUMNS
			Dim newSize: newSize = (rrTotalColums) - UBound(sa)
			
'REDIM LINE ARRAY TO REQUIRED SIZE TO PREVENT OUT OF BOUNDS EXCEPTION
			ReDim Preserve sa(UBound(sa) + newSize)
			
'IF ARRAY HAS MORE VALUES THAN EXPECTED COLUMNS THROW EXCEPTION				
			elseif UBound(sa) > rrTotalColums then
			
			Call Err.Raise(vbObjectError + 10, "ARRAY EXCEEDS EXPECTED COLUMN COUNT! SCRIPT HAS TERMINATED")
			
		end if
		
'CLEAN THE ARRAY OF POSSIBLY UNWANTED CHARACTERS (THIS IS TO AVOID ERRORS UPON RECONSTRUCTION/JOIN)				
		sa = CleanArr(sa)
		
'DETECT IF HEADER LINE AND DELETE IF ENABLED, OTHERWISE CONTINUE
		if sa(rRColumn) <> invalidStr and sa(rRColumn) <> validStr and sa(rRColumn) <> "" then
			
			if delHead = False then : Call push(outArr,(join(sa, rfDel))) : end if
			
		else
			
			cc = sa(rcColumn) 'EXTRACT VALUE
			
'CHECK IF COST CENTRE IS VALID USING THE COST_LOOKUP FUNCTION (MAKE SURE COST CENTRE ARRAY HAS ALREADY BEEN INITIATED USING THE INIT FUNCTION)
			If CostArr_Lookup(costArr,  cc, igSpecial) = true then
				
'POPULATE THE RESULT COLUMN WITH  SPECIFIED VALID STRING
				sa(rrColumn) = validStr
				tValidcc = tValidcc + 1
			else
				
'POPULATE THE RESULT COLUMN WITH  SPECIFIED INVALID STRING
				sa(rrColumn) = invalidStr
				tInvalidcc = tInvalidcc + 1
			end if
			
'CLEAN COST CENTRE STRING IF OPTION IS ENABLED					
			If cleanCC = TRUE then
				
				sa(rcColumn) = RemoveAlphaNumeric(cc) 'REMOVE ALL SPECIAL CHARACTERS (INCLUDING SPACES)
				
			end if
			
			
			Dim outStr: outStr = join(sa, rfDel) 'JOIN ENTIRE LINE USING THE DELIMITER SPECIFIED
			
'REMOVE ALL TRAILING DELIMITERS (EXCEL ABSOLUTLY HATES THESE)
			Do While Right(outStr, 3) = """,""" 
				
				outStr = Left(outStr, Len(outStr) - 3) 
				
			Loop   
			
			if rfDel = """,""" then
				outStr = """" & outStr & """" ' ADD DOUBLE QUOTE TO BEGINNING AND END if using quoted delimiter
			end if
			
			
'PUSH MODIFIED LINE TO OUTPUT ARRAY
			Call push(outArr,outStr)
			
		end if
		
		
	Loop
'___________________________________________________________________________________________'
	
	sr.close 'CLOSE READER
	
'CREATE RESULT FILE IF IT DOES NOT ALREADY EXIST
	If not(fso.FileExists(RESULT_FILE_LOCATION)) Then
		fso.CreateTextFile(RESULT_FILE_LOCATION)
	end if
	
'OPEN STREAM WRITER
	
	
	If IsArray(outArr) then
		
		Set sw = fso.OpenTextFile(RESULT_FILE_LOCATION, FORWRITING)
		
'OUTPUT ALL LINES FROM OUTPUT ARRAY TO RESULT FILE
		for i = 0 to ubound(outArr)
			
			sw.WriteLine(outArr(i))
			
		next
		
		sw.close 'CLOSE WRITER
		
	end if
	
	
	
	
	
'================================================================================'
End Sub
'================================================================================'






'================================================================================'
Function Init_CostArr(ccFile, column, delmiter, igSpecial)
'================================================================================'
'FUNCTION WILL ACCEPT A COST CENTRE 'FILE PATH' AND ATTEMPT TO PARSE THAT FILE TO AN ARRAY
'PARAMS:
'CCFILE - (REQUIRED) FILE TO BE USED AS REFERANCE (THAT STORES MASTER LIST OF COST CENTRES)
'COLUMN - (REQUIRED) PROVIDE THE COLUMN IN THE CSV TO PARSE INTO ARRAY
'DELIMITER - (REQUIRED) PROVIDE THE DELIMITER USED IN THE CSV
'IGSPECIAL - (REQUIRED) BOOLEAN VALUE ON WHETHER TO REMOVE SPECIAL CHARACTERS OR NOT (INCLUDING SPACES)
'NOTE THAT THE CONST CENTRES NEED TO BE IN THE SECOND COLUMN
'RESULT:
'ARRAY OF COST CENTRES
	
	set fso = CreateObject("Scripting.FileSystemObject")
	
	Dim sa 'TEMP VAR WILL STORE LINE WHEN SPLIT
	Dim ccv 'TEMP VAR TO STORE COST CENTRE FOR PROCESSING
	Dim ca 'RESULTING COST CENTRE ARRAY THE WILL BE RETURNED
	
'CHECK THAT THE SUPPLIED 'COLUMNS' & 'DELIMITER' VARS ARE VALID
	If (IsNumeric(column) = False) then
		
'CALL EXCEPTION
		Call Err.Raise(vbObjectError + 10, "INIT_COSTARR", "ERROR WITH SUPPLIED COLUMN (" & column & ") IS NUM = " & IsNumeric(column) & ", OR DELIMITER (" & delmiter & ")" )
		
	end if
	
'CHECK IF CCFILE TRULY EXISTS (BECAUSE WE ACTUALLY NEED IT TO CONTINUE)
'___________________________________________________________________________________________'
	If not(fso.FileExists(ccFile)) Then
		
'CALL EXCEPTION
		Call Err.Raise(vbObjectError + 10, "INIT_COSTARR", "ERROR WITH PATH SUPPLIED")
		
	else
		
'OPEN CONST CENTRE FILE FOR 'READING ONLY'
		Set objFile = fso.OpenTextFile(ccFile, FORREADING)
		
'STREAM TEXTFILE AND COLLECT COST CENTRES ON SECOND COLUMN
'__________________________________________________________________'
		Do Until objFile.AtEndOfStream
			
'SPLIT CURRENT LINE ON COMMA TO ARR
			sa = Split(objFile.ReadLine, delmiter)
			
			ccv = sa(column) 'EXTRACT VALUE
			ccv = Trim(ccv) 'REMOVE LEADING AND TRAILING WHITE SPACES
			ccv = Replace(ccv, """", "") 'REMOVE ANY DOUBLE QUOTES FROM VALUE
			
			if igSpecial = TRUE then
				
				ccv = RemoveAlphaNumeric(ccv) 'REMOVE ALL SPECIAL CHARACTERS (INCLUDING SPACES)
				
			end if					
			
			
'APPEND COST CENTRE TO ARRAY
			Call push(ca, ccv)
			
		Loop
'__________________________________________________________________'
		
		objFile.close 'CLOSE STREAM
		
	End if
'___________________________________________________________________________________________'
	
'VALIDATE THAT AT LEAST ONE COST CENTRE WAS PARSED INTO THE ARRAY BY CHEKCING ITS LENGTH
	If IsArray(ca) And UBound(ca) > 0 then
		
'RETURN THE RESULTING COST CENTRE ARRAY
		Init_CostArr =  ca
	Else
		
'CALL EXCEPTION WHEN ARRAY IS EMPTY
		Call Err.Raise(vbObjectError + 10, "INIT_COSTARR", "RESULTING ARR EMPTY (CHECK CC FILE)")
		
	end if
	
'================================================================================'
End Function 'END INIT_COSTARR FUNCTION
'================================================================================'


'================================================================================'
Function CostArr_Lookup(costArr, costCentre, igSpecial )
'================================================================================'
'FUNCTION WILL ACCEPT A COST CENTRE 'COSTARR' AND ATTEMPT TO DERTERMIN IF THE 'COSTCENTRE' VALUE EXISTS WITHIN IT
'PARAMS:
'COSTARR - (REQUIRED) 1 DIMENTIONAL ARRAY OF COST CENTRES
'COSTCENTRE - (REQUIRED) STRING VALUE TO LOOK FOR IN ARRAY
'IGSPECIAL - (REQUIRED) BOOLEAN VALUE ON WHETHER TO IGNORE SPECIAL CHARACTERS OR NOT (INCLUDING SPACES)
	
'RESULT:
'BOOLEAN RESULT IS RETURNED IF THE COST CENTRE IS VALID OR NOT
	
'DEFAULT RETURN VALUE TO FALSE UNLESS A RESULT IS FOUND
	Dim r: r = false
	
'CLEAN UP GIVEN COST CENTRE
	costCentre = Trim(costCentre) 'REMOVE LEADING AND TRAILING WHITE SPACES
	costCentre = Replace(costCentre, """", "") 'REMOVE ANY DOUBLE QUOTES FROM VALUE
	
	if igSpecial = TRUE then
		
		costCentre = RemoveAlphaNumeric(costCentre) 'REMOVE ALL SPECIAL CHARACTERS (INCLUDING SPACES)
		
	end if
	
'VALIDATE THAT AT LEAST ONE COST CENTRE WAS PARSED INTO THE ARRAY BY CHEKCING ITS LENGTH
'___________________________________________________________________________________________'
	If not IsArray(costArr) And UBound(costArr) = 0 then
		
'CALL EXCEPTION WHEN ARRAY IS EMPTY
		Call Err.Raise(vbObjectError + 10, "COSTARR_LOOKUP", "PROVIDED ARR EMPTY (CHECK CC ARR)")
	Else
		
'LOOP THROUGH COST CENTRE ARRAY AND COMPAIR WILL GIVEN VALUE'
'__________________________________________________'
		For i = 0 To UBound(costArr)
			
'MsgBox("COMPAIRING " &  Trim(costCentre) &  " TO " & Trim(costArr(i)))
			
'COMPAIR GIVEN VALUE WITH CURRENT INDEX IN ARR
			if Trim(costCentre) = Trim(costArr(i)) then
				
'WHEN CONDITION IS REACHED EXIT AND RETURN TRUE
				r = true
				Exit For
			end if
			
		Next
'__________________________________________________'
		
	end if
'___________________________________________________________________________________________'
	
'IF HERE THEN NO RESULT WAS FOUND, EXIT AND RETURN FALSE
	CostArr_Lookup = r
	
'================================================================================'
End Function 'END COSTARR_LOOKUP FUNCTION
'================================================================================'


'================================================================================'
Function push(ByRef arr, ByVal row) 'vbscript does not support array push function, therefore we need to implement one ourselves :)
'================================================================================'
	Dim cellValue
	Dim r: r=0
	Dim c: c=0
	Dim dummyVar
	
	If IsArray(arr) Then
		
		On Error Resume Next
		dummyVar = UBound(arr, 2) 'DETERMINE IF THIS IS A 2D MATRIX ARRAY
		
		If (CLng(Err.Number) > 0) = TRUE Then 'PROCESS 1D ARRAY HERE
			On Error Goto 0
			c = UBound(arr) + 1
			Redim Preserve arr(c)
			arr(UBound(arr)) = row
			
		Else '2D ARRAYS PROCESS HERE
			On Error Goto 0
			Redim Preserve arr(UBound(arr), UBound(arr, 2) + 1) 'ADD A NEW ROW/INDEX
			If IsArray(row) Then
				For r = 0 To UBound(row) 'RUN THROUGH EACH COLUMN
					arr(r, UBound(arr,2)) = row(r) 'ADD CELL
				Next
			Else
				arr(0, UBound(arr,2)) = row
			End If
			
		End If
		
	Else
		
		If IsArray(row) Then
			
			If UBound(row) < 2 Then 'IF ARRAY THEN SHOLD ONLY HAVE ONE INDEX
				arr = row
			End If
		Else
			ReDim arr(0)
			arr(0) = row
		End If
		
	End If
	
	If IsArray(arr) Then push = UBound(arr)
	
'================================================================================'
End Function
'================================================================================'


'================================================================================'
Function ProgressMsg( strMessage, strWindowTitle ) 'will create a progress window that can be called, closed or updated
'================================================================================'
	
	Set wshShell = WScript.CreateObject( "WScript.Shell" )
	strTEMP = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
	If strMessage = "" Then
		On Error Resume Next
		objProgressMsg.Terminate( )
		On Error Goto 0
		Exit Function
	End If
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strTempVBS = strTEMP + "\" & "Message.vbs"
	
	Set objTempMessage = objFSO.CreateTextFile( strTempVBS, True )
	objTempMessage.WriteLine( "MsgBox""" & strMessage & """, 4096, """ & strWindowTitle & """" )
	objTempMessage.Close
	
	On Error Resume Next
	objProgressMsg.Terminate( )
	On Error Goto 0
	
	Set objProgressMsg = WshShell.Exec( "%windir%\system32\wscript.exe " & strTempVBS )
	
	Set wshShell = Nothing
	Set objFSO   = Nothing
	
'================================================================================'
End Function
'================================================================================'


'================================================================================'
Function RemoveAlphaNumeric(str) 'REMOVE ALL NON ALPHANUMERIC CHARACTERS FROM SUPPLIED STRING
'================================================================================'
	
	strAlphaNumeric = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	
	Dim clStr
	
	For i = 1 to len(str)
		
		strChar = mid(str,i,1)
		
		If instr(strAlphaNumeric,strChar) Then
			
			clStr = clStr & strChar
			
		End If
		
	Next
	
	RemoveAlphaNumeric = clStr
'================================================================================'
End Function
'================================================================================'


'================================================================================'
Function CleanArr(a) 'CLEANS EVERY STRING IN ARRAY BY REMOVING ERRONIOUS COMMAS AND DOUBLE QUOTES
'================================================================================'
	
	for i = 0 to ubound(a)
		
		a(i) = Trim(a(i))
		a(i) = Replace(a(i) , """",  "")
		a(i) = Replace(a(i) , ",",  "")
		
	next
	
	CleanArr = a
	
'================================================================================'
End Function
'================================================================================'

'================================================================================'
Function ReplaceCallback(match, position, all)
'================================================================================'
	ReplaceCallback = Replace(match, ",", " ")
'================================================================================'
End Function
'================================================================================'

