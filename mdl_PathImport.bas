Attribute VB_Name = "mdl_PathImport"
Function IMPORTExcel()
      
      On Error GoTo ErrChk
    Dim wsOut As Worksheet
    Dim wb As Workbook
    Dim wbDo As Workbook
    Dim wsconfig As Worksheet
    Dim wsToCopy As Worksheet
    Dim wsRelatory As Worksheet
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim i As Integer
    Set wb = ThisWorkbook
    Set wsconfig = wb.Sheets("MASTER")
    
    strPath = wsconfig.Range("tbl_PathImport")
    If (Right(strPath, 1) <> "\") Then strPath = strPath & "\"
    If (Mid(strPath, 2, 2) <> ":\") Then strPath = Environ("userprofile") & "\" & strPath
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strPath)
    For Each objFile In objFolder.Files
        If ((InStr(1, LCase(objFile.Name), ".xls") > 0 Or InStr(1, LCase(objFile.Name), ".csv") > 0 Or InStr(1, LCase(objFile.Name), ".htm") > 0) And InStr(1, LCase(objFile.Name), "~") = 0) Then
            k = 0
            keepon = True
            While (keepon And k < 10)
                keepon = False
                
                Set wbDo = Excel.Workbooks.Open(objFile.Path, ReadOnly:=True, UpdateLinks:=False)
                
                findws = False
                v = 1
                While (findws = False And v <= wbDo.Sheets.Count)
                    Set wsOut = wbDo.Sheets(v)
                    If wsOut.Visible = xlSheetVisible Then
                        findws = True
                    Else
                        Set wsOut = Nothing
                        v = v + 1
                    End If
                Wend
                If (v > wbDo.Sheets.Count) Then Stop 'Not Find Sheet?
                If ((InStr(1, LCase(objFile.Name), ".htm") > 0) And ((wsOut.Range("A1:H100").Find("ó") Is Nothing) Or (wsOut.Range("A1:H100").Find("ó") Is Nothing) Or (wsOut.Range("A1:H100").Find("ó") Is Nothing))) Then
                    keepon = True
                    Set wsOut = Nothing
                    wbDo.Close
                    Set wbDo = Nothing
                    strOldFileName = objFile.Path
                    strNewFileName = ConvertFileCharset(objFile)
                    objFSO.DeleteFile objFile.Path, True
                    Name strNewFileName As strOldFileName
                    Set objFile = objFSO.GetFile(strOldFileName)
                    
                    Set wbDo = Excel.Workbooks.Open(objFile.Path, ReadOnly:=True, UpdateLinks:=False)
                    
                    findws = False
                    v = 1
                    While (findws = False And v <= wbDo.Sheets.Count)
                        Set wsOut = wbDo.Sheets(v)
                        If wsOut.Visible = xlSheetVisible Then
                            findws = True
                        Else
                            v = v + 1
                        End If
                    Wend
                    If (v > wbDo.Sheets.Count) Then Stop 'Not Find Sheet?
                End If
                k = k + 1
            Wend
            
            If ((Len(wsOut.Range("A1").Value) - Len(Replace(wsOut.Range("A1").Value, ";", ""))) > 8) Then
                wbDo.Close
                Set wbDo = Nothing
                Set wsOut = Nothing
                ChangeCSVCharacter (objFile.Path)

                Set wbDo = Excel.Workbooks.Open(objFile.Path, ReadOnly:=True, UpdateLinks:=False)

                findws = False
                v = 1
                While (findws = False And v <= wbDo.Sheets.Count)
                    Set wsOut = wbDo.Sheets(v)
                    If wsOut.Visible = xlSheetVisible Then
                        findws = True
                    Else
                        v = v + 1
                    End If
                Wend
                If (v > wbDo.Sheets.Count) Then Stop 'Not Find Sheet?
            End If
            If (wsOut.ProtectContents = False And objFile.Size < 1000000) Then
               
                If (InStr(1, objFile.Name, "-HIGH-") > 0) Then
                    wsOut.Copy Before:=wb.Sheets(1)
                Else
                    wsOut.Copy Before:=wb.Sheets("MASTER")
                End If
       
            Else
             
                If (InStr(1, objFile.Name, "-HIGH-") > 0) Then
                    Set wsToCopy = wb.Sheets.Add(After:=wb.Sheets(1))
                Else
                    Set wsToCopy = wb.Sheets.Add(After:=wb.Sheets("MASTER"))
                    Set wsRelatory = wb.Sheets.Add(After:=wb.Sheets(2))
                End If
                
                wsOut.Range("A1:CZ100000").Copy
                wsToCopy.Range("A1:CZ100000").PasteSpecial xlPasteValuesAndNumberFormats
                On Error Resume Next
                wsToCopy.Range("A1:CZ100000").PasteSpecial xlPasteColumnWidths
                On Error GoTo ErrChk
                wsToCopy.Range("A1:CZ100000").PasteSpecial xlPasteFormats
                wsToCopy.Name = wsOut.Name
            End If
            
            wsRelatory.Name = "Sheet_" & wb.Sheets.Count - 1
            wsRelatory.Range("1:1").Insert Shift:=xlDown
            wsRelatory.Range("A1").Value = wbDo.FullName
            
            Set wsOut = wb.Sheets(wsOut.Name)
            wsOut.Name = "Sheet_" & wb.Sheets.Count
            wsOut.Range("1:1").Insert Shift:=xlDown
            wsOut.Range("A1").Value = wbDo.FullName

            wbDo.Close
            
            Call TemplateRelatoryMount(wsToCopy, wsRelatory)
            
        End If
        On Error Resume Next
        wbDo.Close
        On Error GoTo ErrChk
        Set wbDo = Nothing
        
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.Run "cmd.exe /c echo. >NUL | clip", 0, True
        Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 1)
    Next
    
Exit Function

ErrChk:
    MsgBox Err.Description
    Stop
    Resume

End Function

Function ConvertFileCharset(ByVal oFile As Object) As String

    On Error GoTo ErrChk
    
   Const adTypeBinary = 1
   Const adTypeText = 2
   Const bOverwrite = True
   Const bAsASCII = False
 
   Dim oFS: Set oFS = CreateObject("Scripting.FileSystemObject")
 
   Dim oFrom: Set oFrom = CreateObject("ADODB.Stream")
   Dim sFrom: sFrom = "Windows-1252"
   Dim sFFSpec: sFFSpec = oFile.Path
   Dim oTo: Set oTo = CreateObject("ADODB.Stream")
   Dim sTo: sTo = "utf-8"
   Dim sTFSpec: sTFSpec = oFile.ParentFolder.Path & "\utf" & oFile.Name
 
   oFrom.Type = adTypeText
   oFrom.Charset = sFrom
   oFrom.Open
   oFrom.LoadFromFile sFFSpec
   strFullText = oFrom.ReadText
   oFrom.Close
 
   oTo.Type = adTypeText
   oTo.Charset = sTo
   oTo.Open
   oTo.WriteText strFullText
   oTo.SaveToFile sTFSpec
   oTo.Close
   
   ConvertFileCharset = sTFSpec
   
   Exit Function
   
ErrChk:
    MsgBox Err.Description
    Stop
    Resume

End Function

Sub ChangeCSVCharacter(ByVal filePath As String)

    On Error GoTo ErrChk

    sTemp = ""
    Open filePath For Input As #1
    Do Until EOF(1)
        Line Input #1, sBuf
        sTemp = sTemp & sBuf & vbCrLf
    Loop
    Close #1
    
    sTemp = Replace(Replace(sTemp, ",", "."), ";", ",")
    
    Open filePath For Output As #1
    Print #1, sTemp
    Close #1
    
    Exit Sub

ErrChk:
    MsgBox Err.Description
    Stop
    Resume

End Sub

Function CleanPreviousFiles()

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook
    On Error Resume Next
    Sheets("MASTER").Select
    
    For Each mySheet In wb.Sheets
        Set ws = mySheet
        If (Left(ws.Name, 5) = "Sheet") Then
            ws.Delete
        End If
    Next

End Function





