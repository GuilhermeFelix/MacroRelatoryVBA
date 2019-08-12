Attribute VB_Name = "mdl_GetSap"
Function GetRelatoryinSAP()


    On Error GoTo ErrChk

    Dim ws As Worksheet
    Dim wb As Workbook
    'Dim mySheet As Worksheet
    'Dim r As Integer
    Dim k As Integer
    'Dim strErrLog As String
    'Dim iRdr As Integer
    'Dim rsReturn As Object
    'Dim aValues() As String
    'Dim aFields() As String
    'Dim aTypes() As String
    'Dim objNS As Object
    'Dim olFolder As Object
    
    r = 0
    Set wb = ThisWorkbook
    cIsSchedule = ThisWorkbook.Sheets("Config").Range("prfScheduleStatus").Value
    
    
    '========= Open
            On Error Resume Next
            strSAPID = session.ID
            If Err.Number <> 0 Then SAPConnected = False
            On Error GoTo ErrChk
            If Not (SAPConnected) Then
                On Error Resume Next
                k = 0
                While Not (SAPConnected) And k < 10
RetrySAPConn:
                    iErr = 0
                    Set SapGuiAuto = GetObject("SAPGUI")
                    Set SAPApp = SapGuiAuto.GetScriptingEngine
                    iErr = Err.Number
                    Set Connection = SAPApp.Children(0)
                    Set session = Connection.Children(0)
                    session.createsession
                    Set session2 = Connection.Children(1)
                    k = k + 1
                    If (k > 10) Then Stop 'SAP closed?
                    If iErr <> 70 Then
                        SAPConnected = True
                    Else
                        Stop
                    End If
                Wend
                On Error GoTo ErrChk
            End If 'SAP conect is exit
    
    On Error Resume Next
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "operation" 'example /nZSE16N
    If (Err.Number <> 0) Then GoTo RetrySAPConn
    On Error GoTo ErrChk
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 3
    
    '========= Next Process
    SAPConnected = False
    
       
    
    
    Exit Function
    
ErrChk:
        
        MsgBox Err.Description
        Stop
        Resume

End Function
