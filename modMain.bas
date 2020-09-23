Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias _
                 "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
                "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
                ByVal nSize As Long) As Long
Global cnn As cConnection
Global DBName As String
Global username As String
Global AccessLevel As Integer
Sub Main()
   'Prevent multiple instances of this application ..
    If App.PrevInstance = True Then End
   'Just to make sure that FreeImage.dll is present in the system32 folder...
    Dim strsrc As String, strdest As String
    Dim Path As String, strSave As String
    strSave = String(200, Chr$(0))
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\system32"
    If PathFileExists(Path & "\FreeImage.dll") = 0 Then
        strsrc = App.Path & "\FreeImage.dll"
        If PathFileExists(strsrc) <> 0 Then
            strdest = Path & "\FreeImage.dll"
            FileCopy strsrc, strdest
        Else
            End
        End If
    End If
   '---------------------------------
    DBName = App.Path & "\sample.db"
   '---------------------------------
    If PathFileExists(DBName) = 0 Then
        MsgBox "The SQLite Database '" & DBName & "' does not exist !     ", _
                        vbCritical, "Database not found"
        ShutDown
        Exit Sub
    End If
    
    Set cnn = New cConnection
    cnn.OpenDB DBName
    Dim fLogin As New frmLogin
    fLogin.Show vbModal
    If Not fLogin.OK Then
        End
    Else
        MDIForm1.Show
    End If
End Sub
Public Function GetNewID(ByVal sTbl As String, ByVal sFld As String) As Long
    Dim Rs As cRecordset
    Set Rs = cnn.OpenRecordset("select [" & sFld & "] from [" & sTbl & "] order by [" & sFld & "] desc")
    If Not Rs.BOF And Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0).Value) Then
            GetNewID = Rs.Fields(0).Value + 1
        Else
            GetNewID = 1
        End If
    Else
        GetNewID = 1
    End If
    Set Rs = Nothing
End Function
Public Sub ShutDown(Optional ByVal Force As Boolean = False)
    On Error Resume Next
    Set cnn = Nothing
    Dim i As Integer
    For i = Forms.Count - 1 To 0 Step -1
         Unload Forms(i) ' Triggers QueryUnload and Form_Unload
         If Not Force Then
            If Forms.Count > i Then
                If Forms.Count = 0 Then Exit For
            End If
         End If
    Next i
    If Force Or (Forms.Count = 0) Then Close
End Sub

