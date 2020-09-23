VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUserManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " User Manager "
   ClientHeight    =   6765
   ClientLeft      =   5265
   ClientTop       =   2880
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8565
   Begin VB.CommandButton cmdDeletePic 
      Caption         =   "Delete Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6660
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3705
      Width           =   1560
   End
   Begin VB.CommandButton cmdChangePic 
      Caption         =   "Change Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5115
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1500
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3705
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   510
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   6535
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483630
      BackColor       =   -2147483639
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Username"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5940
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5940
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   180
      ScaleHeight     =   2280
      ScaleWidth      =   4725
      TabIndex        =   4
      Top             =   4350
      Width           =   4725
      Begin VB.CommandButton cmdShowPassword 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2550
         TabIndex        =   12
         Top             =   990
         Width           =   315
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1140
         MaxLength       =   8
         PasswordChar    =   "Ÿ"
         TabIndex        =   11
         Text            =   "abcde"
         Top             =   975
         Width           =   1350
      End
      Begin VB.TextBox txtUsername 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   10
         Top             =   555
         Width           =   3375
      End
      Begin VB.TextBox txtFullName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   9
         Top             =   135
         Width           =   3375
      End
      Begin VB.ComboBox cboAL 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   990
         Width           =   810
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1050
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "  Delete the current journal transaction.  "
         Top             =   1590
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1590
         UseMaskColor    =   -1  'True
         Width           =   930
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1935
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1590
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   17
         Top             =   1020
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Top             =   195
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   3090
         TabIndex        =   14
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label lblUID 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4245
         TabIndex        =   13
         Top             =   615
         Width           =   90
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4905
      Top             =   5955
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "This sample program uses dhSQLite DLL ver. 1.02 database engine wrapper from Datenhaus (Germany)."
      Height          =   585
      Left            =   5100
      TabIndex        =   21
      Top             =   4470
      Width           =   3045
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FreeImage3 VB Wrapper by Carsten Klein used for reliable rendering of images from SQLite BLOB data."
      Height          =   585
      Left            =   5100
      TabIndex        =   20
      Top             =   5160
      Width           =   3045
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   3060
      Left            =   5100
      Stretch         =   -1  'True
      Top             =   510
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER ACCOUNT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   195
      TabIndex        =   3
      Top             =   180
      Width           =   1245
   End
End
Attribute VB_Name = "frmUserManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim strSQL As String
    Dim currID As Long
    Dim blnAddMode As Boolean
Sub Form_Load()
    Dim i As Integer
    Me.Move 500, 500, 8655, 7245
    cboAL.Clear
    For i = 0 To 9
        cboAL.AddItem i
    Next i
    If AccessLevel = 9 Then
        Picture1.Visible = True
        lv.Height = 3720
    Else
        Picture1.Visible = False
        lv.Height = 5235
    End If
    cmdRefresh_Click
    Screen.MousePointer = vbNormal
End Sub
Private Sub cmdRefresh_Click()
    blnAddMode = False
    Dim thisID As Long
    If lv.ListItems.Count > 0 Then
        thisID = CLng(lv.SelectedItem.Index)
    Else
        thisID = 1
    End If
    lv.ListItems.Clear
    LoadData
    If lv.ListItems.Count = 0 Then Exit Sub
    If thisID > 0 Then
        lv.ListItems(1).Selected = True
        currID = CLng(lv.SelectedItem)
        lv_ItemClick lv.SelectedItem
    End If
End Sub
Private Sub LoadData()
    Dim i As Integer
    blnAddMode = False
    lv.ListItems.Clear
    lv.ColumnHeaders(2).Width = 2900
    Dim rsLogin As cRecordset
    If AccessLevel = 9 Then
        strSQL = "SELECT * FROM [LOGINS] "
    Else
        strSQL = "SELECT * FROM [LOGINS] WHERE " & _
             "[USERNAME] <> 'admin' AND [ACCESS] <= " & AccessLevel
    End If
    strSQL = strSQL & " ORDER BY [USERNAME]"
    Set rsLogin = cnn.OpenRecordset(strSQL)
    If Not rsLogin.BOF And Not rsLogin.EOF Then
        i = 1
        Do While Not rsLogin.EOF
            lv.ListItems.Add i, , Format$(rsLogin!uID, "000")
            If Not IsNull(rsLogin!FullName) Then
                lv.ListItems(i).SubItems(1) = rsLogin!FullName
            Else
                lv.ListItems(i).SubItems(1) = "-"
            End If
            lv.ListItems(i).SubItems(2) = rsLogin!username
            i = i + 1
            rsLogin.MoveNext
        Loop
    End If
    Set rsLogin = Nothing
End Sub
Private Sub cboFilter_Click()
    cmdRefresh_Click
End Sub
Private Sub DisplayData()
    cmdAdd.Visible = True
    cmdDelete.Visible = True
    blnAddMode = False
    lblUID.Caption = 0
    txtFullName.Text = ""
    txtUserName.Text = ""
    txtPassword.Text = ""
    Dim rsLogin As cRecordset
    strSQL = "SELECT * FROM [LOGINS] WHERE [UID] = " & currID
    Set rsLogin = cnn.OpenRecordset(strSQL)
    If Not rsLogin.BOF And Not rsLogin.EOF Then
        If Not IsNull(rsLogin.Fields("UID").Value) Then
            lblUID.Caption = rsLogin.Fields("UID").Value
        End If
        If Not IsNull(rsLogin.Fields("FULLNAME").Value) Then
            txtFullName.Text = rsLogin.Fields("FULLNAME").Value
        End If
        If Not IsNull(rsLogin.Fields("USERNAME").Value) Then
            txtUserName.Text = rsLogin.Fields("USERNAME").Value
        End If
        If Not IsNull(rsLogin.Fields("PASSWORD").Value) Then
            txtPassword.Text = rsLogin.Fields("PASSWORD").Value
        End If
        cboAL.ListIndex = rsLogin.Fields("ACCESS").Value
    End If
    Set rsLogin = Nothing
    If txtUserName.Text = "admin" Then
        txtFullName.Enabled = False
        txtUserName.Enabled = False
        cboAL.Enabled = False
    Else
        txtFullName.Enabled = True
        txtUserName.Enabled = True
        cboAL.Enabled = True
    End If
    Screen.MousePointer = vbNormal
End Sub
Private Sub cmdUpdate_Click()
    If Trim$(txtUserName.Text) = "" Then
        MsgBox " Cannot add with an empty username !    ", _
                vbCritical, " Add User failed"
        txtUserName.SetFocus
        Exit Sub
    End If
    If Trim$(txtPassword.Text) = "" Or txtPassword.Text = "" Then
        MsgBox " Cannot add with an empty password !    ", _
                vbCritical, " Add User failed"
        txtPassword.SetFocus
        Exit Sub
    End If
    Dim Rs As cRecordset
    If blnAddMode Then
        Set Rs = cnn.OpenRecordset("SELECT * FROM [LOGINS] WHERE [USERNAME] = '" & _
                    txtUserName.Text & "' AND [USERNAME] <> 'admin'")
        If Not Rs.BOF And Not Rs.EOF Then
            MsgBox " That user already exists !         " & _
                "Please try a different username.   ", _
                vbCritical, " Add User Failed"
            txtUserName.SetFocus
            txtUserName.SelStart = 0
            txtUserName.SelLength = Len(txtUserName.Text)
            Set Rs = Nothing
            Exit Sub
        End If
        Set Rs = Nothing
        Dim newID As Long
        newID = GetNewID("LOGINS", "UID")
        cnn.Execute "INSERT INTO [LOGINS] VALUES (" & _
                newID & ", '" & _
                txtFullName.Text & "', '" & _
                txtUserName.Text & "', '" & _
                txtPassword.Text & "', " & _
                CInt(cboAL.Text) & ")"
        cmdRefresh_Click
        blnAddMode = False
        Exit Sub
    Else
        cnn.Execute "UPDATE [LOGINS] SET " & _
                "[FULLNAME] = '" & txtFullName.Text & "', " & _
                "USERNAME = '" & txtUserName.Text & "', " & _
                "[PASSWORD] = '" & txtPassword.Text & "', " & _
                "[ACCESS] = " & CInt(cboAL.Text) & _
                " WHERE [UID] = " & CLng(lblUID.Caption)
        If txtUserName.Text = username Then
            AccessLevel = CInt(cboAL.Text)
        End If
    End If
    Set Rs = Nothing
    cmdRefresh_Click
End Sub
Private Sub cmdAdd_Click()
    blnAddMode = True
    lv.ListItems(lv.ListItems.Count).Selected = True
    cmdAdd.Visible = False
    cmdDelete.Visible = False
    blnAddMode = True
    txtUserName.Text = ""
    txtPassword.Text = ""
    txtFullName.Text = ""
    txtFullName.Enabled = True
    txtUserName.Enabled = True
    txtPassword.Enabled = True
    cboAL.Enabled = True
    cboAL.ListIndex = 1
    txtFullName.SetFocus
End Sub
Private Sub cmdDelete_Click()
    If txtUserName.Text = "admin" Then
        MsgBox "You cannot delete the 'admin' account !         ", _
                    vbExclamation, "Access denied"
        Exit Sub
    End If
    Dim resp As Long
    resp = MsgBox("Are you sure you to delete the account of " & _
            txtUserName.Text & " ?    ", _
            vbQuestion + vbYesNo, "Warning: Confirm delete user")
    If resp = vbYes Then
        cnn.Execute "DELETE FROM LOGINS WHERE USERNAME = '" & _
                txtUserName.Text & "'"
        cmdRefresh_Click
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = 96
        Case 13
          SendKeys "{TAB}"
          KeyAscii = 0
    End Select
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lv.ListItems.Count = 0 Then Exit Sub
    currID = CLng(lv.SelectedItem)
    DisplayData
    ShowPic
End Sub
Private Sub txtFullName_GotFocus()
    txtFullName.SelStart = 0
    txtFullName.SelLength = Len(txtFullName.Text)
End Sub
Private Sub txtUsername_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
Private Sub cmdShowPassword_Click()
    If Trim$(cmdShowPassword.Caption) = "?" Then
        txtPassword.Font = "Tahoma"
        txtPassword.PasswordChar = ""
        cmdShowPassword.Caption = "x"
    Else
        txtPassword.Font = "Fixedsys"
        txtPassword.PasswordChar = "Ÿ"
        cmdShowPassword.Caption = "?"
    End If
    On Error Resume Next
    lv.SetFocus
End Sub
Private Sub ShowPic()
    Image1.Picture = LoadPicture()
    If lv.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim file_name As String
    Dim file_length As Long
    Dim fnum As Integer
    Dim Bytes() As Byte
    Dim rsPic As cRecordset
    file_name = App.Path & "\tmp.jpg"
    If PathFileExists(file_name) <> 0 Then
        Kill file_name
    End If
    Set rsPic = cnn.OpenRecordset("select [PIC], [PIC_LEN] from [PICTURES] " & _
                        "where [ID_NO] = " & CLng(lv.SelectedItem))
    If Not rsPic.BOF And Not rsPic.EOF Then
        If Not IsNull(rsPic.Fields(0).Value) Then
              If rsPic.Fields(0).ActualSize = 0 Then
                Set rsPic = Nothing
                Exit Sub
              End If
              Bytes = rsPic.Fields(0).Value
              fnum = FreeFile
              Open file_name For Binary As #fnum
                  Put #fnum, 1, Bytes
              Close fnum
              Dim dib As Long
              dib = FreeImage_Load(FIF_JPEG, file_name, 0)
              Set Image1.Picture = LoadPictureEx(file_name)
              Image1.Refresh
        End If
    End If
    Set rsPic = Nothing
End Sub
Private Sub cmdChangePic_Click()
    If lv.ListItems.Count = 0 Then
        Exit Sub
    End If
    'On Error GoTo errhand
    Dim Rs As cRecordset
    dlg.DialogTitle = "Select Picture"
    dlg.Flags = _
        cdlOFNFileMustExist Or _
        cdlOFNHideReadOnly Or _
        cdlOFNExplorer
    dlg.CancelError = True
    dlg.InitDir = App.Path & "\pictures"
    dlg.Filter = "Picture Files | *.jpg|GIF | *.gif|Bitmap | *.bmp"
    dlg.CancelError = False
    dlg.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & Err.Description
        Exit Sub
    End If
    If dlg.Filename = "" Then
        Exit Sub
    End If
    Dim f_name As String
    f_name = dlg.Filename

    Set Rs = cnn.OpenRecordset("select * from [PICTURES] where " & _
                               "[ID_NO] = " & CLng(lv.SelectedItem))
    If Rs.BOF Or Rs.EOF Then
        cnn.Execute "insert into [PICTURES] values(" & _
                    CLng(lv.SelectedItem) & ", 0, NULL)"
    End If
    Set Rs = Nothing
    Set Rs = cnn.OpenRecordset("select * from [PICTURES] where " & _
                               "[ID_NO] = " & CLng(lv.SelectedItem))
    setBLOB Rs, "PIC", f_name
    Set Rs = Nothing
    ShowPic
Exit Sub
errhand:
    MsgBox Err.Description
End Sub
Private Sub setBLOB(Rs As cRecordset, Field As String, Source As String)
'Usage: setBLOB myRecordSet, "FileField", "c:\myfile.gif" ' Places file into database
    Dim byBlobData() As Byte
    Dim intFileHandle As Integer
    Dim lngFileLength As Long
    intFileHandle = FreeFile
    Open Source For Binary As intFileHandle
        lngFileLength = LOF(intFileHandle) - 1
        byBlobData = InputB(lngFileLength, intFileHandle)
    Close intFileHandle
    Dim Cmd As cCommand
    Set Cmd = cnn.CreateCommand("update [PICTURES] set " & _
                                "[PIC_LEN] = " & lngFileLength & ", " & _
                                "[PIC] = (?) where " & _
                                "[ID_NO] = " & CLng(lv.SelectedItem))
    On Error Resume Next
    cnn.BeginTrans
    Cmd.setBLOB 1, byBlobData    'second Param to BlobBytes
    Cmd.Execute 'execute the insert-command
    If Err.Number = 0 Then 'success
        cnn.CommitTrans
    Else
        cnn.RollbackTrans
    End If
End Sub
Private Sub cmdDeletePic_Click()
    If lv.ListItems.Count = 0 Then Exit Sub
    cnn.Execute "delete from [PICTURES] where [ID_NO] = " & CLng(lv.SelectedItem)
    ShowPic
End Sub


