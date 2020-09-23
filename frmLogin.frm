VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Q u i t  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   2850
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&E n t e r "
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "OK"
      Top             =   2835
      Width           =   1425
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2565
      PasswordChar    =   "Ÿ"
      TabIndex        =   1
      Text            =   "admin"
      Top             =   2175
      Width           =   1650
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2565
      TabIndex        =   0
      Text            =   "admin"
      Top             =   1710
      Width           =   1650
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your username and password to login."
      Height          =   255
      Left            =   1245
      TabIndex        =   6
      Top             =   1350
      Width           =   3330
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Tag             =   "&Password:"
      Top             =   2235
      Width           =   810
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   1425
      TabIndex        =   4
      Tag             =   "&User Name:"
      Top             =   1770
      Width           =   915
   End
   Begin VB.Image Image2 
      Height          =   1620
      Left            =   480
      Picture         =   "frmLogin.frx":0CB2
      Top             =   60
      Width           =   5130
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean
Private Sub Form_Load()
    Screen.MousePointer = vbNormal
    txtUserName.Text = GetSetting(App.EXEName, "User", "Login", "admin")
End Sub
Private Sub Form_Activate()
     On Error Resume Next
     txtPassword.SetFocus
End Sub
Private Sub cmdOK_Click()
    Screen.MousePointer = vbHourglass
    If txtUserName.Text = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Please enter your login name."
        Exit Sub
    End If
    If txtPassword.Text = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Please enter your password."
        Exit Sub
    End If
    username = txtUserName.Text
    Dim rsL As cRecordset
    Set rsL = cnn.OpenRecordset("SELECT * FROM [LOGINS] WHERE [USERNAME] = '" & _
        UCase$(txtUserName.Text) & "' OR [USERNAME] = '" & _
        LCase$(txtUserName.Text) & "'")
    If rsL.BOF And rsL.EOF Then
        Screen.MousePointer = vbNormal
        MsgBox "Username not found !"
        txtUserName.SetFocus
        txtUserName.SelStart = 0
        txtUserName.SelLength = Len(txtUserName.Text)
        Set rsL = Nothing
        Exit Sub
    Else
        If rsL.Fields("PASSWORD").Value = txtPassword.Text Then
           ' login is ok !
            DoEvents
            modMain.username = txtUserName.Text
            If IsNumeric(rsL.Fields("ACCESS").Value) Then
                AccessLevel = rsL.Fields("ACCESS").Value
            Else
                AccessLevel = 0
            End If
            OK = True
            SaveSetting App.EXEName, "User", "Login", txtUserName.Text
            Set rsL = Nothing
            Unload Me
        Else
            Screen.MousePointer = vbNormal
            MsgBox "Incorrect password !"
            txtPassword.SetFocus
            txtPassword.SelStart = 0
            txtPassword.SelLength = Len(txtPassword.Text)
            Set rsL = Nothing
            Exit Sub
        End If

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
Private Sub cmdQuit_Click()
    OK = False
    ShutDown
End Sub
Private Sub cmdCancel_Click()
    OK = False
    ShutDown
End Sub
Private Sub txtUsername_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub
