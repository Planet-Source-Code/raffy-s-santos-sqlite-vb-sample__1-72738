VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H00404000&
   Caption         =   " Sample VB SQLite Program"
   ClientHeight    =   4170
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7110
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   3855
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6747
            Text            =   " Sample program by raffysantos@hotmail.com Dec. 2009"
            TextSave        =   " Sample program by raffysantos@hotmail.com Dec. 2009"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3825
      ScaleWidth      =   1785
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton btnSQLite 
         Caption         =   "Data Grid View"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   120
         Picture         =   "MDIForm1.frx":2372
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   1545
      End
      Begin VB.CommandButton btnUsers 
         Caption         =   "User Manager"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   120
         Picture         =   "MDIForm1.frx":27E8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   195
         Width           =   1545
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   600
         Top             =   2895
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   120
         Picture         =   "MDIForm1.frx":2CAE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1950
         Width           =   1545
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
    Me.Move 0, 0, 13000, 10000
    Screen.MousePointer = vbNormal
End Sub
Private Sub btnUsers_Click()
    frmUserManager.Show
    frmUserManager.ZOrder 0
End Sub
Private Sub btnSQLite_Click()
    frmSQLite.Show
End Sub
Private Sub btnClose_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ShutDown
End Sub
Private Sub Timer1_Timer()
    sbar.Panels(2).Text = " Username : " & username & " "
    sbar.Panels(3).Text = " Access Level : " & AccessLevel & " "
    sbar.Panels(4).Text = " " & Format$(Date, "Long Date") & " "
    sbar.Panels(5).Text = " " & Format$(Now, "hh:mm:ss AMPM") & " "
End Sub

