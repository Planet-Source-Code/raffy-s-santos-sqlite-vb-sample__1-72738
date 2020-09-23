VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSQLite 
   Caption         =   "Data Grid"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   Begin VB.PictureBox pDGBox 
      Height          =   4965
      Left            =   90
      ScaleHeight     =   327
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   573
      TabIndex        =   1
      Top             =   2160
      Width           =   8655
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1365
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   2408
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1031
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1031
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.ListBox lstTables 
      Height          =   1425
      Left            =   630
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   2175
   End
   Begin VB.Label lTables 
      Caption         =   "Tables"
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   165
      Width           =   645
   End
   Begin VB.Label lTiming 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   1710
      Width           =   8655
   End
End
Attribute VB_Name = "frmSQLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As cRecordset
Private Sub Form_Load()
    Dim Table As cTable
    For Each Table In cnn.DataBases(1).Tables
      If Left$(Table.Name, 7) <> "sqlite_" Then lstTables.AddItem Table.Name
    Next Table
    If lstTables.ListCount > 0 Then
        lstTables.ListIndex = 0
        lstTables_Click
    End If
End Sub
Private Sub lstTables_Click()
    If lstTables.ListIndex = -1 Then Exit Sub
    Screen.MousePointer = 13
    QueryAndVisualize lstTables.Text
    Screen.MousePointer = 0
End Sub
Private Sub QueryAndVisualize(TableOrViewName As String)
      Set Rs = cnn.OpenRecordset("Select * from [" & TableOrViewName & "]")
      Set DataGrid1.DataSource = Rs.DataSource
End Sub
Private Sub Form_Resize()
    pDGBox.Move -1, lTiming.Top + lTiming.Height, ScaleWidth + 2, ScaleHeight - lTiming.Top - lTiming.Height + 2
    lTiming.Move -3, lTiming.Top, ScaleWidth + 6
End Sub
Private Sub pDGBox_Resize()
    DataGrid1.Move 0, 0, pDGBox.ScaleWidth, pDGBox.ScaleHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set DataGrid1.DataSource = Nothing
    Set Rs = Nothing
End Sub

