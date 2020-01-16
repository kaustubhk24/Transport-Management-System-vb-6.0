VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Route Details"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   LinkTopic       =   "Form3"
   ScaleHeight     =   7755
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      Caption         =   " NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   27
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   26
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Height          =   3015
      Left            =   5640
      TabIndex        =   18
      Top             =   1440
      Width           =   2055
      Begin VB.CommandButton Command8 
         Caption         =   " MOVE LAST"
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "MOVE NEXT"
         Height          =   615
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   " MOVE PREVIOUS"
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   " MOVE FIRST"
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   3735
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Text            =   " "
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Text            =   " "
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Text            =   " "
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Text            =   " "
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Text            =   " "
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Route Number"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Route Name"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Bus Registration Number"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Driver Name"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Starting Time"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   2400
      TabIndex        =   0
      Top             =   5520
      Width           =   5535
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5295
         Begin VB.Frame Frame4 
            Height          =   975
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   5055
            Begin VB.CommandButton Command1 
               Caption         =   "&ADD"
               Height          =   615
               Left            =   600
               TabIndex        =   6
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command2 
               Caption         =   "&DELETE"
               Height          =   615
               Left            =   1560
               TabIndex        =   5
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command3 
               Caption         =   "&EDIT"
               Height          =   615
               Left            =   2520
               TabIndex        =   4
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command4 
               Caption         =   "&CLOSE"
               Height          =   615
               Left            =   3480
               TabIndex        =   3
               Top             =   240
               Width           =   975
            End
         End
      End
   End
   Begin VB.Frame Frame6 
      Height          =   3255
      Left            =   5520
      TabIndex        =   24
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Frame Frame7 
      Height          =   4095
      Left            =   960
      TabIndex        =   25
      Top             =   960
      Width           =   3975
   End
   Begin VB.Frame Frame8 
      Height          =   1335
      Left            =   8280
      TabIndex        =   28
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   " ROUTE DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   23
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public DB1 As Database
Public RS1 As Recordset
Public RS2 As Recordset
Public RS3 As Recordset
Public RS4 As Recordset
Public RS5 As Recordset
Public RS6 As Recordset

Public Sub CLEARF()
Combo1.Text = " "
Text1.Text = " "
Text2.Text = " "
Combo2.Text = " "
Combo3.Text = " "
End Sub


Private Sub Command10_Click()
Form4.Show
End Sub

Private Sub Command5_Click()
Command2.Enabled = False
Command3.Enabled = False
RS1.MoveFirst
DISPLAY
End Sub

Private Sub Command6_Click()
Command2.Enabled = False
Command3.Enabled = False
If Not RS1.BOF Then
RS1.MovePrevious
If Not RS1.BOF Then
DISPLAY
Else
RS1.MoveFirst
End If
End If
End Sub

Private Sub Command7_Click()
'Command6.Enabled = False
'Command7.Enabled = False
If Not RS1.EOF Then
RS1.MoveNext
If Not RS1.EOF Then
DISPLAY
Else
RS1.MoveLast
End If
End If
End Sub

Private Sub Command8_Click()
'Command2.Enabled = False
'Command3.Enabled = False
RS1.MoveLast
DISPLAY
End Sub


Private Sub Command9_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Set DB1 = DBEngine.OpenDatabase(App.Path & "\transport.mdb")
Set RS1 = DB1.OpenRecordset("Route")
Command2.Enabled = False
Command3.Enabled = False
COMBOLIST
COMBO2LIST
COMBO3LIST
CLEARF
End Sub
Public Sub DISPLAY()
Combo1.Text = RS1.Fields(0)
Text1.Text = RS1.Fields(1)
Combo2.Text = RS1.Fields(2)
Combo3.Text = RS1.Fields(3)
Text2.Text = RS1.Fields(4)
End Sub



Private Sub Command1_Click()
Dim MDATE
If Command1.Caption = "&ADD" Then
CLEARF
RS1.AddNew
Set RS6 = DB1.OpenRecordset("Select max(rno) from Route")
If RS6.RecordCount = 1 Then
MsgBox RS6.Fields(0)
Combo1.Text = RS6.Fields(0) + 1
'Combo1.Enabled = False
End If
Command2.Enabled = False
Command3.Enabled = False
Command1.Caption = "&UPDATE"
Else
RS1.Fields(0) = Combo1.Text
 RS1.Fields(1) = Text1.Text
 RS1.Fields(2) = Combo2.Text
RS1.Fields(3) = Combo3.Text
 RS1.Fields(4) = Text2.Text
If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS1.Update
End If
Command1.Caption = "&ADD"
COMBOLIST
End If
End Sub

Private Sub Command2_Click()
If MsgBox("Record Deleted", vbQuestion + vbYesNo, "DELETE") = vbYes Then
RS2.Delete
RS2.MoveFirst
DB1.Execute ("Delete From Route where rno=" & (Combo1.Text))
CLEARF
Command2.Enabled = False
COMBOLIST
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "&EDIT" Then
RS2.Edit
Command3.Caption = "&SAVE"
'Command2.Enabled = True
Else
 RS2.Fields(0) = Combo1.Text
 RS2.Fields(1) = Text1.Text
 RS2.Fields(2) = Combo2.Text
RS2.Fields(3) = Combo3.Text
 RS2.Fields(4) = Text2.Text
 If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS2.Update
End If
Command3.Caption = "&EDIT"
COMBOLIST
End If
End Sub

Private Sub Combo1_Click()
Set RS2 = DB1.OpenRecordset("select * from Route Where rno=" _
& (Combo1.Text))
If RS2.RecordCount = 1 Then
Combo1.Text = RS2.Fields(0)
Text1.Text = RS2.Fields(1)
Combo2.Text = RS2.Fields(2)
Combo3.Text = RS2.Fields(3)
Text2.Text = RS2.Fields(4)
Command2.Enabled = True
Command3.Enabled = True
Else
MsgBox "Enter  Data"
End If
End Sub

Public Sub COMBOLIST()
Combo1.Clear
Set RS3 = DB1.OpenRecordset("select * from Route")
If Not (RS3.BOF And RS3.EOF) Then
RS3.MoveFirst
While Not RS3.EOF
Combo1.AddItem RS3.Fields(0), Index
RS3.MoveNext
Wend
End If
End Sub

Public Sub COMBO2LIST()
Combo2.Clear
Set RS4 = DB1.OpenRecordset("select * from Bus")
If Not (RS4.BOF And RS4.EOF) Then
RS4.MoveFirst
While Not RS4.EOF
Combo2.AddItem RS4.Fields(1), Index
RS4.MoveNext
Wend
End If
End Sub

Public Sub COMBO3LIST()
Combo3.Clear
Set RS5 = DB1.OpenRecordset("select * from Driver")
If Not (RS5.BOF And RS5.EOF) Then
RS5.MoveFirst
While Not RS5.EOF
Combo3.AddItem RS5.Fields(1), Index
RS5.MoveNext
Wend
End If
End Sub
Private Sub Command4_Click()
Unload Me
Form6.Show
End Sub





