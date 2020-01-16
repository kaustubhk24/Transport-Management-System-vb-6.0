VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Salary Details"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   LinkTopic       =   "Form5"
   ScaleHeight     =   8595
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      Caption         =   "NEXT"
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
      Left            =   8280
      TabIndex        =   41
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   " PREVIOUS"
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
      Left            =   8280
      TabIndex        =   40
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Height          =   3015
      Left            =   7680
      TabIndex        =   32
      Top             =   1560
      Width           =   2055
      Begin VB.CommandButton Command8 
         Caption         =   " MOVE LAST"
         Height          =   615
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   " MOVE NEXT"
         Height          =   615
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   " MOVE PREVIOUS"
         Height          =   615
         Left            =   240
         TabIndex        =   34
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   " MOVE FIRST"
         Height          =   615
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   1560
      TabIndex        =   17
      Top             =   6480
      Width           =   4815
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   4575
         Begin VB.Frame Frame4 
            Height          =   1095
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   4335
            Begin VB.CommandButton Command4 
               Caption         =   "&CLOSE"
               Height          =   615
               Left            =   3120
               TabIndex        =   23
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command3 
               Caption         =   "&EDIT"
               Height          =   615
               Left            =   2160
               TabIndex        =   22
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command2 
               Caption         =   "&DELETE"
               Height          =   615
               Left            =   1200
               TabIndex        =   21
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command1 
               Caption         =   "&ADD"
               Height          =   615
               Left            =   240
               TabIndex        =   20
               Top             =   240
               Width           =   975
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   6495
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   2520
         TabIndex        =   31
         Text            =   " "
         Top             =   4560
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   4680
         TabIndex        =   29
         Text            =   " "
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   4680
         TabIndex        =   27
         Text            =   " "
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Text            =   " "
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Text            =   " "
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Text            =   " "
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Text            =   " "
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   " "
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4680
         TabIndex        =   3
         Text            =   " "
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   4680
         TabIndex        =   2
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Text            =   " "
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Total Leaves"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   " Net Salary"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Extra Leaves"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Medical Leaves"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Casual Leaves"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "       HRA"
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "        DA"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Basic Salary"
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Driver Name"
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Driver Identity Number"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "       CCA"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "        PF"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   3360
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Height          =   3255
      Left            =   7560
      TabIndex        =   38
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Frame Frame7 
      Height          =   5415
      Left            =   360
      TabIndex        =   39
      Top             =   840
      Width           =   6735
   End
   Begin VB.Frame Frame8 
      Height          =   1335
      Left            =   8160
      TabIndex        =   42
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "  SALARY DETAILS"
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
      TabIndex        =   37
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DB1 As Database
Public RS1 As Recordset
Public RS2 As Recordset
Public RS3 As Recordset
Public RS4 As Recordset

Public Sub CLEARF()
Combo1.Text = " "
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text8.Text = " "
Text9.Text = " "
Text10.Text = " "
End Sub
Public Sub DISPLAY()
Combo1.Text = RS1.Fields(0)
Text1.Text = RS1.Fields(1)
Text2.Text = RS1.Fields(2)
Text3.Text = RS1.Fields(3)
Text4.Text = RS1.Fields(4)
Text5.Text = RS1.Fields(5)

Text6.Text = RS1.Fields(6)
Text7.Text = RS1.Fields(7)
Text8.Text = RS1.Fields(8)
Text9.Text = RS1.Fields(9)
Text10.Text = RS1.Fields(10)
Text11.Text = RS1.Fields(11)

End Sub


Private Sub Command10_Click()
Form7.Show
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
Command2.Enabled = False
Command3.Enabled = False
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
Command2.Enabled = False
Command3.Enabled = False
RS1.MoveLast
DISPLAY
End Sub


Private Sub Command9_Click()
Form4.Show
End Sub

Private Sub Form_Activate()
Text3.Text = 2
Text3.Enabled = False
Text5.Text = 0
End Sub

Private Sub Form_Load()
Set DB1 = DBEngine.OpenDatabase(App.Path & "\transport.mdb")
Set RS1 = DB1.OpenRecordset("Salary")
Command2.Enabled = False
Command3.Enabled = False
COMBOLIST
CLEARF
End Sub



Private Sub Command1_Click()
If Command1.Caption = "&ADD" Then
CLEARF
RS1.AddNew
Command2.Enabled = False
Command3.Enabled = False
Command1.Caption = "&UPDATE"
Text3.Text = 2
Text4.Text = 0
Text5.Text = 0
Text6.Text = 0
Combo1.SetFocus
Else
RS1.Fields(0) = Combo1.Text
 RS1.Fields(1) = Text1.Text
 RS1.Fields(2) = Text2.Text
RS1.Fields(3) = Text3.Text
 RS1.Fields(4) = Text4.Text
 RS1.Fields(5) = Text5.Text
 RS1.Fields(6) = Text6.Text
 RS1.Fields(7) = Text7.Text
 RS1.Fields(8) = Text8.Text
 RS1.Fields(9) = Text9.Text
 RS1.Fields(10) = Text10.Text
 RS1.Fields(11) = Text11.Text
If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS1.Update
End If
Command1.Caption = "&ADD"
End If
COMBOLIST
End Sub

Private Sub Command2_Click()
If MsgBox("Record Deleted", vbQuestion + vbYesNo, "DELETE") = vbYes Then
RS2.Delete
RS2.MoveFirst
DB1.Execute ("Delete From Salary where did='" & (Combo1.Text) & "'")
CLEARF
Command2.Enabled = False
End If
COMBOLIST
End Sub

Private Sub Command3_Click()
If Command3.Caption = "&EDIT" Then
RS2.Edit
Command3.Caption = "&SAVE"
'Command2.Enabled = True
Else
 RS2.Fields(0) = Combo1.Text
 RS2.Fields(1) = Text1.Text
 RS2.Fields(2) = Text2.Text
RS2.Fields(3) = Text3.Text
 RS2.Fields(4) = Text4.Text
 RS2.Fields(5) = Text5.Text
 RS2.Fields(6) = Text6.Text
 RS2.Fields(7) = Text7.Text
 RS2.Fields(8) = Text8.Text
 RS2.Fields(9) = Text9.Text
 RS2.Fields(10) = Text10.Text
 RS2.Fields(11) = Text11.Text

If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS2.Update
End If
Command3.Caption = "&EDIT"
COMBOLIST
End If
End Sub

Private Sub Combo1_Click()
Set RS2 = DB1.OpenRecordset("select * from Salary Where did='" _
& (Combo1.Text) & "'")
If RS2.RecordCount = 1 Then
Combo1.Text = RS2.Fields(0)
Text1.Text = RS2.Fields(1)
Text2.Text = RS2.Fields(2)
Text3.Text = RS2.Fields(3)
Text4.Text = RS2.Fields(4)
Text5.Text = RS2.Fields(5)
Text6.Text = RS2.Fields(6)
Text7.Text = RS2.Fields(7)
Text8.Text = RS2.Fields(8)
Text9.Text = RS2.Fields(9)
Text10.Text = RS2.Fields(10)
Text11.Text = RS2.Fields(11)
Command2.Enabled = True
Command3.Enabled = True
Else
Set RS4 = DB1.OpenRecordset("select * from Driver Where did='" _
& (Combo1.Text) & "'")
If RS4.RecordCount = 1 Then
Text1.Text = RS4.Fields(1)
MsgBox "Enter  Data"
End If
End If
End Sub

Public Sub COMBOLIST()
Combo1.Clear
Set RS3 = DB1.OpenRecordset("select * from Driver")
If Not (RS3.BOF And RS3.EOF) Then
RS3.MoveFirst
While Not RS3.EOF
Combo1.AddItem RS3.Fields(0), Index
RS3.MoveNext
Wend
End If
End Sub

Private Sub Command4_Click()
Unload Me
Form6.Show
End Sub



Private Sub Text2_LostFocus()
Dim s, t
s = Val(Text2.Text) / 30
t = s * 30 - Val(Text6.Text)
Text7.Text = (t * 0.3)
Text8.Text = (t * 0.3)
Text9.Text = (t * 0.15)
Text10.Text = (t * 0.3)
Text11.Text = Val(Text2.Text) + Val(Text7.Text) + Val(Text8.Text) _
+ Val(Text9.Text) - Val(Text10.Text)
End Sub

Private Sub Text5_LostFocus()
Dim N
N = Val(Text3.Text) - Val(Text5.Text)
N = N * -1
If N > 0 Then
Text6.Text = N
ElseIf N <= 0 Then
Text6.Text = 0
End If


End Sub

