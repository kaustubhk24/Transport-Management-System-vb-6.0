VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Diesel Filling Details"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
   LinkTopic       =   "Form4"
   ScaleHeight     =   6885
   ScaleWidth      =   10200
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
      Left            =   8520
      TabIndex        =   25
      Top             =   3120
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
      Left            =   8520
      TabIndex        =   24
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Height          =   3135
      Left            =   5880
      TabIndex        =   14
      Top             =   1320
      Width           =   2055
      Begin VB.CommandButton Command8 
         Caption         =   " MOVE LAST"
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "MOVE NEXT"
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   " MOVE PREVIOUS"
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   " MOVE FIRST"
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   960
      TabIndex        =   7
      Top             =   1200
      Width           =   3735
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Text            =   " "
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1680
         TabIndex        =   13
         Text            =   " "
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1680
         TabIndex        =   12
         Text            =   " "
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Text            =   " "
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Current Speedo Meter Reading"
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Bus Registration Number"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "        Litres"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "          Date"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   1560
      TabIndex        =   0
      Top             =   5280
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
      Height          =   3375
      Left            =   5760
      TabIndex        =   22
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Frame Frame7 
      Height          =   3615
      Left            =   840
      TabIndex        =   23
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Frame Frame8 
      Height          =   1335
      Left            =   8400
      TabIndex        =   26
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "DIESEL DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   21
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form4"
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
Text3.Text = " "
End Sub

Private Sub Command10_Click()
Form5.Show
End Sub

Private Sub Command5_Click()
Command2.Enabled = True
Command3.Enabled = True
RS1.MoveFirst
DISPLAY
End Sub

Private Sub Command6_Click()
Command2.Enabled = True
Command3.Enabled = True
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
Command2.Enabled = True
Command3.Enabled = True
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
Command2.Enabled = True
Command3.Enabled = True
RS1.MoveLast
DISPLAY
End Sub


Private Sub Command9_Click()
Form3.Show
End Sub

Private Sub Form_Activate()
Text2.Text = Format$(Date$, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
Form4.Hide
Form9.Show
Set DB1 = DBEngine.OpenDatabase(App.Path & "\transport.mdb")
Set RS1 = DB1.OpenRecordset("Petro")
Command2.Enabled = False
Command3.Enabled = False
COMBOLIST
CLEARF
End Sub
Public Sub DISPLAY()
Combo1.Text = RS1.Fields(0)
Text1.Text = RS1.Fields(1)
Text2.Text = RS1.Fields(2)
Text3.Text = RS1.Fields(3)
End Sub



Private Sub Command1_Click()
Dim MDATE
If Command1.CAPTION = "&ADD" Then
Text1.Text = 0
Text3.Text = 0
RS1.AddNew
Command2.Enabled = False
Command3.Enabled = False
Command1.CAPTION = "&UPDATE"
Combo1.SetFocus
Else
 RS1.Fields(0) = Combo1.Text
 RS1.Fields(1) = Text1.Text
 RS1.Fields(2) = Text2.Text
 RS1.Fields(3) = Text3.Text
If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS1.Update
End If
Command1.CAPTION = "&ADD"
Command2.Enabled = True
Command3.Enabled = True
End If
COMBOLIST
End Sub

Private Sub Command2_Click()
If MsgBox("Record Deleted", vbQuestion + vbYesNo, "DELETE") = vbYes Then
RS1.Delete
RS1.MoveFirst
DB1.Execute ("Delete From Petro where breg='" & (Combo1.Text) & "'")
CLEARF
Command2.Enabled = False
End If
COMBOLIST
End Sub

Private Sub Command3_Click()

Set RS2 = DB1.OpenRecordset("select * from Petro Where breg='" _
& (Combo1.Text) & "'")
If RS2.RecordCount = 1 Then
RS2.Edit
 RS2.Fields(0) = Combo1.Text
 RS2.Fields(1) = RS2.Fields(1) + Val(Text1.Text)
 RS2.Fields(2) = Text2.Text
 N = Val(Text3.Text) - RS2.Fields(3)
 RS2.Fields(3) = RS2.Fields(3) + N
If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS2.Update
End If
Command3.CAPTION = "&EDIT"
COMBOLIST
End If
End Sub

Private Sub Combo1_Click()
Set RS2 = DB1.OpenRecordset("select * from Petro Where breg='" _
& (Combo1.Text) & "'")
If RS2.RecordCount = 1 Then
Command3.Enabled = True
Command1.Enabled = False
'Combo1.Text = RS2.Fields(0)
'Text1.Text = RS2.Fields(1)
'Text2.Text = RS2.Fields(2)
MsgBox "Previous Meter Reading :: " & RS2.Fields(3)
MsgBox "Diesel Utilized Upto Date :: " & RS2.Fields(1)
Text3.Text = ""
Text1.Text = ""
Text3.SetFocus
'Command2.Enabled = True
'Command3.Enabled = True
Else
MsgBox "Enter  Data"
Command3.Enabled = False
Command1.Enabled = True
End If
End Sub

Public Sub COMBOLIST()
Combo1.Clear
Set RS3 = DB1.OpenRecordset("select * from Bus")
If Not (RS3.BOF And RS3.EOF) Then
RS3.MoveFirst
While Not RS3.EOF
Combo1.AddItem RS3.Fields(1), Index
RS3.MoveNext
Wend
End If
End Sub

Private Sub Command4_Click()
Unload Me
Form6.Show
End Sub



Private Sub Text3_LostFocus()
Dim N
If Text3.Text = "" Then

ElseIf Text3.Text = 0 Then
Text1.Text = 50
End If
Set RS4 = DB1.OpenRecordset("select * from Petro Where breg='" _
& (Combo1.Text) & "'")
If RS4.RecordCount = 1 Then
N = Val(Text3.Text) - RS4.Fields(3)
If N >= 0 And N < 200 Then
Text1.Text = 50
ElseIf N >= 200 And N < 300 Then
Text1.Text = 100
ElseIf N >= 300 And N < 400 Then
Text1.Text = 130
ElseIf N >= 400 And N < 600 Then
Text1.Text = 150
ElseIf N >= 600 And N < 700 Then
Text1.Text = 170
ElseIf N >= 700 And N <= 1000 Then
Text1.Text = 250
End If
End If
End Sub
