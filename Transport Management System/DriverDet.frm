VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Driver Details"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8040
   ScaleWidth      =   11040
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
      Height          =   555
      Left            =   8400
      TabIndex        =   35
      Top             =   3720
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
      TabIndex        =   34
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   5520
      TabIndex        =   26
      Top             =   1920
      Width           =   2055
      Begin VB.CommandButton Command5 
         Caption         =   " MOVE FIRST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   " MOVE PREVIOUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   " MOVE NEXT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   " MOVE LAST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   3855
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   25
         Text            =   " "
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Text            =   " "
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         Text            =   " "
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   22
         Text            =   " "
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Text            =   "  "
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Text            =   " "
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Text            =   " "
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   18
         Text            =   " "
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Text            =   " "
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Driving Licence Expired Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Driving Licence Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Date of Join"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Driver Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Driver Identity Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Driver Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Driver Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Blood Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      TabIndex        =   0
      Top             =   6600
      Width           =   5535
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5295
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   5055
            Begin VB.CommandButton Command1 
               Caption         =   "&ADD"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   600
               TabIndex        =   6
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command2 
               Caption         =   "&DELETE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   1560
               TabIndex        =   5
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command3 
               Caption         =   "&EDIT"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   2520
               TabIndex        =   4
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command4 
               Caption         =   "&CLOSE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
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
      Height          =   5535
      Left            =   360
      TabIndex        =   32
      Top             =   840
      Width           =   4095
   End
   Begin VB.Frame Frame7 
      Height          =   3375
      Left            =   5400
      TabIndex        =   33
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Frame Frame8 
      Height          =   1455
      Left            =   8280
      TabIndex        =   36
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   " DRIVER DETAILS"
      Height          =   615
      Left            =   3120
      TabIndex        =   31
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DB1 As Database
Public RS1 As Recordset
Public RS2 As Recordset
Public RS3 As Recordset

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
End Sub



Private Sub Command10_Click()
Form3.Show
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
'Command6.Enabled = False
'Command7.Enabled = False
RS1.MoveLast
DISPLAY
End Sub


Private Sub Command9_Click()
Form1.Show
End Sub

Private Sub Form_Load()
Set DB1 = DBEngine.OpenDatabase(App.Path & "\transport.mdb")
Set RS1 = DB1.OpenRecordset("Driver")
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
Text4.Text = RS1.Fields(4)
Text5.Text = RS1.Fields(5)
Text6.Text = RS1.Fields(6)
Text7.Text = RS1.Fields(7)
Text8.Text = RS1.Fields(8)
End Sub




Private Sub Command1_Click()
Dim MDATE
If Command1.Caption = "&ADD" Then
CLEARF
RS1.AddNew
Command2.Enabled = False
Command3.Enabled = False
Command1.Caption = "&UPDATE"
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
DB1.Execute ("Delete From Driver where did='" & (Combo1.Text) & "'")
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
If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS2.Update
End If
Command3.Caption = "&EDIT"
COMBOLIST
End If
End Sub

Private Sub Combo1_Click()
Set RS2 = DB1.OpenRecordset("select * from Driver Where did='" _
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
Command2.Enabled = True
Command3.Enabled = True
Else
MsgBox "Enter  Data"
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




