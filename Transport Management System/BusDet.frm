VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "  Bus Details"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   10470
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
      Left            =   8160
      TabIndex        =   29
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
      Height          =   555
      Left            =   8160
      TabIndex        =   28
      Top             =   3120
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
      Height          =   2775
      Left            =   5280
      TabIndex        =   20
      Top             =   2160
      Width           =   1695
      Begin VB.CommandButton Command8 
         Caption         =   "MOVE LAST"
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
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   1455
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
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "MOVE PREVIOUS"
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
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1455
      End
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
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
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
      Left            =   1800
      TabIndex        =   8
      Top             =   5760
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
         TabIndex        =   9
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
            TabIndex        =   10
            Top             =   120
            Width           =   5055
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
               Left            =   3600
               TabIndex        =   14
               Top             =   240
               Width           =   1095
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
               TabIndex        =   13
               Top             =   240
               Width           =   1095
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
               Left            =   1440
               TabIndex        =   12
               Top             =   240
               Width           =   1095
            End
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
               Left            =   360
               TabIndex        =   11
               Top             =   240
               Width           =   1095
            End
         End
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
      Height          =   3855
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   3375
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
         Left            =   1560
         TabIndex        =   19
         Text            =   " "
         Top             =   960
         Width           =   1335
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
         Left            =   1560
         TabIndex        =   18
         Text            =   " "
         Top             =   3000
         Width           =   1335
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
         Height          =   405
         Left            =   1560
         TabIndex        =   17
         Text            =   " "
         Top             =   2520
         Width           =   1335
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
         Left            =   1560
         TabIndex        =   16
         Text            =   " "
         Top             =   1920
         Width           =   1335
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
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Text            =   " "
         Top             =   1440
         Width           =   1335
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
         Left            =   1560
         TabIndex        =   7
         Text            =   " "
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Registration expired Date"
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
         Left            =   240
         TabIndex        =   6
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Date of Registration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Bus Capacity"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Bus Chassi No"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Bus Registration Number"
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
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Bus No"
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Height          =   4335
      Left            =   480
      TabIndex        =   26
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Frame Frame7 
      Height          =   3255
      Left            =   5160
      TabIndex        =   27
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Frame Frame9 
      Height          =   1575
      Left            =   7920
      TabIndex        =   30
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   " BUS DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3120
      TabIndex        =   25
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DB1 As Database
Public RS1 As Recordset
Public RS2 As Recordset
Public RS3 As Recordset
Public RS4 As Recordset
Dim i

Public Sub CLEARF()
Combo1.Text = " "
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
End Sub
Public Sub DISPLAY()
Combo1.Text = RS1.Fields(0)
Text1.Text = RS1.Fields(1)
Text2.Text = RS1.Fields(2)
Text3.Text = RS1.Fields(3)
Text4.Text = RS1.Fields(4)
Text5.Text = RS1.Fields(5)
End Sub



Private Sub Command10_Click()
Form2.Show
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
Form6.Show
End Sub

Private Sub Form_Load()
Set DB1 = DBEngine.OpenDatabase(App.Path & "\transport.mdb")
Set RS1 = DB1.OpenRecordset("bus")
Command2.Enabled = False
Command3.Enabled = False
COMBOLIST
CLEARF
End Sub



Private Sub Command1_Click()
Dim MDATE
If Command1.Caption = "&ADD" Then
CLEARF
RS1.AddNew
Set RS4 = DB1.OpenRecordset("Select max(bno) from Bus")
If RS4.RecordCount = 1 Then
MsgBox RS4.Fields(0)
Combo1.Text = RS4.Fields(0) + 1
Combo1.Enabled = False
'MsgBox Combo1.Text
End If
Command2.Enabled = False
Command3.Enabled = False
Command1.Caption = "&UPDATE"
Text1.SetFocus
Else
RS1.Fields(0) = Combo1.Text
RS1.Fields(1) = Text1.Text
RS1.Fields(2) = Text2.Text
RS1.Fields(3) = Text3.Text
RS1.Fields(4) = Text4.Text
RS1.Fields(5) = Text5.Text
If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS1.Update
End If
Command1.Caption = "&ADD"
COMBOLIST
Combo1.Enabled = True
End If
End Sub

Private Sub Command2_Click()
If MsgBox("Record Deleted", vbQuestion + vbYesNo, "DELETE") = vbYes Then
RS2.Delete
RS2.MoveFirst
DB1.Execute ("Delete From bus where brno='" & (Text1.Text) & "'")
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


If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS2.Update
End If
Command3.Caption = "&EDIT"
COMBOLIST
End If
End Sub

Private Sub Combo1_Click()
Set RS2 = DB1.OpenRecordset("select * from bus Where bno=" _
& (Combo1.Text))
If RS2.RecordCount = 1 Then
Combo1.Text = RS2.Fields(0)
Text1.Text = RS2.Fields(1)
Text2.Text = RS2.Fields(2)
Text3.Text = RS2.Fields(3)
Text4.Text = RS2.Fields(4)
Text5.Text = RS2.Fields(5)
Command2.Enabled = True
Command3.Enabled = True
Else
MsgBox "Enter  Data"
End If
End Sub

Public Sub COMBOLIST()
Combo1.Clear
Set RS3 = DB1.OpenRecordset("select * from bus")
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



Private Sub Text4_LostFocus()
MDATE = DateSerial(Year(Text4.Text) + 10, Month(Text4.Text), Day(Text4.Text))
Text5.Text = MDATE
End Sub
