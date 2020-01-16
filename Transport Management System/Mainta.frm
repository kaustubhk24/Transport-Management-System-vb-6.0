VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Maintanance Details"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form7"
   ScaleHeight     =   8595
   ScaleWidth      =   10500
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
      Left            =   7080
      TabIndex        =   29
      Top             =   5520
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
      Left            =   7080
      TabIndex        =   28
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Height          =   2655
      Left            =   6720
      TabIndex        =   20
      Top             =   1560
      Width           =   1695
      Begin VB.CommandButton Command8 
         Caption         =   " MOVE LAST"
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "MOVE NEXT"
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "MOVE PREVIOUS"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   " MOVE FIRST"
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   1560
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&CLOSE"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&EDIT"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Mainta.frx":0000
         Left            =   1920
         List            =   "Mainta.frx":0010
         TabIndex        =   19
         Text            =   " "
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Text            =   " "
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Serviceing  Date"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "  Rectification             Details"
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "    Serviceing              Details"
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "    Maintanance            Types"
         Height          =   495
         Left            =   600
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "    Number of           Kilometers"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   " Bus Registration        Number"
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   1920
      TabIndex        =   9
      Top             =   6960
      Width           =   5295
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   1800
      TabIndex        =   10
      Top             =   6840
      Width           =   5535
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   1680
      TabIndex        =   11
      Top             =   6720
      Width           =   5775
   End
   Begin VB.Frame Frame6 
      Height          =   2895
      Left            =   6600
      TabIndex        =   26
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Frame Frame7 
      Height          =   5055
      Left            =   1320
      TabIndex        =   27
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Frame Frame8 
      Height          =   1335
      Left            =   6960
      TabIndex        =   30
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   " MAINTANANCE DETAILS"
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
      Left            =   1560
      TabIndex        =   25
      Top             =   360
      Width           =   6855
   End
End
Attribute VB_Name = "Form7"
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
Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
End Sub
Public Sub DISPLAY()
Combo1.Text = RS2.Fields(0)
Text4.Text = RS2.Fields(1)
Combo2.Text = RS2.Fields(2)
Text1.Text = RS2.Fields(3)

If Combo2.Text = "GENERAL" Then
    Label4.Caption = "Serviceing Details"
    Text1.Text = "Water Service Greasing Minor Repairs"
    Label5.Caption = ""
    Text2.Enabled = False
ElseIf Combo2.Text = "REPAIR & SERVICE" Then
    Label4.Caption = "       Faults"
    Label5.Enabled = True
    Text2.Enabled = True
    Label5.Caption = "  Rectifications"
    Text2.Enabled = True
ElseIf Combo2.Text = "SPARE PARTS" Then
    Label4.Caption = "Spare Parts Name"
    Label5.Caption = ""
    Text2.Enabled = False
ElseIf Combo2.Text = "ACCESSORIES" Then
    Label4.Caption = "Accessories Name"
    Label5.Caption = ""
    Text2.Enabled = False
End If


If RS2.Fields(4) = Null Or RS2.Fields(4) = "" Then
Text2.Text = "Nil"
ElseIf RS2.Fields(4) <> "" Then
Text2.Text = RS2.Fields(4)
End If
Text3.Text = RS2.Fields(5)
End Sub


Private Sub Combo1_Click()
Set RS4 = DB1.OpenRecordset("select * from Petro Where breg='" _
& (Combo1.Text) & "'")
If RS4.RecordCount = 1 Then
Text4.Text = RS4.Fields(3)
Else
MsgBox "New Bus.."
End If
End Sub

Private Sub Combo2_Click()
If Combo2.Text = "GENERAL" Then
    Label4.Caption = "Serviceing Details"
    Text1.Text = "Water Service Greasing Minor Repairs"
    Label5.Caption = ""
    Text2.Enabled = False
ElseIf Combo2.Text = "REPAIR & SERVICE" Then
    Label4.Caption = "       Faults"
    Label5.Enabled = True
    Text2.Enabled = True
    Label5.Caption = "  Rectifications"
    Text1.Text = ""
    Text2.Enabled = True
    Text2.Text = ""
    Text1.SetFocus
ElseIf Combo2.Text = "SPARE PARTS" Then
    Label4.Caption = "Spare Parts Name"
    Text1.SetFocus
    Text1.Text = ""
    Label5.Caption = ""
    Text2.Enabled = False
ElseIf Combo2.Text = "ACCESSORIES" Then
    Label4.Caption = "Accessories Name"
    Text1.Text = ""
    Label5.Caption = ""
    Text1.SetFocus
    Text2.Enabled = False
End If
End Sub


Private Sub Command1_Click()
If Command1.Caption = "&ADD" Then
CLEARF
RS2.AddNew
Command2.Enabled = False
Command3.Enabled = False
Command1.Caption = "&UPDATE"
Combo1.SetFocus
Else
RS2.Fields(0) = Combo1.Text
 RS2.Fields(1) = Text4.Text
 RS2.Fields(2) = Combo2.Text
RS2.Fields(3) = Text1.Text
If Text2.Text = Null Then
RS2.Fields(4) = "Nil"

Else
RS2.Fields(4) = Text2.Text



End If
RS2.Fields(5) = Date$
If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS2.Update
End If
Command1.Caption = "&ADD"
Command2.Enabled = True
Command3.Enabled = True
End If
End Sub


Private Sub Command10_Click()
Form8.Show
End Sub

Private Sub Command2_Click()
If MsgBox("Record Deleted", vbQuestion + vbYesNo, "DELETE") = vbYes Then
RS2.Delete
RS2.MoveFirst
DB1.Execute ("Delete From MaintDet where brno='" & (Combo1.Text) & "'")
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
 RS2.Fields(1) = Text4.Text
 RS2.Fields(2) = Combo2.Text
RS2.Fields(3) = Text1.Text
If Text2.Text = Null Then
RS2.Fields(4) = "Nil"
Else
 RS2.Fields(4) = Text2.Text

End If
RS2.Fields(5) = Text3.Text

If MsgBox("Record Saved", vbExclamation + vbOKOnly, "UPDATE") = vbOK Then
RS2.Update
End If
Command3.Caption = "&EDIT"
End If
End Sub

Private Sub Command4_Click()
Unload Me
Form6.Show
End Sub


Private Sub Command5_Click()
Command2.Enabled = True
Command3.Enabled = True
RS2.MoveFirst
DISPLAY
End Sub

Private Sub Command6_Click()
Command2.Enabled = True
Command3.Enabled = True
If Not RS2.BOF Then
RS2.MovePrevious
If Not RS2.BOF Then
DISPLAY
Else
RS2.MoveFirst
End If
End If
End Sub

Private Sub Command7_Click()
Command2.Enabled = True
Command3.Enabled = True
If Not RS2.EOF Then
RS2.MoveNext
If Not RS2.EOF Then
DISPLAY
Else
RS2.MoveLast
End If
End If
End Sub

Private Sub Command8_Click()
Command2.Enabled = True
Command3.Enabled = True
RS2.MoveLast
DISPLAY
End Sub


Private Sub Command9_Click()
Form5.Show
End Sub

Private Sub Form_Activate()
Text3.Text = Format$(Date$, "d,mmm,yyyy")
Text3.Enabled = False
End Sub
Public Sub COMBOLIST()
Combo1.Clear
Set RS3 = DB1.OpenRecordset("select * from bus")
If Not (RS3.BOF And RS3.EOF) Then
RS3.MoveFirst
While Not RS3.EOF
Combo1.AddItem RS3.Fields(1), Index
RS3.MoveNext
Wend
End If
End Sub


Private Sub Form_Load()
Set DB1 = DBEngine.OpenDatabase(App.Path & "\transport.mdb")
Set RS2 = DB1.OpenRecordset("MaintDet")
COMBOLIST
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Timer1_Timer()
Text3.Text = Format$(Date$, "d,mmm,yyyy")
End Sub
