VERSION 5.00
Begin VB.Form system_lock 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "System Locked"
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   Picture         =   "system_lock.frx":0000
   ScaleHeight     =   8580
   ScaleWidth      =   11730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5880
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4275
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   3525
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   10680
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   6360
      Picture         =   "system_lock.frx":266E1
      Top             =   4920
      Width           =   1665
   End
End
Attribute VB_Name = "system_lock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
log_me_in
End Sub

Private Sub Image1_Click()
log_me_in
End Sub

Private Sub log_me_in()


If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Please fill up the form correctly.", vbInformation
Exit Sub
End If


Dim conn As ADODB.Connection

Dim rs As ADODB.Recordset
Dim fld As ADODB.Field

Dim config As String
Dim content(5) As String
Dim ans1 As Double
Dim ans2 As Double
Dim ans3 As Double
Dim ans4 As Double
Dim ans5 As Double
Dim ans6 As Double
Dim password_text As Double

ans1 = Asc(Text2.Text)
ans2 = AscB(StrReverse(Text2.Text))
ans3 = Asc(Left(Text2.Text, 3))
ans4 = AscB(StrReverse(Left(Text2.Text, 2)))
ans5 = Asc(Right(Text2.Text, 3))
ans6 = AscB(StrReverse(Right(Text2.Text, 4)))

password_text = (ans1 + ans2 + ans3 + ans4 + ans5 + ans6) * (ans1 * ans2 * ans3 * ans4 * ans5 * ans6)

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "select user_full_name, designation from users where user_name = '" & Text1.Text & "' and user_password = '" & password_text & ":" & (Len(Text2.Text) * 144) & "'", conn

If rs.EOF = True Then
MsgBox "User and Password did not match!", vbExclamation
Exit Sub


Else
a = 0
For Each fld In rs.Fields
content(a) = fld.Value
a = a + 1
Next

Main.locker.ForeColor = &H404040
Main.Show
Main.Enabled = True
Unload Me





End If


End Sub

Private Sub Label1_Click()
system_lock.WindowState = 1
End Sub
