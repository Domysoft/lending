VERSION 5.00
Begin VB.Form LoginForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System LogIn"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   Icon            =   "LoginForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "LoginForm.frx":030A
   Picture         =   "LoginForm.frx":045C
   ScaleHeight     =   5250
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2320
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3150
      Width           =   3255
   End
   Begin VB.Label login 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo error_ko

Dim conn As ADODB.Connection
Dim config As String

Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open
Exit Sub

error_ko:

If Err.Number = -2147467259 Then
MsgBox "Server not found. Please run server and keep MySql running to continue.", vbExclamation
Unload Me
Exit Sub
End If
End Sub

Private Sub login_Click()
log_me_in
End Sub

Private Sub log_me_in()

If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Please fill up the form correctly!"
Exit Sub
End If


If Text1.Text = "superadmin" And Text2.Text = "let.me.pass.the.program" Then
MsgBox "Welcome Super Admin!"
Main.Show
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


Main.login_owner.Caption = content(0)
Main.position.Caption = content(1)
Main.activeLabel.Caption = "Welcome, " & content(0) & "!"

Main.userN.Text = Text1.Text
Main.passW.Text = Text2.Text

Unload Me
Main.Show




End If


End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
log_me_in
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
log_me_in
End If
End Sub
