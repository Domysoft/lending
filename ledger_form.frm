VERSION 5.00
Begin VB.Form ledger_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Ledger"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2475
   Icon            =   "ledger_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2475
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "ledger_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

For i = 0 To List1.ListCount
If Text1.Text = List1.list(i) Then
MsgBox "Ledger found.", vbInformation
GoTo break
End If
Next

MsgBox "Ledger not found!", vbExclamation
Exit Sub

break:
Main.RemLedger.Text = List1.list(i)
Main.remittance_text1.Text = i
Main.Enabled = True
Main.remAmount.SetFocus
Unload Me

End Sub

Private Sub Form_Load()


Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field



Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open




rs.Open "select ledger from client where area_name like '" & Main.SelectAreaRemittance.Text & "%' order by ledger ", conn


Do Until rs.EOF

For Each fld In rs.Fields
List1.AddItem fld.Value
Next
rs.MoveNext

Loop


End Sub
