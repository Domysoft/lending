VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EditRemittance 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   Icon            =   "EditRemittance.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "EditRemittance.frx":030A
   ScaleHeight     =   5580
   ScaleWidth      =   7710
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox notice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3600
      Width           =   5535
   End
   Begin MSComCtl2.DTPicker update_date 
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2760
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   80478209
      CurrentDate     =   40809
   End
   Begin VB.TextBox date_paid 
      Height          =   285
      Left            =   6840
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   7200
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Text5"
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   7200
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7200
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   7200
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   195
      Left            =   7200
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox edit_remittance_dc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox Edit_remittance_name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox edit_remittnace_ledger 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label update 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2400
      TabIndex        =   8
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Image close_hover 
      Height          =   270
      Left            =   7050
      Picture         =   "EditRemittance.frx":CE92
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image close_click 
      Height          =   270
      Left            =   7050
      Picture         =   "EditRemittance.frx":D917
      Top             =   0
      Width           =   645
   End
End
Attribute VB_Name = "EditRemittance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_click_Click()
Unload Me
End Sub

Private Sub close_hover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
close_hover.Visible = False
End Sub


Private Sub edit_remittance_dc_Change()
Text1.Text = edit_remittance_dc.Text
End Sub

Private Sub Form_Load()
Dim adlaw As Date
Main.Enabled = False

If Main.position.Caption <> "Administrator" Then
update_date.Enabled = False
notice.Enabled = False
Else
update_date.Enabled = True
update_date.Enabled = True
End If


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
close_hover.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Main.Enabled = True
End Sub

Private Sub update_Click()
Dim b_count As Integer

b_count = 0

Dim adlaw_date As Date

adlaw_date = date_paid.Text
If update_date.Value <> adlaw_date And notice.Text = "" Then
MsgBox "If you change the date of payment, please indicate the reason by writing some notes.", vbExclamation
Exit Sub
End If

If notice.Enabled = True And notice.Text = "" And edit_remittance_dc.Text <> Text1.Text Then
MsgBox "If you change the payment on this certain remittance, please indicate the reason by writing some notes.", vbExclamation
Exit Sub
End If

If notice.Enabled = True And (edit_remittance_dc.Text <> Text1.Text Or update_date.Value <> adlaw_date) Then
notice.Text = notice.Text & vbNewLine & " (Changes made from " & Text9.Text & ", " & date_paid.Text & " to " & edit_remittance_dc.Text & ", " & update_date & ")"
End If


Dim conn As ADODB.Connection
Dim config As String
Dim new_area() As String
Dim ok_area As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim query As String
Dim q_balance As Double
Dim q_balance2 As Double

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open


' ----------------   update client for changes

query = "select balance from client where ledger = '" & edit_remittnace_ledger.Text & "'"

rs.Open query, conn

If rs.EOF = True Then
MsgBox "Ledger not found!"
Exit Sub
End If

For Each fld In rs.Fields
q_balance2 = fld.Value
Next

rs.Close
q_balance = (Val(Text6.Text) + Val(Text9.Text)) - Val(edit_remittance_dc.Text)
''q_balance = Val(q_balance2) - (Val(edit_remittance_dc.Text) - Val(Text9.Text))


query = "update client set balance ='" & (Val(q_balance2) + Val(Text9.Text)) - Val(edit_remittance_dc.Text) & "' where ledger ='" & edit_remittnace_ledger.Text & "'"
''query = "update client set balance ='" & q_balance & "' where ledger ='" & edit_remittnace_ledger.Text & "'"

conn.Execute query

' ------------------update area() for changes

new_area = Split(Text8.Text)

ok_area = LCase(Trim(new_area(0))) & Trim(new_area(1))

query = "update " & ok_area & " set balance = '" & q_balance & "', " _
& " d_c ='" & edit_remittance_dc.Text & "', " _
& " date_paid = '" & update_date.Value & "', edit_remarks = '" & notice & "" _
& "' where release_id ='" & Text4.Text & "'"

conn.Execute query

'  ----------------- update loan_release fo changes


''query = "update loan_release set balance ='" & q_balance & "' where release_number = '" & Text7.Text & "'"
query = "update loan_release set balance ='" & (Val(q_balance2) + Val(Text9.Text)) - Val(edit_remittance_dc.Text) & "' where release_number = '" & Text7.Text & "'"

conn.Execute query

Main.view_remittance_blanck.Text = edit_remittance_dc.Text
Main.Text3.Text = update_date.Value
MsgBox "Remittance altered successfuly."

Unload Me

End Sub
