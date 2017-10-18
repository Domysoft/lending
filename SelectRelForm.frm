VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SelectRelForm 
   BorderStyle     =   0  'None
   Caption         =   "Select Release"
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   Icon            =   "SelectRelForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "SelectRelForm.frx":030A
   ScaleHeight     =   8640
   ScaleWidth      =   11700
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text6 
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
      Height          =   240
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1545
      Width           =   4095
   End
   Begin VB.TextBox Text5 
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
      Height          =   240
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2920
      Width           =   4095
   End
   Begin VB.TextBox Text4 
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
      Height          =   240
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox Text3 
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
      Height          =   240
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2925
      Width           =   4095
   End
   Begin VB.TextBox Text2 
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
      Height          =   240
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox Text1 
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
      Height          =   240
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1545
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   50
      Cols            =   16
      BackColorSel    =   16744576
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image close_hover 
      Height          =   270
      Left            =   11040
      Picture         =   "SelectRelForm.frx":F6BD
      Top             =   0
      Width           =   645
   End
   Begin VB.Image close_click 
      Height          =   270
      Left            =   11040
      Picture         =   "SelectRelForm.frx":10142
      Top             =   0
      Width           =   645
   End
End
Attribute VB_Name = "SelectRelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub close_click_Click()
Unload Me
End Sub

Private Sub close_hover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
close_hover.Visible = False
close_click.Visible = True
End Sub

Private Sub Form_Load()

Main.Enabled = False


grid.ColWidth(0) = 20
grid.ColWidth(1) = 1500
grid.ColWidth(2) = 1050
grid.ColWidth(3) = 1450
grid.ColWidth(4) = 1450
grid.ColWidth(5) = 1050
grid.ColWidth(6) = 1300
grid.ColWidth(7) = 1300
grid.ColWidth(8) = 2000

grid.ColAlignment(2) = 1
grid.ColAlignment(1) = 1
grid.ColAlignment(3) = 1
grid.ColAlignment(4) = 1
grid.ColAlignment(8) = 1

grid.TextMatrix(0, 1) = "Date Approved"
grid.TextMatrix(0, 2) = "Release #"
grid.TextMatrix(0, 3) = "Maturity Date"
grid.TextMatrix(0, 4) = "Payment Term"
grid.TextMatrix(0, 5) = "Days Left"
grid.TextMatrix(0, 6) = "Amount"
grid.TextMatrix(0, 7) = "Balance"
grid.TextMatrix(0, 8) = "Remarks"

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
close_hover.Visible = True
close_click.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Main.Enabled = True
End Sub



Private Sub grid_DblClick()
CustomerHistory.Text1.Text = Text1.Text
CustomerHistory.Text2.Text = Text2.Text
CustomerHistory.Text3.Text = Text3.Text
CustomerHistory.Text4.Text = Val(grid.TextMatrix(grid.Row, 6))
CustomerHistory.Text5.Text = grid.TextMatrix(grid.Row, 1)
CustomerHistory.Text6.Text = grid.TextMatrix(grid.Row, 7)
CustomerHistory.Text7.Text = Text6.Text
CustomerHistory.Text8.Text = Text4.Text
CustomerHistory.Text9.Text = Text5.Text
CustomerHistory.Text10.Text = grid.TextMatrix(grid.Row, 4)
CustomerHistory.Text11.Text = grid.TextMatrix(grid.Row, 3)
CustomerHistory.Text12.Text = grid.TextMatrix(grid.Row, 8)
CustomerHistory.Text13.Text = grid.TextMatrix(grid.Row, 1)
CustomerHistory.Text14.Text = grid.TextMatrix(grid.Row, 10)
CustomerHistory.Text15.Text = grid.TextMatrix(grid.Row, 9)
CustomerHistory.Text16.Text = grid.TextMatrix(grid.Row, 12)
CustomerHistory.last_flag.Text = grid.Row
CustomerHistory.Show
End Sub

Private Sub grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Main.position.Caption = "Encoder" Then
Exit Sub
End If


If Button = 2 Then
    ans = (MsgBox("Area you sure you want to cancel this release?" & vbNewLine & vbNewLine & "Release amount : " & grid.TextMatrix(grid.Row, 6) & vbNewLine & "Release number : " & grid.TextMatrix(grid.Row, 2) & vbNewLine & vbNewLine & "Please be noted that once release is cancelled, it will be permanent and can't be change or undo.", vbYesNo))
    
    'ans : 6 = yes,, 7 = no
    
    If ans = 6 Then
        
        Dim conn As ADODB.Connection
        Dim config As String
        
        Set conn = New ADODB.Connection

        Open "config.txt" For Input As #1
        config = Input$(LOF(1), #1)
        Close #1
        conn.ConnectionString = config
        conn.Open
        
        query = "update loan_release set remarks = 'Cancelled', date_approve = 'Cancelled', balance = '0' where release_number ='" & grid.TextMatrix(grid.Row, 2) & "'"
        'MsgBox query
        conn.Execute query
        temp_text = Text6.Text
        Text6.Text = 0
        Text6.Text = temp_text
    
    End If
        
End If
End Sub

Private Sub Text6_Change()
Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim fld As ADODB.Field
Dim fld2 As ADODB.Field
Dim client_content(20) As String
Dim spliter() As String

Set rs = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "select * from client where ledger = '" & Text6.Text & "' and area_name = '" & Text4.Text & "'", conn


b = 0

Do Until rs.EOF
b = b + 1
grid.Rows = b + 1

a = 0
For Each fld In rs.Fields
client_content(a) = fld.Value
a = a + 1
Next

Text1.Text = client_content(1)

If client_content(2) <> "" Then
Text2.Text = client_content(2)
Else
Text2.Text = "Address: _____"
End If

If client_content(3) <> "" Then
Text3.Text = client_content(3)
Else
Text3.Text = "Investment: ______"
End If

Text4.Text = client_content(5)

'Text5.Text = client_content(4)
If client_content(4) <> "" Then
Text5.Text = client_content(4)
Else
Text5.Text = "Mobile Number: ______"
End If


rs.MoveNext

Loop

rs.Close


rs.Open "select * from loan_release where ledger  = '" & Text6.Text & "' and area_name = '" & Text4.Text & "'", conn
b = 0
Do Until rs.EOF
b = b + 1
grid.Rows = b + 1

a = 0
For Each fld In rs.Fields
client_content(a) = fld.Value
a = a + 1
Next


grid.TextMatrix(b, 1) = client_content(16)
grid.TextMatrix(b, 2) = client_content(8)
grid.TextMatrix(b, 3) = client_content(2)
grid.TextMatrix(b, 4) = client_content(3)
grid.TextMatrix(b, 5) = client_content(4)
grid.TextMatrix(b, 6) = Val(client_content(5)) / Val(client_content(11))
grid.TextMatrix(b, 7) = client_content(6)

If client_content(6) = "0" Then
grid.TextMatrix(b, 8) = "Paid"
    If Val(client_content(4)) < 0 Then
    grid.TextMatrix(b, 8) = "Paid (Past Due)"
    End If
Else
grid.TextMatrix(b, 8) = "Unpaid"
    If Val(client_content(4)) < 0 Then
    grid.TextMatrix(b, 8) = "Unpaid (Past Due)"
    End If
End If


If client_content(10) = "Cancelled" Then
    grid.TextMatrix(b, 8) = "Cancelled"
End If
'grid.TextMatrix(b - 1, 8) = client_content(10)
grid.TextMatrix(b, 9) = client_content(14)
grid.TextMatrix(b, 10) = client_content(8)
grid.TextMatrix(b, 11) = client_content(1)
grid.TextMatrix(b, 12) = client_content(13)
grid.TextMatrix(b, 15) = client_content(11)
rs.MoveNext
Loop

End Sub
