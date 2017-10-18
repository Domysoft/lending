VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CustomerHistory 
   BorderStyle     =   0  'None
   Caption         =   "Customer History"
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11760
   Icon            =   "CustomerHistory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CustomerHistory.frx":030A
   ScaleHeight     =   8655
   ScaleWidth      =   11760
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox last_flag 
      Height          =   285
      Left            =   5880
      TabIndex        =   22
      Text            =   "Text18"
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox flag_number 
      Height          =   285
      Left            =   5880
      TabIndex        =   21
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox flag_amount 
      Height          =   285
      Left            =   5880
      TabIndex        =   20
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox print_preview_selectArea 
      Height          =   375
      Left            =   11280
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1420
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2035
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2635
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3235
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3835
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   5005
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1435
      Width           =   3495
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2035
      Width           =   3495
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2635
      Width           =   2655
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3235
      Width           =   3615
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3835
      Width           =   2895
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5005
      Width           =   3255
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4435
      Width           =   2655
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4435
      Width           =   2535
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5575
      Width           =   2655
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5575
      Width           =   2775
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Text            =   "Text17"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   6000
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   50
      Cols            =   7
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   4800
      TabIndex        =   18
      Top             =   360
      Width           =   2415
   End
   Begin VB.Image close_hover 
      Height          =   270
      Left            =   11090
      Picture         =   "CustomerHistory.frx":1C56A
      Top             =   0
      Width           =   645
   End
   Begin VB.Image close_click 
      Height          =   270
      Left            =   11090
      Picture         =   "CustomerHistory.frx":1CFEF
      Top             =   0
      Width           =   645
   End
End
Attribute VB_Name = "CustomerHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook

Private Sub close_click_Click()
Unload Me
End Sub

Private Sub close_hover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
close_hover.Visible = False

End Sub

Private Sub flag_number_Change()




Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim query As String
Dim split_area() As String

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

split_area = Split(Text8.Text, " ")





If flag_number.Text = "" Then
Exit Sub
End If

Dim flag_counter As Integer
Dim new_dc As Double
Dim new_bal As Double

Dim new_bal2 As Double

    new_dc = Val(grid.TextMatrix(Val(flag_number.Text), 2)) + Val(flag_amount.Text)
    new_bal = Val(grid.TextMatrix(Val(flag_number.Text), 3)) - Val(flag_amount.Text)
    
      
    grid.TextMatrix(Val(flag_number.Text), 2) = new_dc
    
    grid.TextMatrix(Val(flag_number.Text), 3) = new_bal   ' grid.TextMatrix(Val(flag_number.Text), 4)

    flag_counter = (Val(flag_number.Text) + 1)



query = "update " & LCase(Trim(split_area(0))) & split_area(1) & " set d_c = '" _
& new_dc & "', balance = '" & new_bal & "' where release_id = '" & grid.TextMatrix(Val(flag_number.Text), 4) & "'"

conn.Execute query




Do Until flag_counter = grid.Rows

    new_bal = Val(grid.TextMatrix(Val(flag_counter - 1), 3)) - Val(grid.TextMatrix(Val(flag_counter), 2))
    grid.TextMatrix(flag_counter, 3) = new_bal ' & grid.TextMatrix(flag_counter, 4)
 
 
    query = "update " & LCase(Trim(split_area(0))) & split_area(1) & " set balance = '" & new_bal & "' where release_id = '" & grid.TextMatrix(flag_counter, 4) & "'"

    conn.Execute query

flag_counter = flag_counter + 1
Loop



    query = "update client set balance = '" & new_bal & "' where ledger = '" & Text7.Text & "'"

    conn.Execute query
    
    query = "update loan_release set balance = '" & new_bal & "' where release_number = '" & Text14.Text & "'"

    conn.Execute query

    Text6.Text = new_bal
    
    SelectRelForm.grid.TextMatrix(last_flag, 7) = new_bal
flag_number.Text = ""




End Sub

Private Sub Form_Load()

SelectRelForm.Enabled = False


grid.ColWidth(0) = 300
grid.ColWidth(1) = 3450
grid.ColWidth(2) = 3450
grid.ColWidth(3) = 3450



grid.TextMatrix(0, 0) = ""
grid.TextMatrix(0, 1) = "Date"
grid.TextMatrix(0, 2) = "D/C"
grid.TextMatrix(0, 3) = "Balance"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
close_hover.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
SelectRelForm.Enabled = True
End Sub

Private Sub grid_DblClick()

If Main.position.Caption = "Encoder" Then
MsgBox "Only Administrator can edit costumer's history", vbExclamation
Exit Sub
End If

If grid.TextMatrix(grid.Row, 0) <> "" Then

ans = (MsgBox("Remarks: " & vbNewLine & grid.TextMatrix(grid.Row, 5) _
         & vbNewLine & vbNewLine & "Do you wish to continue editing?" & vbNewLine, vbYesNo))
End If

If ans = 7 Then
Exit Sub
End If


EditPayment.Text1 = grid.TextMatrix(grid.Row, 1)
EditPayment.Text2 = grid.TextMatrix(grid.Row, 2)
EditPayment.Text3 = grid.TextMatrix(grid.Row, 3)
EditPayment.temp_bal = grid.TextMatrix(grid.Row, 2)
EditPayment.Text4 = grid.TextMatrix(grid.Row, 5)
EditPayment.Text5 = grid.TextMatrix(grid.Row, 4)
EditPayment.grid_number.Text = grid.Row
EditPayment.Text6 = Text8.Text

EditPayment.Show
End Sub

Private Sub Label1_Click()

CustomerHistory.Enabled = False
loading.Show
Dim row_counter As Integer

Set xlwbook = xl.Workbooks.Open(App.Path & "\PRINT\transaction.xls")
Set xlsheet = xlwbook.Sheets.Item(1)
    
        xlsheet.Cells(5, 2) = Text1.Text
        If Text2.Text = "Address: _____" Then
            xlsheet.Cells(6, 2) = ""
            Else:   xlsheet.Cells(6, 2) = Text2.Text
        End If
        
        If Text3.Text = "Investment: ______" Then
            xlsheet.Cells(7, 2) = ""
            Else:   xlsheet.Cells(7, 2) = Text3.Text
        End If
        
        If Text9.Text = "Mobile Number: ______" Then
            xlsheet.Cells(7, 7) = ""
            Else:   xlsheet.Cells(7, 7) = Text9.Text
        End If
        
        'xlsheet.Cells(7, 2) = Text3.Text
        xlsheet.Cells(8, 2) = Text4.Text
        xlsheet.Cells(9, 2) = Text5.Text
        xlsheet.Cells(10, 2) = Text13.Text
        xlsheet.Cells(11, 2) = Text6.Text
        xlsheet.Cells(12, 2) = Text15.Text
        xlsheet.Cells(5, 7) = Text7.Text
        xlsheet.Cells(6, 7) = Text8.Text
        
        xlsheet.Cells(8, 7) = Text10.Text
        xlsheet.Cells(9, 7) = Text11.Text
        xlsheet.Cells(10, 7) = Text4.Text
        xlsheet.Cells(11, 7) = Text12.Text
        xlsheet.Cells(12, 7) = Text16.Text
        
        
        For row_counter = 1 To print_preview_selectArea Step 1
        
        xlsheet.Cells(row_counter + 14, 1) = grid.TextMatrix(row_counter, 1)
        xlsheet.Cells(row_counter + 14, 3) = grid.TextMatrix(row_counter, 2)
        xlsheet.Cells(row_counter + 14, 4) = grid.TextMatrix(row_counter, 3)
        xlsheet.Cells(row_counter + 14, 6) = grid.TextMatrix(row_counter, 5)
        
        
        Next

    
    xlwbook.SaveAs (App.Path & "\PRINT\print.xls")
    xl.ActiveWorkbook.Close False, App.Path & "\PRINT\transaction.xls"
    xl.Quit
    
    Set xlwbook = Nothing
    Set xl = Nothing
    
    MsgBox "Data has been sent to the printer!", vbInformation
     CustomerHistory.Enabled = True
    Unload loading
    Shell (App.Path & "\PRINT\printing.exe")
    Shell (App.Path & "\PRINT\delete.bat")

End Sub

Private Sub Text14_Change()
work
End Sub

Private Sub Text17_Change()
work
End Sub

Private Sub work()


Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim area_content(10) As String
Dim spliter() As String
Dim close_logic As Integer

close_logic = 0
Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

spliter = Split(Trim(Text8.Text), " ")

rs.Open "select * from " & LCase(spliter(0)) & spliter(1) & " where release_number = '" & Text14.Text & "'", conn

b = 0
Do Until rs.EOF
b = b + 1
grid.Rows = b + 1
a = 0
For Each fld In rs.Fields
area_content(a) = fld.Value
a = a + 1
Next

'grid.ColAlignment(4) = 2

If area_content(8) <> "" Then
grid.TextMatrix(b, 0) = "*"
End If


grid.TextMatrix(b, 1) = area_content(5)
grid.TextMatrix(b, 2) = area_content(4)
grid.TextMatrix(b, 3) = area_content(3)
grid.TextMatrix(b, 5) = area_content(8)
grid.TextMatrix(b, 4) = area_content(0)

If grid.TextMatrix(b, 3) = "0" Then
Text12.Text = "Paid"
    If Val(area_content(6)) < 0 Then
    Text12.Text = Text12.Text & "  (past due payment) "
    End If
End If
print_preview_selectArea.Text = b

Dim closed_trim() As String


If LCase(Trim(area_content(8))) = "closed" Then
GoTo close_term
End If





rs.MoveNext
Loop

grid.Col = 4
grid.Sort = 1

close_term:
End Sub

