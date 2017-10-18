VERSION 5.00
Begin VB.Form EditPayment 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   Icon            =   "EditPayment.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "EditPayment.frx":030A
   ScaleHeight     =   4395
   ScaleWidth      =   7740
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox temp_bal 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox grid_number 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   495
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
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   2880
      Width           =   5295
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
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   0
      EndProperty
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
      TabIndex        =   1
      Top             =   1440
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
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Image close_hover 
      Height          =   270
      Left            =   7080
      Picture         =   "EditPayment.frx":C011
      Top             =   0
      Width           =   645
   End
   Begin VB.Image close_click 
      Height          =   270
      Left            =   7080
      Picture         =   "EditPayment.frx":CA96
      Top             =   0
      Width           =   645
   End
End
Attribute VB_Name = "EditPayment"
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

Private Sub Form_Load()
CustomerHistory.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
close_hover.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

CustomerHistory.Enabled = True

End Sub

Private Sub Label1_Click()


If Text4.Text = "" Then
MsgBox "Please fill up the form correctly!", vbExclamation
Exit Sub
End If

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

split_area = Split(Text6.Text, " ")


query = "update " & LCase(Trim(split_area(0))) & split_area(1) & " set edit_remarks = '" & Text4.Text _
& "' where release_id = '" & Text5.Text & "'"

conn.Execute query


CustomerHistory.flag_amount.Text = Val(Text2.Text) - Val(temp_bal.Text)
CustomerHistory.flag_number.Text = grid_number.Text


CustomerHistory.Text17.Text = Now
Unload Me

End Sub

