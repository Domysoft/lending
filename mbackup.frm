VERSION 5.00
Begin VB.Form mbackup 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Load Back Up"
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   Picture         =   "mbackup.frx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Warning!"
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   3855
      Begin VB.CommandButton verify 
         Caption         =   "Verify"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label stat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Status:"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label notice 
         BackStyle       =   0  'Transparent
         Caption         =   "Username and password unverify"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   4440
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.CommandButton load 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Load back Up"
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1695
   End
   Begin VB.ListBox date_sel 
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      ItemData        =   "mbackup.frx":4E8E
      Left            =   240
      List            =   "mbackup.frx":4E90
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.ComboBox day_sel 
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   1
      Text            =   "Year"
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox month_sel 
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Text            =   "Month"
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image close2 
      Height          =   270
      Left            =   7080
      Picture         =   "mbackup.frx":4E92
      Top             =   0
      Width           =   645
   End
   Begin VB.Label date_back 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Back up"
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   4370
      Width           =   1815
   End
   Begin VB.Image close1 
      Height          =   270
      Left            =   7080
      Picture         =   "mbackup.frx":58EF
      Top             =   0
      Width           =   645
   End
End
Attribute VB_Name = "mbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub close1_Click()
If stat.Caption <> "Running" Then
MsgBox "Please start server again and click 'Verfiry' to check server status.", vbExclamation
Exit Sub

Else
Main.Enabled = True
Unload Me

End If

End Sub

Private Sub close2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
close2.Visible = False
End Sub

Private Sub date_sel_Click()


'GoTo skp
If stat.Caption = "Running" Then
MsgBox "Server is still running. Please stop server and click 'Verify' to verify server status.", vbInformation
Exit Sub
End If

If stat.Caption = "Unverify" Then
MsgBox "Please click 'Verify' to check if server is still running or not.", vbInformation
Exit Sub
End If

'skp:


month1 = month_sel.ListIndex + 1

If month1 < 10 Then
month1 = "0" & month1
Else
month1 = month1
End If

date_back.Caption = month1 & "-" & List1.list(date_sel.ListIndex) & "-" & day_sel.Text
End Sub

Private Sub day_sel_Click()
query
End Sub

Private Sub Form_Load()
Label1.Caption = "Warning! Loading back up is not recommended." & vbNewLine & vbNewLine & "To load backup, server must be stop first. Click 'Verify' to check server status."

stat.Caption = "Unverify"
For i = 10 To 50
day_sel.AddItem "20" & i
Next

month_sel.AddItem "January"
month_sel.AddItem "Febuary"
month_sel.AddItem "March"
month_sel.AddItem "April"
month_sel.AddItem "May"
month_sel.AddItem "June"
month_sel.AddItem "July"
month_sel.AddItem "August"
month_sel.AddItem "September"
month_sel.AddItem "October"
month_sel.AddItem "November"
month_sel.AddItem "December"



End Sub









Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
close2.Visible = True
End Sub

Private Sub load_Click()
Dim config As String

'GoTo skp

If date_back.Caption = "Select Back up" Then
MsgBox "Please select back up first!", vbInformation
Exit Sub
End If

If stat.Caption <> "Stoped/not running" Then
MsgBox "Please stop server.", vbInformation
Exit Sub
End If

'skp:

Open "path_config.dll" For Input As #1
config = Input$(LOF(1), #1)
Close #1




Dim FSO As New FileSystemObject
Dim fsoFldr As Scripting.Folder
Dim fsoFile As Scripting.File
Dim intCounter As Integer
Set fsoFldr = FSO.GetFolder(App.Path & "\back_up\" & date_back.Caption) 'source folder
fsoFldr.Copy (config)
Set fsoFldr = FSO.GetFolder(config)


MsgBox "Task completed successfully." & vbNewLine & "Loaded back up on " & date_back.Caption & ".", vbInformation

End Sub

Private Sub month_sel_Click()
query
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then
Text1.Text = "Enter username"
End If

End Sub



Private Sub Text2_GotFocus()
Text2.Text = ""
verify.Default = True
End Sub

Private Sub Text2_LostFocus()
If Text2.Text = "" Then
Text2.Text = "Enter password"
End If
End Sub




Private Sub verify_Click()
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field

Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
On Error GoTo data_err
conn.Open


stat.Caption = "Running"
Exit Sub

data_err:
stat.Caption = "Stoped/not running"

End Sub



Private Sub query()
date_sel.Clear
List1.Clear
'Dim month1 As interger

If month_sel.Text = "Month" Or day_sel.Text = "Year" Then
Exit Sub
End If


month1 = month_sel.ListIndex + 1

If month1 < 10 Then
month1 = "0" & month1
Else
month1 = month1
End If


For i = 1 To 31

dir_path = App.Path & "\back_up\" & month1 & "-" & i & "-" & day_sel.Text
'MsgBox dir_path
If Dir$(dir_path, vbDirectory) = "" Then
Else
date_sel.AddItem month_sel.Text & " " & i & ", " & day_sel.Text
List1.AddItem i
End If


Next

End Sub
