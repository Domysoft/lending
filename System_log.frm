VERSION 5.00
Begin VB.Form System_log 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MM Lending log file"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "System_log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Width = System_log.Width - 450
Text1.Height = System_log.Height - 750
End Sub

Private Sub Form_Resize()
If System_log.Width < 1000 Or System_log.Height < 1000 Then
Exit Sub
Else
Text1.Width = System_log.Width - 450
Text1.Height = System_log.Height - 750
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Main.Enabled = True
End Sub

