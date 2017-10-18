VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form loadBackUpForm 
   BackColor       =   &H80000012&
   Caption         =   "Load Back Up"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   Picture         =   "load_backUp.frx":0000
   ScaleHeight     =   2790
   ScaleWidth      =   7710
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
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
      Format          =   80347137
      CurrentDate     =   40809
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
End
Attribute VB_Name = "loadBackUpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
MsgBox "s"
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
MsgBox "s"
End Sub

