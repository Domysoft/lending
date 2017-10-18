VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Main 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "MM Lending System"
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   FillColor       =   &H0000C000&
   ForeColor       =   &H0000FF00&
   Icon            =   "cifer_lending_program.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   8160
      TabIndex        =   134
      Top             =   11400
      Width           =   3495
      Begin VB.CommandButton Command13 
         Caption         =   "&g"
         Height          =   315
         Left            =   2160
         TabIndex        =   147
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&p"
         Height          =   375
         Left            =   600
         TabIndex        =   146
         Top             =   3120
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&e"
         Height          =   615
         Left            =   1920
         TabIndex        =   145
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&d"
         Height          =   375
         Left            =   1800
         TabIndex        =   144
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&u"
         Height          =   495
         Left            =   600
         TabIndex        =   143
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&v"
         Height          =   495
         Left            =   3120
         TabIndex        =   142
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&t"
         Height          =   495
         Left            =   2400
         TabIndex        =   141
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&r"
         Height          =   615
         Left            =   1440
         TabIndex        =   140
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&l"
         Height          =   375
         Left            =   480
         TabIndex        =   139
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&o"
         Height          =   495
         Left            =   2880
         TabIndex        =   138
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&c"
         Height          =   495
         Left            =   1560
         TabIndex        =   137
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&a"
         Height          =   495
         Left            =   960
         TabIndex        =   136
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&h"
         Height          =   375
         Left            =   360
         TabIndex        =   135
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.TextBox h_counter 
      Height          =   285
      Left            =   12000
      TabIndex        =   133
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox holder 
      Height          =   375
      Left            =   11880
      TabIndex        =   132
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame help_frame 
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   12000
      TabIndex        =   129
      Top             =   10680
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid help_grid 
         Height          =   4455
         Left            =   120
         TabIndex        =   130
         Top             =   120
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   15
         Cols            =   4
         HighLight       =   2
         ScrollBars      =   2
         Appearance      =   0
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
   End
   Begin VB.Frame developerFrame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   4905
      Left            =   12840
      TabIndex        =   113
      Top             =   9600
      Width           =   8415
      Begin VB.Image Image4 
         Height          =   4710
         Left            =   120
         Picture         =   "cifer_lending_program.frx":030A
         Top             =   120
         Width           =   8025
      End
   End
   Begin VB.Frame CostumerSearchFrame 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5500
      Left            =   13440
      TabIndex        =   96
      Top             =   9120
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid CostumerSearchGrid 
         Height          =   4695
         Left            =   360
         TabIndex        =   97
         Top             =   120
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8281
         _Version        =   393216
         Rows            =   20
         Cols            =   4
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
   End
   Begin VB.Frame PrintpreviewFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5500
      Left            =   14040
      TabIndex        =   91
      Top             =   8400
      Width           =   8415
      Begin VB.TextBox print_preview_counter 
         Height          =   375
         Left            =   6960
         TabIndex        =   148
         Top             =   4800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox print_preview_selectArea 
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "cifer_lending_program.frx":C535
         Left            =   1680
         List            =   "cifer_lending_program.frx":C537
         TabIndex        =   95
         Text            =   "Select Area"
         Top             =   120
         Width           =   2055
      End
      Begin VB.ComboBox print_preview_selescttask 
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "cifer_lending_program.frx":C539
         Left            =   5040
         List            =   "cifer_lending_program.frx":C53B
         TabIndex        =   94
         Text            =   "Select Task"
         Top             =   120
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid PrintPreviewGrid 
         Height          =   3975
         Left            =   360
         TabIndex        =   93
         Top             =   600
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   20
         Cols            =   5
         GridColor       =   16761024
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
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
      Begin VB.Label print_prev_status 
         BackStyle       =   0  'Transparent
         Caption         =   "Creating files…."
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   480
         TabIndex        =   149
         Top             =   4920
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Image print_prev_print 
         Height          =   780
         Left            =   2880
         Picture         =   "cifer_lending_program.frx":C53D
         Top             =   4680
         Width           =   2940
      End
      Begin VB.Image Image12 
         Height          =   5610
         Left            =   240
         Picture         =   "cifer_lending_program.frx":12C97
         Top             =   0
         Width           =   7920
      End
   End
   Begin VB.Frame ViewRemittanceFrame 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00000000&
      Height          =   5535
      Left            =   11520
      TabIndex        =   76
      Top             =   5280
      Width           =   8415
      Begin VB.TextBox view_remittance_blanck 
         Height          =   285
         Left            =   480
         TabIndex        =   110
         Text            =   "Text3"
         Top             =   5160
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComCtl2.DTPicker date_ViewRemiottance 
         Height          =   375
         Left            =   4920
         TabIndex        =   99
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
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
         CalendarTrailingForeColor=   16744576
         Format          =   68419585
         CurrentDate     =   40794
      End
      Begin VB.TextBox total_view_remittance 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   90
         Top             =   5160
         Width           =   2055
      End
      Begin VB.ComboBox SelectAreaViewR 
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   89
         Text            =   "Select Area"
         Top             =   120
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid viewRemittanceGrid 
         Height          =   4335
         Left            =   360
         TabIndex        =   77
         Top             =   600
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   50
         Cols            =   9
         BackColorBkg    =   12632256
         GridColor       =   16761024
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
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
      Begin VB.Image Image11 
         Height          =   5610
         Left            =   240
         Picture         =   "cifer_lending_program.frx":19497
         Top             =   0
         Width           =   7920
      End
   End
   Begin VB.Frame UpdateReleaseFrame 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   15360
      TabIndex        =   74
      Top             =   7680
      Width           =   8415
      Begin VB.TextBox updateRel_blanck 
         Height          =   285
         Left            =   120
         TabIndex        =   106
         Text            =   "Text3"
         Top             =   4320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox series_update_release 
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
         Left            =   3360
         TabIndex        =   88
         Top             =   4380
         Width           =   4095
      End
      Begin VB.TextBox book_no_update_release 
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
         Left            =   3360
         TabIndex        =   87
         Top             =   3880
         Width           =   4095
      End
      Begin VB.TextBox page_update_reelase 
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
         Left            =   3360
         TabIndex        =   86
         Top             =   3440
         Width           =   4095
      End
      Begin VB.TextBox place_update_release 
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
         Left            =   3360
         TabIndex        =   85
         Top             =   2880
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker date_update_release 
         Height          =   375
         Left            =   3360
         TabIndex        =   84
         Top             =   2400
         Width           =   4095
         _ExtentX        =   7223
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
         CalendarTrailingForeColor=   16744576
         Format          =   68419585
         CurrentDate     =   40790
      End
      Begin VB.TextBox day_update_release 
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
         Left            =   3360
         TabIndex        =   83
         Text            =   "57"
         Top             =   2040
         Width           =   4095
      End
      Begin VB.ComboBox term_update_release 
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   82
         Text            =   "Select Term"
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox amount_update_release 
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
         Left            =   3360
         TabIndex        =   81
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox name_update_release 
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox rel_no_update_rel 
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
         Left            =   3360
         TabIndex        =   79
         Text            =   "Enter Release Number"
         Top             =   120
         Width           =   4095
      End
      Begin VB.Image back_to_release 
         Height          =   780
         Left            =   1080
         Picture         =   "cifer_lending_program.frx":20302
         Top             =   4800
         Width           =   2940
      End
      Begin VB.Image UpdateRelease 
         Height          =   780
         Left            =   4680
         Picture         =   "cifer_lending_program.frx":26DC6
         Top             =   4800
         Width           =   2940
      End
      Begin VB.Image Image9 
         Height          =   5505
         Left            =   360
         Picture         =   "cifer_lending_program.frx":2D970
         Top             =   0
         Width           =   8190
      End
      Begin VB.Image Image10 
         Height          =   135
         Left            =   960
         Top             =   5040
         Width           =   855
      End
   End
   Begin VB.Frame new_releaseFrame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   16200
      TabIndex        =   68
      Top             =   6720
      Width           =   8415
      Begin VB.ListBox new_rel_list 
         Height          =   1035
         Left            =   3000
         TabIndex        =   105
         Top             =   720
         Visible         =   0   'False
         Width           =   4400
      End
      Begin VB.TextBox new_rel_blanck 
         Height          =   375
         Left            =   7920
         TabIndex        =   104
         Text            =   "Text3"
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComCtl2.DTPicker date_new_release 
         Height          =   375
         Left            =   3120
         TabIndex        =   73
         Top             =   2880
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTrailingForeColor=   16761024
         Format          =   68419585
         CurrentDate     =   40790
      End
      Begin VB.TextBox percent_new_release 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3120
         TabIndex        =   72
         Text            =   "Value in percent"
         Top             =   2090
         Width           =   4215
      End
      Begin VB.TextBox Amount_new_release 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   71
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox name_new_release 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   70
         Text            =   "Enter the customer's ledger/name to maker release"
         Top             =   360
         Width           =   4215
      End
      Begin VB.Image backToRelease 
         Height          =   780
         Left            =   960
         Picture         =   "cifer_lending_program.frx":3D083
         Top             =   3840
         Width           =   2940
      End
      Begin VB.Image Image8 
         Height          =   780
         Left            =   4560
         Picture         =   "cifer_lending_program.frx":43B47
         Top             =   3840
         Width           =   2940
      End
      Begin VB.Image Image7 
         Height          =   3675
         Left            =   240
         Picture         =   "cifer_lending_program.frx":4A309
         Top             =   0
         Width           =   7380
      End
   End
   Begin VB.Frame userOptionFrame 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   -480
      TabIndex        =   54
      Top             =   11400
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid user_grid 
         Height          =   615
         Left            =   7320
         TabIndex        =   112
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         _Version        =   393216
         Cols            =   7
      End
      Begin VB.TextBox userOp_blanck 
         Height          =   285
         Left            =   7680
         TabIndex        =   111
         Text            =   "Text3"
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox UserOpConfirmPass 
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
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   60
         Top             =   2835
         Width           =   4095
      End
      Begin VB.TextBox UserOpPass 
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
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   59
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox UserOpUsername 
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
         Left            =   2640
         TabIndex        =   58
         Top             =   1875
         Width           =   4095
      End
      Begin VB.TextBox UserOpName 
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
         Left            =   2640
         TabIndex        =   57
         Top             =   1365
         Width           =   4095
      End
      Begin VB.ComboBox UserDesigntion 
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   61
         Text            =   "Select designation"
         Top             =   3360
         Width           =   4335
      End
      Begin VB.ComboBox UserOpSelName 
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   56
         Text            =   "Select Name"
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox UserOptask 
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "cifer_lending_program.frx":53205
         Left            =   1680
         List            =   "cifer_lending_program.frx":53207
         TabIndex        =   55
         Text            =   "Select Task"
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Select designation"
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
         Left            =   360
         TabIndex        =   66
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
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
         Left            =   360
         TabIndex        =   65
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   360
         TabIndex        =   64
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
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
         Left            =   360
         TabIndex        =   63
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   360
         TabIndex        =   62
         Top             =   1440
         Width           =   735
      End
      Begin VB.Image userOpsave 
         Height          =   870
         Left            =   2640
         Picture         =   "cifer_lending_program.frx":53209
         Top             =   3960
         Width           =   3015
      End
      Begin VB.Image uoption 
         Height          =   4020
         Left            =   240
         Picture         =   "cifer_lending_program.frx":59785
         Top             =   240
         Width           =   6990
      End
   End
   Begin VB.Frame remittanceframe 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   3000
      TabIndex        =   40
      Top             =   2760
      Width           =   8415
      Begin VB.TextBox remittance_text6 
         Height          =   285
         Left            =   8040
         TabIndex        =   109
         Text            =   "Text4"
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox remittance_text1 
         Height          =   285
         Left            =   7440
         TabIndex        =   108
         Text            =   "Text3"
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox remittance_blanck_list 
         Height          =   255
         Left            =   7560
         TabIndex        =   107
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox remTotalCollector 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "0.00"
         Top             =   3320
         Width           =   4215
      End
      Begin VB.TextBox remCollector 
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
         Left            =   2520
         TabIndex        =   50
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox RemLedger 
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
         Left            =   2520
         TabIndex        =   48
         Top             =   950
         Width           =   4215
      End
      Begin VB.TextBox remAmount 
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
         Left            =   2520
         TabIndex        =   47
         Top             =   360
         Width           =   4215
      End
      Begin VB.ComboBox SelectAreaRemittance 
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   49
         Text            =   "Select Area"
         Top             =   1440
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker remittance_date 
         Height          =   375
         Left            =   2520
         TabIndex        =   51
         Top             =   2640
         Width           =   4215
         _ExtentX        =   7435
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
         CalendarTrailingForeColor=   16761024
         Format          =   68419585
         CurrentDate     =   40739
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Collection"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   480
         TabIndex        =   46
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Collected Date"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Collector"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   480
         TabIndex        =   44
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   480
         TabIndex        =   43
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ledger"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
      Begin VB.Image saveRel 
         Height          =   870
         Left            =   2520
         Picture         =   "cifer_lending_program.frx":5E94E
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Image Image6 
         Height          =   4290
         Left            =   -360
         Picture         =   "cifer_lending_program.frx":64ECA
         Top             =   0
         Width           =   7695
      End
   End
   Begin VB.Frame releaseFrame 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   -2280
      TabIndex        =   37
      Top             =   10680
      Width           =   8415
      Begin VB.Image upRel 
         Height          =   885
         Left            =   2880
         Picture         =   "cifer_lending_program.frx":6DAC0
         Top             =   2400
         Width           =   2940
      End
      Begin VB.Image new_re 
         Height          =   780
         Left            =   2760
         Picture         =   "cifer_lending_program.frx":74770
         Top             =   1080
         Width           =   3090
      End
   End
   Begin VB.Frame clientOpframe 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   -2520
      TabIndex        =   36
      Top             =   10080
      Width           =   8415
      Begin VB.TextBox SelectAreaClientOp 
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2640
         Width           =   4095
      End
      Begin VB.ListBox clientOpList 
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         ItemData        =   "cifer_lending_program.frx":7AEF0
         Left            =   2640
         List            =   "cifer_lending_program.frx":7AEF2
         TabIndex        =   103
         Top             =   550
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox COpLedger 
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
         Height          =   345
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox COpnumber 
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
         Left            =   2760
         TabIndex        =   20
         Top             =   2160
         Width           =   4095
      End
      Begin VB.TextBox COpIn 
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
         Left            =   2760
         TabIndex        =   19
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox COpAdd 
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
         Left            =   2760
         TabIndex        =   18
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox COpLname 
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
         Left            =   2760
         TabIndex        =   17
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox COpName 
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
         Left            =   2760
         TabIndex        =   16
         Text            =   "Enter the name of person to be search."
         Top             =   240
         Width           =   4095
      End
      Begin VB.Image savecOp 
         Height          =   870
         Left            =   2520
         Picture         =   "cifer_lending_program.frx":7AEF4
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Image Image5 
         Height          =   3705
         Left            =   240
         Picture         =   "cifer_lending_program.frx":81470
         Top             =   0
         Width           =   7125
      End
   End
   Begin VB.Frame areaOptionframe 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   -4080
      TabIndex        =   34
      Top             =   9840
      Width           =   8415
      Begin VB.TextBox areaopblanck 
         Height          =   285
         Left            =   7560
         TabIndex        =   102
         Text            =   "Text3"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox collectorAreaOp 
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
         Left            =   2880
         TabIndex        =   15
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox areaLocAreaOp 
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   4215
      End
      Begin VB.ComboBox AreaSelectAreaOP 
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   13
         Text            =   "    Select Area"
         Top             =   360
         Width           =   4695
      End
      Begin VB.Image sav 
         Height          =   870
         Left            =   2760
         Picture         =   "cifer_lending_program.frx":8E4EC
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Image areaN 
         Height          =   285
         Left            =   480
         Picture         =   "cifer_lending_program.frx":94A68
         Top             =   360
         Width           =   705
      End
      Begin VB.Image saveAOp 
         Height          =   2010
         Left            =   360
         Picture         =   "cifer_lending_program.frx":97AAB
         Top             =   840
         Width           =   8250
      End
   End
   Begin VB.Frame addClientFrame 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   -4680
      TabIndex        =   32
      Top             =   9600
      Width           =   8415
      Begin VB.ComboBox addclientArea 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         Left            =   2520
         TabIndex        =   10
         Text            =   "   Select Area"
         Top             =   3050
         Width           =   4650
      End
      Begin VB.TextBox addclientLegder 
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
         Height          =   240
         Left            =   2760
         TabIndex        =   12
         Top             =   3600
         Width           =   4215
      End
      Begin VB.TextBox addclientNumber 
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
         Height          =   240
         Left            =   2760
         TabIndex        =   9
         Top             =   2640
         Width           =   4215
      End
      Begin VB.TextBox addclientInvestment 
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
         Height          =   240
         Left            =   2760
         TabIndex        =   8
         Top             =   2160
         Width           =   4215
      End
      Begin VB.TextBox addclientAddress 
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
         Height          =   240
         Left            =   2760
         TabIndex        =   7
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox addclientLname 
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
         Height          =   240
         Left            =   2760
         TabIndex        =   6
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox addclientMname 
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
         Height          =   240
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox addclientFname 
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
         Height          =   240
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   4215
      End
      Begin VB.Image addClientUi 
         Height          =   3975
         Left            =   240
         Picture         =   "cifer_lending_program.frx":9E477
         Top             =   0
         Width           =   6960
      End
      Begin VB.Image Image3 
         Height          =   810
         Left            =   4320
         Picture         =   "cifer_lending_program.frx":AB137
         Top             =   3960
         Width           =   2940
      End
   End
   Begin VB.Frame addAreaframe 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   -5280
      TabIndex        =   27
      Top             =   9360
      Width           =   8415
      Begin VB.TextBox addarea_collector 
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
         Left            =   2880
         TabIndex        =   3
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox addarea_location 
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
         Height          =   240
         Left            =   2880
         TabIndex        =   2
         Top             =   1000
         Width           =   4335
      End
      Begin VB.Image addbutton 
         Height          =   810
         Left            =   4440
         Picture         =   "cifer_lending_program.frx":B13F2
         Top             =   2640
         Width           =   2940
      End
      Begin VB.Image addareaUI 
         Height          =   2010
         Left            =   360
         Picture         =   "cifer_lending_program.frx":B76AD
         Top             =   600
         Width           =   8250
      End
   End
   Begin VB.Frame homeFrame 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   -6240
      TabIndex        =   26
      Top             =   9000
      Width           =   8415
      Begin VB.Label last_rel 
         BackStyle       =   0  'Transparent
         Caption         =   "Last release was reformed  6/4/2011"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6360
         TabIndex        =   128
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label bad_release 
         BackStyle       =   0  'Transparent
         Caption         =   "3 release has not yet approved"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   6360
         MouseIcon       =   "cifer_lending_program.frx":BE079
         MousePointer    =   99  'Custom
         TabIndex        =   127
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label bad_client 
         BackStyle       =   0  'Transparent
         Caption         =   "200 clients are in Pass Due Status"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   126
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label good_client 
         BackStyle       =   0  'Transparent
         Caption         =   "1000 clients are in good payment status"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         TabIndex        =   125
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label ave_area 
         BackStyle       =   0  'Transparent
         Caption         =   "Average of 200 clients per Area"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   124
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label release_count 
         BackStyle       =   0  'Transparent
         Caption         =   "20000"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   7080
         TabIndex        =   123
         Top             =   960
         Width           =   615
      End
      Begin VB.Label client_count 
         BackStyle       =   0  'Transparent
         Caption         =   "1200"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4200
         TabIndex        =   122
         Top             =   960
         Width           =   495
      End
      Begin VB.Label area_count 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   121
         Top             =   960
         Width           =   255
      End
      Begin VB.Label total_area 
         BackStyle       =   0  'Transparent
         Caption         =   "Total of 16 Areas."
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
         Left            =   720
         TabIndex        =   120
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label make_back_up 
         BackStyle       =   0  'Transparent
         Caption         =   "Make Back Up"
         BeginProperty Font 
            Name            =   "Catriel"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         MouseIcon       =   "cifer_lending_program.frx":BE1CB
         MousePointer    =   99  'Custom
         TabIndex        =   119
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label back_upwarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Data back up is not yet performed! Click ""Make Back Up "" now!"
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
         Left            =   600
         TabIndex        =   118
         Top             =   3960
         Width           =   5775
      End
      Begin VB.Image back_up_not_ok 
         Height          =   225
         Left            =   240
         Picture         =   "cifer_lending_program.frx":BE31D
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image back_up_ok 
         Height          =   240
         Left            =   240
         Picture         =   "cifer_lending_program.frx":BEA1E
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image release_not_ok1 
         Height          =   225
         Left            =   6000
         Picture         =   "cifer_lending_program.frx":BF126
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image release_ok2 
         Height          =   240
         Left            =   6000
         Picture         =   "cifer_lending_program.frx":BF827
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image release_ok1 
         Height          =   240
         Left            =   6000
         Picture         =   "cifer_lending_program.frx":BFF2F
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image client_not_ok 
         Height          =   225
         Left            =   3120
         Picture         =   "cifer_lending_program.frx":C0637
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image client_ok 
         Height          =   240
         Left            =   3120
         Picture         =   "cifer_lending_program.frx":C0D38
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image area_not_ok 
         Height          =   240
         Left            =   240
         Picture         =   "cifer_lending_program.frx":C1440
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image area_ok 
         Height          =   240
         Left            =   240
         Picture         =   "cifer_lending_program.frx":C1B48
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   6480
         TabIndex        =   117
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   3600
         TabIndex        =   116
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   840
         TabIndex        =   115
         Top             =   960
         Width           =   495
      End
      Begin VB.Image area 
         Height          =   3555
         Left            =   0
         Picture         =   "cifer_lending_program.frx":C2250
         Top             =   120
         Width           =   2685
      End
      Begin VB.Image Image1 
         Height          =   3555
         Left            =   2850
         Picture         =   "cifer_lending_program.frx":C4284
         Top             =   120
         Width           =   2685
      End
      Begin VB.Image Image2 
         Height          =   3555
         Left            =   5700
         Picture         =   "cifer_lending_program.frx":C6707
         Top             =   120
         Width           =   2685
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   11880
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   6240
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   12000
      Top             =   7200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   9240
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   12120
      TabIndex        =   1
      Top             =   7920
      Width           =   255
   End
   Begin VB.TextBox search 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   9600
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "Customer Search"
      Top             =   720
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   9000
   End
   Begin VB.Label help_label 
      BackStyle       =   0  'Transparent
      Caption         =   "Double click topic to view instructions!"
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
      Left            =   11880
      TabIndex        =   131
      Top             =   5160
      Width           =   5655
   End
   Begin VB.Label developerLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't hesitate to contact us!"
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
      Left            =   11760
      TabIndex        =   114
      Top             =   4800
      Width           =   7095
   End
   Begin VB.Label login_owner 
      Height          =   135
      Left            =   11880
      TabIndex        =   101
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label position 
      Caption         =   "Label9"
      Height          =   255
      Left            =   12000
      TabIndex        =   100
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label costumersearchLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Double click the name of the costumer to select."
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
      Left            =   11760
      TabIndex        =   98
      Top             =   4200
      Width           =   7095
   End
   Begin VB.Label PrintPreviewLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Select area and the task you want to perform. And click ""Print"" to print."
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
      Left            =   11880
      TabIndex        =   92
      Top             =   3840
      Width           =   7095
   End
   Begin VB.Label view_remittance_label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Double click the date to edit remittance"
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11760
      TabIndex        =   78
      Top             =   4560
      Width           =   7335
   End
   Begin VB.Label update_release_label 
      BackStyle       =   0  'Transparent
      Caption         =   "Click ""Update Release"" to save and update release."
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
      Left            =   11880
      TabIndex        =   75
      Top             =   3360
      Width           =   6495
   End
   Begin VB.Label new_release_label 
      BackStyle       =   0  'Transparent
      Caption         =   "Click ""Save and Print"" to save and automatically print the release."
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
      Left            =   11880
      TabIndex        =   69
      Top             =   3000
      Width           =   7215
   End
   Begin VB.Label useroptionlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Select task first to proceed. Click ""Save"" to save changes and ""Delete"" to delete user."
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
      Left            =   11880
      TabIndex        =   67
      Top             =   2640
      Width           =   7575
   End
   Begin VB.Label Remittancelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Select ""Area"" first to genarate ledger. Press ""Enter"" or click ""Save"" to save remittance and go to next  ledger."
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   53
      Top             =   2085
      Width           =   7335
   End
   Begin VB.Label releaselabel2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click ""Update Release"" button to update release."
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
      Left            =   11880
      TabIndex        =   39
      Top             =   1680
      Width           =   6975
   End
   Begin VB.Label releaselabel1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click ""New Release"" button the make new release."
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
      Left            =   12000
      TabIndex        =   38
      Top             =   1440
      Width           =   6615
   End
   Begin VB.Label ArOplabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Click the ""Save"" button to save changes."
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
      Left            =   11880
      TabIndex        =   35
      Top             =   1080
      Width           =   7455
   End
   Begin VB.Label addclientlabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Click the ""Add"" Button to Add Client"
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11880
      TabIndex        =   33
      Top             =   720
      Width           =   7575
   End
   Begin VB.Image addClient1 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":C84BB
      Top             =   3015
      Width           =   2700
   End
   Begin VB.Label addareaLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Click the ""Add"" button to Add Area"
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   31
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label homeLabel2 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Keyboard shorcuts are not case sensitive."
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   30
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label homeLabel1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use keyboard shorcuts for quiick navigation. Example , press ""Alt + A"" to Add Area."
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
      Left            =   11880
      TabIndex        =   29
      Top             =   240
      Width           =   7575
   End
   Begin VB.Image notice 
      Height          =   630
      Left            =   2880
      Picture         =   "cifer_lending_program.frx":C981F
      Top             =   2040
      Width           =   8565
   End
   Begin VB.Image close1 
      Height          =   270
      Left            =   11040
      Picture         =   "cifer_lending_program.frx":CDBED
      Top             =   0
      Width           =   645
   End
   Begin VB.Image close2 
      Height          =   270
      Left            =   11040
      Picture         =   "cifer_lending_program.frx":CE64A
      Top             =   0
      Width           =   645
   End
   Begin VB.Label activeLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Wecome back, John Doe"
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AB6121&
      Height          =   495
      Left            =   3120
      TabIndex        =   28
      Top             =   1440
      Width           =   8295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   24
      Top             =   0
      Width           =   9135
   End
   Begin VB.Image min1 
      Height          =   270
      Left            =   10680
      Picture         =   "cifer_lending_program.frx":CF0CF
      Top             =   0
      Width           =   375
   End
   Begin VB.Image min2 
      Height          =   270
      Left            =   10680
      Picture         =   "cifer_lending_program.frx":CF7A1
      Top             =   0
      Width           =   375
   End
   Begin VB.Image logOut1 
      Height          =   330
      Left            =   9240
      Picture         =   "cifer_lending_program.frx":CFE7E
      Top             =   0
      Width           =   1410
   End
   Begin VB.Image logout2 
      Height          =   330
      Left            =   9240
      Picture         =   "cifer_lending_program.frx":D0C07
      Top             =   0
      Width           =   1410
   End
   Begin VB.Image printPreview1 
      Height          =   420
      Left            =   7200
      Picture         =   "cifer_lending_program.frx":D1E26
      Top             =   660
      Width           =   2010
   End
   Begin VB.Image printpreview2 
      Height          =   420
      Left            =   7200
      Picture         =   "cifer_lending_program.frx":D2C80
      Top             =   660
      Width           =   2010
   End
   Begin VB.Image help1 
      Height          =   420
      Left            =   5280
      Picture         =   "cifer_lending_program.frx":D3AA2
      Top             =   660
      Width           =   1215
   End
   Begin VB.Image help2 
      Height          =   420
      Left            =   5280
      Picture         =   "cifer_lending_program.frx":D4759
      Top             =   660
      Width           =   1215
   End
   Begin VB.Image viewRemittance1 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":D53DB
      Top             =   5790
      Width           =   2700
   End
   Begin VB.Image viewRemittance2 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":D66D6
      Top             =   5790
      Width           =   2700
   End
   Begin VB.Image remittance1 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":D866F
      Top             =   5265
      Width           =   2700
   End
   Begin VB.Image remittance2 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":D97B2
      Top             =   5260
      Width           =   2700
   End
   Begin VB.Image release1 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":DB4A9
      Top             =   4680
      Width           =   2700
   End
   Begin VB.Image release2 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":DC455
      Top             =   4680
      Width           =   2700
   End
   Begin VB.Image clientOption1 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":DDFAB
      Top             =   4120
      Width           =   2700
   End
   Begin VB.Image clientOption2 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":DF279
      Top             =   4120
      Width           =   2700
   End
   Begin VB.Image areaOption1 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":E1174
      Top             =   3560
      Width           =   2700
   End
   Begin VB.Image areaOption2 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":E2304
      Top             =   3560
      Width           =   2700
   End
   Begin VB.Image addArea1 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":E410A
      Top             =   2445
      Width           =   2700
   End
   Begin VB.Image addArea2 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":E523D
      Top             =   2450
      Width           =   2700
   End
   Begin VB.Image addClient2 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":E6ED0
      Top             =   3010
      Width           =   2700
   End
   Begin VB.Image home1 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":E8D56
      Top             =   1850
      Width           =   2700
   End
   Begin VB.Image home2 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":E9CFF
      Top             =   1850
      Width           =   2700
   End
   Begin VB.Label lbltime 
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
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   8280
      Width           =   4815
   End
   Begin VB.Image developers1 
      Height          =   390
      Left            =   9720
      Picture         =   "cifer_lending_program.frx":EB761
      Top             =   8160
      Width           =   1905
   End
   Begin VB.Image customerSearch1 
      Height          =   360
      Left            =   9360
      Picture         =   "cifer_lending_program.frx":EC3B7
      Top             =   660
      Width           =   2280
   End
   Begin VB.Image userOption1 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":ECE3B
      Top             =   6360
      Width           =   2700
   End
   Begin VB.Image useroption2 
      Height          =   510
      Left            =   120
      Picture         =   "cifer_lending_program.frx":EE148
      Top             =   6360
      Width           =   2700
   End
   Begin VB.Image developers2 
      Height          =   390
      Left            =   9720
      Picture         =   "cifer_lending_program.frx":F001A
      Top             =   8160
      Width           =   1905
   End
   Begin VB.Image body 
      Height          =   7530
      Left            =   0
      Picture         =   "cifer_lending_program.frx":F0D42
      Top             =   1080
      Width           =   11715
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MM Lending System"
      BeginProperty Font 
         Name            =   "Catriel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11775
   End
   Begin VB.Image header 
      Height          =   1080
      Left            =   0
      Picture         =   "cifer_lending_program.frx":F6BB9
      Top             =   0
      Width           =   11715
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pindot As Boolean
Dim x_1 As Integer
Dim y_1 As Integer
Dim active As String

Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field

Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook


Dim userlist_string As String






Private Sub activeLabel_Click()
If search.Text = "" Then
search.Text = "Customer Search"
search.FontItalic = True
search.ForeColor = &H808080
End If
End Sub

Private Sub addArea1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If active <> "addArea" Then
addArea1.Visible = False
End If

If active = "home" Then
home1.Visible = False
Else
home1.Visible = True
End If

If active = "addClient" Then
addClient1.Visible = False
Else
addClient1.Visible = True
End If

End Sub

Private Sub addArea2_Click()

If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

addarea_location.Text = ""
addarea_collector.Text = ""
hide_all
show_area
End Sub

Private Sub addbutton_Click()

If addarea_location.Text = "" Or addarea_collector.Text = "" Then
GoTo error_savings
End If

Dim newArea() As String
Dim area_splited As String


Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open


area_splited = Trim(addarea_location.Text)
newArea = Split(area_splited, " ")

On Error GoTo error_savings

If newArea(0) <> "Area" Then
GoTo error_savings
End If

Dim area_no As Integer
area_no = Val(Right(area_splited, Len(area_splited) - 4))

If area_no = 0 Then
GoTo error_savings
End If



rs.Open "SELECT area_name FROM area WHERE area_name = '" & Trim(newArea(0)) & " " & area_no & "'", conn

If rs.EOF Then
Else
MsgBox "Area name already exist!"
Exit Sub
End If



ans = (MsgBox("Area you sure you want to save?", vbYesNoCancel))

If ans = 2 Then
return_home
Exit Sub
End If

If ans = 7 Then
Exit Sub
End If





'area_splited = Trim(newArea(0)) & Trim(newArea(1))


query = "CREATE TABLE  " & Trim(newArea(0)) & area_no & " (release_id double NOT NULL AUTO_INCREMENT PRIMARY KEY , area_name VARCHAR( 255 ) NOT NULL , ledger VARCHAR( 255 ) NOT NULL , balance VARCHAR( 255 ) NOT NULL , d_c VARCHAR( 255 ) NOT NULL , date_paid VARCHAR( 255 ) NOT NULL, days_left VARCHAR( 255 ) NOT NULL, release_number VARCHAR( 255 ) NOT NULL, edit_remarks VARCHAR( 255 ) NOT NULL);"
query2 = "INSERT INTO area (area_name, collector) values ('" & Trim(newArea(0)) & " " & area_no & "', ' " & addarea_collector.Text & "')"
conn.Execute query
conn.Execute query2

MsgBox "Adding area completed successfully!"
return_home

Exit Sub
error_savings:
MsgBox "Error saving data. Please fill up the form correctly!", vbExclamation



End Sub

Private Sub addClient1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If active <> "addClient" Then
addClient1.Visible = False
End If



If active = "addArea" Then
addArea1.Visible = False
Else
addArea1.Visible = True
End If


If active = "areaOption" Then
areaOption1.Visible = False
Else
areaOption1.Visible = True
End If

End Sub

Private Sub addClient2_Click()

If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

hide_all
show_client

End Sub

Private Sub addclientArea_Click()
addclientArea.Locked = False

Dim conn As ADODB.Connection
Dim config As String
Dim query As String
Dim ledger_val As String
Dim area As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

query = "select ledger from client where area_name = '" & addclientArea.Text & "' order by ledger desc limit 1"

rs.Open query, conn

If rs.EOF = True Then
addclientLegder.Text = "Please initialize ledger"
Exit Sub
End If

For Each fld In rs.Fields
ledger_val = fld.Value
Next

addclientLegder.Text = Val(ledger_val) + 1

End Sub

Private Sub addclientArea_KeyDown(KeyCode As Integer, Shift As Integer)
addclientArea.Locked = True
End Sub

Private Sub addclientArea_KeyUp(KeyCode As Integer, Shift As Integer)
addclientArea.Locked = False
End Sub

Private Sub addclientLegder_Click()
If addclientLegder.Text = "Please initialize ledger" Then
addclientLegder.Text = ""
End If
End Sub

Private Sub areaLocAreaOp_Click()
MsgBox "You can't edit the Area Name/Location due to previous transactions." & vbNewLine _
    & "You can only change the Preferred Area Collector on this option."
End Sub

Private Sub areaOption1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If active <> "areaOption" Then
areaOption1.Visible = False
End If

If active = "addArea" Then
addArea1.Visible = False
Else
addArea1.Visible = True
End If


If active = "clientOption" Then
clientOption1.Visible = False
Else
clientOption1.Visible = True
End If

End Sub

Private Sub areaOption2_Click()

If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

hide_all
show_areaOption


End Sub

Private Sub AreaSelectAreaOP_Click()
AreaSelectAreaOP.Locked = False

Dim conn As ADODB.Connection
Dim config As String
Dim area_field(4) As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim array_counter As Integer

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "select * from area where area_name = '" & AreaSelectAreaOP.Text & "'", conn
If rs.EOF Then
Exit Sub
End If
array_counter = 0

For Each fld In rs.Fields
area_field(array_counter) = fld.Value
array_counter = array_counter + 1
Next

areaLocAreaOp.Text = area_field(1)
collectorAreaOp.Text = area_field(2)
areaopblanck.Text = area_field(0)

End Sub

Private Sub AreaSelectAreaOP_KeyDown(KeyCode As Integer, Shift As Integer)
AreaSelectAreaOP.Locked = True
End Sub

Private Sub AreaSelectAreaOP_KeyUp(KeyCode As Integer, Shift As Integer)
AreaSelectAreaOP.Locked = False
End Sub

Private Sub back_to_release_Click()

hide_all
show_release

If active <> "release" Then
release1.Visible = False
End If

If active = "clientOption" Then
clientOption1.Visible = False
Else
clientOption1.Visible = True
End If

If active = "remittance" Then
remittance1.Visible = False
Else
remittance1.Visible = True
End If

End Sub

Private Sub backToRelease_Click()

hide_all
show_release

If active <> "release" Then
release1.Visible = False
End If

If active = "clientOption" Then
clientOption1.Visible = False
Else
clientOption1.Visible = True
End If

If active = "remittance" Then
remittance1.Visible = False
Else
remittance1.Visible = True
End If

End Sub

Private Sub bad_release_Click()
If bad_release <> "All relase has been approved." Then
MsgBox "The following are the relase number(s) that has not been approved." & vbNewLine & holder.Text, vbInformation
Else: MsgBox "All relase has been approved.", vbInformation
End If
End Sub

Private Sub body_Click()
If search.Text = "" Then
search.Text = "Customer Search"
search.FontItalic = True
search.ForeColor = &H808080
End If
End Sub

Private Sub body_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

home1.Visible = True
addClient1.Visible = True
addArea1.Visible = True
areaOption1.Visible = True
clientOption1.Visible = True
release1.Visible = True
remittance1.Visible = True
viewRemittance1.Visible = True
userOption1.Visible = True
help1.Visible = True
printPreview1.Visible = True
developers1.Visible = True

If active = "home" Then
home1.Visible = False

ElseIf active = "addClient" Then
addClient1.Visible = False

ElseIf active = "addArea" Then
addArea1.Visible = False

ElseIf active = "areaOption" Then
areaOption1.Visible = False

ElseIf active = "clientOption" Then
clientOption1.Visible = False

ElseIf active = "release" Then
release1.Visible = False

ElseIf active = "remittance" Then
remittance1.Visible = False

ElseIf active = "viewRemittance" Then
viewRemittance1.Visible = False

ElseIf active = "userOption" Then
userOption1.Visible = False

ElseIf active = "developers" Then
developers1.Visible = False

ElseIf active = "help" Then
help1.Visible = False

ElseIf active = "printPreview" Then
printPreview1.Visible = False
End If
End Sub

Private Sub clientOpList_DblClick()
clientOpList.Visible = False
Dim list() As String
Dim list_1 As String
list_2 = clientOpList.list(clientOpList.ListIndex)


list_2 = Trim(list_2)
list = Split(list_2, ",")


Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim client_content(7) As String

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "select * from client where ledger = '" & list(0) & "'", conn


client_counter = 0
For Each fld In rs.Fields
client_content(client_counter) = fld.Value
client_counter = client_counter + 1
Next


COpLname.Text = client_content(1)
COpAdd.Text = client_content(2)
COpIn.Text = client_content(3)
COpnumber.Text = client_content(4)
SelectAreaClientOp.Text = client_content(5)
COpLedger.Text = client_content(0)
COpName.Text = client_content(1)
clientOpList.Visible = False
End Sub

Private Sub clientOption1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If active <> "clientOption" Then
clientOption1.Visible = False
End If

If active = "areaOption" Then
areaOption1.Visible = False
Else
areaOption1.Visible = True
End If

If active = "release" Then
release1.Visible = False
Else
release1.Visible = True
End If

End Sub

Private Sub clientOption2_Click()

If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

hide_all
show_ClientOption



End Sub

Private Sub close1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
close1.Visible = False
close2.Visible = True
End Sub

Private Sub close2_Click()
Timer2.Enabled = True
End Sub

Private Sub Command1_Click()

hide_all
show_home
End Sub

Private Sub Command10_Click()
hide_all
show_dev
End Sub

Private Sub Command11_Click()
hide_all
help_show
End Sub

Private Sub Command12_Click()
hide_all
show_print
End Sub

Private Sub Command13_Click()
Unload Me
LoginForm.Show
End Sub

Private Sub Command2_Click()
If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

addarea_location.Text = ""
addarea_collector.Text = ""
hide_all
show_area
End Sub

Private Sub Command3_Click()
If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

hide_all
show_client
End Sub

Private Sub Command4_Click()
If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

hide_all
show_areaOption
End Sub

Private Sub Command5_Click()

If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

hide_all
show_ClientOption
End Sub

Private Sub Command6_Click()
hide_all
show_release
End Sub

Private Sub Command7_Click()
hide_all
show_remittance
End Sub

Private Sub Command8_Click()
hide_all
show_viewRemittance
End Sub

Private Sub Command9_Click()
If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

hide_all
show_userOption
End Sub

Private Sub COpLedger_GotFocus()
MsgBox "Client's Ledger can't be change due to previous transactions."
End Sub

Private Sub COpName_Change()
clientOpList.Visible = True

clientOpList.Clear

Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim client_content(7) As String

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "select * from client where client_name like '" & COpName.Text & "%'", conn


Do Until rs.EOF
client_counter = 0
For Each fld In rs.Fields
client_content(client_counter) = fld.Value
client_counter = client_counter + 1
Next

clientOpList.AddItem client_content(0) & ",      " & client_content(1)

rs.MoveNext
Loop
rs.Close

rs.Open "select * from client where ledger like '" & COpName.Text & "%'", conn


Do Until rs.EOF
client_counter = 0
For Each fld In rs.Fields
client_content(client_counter) = fld.Value
client_counter = client_counter + 1
Next

clientOpList.AddItem client_content(0) & ",      " & client_content(1)

rs.MoveNext
Loop


End Sub

Private Sub COpName_Click()
COpName.Text = ""
End Sub

Private Sub CostumerSearchGrid_DblClick()
If CostumerSearchGrid.TextMatrix(CostumerSearchGrid.Row, 1) = "" Then
Exit Sub
End If
SelectRelForm.Text4.Text = CostumerSearchGrid.TextMatrix(CostumerSearchGrid.Row, 1)
SelectRelForm.Text6.Text = CostumerSearchGrid.TextMatrix(CostumerSearchGrid.Row, 2)


SelectRelForm.Show
End Sub

Private Sub date_new_release_Change()
Dim date_1() As String
date_1 = Split(date_new_release.Value, "/")
new_rel_blanck.Text = date_1(0) & "-" & date_1(1) & "-" & date_1(2)

End Sub

Private Sub exe_prog()
viewRemittanceGrid.Rows = 1
If SelectAreaViewR.Text = "Select Area" Then
Exit Sub
End If

total_view_remittance.Text = 0
Dim conn As ADODB.Connection
Dim config As String
Dim newArea() As String
Dim newdate() As String
Dim content(10) As String
Dim area As String
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim fld As ADODB.Field
Dim fld2 As ADODB.Field
Dim client_name As String

Set rs = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open


newArea = Split(SelectAreaViewR.Text, " ")
newdate = Split(date_ViewRemiottance.Value, "/")



'If Len(newdate(0)) = 1 Then
'newdate(0) = "0" & newdate(0)
'End If

'If Len(newdate(1)) = 1 Then
'newdate(1) = "0" & newdate(1)
'End If

'MsgBox "SELECT ledger,d_c,area_name,release_id, edit_remarks, balance,release_number FROM " & LCase(newArea(0)) & newArea(1) & " where date_paid = '" & newdate(0) & "/" & newdate(1) & "/" & newdate(2) & "'"

rs.Open "SELECT ledger,d_c,area_name,release_id, edit_remarks, balance,release_number, date_paid FROM " & LCase(newArea(0)) & newArea(1) & " where date_paid = '" & newdate(0) & "/" & newdate(1) & "/" & newdate(2) & "'", conn

b = 0

Do Until rs.EOF
b = b + 1
viewRemittanceGrid.Rows = b + 1
a = 0
For Each fld In rs.Fields

content(a) = fld.Value
    rs2.Open "select client_name from client where ledger = '" & content(0) & "' and area_name = '" & content(2) & "'", conn
    Do Until rs2.EOF
    For Each fld2 In rs2.Fields
      client_name = fld2.Value
    Next
    rs2.MoveNext
    Loop
    rs2.Close
a = a + 1
Next


viewRemittanceGrid.TextMatrix(b, 1) = content(0)
viewRemittanceGrid.TextMatrix(b, 3) = content(1)
total_view_remittance.Text = Val(total_view_remittance.Text) + Val(content(1))
viewRemittanceGrid.TextMatrix(b, 2) = client_name
viewRemittanceGrid.TextMatrix(b, 4) = content(3)
viewRemittanceGrid.TextMatrix(b, 6) = content(5)
viewRemittanceGrid.TextMatrix(b, 7) = content(6)
viewRemittanceGrid.TextMatrix(b, 8) = content(7)


next_loop:
rs.MoveNext
Loop
viewRemittanceGrid.ColAlignment(1) = 0
viewRemittanceGrid.ColAlignment(3) = 0
viewRemittanceGrid.Col = 4
viewRemittanceGrid.Sort = 1

End Sub

Private Sub date_ViewRemiottance_Change()
exe_prog
End Sub

Private Sub day_update_release_Click()
day_update_release.Text = ""
End Sub

Private Sub developers1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If active <> "developers" Then
developers1.Visible = False
End If
End Sub

Private Sub developers2_Click()
hide_all
show_dev
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()
load_query
load_back_up
Main.Width = 11715
Main.Height = 8610

homeFrame.Appearance = 0
addAreaframe.Appearance = 0
addClientFrame.Appearance = 0
areaOptionframe.Appearance = 0
clientOpframe.Appearance = 0
releaseFrame.Appearance = 0
remittanceframe.Appearance = 0
userOptionFrame.Appearance = 0
new_releaseFrame.Appearance = 0
UpdateReleaseFrame.Appearance = 0
ViewRemittanceFrame.Appearance = 0
PrintpreviewFrame.Appearance = 0
CostumerSearchFrame.Appearance = 0
developerFrame.Appearance = 0
help_frame.Appearance = 0


' for help grid ----------------

help_grid.ColWidth(0) = 300
help_grid.ColWidth(1) = 2500
help_grid.ColWidth(2) = 2500
help_grid.ColWidth(3) = 2800

help_grid.TextMatrix(0, 1) = "What's on?"
help_grid.TextMatrix(0, 2) = "How to "
help_grid.TextMatrix(0, 3) = "How to "

help_grid.TextMatrix(1, 0) = "1"
    help_grid.TextMatrix(1, 1) = "Home (H)"
    help_grid.TextMatrix(1, 2) = "make back up?"
    help_grid.TextMatrix(1, 3) = "know unapproved release?"

help_grid.TextMatrix(2, 0) = "2"
    help_grid.TextMatrix(2, 1) = "Add Area (A)"
    help_grid.TextMatrix(2, 2) = "add area?"
    
help_grid.TextMatrix(3, 0) = "3"
    help_grid.TextMatrix(3, 1) = "Add Client (C)"
    help_grid.TextMatrix(3, 2) = "add client?"
    
help_grid.TextMatrix(4, 0) = "4"
    help_grid.TextMatrix(4, 1) = "Area Option (O)"
    help_grid.TextMatrix(4, 2) = "edit area?"
    
help_grid.TextMatrix(5, 0) = "5"
    help_grid.TextMatrix(5, 1) = "Client Option (L)"
    help_grid.TextMatrix(5, 2) = "edti client info?"
    
help_grid.TextMatrix(6, 0) = "6"
    help_grid.TextMatrix(6, 1) = "Release (R)"
    help_grid.TextMatrix(6, 2) = "make new release?"
    help_grid.TextMatrix(6, 3) = "update approved release?"
    
help_grid.TextMatrix(7, 0) = "7"
    help_grid.TextMatrix(7, 1) = "Remittance (T)"
    help_grid.TextMatrix(7, 2) = "encode daily remittance?"
    
help_grid.TextMatrix(8, 0) = "8"
    help_grid.TextMatrix(8, 1) = "View Remittance (V)"
    help_grid.TextMatrix(8, 2) = "view daily remittance?"
    help_grid.TextMatrix(8, 3) = "edit/change wrong remittance?"

help_grid.TextMatrix(9, 0) = "9"
    help_grid.TextMatrix(9, 1) = "User Option (U)"
    help_grid.TextMatrix(9, 2) = "add user?"
    help_grid.TextMatrix(9, 3) = "edit/change user?"
    
help_grid.TextMatrix(10, 0) = "10"
    help_grid.TextMatrix(10, 1) = "Customer Search"
    help_grid.TextMatrix(10, 2) = "search individual customer?"
    help_grid.TextMatrix(10, 3) = "view customer transaction?"
    help_grid.TextMatrix(11, 2) = "to print transaction?"
    help_grid.TextMatrix(11, 3) = "edit payment transaction?"
    
help_grid.TextMatrix(12, 0) = "11"
    help_grid.TextMatrix(12, 1) = "Print Preview (P)"
    help_grid.TextMatrix(12, 2) = "to view pass due accounts?"
    help_grid.TextMatrix(12, 3) = "to view master list?"
    help_grid.TextMatrix(13, 2) = "to print master list?"
    help_grid.TextMatrix(13, 3) = "to print pass due list?"

help_grid.TextMatrix(14, 0) = "12"
    help_grid.TextMatrix(14, 1) = "Developers (D)"
  
' help grid ends -----------------


'  for release date ---------------

date_new_release.Value = Format(Now, "mm d yyyy")
Dim date_1() As String
date_1 = Split(date_new_release.Value, "/")
new_rel_blanck.Text = date_1(0) & "-" & date_1(1) & "-" & date_1(2)

'release date ends ------------------




' for area combo boxes -----

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "SELECT area_name FROM area ", conn

AreaSelectAreaOP.Clear
addclientArea.Clear
SelectAreaRemittance.Clear
SelectAreaViewR.Clear
print_preview_selectArea.Clear


Do Until rs.EOF

For Each fld In rs.Fields
AreaSelectAreaOP.AddItem fld.Value
addclientArea.AddItem fld.Value
'SelectAreaClientOp.AddItem fld.Value
SelectAreaRemittance.AddItem fld.Value
SelectAreaViewR.AddItem fld.Value
print_preview_selectArea.AddItem fld.Value
Next
rs.MoveNext
Loop
SelectAreaViewR.Text = "Select Area"
print_preview_selectArea.Text = "Select Area"
AreaSelectAreaOP.Text = "Select Area"
addclientArea.Text = "Select Area"
SelectAreaRemittance.Text = "Select Area"
' area comboxes ends --------




'   for view remittance grids --------------

viewRemittanceGrid.ColWidth(0) = 500
viewRemittanceGrid.ColWidth(1) = 1000
viewRemittanceGrid.ColWidth(2) = 4400
viewRemittanceGrid.ColWidth(3) = 1500
viewRemittanceGrid.TextMatrix(0, 0) = ""
viewRemittanceGrid.TextMatrix(0, 1) = "Ledger"
viewRemittanceGrid.TextMatrix(0, 2) = "Name"
viewRemittanceGrid.TextMatrix(0, 3) = "D/C"

'   view remittance ends -----------------


' for print preview grids  --------------------


PrintPreviewGrid.ColAlignment(4) = 0
PrintPreviewGrid.ColAlignment(1) = 0
PrintPreviewGrid.ColAlignment(3) = 0

PrintPreviewGrid.ColWidth(0) = 500
PrintPreviewGrid.ColWidth(1) = 900
PrintPreviewGrid.ColWidth(2) = 3500
PrintPreviewGrid.ColWidth(3) = 1400
PrintPreviewGrid.ColWidth(4) = 20006


PrintPreviewGrid.TextMatrix(0, 0) = ""
PrintPreviewGrid.TextMatrix(0, 1) = "Ledger"
PrintPreviewGrid.TextMatrix(0, 2) = "Name"
PrintPreviewGrid.TextMatrix(0, 3) = "Balance"
PrintPreviewGrid.TextMatrix(0, 4) = "Days Left"


print_preview_selescttask.AddItem "Preview Past Due"
print_preview_selescttask.AddItem "Prview Master List"


' print preview ends -----------


' for costumer search elements ---------

CostumerSearchGrid.ColWidth(0) = 100
CostumerSearchGrid.ColWidth(1) = 1050
CostumerSearchGrid.ColWidth(2) = 1550
CostumerSearchGrid.ColWidth(3) = 5000

CostumerSearchGrid.ColAlignment(1) = 0
CostumerSearchGrid.ColAlignment(2) = 0
CostumerSearchGrid.ColAlignment(3) = 0
CostumerSearchGrid.TextMatrix(0, 1) = "Area"
CostumerSearchGrid.TextMatrix(0, 2) = "Ledger"
CostumerSearchGrid.TextMatrix(0, 3) = "Name"

' costumer search  ends ---------------


' for update release -----------

term_update_release.AddItem "Daily"
term_update_release.AddItem "Weekly"
term_update_release.AddItem "Every Pay Day"
term_update_release.AddItem "Monthly"

'update release ends ------------

' for user option -------------

UserOptask.AddItem "Add User"
UserOptask.AddItem "Edit User"
UserDesigntion.AddItem "Administrator"
UserDesigntion.AddItem "Encoder"


'user option ends --------------





show_home

End Sub


Private Sub header_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
help1.Visible = True
printPreview1.Visible = True
logOut1.Visible = True
min1.Visible = True
close1.Visible = True
End Sub

Private Sub help_grid_DblClick()
If help_grid.Row = 1 And help_grid.Col = 1 Then
MsgBox "A place where you can find the summarized system status information.", vbInformation

ElseIf help_grid.Row = 1 And help_grid.Col = 2 Then
MsgBox "Click the " & Chr(34) & "Make Back Up" & Chr(34) & " phrase!", vbInformation

ElseIf help_grid.Row = 1 And help_grid.Col = 3 Then
MsgBox "Click on the underlined blue color phrase/wors.", vbInformation

ElseIf help_grid.Row = 2 And help_grid.Col = 1 Then
MsgBox "A place where you can add new Area and its collector.", vbInformation

ElseIf help_grid.Row = 2 And help_grid.Col = 2 Then
MsgBox "1) Enter the proper Area Name/Location [example, Area 1, Area7]." & vbNewLine _
        & "2) Enter the preferred collector on that area." & vbNewLine _
        & "3) Click the " & Chr(34) & "Add" & Chr(34) & "  button.", vbInformation
        
ElseIf help_grid.Row = 3 And help_grid.Col = 1 Then
MsgBox "A place where you can add new client.", vbInformation

ElseIf help_grid.Row = 3 And help_grid.Col = 2 Then
MsgBox "1) Enter Client's info from First name to Mobile Number." & vbNewLine _
        & "2) Select Area Location." & vbNewLine _
        & "3) The Ledger will be automatically filled up," & vbNewLine _
        & "     if not, initialize ledger." & vbNewLine _
        & "4) Click the " & Chr(34) & "Add" & Chr(34) & "  button.", vbInformation

ElseIf help_grid.Row = 4 And help_grid.Col = 1 Then
MsgBox "A place where you can change Area collector.", vbInformation

ElseIf help_grid.Row = 4 And help_grid.Col = 2 Then
MsgBox "1) Select Area" & vbNewLine _
        & "2) Change collector's name." & vbNewLine _
        & "3) Click the " & Chr(34) & "Save" & Chr(34) & "  button.", vbInformation

ElseIf help_grid.Row = 5 And help_grid.Col = 1 Then
MsgBox "A place where you can edit client's information.", vbInformation

ElseIf help_grid.Row = 5 And help_grid.Col = 2 Then
MsgBox "1) Click the " & Chr(34) & "Enter the name of the person to search" & Chr(34) & "." & vbNewLine _
        & "2) Type the Surename/Ledger of the client." & vbNewLine _
        & "3) Double click on the dropdown list." & vbNewLine _
        & "4) Change the wrong information(s)." & vbNewLine _
        & "5) Click the " & Chr(34) & "Save" & Chr(34) & "  button.", vbInformation
        
ElseIf help_grid.Row = 6 And help_grid.Col = 1 Then
MsgBox "A place where you make or update/approve release." & vbNewLine _
        & "Click " & Chr(34) & "New Release" & Chr(34) & " to make new release." & vbNewLine _
        & "Click " & Chr(34) & "Update Release" & Chr(34) & " to update/approve release.", vbInformation
       

ElseIf help_grid.Row = 6 And help_grid.Col = 2 Then
MsgBox "1) Click the " & Chr(34) & "Enter the customer's .... " & Chr(34) & "." & vbNewLine _
        & "2) Type the Surename/Ledger of the client." & vbNewLine _
        & "3) Double click on the dropdown list." & vbNewLine _
        & "4) Enter the amount, percent, and date of release." & vbNewLine _
        & "5) Click the " & Chr(34) & "Print and Save Release" & Chr(34) & "  button." & vbNewLine _
        & "6) Wait for printer to print.", vbInformation

ElseIf help_grid.Row = 6 And help_grid.Col = 3 Then
MsgBox "1) Click the " & Chr(34) & "Enter Release Number" & Chr(34) & "." & vbNewLine _
        & "2) Enter the release number found in the upper right side on Promissory Note." & vbNewLine _
        & "3) Fill up all the information needed." & vbNewLine _
        & "4) Click the " & Chr(34) & "Update Release" & Chr(34) & "  button.", vbInformation


ElseIf help_grid.Row = 7 And help_grid.Col = 1 Then
MsgBox "A place where you encode the dailly remittance.", vbInformation

ElseIf help_grid.Row = 7 And help_grid.Col = 2 Then
MsgBox "1) Select the Area Location/Name." & vbNewLine _
        & "2) Enter the Amount on the Amount text box." & vbNewLine _
        & "3) You can hit " & Chr(34) & "Enter" & Chr(34) & " or click " & Chr(34) & "Save" & Chr(34) & "  button.", vbInformation
        

ElseIf help_grid.Row = 8 And help_grid.Col = 1 Then
MsgBox "A place where you can view/edit the dailly remittance.", vbInformation


ElseIf help_grid.Row = 8 And help_grid.Col = 2 Then
MsgBox "1) Select the Area Location/Name and remittance's date.", vbInformation

ElseIf help_grid.Row = 8 And help_grid.Col = 3 Then
MsgBox "1) Select the Area Location/Name and remittance's date." & vbNewLine _
        & "2) Double click client's name." & vbNewLine _
        & "2) Change the D/C." & vbNewLine _
        & "3) Click the " & Chr(34) & "Update" & Chr(34) & "  button.", vbInformation

ElseIf help_grid.Row = 9 And help_grid.Col = 1 Then
MsgBox "A place where you can add/edit user.", vbInformation


ElseIf help_grid.Row = 9 And help_grid.Col = 2 Then
MsgBox "1) Click on " & Chr(34) & "Select Task" & Chr(34) & " and select " & Chr(34) & "Add User" & Chr(34) & "." & vbNewLine _
        & "2) Fill up all the information" & vbNewLine _
        & "3) Click the " & Chr(34) & "Save" & Chr(34) & "  button.", vbInformation


ElseIf help_grid.Row = 9 And help_grid.Col = 3 Then
MsgBox "1) Click on " & Chr(34) & "Select Task" & Chr(34) & " and select " & Chr(34) & "Edit User" & Chr(34) & "." & vbNewLine _
        & "2) Click on " & Chr(34) & "Select User" & Chr(34) & " and select user." & vbNewLine _
        & "3) Fill up all the information" & vbNewLine _
        & "4) Click the " & Chr(34) & "Save" & Chr(34) & "  button.", vbInformation


ElseIf help_grid.Row = 10 And help_grid.Col = 1 Then
MsgBox "A place where you can search/edit/print/view client's transactions.", vbInformation


ElseIf help_grid.Row = 10 And help_grid.Col = 2 Then
MsgBox "1) Click on " & Chr(34) & "Customer Search" & Chr(34) & " on the upper right side of the window." & vbNewLine _
        & "2) Type the Surename of the client.", vbInformation
        
ElseIf help_grid.Row = 10 And help_grid.Col = 3 Then
MsgBox "1) Click on " & Chr(34) & "Customer Search" & Chr(34) & " on the upper right side of the window." & vbNewLine _
        & "2) Type the Surename of the client." & vbNewLine _
        & "3) Double click Customer's name." & vbNewLine _
        & "4) Select a release by clicking it twice [double click].", vbInformation
        
ElseIf help_grid.Row = 11 And help_grid.Col = 2 Then
MsgBox "1) Click on " & Chr(34) & "Customer Search" & Chr(34) & " on the upper right side of the window." & vbNewLine _
        & "2) Type the Surename of the client." & vbNewLine _
        & "3) Double click Customer's name." & vbNewLine _
        & "4) Select a release by clicking it twice [double click]." & vbNewLine _
        & "5) Click the " & Chr(34) & "Print" & Chr(34) & "  button.", vbInformation
        
        
ElseIf help_grid.Row = 11 And help_grid.Col = 3 Then
MsgBox "1) Click on " & Chr(34) & "Customer Search" & Chr(34) & " on the upper right side of the window." & vbNewLine _
        & "2) Type the Surename of the client." & vbNewLine _
        & "3) Double click Customer's name." & vbNewLine _
        & "4) Select a release by clicking it twice [double click]." & vbNewLine _
        & "5) Select wrong payemt by clicking it twice [double click]." & vbNewLine _
        & "6) Change the wrong values." & vbNewLine _
        & "7) Click the " & Chr(34) & "Update" & Chr(34) & "  button.", vbInformation

ElseIf help_grid.Row = 12 And help_grid.Col = 1 Then
MsgBox "A place where you can print/preview master list and pass due transactions.", vbInformation

ElseIf help_grid.Row = 12 And help_grid.Col = 2 Then
MsgBox "1) Click on " & Chr(34) & "Select Area" & Chr(34) & " and select an Area" & vbNewLine _
        & "2) Click on " & Chr(34) & "Select Task" & Chr(34) & " and select " & Chr(34) & "Preview Pass Due" & Chr(34) & ".", vbInformation

ElseIf help_grid.Row = 12 And help_grid.Col = 3 Then
MsgBox "1) Click on " & Chr(34) & "Select Area" & Chr(34) & " and select an Area" & vbNewLine _
        & "2) Click on " & Chr(34) & "Select Task" & Chr(34) & " and select " & Chr(34) & "Preview Master List" & Chr(34) & ".", vbInformation

ElseIf help_grid.Row = 13 And help_grid.Col = 2 Then
MsgBox "1) Click on " & Chr(34) & "Select Area" & Chr(34) & " and select an Area" & vbNewLine _
        & "2) Click on " & Chr(34) & "Select Task" & Chr(34) & " and select " & Chr(34) & "Preview Master List" & Chr(34) & "." & vbNewLine _
        & "3) Click the " & Chr(34) & "Print" & Chr(34) & "  button.", vbInformation


ElseIf help_grid.Row = 13 And help_grid.Col = 3 Then
MsgBox "1) Click on " & Chr(34) & "Select Area" & Chr(34) & " and select an Area" & vbNewLine _
        & "2) Click on " & Chr(34) & "Select Task" & Chr(34) & " and select " & Chr(34) & "Preview Pass Due" & Chr(34) & "." & vbNewLine _
        & "3) Click the " & Chr(34) & "Print" & Chr(34) & "  button.", vbInformation


ElseIf help_grid.Row = 14 And help_grid.Col = 1 Then
MsgBox "A place where you can find the developers information and how to contact them if problems exist.", vbInformation

End If

End Sub

Private Sub help1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
help1.Visible = False
End Sub

Private Sub help2_Click()
hide_all
help_show

End Sub

Private Sub home1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If active <> "home" Then
home1.Visible = False
End If

If active = "addClient" Then
addClient1.Visible = False
Else
addClient1.Visible = True
End If

End Sub

Private Sub home2_Click()



hide_all
show_home

End Sub

Private Sub Image13_Click()

End Sub

Private Sub Image3_Click()

On Error GoTo error_saving
Dim conn As ADODB.Connection
Dim config As String
Dim newArea() As String
Dim area As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim name_split() As String
Dim s_name As String
Dim f_name As String


'name_split = Split(addclientFname.Text, ",")
'On Error GoTo error_name
's_name = name_split(0)
'f_name = Trim(name_split(1))

'If f_name = "" Then
'GoTo error_name
'End If

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection




If addclientFname.Text = "" Or addclientMname.Text = "" Or addclientLname.Text = "" Then
GoTo error_saving
End If

If addclientAddress.Text = "" Or addclientInvestment.Text = "" Or addclientNumber.Text = "" Or Val(addclientLegder.Text) = 0 Then
GoTo error_saving
End If





If Trim(addclientArea.Text) = "Select Area" Then GoTo error_saving

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open


query = "select ledger from client where area_name = '" & addclientArea.Text & "'"
rs.Open query, conn

If rs.EOF = True Then
GoTo next_part
End If


For Each fld In rs.Fields
    If addclientLegder.Text = fld.Value Then
    MsgBox "Error saving!" & vbNewLine _
    & "Duplication of ledger, suggested ledger = " & Val(fld.Value) + 1
    Exit Sub
    End If
Next


next_part:

ans = (MsgBox("Area you sure you want to save?", vbYesNoCancel))

If ans = 2 Then
return_home
Exit Sub
End If

If ans = 7 Then
Exit Sub
End If



query = "insert into client (ledger, client_name, address, investment, mobile_number, area_name) " _
& "values ('" & addclientLegder.Text & "','" & addclientLname.Text & ", " & addclientFname.Text & " " & Left(Trim(addclientMname.Text), 1) & "','" & addclientAddress.Text & "','" & addclientInvestment.Text & "','" & addclientNumber.Text & "','" & addclientArea.Text & "')"
conn.Execute query

MsgBox "Adding client completed successfully!"
return_home
Exit Sub

error_saving:
MsgBox "Error saving data. Please fill up the form correctly!", vbExclamation
Exit Sub

error_name:
MsgBox "Please enter the full name." & vbNewLine & "Enter first the lastname followed by ',' and then the firstname."


End Sub


Private Sub Image8_Click()

Dim print_string As String


Dim pass As Boolean
pass = False
If Val(Amount_new_release.Text) = 0 Or Val(percent_new_release.Text) = 0 Or name_new_release.Text = "Enter the customer's ledger/name to maker release" Then
GoTo error_saving
End If

ans = (MsgBox("Area you sure you want to save?", vbYesNoCancel))

If ans = 2 Then
Unload Me
Exit Sub
End If

If ans = 7 Then
Exit Sub
End If

Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim ledger_1() As String
Dim ledger As String
Dim pd_amount(10) As String

'On Error GoTo error_saving
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

ledger_1 = Split(name_new_release.Text, ",")

query = "select balance,release_number, maturity from loan_release where ledger= '" & ledger_1(0) & "' order by release_number desc limit 1"

rs.Open query, conn

If rs.EOF = True Then
pd_amount(0) = 0
GoTo endOfFile
End If


a = 0
For Each fld In rs.Fields
pd_amount(a) = fld.Value
a = a + 1
Next


If pd_amount(2) = "" Then
MsgBox "Previous release has not been approved!"
Exit Sub
End If

If Val(pd_amount(0)) <> 0 Then
pass = True
End If

endOfFile:

rs.Close




query = "INSERT INTO loan_release (date_release, ledger, percent,  balance, amount, area_name, client_name) values" _
& "('" & date_new_release.Value & "', '" & ledger_1(0) & "','" & percent_new_release.Text & "','" & (Val(Amount_new_release.Text) * Val(percent_new_release.Text)) - Val(pd_amount(0)) & "','" & Val(Amount_new_release.Text) * Val(percent_new_release.Text) & "','" & Trim(ledger_1(3)) & "','" & Trim(ledger_1(1)) & ", " & Trim(ledger_1(2)) & "')"

conn.Execute query

query = "update client set balance = '" & (Val(Amount_new_release.Text) * Val(percent_new_release.Text)) - Val(pd_amount(0)) & "' where ledger = '" & ledger_1(0) & "'"
    conn.Execute query



' for pass due accounts   ---------------------------------

If pass = True Then

    Dim area_split() As String
    Dim month As String
    Dim days As String
    Dim split_date() As String

    area_split = Split(Trim(ledger_1(3)), " ")
    split_date = Split(date_new_release.Value, "/")

    days = split_date(1)
  
    month = split_date(0)
    
' ------------------ insert into area

    query = "insert into " & LCase(Trim(area_split(0))) & Trim(area_split(1)) & " (area_name,  ledger, balance, d_c, days_left, release_number,date_paid) " _
    & " values ('" & Trim(ledger_1(3)) & "','" & ledger_1(0) & "','0','" & pd_amount(0) & "','" & DateDiff("d", date_new_release.Value, pd_amount(2)) & "','" & pd_amount(1) & "','" & month & "/" & days & "/" & split_date(2) & "')"
    conn.Execute query

    '  ------------------  insert into loan_release


    query = "update loan_release set remarks = '' , balance = '0' , days_left = '" & DateDiff("d", date_new_release.Value, pd_amount(2)) & "' where release_number = '" & pd_amount(1) & "'"
    conn.Execute query

    query = "update client set balance = '" & (Val(Amount_new_release.Text) * Val(percent_new_release.Text)) - Val(pd_amount(0)) & "' where ledger = '" & ledger_1(0) & "'"
    conn.Execute query


End If








' printing starts ----------


query = "select release_number from loan_release where  ledger = '" & Trim(ledger_1(0)) & "' order by release_number desc limit 1"

rs.Open query, conn


For Each fld In rs.Fields
rn_number = fld.Value
Next

rs.Close


query = "select address from client where  ledger = '" & Trim(ledger_1(0)) & "'"

rs.Open query, conn

For Each fld In rs.Fields
addrss = fld.Value
Next

print_string = Amount_new_release.Text & " * " & rn_number & " * " & Trim(ledger_1(1)) & ", " & Trim(ledger_1(2)) & " * " & addrss & " * " & percent_new_release.Text & " * " & Round(((Val(Amount_new_release.Text) * Val(percent_new_release.Text)) / 57), 2) & " * " & _
                date_new_release.Value & " * " & Trim(ledger_1(0)) & " * " & Val(pd_amount(0)) & " * " & (Val(Amount_new_release.Text) - Val(pd_amount(0)) - 20)

Open "print.dll" For Output As #1
Print #1, print_string
Close "1"

Shell ("printing.exe")
' printing ends --------------

return_home
MsgBox "Release successfully made!"
Exit Sub
error_saving:
MsgBox "Error saving. Please fill up the form correctly!", vbExclamation

Exit Sub

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If active = "help" Then
help1.Visible = False
Else
help1.Visible = True
End If

If active = "printPreview" Then
printPreview1.Visible = False
Else
printPreview1.Visible = True
End If


logOut1.Visible = True
min1.Visible = True
close1.Visible = True
End Sub

Private Sub Label2_Click()
If search.Text = "" Then
search.Text = "Customer Search"
search.FontItalic = True
search.ForeColor = &H808080
End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pindot = True
x_1 = X
y_1 = Y
pindot = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x2 As Integer
Dim y2 As Integer
If pindot = True Then
x2 = x_1 - X
y2 = y_1 - Y
Main.Left = Main.Left - x2
Main.Top = Main.Top - y2
End If
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pindot = False
End Sub

Private Sub logOut1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
logOut1.Visible = False
End Sub

Private Sub logout2_Click()
Unload Me
LoginForm.Show
End Sub



Private Sub min1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
min1.Visible = False
End Sub

Private Sub min2_Click()
Main.WindowState = 1
End Sub

Private Sub name_new_release_Change()

If name_new_release.Text = "Enter the customer's ledger/name to maker release" Then
Exit Sub
End If

new_rel_list.Visible = True

new_rel_list.Clear

Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim client_content(7) As String

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "select * from client where client_name like '" & name_new_release.Text & "%'", conn


Do Until rs.EOF
client_counter = 0
For Each fld In rs.Fields
client_content(client_counter) = fld.Value
client_counter = client_counter + 1
Next

new_rel_list.AddItem client_content(0) & ", " & client_content(1) & ", " & client_content(5)

rs.MoveNext
Loop
rs.Close

rs.Open "select * from client where ledger like '" & name_new_release.Text & "%'", conn


Do Until rs.EOF
client_counter = 0
For Each fld In rs.Fields
client_content(client_counter) = fld.Value
client_counter = client_counter + 1
Next

new_rel_list.AddItem client_content(0) & ",      " & client_content(1) & ", " & client_content(5)

rs.MoveNext
Loop

End Sub

Private Sub name_new_release_Click()
name_new_release.Text = ""
End Sub

Private Sub name_update_release_GotFocus()
MsgBox "Name will be automatically filled up when you type the Release Number." & vbNewLine & vbNewLine _
        & "You can't alter the customer's name.", vbExclamation

End Sub

Private Sub new_re_Click()

hide_all
release1.Visible = False
show_newRelease
End Sub



Private Sub new_rel_list_Click()
name_new_release.Text = new_rel_list.list(new_rel_list.ListIndex)
new_rel_list.Visible = False
End Sub

Private Sub new_releaseFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
release1.Visible = False
End Sub

Private Sub percent_new_release_Click()
percent_new_release.Text = ""
End Sub


Private Sub print_preview_selectArea_Click()
print_preview_selectArea.Locked = False
a = list_query(print_preview_selectArea.Text, print_preview_selescttask.Text)
End Sub

Private Function list_query(area As String, task As String)
PrintPreviewGrid.Rows = 1
Dim query_value As String
Dim query2 As String

If task = "Preview Past Due" Then
PrintPreviewGrid.TextMatrix(0, 4) = "Due Days"
PrintPreviewGrid.ColWidth(0) = 500
PrintPreviewGrid.ColWidth(1) = 900
PrintPreviewGrid.ColWidth(2) = 3500
PrintPreviewGrid.ColWidth(3) = 1400
PrintPreviewGrid.ColWidth(4) = 20006
query_value = "select * from loan_release where area_name ='" & area & "' and remarks = 'Past due'"
ElseIf task = "Prview Master List" Then
PrintPreviewGrid.TextMatrix(0, 4) = "Days Left"
query_value = "select * from client where area_name = '" & area & "'"
PrintPreviewGrid.ColWidth(0) = 500
PrintPreviewGrid.ColWidth(1) = 900
PrintPreviewGrid.ColWidth(2) = 4100
PrintPreviewGrid.ColWidth(3) = 1900
PrintPreviewGrid.ColWidth(4) = 20006
Else
Exit Function
End If

Dim conn As ADODB.Connection
Dim config As String
Dim q_value(20) As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open query_value, conn
b = 0
Do Until rs.EOF
b = b + 1
PrintPreviewGrid.Rows = b + 1
a = 0
    For Each fld In rs.Fields
     q_value(a) = fld.Value
     a = a + 1
    Next



If task = "Preview Past Due" Then
PrintPreviewGrid.TextMatrix(b, 4) = "   " & Abs(Val(q_value(4)))
PrintPreviewGrid.TextMatrix(b, 1) = q_value(0)
PrintPreviewGrid.TextMatrix(b, 2) = q_value(7)
PrintPreviewGrid.TextMatrix(b, 3) = "Php. " & q_value(6)
End If

If task = "Prview Master List" Then
PrintPreviewGrid.TextMatrix(b, 4) = "   " & Abs(Val(q_value(4)))
PrintPreviewGrid.TextMatrix(b, 1) = q_value(0)
PrintPreviewGrid.TextMatrix(b, 2) = q_value(1)
PrintPreviewGrid.TextMatrix(b, 3) = "Php. " & q_value(6)
End If

rs.MoveNext
Loop

print_preview_counter.Text = b

End Function

Private Sub print_preview_selectArea_KeyDown(KeyCode As Integer, Shift As Integer)
print_preview_selectArea.Locked = True
End Sub

Private Sub print_preview_selectArea_KeyUp(KeyCode As Integer, Shift As Integer)
print_preview_selectArea.Locked = False
End Sub

Private Sub print_preview_selescttask_Click()
a = list_query(print_preview_selectArea.Text, print_preview_selescttask.Text)
print_preview_selescttask.Locked = False
End Sub

Private Sub print_preview_selescttask_KeyDown(KeyCode As Integer, Shift As Integer)
print_preview_selescttask.Locked = True
End Sub

Private Sub print_preview_selescttask_KeyUp(KeyCode As Integer, Shift As Integer)
print_preview_selescttask.Locked = False
End Sub

Private Sub printPreview1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
printPreview1.Visible = False


End Sub

Private Sub printpreview2_Click()
hide_all
show_print

End Sub



Private Sub rel_no_update_rel_Click()
rel_no_update_rel.Text = ""
End Sub

Private Sub rel_no_update_rel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim rs2 As ADODB.Recordset


Set rs = New ADODB.Recordset
Set rs2 = New ADODB.Recordset


Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "select client_name from loan_release where release_number = '" & rel_no_update_rel.Text & "'", conn
rs2.Open "select date_approve from loan_release where release_number = '" & rel_no_update_rel.Text & "' and date_approve != ''", conn

If rs2.EOF = False Then
For Each fld In rs2.Fields
MsgBox "The release is already approved on " & fld.Value & ".", vbExclamation

Next
Exit Sub
End If


If rs.EOF = True Then
MsgBox "Release number was not found!", vbExclamation

Exit Sub
End If

For Each fld In rs.Fields
name_update_release.Text = fld.Value
Next

End If
End Sub

Private Sub release1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If active <> "release" Then
release1.Visible = False
End If

If active = "clientOption" Then
clientOption1.Visible = False
Else
clientOption1.Visible = True
End If

If active = "remittance" Then
remittance1.Visible = False
Else
remittance1.Visible = True
End If

End Sub

Private Sub release2_Click()
hide_all
show_release
End Sub

Private Sub remAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
save_remittance
remAmount.Text = ""
remAmount.SetFocus
End If
End Sub

Private Sub remittance1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If active <> "remittance" Then
remittance1.Visible = False
End If

If active = "release" Then
release1.Visible = False
Else
release1.Visible = True
End If

If active = "viewRemittance" Then
viewRemittance1.Visible = False
Else
viewRemittance1.Visible = True
End If
End Sub

Private Sub remittance2_Click()
hide_all
show_remittance
End Sub

Private Sub RemLedger_Click()
If position.Caption <> "Administrator" Then
RemLedger.Locked = True
MsgBox "Only Administrators can have an individual client remittance.", vbExclamation
Else
RemLedger.Locked = False
End If
End Sub

Private Sub sav_Click()

If AreaSelectAreaOP.Text = "Select Area" Then
MsgBox "Please select an area first before doing some task!"
Exit Sub
End If

If collectorAreaOp.Text = "" Or areaLocAreaOp.Text = "" Then
GoTo error_saving
End If

ans = (MsgBox("Area you sure you want to save changes?", vbYesNoCancel))

If ans = 2 Then
return_home
Exit Sub
End If

If ans = 7 Then
Exit Sub
End If

On Error GoTo error_saving
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


query = "update area set collector = '" & collectorAreaOp.Text & "' where area_id = '" & areaopblanck.Text & "'"
conn.Execute query

MsgBox "Saving successfully done!"
return_home
Exit Sub
error_saving:
MsgBox "Error saving data. Please fill up the form correctly!", vbExclamation

End Sub

Private Sub savecOp_Click()

'count_me = 0
'For Each DataField In clientOption

'    If DataField = TextBox And DataField = "" And count_me <> 1 Then
    'MsgBox DataField & " " & count_me
'    GoTo error_saving
'    End If
'    count_me = count_me + 1
'Next

If COpLname.Text = "" Or COpIn.Text = "" Or COpnumber.Text = "" Or COpAdd.Text = "" Then
GoTo error_saving
End If

ans = (MsgBox("Area you sure you want to save changes?", vbYesNoCancel))

If ans = 2 Then
return_home
Exit Sub
End If

If ans = 7 Then
Exit Sub
End If

On Error GoTo error_saving
Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim client_content(7) As String

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

query = "update client set client_name = '" & COpLname.Text & "', address = '" & COpAdd.Text & "', investment = '" & COpIn.Text & "', area_name = '" & SelectAreaClientOp.Text & "', mobile_number = '" & COpnumber.Text & "' where ledger = '" & COpLedger.Text & "'"
conn.Execute query

MsgBox "Saving successfully done!"
return_home
Exit Sub
error_saving:
MsgBox "Error saving data. Please fill up the form correctly!", vbExclamation

End Sub

Private Sub saveRel_Click()
save_remittance
End Sub
Private Sub save_remittance()



If SelectAreaRemittance.Text = "Select Area" Then
MsgBox "Please select an area first!", vbExclamation

Exit Sub
End If

If RemLedger = "" Then
MsgBox "No client to update!", vbExclamation
Exit Sub
End If

remittance_text1.Text = Val(remittance_text1.Text) + 1

remTotalCollector.Text = Val(remTotalCollector.Text) + Val(remAmount.Text)


Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim content(20) As String
Dim client_content(20) As String
Dim area_spliter() As String
Dim new_area As String
Dim month As String
Dim days As String
Dim split_date() As String
Dim rem_date As String

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open




'rs.Open "select * from client where ledger = '" & Text3.Text & "'", conn


'Do Until rs.EOF
'a = 0
'For Each fld In rs.Fields
'client_content(a) = fld.Value
'a = a + 1
'Next
'rs.MoveNext
'Loop

'rs.Close

split_date = Split(remittance_date.Value, "/")

'If Len(split_date(1)) = 1 Then
'days = "0" & split_date(1)
'Else
days = split_date(1)
'End If

'If Len(split_date(0)) = 1 Then
'month = "0" & split_date(0)
'Else
month = split_date(0)
'End If


rs.Open "select * from loan_release where ledger = '" & RemLedger.Text & "' and area_name = '" & SelectAreaRemittance.Text & "' order by release_number desc limit 1", conn


Do Until rs.EOF
a = 0
For Each fld In rs.Fields
content(a) = fld.Value
a = a + 1
Next
rs.MoveNext
Loop

rs.Close


rs.Open "select area_id from area where area_name = '" & SelectAreaRemittance.Text & "'", conn

Do Until rs.EOF
For Each fld In rs.Fields
area_id = fld.Value
Next
rs.MoveNext
Loop


area_spliter = Split(Trim(SelectAreaRemittance.Text), " ")
new_area = LCase(area_spliter(0)) & area_spliter(1)

If content(8) = "" Or (Val(content(6)) - Val(remAmount.Text)) <= 0 Then
remAmount.Text = "0"
End If

If content(2) <> "" Then
rem_date = DateDiff("d", remittance_date.Value, content(2))
Else
rem_date = 0
query = "insert into " & new_area & " (ledger,area_name,d_c,date_paid) " _
& " values ('" & RemLedger.Text & "','" & SelectAreaRemittance.Text & "','" & remAmount.Text & "','" & month & "/" & days & "/" & split_date(2) & "')"
conn.Execute query

GoTo not_release
End If

query = "insert into " & new_area & " (area_name,  ledger, balance, d_c, days_left, release_number,date_paid) " _
& " values ('" & SelectAreaRemittance.Text & "','" & RemLedger.Text & "','" & Val(content(6)) - Val(remAmount.Text) & "','" & remAmount.Text & "','" & rem_date & "','" & content(8) & "','" & month & "/" & days & "/" & split_date(2) & "')"


conn.Execute query


query2 = "update loan_release set days_left = '" & rem_date & "', balance = '" & Val(content(6)) - Val(remAmount.Text) & "' where release_number = '" & content(8) & "'"
conn.Execute query2

query3 = "update client set balance = '" & Val(content(6)) - Val(remAmount.Text) & "' where area_name = '" & SelectAreaRemittance.Text & "' and ledger = '" & RemLedger.Text & "'"
conn.Execute query3


If (Val(rem_date) - 1) < 0 Then
conn.Execute "update loan_release set remarks = 'Past due' where release_number = '" & content(8) & "'"
Else
conn.Execute "update loan_release set remarks = '' where release_number = '" & content(8) & "'"
End If
remAmount.Text = ""

RemLedger.Text = remittance_blanck_list.list(remittance_text1.Text)
Exit Sub

not_release:
remAmount.Text = ""
RemLedger.Text = remittance_blanck_list.list(remittance_text1.Text)
Exit Sub

End Sub

Private Sub search_Change()
If search.Text <> "Customer Search" Then
search_show
End If

If search.Text <> "" Then
Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim client_content(7) As String

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "select * from client where client_name like '" & search.Text & "%'", conn


b = 0

Do Until rs.EOF
b = b + 1
CostumerSearchGrid.Rows = b + 1

a = 0
For Each fld In rs.Fields
client_content(a) = fld.Value
a = a + 1
Next


CostumerSearchGrid.TextMatrix(b, 1) = client_content(5)
CostumerSearchGrid.TextMatrix(b, 2) = client_content(0)
CostumerSearchGrid.TextMatrix(b, 3) = client_content(1)

rs.MoveNext

Loop
CostumerSearchGrid.Sort = 1

End If

End Sub

Private Sub search_Click()
If search.Text = "Customer Search" Then
search.Text = ""
search.ForeColor = vbBlack
search.FontItalic = False
search_show
End If
End Sub

Private Sub search_GotFocus()
If search.Text = "Customer Search" Then
search_show
search.Text = ""
search.ForeColor = vbBlack
search.FontItalic = False
End If
End Sub

Private Sub search_KeyDown(KeyCode As Integer, Shift As Integer)
If search.Text = "Customer Search" Then
search.ForeColor = vbBlack
search.FontItalic = False
search.Text = ""
End If
End Sub

Private Sub search_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 And search.Text = "" Then
search.Text = "Customer Search"
search.FontItalic = True
search.ForeColor = &H808080
End If
End Sub

Private Sub search_LostFocus()
If search.Text = "" Then
search.Text = "Customer Search"
search.FontItalic = True
search.ForeColor = &H808080
End If
End Sub



Private Sub SelectAreaClientOp_GotFocus()
MsgBox "Area Name/Location can't be change due to previous transactions."
End Sub

Private Sub SelectAreaRemittance_Click()
remittance_blanck_list.Clear
remittance_text6.Text = 1

Dim conn As ADODB.Connection
Dim config As String
Dim rsa As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim fld As ADODB.Field



Set rsa = New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rsa.Open "select maturity from loan_release where maturity = '' and area_name = '" & SelectAreaRemittance.Text & "'", conn

If rsa.EOF = False Then
MsgBox "Warning! A release has not been approve on this area!", vbCritical
End If

rs.Open "select ledger from client where area_name like '" & SelectAreaRemittance.Text & "%'", conn
rs2.Open "select collector from area where area_name like '" & SelectAreaRemittance.Text & "%'", conn

For Each fld In rs2.Fields
remCollector.Text = fld.Value
Next

Do Until rs.EOF

For Each fld In rs.Fields
remittance_blanck_list.AddItem fld.Value
Next
rs.MoveNext
remittance_text6.Text = Val(remittance_text6.Text) + 1
Loop


remittance_text1.Text = 0
RemLedger.Text = remittance_blanck_list.list(remittance_text1.Text)

End Sub

Private Sub SelectAreaViewR_Click()
exe_prog
End Sub

Private Sub term_update_release_Click()
term_update_release.Locked = False
End Sub

Private Sub term_update_release_KeyDown(KeyCode As Integer, Shift As Integer)
term_update_release.Locked = True
End Sub

Private Sub term_update_release_KeyUp(KeyCode As Integer, Shift As Integer)
term_update_release.Locked = False
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii
End Sub

Private Sub Text8_Change()

End Sub

Private Sub Timer1_Timer()
Dim today As Variant
today = Format(Now, "dddd,  " & "mmmm  d, yyyy" & "           hh : mm ampm")
lbltime.Caption = today
End Sub

Private Sub Timer2_Timer()
Main.Top = Main.Top - 500
Main.Width = Main.Width - 300

If Main.Width < 300 Or Main.Height < 300 Then
Timer2.Enabled = False
Unload Me
End If
End Sub



Private Sub UpdateRelease_Click()
If Val(day_update_release.Text) = 0 Then
GoTo error_saving
End If

If term_update_release.Text = "Select Term" Then
GoTo error_saving
End If

If amount_update_release.Text = "" Then
GoTo error_saving
End If

If place_update_release.Text = "" Then
GoTo error_saving
End If


If page_update_reelase.Text = "" Then
GoTo error_saving
End If


If book_no_update_release.Text = "" Then
GoTo error_saving
End If



If series_update_release.Text = "" Then
GoTo error_saving
End If



If rel_no_update_rel.Text = "" Then
GoTo error_saving
End If


If name_update_release.Text = "" Then
GoTo error_saving
End If
days = DateAdd("d", day_update_release.Text, date_update_release.Value)
Dim name() As String

Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim content(5) As String
Dim due_amount As Double
Dim real_amount As Double


Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

query = "select amount, balance, percent, ledger from loan_release where release_number = '" & rel_no_update_rel.Text & "'"
rs.Open query, conn

a = 0
For Each fld In rs.Fields
content(a) = fld.Value
a = a + 1
Next


due_amount = Val(content(0)) - Val(content(1))

real_amount = (Val(amount_update_release.Text) * Val(content(2))) - due_amount

If due_amount = 0 Then


query = "update client set balance = '" & real_amount & "' where ledger ='" & content(3) & "'"
conn.Execute query

query = " update loan_release set maturity = '" & days & "', date_approve = '" & date_update_release.Value & "', payment_term = '" & term_update_release.Text _
        & "', days_left = '" & day_update_release.Text & "', place_issued = '" _
        & place_update_release.Text & "', page_number = '" & page_update_reelase.Text & "', book_number = '" & book_no_update_release.Text & "', series_of = '" & series_update_release.Text & "', amount = '" & real_amount _
        & "', balance = '" & real_amount & "' where release_number = '" & rel_no_update_rel.Text & "'"
        
conn.Execute query

Else


query = " update loan_release set maturity = '" & days & "', date_approve = '" & date_update_release.Value & "', payment_term = '" & term_update_release.Text _
        & "', days_left = '" & day_update_release.Text & "', place_issued = '" _
        & place_update_release.Text & "', page_number = '" & page_update_reelase.Text & "', book_number = '" & book_no_update_release.Text & "', series_of = '" & series_update_release.Text & "', amount = '" & (Val(amount_update_release.Text) * Val(content(2))) _
        & "', balance = '" & real_amount & "' where release_number = '" & rel_no_update_rel.Text & "'"

conn.Execute query

query = "update client set balance = '" & real_amount & "' where ledger ='" & content(3) & "'"

conn.Execute query

End If

return_home
Exit Sub

error_saving:
MsgBox "Error updating loan release. Please fi up the form correctly!", vbExclamation
Exit Sub

End Sub

Private Sub upRel_Click()
hide_all
release1.Visible = False
show_updateRelease
End Sub



Private Sub userOpsave_Click()

If UserOpName.Text = "" Or UserOpUsername.Text = "" Or UserOpPass.Text = "" Or UserDesigntion.Text = "Select designation" Or UserOptask = "Select Task" Then
MsgBox "Saving error, please fill up the form correctly!", vbExclamation
Exit Sub
End If

If Len(UserOpPass.Text) < 6 Then
MsgBox "Password mass be more than or equal to 6 characters.", vbExclamation
Exit Sub
End If

If (UserOpPass.Text <> UserOpConfirmPass.Text) Then
MsgBox "Password did not match!", vbExclamation
Exit Sub
End If

If UserOptask.Text = "Select Task" Then
MsgBox "Please select a Task!", vbExclamation
Exit Sub
End If




Dim ans1 As Double
Dim ans2 As Double
Dim ans3 As Double
Dim ans4 As Double
Dim ans5 As Double
Dim ans6 As Double
Dim password_text As Double


ans1 = Asc(UserOpPass.Text)
ans2 = AscB(StrReverse(UserOpPass.Text))
ans3 = Asc(Left(UserOpPass.Text, 3))
ans4 = AscB(StrReverse(Left(UserOpPass.Text, 2)))
ans5 = Asc(Right(UserOpPass.Text, 3))
ans6 = AscB(StrReverse(Right(UserOpPass.Text, 4)))

password_text = (ans1 + ans2 + ans3 + ans4 + ans5 + ans6) * (ans1 * ans2 * ans3 * ans4 * ans5 * ans6)




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


If UserOptask.Text = "Add User" Then

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "select user_name from users where user_name = '" & UserOpUsername & "'", conn

If rs.EOF = True Then
Else
MsgBox "User name already in use!", vbExclamation
Exit Sub
End If



query = "insert into users (user_full_name, user_name,user_password,designation)" _
        & "values ('" & UserOpName.Text & "','" & UserOpUsername.Text & "','" & password_text & ":" & (Len(UserOpPass.Text) * 144) & "','" & UserDesigntion.Text & "')"
End If

If UserOptask.Text = "Edit User" Then
query = "update users set user_full_name = '" & UserOpName.Text & "', user_name = '" _
        & UserOpUsername.Text & "', user_password ='" & password_text & ":" & (Len(UserOpPass.Text) * 144) & "', designation = '" & UserDesigntion.Text & "' where user_id = '" & userOp_blanck.Text & "'"
        
End If

conn.Execute query
If UserOptask.Text = "Add User" Then
MsgBox "New user was successfully added!"
End If

If UserOptask.Text = "Edit User" Then
MsgBox "User updated!"
End If

return_home

End Sub

Private Sub UserOpSelName_Click()
'MsgBox user_grid.TextMatrix(UserOpSelName.ListIndex + 1, 1)

UserOpName.Text = user_grid.TextMatrix(UserOpSelName.ListIndex + 1, 1)
UserOpUsername.Text = user_grid.TextMatrix(UserOpSelName.ListIndex + 1, 2)
'Text3.Text = user_grid.TextMatrix(UserOpSelName.ListIndex + 1, 3)
UserDesigntion.Text = user_grid.TextMatrix(UserOpSelName.ListIndex + 1, 4)
userOp_blanck.Text = user_grid.TextMatrix(UserOpSelName.ListIndex + 1, 5)
UserOpPass.Text = ""
UserOpConfirmPass.Text = ""
End Sub

Private Sub UserOptask_Click()
If UserOptask.Text = "Edit User" Then

Dim conn As ADODB.Connection
Dim config As String
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim content(10) As String


Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "select * from users", conn

b = 0
Do Until rs.EOF
b = b + 1
user_grid.Rows = b + 1
a = 0
For Each fld In rs.Fields
content(a) = fld.Value
a = a + 1
Next

user_grid.TextMatrix(b, 1) = content(1)
user_grid.TextMatrix(b, 2) = content(2)
user_grid.TextMatrix(b, 3) = content(3)
user_grid.TextMatrix(b, 4) = content(4)
user_grid.TextMatrix(b, 5) = content(0)
UserOpSelName.AddItem content(1)
rs.MoveNext
Loop


Else

UserOpSelName.Clear
UserOpSelName.Text = "Select Name"

End If

UserOpName.Text = ""
UserOpUsername.Text = ""
UserOpPass.Text = ""
UserOpConfirmPass.Text = ""
UserDesigntion.Text = "Select designation"

End Sub

Private Sub userOption1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If active <> "userOption" Then
userOption1.Visible = False
End If

If active = "viewRemittance" Then
viewRemittance1.Visible = False
Else
viewRemittance1.Visible = True
End If

End Sub

Private Sub useroption2_Click()

If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

hide_all
show_userOption
End Sub

Private Sub view_remittance_blanck_Change()
exe_prog
End Sub

Private Sub viewRemittance1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If active <> "viewRemittance" Then
viewRemittance1.Visible = False
End If

If active = "remittance" Then
remittance1.Visible = False
Else
remittance1.Visible = True
End If

If active = "userOption" Then
userOption1.Visible = False
Else
userOption1.Visible = True
End If

End Sub

Private Sub viewRemittance2_Click()
hide_all
show_viewRemittance

End Sub

Private Sub hide_all()

homeFrame.Left = 30000
homeLabel1.Left = 30000
homeLabel2.Left = 30000

addAreaframe.Left = 100000
addareaLabel.Left = 100000

addClientFrame.Left = 30000
addclientlabel.Left = 30000

areaOptionframe.Top = 30000
ArOplabel.Left = 30000

clientOpframe.Left = 30000

releaseFrame.Top = 30000
releaselabel1.Left = 30000
releaselabel2.Left = 30000

remittanceframe.Left = 30000
Remittancelabel.Left = 30000

userOptionFrame.Top = 30000
useroptionlabel.Top = 30000

new_releaseFrame.Top = 30000
new_release_label.Top = 30000

update_release_label.Top = 30000
UpdateReleaseFrame.Top = 30000

view_remittance_label.Top = 30000
ViewRemittanceFrame.Top = 30000

PrintpreviewFrame.Top = 30000
PrintPreviewLabel.Top = 30000

CostumerSearchFrame.Top = 30000
costumersearchLabel.Top = 30000

developerFrame.Top = 30000
developerLabel.Top = 30000

help_frame.Top = 30000
help_label.Top = 30000

home1.Visible = True
addClient1.Visible = True
addArea1.Visible = True
areaOption1.Visible = True
clientOption1.Visible = True
release1.Visible = True
remittance1.Visible = True
viewRemittance1.Visible = True
userOption1.Visible = True
help1.Visible = True
printPreview1.Visible = True
developers1.Visible = True
notice.Visible = True

End Sub
Private Sub show_home()
load_area
load_query
load_back_up
active = "home"
activeLabel = "Welcome, " & login_owner.Caption & "!"
home1.Visible = False

homeFrame.Top = 2700
homeFrame.Left = 3100
homeLabel1.Top = 2080
homeLabel1.Left = 3840
homeLabel2.Top = 2350
homeLabel2.Left = 3840
End Sub
Private Sub show_area()

If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

active = "addArea"
activeLabel = "Add Area"
addArea1.Visible = False

addAreaframe.Top = 2700
addAreaframe.Left = 3100
addareaLabel.Left = 3840
addareaLabel.Top = 2205
End Sub

Private Sub show_client()

If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If


active = "addClient"
activeLabel = "Add Client"
addClient1.Visible = False

addClientFrame.Top = 2700
addClientFrame.Left = 3100
addclientlabel.Left = 3840
addclientlabel.Top = 2205

addclientFname.Text = ""
addclientMname.Text = ""
addclientLname.Text = ""
addclientAddress.Text = ""
addclientInvestment.Text = ""
addclientNumber.Text = ""
addclientLegder.Text = ""
addclientArea.Text = "Select Area"

addclientFname.SetFocus


End Sub

Private Sub show_areaOption()

If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

active = "areaOption"
activeLabel = "Area Option"
areaOption1.Visible = False

areaOptionframe.Top = 2700
areaOptionframe.Left = 3100
ArOplabel.Left = 3840
ArOplabel.Top = 2205

AreaSelectAreaOP.Text = "Select Area"
areaLocAreaOp.Text = ""
collectorAreaOp.Text = ""
End Sub

Private Sub show_ClientOption()


If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If


active = "clientOption"
activeLabel = "Client Option"
clientOption1.Visible = False

clientOpframe.Top = 2700
clientOpframe.Left = 3100
ArOplabel.Left = 3840
ArOplabel.Top = 2205

COpName.Text = "Enter the name of person to be search."
COpLname.Text = ""
COpAdd.Text = ""
COpIn.Text = ""
COpnumber.Text = ""
SelectAreaClientOp.Text = ""
COpLedger.Text = ""
clientOpList.Visible = False
End Sub

Private Sub show_release()
active = "release"
activeLabel = "Loan Release"
release1.Visible = False

releaseFrame.Top = 2700
releaseFrame.Left = 3100
releaselabel1.Top = 2080
releaselabel1.Left = 3840
releaselabel2.Top = 2350
releaselabel2.Left = 3840


End Sub

Private Sub show_remittance()
active = "remittance"
activeLabel = "Daily Remittance"
remittance1.Visible = False

remittanceframe.Top = 2700
remittanceframe.Left = 3100
Remittancelabel.Top = 2080
Remittancelabel.Left = 3840
remittance_date.Value = Format(Now, "mm  dd yyyy")

remAmount.Text = ""
RemLedger.Text = ""
SelectAreaRemittance.Text = "Select Area"
remCollector.Text = ""
remTotalCollector.Text = "0.00"
End Sub

Private Sub show_viewRemittance()
active = "viewRemittance"
activeLabel = "View Remittnace"
viewRemittance1.Visible = False

ViewRemittanceFrame.Top = 2700
ViewRemittanceFrame.Left = 3100
view_remittance_label.Top = 2205
view_remittance_label.Left = 3840

date_ViewRemiottance.Value = Format(Now, "  mm d yyyy")

End Sub

Private Sub show_userOption()

If Main.position.Caption = "Encoder" Then
MsgBox "This functionality is for Administrators only.", vbInformation
return_home
Exit Sub
End If

active = "userOption"
activeLabel = "User Option"
userOption1.Visible = False

userOptionFrame.Top = 2700
userOptionFrame.Left = 3100
useroptionlabel.Left = 3840
useroptionlabel.Top = 2205

UserOpName.Text = ""
UserOpUsername.Text = ""
UserOpPass.Text = ""
UserOpConfirmPass.Text = ""
UserDesigntion.Text = "Select designation"
UserOpSelName.Text = "Select Name"

End Sub

Private Sub search_show()
hide_all
'notice.Visible = False
activeLabel = "Customer Search"
active = "search"

CostumerSearchFrame.Top = 2700
CostumerSearchFrame.Left = 3100
costumersearchLabel.Left = 3840
costumersearchLabel.Top = 2205
End Sub

Private Sub help_show()
active = "help"
activeLabel = "System Help"
help1.Visible = False

help_frame.Top = 2700
help_frame.Left = 3100
help_label.Left = 3840
help_label.Top = 2205

End Sub

Private Sub show_print()
active = "printPreview"
activeLabel = "Preview and Print"
printPreview1.Visible = False

PrintpreviewFrame.Top = 2700
PrintpreviewFrame.Left = 3100
PrintPreviewLabel.Top = 2205
PrintPreviewLabel.Left = 3840

End Sub

Private Sub show_dev()
active = "developers"
activeLabel = "Developer's Profile"
developers1.Visible = False


developerFrame.Top = 2700
developerFrame.Left = 3100
developerLabel.Top = 2205
developerLabel.Left = 3840
End Sub

Private Sub show_newRelease()
new_releaseFrame.Top = 2700
new_releaseFrame.Left = 3100
new_release_label.Top = 2205
new_release_label.Left = 3840
activeLabel = "New Release"

new_rel_list.Visible = False
name_new_release.Text = "Enter the customer's ledger/name to maker release"
Amount_new_release.Text = ""
percent_new_release.Text = "Value in percent"
End Sub

Private Sub show_updateRelease()
UpdateReleaseFrame.Top = 2700
UpdateReleaseFrame.Left = 3100
update_release_label.Top = 2205
update_release_label.Left = 3840
activeLabel = "Update Release"

date_update_release.Value = Format(Now, "mm d yyyy")
Dim date_2() As String
date_2 = Split(date_update_release.Value, "/")
updateRel_blanck.Text = date_2(0) & "-" & date_2(1) & "-" & date_2(2)

day_update_release.Text = "57"
term_update_release.Text = "Select Term"
amount_update_release.Text = ""
place_update_release.Text = ""
page_update_reelase.Text = ""
book_no_update_release.Text = ""
series_update_release.Text = ""
rel_no_update_rel.Text = "Enter Release Number"
name_update_release.Text = ""

End Sub

Private Sub return_home()
hide_all
show_home
End Sub

Private Sub viewRemittanceGrid_DblClick()

Dim adlaw As Date
adlaw = Format(Now, "mm/d/yyyy")

If (adlaw <> date_ViewRemiottance.Value) And position.Caption <> "Administrator" Then
MsgBox "Only Admin can edit the previous remittance.", vbExclamation
'Exit Sub
End If

If viewRemittanceGrid.TextMatrix(viewRemittanceGrid.Row, 1) = "" Then
Exit Sub
End If

EditRemittance.edit_remittnace_ledger.Text = viewRemittanceGrid.TextMatrix(viewRemittanceGrid.Row, 1)
EditRemittance.Edit_remittance_name.Text = viewRemittanceGrid.TextMatrix(viewRemittanceGrid.Row, 2)
EditRemittance.Text8.Text = SelectAreaViewR.Text
EditRemittance.Text4.Text = viewRemittanceGrid.TextMatrix(viewRemittanceGrid.Row, 4)
EditRemittance.Text6.Text = viewRemittanceGrid.TextMatrix(viewRemittanceGrid.Row, 6)
EditRemittance.Text7.Text = viewRemittanceGrid.TextMatrix(viewRemittanceGrid.Row, 7)
EditRemittance.Text9.Text = viewRemittanceGrid.TextMatrix(viewRemittanceGrid.Row, 3)
EditRemittance.edit_remittance_dc = viewRemittanceGrid.TextMatrix(viewRemittanceGrid.Row, 3)
EditRemittance.date_paid.Text = viewRemittanceGrid.TextMatrix(viewRemittanceGrid.Row, 8)
EditRemittance.update_date.Value = viewRemittanceGrid.TextMatrix(viewRemittanceGrid.Row, 8)


EditRemittance.Show
End Sub


Private Sub load_area()

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open

rs.Open "SELECT area_name FROM area ", conn

AreaSelectAreaOP.Clear
addclientArea.Clear
SelectAreaRemittance.Clear
SelectAreaViewR.Clear
print_preview_selectArea.Clear

Do Until rs.EOF

For Each fld In rs.Fields
AreaSelectAreaOP.AddItem fld.Value
addclientArea.AddItem fld.Value
'SelectAreaClientOp.AddItem fld.Value
SelectAreaRemittance.AddItem fld.Value
SelectAreaViewR.AddItem fld.Value
print_preview_selectArea.AddItem fld.Value
Next
rs.MoveNext
Loop
SelectAreaViewR.Text = "Select Area"
print_preview_selectArea.Text = "Select Area"
AreaSelectAreaOP.Text = "Select Area"
addclientArea.Text = "Select Area"
SelectAreaRemittance.Text = "Select Area"
End Sub


Private Sub load_query()

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection

Open "config.txt" For Input As #1
config = Input$(LOF(1), #1)
Close #1
conn.ConnectionString = config
conn.Open



rs.Open "SELECT ledger FROM client ", conn
counter_1 = 0
Do Until rs.EOF
counter_1 = counter_1 + 1
rs.MoveNext
Loop
client_count.Caption = counter_1

rs.Close



pass_due_counter = 0
rs.Open "SELECT remarks FROM loan_release where remarks = 'Past due'", conn

Do Until rs.EOF
pass_due_counter = pass_due_counter + 1
rs.MoveNext
Loop

' for clients -----------------

good_client.Caption = counter_1 - pass_due_counter & " clients are in good payment status."
bad_client.Caption = pass_due_counter & " clients are in Pass Due status."

rs.Close
rs.Open "SELECT area_id FROM area", conn
area_counter = 0

Do Until rs.EOF
area_counter = area_counter + 1
rs.MoveNext
Loop

' clients ends ---------------


' for total area ------------

total_area.Caption = "Total of " & area_counter & " Areas."
area_count.Caption = area_counter
ave_area.Caption = "Average of " & Round((counter_1 / area_counter), 0) & " clients per Area"

rs.Close

' area ends --------------


' for release ---------------

rs.Open "SELECT date_approve FROM loan_release order by release_number", conn

Dim not_approved(99999) As String

loan_counter = 0
unapprove_counter = 0
last_value = 0
Do Until rs.EOF
For Each fld In rs.Fields
If fld.Value = "" Then
unapprove_counter = unapprove_counter + 1
not_approved(unapprove_counter) = loan_counter + 1
Else
last_value = fld.Value
End If
Next
loan_counter = loan_counter + 1
rs.MoveNext
Loop


release_count.Caption = loan_counter

If unapprove_counter <> 0 Then

Dim msgboxstring As String
h_counter.Text = unapprove_counter
bad_release.Caption = unapprove_counter & " release has not yet approved."
release_ok1.Visible = False
release_not_ok1.Visible = True

msgboxstring = not_approved(1)
For i = 2 To unapprove_counter Step 1
msgboxstring = msgboxstring & "," & not_approved(i)
Next

holder.Text = msgboxstring

Else
release_ok1.Visible = True
release_not_ok1.Visible = False
bad_release.Caption = "All relase has been approved."
End If



last_rel.Caption = "Last release was reformed " & last_value & "."

' release ends ------------


End Sub

Private Sub make_back_up_Click()
Dim now_date As String
Open App.Path & "\back_up\back_up_config.dll" For Input As #1
config = Input$(LOF(1), #1)
Close #1

' make dir---
now_date = Format(Now, "mm-d-yyyy")
MkDir (App.Path & "\back_up\" & now_date)

'copy back up -----
Dim FSO As New FileSystemObject
Dim fsoFldr As Scripting.Folder
Dim fsoFile As Scripting.File
Dim intCounter As Integer
Set fsoFldr = FSO.GetFolder(config)
fsoFldr.Copy (App.Path & "\back_up\" & now_date)
Set fsoFldr = FSO.GetFolder(App.Path & "\back_up\" & now_date)

MsgBox "Task completed successfully.", vbInformation

back_up_not_ok.Visible = False
back_up_ok.Visible = True
make_back_up.Visible = False
back_upwarning.Caption = "Data Back Up is already performed. Back Up is updated."
End Sub


Private Sub load_back_up()
now_date = Format(Now, "mm-d-yyyy")
dir_path = App.Path & "\back_up\" & now_date
If Dir$(dir_path, vbDirectory) = "" Then
back_upwarning.Caption = "Data back up is not yet performed! Click " & Chr(34) & "Make Back Up" & Chr(34) & " now!"
back_up_not_ok.Visible = True
back_up_ok.Visible = False
make_back_up.Visible = True
Else
back_up_not_ok.Visible = False
back_up_ok.Visible = True
make_back_up.Visible = False
back_upwarning.Caption = "Data Back Up is already performed. Back Up is updated."
End If
End Sub



Private Sub print_prev_print_Click()


If print_preview_selectArea.Text = "Select Area" Or print_preview_selescttask.Text = "Select Task" Then
MsgBox "Please select Area/Task first!", vbExclamation
Exit Sub
End If

If print_preview_selescttask.Text = "Preview Past Due" Then
print_pass_due
ElseIf print_preview_selescttask.Text = "Prview Master List" Then
ans = (MsgBox("Do you want to include balance?", vbYesNoCancel))

If ans = 2 Then   '  cancel
Exit Sub
End If

If ans = 7 Then   ' no
master_no_bal
End If

If ans = 6 Then   ' yes
master_list
End If

End If


End Sub

Private Sub print_pass_due()

print_prev_status.Visible = True
Main.Enabled = False
loading.Show
Dim row_counter As Integer

Set xlwbook = xl.Workbooks.Open(App.Path & "\PRINT\pass_due.xls")
Set xlsheet = xlwbook.Sheets.Item(1)
    
        xlsheet.Cells(5, 2) = print_preview_selectArea.Text
        
        For row_counter = 1 To print_preview_counter Step 1
        
        xlsheet.Cells(row_counter + 7, 1) = row_counter
        xlsheet.Cells(row_counter + 7, 2) = PrintPreviewGrid.TextMatrix(row_counter, 1)
        xlsheet.Cells(row_counter + 7, 4) = PrintPreviewGrid.TextMatrix(row_counter, 2)
        xlsheet.Cells(row_counter + 7, 6) = PrintPreviewGrid.TextMatrix(row_counter, 3)
        xlsheet.Cells(row_counter + 7, 8) = PrintPreviewGrid.TextMatrix(row_counter, 4)
        
        print_prev_status.Caption = "Creating files….  " & (row_counter / print_preview_counter) * 100 & "%"
        Next

    
    xlwbook.SaveAs (App.Path & "\PRINT\print.xls")
    xl.ActiveWorkbook.Close False, App.Path & "\PRINT\pass_due.xls"
    xl.Quit
    
    Set xlwbook = Nothing
    Set xl = Nothing
    
    MsgBox "Data has been sent to the printer!", vbInformation
    print_prev_status.Visible = False
    Main.Enabled = True
    Unload loading
    Shell (App.Path & "\PRINT\printing.exe")
    Shell (App.Path & "\PRINT\delete.bat")
End Sub


Private Sub master_list()
print_prev_status.Visible = True
Main.Enabled = False
loading.Show
loading.Enabled = False
Dim row_counter As Integer

Set xlwbook = xl.Workbooks.Open(App.Path & "\PRINT\masterlist_bal.xls")
Set xlsheet = xlwbook.Sheets.Item(1)

        xlsheet.Cells(5, 2) = print_preview_selectArea.Text
        
        For row_counter = 1 To print_preview_counter Step 1
        
        xlsheet.Cells(row_counter + 7, 1) = row_counter
        xlsheet.Cells(row_counter + 7, 2) = PrintPreviewGrid.TextMatrix(row_counter, 1)
        xlsheet.Cells(row_counter + 7, 4) = PrintPreviewGrid.TextMatrix(row_counter, 2)
        xlsheet.Cells(row_counter + 7, 6) = PrintPreviewGrid.TextMatrix(row_counter, 3)
       
        
        print_prev_status.Caption = "Creating files….  " & (row_counter / print_preview_counter) * 100 & "%"
        Next

    
    xlwbook.SaveAs (App.Path & "\PRINT\print.xls")
    xl.ActiveWorkbook.Close False, App.Path & "\PRINT\masterlist_bal.xls"
    xl.Quit
    
    Set xlwbook = Nothing
    Set xl = Nothing
    
    MsgBox "Data has been sent to the printer!", vbInformation
    print_prev_status.Visible = False
    Main.Enabled = True
    Unload loading
    Shell (App.Path & "\PRINT\printing.exe")
    Shell (App.Path & "\PRINT\delete.bat")

End Sub

Private Sub master_no_bal()
print_prev_status.Visible = True
Main.Enabled = False
loading.Show
loading.Enabled = False
Dim row_counter As Integer

Set xlwbook = xl.Workbooks.Open(App.Path & "\PRINT\master_list_no_bal.xls")
Set xlsheet = xlwbook.Sheets.Item(1)
        
        For row_counter = 1 To print_preview_counter Step 1
        
        xlsheet.Cells(row_counter + 1, 1) = PrintPreviewGrid.TextMatrix(row_counter, 1)
        xlsheet.Cells(row_counter + 1, 2) = PrintPreviewGrid.TextMatrix(row_counter, 2)
        xlsheet.Cells(row_counter + 1, 3) = "_______"
        xlsheet.Cells(row_counter + 1, 4) = PrintPreviewGrid.TextMatrix(row_counter, 1)
        xlsheet.Cells(row_counter + 1, 5) = PrintPreviewGrid.TextMatrix(row_counter, 2)
        xlsheet.Cells(row_counter + 1, 6) = "_______"
       
        
        print_prev_status.Caption = "Creating files….  " & (row_counter / print_preview_counter) * 100 & "%"
        Next

    
    xlwbook.SaveAs (App.Path & "\PRINT\print.xls")
    xl.ActiveWorkbook.Close False, App.Path & "\PRINT\master_list_no_bal.xls"
    xl.Quit
    
   
    
    
    Set xlwbook = Nothing
    Set xl = Nothing
    
    MsgBox "Data has been sent to the printer!", vbInformation
    print_prev_status.Visible = False
    Main.Enabled = True
    Unload loading

    Shell (App.Path & "\PRINT\printing.exe")
    Shell (App.Path & "\PRINT\delete.bat")
End Sub

