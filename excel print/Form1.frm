VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

Dim xlApp As Excel.Application
Dim xlWB As Excel.Workbook
Dim xlSH As Excel.Worksheet

Set xlApp = New Excel.Application
Set xlWB = xlApp.Workbooks.Open(FileName:=App.Path & "\print.xls")
Set xlSH = xlWB.Worksheets(1)
xlSH.PrintOut



xlWB.Close
xlApp.Quit
    Set xlWB = Nothing
    Set xlApp = Nothing
    
Shell ("delete.bat")

Unload Me
End Sub
