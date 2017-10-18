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

Dim print_string As String
Dim print_split() As String
Dim amount  As String
Dim rel_no As String
Dim name As String
Dim add As String
Dim percent As String
Dim rate As String
Dim day_date As String
Dim m_y_date As String
Dim r_date As String
Dim ledger As String
Dim balance As String
Dim ret_amount As String
Dim date_split() As String
Dim day As String
Dim month As String
Dim year As String




Open App.Path & "\print.dll" For Input As #1
print_string = Input$(LOF(1), #1)
Close #1



print_split = Split(print_string, "*")



amount = Trim(print_split(0))
rel_no = Trim(print_split(1))
name = Trim(print_split(2))
add = Trim(print_split(3))
percent_amount = Trim(print_split(4))
rate = Trim(print_split(5))
r_date = Trim(print_split(6))
ledger = Trim(print_split(7))
balance = Trim(print_split(8))
net_amount = Trim(print_split(9))

date_split = Split(Trim(print_split(6)), "/")

month = date_split(0)
day = date_split(1)
year = date_split(2)

If month = "1" Then
month = "January"

ElseIf month = "2" Then
month = "February"

ElseIf month = "3" Then
month = "March"

ElseIf month = "4" Then
month = "April"

ElseIf month = "5" Then
month = "May"

ElseIf month = "6" Then
month = "June"

ElseIf month = "7" Then
month = "July"

ElseIf month = "8" Then
month = "August"

ElseIf month = "9" Then
month = "September"

ElseIf month = "10" Then
month = "October"

ElseIf month = "11" Then
month = "November"

ElseIf month = "12" Then
month = "December"

End If



day_date = day
m_y_date = month & " " & year





' doc works starts ------------



Shell ("copy.bat")

Dim WordObj As Word.Application
Dim Worddoc As Word.Document


Set WordObj = New Word.Application
Set Worddoc = WordObj.Documents.Open(FileName:=App.Path & "\pnpn.doc")


With Worddoc.Bookmarks
    .Item("Address").Range.Text = add
    .Item("Address2").Range.Text = add
    .Item("Address3").Range.Text = add
    .Item("Amount").Range.Text = amount
    .Item("Amount2").Range.Text = amount
    .Item("Amount3").Range.Text = amount
    .Item("Amount4").Range.Text = amount
    .Item("Balance").Range.Text = balance
    .Item("Balance2").Range.Text = balance
    .Item("rate").Range.Text = rate
    .Item("date").Range.Text = r_date
    .Item("date2").Range.Text = r_date
    .Item("day").Range.Text = day_date
    .Item("day2").Range.Text = day_date
    .Item("Ledger").Range.Text = ledger
    .Item("Ledger2").Range.Text = ledger
    .Item("month_year").Range.Text = m_y_date
    .Item("month_year2").Range.Text = m_y_date
    .Item("Name").Range.Text = name
    .Item("Name2").Range.Text = name
    .Item("Name3").Range.Text = name
    .Item("Name4").Range.Text = name
    .Item("Net").Range.Text = net_amount
    .Item("Net2").Range.Text = net_amount
    .Item("Percent").Range.Text = percent_amount
    .Item("Rel_no").Range.Text = rel_no

End With
Dim PNname As String

PNname = name & " ---- " & rel_no ' added 1-1-12
ActiveDocument.SaveAs (App.Path & "\PN\" & PNname & ".doc") ' added 1-1-12
'ActiveDocument.SaveAs (App.Path & "\copy_pn.Doc") ------ changes made due printing problem - 1-1-12

Shell ("delete.bat") ' added 1-1-12
Unload Me

WordObj.Quit False
Set Worddoc = Nothing
Set WordObj = Nothing

Shell ("delete.bat") ' added 1-1-12
Unload Me


' changes made due printing problem - 1-1-12

Shell ("delete.bat") ' added 1-1-12
Unload Me

Exit Sub
' changes made due printing problem - 1-1-12


'Exit Sub




' printing starts---------- this code it not executed due to 1-1-112 changes


Dim objWord As Word.Application
Dim objBookmark As Word.Bookmark

    Set objWord = New Word.Application
    objWord.Documents.add App.Path & "\copy_pn.Doc", , , True
    objWord.ActiveDocument.PrintOut
    

    Do While objWord.BackgroundPrintingStatus > 0
    Loop
objWord.Quit False
Set objWord = Nothing

Shell ("delete.bat")

Unload Me
End Sub
