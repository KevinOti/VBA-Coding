VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Quotation 
   Caption         =   "Quotation"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6150
   OleObjectBlob   =   "Quotation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Quotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Vlookup code to extract price of the product'''
Private Sub cmbproduct_Change()
On Error Resume Next
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Products")
If Me.cmbProduct.Value = "" Then Me.txtPrice.Value = "" And Me.txtPrice.Value = ""

If Me.cmbProduct.Value <> "" Then
    Me.txtPrice.Value = Application.WorksheetFunction.VLookup(Me.cmbProduct.Value, sh.Range("A:D"), 4, 0)
End If

End Sub

Private Sub CommandButton1_Click()
'' Code for adding products to generate the quote''
Dim sh As Worksheet
Dim rsh As Worksheet
Set sh = ThisWorkbook.Sheets("Products")
Set rsh = ThisWorkbook.Sheets("Quotation")

If Me.cmbProduct.Value = "" Then
    MsgBox "Select product"
    Exit Sub
End If
If Me.txtClientiel.Value = "" Then
    MsgBox "please enter the client Name"
    Exit Sub
End If
If Me.txtQty.Value = "" Then
    MsgBox "Please enter the price"
    Exit Sub
End If

Set sh = ThisWorkbook.Sheets("Products")
Set rsh = ThisWorkbook.Sheets("Quotation")

If Me.cmbProduct.Value <> "" Then
    Me.txtPrice.Value = Application.WorksheetFunction.VLookup(Me.cmbProduct.Value, sh.Range("A:D"), 4, 0)
End If

Set sh = ThisWorkbook.Sheets("Products")
Set rsh = ThisWorkbook.Sheets("Quotation")
rsh.Range("A4").Value = Me.txtClientiel.Value
Me.txtClientiel.Value = rsh.Range("A4").Value
rsh.Range("E5").Value = Left(Me.txtClientiel.Value, 3) & "/" & Format(Date, "DD")
Me.txtQuoteID.Value = rsh.Range("E5").Value
rsh.Range("E4").Value = Left(Range("A1"), 3) & "/" & Format(Date, "DD")
Me.txtQuoteID.Value = rsh.Range("E4")


Dim lr As Integer
Set rsh = ThisWorkbook.Sheets("Quotation")
lr = Application.WorksheetFunction.CountA(rsh.Range("A:A"))
rsh.Range("A" & lr).Value = lr - 8
rsh.Range("B" & lr).Value = Me.cmbProduct.Value
rsh.Range("C" & lr).Value = Me.txtPrice.Value
rsh.Range("D" & lr).Value = Me.txtQty.Value
'rsh.Range("E" & lr).Value = Me.txtQty * Me.txtPrice

Me.cmbProduct.Value = ""
Me.txtQty.Value = ""
Me.txtPrice.Value = ""

quote_name = sh.Range("E4").Value
'rsh.ExportAsFixedFormat xlTypePDF, Environ("Userprofile") & "\desktop\"
 



End Sub

Sub products()
'''Adding products as drop down'''

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Products")
Dim lr As Integer
For lr = 2 To Application.WorksheetFunction.CountA(sh.Range("A:A"))
Me.cmbProduct.AddItem sh.Range("A" & lr)
Next lr


End Sub

Private Sub CommandButton2_Click()
'''converting quote to pdf'''
Dim rsh As Worksheet
Set rsh = ThisWorkbook.Sheets("Quotation")

rsh.ExportAsFixedFormat xlTypePDF, Environ("Userprofile") & "\Desktop\Quote.pdf"

End Sub

Private Sub UserForm_Initialize()

Call products
Call reveal

End Sub

Sub reveal()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Quotation")

End Sub









