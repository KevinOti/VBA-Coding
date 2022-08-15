VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OrderEntry 
   Caption         =   "UserForm1"
   ClientHeight    =   8490.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9060.001
   OleObjectBlob   =   "OrderEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OrderEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmboutlet_Change()

On Error Resume Next
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Outlets")

If Me.cmboutlet.Value = "" Then Me.txtchain.Value = "" And Me.txtRegion.Value = ""
If Me.cmboutlet.Value <> "" Then Me.txtchain.Value = Application.WorksheetFunction.VLookup(Me.cmboutlet.Value, sh.Range("A:C"), 2, 0)
If Me.cmboutlet.Value <> "" Then Me.txtRegion.Value = Application.WorksheetFunction.VLookup(Me.cmboutlet.Value, sh.Range("A:C"), 3, 0)

        
End Sub

Private Sub cmbproduct_Change()
On Error Resume Next
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("products")

If Me.cmbProduct.Value = "" Then Me.txtPrice.Value = "" And Me.txtuom.Value = "" And Me.txtorder.Value = ""
If Me.cmbProduct.Value <> "" Then Me.txtcode.Value = _
    Application.WorksheetFunction.VLookup(Me.cmbProduct.Value, sh.Range("A:D"), 2, 0)

If Me.cmbProduct.Value <> "" Then Me.txtuom.Value = _
    Application.WorksheetFunction.VLookup(Me.cmbProduct.Value, sh.Range("A:D"), 3, 0)
    
If Me.cmbProduct.Value <> "" Then Me.txtPrice.Value = _
    Application.WorksheetFunction.VLookup(Me.cmbProduct.Value, sh.Range("A:D"), 4, 0)

'If Me.cmboutlet.Value <> "" Then Me.txtchain.Value = _
   ' Application.WorksheetFunction.VLookup(Me.cmbproduct.Value, sh.Range("A:D"), 2, 0)
'End If
'If Me.cmboutlet.Value <> "" Then Me.txtRegion.Value = _
'    Application.WorksheetFunction.VLookup(Me.cmbproduct.Value, sh.Range("A:D"), 3, 0)
'End If
End Sub


Private Sub cmbSave_Click()
ThisWorkbook.Save
End Sub

Private Sub CommandButton1_Click()

If Me.cmboutlet.Value = "" Then
MsgBox "Please select the outlet", vbCritical
    Exit Sub
End If

If Me.cmbProduct.Value = "" Then
MsgBox "Please select the product", vbCritical
    Exit Sub
End If


If IsNumeric(Me.txtorder.Value) = False Then
MsgBox "Please enter the correct Quantity", vbCritical
    Exit Sub
End If
    
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Ordering")
Dim lr As Integer

'If Application.WorksheetFunction.CountIf(sh.Range("E:E"), Me.cmbproduct.Value) > 0 Then
    'MsgBox "The product already exits", vbInformation
    'Exit Sub
'End If
lr = Application.WorksheetFunction.CountA(sh.Range("A:A"))

sh.Range("A" & lr + 1).Value = lr
sh.Range("B" & lr + 1).Value = Me.cmboutlet.Value
sh.Range("C" & lr + 1).Value = Me.txtchain.Value
sh.Range("D" & lr + 1).Value = Me.txtRegion.Value
sh.Range("E" & lr + 1).Value = Me.cmbProduct.Value
sh.Range("F" & lr + 1).Value = Me.txtcode.Value
sh.Range("G" & lr + 1).Value = Me.txtuom.Value
sh.Range("H" & lr + 1).Value = Me.txtorder.Value
sh.Range("I" & lr + 1).Value = Me.txtorder * Me.txtPrice.Value
sh.Range("J" & lr + 1).Value = Me.txtDate.Value
sh.Range("K" & lr + 1).Value = Me.cmbStatus.Value


Me.cmbProduct.Value = ""
Me.txtcode.Value = ""
Me.txtuom.Value = ""
Me.txtorder.Value = ""
Me.txtPrice.Value = ""
Me.cmbStatus.Value = ""

MsgBox "Product successfully added"
sh.Range("A:L").EntireColumn.AutoFit

End Sub

Private Sub CommandButton2_Click()
If Me.cmboutlet.Value = "" Then
MsgBox "Please select the outlet", vbCritical
    Exit Sub
End If

If Me.cmbProduct.Value = "" Then
MsgBox "Please select the product", vbCritical
    Exit Sub
End If


If IsNumeric(Me.txtorder.Value) = False Then
MsgBox "Please enter the correct Quantity", vbCritical
    Exit Sub
End If
    
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Ordering")
Dim lr As Integer
'If Application.WorksheetFunction.CountIf(sh.Range("E:E"), Me.cmbproduct.Value) > 0 Then
    'MsgBox "The product already exits", vbInformation
    'Exit Sub
'End If
lr = Me.TextBox1.Value

sh.Range("A" & lr + 1).Value = lr
sh.Range("B" & lr + 1).Value = Me.cmboutlet.Value
sh.Range("C" & lr + 1).Value = Me.txtchain.Value
sh.Range("D" & lr + 1).Value = Me.txtRegion.Value
sh.Range("E" & lr + 1).Value = Me.cmbProduct.Value
sh.Range("F" & lr + 1).Value = Me.txtcode.Value
sh.Range("G" & lr + 1).Value = Me.txtuom.Value
sh.Range("H" & lr + 1).Value = Me.txtorder.Value
sh.Range("I" & lr + 1).Value = Me.txtorder * Me.txtPrice.Value
sh.Range("J" & lr + 1).Value = Me.txtDate.Value
sh.Range("K" & lr + 1).Value = Me.cmbStatus.Value


Me.cmbProduct.Value = ""
Me.txtcode.Value = ""
Me.txtuom.Value = ""
Me.txtorder.Value = ""
Me.txtPrice.Value = ""
Me.cmbStatus.Value = ""

MsgBox "Product successfully updated"
sh.Range("A:L").EntireColumn.AutoFit

End Sub

Private Sub CommandButton4_Click()
Quotation.Show
End Sub

Private Sub CommandButton5_Click()
Call invoice_gen
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Me.TextBox1.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
Me.cmboutlet.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
Me.txtchain.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
Me.txtRegion.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 3)
Me.cmbProduct.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 4)
Me.txtuom.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 6)
Me.txtcode.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 5)
Me.txtPrice.Value = ""
Me.txtorder.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 7)
Me.cmbStatus.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 10)
Me.txtDate.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 9)

End Sub

Private Sub UserForm_Initialize()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Outlets")
Dim i As Integer
For i = 2 To Application.WorksheetFunction.CountA(sh.Range("A:A"))
Me.cmboutlet.AddItem sh.Range("A" & i)
Next i

With Me.cmbStatus
    .AddItem "Pending"
    .AddItem "Supplied"
End With
Call chainNregion
Call add
Me.txtDate.Value = Format(Date, "D-MMM-YYYY")
End Sub

Sub chainNregion()

Dim i As Integer
Dim pr As Worksheet
Set pr = ThisWorkbook.Sheets("Products")
For i = 2 To Application.WorksheetFunction.CountA(pr.Range("A:A"))
    Me.cmbProduct.AddItem pr.Range("A" & i)

Next i

End Sub

Sub add()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Ordering")
Dim lr As Integer
lr = Application.WorksheetFunction.CountA(sh.Range("A:A"))
If lr = 1 Then lr = 2
With Me.ListBox1
    .ColumnHeads = True
    .ColumnCount = 6
    .ColumnWidths = "50, 80, 0, 50, 150, 80"
    .RowSource = "Ordering!A2:l" & lr
End With

End Sub

Sub invoice_gen()
Dim sh As Worksheet
Dim ish As Worksheet
Dim pd As Worksheet

Set sh = ThisWorkbook.Sheets("Ordering")
Set ish = ThisWorkbook.Sheets("Invoice")
Set pd = ThisWorkbook.Sheets("Create")


sh.AutoFilterMode = False
sh.UsedRange.AutoFilter 2, "=" & Me.cmboutlet.Value
    sh.Range("E:I").Copy
    pd.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    pd.Rows("1").Delete
    pd.Columns("A:E").AutoFit
    pd.Columns("B:C").Delete
    pd.Range("A1").CurrentRegion.Copy
    'ish.Range("B11:B36").Clear
    ish.Range("B11").PasteSpecial xlPasteValuesAndNumberFormats
    ish.Range("A4").Value = sh.Range("B2")
    ish.Range("e3").Value = Left(ish.Range("A4").Value, 3) & "/" & Format(ish.Range("E4").Value, "DD")
    'ish.Range("A11:B36").Borders(xlDiagonalDown).LineStyle = xlNone
    'ish.Range("A11:B36").Borders(xlDiagonalUp).LineStyle = xlNone
    With ish.Range("A11:E36").Borders(xlEdgeLeft)
        '.LineStyle = xlContinuous
        '.Weight = xlThin
    End With
    With ish.Range("A11:E36").Borders(xlEdgeTop)
        '.LineStyle = xlContinuous
        '.Weight = xlThin
    End With
    With ish.Range("A11:E36").Borders(xlEdgeBottom)
        '.LineStyle = xlContinuous
        '.Weight = xlThin
    End With
    With ish.Range("A11:E36").Borders(xlEdgeRight)
        '.LineStyle = xlContinuous
        '.Weight = xlThin
    End With
    With ish.Range("A11:E36").Borders(xlInsideVertical)
        '.LineStyle = xlContinuous
        '.Weight = xlThin
    End With
    With ish.Range("A11:E36").Borders(xlInsideHorizontal)
        '.LineStyle = xlContinuous
        '.Weight = xlThin
    End With
    
    convention = sh.Range("e3").Value
    ish.ExportAsFixedFormat xlTypePDF, Environ("Userprofile") & "\Desktop\invoice.pdf"
''' Saving the invoice copy'''


    
'sh.Range("Y2").PasteSpecial xlPasteValuesAndNumberFormats
    
'

sh.AutoFilterMode = False



End Sub

Sub invoice_tracker()



End Sub







