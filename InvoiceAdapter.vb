Sub DataCleansing()

Dim IVA_Staging As Single
Dim IVA As Single
Dim RowNumberID As Integer
Dim Neto As Single
Dim IVA_Percent As Single
Dim Num_Factura As String


lastrow = Range("B" & Rows.Count).End(xlUp).Row


For i = 2 To lastrow

  
If Cells(i, 3).Value <> vbNullString Then

'Translate Tipo de Factura

    If Cells(i, 3).Value = "Bill" Then
    Cells(i, 3).Value = "Factura"
    ElseIf Cells(i, 3).Value = "Supplier Credit" Then Cells(i, 3).Value = "Nota de Credito"
    Else: Cells(i, 3).Value = "Gasto"
    End If

Cells(i, 19).Value = Cells(i, 9).Value
Cells(i, 8).Value = Cells(i + 1, 8).Value
Num_Factura = Cells(i, 4).Value

'if Bill Type is C or Cloud Service --> insert value in column "Exento" else inert in column "Neto"

    If Left(Num_Factura, 1) = "C" Or Cells(i, 8) = "Cloud Services" Then
    Cells(i, 13).Value = Cells(i, 9).Value
    Else
    Cells(i, 9).Value = Cells(i + 1, 9).Value
    End If

'Cells(i, 1).Value = Month(Cells(i, 2).Value)
RowNumberID = i
Neto = Cells(i, 9).Value
Rows(RowNumberID + 1).EntireRow.Delete


End If

'Calculate and Set IVA (10,5-21-27) at the relative column

If Cells(i, 8).Value = "New_IVA Payable" And Cells(i + 1, 8).Value <> "New_IVA Payable" Then
IVA_Staging = Cells(i, 9).Value
IVA = Abs(IVA_Staging)
IVA_Percent = IVA / Neto * 100

Select Case IVA_Percent
   Case 20 To 22
     Cells(RowNumberID, 10).Value = IVA
  Case 10 To 11
      Cells(RowNumberID, 11).Value = IVA
  Case 26 To 28
      Cells(RowNumberID, 12).Value = IVA
   End Select
   
End If

' Set Percepcion IVA
If Cells(i, 8).Value = "Percepcion IVA" Then Cells(RowNumberID, 14).Value = Cells(i, 9).Value

' Set IIBB
If Cells(i, 8).Value = "Percepcion Ingresos Brutos" Then Cells(RowNumberID, 15).Value = Cells(i, 9).Value

' Set Comercio e Industria
If Cells(i, 8).Value = "Percepciones Comercio e Industria Cba" Then Cells(RowNumberID, 16).Value = Cells(i, 9).Value

' Set Otros Impuestos
If Cells(i, 8).Value = "Impuestos Internos" Then Cells(RowNumberID, 17).Value = Cells(i, 9).Value


Next i




End Sub







