Sub VariacionesaActivos()


Sheets("ACTIVOS").Activate
Rows("1:1").Select
Selection.AutoFilter

Var = Sheets("ACTIVOS").Range("A" & Rows.Count).End(xlUp).Row

For i = 2 To Var

Sheets("ACTIVOS").Cells(i, 54) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 2), Sheets("VARIACION").Range("A:P"), 16, 0)

If Not IsError(Sheets("ACTIVOS").Cells(i, 54)) Then
 If Trim(Sheets("ACTIVOS").Cells(i, 54)) = "SI REPORTAR" Then
         Sheets("ACTIVOS").Cells(i, 46) = "ACTUALIZACION"
 End If
 End If
Next i
  
    
 Rows("1:1").Select
 Range("AT1").Activate
 Selection.AutoFilter
 Sheets("ACTIVOS").Range("AT1").AutoFilter Field:=46, Criteria1:="<>*NO REPORTAR*"
    
MsgBox ("Variaciones Listas"), vbInformation, "Op plus"




End Sub

