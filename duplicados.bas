Attribute VB_Name = "Módulo2"
Sub eliminar_duplicados()
Attribute eliminar_duplicados.VB_ProcData.VB_Invoke_Func = " \n14"
'
' borrar_duplicados Macro
'

'
  Application.ScreenUpdating = False
    Columns("A:G").Select
    UltimaFila = Cells(Rows.Count, "G").End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        "E2:E" & UltimaFila), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        "G2:G" & UltimaFila), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        "B2:B" & UltimaFila), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:G" & UltimaFila)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("H3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]-R[-1]C[-1]<10/1440,RC[-3]=R[-1]C[-3],RC[-6]=R[-1]C[-6]),IFERROR(AND(RC[-1]-R[-2]C[-1]<10/1440,RC[-3]=R[-2]C[-3],RC[-6]=R[-2]C[-6]),0),IFERROR(AND(RC[-1]-R[-3]C[-1]<10/1440,RC[-3]=R[-3]C[-3],RC[-6]=R[-3]C[-6]),0)),1,0)"
    Range("H3").Select
    UltimaFila = Cells(Rows.Count, "G").End(xlUp).Row
    Selection.AutoFill Destination:=Range("H3:H" & UltimaFila), Type:=xlFillDefault
    Columns("G:G").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=H1>0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    numDuplicados = WorksheetFunction.CountIf(Range("H1:H" & UltimaFila), ">0")
    resp = MsgBox(numDuplicados & " dupicados resaltados ¿Desea eliminarlos?", vbYesNo + vbQuestion)
    If resp = vbYes Then
  Dim i As Long

  ' Obtener la última fila con datos en la hoja
  UltimaFila = Cells(Rows.Count, 1).End(xlUp).Row

  ' Recorrer las filas de abajo hacia arriba
  For i = UltimaFila To 1 Step -1
    '  Aquí va el criterio. Por ejemplo, si la columna A contiene "X"
    If Cells(i, 8).Value > 0 Then
      Rows(i).Delete
'      Call resaltar_duplicados
    End If
  Next i
'  Call resaltar_duplicados
  Columns(8).Delete
  Application.ScreenUpdating = True
  MsgBox "Duplicados eliminados", vbOKOnly
  Else
    MsgBox "Duplicados resaltados", vbOKOnly
  End If
End Sub
