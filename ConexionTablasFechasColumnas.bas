Attribute VB_Name = "Visitas"
Option Explicit

Sub PendienteAdjuntado()

   Dim ws As Worksheet
   Set ws = ThisWorkbook.Worksheets("Datos Visitas")
   
   Dim lastrow As Integer
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
   Dim i As Integer 'contador
   
   For i = 2 To lastrow
   
      If ws.Cells(i, 8).Value = 0 And ws.Cells(i, 9).Value <> 0 Then
         
         ws.Cells(i, 12).Value = ws.Cells(i, 8).Value
         ws.Cells(i, 13).Value = ws.Cells(i, 9).Value
         
      End If
   
      If ws.Cells(i, 8).Value = ws.Cells(i, 9).Value Then
         
         ws.Cells(i, 12).Value = ws.Cells(i, 8).Value - ws.Cells(i, 9).Value
         ws.Cells(i, 13).Value = ws.Cells(i, 8).Value + ws.Cells(i, 9).Value
      
      End If
      
      If ws.Cells(i, 8).Value < ws.Cells(i, 9).Value And ws.Cells(i, 8).Value <> 0 Then
         
         ws.Cells(i, 12).Value = ws.Cells(i, 8).Value - ws.Cells(i, 8).Value
         ws.Cells(i, 13).Value = ws.Cells(i, 8).Value + ws.Cells(i, 9).Value
      
      End If
      
      If ws.Cells(i, 8).Value > ws.Cells(i, 9).Value And ws.Cells(i, 9).Value <> 0 Then
         
         ws.Cells(i, 12).Value = ws.Cells(i, 8).Value
         ws.Cells(i, 13).Value = ws.Cells(i, 9).Value
      
      End If
      
   Next
   
   MsgBox "Cáclulos finalizados"
   

End Sub

Sub FechasVisitas()

  'Variables
   
   Dim wsVisitas As Worksheet 'Hoja"1"
   Set wsVisitas = ThisWorkbook.Worksheets("Datos Visitas")
   
   Dim wsFechas As Worksheet 'Hoja"2"
   Set wsFechas = ThisWorkbook.Worksheets("Visitas Info")
   
   Dim lastRowVisitas As Integer 'Ultima fila de "hoja1"
   lastRowVisitas = wsVisitas.Cells(Rows.Count, 1).End(xlUp).Row
   
   Dim lastColumnVisitas As Integer 'Ultima columna ed "hoja1"
   lastColumnVisitas = 17
   
   Dim lastRowFechas As Integer 'Ultima fila de "hoja2"
   lastRowFechas = wsFechas.Cells(Rows.Count, 1).End(xlUp).Row
   
   Dim i As Integer 'Contador 1
   Dim j As Integer 'Contador 2
   Dim k As Byte 'Contador 3
   k = 0
   
   For i = 2 To lastRowVisitas
   
      For j = 2 To lastRowFechas
      
         lastColumnVisitas = 17
         
            While (wsVisitas.Cells(i, lastColumnVisitas) <> "")
         
               lastColumnVisitas = lastColumnVisitas + 1
            
            Wend
            
            Debug.Print wsVisitas.Cells(i, 1).Value
            Debug.Print wsFechas.Cells(j, 1).Value
            
         
         If wsVisitas.Cells(i, 1).Value = wsFechas.Cells(j, 1).Value And wsFechas.Cells(j, 9).Value = "x" Then
         
            wsVisitas.Cells(i, lastColumnVisitas) = wsFechas.Cells(j, 2).Value
            
            Debug.Print wsFechas.Cells(j, 2).Value
            
         End If
         
         
      Next
      
   Next
            
  
   
   MsgBox "Fin"
   
End Sub
