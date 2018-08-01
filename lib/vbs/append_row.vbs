Public Function appendRow(r, value, valueColl, collsRange, ws)  
  Dim splittedAddr : splittedAddr = Split(collsRange, ":")
  Dim first : first = splittedAddr(0)
  Dim last : last = splittedAddr(1)
  copyRowWithFormulas first & r, last & r, ws
  ws.Range(valueColl & r + 1).value = value
End Function

Sub copyRowWithFormulas(first, last, ws)
  strPattern = ".*[a-z]+($)?[0-9]+.*"
  Set regEx = New RegExp
  
  regEx.Pattern = strPattern
  regEx.IgnoreCase = True
  regEx.Global = True

  ws.Range(first, last).Copy
  ws.Range(first, last).Offset(1).Insert  

  For Each c In ws.Range(first, last).Offset(1, 0)
    If Not regEx.Test(c.Formula) Then
      c.value = ""
    End If
  Next
End Sub

Public Sub main(path, sheet, rowIndex, collumn, value, collsRange)
  WScript.Stdout.WriteLine "path: " & path
  WScript.Stdout.WriteLine "sheet: " & sheet
  
  Set objExcel = CreateObject("Excel.Application") 
  
  Set wb = objExcel.Workbooks.Open(path, True, False)
  Set ws = wb.Worksheets(sheet)
  
  i = appendRow((rowIndex), (value), (collumn), (collsRange), ws)

  wb.Save
  wb.Close
  objExcel.Quit
End Sub


If WScript.Arguments.Count <= 5 Then
  WScript.Stdout.WriteLine "Invalid number of arguments"
  Wscript.Quit
End if

Dim path : path = WScript.Arguments(0)
Dim sheet : sheet = WScript.Arguments(1)
Dim rowIndex : rowIndex = WScript.Arguments(2)
Dim collumn : collumn = WScript.Arguments(3)
Dim collsRange : collsRange = WScript.Arguments(4)
Dim value : value = WScript.Arguments(5)
WScript.Stdout.WriteLine "Append starts"
main path, sheet, rowIndex, collumn, value, collsRange

WScript.Stdout.WriteLine "Append complete"