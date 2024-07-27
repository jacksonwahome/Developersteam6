Attribute VB_Name = "arrays"
Option Explicit
Sub multiarrays()
Dim multiarray(1 To 10, 1 To 3)
Dim dimension1 As Long, dimension2 As Long

multiarray(1, 1) = Range("b2").Value
multiarray(1, 2) = Range("c2").Value
multiarray(1, 3) = Range("d2").Value
End Sub
 or
 
 
 Sub multiarrays2()
Dim multiarray(1 To 10, 1 To 3)
Dim dimension1 As Long, dimension2 As Long

For dimension1 = 1 To 10
    For dimension2 = 1 To 3
    multiarray(dimension1, dimension2) = Range("b2").Offset(dimension1 - 1, dimension2 - 1).Value
    Next dimension2
Next dimension1
End Sub
 Sub multiarrays3()
Dim multiarray(1 To 10, 1 To 3)
Dim dimension1 As Long, dimension2 As Long
For dimension1 = 1 To 10
    For dimension2 = 1 To 3
    multiarray(dimension1, dimension2) = Range("b2").Offset(dimension1 - 1, dimension2 - 1).Value
    Next dimension2
Next dimension1

For dimension1 = 1 To 10
    For dimension2 = 1 To 3
      Range("h16").Offset(dimension1 - 1, dimension2 - 1).Value = multiarray(dimension1, dimension2)
    Next dimension2
Next dimension1
End Sub

'using lbound and Ubound
 Sub multiarrays4()
Dim multiarray(1 To 10, 1 To 5)
Dim dimension1 As Long, dimension2 As Long
For dimension1 = LBound(multiarray, 1) To UBound(multiarray, 1)
    For dimension2 = LBound(multiarray, 2) To UBound(multiarray, 2)
    multiarray(dimension1, dimension2) = Range("b2").Offset(dimension1 - 1, dimension2 - 1).Value
    Next dimension2
Next dimension1

For dimension1 = LBound(multiarray, 1) To UBound(multiarray, 1)
    For dimension2 = LBound(multiarray, 2) To UBound(multiarray, 2)
      Range("h16").Offset(dimension1 - 1, dimension2 - 1).Value = multiarray(dimension1, dimension2)
    Next dimension2
Next dimension1
End Sub



'using dynamic arrays 'using redim function
 Sub multiarrays5()
Dim multiarray()
Dim dimension1 As Long, dimension2 As Long
Dim numberdimension1 As Long, numberdimension2 As Long
numberdimension1 = Range("b2", Range("b2").End(xlDown)).Cells.Count
numberdimension2 = Range("b2", Range("b2").End(xlToRight)).Cells.Count

ReDim multiarray(1 To numberdimension1, 1 To numberdimension2)

For dimension1 = LBound(multiarray, 1) To UBound(multiarray, 1)
    For dimension2 = LBound(multiarray, 2) To UBound(multiarray, 2)
    multiarray(dimension1, dimension2) = Range("b2").Offset(dimension1 - 1, dimension2 - 1).Value
    Next dimension2
Next dimension1

For dimension1 = LBound(multiarray, 1) To UBound(multiarray, 1)
    For dimension2 = LBound(multiarray, 2) To UBound(multiarray, 2)
      Range("h16").Offset(dimension1 - 1, dimension2 - 1).Value = multiarray(dimension1, dimension2)
    Next dimension2
Next dimension1

'the errase array line is optional but there are some situation where you might need it
'eg usign the same array with different elements or where you have other lines of code below that you do not want to interfere
'Erase multiarray
End Sub

