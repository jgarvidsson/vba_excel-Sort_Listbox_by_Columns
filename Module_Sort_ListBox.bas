Attribute VB_Name = "Module_Sort_ListBox"
''' Sort Listbox
''' Array module
'   SortListBox     --> Created by J.G.Arvidsson (2020-07-18)
'   SortArrayAtoZ   --> From page ExcelOffTheGrid
'   SortArrayZtoA   --> From page ExcelOffTheGrid
'
' The Separator inside the SortListBox is a value that the user must choose so as not to interfere with the data in the list.


Function SortListBox(ListBoxName As MSForms.ListBox, Ascendent As Boolean, Optional SortByColumn As Double = 0)

''' Variables
    Dim CountItems As Double
    Dim CountColumns As Double
    Dim Separator As String                 ' String that separate each column ;)
    Dim Container As String                 ' String with all information for row
    Dim ListBoxArray() As String
    Dim ListBoxReturn() As String
    
''' Values Assignation
    CountItems = ListBoxName.ListCount
    CountColumns = ListBoxName.ColumnCount ' Remember, start in 0 to ColumCount-1
    Separator = "___"
    
''' Errors are discarded
    If CountItems = 0 Then Exit Function
    If SortByColumn > CountColumns - 1 Then SortByColumn = CountColumns - 1
    
''' Create Array
    ReDim ListBoxArray(CountItems - 1)
    ReDim ListBoxReturn(CountColumns)
    
''' Convert ListBox in an Arran "momo"-dimensional
    For i = 0 To CountItems - 1
        For n = 0 To CountColumns - 1
            Container = IIf(Container = "", "", Container & Separator) & ListBoxName.Column(n, i)
        Next n
        ListBoxArray(i) = ListBoxName.Column(SortByColumn, i) & Separator & Container
        Container = ""
    Next i
    
''' Using SortArray functions the array is sort according with the user
    If Ascendent = True Then
        SortArrayAtoZ ListBoxArray
    Else
        SortArrayZtoA ListBoxArray
    End If

''' Return the ListBox sort
    ListBoxName.Clear
    For i = 0 To CountItems - 1
        ListBoxReturn = Split(ListBoxArray(i), Separator)
        ListBoxName.AddItem
        For n = 1 To CountColumns       ' Remember that exist a new value in the list and it must be eliminated :) (THE ORDER VALUE)
            ListBoxName.List(i, n - 1) = ListBoxReturn(n)
        Next n
    Next i
    
''' Array content is erased to free up memory space
    Erase ListBoxArray
    Erase ListBoxReturn
End Function


Function SortArrayAtoZ(myArray As Variant)

Dim i As Long
Dim j As Long
Dim Temp

'Sort the Array A-Z
For i = LBound(myArray) To UBound(myArray) - 1
    For j = i + 1 To UBound(myArray)
        If UCase(myArray(i)) > UCase(myArray(j)) Then
            Temp = myArray(j)
            myArray(j) = myArray(i)
            myArray(i) = Temp
        End If
    Next j
Next i

SortArrayAtoZ = myArray

End Function


Function SortArrayZtoA(myArray As Variant)

Dim i As Long
Dim j As Long
Dim Temp

'Sort the Array Z-A
For i = LBound(myArray) To UBound(myArray) - 1
    For j = i + 1 To UBound(myArray)
        If UCase(myArray(i)) < UCase(myArray(j)) Then
            Temp = myArray(j)
            myArray(j) = myArray(i)
            myArray(i) = Temp
        End If
    Next j
Next i

SortArrayZtoA = myArray

End Function
