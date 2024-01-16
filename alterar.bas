Attribute VB_Name = "alterar"
Sub RefreshPivotTables()

  Dim pivotTable As pivotTable

  For Each plan In ActiveWorkbook.Sheets
    For Each pivotTable In plan.PivotTables
        pivotTable.RefreshTable
    Next
  Next

End Sub
Sub ordenar()

    Dim iForsta, iSista As Integer
    Dim i, j As Integer
    Dim sTempA, sTempB, sTempC As String
    iForsta = 0
    iSista = UserForm_copy.ComboBox1.ListCount - 1
    For i = iForsta To iSista - 1
        For j = i + 1 To iSista
            If UserForm_copy.ComboBox1.List(i) > UserForm_copy.ComboBox1.List(j) Then
                sTempC = UserForm_copy.ComboBox1.List(j)
                 UserForm_copy.ComboBox1.List(j) = UserForm_copy.ComboBox1.List(i)
                UserForm_copy.ComboBox1.List(i) = sTempC
            End If
        Next j
    Next i
    
    iSista = UserForm_copy.ComboBox2.ListCount - 1
    
    For i = iForsta To iSista - 1
        For j = i + 1 To iSista
            If UserForm_copy.ComboBox2.List(i) > UserForm_copy.ComboBox2.List(j) Then
                sTempC = UserForm_copy.ComboBox2.List(j)
                 UserForm_copy.ComboBox2.List(j) = UserForm_copy.ComboBox2.List(i)
                UserForm_copy.ComboBox2.List(i) = sTempC
            End If
        Next j
    Next i
    
    iSista = UserForm_copy.ComboBox3.ListCount - 1
    For i = iForsta To iSista - 1
        For j = i + 1 To iSista
            If UserForm_copy.ComboBox3.List(i) > UserForm_copy.ComboBox3.List(j) Then
                sTempC = UserForm_copy.ComboBox3.List(j)
                 UserForm_copy.ComboBox3.List(j) = UserForm_copy.ComboBox3.List(i)
                UserForm_copy.ComboBox3.List(i) = sTempC
            End If
        Next j
        
    Next i

End Sub
Sub FormControlSortAZ()
    Dim dados As Variant
    Dim vetor() As Variant
    Dim inf As Long, sup As Long
    Dim i As Integer
    
    dados = UserForm_copy.ListBox3.List
    inf = LBound(dados, 1)
    sup = UBound(dados, 1)
    ReDim vetor(inf To sup, 1 To 1)
    
    For i = inf To sup
        vetor(i, 1) = dados(i, 2)
    Next
    
    UserForm_copy.ListBox3.List = WorksheetFunction.Rank(dados, vetor, 1)
End Sub
Function SortArrayZtoA()
    Dim myArray As Variant
    Dim i As Long
    Dim j As Long
    Dim Temp
    
    myArray = UserForm_copy.ListBox3.List
    
    'Sort the Array Z-A
    For i = LBound(myArray) To UBound(myArray) - 1
        For j = i + 1 To UBound(myArray)
        a = Format(myArray(i, 2), "d")
        b = Format(myArray(j, 2), "d")
            If a < b Then
                Temp = myArray(j, 2)
                myArray(j, 2) = myArray(i, 2)
                myArray(i, 2) = Temp
            End If
        Next j
    Next i
    
    SortArrayZtoA = myArray
End Function
Sub ordenar_comp()
    Dim iForsta, iSista As Integer
    Dim i, j As Integer
    Dim sTempA, sTempB, sTempC As String
    iForsta = 0
    iSista = UserForm_copy.ListBox3.ListCount - 1
    
    For i = iForsta To iSista - 1
        
        For j = i + 1 To iSista
        class_a = Format(UserForm_copy.ListBox3.List(i, 2), "d")
        class_b = Format(UserForm_copy.ListBox3.List(j, 2), "d")
            If class_a < class_b Then
               UserForm_copy.ListBox3.ListIndex = j
             End If
        Next j
    Next i

End Sub
Sub ordenar_fat()
    Dim iForsta, iSista As Integer
    Dim i, j As Integer
    Dim sTempA, sTempB, sTempC As String
    iForsta = 0
    iSista = UserForm_copy.ListBox4.ListCount - 1
    
    For i = iForsta To iSista - 1
    
        For j = i + 1 To iSista
        d1 = Format(UserForm_copy.ListBox4.List(i, 2), "dd")
        d2 = Format(UserForm_copy.ListBox4.List(j, 2), "dd")
            If d1 > d2 Then
                sTempC = Format(UserForm_copy.ListBox4.List(j, 2), "dd")
                UserForm_copy.ListBox4.List(j, 2) = Format(UserForm_copy.ListBox4.List(i, 2), "dd")
                UserForm_copy.ListBox4.List(i, 2) = sTempC
            End If
        Next j
    Next i

End Sub
Sub reload()
    Dim linha As Integer, i As Integer
    UserForm_copy.ListBox1.Clear
    linha = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    
    For i = linha To 2 Step -1
    
        With UserForm_copy.ListBox1
            .AddItem
            .List(.ListCount - 1, 0) = ActiveSheet.Range("B1").Cells(i, 1)
            .List(.ListCount - 1, 1) = ActiveSheet.Range("B1").Cells(i, 2)
            .List(.ListCount - 1, 2) = ActiveSheet.Range("B1").Cells(i, 3)
            .List(.ListCount - 1, 3) = ActiveSheet.Range("B1").Cells(i, 4)
            .List(.ListCount - 1, 4) = ActiveSheet.Range("B1").Cells(i, 5)
            .List(.ListCount - 1, 5) = ActiveSheet.Range("B1").Cells(i, 6)
            .List(.ListCount - 1, 6) = ActiveSheet.Range("B1").Cells(i, 7)
            .List(.ListCount - 1, 7) = ActiveSheet.Range("B1").Cells(i, 8)
            .List(.ListCount - 1, 8) = ActiveSheet.Range("B1").Cells(i, 9)
        End With
    Next i
    registros = UserForm_copy.ListBox1.ListCount
    UserForm_copy.ListBox1.ColumnWidths = "300;40;60;40;40;40;40;40;40;40"
    
    For m = linha To 2 Step -1
        With UserForm_copy.ListBox2
            .AddItem
            .List(.ListCount - 1, 0) = ActiveSheet.Range("B1").Cells(m, 10)
            .List(.ListCount - 1, 1) = ActiveSheet.Range("B1").Cells(m, 11)
            .List(.ListCount - 1, 2) = registros - .ListCount + 2
        End With
        With ComboBox1
        .AddItem ActiveSheet.Range("B1").Cells(m, 1) & " " & registros - .ListCount + 1
        End With
    Next m
    
    Call ordenar

End Sub
Sub LookupTableValue()
    Dim tbl As ListObject
    Dim FoundCell As Range
    Dim LookupValue As String
    
    'Lookup Value
      LookupValue = "ID-123"
    
    'Store Table Object to a variable
      Set tbl = ActiveSheet.ListObjects("Table1")
    
    'Attempt to find value in Table's first Column
      On Error Resume Next
      Set FoundCell = tbl.DataBodyRange.Columns(1).Find(LookupValue, LookAt:=xlWhole)
      On Error GoTo 0
    
    'Return Table Row number if value is found
      If Not FoundCell Is Nothing Then
        MsgBox "Found in table row: " & _
          tbl.ListRows(FoundCell.Row - tbl.HeaderRowRange.Row).Index
      Else
        MsgBox "Value not found"
      End If

End Sub
Sub MultiColumnTable_To_Array()
    Dim myTable As ListObject
    Dim myArray As Variant
    Dim X As Long
    
    'Set path for Table variable
      Set myTable = ActiveSheet.ListObjects("Table1")
    
    'Create Array List from Table
      myArray = myTable.DataBodyRange
    
    'Loop through each item in Third Column of Table (displayed in Immediate Window [ctrl + g])
      For X = LBound(myArray) To UBound(myArray)
        Debug.Print myArray(X, 3)
      Next X
  
End Sub
