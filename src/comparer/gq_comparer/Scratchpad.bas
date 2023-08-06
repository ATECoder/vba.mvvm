Attribute VB_Name = "Scratchpad"
'@Folder "MVVM"
Option Explicit

Sub DoSomething()
    
    Dim list As New ArrayList
    'populate with unsorted things - these don't have to be strings
    list.Add "d"
    list.Add "b"
    list.Add "c"
    list.Add "a"
    
    Debug.Print "UnSorted:", Join(list.ToArray, ", ")
    
    'the sort_2 overload lets us pass in an IComparer instance, here it is a GUIComparer with a testbox view
    'You could envisage a different view such as one which displays two photos
    list.sort_2 GUIComparer.Create(ViewFactory:=TextRepresentableThingsView)
    
    Debug.Print "Sorted:", Join(list.ToArray, ", ")
    
    
End Sub


Sub SortSelection()

    Dim list As New mscorlib.ArrayList
    Dim rangeToSort As Range
    Set rangeToSort = Selection
    
    Dim item As Range
    For Each item In rangeToSort
        list.Add item.Value
    Next item
    
    list.sort_2 GUIComparer.Create(ViewFactory:=TextRepresentableThingsView)
    
    dumpList list, dumpWhere:=rangeToSort
    
End Sub

Private Sub dumpList(ByVal list As mscorlib.ArrayList, Optional ByVal dumpWhere As Range)
    If dumpWhere Is Nothing Then Set dumpWhere = ThisWorkbook.Sheets.Add().Range("A1")
    dumpWhere.Resize(list.Count, 1).Value = WorksheetFunction.Transpose(list.ToArray)
End Sub
