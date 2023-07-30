Attribute VB_Name = "RemoveLegacyElements"
Sub RemoveElements()

    Dim p_element As VBComponent
    For Each p_element In ActiveWorkbook.VBProject.VBComponents
        If p_element.Name = "Sheet1" Then
        ElseIf p_element.Name = ActiveWorkbook.Name Then
        ElseIf p_element.Name = "RemoveLegacyElements" Then
        ElseIf p_element.Name = "CustomErrors" Then
        ElseIf p_element.Name = "FormsProgID" Then
        ElseIf p_element.Name = "GuardClauses" Then
        ElseIf p_element.Name = "Resources" Then
        ElseIf 1 = VBA.InStr(1, p_element.Name, "Test", vbTextCompare) Then
        ElseIf 1 = VBA.InStr(1, p_element.Name, "ITest", vbTextCompare) Then
        ElseIf 1 = VBA.InStr(1, VBA.Right(p_element.Name, 5), "Tests", vbTextCompare) Then
        
        Else
            If p_element.Type <> vbext_ct_Document Then
            ActiveWorkbook.VBProject.VBComponents.Remove p_element
            End If
        End If
    Next

End Sub


