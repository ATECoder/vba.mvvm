Attribute VB_Name = "RemoveReferencedElements"
Sub RemoveElements()

    Dim p_element As VBComponent
    For Each p_element In ActiveWorkbook.VBProject.VBComponents
        If p_element.Name = "Sheet1" Then
        ElseIf p_element.Name = "Sheet2" Then
        ElseIf p_element.Name = "Sheet3" Then
        ElseIf p_element.Name = ActiveWorkbook.Name Then
        ElseIf p_element.Name = "ExampleDynamicView" Then
        ElseIf p_element.Name = "ExampleView" Then
        ElseIf p_element.Name = "ExploreTextboxEvents" Then
        ElseIf p_element.Name = "Resources" Then
        ElseIf p_element.Name = "BindingManagerTests" Then
        ElseIf p_element.Name = "BindingPathTests" Then
        ElseIf p_element.Name = "CommandManagerTests" Then
        ElseIf p_element.Name = "CustomErrors" Then
        ElseIf p_element.Name = "Example" Then
        ElseIf p_element.Name = "FormsProgID" Then
        ElseIf p_element.Name = "GuardClauses" Then
        ElseIf p_element.Name = "ValidationManagerTests" Then
        
        Else
            If p_element.Type <> vbext_ct_Document Then
            ActiveWorkbook.VBProject.VBComponents.Remove p_element
            End If
        End If
    Next

End Sub


