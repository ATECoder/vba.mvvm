Attribute VB_Name = "RemoveLegacyElements"
Sub RemoveElements()

    Dim p_element As VBComponent
    For Each p_element In ActiveWorkbook.VBProject.VBComponents
        If p_element.Name = "ExampleDynamicView" Or _
           p_element.Name = "ExampleView" Or _
           p_element.Name = "ExploreTextBoxEvents" Or _
           p_element.Name = "BindingManagerTests" Or _
           p_element.Name = "BindingPathTests" Or _
           p_element.Name = "CommandManagerTests" Or _
           p_element.Name = "Example" Or _
           p_element.Name = "ValidationManagerTests" Or _
           False Then
            If p_element.Type <> vbext_ct_Document Then
                ActiveWorkbook.VBProject.VBComponents.Remove p_element
            End If
        End If
    Next

End Sub


