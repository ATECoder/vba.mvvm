VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub TestCreatingObject()

    Dim obj As Object
    Set obj = New AcceptCommand
    If obj Is Nothing Then _
        Debug.Print "object failed to construct": Exit Sub
    
    Set obj = Factory.NewAcceptCommand
    If obj Is Nothing Then _
        Debug.Print "object failed to construct": Exit Sub
    
    
End Sub

