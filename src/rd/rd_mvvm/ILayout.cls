VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@Folder rd_MVVM.Infrastructure.View.Dynamic.Layout
'@Exposed
'@Interface

Public Sub BindControlLayout(ByVal Parent As Object, ByVal Child As MSForms.Control, ByVal Anchor As AnchorEdges)
End Sub

Public Sub ResizeLayout()
End Sub
