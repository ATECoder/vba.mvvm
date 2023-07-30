VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestStubBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Tests.Stubs
Option Explicit

Private Type TState
    MemberInvokes As Dictionary
End Type

Private This As TState

'@Description "Tracks a new invoke of the specified member."
Public Sub OnInvoke(ByVal a_memberName As String)
Attribute OnInvoke.VB_Description = "Tracks a new invoke of the specified member."
    Dim NewValue As Long
    If This.MemberInvokes.Exists(a_memberName) Then
        NewValue = This.MemberInvokes.Item(a_memberName) + 1
        This.MemberInvokes.Remove a_memberName
    Else
        NewValue = 1
    End If
    This.MemberInvokes.Add a_memberName, NewValue
    Debug.Print TypeName(Me) & ": OnInvoke(" & a_memberName & ") was invoked"
End Sub

'@Description "Gets the number of invokes made against the specified member in the lifetime of this object."
Public Property Get MemberInvokes(ByVal a_memberName As String) As Long
Attribute MemberInvokes.VB_Description = "Gets the number of invokes made against the specified member in the lifetime of this object."
    If This.MemberInvokes.Exists(a_memberName) Then
        MemberInvokes = This.MemberInvokes.Item(a_memberName)
    Else
        MemberInvokes = 0
    End If
End Property

'@Description "Gets a string listing the MemberInvokes cache content."
Public Function ToString() As String
Attribute ToString.VB_Description = "Gets a string listing the MemberInvokes cache content."
    Dim p_memberNames As Variant
    p_memberNames = This.MemberInvokes.Keys
    
    Dim p_builder As StringBuilder: p_builder = cc_isr_Core.Constructor.CreateStringBuilder
    Dim i As Long
    For i = LBound(p_memberNames) To UBound(p_memberNames)
        Dim Name As String
        Name = p_memberNames(i)
        p_builder.AppendFormat "{0} was invoked {1} time(s)", Name, This.MemberInvokes.Item(Name)
    Next
    ToString = p_builder.ToString
    
End Function

Public Function Verify(ByVal a_assert As cc_isr_Test_Fx.Assert, ByVal a_memberName As String, _
                       ByVal a_expectedInvokes As Long) As cc_isr_Test_Fx.Assert
    
    Dim p_invokes As Long
    p_invokes = MemberInvokes(a_memberName)
    a_assert.AreEqual a_expectedInvokes, p_invokes, _
                      "Member was invoked " & p_invokes & " time(s), but test expected " & a_expectedInvokes & "."

End Function

Private Sub Class_Initialize()
    Set This.MemberInvokes = New Dictionary
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - -
'
' ITestSub implementation
'
' - - - - - - - - - - - - - - - - - - - - - - - -

Private Property Get ITestStub_MemberInvokes(ByVal a_memberName As String) As Long
    
    ITestStub_MemberInvokes = This.MemberInvokes.Item(a_memberName)

End Property

Private Function ITestStub_ToString() As String
    
    ITestStub_ToString = ToString

End Function

Private Function ITestStub_Verify(ByVal a_assert As cc_isr_Test_Fx.Assert, ByVal a_memberName As String, ByVal a_expectedInvokes As Long) As cc_isr_Test_Fx.Assert
    
    ITestStub_Verify = Verify(a_assert, a_memberName, a_expectedInvokes)

End Function

