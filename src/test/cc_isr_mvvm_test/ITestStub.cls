VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that stubs an interface for testing purposes."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that stubs an interface for testing purposes. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Gets the number of times the specified member was invoked in the lifetime of the object. </summary>
Public Property Get MemberInvokes(ByVal a_memberName As String) As Long
Attribute MemberInvokes.VB_Description = "Gets the number of times the specified member was invoked in the lifetime of the object."
End Property

''' <summary>   Gets a string representation of the object's internal state, for debugging purposes (not intended for asserts!). </summary>
Public Function ToString() As String
Attribute ToString.VB_Description = "Gets a string representation of the object's internal state, for debugging purposes (not intended for asserts!)."
End Function

''' <summary>   Fails a test when the member invokes for the specified member name don't match the expected count. </summary>
Public Function Verify(ByVal a_assert As cc_isr_Test_Fx.Assert, ByVal a_memberName As String, _
                       ByVal a_expectedInvokes As Long) As cc_isr_Test_Fx.Assert
Attribute Verify.VB_Description = "Fails a test when the member invokes for the specified member name don't match the expected count."
End Function
