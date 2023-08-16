VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlEventsPunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides an event sink to relay MSForms.Control events."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Provides an event sink to relay MSForms.Control events. </summary>
''' <remarks>   based on <see heref="https://stackoverflow.com/a/51936950"/> and
'''             <see href="https://stackoverflow.com/a/61893857"/> </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IControlEvents

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


#If VBA7 Then
''' <summary>   Establishes or terminates a connection between a client's sink and a connection point container. </summary>
''' <remarks>
''' This function is available through Windows XP and Windows Server 2003. It might be altered or unavailable
''' in subsequent versions of Windows.
''' <see href="https://docs.microsoft.com/en-us/windows/win32/api/shlwapi/nf-shlwapi-connecttoconnectionpoint"/>
''' </remarks>
''' <param name="a_punk">        [stdole.IUnknown] A pointer to the IUnknown interface of the object to be connected
'''                              to the connection point container. If you set fConnect to FALSE to indicate that you
'''                              are disconnecting the object, this parameter is ignored and can be set to NULL. </param>
''' <param name="riidEvent">     [GUID] The IID of the interface on the connection point container whose
'''                              connection point object is being requested.</param>
''' <param name="fConnect">      [Long/Boolean] TRUE if a connection is being established;
'''                              FALSE if a connection is being broken. </param>
''' <param name="s_punkTarget">  [stdole.IUnknown] A pointer to the connection point container's IUnknown interface. </param>
''' <param name="pdwCookie">     [Long] A connection token. If you set fConnect to TRUE to make a new connection,
'''                              this parameter receives a token that uniquely identifies the connection. If you set
'''                              fConnect to FALSE to break a connection, this parameter must point to the token that
'''                              you received when you called ConnectToConnectionPoint to establish the connection.</param>
''' <param name="ppcpOut">       [LongPtr] A pointer to the connection point container's IConnectionPoint interface,
'''                              if the operation was successful. The calling application must release this pointer
'''                              when it is no longer needed. If the request is unsuccessful, the pointer receives NULL.
'''                              This parameter is optional and can be NULL.</param>
Private Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal a_punk As stdole.IUnknown, _
        ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal a_punkTarget As stdole.IUnknown, _
        ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long
#Else
Private Declare Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal Punk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal PunkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As Long) As Long
#End If

Private Type TState
    ''' <summary>   The IID of the interface on the connection point container whose connection point object
    ''' is being requested. </summary>
    RefIID As GUID
    Connected As Boolean
    PunkTarget As Object
    Cookie As Long
    
    Handlers As Collection
End Type

Private This As TState

''' <summary>   Gets/sets the target MSForms.Control reference. </summary>
Public Property Get Target() As Object
Attribute Target.VB_Description = "Gets/sets the target MSForms.Control reference."
    Set Target = This.PunkTarget
End Property

Public Property Set Target(ByVal a_value As Object)
    Set This.PunkTarget = a_value
End Property

''' <summary>   Registers the listener. </summary>
Public Function Connect() As Boolean
Attribute Connect.VB_Description = "Registers the listener."

    cc_isr_Core.GuardClauses.GuardNullReference This.PunkTarget, VBA.Information.TypeName(Me), "Target is not set."
    ConnectToConnectionPoint Me, This.RefIID, True, This.PunkTarget, This.Cookie, 0&
    This.Connected = This.Cookie <> 0
    Connect = This.Connected
    
End Function

''' <summary>   De-registers the listener. </summary>
Public Function Disconnect() As Boolean
Attribute Disconnect.VB_Description = "De-registers the listener."

    If Not This.Connected Then Exit Function
    ConnectToConnectionPoint Me, This.RefIID, False, This.PunkTarget, This.Cookie, 0&
    This.Connected = False
    Disconnect = True

End Function

''' <summary>   A callback that handles MSForms.Control.AfterUpdate events for the registered target control. </summary>
Public Sub OnAfterUpdate()
Attribute OnAfterUpdate.VB_Description = "A callback that handles MSForms.Control.AfterUpdate events for the registered target control."
Attribute OnAfterUpdate.VB_UserMemId = -2147384832
    Dim p_handler As IHandleControlEvents
    For Each p_handler In This.Handlers
        p_handler.HandleAfterUpdate
    Next
End Sub

''' <summary>   A callback that handles MSForms.Control.BeforeUpdate events for the registered target control. </summary>
Public Sub OnBeforeUpdate(ByVal a_cancel As MSForms.IReturnBoolean)
Attribute OnBeforeUpdate.VB_Description = "A callback that handles MSForms.Control.BeforeUpdate events for the registered target control."
Attribute OnBeforeUpdate.VB_UserMemId = -2147384831
    
    Dim p_cancel As Boolean
    p_cancel = a_cancel.Value
    
    Dim p_handler As IHandleControlEvents
    For Each p_handler In This.Handlers
       
        p_handler.HandleBeforeUpdate p_cancel
        a_cancel.Value = p_cancel
        If p_cancel Then Exit For
    
    Next
    
End Sub

''' <summary>   A callback that handles MSForms.Control.Exit events for the registered target control. </summary>
Public Sub OnExit(ByVal a_cancel As MSForms.IReturnBoolean)
Attribute OnExit.VB_Description = "A callback that handles MSForms.Control.Exit events for the registered target control."
Attribute OnExit.VB_UserMemId = -2147384829
    
    Dim p_cancel As Boolean
    p_cancel = a_cancel.Value
    
    Dim p_handler As IHandleControlEvents
    For Each p_handler In This.Handlers
        p_handler.HandleExit p_cancel
        a_cancel.Value = p_cancel
        If p_cancel Then Exit For
    Next
    
End Sub

''' <summary>   A callback that handles MSForms.Control.Enter events for the registered target control. </summary>
Public Sub OnEnter()
Attribute OnEnter.VB_Description = "A callback that handles MSForms.Control.Enter events for the registered target control."
Attribute OnEnter.VB_UserMemId = -2147384830
    
    Dim p_handler As IHandleControlEvents
    For Each p_handler In This.Handlers
        p_handler.HandleEnter
    Next
    
End Sub

''' <summary>   A callback that handles MSForms.Control.Change events for the registered target control. </summary>
''' <remarks>   VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox, replaced
''' OnKeyPress in TextBox</remarks>
Public Sub OnChange()
    Dim p_handler As IHandleControlEvents
    For Each p_handler In This.Handlers
        p_handler.HandleChange
    Next
End Sub

''' <summary>   Registers the specified object to handle the relayed control events. </summary>
Public Sub RegisterHandler(ByVal a_handler As IHandleControlEvents)
Attribute RegisterHandler.VB_Description = "Registers the specified object to handle the relayed control events."
    This.Handlers.Add a_handler
End Sub

Private Sub Class_Initialize()
    Set This.Handlers = New Collection
    This.RefIID.Data1 = &H20400
    This.RefIID.Data4(0) = &HC0
    This.RefIID.Data4(7) = &H46
End Sub

Private Sub Class_Terminate()
    If This.Cookie <> 0 Then Disconnect
End Sub

Private Sub IControlEvents_OnAfterUpdate()
    OnAfterUpdate
End Sub

Private Sub IControlEvents_OnBeforeUpdate(ByVal a_cancel As MSForms.IReturnBoolean)
    OnBeforeUpdate a_cancel
End Sub
' VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox control, basically replacing OnKeyPress in TextBox
Private Sub IControlEvents_OnChange()
    OnChange
End Sub

Private Sub IControlEvents_OnEnter()
    OnEnter
End Sub

Private Sub IControlEvents_OnExit(ByVal a_cancel As MSForms.IReturnBoolean)
    OnExit a_cancel
End Sub

Private Sub IControlEvents_RegisterHandler(ByVal a_handler As IHandleControlEvents)
    RegisterHandler a_handler
End Sub
