VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GUIComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ThingComparer"
'@PredeclaredId
Option Explicit

Implements IComparer

Private Type TGUIComparer
    View As IView
End Type

Private this As TGUIComparer

Private Property Get View() As IView
    Set View = this.View
End Property

Friend Property Set View(ByVal a_value As IView)
    Set this.View = a_value
End Property

Public Function Create(Optional ByVal ViewFactory As IThingComparisonViewFactory) As GUIComparer
    
    GuardClauses.GuardNonDefaultInstance Me, GUIComparer

    Dim result As GUIComparer
    Set result = New GUIComparer
    
    Dim Context As IAppContext
    Set Context = cc_isr_MVVM.Factory.NewAppContext.Initialize()
    
    Dim ViewModel As ThingComparisonViewModel
    Set ViewModel = New ThingComparisonViewModel
    
    If ViewFactory Is Nothing Then Set ViewFactory = TextRepresentableThingsView
    Set result.View = ViewFactory.Create(Context, ViewModel)
        
    Set Create = result
    
End Function

Private Function IComparer_Compare(ByVal x As Variant, ByVal y As Variant) As Long

    GuardClauses.GuardDefaultInstance Me, GUIComparer 'cache won't be clear in the default instance
    GuardClauses.GuardNullReference View, a_Message:="Class must be Created with the .Create Method not `New`"
    
    
    'short circuit default condition
    'our list has no dupes but arraylist asserts x.compareTo(x) = 0
    If x = y Then IComparer_Compare = 0: Exit Function
           
    'only need to check x & y as reverse is already in cache by default
    Static cache As New Scripting.Dictionary
    If TryGetComparisonResult(cache, x, y, IComparer_Compare) Then Exit Function

    With View
        Dim ViewModel As ThingComparisonViewModel
        Set ViewModel = .ViewModel
        ViewModel.ThingX = x
        ViewModel.ThingY = y
    
    
        'show dialog so we can capture cancellation
        If Not .ShowDialog Then Err.Raise 5, , "Cancelled"
        
        Select Case ViewModel.Choice
            Case xChosen
                IComparer_Compare = -1
                
            Case yChosen
                IComparer_Compare = 1
                
            Case Else
                Err.Raise 5, , "Invalid Selection"
                
        End Select
    End With
    
    CacheComparisonResult cache, x, y, IComparer_Compare
    
End Function

Private Function TryGetComparisonResult(ByVal cache As Dictionary, ByVal x As Variant, ByVal y As Variant, ByRef outValue As Long) As Boolean
    Dim key As String
    key = x & vbNullChar & y
    If Not cache.Exists(key) Then Exit Function
    TryGetComparisonResult = True
    outValue = cache.item(key)
End Function

Private Sub CacheComparisonResult(ByVal cache As Dictionary, ByVal x As Variant, ByVal y As Variant, ByVal result As Long)
    cache.Add x & vbNullChar & y, result
    cache.Add y & vbNullChar & x, -result
End Sub

