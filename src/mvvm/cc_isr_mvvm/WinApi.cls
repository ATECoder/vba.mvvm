VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Windows API"
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Provides Windows API"
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As LongPtr) As Long
#Else
    Private Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0               'Horizontal Resolution
Private Const SM_CYSCREEN As Long = 1               'Vertical Resolution
Private Const LOGPIXELSX As Long = 88               'Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72          'A point is defined as 1/72 inches

Public Property Get HorizontalResInPixel() As Long
    HorizontalResInPixel = GetSystemMetrics32(SM_CXSCREEN) * PointsPerPixel
End Property

Public Property Get VerticalResInPixel() As Long
    VerticalResInPixel = GetSystemMetrics32(SM_CYSCREEN) * PointsPerPixel
End Property

''' <summary>   Get Points Per Pixel Screen resolution. </summary>
Public Function PointsPerPixel() As Double

    #If VBA7 Then
        Dim p_handleContext As LongPtr
        Dim p_dotsPerInch As LongPtr
    #Else
        Dim p_handleContext As Long
        Dim p_dotsPerInch As Long
    #End If

    p_handleContext = GetDC(0)
    p_dotsPerInch = GetDeviceCaps(p_handleContext, LOGPIXELSX)
    PointsPerPixel = POINTS_PER_INCH / p_dotsPerInch
    ReleaseDC 0, p_handleContext

End Function

