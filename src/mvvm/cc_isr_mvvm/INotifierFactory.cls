VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotifierFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An abstract factory that creates INotifyPropertyChanged helper objects."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An abstract factory that creates INotifyPropertyChanged helper objects. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Creates and returns a new notifier object. </summary>
Public Function Create() As INotifyPropertyChanged
Attribute Create.VB_Description = "Creates and returns a new notifier object."
End Function
