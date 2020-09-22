VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBAnimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IObject

Private mEyesCount             As Long
Private mLegsCount             As Long

Sub Move(lDistance As Double)

End Sub

Sub Bite()

End Sub

Public Property Get EyesCount() As Long
  EyesCount = mEyesCount
End Property

Public Property Let EyesCount(ByRef Newval As Long)
  mEyesCount = Newval
End Property

Public Property Get LegsCount() As Long
  LegsCount = mLegsCount
End Property

Public Property Let LegsCount(ByRef Newval As Long)
  'mLegsCount = Newval
End Property

Private Sub Class_Initialize()
  g.Count = g.Count + 1
End Sub

Private Sub Class_Terminate()
  g.Count = g.Count - 1
End Sub

Private Function IObject_ToString() As String
  IObject_ToString = "VBAnimal"
End Function
