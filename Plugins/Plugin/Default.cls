VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Default"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public HostForm As Form
Public ID As String

Private Sub Class_Terminate()
Set HostForm = Nothing
End Sub

Public Function PerformSomething(What As Variant)
Set Module1.HForm = HostForm
Module1.IDX = ID
Select Case LCase(What(0))
Case "functions"

End Select
End Function
