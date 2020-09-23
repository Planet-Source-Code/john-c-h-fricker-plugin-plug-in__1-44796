Attribute VB_Name = "Module1"
Public HForm As Form
Public IDX As String

Public Function ReturnInfo(What As Variant)
Module1.HForm.MsgPush What, Module1.IDX
End Function
