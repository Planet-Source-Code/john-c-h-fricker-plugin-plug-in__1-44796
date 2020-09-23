Attribute VB_Name = "Plugin"
Private Type PluginX
POBJ As Object
Name As String
Info As String
End Type
Dim Plugins(1 To 9000) As PluginX

Private Function PPos(Name As String) As Integer
For i = 1 To 9000
 If LCase(Plugins(i).Name) = LCase(Name) Then PPos = i: Exit Function
Next
End Function

Public Function OpenPlugin(ID As String, Plugin As String, HostForm As Form) As Integer
If PPos(ID) Then OpenPlugin = 1: Exit Function
On Error GoTo err
With Plugins(PPos(""))
.Name = ID
.Info = Plugin
Set .POBJ = CreateObject(Plugin & ".Default")
Set .POBJ.HostForm = Form1
.POBJ.ID = .Name
End With
Exit Function
err:
OpenPlugin = 1
End Function

Public Function ClosePlugin(ID As String)
If PPos(ID) < 1 Then Exit Function
With Plugins(PPos(ID))
.Name = ""
.Info = ""
Set .POBJ = Nothing
End With
End Function

Public Function PushInfo(ID As String, Info As Variant)
If PPos(ID) < 1 Then Exit Function
With Plugins(PPos(ID))
.POBJ.performsomething Info
End With
End Function

Public Function ReturnedInfo(ID As String, What As Variant)
Select Case LCase(ID)
Case "primary"
    Select Case LCase(What(0))
    Case "rgb"
    Form1.BackColor = Val(What(1))
    End Select
End Select
End Function
