Attribute VB_Name = "Module1"
Sub Pause(interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

'Function PrivateMsg(strUser As String, strMsg As String)
Public Function Ping()
  With frmYbot
    .WS.SendData .HEADER & Chr(98) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
  End With
End Function
