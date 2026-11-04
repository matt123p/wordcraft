Attribute VB_Name = "Package"
Option Explicit
Public sink As AppSink
Public Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = " \n14"
Dim content As Object
Set content = ActiveDocument.Content
With content
.Text = "package"
End With
On Error Resume Next
Dim value As Long
value = 1 / 0
ActiveDocument.Content.InsertAfter CStr(Err.Number)
End Sub
Public Sub Auto_Open()
ActiveDocument.Content.Text = "must not auto-run"
End Sub
Public Sub Connect()
Set sink = New AppSink
sink.Connect
End Sub
