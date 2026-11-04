Attribute VB_Name = "AppSink"
Private WithEvents app As Word.Application
Public Sub Connect()
Set app = Application
End Sub
Private Sub app_DocumentBeforeSave(ByVal Doc As Word.Document, SaveAsUI As Boolean, Cancel As Boolean)
Doc.Content.Text = "before-save"
Cancel = True
End Sub
