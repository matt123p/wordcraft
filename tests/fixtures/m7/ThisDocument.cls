Attribute VB_Name = "ThisDocument"
Attribute VB_PredeclaredId = True
Private Sub Document_Open()
Content.Text = "owner"
Me.Content.InsertAfter "-bound"
End Sub
