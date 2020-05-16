Attribute VB_Name = "GuideMakerModule"
Sub GuideMaker()
    If ActivePage Is Nothing Then
        MsgBox "There is no open document"
        Exit Sub
    End If
    
    GuideMk.Show

End Sub
