'Paste this code in ThisOutlookSession
Private Sub Application_ItemLoad(ByVal Item As Object)
    On Error GoTo ErrHandler
    If TypeName(Outlook.ActiveExplorer.Selection(1)) = "MailItem" Then
            Call FindAppts
            'saveAttachtoDisk Item
    End If
    Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub
