Attribute VB_Name = "CreateMeetingMinutes"
Function insertField(ByVal objSelection As Object, ByVal field As String, ByVal val As String)
    With objSelection
            .Font.Bold = True
            .Font.Size = 12
            .ParagraphFormat.SpaceAfter = 0
        End With
        objSelection.TypeText field & " "
        With objSelection
            .Font.Bold = False
            .Font.Size = 12
        End With
        objSelection.TypeText val
            objSelection.TypeParagraph
End Function
Sub createMinutes()

    Dim wApp As Object, wDoc As Object, olkMsg As Object, objSelection As Object
    
    On Error Resume Next
    Set wApp = GetObject(, "Word.Application")
    If wApp Is Nothing Then
        Set wApp = CreateObject("Word.Application")
    End If
    Set wDoc = wApp.Documents.Add(, , , True)
    Set objSelection = wDoc.ActiveWindow.Selection
    
    With objSelection
        .Font.Bold = True
        .Font.Size = 16
        .ParagraphFormat.SpaceAfter = 18
    End With
    objSelection.TypeText "Meeting Minutes"
        objSelection.TypeParagraph
    
    
    For Each olkMsg In Application.ActiveExplorer.Selection
        ' SUBJECT
        Call insertField(objSelection, "Subject:", olkMsg.Subject)
        ' IMPORTANCE
        Call insertField(objSelection, "Importance:", olkMsg.Importance)
        ' LOCATION
        Call insertField(objSelection, "Location:", olkMsg.Location)
        ' TIME
        Call insertField(objSelection, "Start:", olkMsg.StartInStartTimeZone)
        ' Organizer
        Call insertField(objSelection, "Organizer:", olkMsg.Organizer)
        ' ATTENDEES
        Call insertField(objSelection, "Required:", olkMsg.RequiredAttendees)
        ' ATTENDEES2
        Call insertField(objSelection, "Optional:", olkMsg.OptionalAttendees)
        
    With objSelection
        .Font.Bold = True
        .Font.Italic = True
        .Font.Size = 12
        .ParagraphFormat.SpaceBefore = 12
        .ParagraphFormat.SpaceAfter = 0
        .TypeText vbTab & "Present:"
            
        .Font.Bold = False
        
        .TypeText " "
        .TypeParagraph
        
        .Font.Bold = True
        .Font.Italic = False
        .Font.Size = 14
        .ParagraphFormat.SpaceBefore = 27
        .ParagraphFormat.SpaceAfter = 18
        .TypeText "Results:"
        .TypeParagraph
        
        .Font.Bold = False
        .Font.Size = 12
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        
        .TypeText vbTab
    End With

    Next olkMsg
    
        
    wApp.Visible = True
    wApp.Activate
    
    ' UNCOMENT IF YOU WANT THE SAVEAS Dialog TO POP UP as the file is created
    'wApp.FileDialog(msoFileDialogSaveAs).InitialFileName _
    '= "C:\Users\" & Environ("username") & "\Documents\MeetingMinutes"
    'wApp.FileDialog(msoFileDialogSaveAs).Show
    
    
    
    Set wApp = Nothing
    Set wDoc = Nothing
    Set olkMsg = Nothing
    Set objSelection = Nothing
   
End Sub
