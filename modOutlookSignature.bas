'Import this file into Microsoft Outlook VBA (Press Alt+F11 from within Outlook)
Option Explicit
Const CAL_TEXT = "Meeting(s) Today: "
Const SIGNATURE_FILE_NAME = "2018"
Sub FindAppts()
    Dim myStart, myEnd, myEndDte As Date
    Dim oCalendar As Outlook.Folder
    Dim oItems As Outlook.Items
    Dim oItemsInDateRange As Outlook.Items
    Dim oFinalItems As Outlook.Items
    Dim oAppt As Outlook.AppointmentItem
       
    'myStart = Format("04/18/2014", "mm/dd/yyyy hh:mm AMPM")
    myStart = Format(Date, "mm/dd/yyyy hh:mm AMPM")
    myEndDte = DateAdd("d", 1, myStart)
    myEnd = Format(myEndDte, "mm/dd/yyyy hh:mm AMPM")
          
    'Construct filter for the next 1-day date range
    Dim strRestriction As String
    strRestriction = "[Start] >= '" & myStart _
    & "' AND [End] <= '" & myEnd & "'" _
    '& " AND ([End] - [Start]) >=10"
    
    Set oCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
    Set oItems = oCalendar.Items
    
    'Including recurrent appointments requires sorting by the Start property
    oItems.IncludeRecurrences = True
    oItems.Sort "[Start]"
    
    'Restrict the Items collection for the 1-day date range
    Set oItemsInDateRange = oItems.Restrict(strRestriction)
    
    'Sort and print final results
    oItemsInDateRange.Sort "[Start]"
    Dim intDuration As Integer, intCount As Integer, col As Collection
    Dim arrStr(96) As String
    For Each oAppt In oItemsInDateRange
        If oAppt.AllDayEvent = False Then
            intDuration = DateDiff("n", oAppt.Start, oAppt.End)
            intCount = intDuration / 4
            If intDuration >= 10 Then
                populateData oAppt.Start, oAppt.End, arrStr()
            End If
        End If
    Next
    
    Dim i As Integer, strCalValue As String, strCalText As String, strTime As String
    strCalText = CAL_TEXT
    Dim intHr As Integer
    intHr = 0
    For i = 0 To 95
        
        If i Mod 4 = 0 Then
            If i = 0 Then '12 am
                strTime = 12 & "am"
            ElseIf i / 48 < 1 Then '1am to 11am
                strTime = i / 4 & "am"
            ElseIf i / 48 = 1 Then '12pm
                strTime = i / 4 & "pm"
            ElseIf i / 48 = 2 Then '12am or '12pm
                strTime = (i / 4) - 12 & "am"
            Else '1pm to 11pm
                strTime = (i / 4) - 12 & "pm"
            End If
            strCalValue = strCalValue & "<sup>" & strTime & "</sup>"
        End If
        
        If (arrStr(i) <> "") Then
            strCalValue = strCalValue & arrStr(i)
        Else
            strCalValue = strCalValue & "&nbsp;"
        End If
    Next i
    
    Call removeBlankHours(strCalValue)
    If Trim(strCalValue) = "" Then
        strCalText = ""
        strCalValue = ""
    End If
    
    Dim strSignFileToUse As String, strSignFileTemplate As String
    strSignFileToUse = Environ("appdata") & _
                "\Microsoft\Signatures\" & SIGNATURE_FILE_NAME & ".htm"
    strSignFileTemplate = Environ("appdata") & _
                "\Microsoft\Signatures\" & SIGNATURE_FILE_NAME & "Template.htm"
    
    'Read the signature
    Dim intFile As Integer
    intFile = FreeFile
    Dim strSignContent  As String
    Open strSignFileTemplate For Input As #intFile
    strSignContent = Input$(LOF(intFile), intFile)
    Close #intFile
    
    strSignContent = Replace(strSignContent, SIGNATURE_FILE_NAME & "Template", SIGNATURE_FILE_NAME)
    strSignContent = Replace(strSignContent, "{calendarText}", strCalText)
    strSignContent = Replace(strSignContent, "{calendarValue}", strCalValue)
    
    'Write the signature with dynamic part
    intFile = FreeFile
    Open strSignFileToUse For Output As #intFile
    Print #intFile, strSignContent
    Close #intFile
    
End Sub
Private Sub populateData(dteStart As Date, dteEnd As Date, ByRef arrStr() As String)
    Dim intMinutesStart As Integer, intMinutesEnd As Integer
    
    intMinutesStart = DateDiff("n", DateValue(dteStart), dteStart)
    intMinutesStart = intMinutesStart - (intMinutesStart Mod 15)
    
    intMinutesEnd = DateDiff("n", DateValue(dteEnd), dteEnd)
    If (intMinutesEnd Mod 15 <> 0) Then
        intMinutesEnd = intMinutesEnd + 15 - (intMinutesEnd Mod 15)
    End If
    
    Dim i As Integer
    For i = intMinutesStart To intMinutesEnd - 1 Step 15
        arrStr(i / 15) = "|"
    Next i
End Sub

Private Sub removeBlankHours(ByRef strContent As String)
    
    Dim i As Integer
    Dim strHr As String
    
    For i = 0 To 24
        If i = 0 Then
            strHr = "12am"
        ElseIf i > 0 And i < 12 Then
            strHr = CStr(i) & "am"
        ElseIf i = 12 Then
            strHr = "12pm"
        ElseIf i > 12 And i < 24 Then
            strHr = CStr(i - 12) & "pm"
        ElseIf i = 24 Then
            strHr = "12am"
        End If
        strContent = Replace(strContent, "<sup>" & strHr & "</sup>&nbsp;&nbsp;&nbsp;&nbsp;", "")
    Next i
End Sub

