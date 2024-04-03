Sub PullEmailSubjectAndDateUsingUniqueSearchQuery()
    ' Declare variables for Outlook objects.
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olFolder As Object
    Dim selectedCells As Range
    Dim cell As Range
    Dim subjectColumn As Long
    Dim dateColumn As Long
    Dim mailItem As Object
    Dim found As Boolean
    Dim staticEmail As String
    Dim headerRow As Integer
    Dim subjectColumnExists As Boolean
    Dim dateColumnExists As Boolean
    Dim searchColumn As Long
    
    ' Static email address to check against.
    staticEmail = "groupemail@domain.tld"
    headerRow = 1 ' Assuming the header is in the first row.
    subjectColumnExists = False
    dateColumnExists = False
    
    ' Verify selection is in Column A (Unique ID/Search Query).
    If Not Intersect(Selection, Columns("A")) Is Nothing Then
        Set selectedCells = Selection
    Else
        MsgBox "Please select cells within Column A (ID)"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    ' Initialize Outlook.
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Specify the folder to search; default is Inbox.
    ' For a specific folder under Inbox, uncomment and modify the next line appropriately.
    Set olFolder = olNamespace.GetDefaultFolder(6) ' 6 refers to the Inbox.
    ' Set olFolder = olNamespace.GetDefaultFolder(6).Folders("SpecificFolder") ' Example for a specific folder.
    
    ' Check for the existence of "Email Subject" and "Email Date" columns.
    For searchColumn = 1 To Cells(headerRow, Columns.Count).End(xlToLeft).Column
        If Cells(headerRow, searchColumn).Value = "Email Subject" Then
            subjectColumnExists = True
            subjectColumn = searchColumn
        ElseIf Cells(headerRow, searchColumn).Value = "Email Date" Then
            dateColumnExists = True
            dateColumn = searchColumn
        End If
    Next searchColumn
    
    ' Create "Email Subject" column if it doesn't exist.
    If Not subjectColumnExists Then
        subjectColumn = Cells(headerRow, Columns.Count).End(xlToLeft).Column + 1
        Cells(headerRow, subjectColumn).Value = "Email Subject"
    End If
    
    ' Create "Email Date" column right after "Email Subject" if it doesn't exist.
    If Not dateColumnExists Then
        dateColumn = subjectColumn + 1
        Cells(headerRow, dateColumn).Value = "Email Date"
    End If
    
    ' Search emails for each selected cell value.
    For Each cell In selectedCells
        found = False
        
        For Each mailItem In olFolder.Items
            If InStr(1, mailItem.Body, cell.Value, vbTextCompare) > 0 Then
                Dim rec As Object
                For Each rec In mailItem.Recipients
                    If LCase(rec.Address) = LCase(staticEmail) Or LCase(rec.Name) = LCase(staticEmail) Then
                        ' Copy the email subject and date for the first found email.
                        Cells(cell.Row, subjectColumn).Value = mailItem.Subject
                        Cells(cell.Row, dateColumn).Value = mailItem.SentOn
                        found = True
                        Exit For
                    End If
                Next rec
                If found Then Exit For
            End If
        Next mailItem
        
        If Not found Then
            ' Mark as "Not Found" if email not found.
            Cells(cell.Row, subjectColumn).Value = "Not Found"
            Cells(cell.Row, dateColumn).Value = ""
        End If
    Next cell

    ' Notify user upon completion.
    MsgBox "Task completed. Outlook search results have been updated.", vbInformation, "Search Completed"
    Exit Sub

ErrorHandler:
    MsgBox "An error has occurred. Please check your Outlook setup and try again.", vbCritical, "Error"
    Set olApp = Nothing
    Set olNamespace = Nothing
    Set olFolder = Nothing
    Set selectedCells = Nothing
End Sub
