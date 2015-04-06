Attribute VB_Name = "Main"
'This function uses the regex library to parse the data we want out of a string with avatar form tags.
' Returns the data encased within the tags.
Public Function getValue(ByVal line As String) As String

    'Define a local variable in this scope to hold the regex matches for fileName
    Dim fieldNameRegExMatches As Object
    
    'Make sure the Microsoft VBScript Regular Expression reference is added to the project.
    'This will be used to pull relevent data from a string
    Dim fieldNameRegEx As New VBScript_RegExp_55.RegExp
    fieldNameRegEx.Pattern = ">(.+)<"
    fieldNameRegEx.IgnoreCase = True
    fieldNameRegEx.Global = False
    
    'Test the fieldNameRegEx against the regular expression
    If fieldNameRegEx.test(line) Then
        'Get the field name from the line
        Set fieldNameRegExMatches = fieldNameRegEx.Execute(line)
        
        If (fieldNameRegExMatches.Count <> 0) Then
            line = fieldNameRegExMatches.Item(0).SubMatches.Item(0)
        End If
    Else 'end regEx test
        line = ""
    End If
    
    'return the string we found
    getValue = line
    
End Function


'This is the main subroutine for the site specific modeling spec tool. All of the user defined
' functions above will be called in here. This macro will open a .txt file and parse out the information
' needed t build a spec sheet based off of the Avatar site specific dump file.
Sub mainSub()

    Call clearFields
        
    'only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    'set the title of the open file dialog
    Application.FileDialog(msoFileDialogOpen).Title = "Select Avatar site specific dump file"
    'remove all other filters
    Call Application.FileDialog(msoFileDialogOpen).Filters.Clear
    'add a custom filter
    Call Application.FileDialog(msoFileDialogOpen).Filters.Add("Text Files Only", "*.txt")
    
    'if the dialog returns anything
    If Application.FileDialog(msoFileDialogOpen).Show <> 0 Then
        'set up some local shit
        Dim filePath As String
        Dim dumpFile As Integer: dumpFile = 1
        Dim specRowIndex As Integer: specRowIndex = 17
        Dim excludeFromDCI As String: excludeFromDCI = "1"
        
        'set up an array which is used to look up the type of field
        Dim fieldTypes(1 To 30) As String
        fieldTypes(1) = "Single Response Dictionary"
        fieldTypes(2) = "Multiple Response Dictionary"
        fieldTypes(3) = "Staff"
        fieldTypes(4) = "Free Text"
        fieldTypes(5) = "Scrolling Free Text"
        fieldTypes(6) = "00000"
        fieldTypes(7) = "Axis I"
        fieldTypes(8) = "Axis II"
        fieldTypes(9) = "Axis III"
        fieldTypes(10) = "Date"
        fieldTypes(11) = "00000"
        fieldTypes(12) = "Label"
        fieldTypes(13) = "00000"
        fieldTypes(14) = "00000"
        fieldTypes(15) = "Service Code"
        fieldTypes(16) = "00000"
        fieldTypes(17) = "Time"
        fieldTypes(18) = "00000"
        fieldTypes(19) = "00000"
        fieldTypes(20) = "00000"
        fieldTypes(21) = "00000"
        fieldTypes(22) = "00000"
        fieldTypes(23) = "00000"
        fieldTypes(24) = "00000"
        fieldTypes(25) = "Sign"
        
        'get the filepath of the selected file
        filePath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        'print the path of the file so we know where this spec data came from
        Range("B9") = filePath
        'open the file and start reading it's goodies
        Open filePath For Input As #dumpFile
        
        'loop through each line of the file until we hid the end of file
        Do Until EOF(1)
            'get the line and put it in textLine string
            Line Input #dumpFile, textLine
            
            If (InStr(textLine, "<formname>")) Then
                Range("A13") = getValue(textLine)
            ElseIf (InStr(textLine, "<entitydatabase>")) Then
                Range("B13") = getValue(textLine)
            ElseIf (InStr(textLine, "<optionid>")) Then
                Range("C13") = getValue(textLine)
            ElseIf (InStr(textLine, "<excludefromdci>")) Then
                excludeFromDCI = getValue(textLine)
            End If
            
            If (excludeFromDCI = "0") Then
                If (InStr(textLine, "<promptorder>")) Then
                    Cells(specRowIndex, "A") = getValue(textLine)
                ElseIf (InStr(textLine, "<fieldtype>")) Then
                    Dim gotValue As String: gotValue = getValue(textLine)
                    
                    If IsNumeric(gotValue) Then
                        Cells(specRowIndex, "B") = fieldTypes(CInt(gotValue))
                        'Cells(specRowIndex, "B") = gotValue
                    Else
                        Cells(specRowIndex, "B") = gotValue
                    End If
                ElseIf (InStr(textLine, "<fieldlabel>")) Then
                    Cells(specRowIndex, "C") = getValue(textLine)
                ElseIf (InStr(textLine, "<initrequired>")) Then
                    Cells(specRowIndex, "D") = getValue(textLine)
                ElseIf (InStr(textLine, "<initenabled>")) Then
                    Cells(specRowIndex, "E") = getValue(textLine)
                ElseIf (InStr(textLine, "</promptdata>")) Then
                    specRowIndex = specRowIndex + 1
                End If
            End If
            
        Loop 'next line
        
        'close the file because we have to
        Close #dumpFile
        
    End If
    
End Sub


'This will reset the spec sheet, clearing all fields that are automatically filled by main()
Sub clearFields()

    'ActiveSheet.Unprotect
    
    Range("B9").ClearContents
    Range("A13:C13").ClearContents
    Range(Cells(17, "A"), Cells(ActiveSheet.UsedRange.Rows.Count + 1, ActiveSheet.UsedRange.Columns.Count)).ClearContents
    
    'delete previous IntegrationTest sheets if they exist, do nothing if they don't.
    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("IntegrationTest").Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
    
    'ActiveSheet.Protect
    
End Sub

'This will generate an integration test script that folks will use to make sure their modeled forms
' are working properly.
Sub generateTestScript()
    
    'turn off screen updating so the screen does look all flashy and stupid, also, macros just run faster
    ' when they don't have to render all of the actions they take.
    Application.ScreenUpdating = False
    
    'delete previous IntegrationTest sheets if they exist, do nothing if they don't.
    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("IntegrationTest").Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
    
    'add a worksheet for the integration test script
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "IntegrationTest"
    
    'copy the Netsmart logo so the sheet looks nice :)
    'for some reason the picture is called "picture 3" instead of logically naming it as the damn filename
    Sheets("FormSpec").Shapes("Picture 3").Copy
    Sheets("IntegrationTest").PasteSpecial
    
    'set up header stuffs
    Cells(5, "B") = "Test case:"
    Cells(6, "B") = "Description:"
    Cells(7, "B") = "Number:"
    Cells(8, "B") = "Overall pass criteria:"
    Cells(9, "B") = "Tester name:"
    Cells(10, "B") = "Test data used:"
    Cells(11, "B") = "Comments:"
    Cells(12, "B") = "Date/time run:"
    Cells(13, "B") = "Status:"
    
    'allign right, white, bold, and background
    Range("B5:B13").HorizontalAlignment = xlRight
    Range("B5:B13").Interior.Color = RGB(67, 172, 106)
    Range("B5:B13").Font.Bold = True
    Range("B5:B13").Font.Color = RGB(255, 255, 255)

    'set up column header for test table
    Cells(16, "A") = "Step"
    Cells(16, "B") = "Action"
    Cells(16, "C") = "Expected result"
    Cells(16, "D") = "Pass"
    Cells(16, "E") = "Fail"
    Cells(16, "F") = "N/A"
    Cells(16, "G") = "Comments"
    
    'background blue, font bold and white
    Range("A16:G16").Interior.Color = RGB(0, 172, 226)
    Range("A16:G16").Font.Bold = True
    Range("A16:G16").Font.Color = RGB(255, 255, 255)
    
    'adjust width to allow for lots o' text in the test table
    Columns("A").ColumnWidth = 4
    Columns("B").ColumnWidth = 30
    Columns("C").ColumnWidth = 30
    Columns("D").ColumnWidth = 4
    Columns("E").ColumnWidth = 4
    Columns("F").ColumnWidth = 4
    Columns("G").ColumnWidth = 30
    
    'loop through each row until we find a null cell
    Dim i As Integer: i = 1
    Do While Not IsEmpty(Sheets("FormSpec").Cells(i + 16, "A"))
    
        'set some instance variable so we don't have to retrieve these values twelve times
        Dim rowIndex As Integer: rowIndex = i + 16
        Dim rowFieldType As String: rowFieldType = Sheets("FormSpec").Cells(rowIndex, "B")
        Dim rowFieldLabel As String: rowFieldLabel = Sheets("FormSpec").Cells(rowIndex, "C")
        Dim action As Range: Set action = Sheets("IntegrationTest").Cells(rowIndex, "B")
        Dim expectedResult As Range: Set expectedResult = Sheets("IntegrationTest").Cells(rowIndex, "C")
        
        Sheets("IntegrationTest").Cells(rowIndex, "A").Value = i
        action.WrapText = True
        expectedResult.WrapText = True
        
        'check the fieldType if the row in FromSpec
        If (rowFieldType = "Single Response Dictionary") Then
            'set the action and expected value for the integration test
            action.Value = "Select one option from " + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Multiple Response Dictionary") Then
            action.Value = "Select multiple options from " + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Staff") Then
            action.Value = "Select a staff member from " + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Free Text") Then
            action.Value = "Type in a value for " + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Scrolling Free Text") Then
            action.Value = "Type in a value for " + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Axis I") Then
            action.Value = "I don't know what this is" + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Axis II") Then
            action.Value = "I don't know what this is" + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Axis III") Then
            action.Value = "I don't know what this is" + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Date") Then
            action.Value = "Select a date for " + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Service Code") Then
            action.Value = "Select a service code from " + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Time") Then
            action.Value = "Enter a time for " + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        ElseIf (rowFieldType = "Sign") Then
            action.Value = "I don't know what this is" + rowFieldLabel
            expectedResult.Value = rowFieldLabel + " will be recorded in the form."
            
        End If
        
        i = i + 1
    Loop
    
    
    
End Sub








