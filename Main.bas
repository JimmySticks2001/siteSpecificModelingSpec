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
        Dim fieldTypes(1 To 20) As String
        fieldTypes(1) = "Single Response Dictionary"
        fieldTypes(2) = "Multiple Response Dictionary"
        fieldTypes(3) = "Staff"
        fieldTypes(4) = "Free Text"
        fieldTypes(5) = "Scrolling Free Text"
        fieldTypes(10) = "Date"
        fieldTypes(12) = "Label"
        fieldTypes(15) = "Service Code"
        fieldTypes(17) = "Time"
        
        
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
                    Cells(specRowIndex, "B") = fieldTypes(CInt(getValue(textLine)))
                    'Cells(specRowIndex, "B") = getValue(textLine)
                ElseIf (InStr(textLine, "<fieldlabel>")) Then
                    Cells(specRowIndex, "D") = getValue(textLine)
                ElseIf (InStr(textLine, "<initrequired>")) Then
                    Cells(specRowIndex, "F") = getValue(textLine)
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

    ActiveSheet.Unprotect
    
    Range("B9").ClearContents
    Range("A13:C13").ClearContents
    Range(Cells(17, "A"), Cells(ActiveSheet.UsedRange.Rows.Count + 1, ActiveSheet.UsedRange.Columns.Count)).ClearContents
    
    ActiveSheet.Protect
    
    
End Sub

