Attribute VB_Name = "a_number_processing"
Sub ReplaceANumbersWithUIDs()
    Dim find_a_numbers As New RegExp
    Dim non_digits As New RegExp
    Dim sheet As Worksheet
    Dim cell As Range
    Dim match As Object
    Dim prevUid As Long
    Dim a_number_2_uid As Object
    Dim uid As String
    Dim matches As Object
    Dim a_number As String
    Dim filepath As String
    
    ' The serializtion path of the A-number to UID map is hard coded here, but
    ' could be extracted as parameter
    filepath = "./a_number_2_uid.txt"
    
    ' Set-up regex for finding A-numbers
    find_a_numbers.Pattern = "[aA]?#?-?\d{2,3}[- ]?\d{3}[- ]?\d{3}\b"
    find_a_numbers.Global = True
    
    ' Set-up regex for finding non-digits
    non_digits.Pattern = "\D"
    non_digits.Global = True
    
    ' LoadCreate a dictionary to store mappings between A-numbers and UIDs
    Set a_number_2_uid = LoadDictionaryFromFile(filepath)
    
    ' Initialize prevUid for uid generation
    If a_number_2_uid.Count = 0 Then
        prevUid = -1
    Else
        prevUid = Application.Max(a_number_2_uid.items)
    End If
    
    ' Loop over all cells in all sheets in the workbook
    For Each sheet In ActiveWorkbook.Worksheets
        For Each cell In sheet.UsedRange
            ' Check if the cell contains text
            If Not IsEmpty(cell.value) And IsString(cell.value) Then
                ' For cells containing text, find the A-numbers
                For Each match In find_a_numbers.Execute(cell.value)
                    ' Canonicalize the A-number
                    a_number = CLng(non_digits.Replace(match.value, ""))
                    
                    ' Check if the A-number is already in the dictionary
                    If Not a_number_2_uid.Exists(a_number) Then
                        ' Generate a unique identifier (UID) for any new A-numbers
                        prevUid = prevUid + 1
                        a_number_2_uid.Add a_number, prevUid
                    End If
                    
                    ' Get the UID for the A-number from the dictionary
                    uid = "UID-" & CStr(a_number_2_uid(a_number))
                    
                    ' Replace the matched A-number in the cell with the UID
                    cell.value = Replace(cell.value, match.value, uid, 1, -1, vbTextCompare)
                Next
            End If
        Next cell
    Next sheet
    
    SaveDictionaryToFile a_number_2_uid, filepath
End Sub

Function IsString(value As Variant) As Boolean
    IsString = VarType(value) = vbString
End Function

Function LoadDictionaryFromFile(filepath As String) As Object
    Dim fso As New FileSystemObject
    Dim file As Object
    Dim dict As New Dictionary
    Dim line As String
    Dim key As String
    Dim value As String
    
    ' Check if the file exists
    If fso.FileExists(filepath) Then
        
        ' Open the text file for reading
        Set file = fso.OpenTextFile(filepath, 1)
        
        ' Read each line from the file
        Do While Not file.AtEndOfStream
            line = file.ReadLine
            
            ' Split the line into key and value
            key = Split(line, ":")(0)
            value = Split(line, ":")(1)
            
            ' Add the key-value pair to the dictionary
            dict.Add key, CLng(value)
        Loop
        
        ' Close the file
        file.Close
        
        ' Return the loaded dictionary
        Set LoadDictionaryFromFile = dict
    Else
        ' Return the empty dictionary
        Set LoadDictionaryFromFile = dict
    End If
End Function

Sub SaveDictionaryToFile(dict As Object, filepath As String)
    Dim fso As New Scripting.FileSystemObject
    Dim file As Object
    Dim key As Variant
    Dim value As Variant
    Dim dictString As String

    ' Create a new text file
    Set file = fso.CreateTextFile(filepath, True)
    
    ' Serialize the dictionary into a string
    For Each key In dict.Keys
        dictString = dictString & key & ":" & CStr(dict(key)) & vbCrLf
    Next key
    
    ' Write the serialized dictionary to the file
    file.Write dictString
    
    ' Close the file
    file.Close
End Sub

