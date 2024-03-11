Attribute VB_Name = "a_number_processing"
Sub ReplaceANumbersWithUIDs()
    Dim regex As Object
    Dim cell As Range
    Dim match As Object
    Dim prevUid As Integer
    Dim a_number_2_uid As Object
    Dim uid As String
    Dim matches As Object
    Dim a_number As String
    Dim filepath As String
    
    ' The serializtion path of the A-number to UID map is hard coded here, but
    ' could be extracted as parameter
    filepath = "./a_number_2_uid.txt"
    
    ' Create a new regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Set the pattern of the regular expression object to match on A-numbers
    regex.Pattern = "[aA]?#?-?[0-9]{2,3}[- ]?[0-9]{3}[- ]?[0-9]{3}\b"
    regex.Global = True
    
    ' LoadCreate a dictionary to store mappings between A-numbers and UIDs
    Set a_number_2_uid = LoadDictionaryFromFile(filepath)
    
    ' Initialize prevUid for uid generation
    If a_number_2_uid.Count = 0 Then
        prevUid = -1
    Else
        prevUid = Application.Max(a_number_2_uid.items)
    End If
    
    
    Debug.Print "Set-up Complete"
    
    ' Loop over all cells in the active worksheet (`ActiveSheet.Range` could be be swapped with
    ' `Selection` to only process the selected/highlighted cells)
    For Each cell In ActiveSheet.UsedRange
        ' Check if the cell contains text
        If Not IsEmpty(cell.value) And IsString(cell.value) Then
            ' For cells containing text, find the A-numbers
            For Each match In regex.Execute(cell.value)
                ' Canonicalize the A-number
                a_number = CanonicalizeANumber(match.value)
                
                
                ' Check if the A-number is already in the dictionary
                If Not a_number_2_uid.Exists(a_number) Then
                    ' Generate a unique identifier (UID) for new A-numbers
                    prevUid = prevUid + 1
                    a_number_2_uid.Add a_number, prevUid
                End If
                
                ' Get the UID for the A-number from the dictionary
                uid = UIDToString(a_number_2_uid(a_number))
                
                ' Replace the matched word with the UID
                Debug.Print match.value
                Debug.Print a_number
                Debug.Print uid
                Debug.Print ""
                cell.value = Replace(cell.value, match.value, uid, 1, -1, vbTextCompare)
            Next
        End If
    Next cell
    
    SaveDictionaryToFile a_number_2_uid, filepath
End Sub

Function IsString(value As Variant) As Boolean
    ' Check if a value is a string
    IsString = VarType(value) = vbString
End Function


Function CanonicalizeANumber(inputString As String) As String
    Dim resultString As String
    Dim i As Integer
    Dim charToRemove As String
    
    ' Define characters to remove
    Dim charactersToRemove As Variant
    charactersToRemove = Array("a", "-", "#", " ")
    
    ' Convert input string to lowercase
    resultString = LCase(inputString)
    
    ' Loop over the string removing characters
    For i = LBound(charactersToRemove) To UBound(charactersToRemove)
        charToRemove = charactersToRemove(i)
        
        ' Replace each occurrence of the character with an empty string
        resultString = Replace(resultString, charToRemove, "")
    Next i
    
    ' If the string only has 8 characters prepend a zero
    If Len(resultString) = 8 Then
        resultString = "0" & resultString
    End If
    
    ' Return the canonicalized A-number
    CanonicalizeANumber = resultString
End Function

Function UIDToString(uid As Integer) As String
    UIDToString = "UID-" & CStr(uid)
End Function

Function LoadDictionaryFromFile(filepath As String) As Object
    Dim fso As Object
    Dim file As Object
    Dim dict As Object
    Dim line As String
    Dim key As String
    Dim value As String
    
    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the file exists
    If fso.FileExists(filepath) Then
        ' Create a new dictionary
        Set dict = CreateObject("Scripting.Dictionary")
        
        ' Open the text file for reading
        Set file = fso.OpenTextFile(filepath, 1)
        
        ' Read each line from the file
        Do While Not file.AtEndOfStream
            line = file.ReadLine
            
            ' Split the line into key and value
            key = Split(line, ":")(0)
            value = Split(line, ":")(1)
            
            ' Add the key-value pair to the dictionary
            dict.Add key, CInt(value)
        Loop
        
        ' Close the file
        file.Close
        
        
        ' Return the dictionary
        Debug.Print "Loaded A-number to UID dictionary"
        Set LoadDictionaryFromFile = dict
    Else
        ' Return the empty dictionary
        Debug.Print "Created an empty A-number to UID dictionary"
        Set dict = CreateObject("Scripting.Dictionary")
        Set LoadDictionaryFromFile = dict
    End If
End Function

Sub SaveDictionaryToFile(dict As Object, filepath As String)
    Dim fso As Object
    Dim file As Object
    Dim key As Variant
    Dim value As Variant
    Dim dictString As String
    
    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
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

