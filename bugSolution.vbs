Function MyFunc(param)
  If IsEmpty(param) Or IsNull(param) Then
    Err.Raise vbError, , "Parameter cannot be empty or null" ' Raise an error for better error handling
    'Alternatively, return a specific value to indicate an error condition
    'MyFunc = -1
  Else
    MyFunc = param * 2
  End If
End Function

On Error GoTo ErrHandler
MsgBox MyFunc(10) 'Output 20
MsgBox MyFunc(Empty) 'Raises error
MsgBox MyFunc("") 'Raises error
Exit Sub

ErrHandler:
MsgBox "Error: " & Err.Description
End Sub