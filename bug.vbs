Function MyFunc(param)
  If IsEmpty(param) Then
    ' Handle empty parameter
    MyFunc = Null 'Should return Null or throw an error
  Else
    ' Process the parameter 
    MyFunc = param * 2
  End If
End Function

MsgBox MyFunc(10) 'Output 20 
MsgBox MyFunc(Empty) 'Output is 0, which is not correct. Should be Null or error
MsgBox MyFunc("") 'Output is 0, which is not correct. Should be Null or error