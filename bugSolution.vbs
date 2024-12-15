Function overloadedFunction(param1)
  If IsMissing(param2) Then
    MsgBox "One parameter version: " & param1
  Else
    MsgBox "Two parameter version: " & param1 & ", " & param2
  End If
End Function

'Incorrect attempt at overloading
'Function overloadedFunction(param1, param2)
'  MsgBox "Two parameter version: " & param1 & ", " & param2
'End Function

'Testing the function
overloadedFunction "single parameter"
overloadedFunction "param1", "param2"