Function MyFunction(param1, param2)
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise vbError, , "Parameters cannot be empty. Please provide values for both param1 and param2."
  End If
  ' ... rest of the function ...
End Function