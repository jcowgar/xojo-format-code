// Alternate tokenizer

//
// Global methods & functions
//

Function ReplaceLineEndings (source As String, lineEnding As String = &uA) As String
source = source.ReplaceAll( &uD + &uA, &uA )
source = source.ReplaceAll( &uD, &uA )
source = source.ReplaceAll( &uA, lineEnding )

return source
End Function

Function IsNumber (v As String) As Boolean

End Function

Function IsPuctuation(v As String) As Boolean

End Function

//
// Classes
//

Class Token
Dim Value As String

Sub Constructor()

End Sub
End Class

dim t as new Token
print t.Value
