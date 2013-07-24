Class TestCase
Dim Bad As String
Dim Good As String
Dim Actual As String

Sub Constructor(b As String, g As String)
Bad = b
Good = g
End Sub

Function Passed() As Boolean
Return (StrComp(Actual, Good, 0) = 0)
End Function

Function Failed() As Boolean
Return Not Passed()
End Function
End Class

NewConsoleProject
DoCommand "NewMethod"

Dim results() As String
Dim tests() As TestCase

tests.Append New TestCase(_
"dim a as integer=(18*2  )   ", _
"Dim a As Integer = (18 * 2)")

tests.Append New TestCase(_
"for EACH p aS PERSON in people(  )", _
"For Each p As PERSON In people()")

tests.Append New TestCase("apples += 10", "apples = apples + 10")
tests.Append New TestCase("apples -= 10", "apples = apples - 10")
tests.Append New TestCase("apples *= 10", "apples = apples * 10")
tests.Append New TestCase("apples /= 10", "apples = apples / 10")

Dim i As Integer
For i = 0 To tests.Ubound
Dim tc As TestCase = tests(i)

Text = tc.Bad
RunScript "Format Code.xojo_script"
tc.Actual = Text

If tc.Failed Then
results.Append(_
"'   Test #: " + Str(i) + " FAILED" + EndOfLine + _
"'     Sent: [" + tc.Bad + "]" + EndOfLine + _
"' Expected: [" + tc.Good + "]" + EndOfLine + _
"'      Got: [" + tc.Actual + "]" + EndOfLine + _
"'")
End If
Next

If results.Ubound = -1 Then
Text = "' Everything passed!"
Else
Text = Join(results, EndOfLine)
End If
