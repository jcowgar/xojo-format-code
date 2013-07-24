' Testing is dramatically different for various configuration options so
' tests are added dynamically based on settings here. These settings must
' match the settings in your Format Code.xojo_script file otherwise failures
' are sure to occur, most probably false.

Dim PadParens As Boolean = False

'
' Represent a test case, good, bad and actual
'

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

'
' Start a new console project and add a new method, this is where
' we will actually do our testing.
'
NewConsoleProject
DoCommand "NewMethod"

Dim results() As String
Dim tests() As TestCase

'
' These tests are preference neutral and can be tested in various
' configurations.
'
tests.Append New TestCase("apples += 10", "apples = apples + 10")
tests.Append New TestCase("apples -= 10", "apples = apples - 10")
tests.Append New TestCase("apples *= 10", "apples = apples * 10")
tests.Append New TestCase("apples /= 10", "apples = apples / 10")

'
' Add tests if paren padding
'
If PadParens Then
tests.Append New TestCase("Add(10,20)", "Add( 10, 20 )")
tests.Append New TestCase("a=((b*Abs(c))/d)", "a = ( ( b * Abs( c ) ) / d )")

Else
'
' Add all other tests. These tests can not be mixed with the above
' because they assume paren padding is OFF and have parens. These
' tests will surely fail if PadParens = True
'
tests.Append New TestCase(_
"dim a as integer=(18*2  )   ", _
"Dim a As Integer = (18 * 2)")

tests.Append New TestCase(_
"for EACH p aS PERSON in people( 5,8 )", _
"For Each p As PERSON In people(5, 8)")
End If

'
' Do our testing
'

Dim i As Integer
For i = 0 To tests.Ubound
Dim tc As TestCase = tests(i)

Text = tc.Bad
RunScript "Format Code.xojo_script"
tc.Actual = Text

If tc.Failed Then
results.Append("'" + EndOfLine + _
"' Test #: " + Str(i) + " FAILED" + EndOfLine + _
"'" + EndOfLine + _
"' Sent:" + EndOfLine + tc.Bad + EndOfLine + EndOfLine + _
"' Expected:" + EndOfLine + tc.Good + EndOfLine + EndOfLine + _
"' Got: " + EndOfLine + tc.Actual + EndOfLine)
End If
Next

'
' Display the results
'

If results.Ubound = -1 Then
Text = "' Everything passed!"
Else
Text = Join(results, EndOfLine)
End If