' Testing is dramatically different for various configuration options so
' tests are added dynamically based on settings here. These settings must
' match the settings in your Format Code.xojo_script file otherwise failures
' are sure to occur, most probably false.

Dim savedClipboard As String = Clipboard

Dim PadParensInner As Boolean = False
Dim PadParensOuter As Boolean = False
Dim PadOperators As Boolean = True
Dim PadComma As Boolean = True

'
' Represent a configuration
'
Class FormatSettings
Dim StringValue As String

Sub Constructor(caseConversion As Integer, padParensInner As Boolean, padParensOuter As Boolean, _
padOperators As Boolean, padComma As Boolean)
Dim results() As String = Array("F", "C", ":", Str(caseConversion))

If padParensInner Then
results.Append "Y"
Else
results.Append "N"
End If

If padParensOuter Then
results.Append "Y"
Else
results.Append "N"
End If

If padOperators Then
results.Append "Y"
Else
results.Append "N"
End If

If padComma Then
results.Append "Y"
Else
results.Append "N"
End If

StringValue = Join(results, "")
End Sub
End Class

'
' Represent a test case, good, bad and actual
'
Class TestCase
Dim Settings As FormatSettings
Dim Bad As String
Dim Good As String
Dim Actual As String

Sub Constructor(s As FormatSettings, b As String, g As String)
Settings = s
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

Dim standardSettings As FormatSettings = New FormatSettings(1, False, False, True, True)
Dim lowerSettings As FormatSettings = New FormatSettings(2, False, False, True, True)
Dim upperSettings As FormatSettings = New FormatSettings(3, False, False, True, True)
Dim padOuterSettings As FormatSettings = New FormatSettings(1, False, True, True, True)
Dim padInnerSettings As FormatSettings = New FormatSettings(1, True, False, True, True)
Dim padAllSettings As FormatSettings = New FormatSettings(1, True, True, True, True)
Dim compressedSettings As FormatSettings = New FormatSettings(1, False, False, False, False)

'
' These tests all rely on PadOperators being True
'
tests.Append New TestCase(standardSettings, "apples += 10", "apples = apples + 10")
tests.Append New TestCase(standardSettings, "apples -= 10", "apples = apples - 10")
tests.Append New TestCase(standardSettings, "apples *= 10", "apples = apples * 10")
tests.Append New TestCase(standardSettings, "apples /= 10", "apples = apples / 10")

' Number parsing
tests.Append New TestCase(standardSettings, "a=10+5", "a = 10 + 5")
tests.Append New TestCase(standardSettings, "a=-10+5", "a = -10 + 5")
tests.Append New TestCase(standardSettings, "a=-10.5+5", "a = -10.5 + 5")

'
' Comment parsing
'
tests.Append New TestCase(standardSettings, "a=10' Assign 10 to A", "a = 10 ' Assign 10 to A")
tests.Append New TestCase(standardSettings, "a=10// Assign 10 to A", "a = 10 // Assign 10 to A")
tests.Append New TestCase(standardSettings, "a=10+5// comment", "a = 10 + 5 // comment")

' Comments should be read until the end of a line
tests.Append New TestCase(standardSettings, "a=10// comment" + EndOfLine + "b=20", _
"a = 10 // comment" + EndOfLine + "b = 20")

' Comments should work on the first line (above cases) and subsequent lines
tests.Append New TestCase(standardSettings, "a=10" + EndOfLine + "// comment?" + EndOfLine + "b=20", _
"a = 10" + EndOfLine + "// comment?" + EndOfLine + "b = 20")

' Comments should have leading space removed
tests.Append New TestCase(standardSettings, "a=10      // Hi", "a = 10 // Hi")

' Keywords shouldn't be parsed in a comment
tests.Append New TestCase(standardSettings, "' if a isa b THEN", "' if a isa b THEN")
tests.Append New TestCase(standardSettings, "// if a isa b THEN", "// if a isa b THEN")

tests.Append New TestCase(padInnerSettings, "Add(10,20)", "Add( 10, 20 )")
tests.Append New TestCase(padInnerSettings, "a=((b*Abs(c))/d)", "a = ( ( b * Abs( c ) ) / d )")

tests.Append New TestCase(padOuterSettings, "Add(10,20)", "Add (10, 20)")
tests.Append New TestCase(padOuterSettings, "a=((b*Abs(c))/d)", "a = ((b * Abs (c)) / d)")

tests.Append New TestCase(padAllSettings, "Add(10,20)", "Add ( 10, 20 )")
tests.Append New TestCase(padAllSettings, "a=((b*Abs(c))/d)", "a = ( ( b * Abs ( c ) ) / d )")

tests.Append New TestCase(standardSettings, _
"dim a as integer=(18*2  )   ", _
"Dim a As Integer = (18 * 2)")

tests.Append New TestCase(standardSettings, _
"for EACH p aS PERSON in people( 5,8 )" + EndOfLine + "Next", _
"For Each p As PERSON In people(5, 8)" + EndOfLine + "Next")

tests.Append New TestCase(compressedSettings, "a = 5  + 12", "a=5+12")
tests.Append New TestCase(compressedSettings, "Add(5, 12)", "Add(5,12)")

' me, self and super should all adhere to case conversion
tests.Append New TestCase(standardSettings, "me.hello()", "Me.hello()")
tests.Append New TestCase(standardSettings, "self.hello()", "Self.hello()")
tests.Append New TestCase(standardSettings, "super.hello()", "Super.hello()")

tests.Append New TestCase(standardSettings, "iF tRUE thEN", "If True Then")
tests.Append New TestCase(lowerSettings, "iF tRUE thEN", "if true then")
tests.Append New TestCase(upperSettings, "iF tRUE thEN", "IF TRUE THEN")

' Strings shouldn't be messed with
tests.Append New TestCase(standardSettings, """if TrUe ThEn""", """if TrUe ThEn""")

'
' Do our testing
'

Dim i As Integer
For i = 0 To tests.Ubound
Dim tc As TestCase = tests(i)

Clipboard = tc.Settings.StringValue

Text = tc.Bad
RunScript "Format Code.xojo_script"
tc.Actual = Text

If tc.Failed Then
results.Append("'------------------------------------------------------------" + EndOfLine + _
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
Dim msg As String

If results.Ubound = -1 Then
Msg = "' " + Str(tests.UBound + 1) + " tests passed, 0 failed"
Else
Dim passedCount As Integer = tests.UBound - results.UBound
Dim failedCount As Integer = tests.UBound + 1 - passedCount

Msg = "' " + Str(passedCount) + " tests passed, " + _
Str(failedCount) + " failed" + EndOfLine + EndOfLine + _
Join(results, "")
End If

Text = msg

Clipboard = savedClipboard
