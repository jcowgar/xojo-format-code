' Testing is dramatically different for various configuration options so
' tests are added dynamically based on settings here. These settings must
' match the settings in your Format Code.xojo_script file otherwise failures
' are sure to occur, most probably false.

Const preferencesModuleName = "FormatCodePreferences"
Const kTitleCase = 1
Const kLowerCase = 2
Const kUpperCase = 3

Dim testIndex, passCount, failCount As Integer
Dim results() As String

'
' Quick methods to change preferences
'

Sub CaseConversion(Assigns value As Integer)
ConstantValue(preferencesModuleName + ".CaseConversion") = Str(value)
End Sub

Sub PadParensInner(Assigns value As Boolean)
ConstantValue(preferencesModuleName + ".PadParensInner") = If(value, "Yes", "No")
End Sub

Sub PadParensOuter(Assigns value As Boolean)
ConstantValue(preferencesModuleName + ".PadParensOuter") = If(value, "Yes", "No")
End Sub

Sub PadOperators(Assigns value As Boolean)
ConstantValue(preferencesModuleName + ".PadOperators") = If(value, "Yes", "No")
End Sub

Sub PadComma(Assigns value As Boolean)
ConstantValue(preferencesModuleName + ".PadComma") = If(value, "Yes", "No")
End Sub

Sub KeywordsToTitleCase(Assigns value As String)
ConstantValue(preferencesModuleName + ".KeywordsToTitleCase") = value
End Sub

Sub KeywordsToUpperCase(Assigns value As String)
ConstantValue(preferencesModuleName + ".KeywordsToUpperCase") = value
End Sub

Sub KeywordsToLowerCase(Assigns value As String)
ConstantValue(preferencesModuleName + ".KeywordsToLowerCase") = value
End Sub

Sub AdditionalKeywords(Assigns value As String)
ConstantValue(preferencesModuleName + ".AdditionalKeywords") = value
End Sub

'
' Unit Test
'

Sub Test(bad As String, expected As String)
Dim actual As String

testIndex = testIndex + 1

Text = bad
RunScript "Format Code.xojo_script"
actual = Text

If StrComp(expected, actual, 0) <> 0 Then
failCount = failCount + 1

results.Append("'------------------------------------------------------------" + EndOfLine + _
"' Test #: " + Str(testIndex) + " FAILED" + EndOfLine + _
"'" + EndOfLine + _
"' Sent:" + EndOfLine + bad + EndOfLine + EndOfLine + _
"' Expected:" + EndOfLine + expected + EndOfLine + EndOfLine + _
"' Got: " + EndOfLine + actual + EndOfLine)

Else
passCount = passCount + 1
End If
End Sub

'
' Start a new console project and add a new method, this is where
' we will actually do our testing.
'

DoCommand "NewModule"
DoCommand "NewMethod"

Test("apples += 10", "apples = apples + 10")
Test("apples -= 10", "apples = apples - 10")
Test("apples *= 10", "apples = apples * 10")
Test("apples /= 10", "apples = apples / 10")

' Number parsing
Test("a=10+5", "a = 10 + 5")
Test("a=-10+5", "a = -10 + 5")
Test("a=-10.5+5", "a = -10.5 + 5")

'
' Comment parsing
'
Test("a=10' Assign 10 to A", "a = 10 ' Assign 10 to A")
Test("a=10// Assign 10 to A", "a = 10 // Assign 10 to A")
Test("a=10+5// comment", "a = 10 + 5 // comment")

' Comments should be read until the end of a line
Test("a=10// comment" + EndOfLine + "b=20", _
"a = 10 // comment" + EndOfLine + "b = 20")

' Comments should work on the first line (above cases) and subsequent lines
Test("a=10" + EndOfLine + "// comment?" + EndOfLine + "b=20", _
"a = 10" + EndOfLine + "// comment?" + EndOfLine + "b = 20")

' Comments should have leading space removed
Test("a=10      // Hi", "a = 10 // Hi")

' Keywords shouldn't be parsed in a comment
Test("' if a isa b THEN", "' if a isa b THEN")
Test("// if a isa b THEN", "// if a isa b THEN")

PadParensInner = True
PadParensOuter = False
PadComma = True
PadOperators = True

Test("Add(10,20)", "Add( 10, 20 )")
Test("a=((b*Abs(c))/d)", "a = ( ( b * Abs( c ) ) / d )")

PadParensOuter = True
PadParensInner = False

Test("Add(10,20)", "Add (10, 20)")
Test("a=((b*Abs(c))/d)", "a = ((b * Abs (c)) / d)")

PadParensInner = True
PadParensOuter = True

Test("Add(10,20)", "Add ( 10, 20 )")
Test("a=((b*Abs(c))/d)", "a = ( ( b * Abs ( c ) ) / d )")

PadParensOuter = False
PadParensInner = False

Test(_
"dim a as integer=(18*2  )   ", _
"Dim a As Integer = (18 * 2)")

Test(_
"for EACH p aS PERSON in people( 5,8 )" + EndOfLine + "Next", _
"For Each p As PERSON In people(5, 8)" + EndOfLine + "Next")

PadOperators = False
PadParensInner = False
PadParensOuter = False
PadComma = False

Test("a = 5  + 12", "a=5+12")
Test("Add(5, 12)", "Add(5,12)")

' me, self and super should all adhere to case conversion
Test("me.hello()", "Me.hello()")
Test("self.hello()", "Self.hello()")
Test("super.hello()", "Super.hello()")

CaseConversion = kTitleCase

Test("iF tRUE thEN", "If True Then")

CaseConversion = kLowerCase

Test("iF tRUE thEN", "if true then")

CaseConversion = kUpperCase

Test("iF tRUE thEN", "IF TRUE THEN")

CaseConversion = kTitleCase

' Strings shouldn't be messed with
Test("""if TrUe ThEn""", """if TrUe ThEn""")

' Test to make sure combination operators are properly output as one unit
PadOperators = True
Test("8>=20 And 8<=99 And 8<>9", "8 >= 20 And 8 <= 99 And 8 <> 9")

' Padding around the pair operator
Test("1:2", "1 : 2")

' Continuation character properly formatted
Test("a=10+ _", "a = 10 + _")
Test("a=10+ _' Hi", "a = 10 + _ ' Hi")

' Reset things to the standard
PadParensOuter = False
PadParensInner = False
PadComma = True
PadOperators = True

' Optional keywords with specific conversion types
KeywordsToTitleCase = "SayHello,SayGoodBye,Val"
KeywordsToUpperCase = "ABC,DEF"
KeywordsToLowerCase = "xyz,www"

Test("sayhello(abc)", "SayHello(ABC)")
Test("sayGOODBYE(WWW)", "SayGoodBye(www)")
Test("abc=XYZ", "ABC = xyz")
Test("def=abc", "DEF = ABC")
Test("mail=good", "mail = good")
Test("ABC=val(""8374 abc XYZ saYHEllo"")", "ABC = Val(""8374 abc XYZ saYHEllo"")")

' Additional Keywords with default case conversion
AdditionalKeywords = "GetName,GetAddress"

CaseConversion = kTitleCase
Test("getname(10)", "GetName(10)")
Test("GETADDRESS(30)", "GetAddress(30)")

CaseConversion = kLowerCase
Test("getNAme(10)", "getname(10)")
Test("GETADDRESS(30)", "getaddress(30)")

CaseConversion = kUpperCase
Test("GetName(10)", "GETNAME(10)")
Test("GETADDRESS(30)", "GETADDRESS(30)")

'
' Display the results
'

Text = Join(results, EndOfLine) + EndOfLine + EndOfLine + _
"' ---------------------------------------------------------------------" + EndOfLine + _
"' " + Str(passCount) + " test(s) passed, " + Str(failCount) + " test(s) failed"

CaseConversion = kTitleCase
PadParensInner = False
PadParensOuter = False
PadOperators = True
PadComma = True
KeywordsToTitleCase = ""
KeywordsToUpperCase = ""
KeywordsToLowerCase = ""
