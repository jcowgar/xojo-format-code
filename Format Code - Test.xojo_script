'
' Our testing is done by first opening the Format Code Unit Test project.
' This contains a module named "FormatCodePreferences" which allows us to
' change various settings from the test script.
'
' We then, from the test script, create a new module and a new method in
' that module. That method's text area is where the test script will work.
'
' Each call to Test(input, expected) will:
'     1. Set the Text of the editor to `input`
'     2. Call the 'Format Code' Xojo Script
'     3. Compare the Text of the editor to the `expected` value
'
' If the newly formatted Text is equal (case sensitive) to the `expected`
' value then our test has passed, otherwise it has failed.
'
' The Test method keeps track of how many tests have passed, how many have
' failed as well as the input, expected and actual output for those tests
' that failed.
'
' When all the tests are done running, the content of the editor is set to
' an output string containing testing statistics as well as details on any
' failed tests.
'
' To add new tests, simply call:
'     Test("non-formatted-input", "expected-formatted-output")
'
' Be sure to set any specific options you may need/want for your test
' prior to calling Test.
'

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

Sub AlignAs(Assigns value As Boolean)
ConstantValue(preferencesModuleName + ".AlignAs") = If(value, "Yes", "No")
End Sub

Sub AlignEqual(Assigns value As Boolean)
ConstantValue(preferencesModuleName + ".AlignEqual") = If(value, "Yes", "No")
End Sub

Sub DoDebug(Assigns value As Boolean)
ConstantValue(preferencesModuleName + ".DoDebug") = If(value, "Yes", "No")
End Sub

'
' Unit Test
'

Sub Test(bad As String, expected As String)
Const kDebugOutputIdentifier = "' ==== Format Code Script Debug Information ===="

Dim actualWithDebug As String

testIndex = testIndex + 1

Text = bad
RunScript "Format Code Script.xojo_script"
actualWithDebug = Text

' Remove the debug info 
dim actual As String = actualWithDebug
Dim pos As Integer = actual.InStr(kDebugOutputIdentifier)
If pos <> 0 Then
pos = pos - 2 // Get the EOL too
actual = actual.Left(pos)
actual = actual.Trim
End If

If StrComp(expected, actual, 0) <> 0 Then
failCount = failCount + 1

results.Append("'------------------------------------------------------------" + EndOfLine + _
"' Test #: " + Str(testIndex) + " FAILED" + EndOfLine + _
"'" + EndOfLine + _
"' Sent:" + EndOfLine + bad + EndOfLine + EndOfLine + _
"' Expected:" + EndOfLine + expected + EndOfLine + EndOfLine + _
"' Got: " + EndOfLine + actualWithDebug + EndOfLine)

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

DoDebug = True

' =================================================================
' =
' =                         TESTS
' =
' =================================================================

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

Test("RaiseEvent Open (    )", "RaiseEvent Open()")
Test("RaiseEvent Open((18))", "RaiseEvent Open( ( 18 ) )")

Test("while(a)", "While ( a )")
Test("if(true)then b=if(""a""=a,true,false)","If ( True ) Then b = If( ""a"" = a, True, False )")
Test("Add(10,20)", "Add( 10, 20 )")
Test("a=((b*Abs(c))/d)", "a = ( ( b * Abs( c ) ) / d )")

PadParensOuter = True
PadParensInner = False

Test("while(b)", "While (b)")
Test("if(true)then b=if(""b""=a,true,false)","If (True) Then b = If (""b"" = a, True, False)")
Test("Add(10,20)", "Add (10, 20)")
Test("a=((b*Abs(c))/d)", "a = ((b * Abs (c)) / d)")

PadParensInner = True
PadParensOuter = True

Test("while(c)", "While ( c )")
Test("if(true)then b=if(""c""=a,true,false)","If ( True ) Then b = If ( ""c"" = a, True, False )")
Test("Add(10,20)", "Add ( 10, 20 )")
Test("a=((b*Abs(c))/d)", "a = ( ( b * Abs ( c ) ) / d )")

PadParensOuter = False
PadParensInner = False

Test("while(d)", "While (d)")
Test("if(true)then b=if(""d""=a,true,false)","If (True) Then b = If(""d"" = a, True, False)")
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

' Line continuation character with IF and comment
Test("if(true) _'Comment" + EndOfLine + "if(true,true,false)then", "If (True) _ 'Comment" + EndOfLine + "If(True, True, False) Then")
Test("if(true) _ 'Comment" + EndOfLine + "if(true,true,false)then", "If (True) _ 'Comment" + EndOfLine + "If(True, True, False) Then")
Test("if(true) _//Comment" + EndOfLine + "if(true,true,false)then", "If (True) _ //Comment" + EndOfLine + "If(True, True, False) Then")
Test("if(true) _ //Comment" + EndOfLine + "if(true,true,false)then", "If (True) _ //Comment" + EndOfLine + "If(True, True, False) Then")

' Optional keywords with specific conversion types
KeywordsToTitleCase = "SayHello,SayGoodBye,Val"
KeywordsToUpperCase = "ABC,DEF"
KeywordsToLowerCase = "xyz,www"

Test("sayhello(abc)", "SayHello(ABC)")
Test("sayGOODBYE(WWW)", "SayGoodBye(www)")
Test("abc=XYZ", "ABC = xyz")
Test("def=abc", "DEF = ABC")
Test("mail=good", "mail = good")

' Reset things to the standard
PadParensOuter = False
PadParensInner = False
PadComma = True
PadOperators = True

' Optional keywords with specific conversion types
KeywordsToTitleCase = "SayHello,SayGoodBye,Val"
KeywordsToUpperCase = "ABC,DEF"
KeywordsToLowerCase = "xyz,www"

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

' KeepCase tests
KeywordsToTitleCase = ""
KeywordsToUpperCase = ""
KeywordsToLowerCase = ""

CaseConversion = kTitleCase

Test("Dim abcXYZ As Integer = 12" + EndOfLine + _
"Dim HELLO AS Integer = 55" + EndOfLine + _
"abcxyz=hello", _
"Dim abcXYZ As Integer = 12" + EndOfLine + _
"Dim HELLO As Integer = 55" + EndOfLine + _
"abcXYZ = HELLO")

Test("Dim ABc, xYZ As Integer" + EndOfLine + "abc=xyz", _
"Dim ABc, xYZ As Integer" + EndOfLine + "ABc = xYZ")

' Alignment

AlignAs = True
AlignEqual = False
PadComma = True

Test("Dim a As Integer" + EndOfLine + _
"Dim name As String", _
"Dim a    As Integer" + EndOfLine + _
"Dim name As String")

Test("Dim a,b As Integer" + EndOfLine + _
"Dim helloWorldHowAreYou As Integer", _
"Dim a, b                As Integer" + EndOfLine + _
"Dim helloWorldHowAreYou As Integer")

Test("Dim a As Integer" + EndOfLine + _
"Dim abc As Integer" + EndOfLine + _
"Hi()" + EndOfLine + _
"Dim aaaaa As Integer" + EndOfLine + _
"Dim b As Integer", _
"Dim a   As Integer" + EndOfLine + _
"Dim abc As Integer" + EndOfLine + _
"Hi()" + EndOfLine + _
"Dim aaaaa As Integer" + EndOfLine + _
"Dim b     As Integer")

AlignEqual = True

Test("Dim a As Integer = 10" + EndOfLine + _
"Dim xyz As Int32 = 5", _
"Dim a   As Integer = 10" + EndOfLine + _
"Dim xyz As Int32   = 5")

PadComma = False

Test("Dim a,   b As  Integer" + EndOfLine + _
"Dim helloWorld As Integer", _
"Dim a,b        As Integer" + EndOfLine + _
"Dim helloWorld As Integer")

'
' Intentional fail to test debug output in this script.
' Normally this line would be commented out.
'

'Test( "RaiseEvent Open( )", "RaiseEvent Open( )") ' Not really what we'd expect

' =================================================================
' =
' =                    Display the results
' =
' =================================================================

Text = Join(results, EndOfLine) + EndOfLine + EndOfLine + _
"' ---------------------------------------------------------------------" + EndOfLine + _
"' " + Str(passCount) + " test(s) passed, " + Str(failCount) + " test(s) failed"

' Reset our preferences so that each time we run a test and then commit
' to our SCM system, the test project hasn't changed.
CaseConversion = kTitleCase
PadParensInner = False
PadParensOuter = False
PadOperators = True
PadComma = True
AlignAs = False
AlignEqual = False
KeywordsToTitleCase = ""
KeywordsToUpperCase = ""
KeywordsToLowerCase = ""
AdditionalKeywords = ""
DoDebug = False
