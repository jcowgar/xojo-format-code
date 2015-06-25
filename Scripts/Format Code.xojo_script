//
// Format Xojo code in the currently opened method
//
// Author: Jeremy Cowgar <jeremy@cowgar.com>
//
// Contributors: 
//   * Kem Tekinay <ktekinay@mactechnologies.com>
// 

Const kVersion = "2015r2"

Const kTitleCase = 1
Const kLowerCase = 2
Const kUpperCase = 3

//
//
// User Preferences:
//
//

// Convert badly formatted "byREF" into ByRef (kTitleCase), byref (kLowerCase) or BYREF (kUpperCase)
Dim CaseConversion As Integer = kTitleCase

// Pad inside parens with a space?
Dim PadParensInner As Boolean = False

// Pad outside parens with space?
Dim PadParensOuter As Boolean = False

// Pad operators with a space?
Dim PadOperators As Boolean = True

// Pad a comma with a trailing space?
Dim PadComma As Boolean = True

// Align 'As' keywords
Dim AlignAs As Boolean = False

// Align '=' assignments
Dim AlignEqual As Boolean = False

// Appends debug information to the end of the editor. This should be
// set to true only for those working on Code Formatter.
Dim DoDebug As Boolean = False

Dim KeywordsToCapitalize() As String

// Keywords that you want Code Formatter to correct the case on. By default, all
// keywords and pragmas are listed
KeywordsToCapitalize = Array("AddHandler", "AddressOf", "And", "Array", "As", "Assigns", _
"Break", "ByRef", "ByVal", "CType", "Call", "Case", "Catch", "Const", "Continue", _
"Declare", "Dim", "Do", "Loop", "DownTo", "Each", "Else", "ElseIf", "End", "Enum", "Exception", _
"Exit", "Extends", "False", "Finally", "For", "Not", "Next", "Function", "GOTO", "GetTypeInfo", _
"If", "Then", "In", "Is", "IsA", "Lib", "Loop", "Me", "New", "Nil", "Optional", "ParamArray", _
"Raise", "RaiseEvent", "Redim", "Rem", "RemoveHandler", "Return", "Select", "Self", _
"Soft", "Static", "Step", "Structure", "Sub", "Super", "Then", "To", "True", "Try", "Until", "Using", _
"Wend", "While", "#If", "#ElseIf", "#EndIf", "#Pragma", "DebugBuild", "RBVersion", _
"RBVersionString", "Target32Bit", "Target64Bit", "TargetBigEndian", "TargetCarbon", _
"TargetCocoa", "TargetHasGUI", "TargetLinux", "TargetLittleEndian", _
"TargetMacOS", "TargetMachO", "TargetWeb", "TargetWin32", "TargetX86", _
"BackgroundTasks", "BoundsChecking", "BreakOnExceptions", "DisableAutoWaitCursor", _
"DisableBackgroundTasks", "DisableBoundsChecking", "Error", "NilObjectChecking", _
"StackOverflowChecking", "Unused", "Warning", "X86CallingConvention", "Boolean", _
"CFStringRef", "CString", "Currency", "Delegate", "Double", "Int16", "Int32", "Int64", _
"Int8", "Integer", "OSType", "PString", "Ptr", "Short", "Single", "String", _
"Structure", "Text", "UInt16", "UInt32", "UInt64", "UInt8", "UShort", "WindowPtr", _
"WString", "XMLNodeType")

// Set up additional arrays of keywords that should be treated in a special way.
// Keywords that do not appear in any of these arrays will be treated with the default behavior.
// Values in any of these arrays will override the default.
//
// For example, if you want all of your keywords title cased except for if/then/else,
// you'd set CaseConversion to kTitleCase and add "if", "then", and "else" to 
// KeywordsToLowercase. You could also add something like MyClass to KeywordsToTitleCase
// and it will be replaced too.
//
// You can add these to your FormatCodePreferences module.

Dim KeywordsToTitleCase() As String 
Dim KeywordsToUpperCase() As String
Dim KeywordsToLowerCase() As String

// Fill in your preferences here
KeywordsToTitleCase = Array("")

KeywordsToUpperCase = Array("")

KeywordsToLowerCase = Array("")

//
// 
// END OF USER PREFERENCES
//
//

Dim XojoKeywords() As String = Array("As", "Assigns", "Break", "ByRef", "ByVal", "Call", "Case", _
"Catch", "Const", "Continue", "Declare", "Dim", "Do", "Loop", "DownTo", "Each", "Else", "End", _
"Enum", "Exception", "Exit", "Extends", "False", "Finally", "For", "Not", "Next", "Function", _
"GOTO", "If", "Then", "In", "Is", "IsA", "Lib", "Loop", "New", "Nil", "Optional", "ParamArray", _
"Raise", "RaiseEvent", "Redim", "Rem", "Return", "Select", "Case", "Soft", "Static", "Step", _
"Structure", "Sub", "Super", "Then", "To", "True", "Try", "Until", "Using", "Wend", "While", _
"#If", "#ElseIf", "#EndIf", "#Pragma", "DebugBuild", "RBVersion", "RBVersionString", _
"Target32Bit", "Target64Bit", "TargetBigEndian", "TargetCarbon", _
"TargetCocoa", "TargetHasGUI", "TargetLinux", "TargetLittleEndian", _
"TargetMacOS", "TargetMachO", "TargetWeb", "TargetWin32", "TargetX86", _
"BackgroundTasks", "BoundsChecking", "BreakOnExceptions", "DisableAutoWaitCursor", _
"DisableBackgroundTasks", "DisableBoundsChecking", "Error", "NilObjectChecking", _
"StackOverflowChecking", "Unused", "Warning", "X86CallingConvention")

// Some keywords can double as functions.
// The rule is that a keyword in this array that is NOT the first token will be treated as a function.
Dim XojoKeyWordsAsFunction() As String = Array("If")

Dim KeepCaseIdentifiers() As String

//
// Try to read preferences from a `FormatCode` module in the application. If a success,
// then those properties override the above settings.
//

Const preferencesModuleName = "FormatCodePreferences"

Function BooleanConstantValue(key As String, defaultValue As Boolean) As Boolean
Select Case ConstantValue(preferencesModuleName + "." + key)
Case "1", "Yes", "True"
Return True

Case "0", "No", "False"
Return False

Case Else
Return defaultValue
End Select
End Function

Function ArrayConstantValue(key As String, defaultValue() As String) As String()
Dim value As String = ConstantValue(preferencesModuleName + "." + key).Trim

If value = "" Then
Return defaultValue

Else
value = value.ReplaceAll(EndOfLine, ",")

Dim arr() As String = Split(value, ",")

// Make sure each value is free of surrounding spaces
For i As Integer = 0 To arr.Ubound
arr(i) = arr(i).Trim
Next i

// Remove any blank items
For i As Integer = arr.Ubound DownTo 0
If arr(i) = "" Then
arr.Remove i
End If
Next i

Return arr()
End If
End Function

Function MergeArrays(arr1() As String, arr2() As String) As String()
Dim result() As String
Dim resultIndex As Integer = -1
ReDim result(arr1.Ubound + arr2.Ubound + 1)

For i As Integer = 0 To arr1.Ubound
resultIndex = resultIndex + 1
result(resultIndex) = arr1(i)
Next i

For i As Integer = 0 To arr2.Ubound
resultIndex = resultIndex + 1
result(resultIndex) = arr2(i)
Next i

Return result()
End Function

Function ReplaceLineEndings(sourceString As String) As String
sourceString = sourceString.ReplaceAll(Chr(13) + Chr(10), Chr(10))
sourceString = sourceString.ReplaceAll(Chr(13), Chr(10))

return sourceString
End Function

Select Case ConstantValue(preferencesModuleName + ".CaseConversion")
Case "kTitleCase", Str(kTitleCase)
CaseConversion = kTitleCase

Case "kLowerCase", Str(kLowerCase)
CaseConversion = kLowerCase

Case "kUpperCase", Str(kUpperCase)
CaseConversion = kUpperCase
End Select

PadParensInner = BooleanConstantValue("PadParensInner", PadParensInner)
PadParensOuter = BooleanConstantValue("PadParensOuter", PadParensOuter)
PadOperators = BooleanConstantValue("PadOperators", PadOperators)
PadComma = BooleanConstantValue("PadComma", PadComma)
AlignAs = BooleanConstantValue("AlignAs", AlignAs)
AlignEqual = BooleanConstantValue("AlignEqual", AlignEqual)
KeywordsToTitleCase = ArrayConstantValue("KeywordsToTitleCase", KeywordsToTitleCase)
KeywordsToUpperCase = ArrayConstantValue("KeywordsToUpperCase", KeywordsToUpperCase)
KeywordsToLowerCase = ArrayConstantValue("KeywordsToLowerCase", KeywordsToLowerCase)
DoDebug = BooleanConstantValue("DoDebug", DoDebug)

// Grab any additional, user-defined keywords from the module. These will be added to the list above.
Dim AdditionalKeywords() As String
AdditionalKeywords = ArrayConstantValue("AdditionalKeywords", AdditionalKeywords)
KeywordsToCapitalize = MergeArrays(KeywordsToCapitalize, AdditionalKeywords)
Redim AdditionalKeywords(-1) ' We don't need this anymore

//
// Code Formatting Code
//

Dim SpecialCharacters() As String = Array("<", ">", "<>", ">=", "<=", "=", "+", "-", "*", "/", _
"^", "(", ")", ",", ":", ".")

//
// Helper functions
//

Function IsASpecial(value As String) As Boolean
Return (SpecialCharacters.IndexOf(value) > -1)
End Function

Function IsANumber(value As String) As Boolean
Dim hasDecimal As Boolean = False

Dim chars() As String = Split(value, "")
For i As Integer = 0 To chars.Ubound
Dim chCode As Integer
chCode = Asc(chars(i))

If i = 0 And (chCode = 43 Or chCode = 45) Then
// Good

ElseIf chCode >= 48 And chCode <= 57 Then
// Good

ElseIf chCode = 46 And hasDecimal = False Then
// Good
hasDecimal = True

Else
Return False
End If
Next

Return True
End Function

Function IsAString(value As String) As Boolean
Return (value.Left(1) = """" And value.Right(1) = """")
End Function

Function IsAComment(value As String) As Boolean
Return (value.Left(1) = "'" Or value.Left(2) = "//")
End Function

Function FirstLetterCap(value As String) As String
Return value.Left(1).Uppercase + value.Mid(2)
End Function

Function CaseConvert(value As String, thisCaseConversion As Integer) As String
Select Case thisCaseConversion
Case kTitleCase ' TitleCase
Return FirstLetterCap(value) ' Not using native TitleCase since that will lowercase the remaining letters

Case kLowerCase ' lower case
Return Lowercase(value)

Case kUpperCase ' UPPER CASE
Return Uppercase(value)

Else
Print "Something went wrong with capitalizing!"

Return value
End Select
End Function

//
// Represent a single token
//
Class Token
Const Keyword = 1
Const Identifier = 2
Const Number = 3
Const Special = 4
Const NewLine = 5
Const StringLiteral = 6
Const Comment = 7

Dim Value As String
Dim Type As Integer
Dim StartIndex As Integer
Dim Length As Integer
Dim PadLengthBefore As Integer
Dim PadLengthAfter As Integer

Sub Constructor(v As String, keepCase As Boolean, startOfLine As Boolean)
Dim capitalizeIndex As Integer = -1
Dim thisCaseConversion As Integer
Dim useArray() As String

If Not keepCase Then
thisCaseConversion = kUpperCase
useArray = KeywordsToUpperCase
capitalizeIndex = KeywordsToUpperCase.IndexOf(v)

If capitalizeIndex = -1 Then
thisCaseConversion = kLowerCase
useArray = KeywordsToLowerCase
capitalizeIndex = KeywordsToLowerCase.IndexOf(v)
End If

If capitalizeIndex = -1 Then
thisCaseConversion = kTitleCase
useArray = KeywordsToTitleCase
capitalizeIndex = KeywordsToTitleCase.IndexOf(v)
End If

If capitalizeIndex = -1 Then
thisCaseConversion = CaseConversion
useArray = KeywordsToCapitalize
capitalizeIndex = KeywordsToCapitalize.IndexOf(v)
End If
End If

If capitalizeIndex <> -1 Then
Value = CaseConvert(useArray(capitalizeIndex), thisCaseConversion)

If XojoKeywords.IndexOf(Value) <> -1 Then
If startOfLine Or XojoKeywordsAsFunction .IndexOf(Value) = -1 Then
Type = Keyword
Else
Type = Identifier
End If

Else
Type = Identifier
End If

Else
Value = v

If Value = Chr(10) Then
Type = NewLine

ElseIf IsASpecial(Value) Then
Type = Special

ElseIf IsANumber(Value) Then
Type = Number

ElseIf IsAComment(Value) Then
Type = Comment

ElseIf IsAString(Value) Then
Type = StringLiteral

Else
Type = Identifier
End If
End If
End Sub
End Class

//
// Convert a string into a stream of tokens.
//
Class Tokenizer
Dim Tokens() As Token
Dim Code As String
Dim CodeLength As Integer

// Parsing state variables
Private mCurrentPosition As Integer
Private mTokenStartPosition As Integer
Private mInString As Boolean
Private mIsDimming As Boolean
Private mStartOfLine As Boolean = True

Sub MaybeAddToken(incrementPosition As Boolean = True)
If mCurrentPosition <= mTokenStartPosition Then
Return
End If

Dim code As String = Trim(Code.Mid(mTokenStartPosition, mCurrentPosition - mTokenStartPosition))

If code <> "" Then
Dim keepCase As Boolean

If mIsDimming Then
If code = "As" Then
mIsDimming = False
Else
KeepCaseIdentifiers.Append code
End If

Else
Dim keepCaseIndex As Integer = KeepCaseIdentifiers.IndexOf(code)
If keepCaseIndex > -1 Then
code = KeepCaseIdentifiers(keepCaseIndex)
keepCase = True
End If
End If

Dim tok As Token = New Token(code, keepCase, mStartOfLine)

tok.StartIndex = mTokenStartPosition
tok.Length = mCurrentPosition - mTokenStartPosition

Tokens.Append(tok)

If tok.Type = Token.Keyword And tok.Value = "Dim" Then
mIsDimming = True
End If

mStartOfLine = False
End If

mTokenStartPosition = mCurrentPosition

If incrementPosition Then
mTokenStartPosition = mTokenStartPosition + 1
End If
End Sub

Sub AddToken(value As String)
Dim tok As New Token(value, False, mStartOfLine)

tok.StartIndex = mTokenStartPosition
tok.Length = value.Len

Tokens.Append(tok)

mTokenStartPosition = mCurrentPosition + 1
mStartOfLine = False
End Sub

Sub AddCommentToken()
Dim eol As Integer = InStr(mCurrentPosition, Code, Chr(10))

If eol = 0 Then
eol = CodeLength + 1
End If

mCurrentPosition = eol

MaybeAddToken
End Sub

Function Tokenize(sourceCode As String) As Boolean
Code = sourceCode
CodeLength = sourceCode.Len

Redim Tokens(-1)
mCurrentPosition = 1
mTokenStartPosition = 1
mInString = False
mStartOfLine = True

Dim encounteredUnderscore As Boolean

While mCurrentPosition <= CodeLength
Dim ch As String = Code.Mid(mCurrentPosition, 1)
Dim nextCh, nextNextCh As String

If mCurrentPosition < CodeLength Then
nextCh = Code.Mid(mCurrentPosition + 1, 1)
If (mCurrentPosition + 1) < CodeLength Then
nextNextCh = Code.Mid(mCurrentPosition + 2, 1) ' Needed to determine `//` type comments
End If
End If

If mInString Then
If ch = """" Then
If nextCh = """" Then
// Increment past the next quote, it is a double quote
mCurrentPosition = mCurrentPosition + 1

Else
mInString = False

AddToken(Trim(Code.Mid(mTokenStartPosition, mCurrentPosition - mTokenStartPosition + 1)))
End If
End If

Else
Select Case ch
Case """"
mInString = True

Case " "
MaybeAddToken

Case "-"
// Could be a plus symbol or a negative number. This detection logic isn't the best...
If Asc(nextCh) < 48 Or Asc(nextCh) > 57 Then
// Next character is not a number, add this token as a math operation of its own
MaybeAddToken

AddToken(ch)
End If

Case "="
// Look at the previous tokens, if it was a +, -, * or /, do magic
// to come up with the expanded version
Dim addToken As Boolean = True

If Tokens.UBound >= 1 Then
Dim oneTokenBack As Token = Tokens(Tokens.UBound)
Dim twoTokensBack As Token = Tokens(Tokens.UBound - 1)
Dim expansionOps As String = "+-*/"

If expansionOps.InStr(oneTokenBack.Value) > 0 Then
addToken = False

// Delete the previous expansionOp
Tokens.Remove(Tokens.UBound)

AddToken("=")
AddToken(twoTokensBack.Value)
AddToken(oneTokenBack.Value)
End If
End If

If addToken Then
MaybeAddToken

AddToken(ch)
End If

Case "(", ")", ",", "+", "*", "^", ":"
MaybeAddToken

AddToken(ch)

Case Chr(10)
MaybeAddToken

AddToken(ch)

If encounteredUnderscore Then
encounteredUnderscore = False ' Reset this
Else
mStartOfLine = True
End If

Case "_"
// Only process the _ character as an individual token if it is the start of a token
// and it is followed by a space, EndOfLine or comment character
If mCurrentPosition = mTokenStartPosition And _
(nextCh = Chr(10) Or nextCh = " " Or nextCh = "'" Or _
(nextCh = "/" And nextNextCh = "/")) Then
encounteredUnderscore = True
MaybeAddToken

AddToken(ch)
End If

Case "'"
MaybeAddToken(False)
AddCommentToken

Case "/"
// We could have // which indicates a comment and should be a single token, not
// two forward slash tokens.
If nextCh = "/" Then
MaybeAddToken(False)

AddCommentToken

Else
// Must have been division
MaybeAddToken

AddToken("/")
End If

Case "<", ">" 
MaybeAddToken

// We could have <>, >=, <=
If nextCh = ">" Or nextCh = "<" Or nextCh = "=" Then
mCurrentPosition = mCurrentPosition + 1

AddToken(ch + nextCh)

Else
AddToken(ch)
End if

Case "."
MaybeAddToken

AddToken(ch)
End Select
End If

mCurrentPosition = mCurrentPosition + 1
Wend

MaybeAddToken

Return True
End Function
End Class

//
// StringWriter - Take a stream of tokens and write them to a string
//
Class StringWriter
Dim Tokens() As Token
Dim DebugString As String
Dim mRow As Integer
Dim mColumn As Integer
Private mResult As String

Private Sub AddSpace()
mColumn = mColumn + 1
mResult = mResult + " "
End Sub

Private Sub AddString(value As String)
mColumn = mColumn + value.Len
mResult = mResult + value
End Sub

Private Sub AddEndOfLine()
mResult = mResult + Chr(10)
mColumn = 0
mRow = mRow + 1
End Sub

Private Function AlignAsFindBlock(start As Integer, ByRef blockBegin As Integer, ByRef blockEnd As Integer) As Boolean
If start >= Tokens.UBound Then
Return False
End If

Dim foundBlock As Boolean
Dim nextBetterBeDim As Boolean

blockBegin = -1
blockEnd = -1

For i As Integer = start To Tokens.UBound
Dim tok As Token = Tokens(i)

If blockBegin = -1 And tok.Type = Token.Keyword And tok.Value = "Dim" Then
blockBegin = i
blockEnd = i

ElseIf blockBegin > -1 And tok.Type = Token.NewLine Then
nextBetterBeDim = True

ElseIf nextBetterBeDim And tok.Type = Token.Keyword And tok.Value = "Dim" Then
nextBetterBeDim = False
blockEnd = i

ElseIf nextBetterBeDim Then
Exit
End If
Next

If blockBegin > -1 And blockBegin <> blockEnd Then
For i As Integer = blockEnd To Tokens.UBound
blockEnd = i

If Tokens(i).Type = Token.NewLine Then
Exit
End If
Next

Return True
End If

Return False
End Function

Private Sub AlignAsStatements()
If AlignAs = False And AlignEqual = False Then
Return
End If

Dim start As Integer
Dim blockBegin As Integer
Dim blockEnd As Integer

While AlignAsFindBlock(start, blockBegin, blockEnd)
Dim maxVariableLength As Integer
Dim thisVariableLength As Integer
Dim maxTypeLength As Integer
Dim measureNext As Boolean

For i As Integer = blockBegin To blockEnd
Dim tok As Token = Tokens(i)

If tok.Type = Token.Keyword And tok.Value = "Dim" Then
measureNext = True

thisVariableLength = 0

ElseIf tok.Type = Token.Keyword And tok.Value = "As" Then
measureNext = False

If thisVariableLength > maxVariableLength Then
maxVariableLength = thisVariableLength
End If

tok.PadLengthAfter = thisVariableLength

ElseIf tok.Type = Token.Special And tok.Value = "=" Then
Dim last As Token = Tokens(i - 1)
Dim thisTypeLength As Integer = last.Value.Len

If thisTypeLength > maxTypeLength Then
maxTypeLength = thisTypeLength
End If

ElseIf measureNext Then
thisVariableLength = thisVariableLength + tok.Value.Len

If tok.Value = "," And PadComma Then
thisVariableLength = thisVariableLength + 1
End If
End If
Next

For i As Integer = blockBegin To blockEnd
Dim tok As Token = Tokens(i)

If tok.Type = Token.Keyword And tok.Value = "As" Then
Dim lastTok As Token = Tokens(i - 1)

lastTok.PadLengthAfter = maxVariableLength - tok.PadLengthAfter

tok.PadLengthAfter = 0

ElseIf tok.Type = Token.Special And tok.Value = "=" And AlignEqual Then
Dim lastTok As Token = Tokens(i - 1)

lastTok.PadLengthAfter = maxTypeLength - lastTok.Value.Len
End If
Next

start = blockEnd + 1
Wend
End Sub

Function Format(theTokens() As Token) As String
Dim i As Integer
Dim tok, lastTok, nextTok As Token = Nil

mRow = 1
mColumn = 0
mResult = ""

Tokens = theTokens

AlignAsStatements

For i = 0 To Tokens.UBound
tok = Tokens(i)
If i < Tokens.UBound Then
nextTok = Tokens(i + 1)

Else
nextTok = Nil
End If

If DoDebug Then
DebugString = DebugString + "' Token: '" + tok.Value.Trim + "', Type: " + Str(tok.Type) + ", " + _
"Start: " + Str(tok.StartIndex) + ", Length: " + Str(tok.Length) + Chr(10)
End If

Select Case tok.Value
Case Chr(10)
AddEndOfLine

Else
For j As Integer = 1 To tok.PadLengthBefore
AddString(" ")
Next

AddString(tok.Value)

For j As Integer = 1 To tok.PadLengthAfter
AddString(" ")
Next
End Select

// Add a space between tokens, if necessary
If mColumn > 0 Then
If nextTok <> Nil Then
If tok.Type = Token.Comment Then
AddEndOfLine

ElseIf tok.Value = "," Then
If PadComma Then
AddSpace
End If

ElseIf tok.Value = ")" And nextTok.Value.Left(1) = "." Then
// Do nothing

ElseIf tok.Type = Token.Special And tok.Value = "." Then
// Do nothing

ElseIf tok.Type = Token.Special And nextTok.Value = "(" Then
If Not PadOperators Then
// Do Nothing
ElseIf tok.Value <> "(" Or PadParensInner Then
AddSpace
End If

ElseIf nextTok.Type = Token.Newline Then
// Do nothing

ElseIf nextTok.Value = "," Then
// Do nothing

ElseIf nextTok.Value = "." Then
// Do nothing

ElseIf nextTok.Value = "(" Then
If tok.Type = Token.Keyword Or PadParensOuter Then
AddSpace
Else
// Do nothing
End If

ElseIf nextTok.Value = ")" Then
If PadParensInner And tok.Value <> "(" Then
AddSpace
End If

ElseIf tok.Type = Token.Special And Not PadOperators Then
// Do nothing

ElseIf nextTok.Type = Token.Special And Not PadOperators Then
// Do nothing

ElseIf tok.Value <> "(" Or PadParensInner Then
AddSpace

Else
// Do nothing
End If
End If
End If

lastTok = tok
Next

Return Trim(mResult)
End Function
End Class

//
// Magic communication between the Format Code script and the Unit Testing
//

Const kVersionMagic = "$FormatCodeVersion$"

//
// Actual program to interact with Xojo IDE to tokenize and format code
//

Sub Main()
//
// Check for any "Magic" communication with the outside world, mainly the
// unit testing script
//

select case Text
case kVersionMagic
Text = kVersion
return

end select

//
// Check for any unsupported settings
//

If AlignEqual And Not AlignAs Then
Print "AlignEqual does not yet work without AlignAs also, please update your settings"

Return
End If

//
// Start the format process!
//

Dim code As String

// If the editor as text selected, assume the user wants to format only the
// selected text. Otherwise, format the entire editor content

If SelLength > 0 Then
code = SelText

Else
code = Text
End If

code = ReplaceLineEndings(code)

Dim tokenize As New Tokenizer
Dim writer As New StringWriter

If tokenize.Tokenize(code) = False Then
Call ShowDialog("Error", "Could not convert the code into a valid stream of tokens", "OK")

Return
End If

Dim result As String = writer.Format(tokenize.Tokens)

If DoDebug Then
result = result + EndOfLine + EndOfLine + "' ==== Format Code Script Debug Information ====" 
result = result + EndOfLine + "' Version: " + kVersion
result = result + EndOfLine + "' (this output is here because DoDebug is set to True within the script)"
result = result + EndOfLine + EndOfLine + writer.DebugString
End If

//
// Make sure something has changed before updating the text editor
//
If StrComp(result.Trim, code.Trim, 0) <> 0 Then
// Save the cursor position (simple, doesn't always restore the position contextually)
// as formatting could have changed enough to make your old cursor position no longer
// the same as the new with the same index.
Dim cursorPosition As Integer = SelStart

If SelLength > 0 Then
SelText = result

Else
Text = result
End If

// Restore the cursor position
SelStart = cursorPosition

End If
End Sub

Main()
