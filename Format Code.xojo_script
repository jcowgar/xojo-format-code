'
' Format Xojo code in the currently opened method
'
' Version: 0.1
' Author: Jeremy Cowgar <jeremy@cowgar.com>
' Contributors: 
' 

Dim DoDebug As Boolean
DoDebug = False

Dim KeywordsToCapitalize() As String
KeywordsToCapitalize = Array("AddHandler", "AddressOf", "Array", "As", "Assigns", _
"Break", "ByRef", "ByVal", "CType", "Call", "Case", "Catch", "Const", "Continue", _
"Declare", "Dim", "Do", "Loop", "DownTo", "Each", "Else", "End", "Enum", "Exception", _
"Exit", "Extends", "False", "Finally", "For", "Next", "Function", "GOTO", "GetTypeInfo", _
"If", "Then", "In", "Is", "IsA", "Lib", "Loop", "Next", "Nil", "Optional", "ParamArray", "Raise", _
"RaiseEvent", "Redim", "Rem", "RemoveHandler", "Return", "Select", "Case", "Soft", _
"Static", "Step", "Structure", "Sub", "Super", "Then", "To", "True", "Try", "Until", _
"Wend", "While", "#If", "#ElseIf", "#EndIf", "#Pragma", "DebugBuild", "RBVersion", _
"RBVersionString", "Target32Bit", "Target64Bit", "TargetBigEndian", "TargetCarbon", _
"TargetCocoa", "TargetHasGUI", "TargetLinux", "TargetLittleEndian", _
"TargetMacOS", "TargetMachO", "TargetWeb", "TargetWin32", "TargetX86", _
"BackgroundTasks", "BoundsChecking", "BreakOnExceptions", "DisableAutoWaitCursor", _
"DisableBackgroundTasks", "DisableBoundsChecking", "Error", "NilObjectChecking", _
"StackOverflowChecking", "Unused", "Warning", "X86CallingConvention")

Class Token
Dim value As String

Sub Constructor(v As String)
Dim capitalizeIndex As Integer
capitalizeIndex = KeywordsToCapitalize.IndexOf(v)

If capitalizeIndex > -1 Then
value = KeywordsToCapitalize(capitalizeIndex)
Else
value = v
End If
End Sub
End Class

Sub MaybeAddToken()
If currentPosition <= tokenStartPosition Then
Return
End If

Dim tok As Token

tok = New Token(Trim(code.Mid(tokenStartPosition, currentPosition - tokenStartPosition)))

If tok.value.len > 0 Then
tokens.Append(tok)
End If

tokenStartPosition = currentPosition + 1
End Sub

Sub AddToken(value As String)
tokens.Append(New Token(value))

tokenStartPosition = currentPosition + 1
End Sub

Dim tokens() As Token
Dim currentPosition As Integer = 0
Dim tokenStartPosition As Integer = 0
Dim code As String = Text
Dim codeLength As Integer = code.Len
Dim inString As Boolean = False

While currentPosition <= codeLength
Dim ch As String = code.Mid(currentPosition, 1)
Dim nextCh As String

If currentPosition < codeLength Then
nextCh = code.Mid(currentPosition + 1, 1)
End If

If inString Then
If ch = """" Then
If currentPosition < codeLength And code.Mid(currentPosition + 1, 1) = """" Then
' Increment past the next quote, it is a double quote
currentPosition = currentPosition + 1
Else
inString = False

AddToken(Trim(code.Mid(tokenStartPosition, currentPosition - tokenStartPosition + 1)))
End If
End If

Else
Select Case ch
Case """"
inString = True

Case " "
MaybeAddToken

Case "+", "-"
' Could be a plus symbol or a negative number. This detection logic isn't the best...
If Asc(nextCh) < 48 Or Asc(nextCh) > 57 Then
' Next character is not a number, add this token as a math operation of its own
MaybeAddToken

AddToken(ch)
End If

Case "(", ")", ",", "_", "/", "*", "^", ":", "=", "<", ">", EndOfLine
MaybeAddToken

AddToken(ch)
End Select
End If

currentPosition = currentPosition + 1
Wend

MaybeAddToken

Dim i, columnPosition As Integer
Dim tok, lastTok, nextTok As Token
Dim result, verbose As String

tok = Nil
lastTok = Nil
nextTok = Nil
result = ""
columnPosition = 0

For i = 0 To tokens.UBound
tok = tokens(i)
If i < tokens.UBound Then
nextTok = tokens(i + 1)
Else
nextTok = Nil
End If

If DoDebug Then
verbose = verbose + "' Token: '" + tok.value + "'" + EndOfLine
End If

columnPosition = columnPosition + tok.value.Len

Select Case tok.value
Case "(", ")"
result = result + tok.value

columnPosition = columnPosition + tok.value.Len

Case EndOfLine
result = result + EndOfLine
columnPosition = 0

Else
result = result + tok.value

columnPosition = columnPosition + tok.value.Len
End Select

If tok.value = "<" And nextTok <> Nil And nextTok.value = ">" Then
' Don't add any additional spacing

ElseIf tok.value <> "(" Then
If nextTok <> Nil And nextTok.value <> "(" And nextTok.value <> ")" And nextTok.value <> "," And columnPosition > 0 Then
result = result + " "
columnPosition = columnPosition + 1
End If
End If

lastTok = tok
Next

If DoDebug Then
result = result + EndOfLine + EndOfLine + verbose
End If

Text = Trim(result)
