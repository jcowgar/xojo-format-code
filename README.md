Format Code - Xojo IDE Script
=============================

Todo
----

* Track current column better and wrap long lines where able to
* If `SelLen > 0 Then` work on `SelText`, not `Text`
* Split `.` as its own token amoungst words, and capitalize them as 
  needed.
* Better detection of positive/negative numbers vs. plus/minus
* Detect comments and format better, for example:
    If 1 Then'blah 
should be translated to
    If 1 Then ' blah
* Don't rely on source code's `EndOfLine`, skip that as a token and insert
  our own `EndOfLine`'s.

Additional Formatting
---------------------

With an option, blank lines should be added for clarity, for example:

    Dim i As Integer
    Dim b As String
    DoCode
 
Should become
 
    Dim i As Integer
    Dim b As String

    DoCode

A `Select Case` statement example:

    Select Case Name
    Case "John"
      DoSomething
    Case "Jeff"
      DoSomethingElse
    End Select
 
Should become

    Select Case Name
    Case "John"
      DoSomething

    Case "Jeff"
      DoSomethingElse
    End Select
