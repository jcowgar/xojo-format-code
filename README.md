Format Code - Xojo IDE Script
=============================

Format's the code in the current editor of the [Xojo IDE](http://www.xojo.com).
For example, badly formatted code:

    DIM i as Integer= 10
    iF i<18    then
      SayHello  ("Howdy"  , "World"  )
    end if
    
Automagically becomes:

    Dim i As Integer = 10
    If i < 18 Then
      SayHello("Howdy", "World")
    End If    

Todo
----

* Track current column and wrap long lines where able to
* Split `.` as its own token amoungst words, and capitalize them as 
  needed.
* Better detection of positive/negative numbers vs. plus/minus
* Detect comments and format better, for example:
    `If 1 Then'blah`
should be translated to
    `If 1 Then ' blah`
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
