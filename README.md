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

It also does some other semi-format tasks, for example it expands common short-hand
assignments:

    appleCount += applesInDeliveryTruck
    appleCount -= applesAteWhileUnloading
    appleCount *= bushelsInTruck
    appleCount /= peopleHelpingToUnload

Automagically becomes:

    appleCount = appleCount + applesInDeliveryTruck
    appleCount = appleCount - applesAteWhileUnloading
    appleCount = appleCount * bushelsInTruck
    appleCount = appleCount / peopleHelpingToUnload

Installation
============

Copy `Format Code.xojo_script` to the `Scripts` directory of your Xojo Installation.
Optionally copy the `Format Code - Test.xojo_script` there also if you wish to run
unit testing, make new features or submit bug reports with failing unit tests.

Options
=======

Users can set options to control various aspects of Format Code.

CaseConversion: The capitalization of keywords
----------------------------------------------

    byREF

Would become:

    ByRef // If CaseConversion = kTitleCase
    byref // If CaseConversion = kLowerCase
    BYREF // If CaseConversion = kUpperCase

PadParensInner: Inner parenthesis padding
-----------------------------------------

    Add ( 1 ,  Abs(something ) )

Would become:

    Add(1, Abs(something))     // If PadParens = False
    Add( 1, Abs( something ) ) // If PadParens = True

PadParensOuter: Outer parenthesis padding
-----------------------------------------

    Add  ( 1 , Abs ( something    ) )

Would become:

    Add(1, Abs(something))   // If PadParensOuter = False
    Add (1, Abs (something)) // If PadParensOuter = True

PadOperators: Operator padding
------------------------------

    a   = 10   + 5

Would become:

    a=10+5     // If PadOperators = False
    a = 10 + 5 // If PadOperators = True

PadComma: Comma padding
-----------------------

    Add(10,   5)

Would become:

    Add(10,5)  // If PadComma = False
    Add(10, 5) // If PadComma = True

Keywords
--------

A user definable array lists all the keywords that you wish to have case conversion
performed on. The array comes with all the keywords of the Xojo language already defined.