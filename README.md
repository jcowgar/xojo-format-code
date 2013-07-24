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

Options
=======

Users can set options to control various aspects of Format Code.

The capitalization of keywords
------------------------------

    byREF

Would become:

    ByRef // If CaseConversion = kTitleCase
    byref // If CaseConversion = kLowerCase
    BYREF // If CaseConversion = kUpperCase

Padding of parens
-----------------

    Add(1, Abs(something))

Would become:

    Add(1, Abs(something))     // If PadParens = False
    Add( 1, Abs( something ) ) // If PadParens = True

Keywords
--------

A user definable array lists all the keywords that you wish to have case conversion
performed on. The array comes with all the keywords of the Xojo language already defined.