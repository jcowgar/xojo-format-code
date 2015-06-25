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

Copy `Scripts/Format Code.xojo_script` to the `Scripts` directory of your Xojo
Installation. To run unit testing, simply open `Format Code Unit Test.xojo_project`
and click the menu item *File > IDE Scripts > Format Code - Test.xojo_script*.

Bleeding Edge Installation
--------------------------

If you are keeping up with the bleeding edge Format Code development via git, you may
wish to clone the repository, then create symbolic links instead of copying the script
each update. An example for OS X at the time of this writing:

    $ cd $HOME/development
    $ git clone https://github.com/jcowgar/xojo-format-code.git
    $ cd xojo-format-code
    $ git checkout develop
    $ ln -s Format\ Code.xojo_script /Applications/Xojo\ 2015\ Release\ 2.2/Scripts/

Now when a new commit is pushed or release made, simply:

    $ cd $HOME/development/xojo-format-code
    $ git pull

and you will be running the bleeding edge.

Creating a Shortcut
-------------------

On OS X you can use the System Preferences to create a shortcut to execute the Format
Code script.

1. Load System Preferences
2. Select *Keyboard*
3. Select the *Shortcuts* tab
4. Select *App Shortcuts*
5. Click the *+* button
6. Choose Application *Xojo*
7. Enter *File->IDE Scripts->Format Code.xojo_script* into the Menu Title
8. Choose a keyboard shortcut (I use Ctrl+F for Format)

Options
=======

Users can set options to control various aspects of Format Code. Options can be set one
of two ways, 

1. Change configuration settings at the top of `Format Code.xojo_script`: This will
   enable your preferences to be applied across multiple projects w/o change. It does
   not allow for custom settings per project and each time you update
   `Format Code.xojo_script` you will have to re-apply your custom settings.
2. Create a module named `FormatCodePreferences` and add constants with the same names
   as the variables in the script. This allows you to have per-project settings and
   be able to update `Format Code.xojo_script` easily. This module does not have to
   contain all, or any, of the preferences but any preference that is found in the
   module will override any preference set manually in the `Format Code.xojo_script`
   file.
   
You can also use both options, #1 to provide defaults to simple projects and #2 to override
those defaults for specific projects.

When setting options via the `FormatCodePreferences`, the CaseConversion constant can
be numeric (1, 2, 3) or a string, "kTitleCase", "kLowerCase", "kUpperCase" and boolean
constants can be a boolean value or a string, "Yes" or "True"

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

KeywordsToTitleCase
-------------------

A user definable array defined in the `FormatCodePreferences` module delimited by commas.
All items will be converted to `TitleCase` regardless of any case conversion setting.

KeywordsToUpperCase
-------------------

A user definable array defined in the `FormatCodePreferences` module delimited by commas.
All items will be converted to `UPPERCASE` regardless of any case conversion setting.

KeywordsToLowerCase
-------------------

A user definable array defined in the `FormatCodePreferences` module delimited by commas.
All items will be converted to `lowercase` regardless of any case conversion setting.
