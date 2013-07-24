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
