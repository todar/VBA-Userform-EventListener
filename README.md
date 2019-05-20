# VBA Userform EventListener
A very easy way to add event listeners to a userform.

## Getting Started
> Importing or copying both **EventListenerEmitter.cls** and **EventListenerItem.cls** is **required** in order to work!

Here is a basic template, simply add this to a userform.
```vb
Private WithEvents Emitter As EventListenerEmitter

Private Sub UserForm_Activate()
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
End Sub
```

That's it, now you can start listening for events!

## Listening for the events

You can listen for all events in one event handler **Emitter_EmittedEvent**. see the example below.

```vb
'EXAMPLE SHOWING A BASIC WAY OF DOING A HOVER EFFECT
Private Sub Emitter_EmittedEvent(Control As Object, ByVal EventName As EmittedEvent, EventParameters As Scripting.Dictionary)

        'Select statements are really handy working with these events
    Select Case True

        'Change color when mouseover, for a fun hover effect :)
        Case EventName = MouseOver And TypeName(Control) = "CommandButton"
            Control.BackColor = 9029664

        'Don't forget to change it back!
        Case EventName = MouseOut And TypeName(Control) = "CommandButton"
            Control.BackColor = 8435998

    End Select

End Sub
```

You can also listen just to specific events as well.

```vb
Private Sub Emitter_Focus(Control As Object)
    
    'CHANGE BORDER COLOR FOR TEXTBOX TO A LIGHT BLUE
    If TypeName(Control) = "TextBox" Then
        Control.BorderColor = 16034051
    End If
    
End Sub

Private Sub Emitter_Blur(Control As Object)
    
    'CHANGE BORDER COLOR BACK TO A LIGHT GREY
    If TypeName(Control) = "TextBox" Then
        Control.BorderColor = 12434877
    End If
    
End Sub
```

## Note
This is in the early stages, so feel free to use it as you wish. Currently, the events emitted are pretty simple: Click, DoubleClick, MouseOver, MouseOut, MouseMove, MouseDown, and MouseUp. 

As I have time I'll be adding more events and seeing if I have any needed improvements.

Feel free to do a pull request if you added to it or improved it in any way!

**Also, I've posted this code on <a href="https://codereview.stackexchange.com/questions/220370/userform-event-listener-and-emitter">codereview</a>. Feel free to make suggestions or improvements there as well!**
