# VBA Userform EventListener
A very easy way to add event listners to a userform.

## Getting Started
> Imported or copying both **EventListnerEmitter.cls** and **EventListnerItem.cls** are **required** in order to work!

To use in a userform try this basic template to get started
```vb
Private WithEvents Emitter As EventListnerEmitter

Private Sub UserForm_Activate()
    Set Emitter = New EventListnerEmitter
    Emitter.AddEventListnerAll Me
End Sub
```

That's it, now you can start listening for events!

## Listening for the events
Currently all events are sent to one event hanlder, see the example below to get started.

```vb
Private Sub Emitter_EmittedEvent(Control As Object, ByVal EventName As String)
    
    'display name of control that was clicked
    If EventName = "Click" Then
        msgbox Control.Name

    'Change color when mouseover, for a fun hover effect :)
    ElseIf EventName = "MouseOver" And TypeName(Control) = "CommandButton" Then
        Control.Backcolor = 9029664

    'Don't forget to change it back!    
    ElseIf EventName = "MouseOut" And TypeName(Control) = "CommandButton" Then
         Control.Backcolor = 8435998
        
    End If

End Sub
```

## Note
This is in the early stages, so feel free to use it as you wish. Currently the events emitted are pretty simple: Click, DblClick, MouseOver, and MouseOut. 

As I have time I'll be adding more events and seeing if I have any needed improvements.

Feel free to do a pull request if you added to it!
