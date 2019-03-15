# VBA Userform EventListener
A very easy way to add event listners to a userform.

## Getting Started
> Importing or copying both **EventListnerEmitter.cls** and **EventListnerItem.cls** is **required** in order to work!

Here is basic template, simply add this to a userform.
```vb
Private WithEvents Emitter As EventListnerEmitter

Private Sub UserForm_Activate()
    Set Emitter = New EventListnerEmitter
    Emitter.AddEventListnerAll Me
End Sub
```

That's it, now you can start listening for events!

## Listening for the events

You can listen for all events in one event hanlder **Emitter_EmittedEvent**. see the example below.

```vb
'EXAMPLE SHOWING A BASIC WAY OF DOING A HOVER EFFECT
Private Sub Emitter_EmittedEvent(Control As Object, ByVal EventName As String, ByRef EventValue As Variant)
    
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
This is in the early stages, so feel free to use it as you wish. Currently the events emitted are pretty simple: Click, DblClick, MouseOver, MouseOut, MouseMove, MouseDown, and MouseUp. 

As I have time I'll be adding more events and seeing if I have any needed improvements.

Feel free to do a pull request if you added to it!
