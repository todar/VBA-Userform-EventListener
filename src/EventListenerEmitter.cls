VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListenerEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' Array of all the different event listeners for every userform control and the form itself
Private EventList() As New EventListenerItem

' All the current possible events that can be emitted.
' Note, EmittedEvent is sent for all events!
Public Event EmittedEvent(ByRef control As Object, ByVal EventType As EmittedEvent, ByRef EventParameters As Collection)
Public Event Click(ByRef control As Object)
Public Event DblClick(ByRef control As Object, ByRef Cancel As MSForms.ReturnBoolean)
Public Event KeyUp(ByRef control As Object, ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
Public Event KeyDown(ByRef control As Object, ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
Public Event MouseOver(ByRef control As Object)
Public Event MouseOut(ByRef control As Object)
Public Event MouseMove(ByRef control As Object, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
Public Event Focus(ByRef control As Object)
Public Event Blur(ByRef control As Object)
Public Event Change(ByRef control As Object)

' Events to Labels
Public Event LabelMouseOver(ByRef Label As MSForms.Label)
Public Event LabelMouseOut(ByRef Label As MSForms.Label)
Public Event LabelClick(ByRef Label As MSForms.Label)
Public Event LabelDoubleClick(ByRef Label As MSForms.Label, ByRef Cancel As MSForms.ReturnBoolean)
Public Event LabelMouseMove(ByRef Label As MSForms.Label, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)

' Events to Textboxes
Public Event TextboxFocus(ByRef Textbox As MSForms.Textbox)
Public Event TextboxBlur(ByRef Textbox As MSForms.Textbox)
Public Event TextboxMouseOver(ByRef Textbox As MSForms.Textbox)
Public Event TextboxMouseOut(ByRef Textbox As MSForms.Textbox)
Public Event TextboxClick(ByRef Textbox As MSForms.Textbox)
Public Event TextboxDoubleClick(ByRef Textbox As MSForms.Textbox, ByRef Cancel As MSForms.ReturnBoolean)
Public Event TextboxMouseMove(ByRef Textbox As MSForms.Textbox, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)

' Events to CommandButtons
Public Event CommandButtonMouseOver(ByRef CommandButton As MSForms.CommandButton)
Public Event CommandButtonMouseOut(ByRef CommandButton As MSForms.CommandButton)
Public Event CommandButtonClick(ByRef CommandButton As MSForms.CommandButton)
Public Event CommandButtonDoubleClick(ByRef CommandButton As MSForms.CommandButton, ByRef Cancel As MSForms.ReturnBoolean)
Public Event CommandButtonMouseMove(ByRef CommandButton As MSForms.CommandButton, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)

' Types of events that can occur
Public Enum EmittedEvent
    Click
    DoubleClick
    MouseMove
    MouseOut
    MouseOver
    MouseDown
    MouseUp
    KeyUp
    KeyDown
    Focus
    Blur
End Enum

' Called by EventListenerItem class - main entryway of emitting all events
Public Sub EmitEvent(ByRef control As Object, ByVal EventType As EmittedEvent, Optional ByRef EventParameters As Collection)
    ' Event raised for all events. This is a way for the user to collect from a single location.
    RaiseEvent EmittedEvent(control, EventType, EventParameters)
    
    ' Specific events
    Select Case EventType
    
        Case Click
            RaiseEvent Click(control)
        
        Case DoubleClick
            RaiseEvent DblClick(control, EventParameters("Cancel"))
        
        Case KeyUp
            RaiseEvent KeyUp(control, EventParameters("KeyCode"), EventParameters("Shift"))
            
        Case KeyDown
            RaiseEvent KeyDown(control, EventParameters("KeyCode"), EventParameters("Shift"))
        
        Case MouseOver
            RaiseEvent MouseOver(control)
            
        Case MouseOut
            RaiseEvent MouseOut(control)
            
        Case Focus
            RaiseEvent Focus(control)
            
        Case Blur
            RaiseEvent Blur(control)
            
        Case MouseMove
            RaiseEvent MouseMove(control, EventParameters("Shift"), EventParameters("X"), EventParameters("Y"))
        
    End Select
    
    ' Call the specific control type events
    Select Case TypeName(control)
        Case "Label"
            EmitLabelEvent control, EventType, EventParameters
            
        Case "Textbox"
            EmitTextboxEvent control, EventType, EventParameters
            
        Case "CommandButton"
            EmitCommandButtonEvent control, EventType, EventParameters
        
    End Select
End Sub

' Events for Labels
Private Sub EmitLabelEvent(ByRef Label As MSForms.Label, ByVal EventType As String, ByRef EventParameters As Collection)
    Select Case EventType
        Case EmittedEvent.Click
            RaiseEvent LabelClick(Label)
            
        Case EmittedEvent.DoubleClick
            RaiseEvent LabelDoubleClick(Label, EventParameters("Cancel"))
            
        Case EmittedEvent.MouseOver
            RaiseEvent LabelMouseOver(Label)
            
        Case EmittedEvent.MouseOut
            RaiseEvent LabelMouseOut(Label)

        Case EmittedEvent.MouseMove
            RaiseEvent MouseMove(Label, EventParameters("Shift"), EventParameters("X"), EventParameters("Y"))
    End Select
End Sub

' Events for Textboxes
Private Sub EmitTextboxEvent(ByRef Textbox As MSForms.Textbox, ByVal EventType As String, ByRef EventParameters As Collection)
    Select Case EventType
        Case EmittedEvent.Blur
            RaiseEvent TextboxBlur(Textbox)
        
        Case EmittedEvent.Focus
            RaiseEvent TextboxFocus(Textbox)
    
        Case EmittedEvent.Click
            RaiseEvent TextboxClick(Textbox)
            
        Case EmittedEvent.DoubleClick
            RaiseEvent TextboxDoubleClick(Textbox, EventParameters("Cancel"))
            
        Case EmittedEvent.MouseOver
            RaiseEvent TextboxMouseOver(Textbox)
            
        Case EmittedEvent.MouseOut
            RaiseEvent TextboxMouseOut(Textbox)

        Case EmittedEvent.MouseMove
            RaiseEvent MouseMove(Textbox, EventParameters("Shift"), EventParameters("X"), EventParameters("Y"))
    End Select
End Sub

' Events for CommandButton
Private Sub EmitCommandButtonEvent(ByRef CommandButton As MSForms.CommandButton, ByVal EventType As String, ByRef EventParameters As Collection)
    Select Case EventType
        Case EmittedEvent.Click
            RaiseEvent CommandButtonClick(CommandButton)
            
        Case EmittedEvent.DoubleClick
            RaiseEvent CommandButtonDoubleClick(CommandButton, EventParameters("Cancel"))
            
        Case EmittedEvent.MouseOver
            RaiseEvent CommandButtonMouseOver(CommandButton)
            
        Case EmittedEvent.MouseOut
            RaiseEvent CommandButtonMouseOut(CommandButton)

        Case EmittedEvent.MouseMove
            RaiseEvent MouseMove(CommandButton, EventParameters("Shift"), EventParameters("X"), EventParameters("Y"))
    End Select
End Sub

' MUST CALL THIS IF YOU WANT TO programmatically SET CONTROL! OTHERWISE, EVENT'S WILL BE OFF!
Public Sub SetFocusToControl(ByRef control As Object)
    'If the user was to set focus through VBA then this code will fall apart considering
    'it is unaware of that event occurring.
    If Not control Is Nothing Then
        control.setFocus
        EmitEvent control, Focus
    End If
End Sub

' ADD EVENT Listeners ON SPECIFIC CONTROLS - ALSO CALLED BY AddEventListenerAll
Public Sub AddEventListener(ByRef control As Object)
    ' Events are stored in a private EventListenerItem array
    If IsArrayEmpty(EventList) Then
        ReDim EventList(0 To 0)
    Else
        ReDim Preserve EventList(0 To UBound(EventList) + 1)
    End If
    
    'CALL AddEventListener IN EventListenerItem. THIS IS KEPT IN
    EventList(UBound(EventList)).AddEventListener control, Me
End Sub

'ADD EVENT Listener TO ALL CONTROLS INCLUDING THE FORM
Public Sub AddEventListenerAll(ByRef Form As Object)
    AddEventListener Form
    
    Dim control As MSForms.control
    For Each control In Form.Controls
        AddEventListener control
    Next control
End Sub

Private Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CPEARSON: http://www.cpearson.com/excel/VBAArrays.htm
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    Dim ub As Long
    ub = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' On rare occasion, under circumstances I cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array. On these occasions, LBound is 0 and
        ' UBound is -1. To accommodate the weird behavior, test to see if LB > UB.
        ' If so, the array is not allocated.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Dim LB As Long
        LB = LBound(Arr)
        If LB > ub Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If
End Function
