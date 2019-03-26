VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListnerEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'ARRAY OF ALL THE DIFFERENT EVENT LISTNERS FOR EVERY USERFORM CONTROL +FORM ITSELF
Private EventList() As New EventListnerItem

'ALL CURRENT POSSIBLE EVENTS THAT CAN BE EMMITED. NOTE, EMITTEDEVENT IS SENT FOR ALL!
Public Event EmittedEvent(ByRef Control As Object, ByVal EventName As String, ByRef EventValue As Variant)

Public Event Click(ByRef Control As Object)
Public Event DblClick(ByRef Control As Object, ByRef Cancel As MSForms.ReturnBoolean)
Public Event KeyUp(ByRef Control As Object, ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
Public Event KeyDown(ByRef Control As Object, ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
Public Event MouseOver(ByRef Control As Object)
Public Event MouseOut(ByRef Control As Object)
Public Event MouseMove(ByRef Control As Object, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
Public Event Focus(ByRef Control As Object)
Public Event Blur(ByRef Control As Object)

'***********************************************************************************
' PUBLIC METHODS
'***********************************************************************************

'CALLED BY EVENTLISTENERCOLLECTION CLASS - MAIN ENTRYWAY OF EMITTING ALL EVENTS
Public Sub EmitEvent(ByRef Control As Object, ByVal EventName As String, ByRef EventValue As Variant)
    
    'EVENT RAISED FOR ALL EVENTS. A WAY FOR THE USER TO COLLECT FROM A SINGLE LOCATION.
    RaiseEvent EmittedEvent(Control, EventName, EventValue)
    
    'SPECIFIC EVENTS PER OBJECT TYPE
    Select Case EventName
    
        Case "Click"
            RaiseEvent Click(Control)
        
        Case "DblClick"
            RaiseEvent DblClick(Control, CVar(EventValue(0)))
        
        Case "KeyUp"
            RaiseEvent KeyUp(Control, CVar(EventValue(0)), CVar(EventValue(1)))
            
        Case "KeyDown"
            RaiseEvent KeyDown(Control, CVar(EventValue(0)), CVar(EventValue(1)))
        
        Case "MouseOver"
            RaiseEvent MouseOver(Control)
            
        Case "MouseOut"
            RaiseEvent MouseOut(Control)
            
        Case "Focus"
            RaiseEvent Focus(Control)
            
        Case "Blur"
            RaiseEvent Blur(Control)
            
        Case "MouseMove"
            RaiseEvent MouseMove(Control, CVar(EventValue(0)), CVar(EventValue(1)), CVar(EventValue(2)))
        
    End Select

End Sub

'MUST CALL THIS IF YOU WANT TO PROGRAMATICALLY SET CONTROL. OTHERWISE EVENT'S WILL BE OFF!
Public Sub SetFocusToControl(ByRef Control As Object)
    
    If Not Control Is Nothing Then
        Control.setFocus
        EmitEvent Control, "Focus", ""
    End If
    
End Sub

'ADD EVENT LISTNERS ON SPECIFIC CONTROLS - ALSO CALLED BY AddEventListnerAll
Public Sub AddEventListner(ByRef Control As Object)
    
    If IsArrayEmpty(EventList) Then
        ReDim EventList(0 To 0)
    Else
        ReDim Preserve EventList(0 To UBound(EventList) + 1)
    End If
    
    EventList(UBound(EventList)).AddEventListner Control, Me
    
End Sub

'ADD EVENT LISTNER TO ALL CONTROLS INCLUDING THE FORM
Public Sub AddEventListnerAll(ByRef Form As Object)
    
    AddEventListner Form
    
    Dim ctrl As MSForms.Control
    For Each ctrl In Form.Controls
        AddEventListner ctrl
    Next ctrl
    
End Sub

'HELPER FUNCTION TO TEST FOR EMPTY ARRAY
Private Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CPEARSON
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


