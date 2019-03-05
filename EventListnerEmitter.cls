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

'ARRAY OF ALL THE DIFFERENT EVENT LISTNERS
Private EventList() As New EventListnerItem

Public Event EmittedEvent(ByRef Control As Object, ByVal EventName As String)

'CALLED BY EVENTLISTENERCOLLECTION CLASS - WAY OF EMITTING AN EVENT
Public Sub EmitEvent(ByRef Control As Object, ByVal EventName As String)
    
    RaiseEvent EmittedEvent(Control, EventName)

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
