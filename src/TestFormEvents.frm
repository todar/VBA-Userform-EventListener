VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestFormEvents 
   Caption         =   "Event Testing"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4695
   OleObjectBlob   =   "TestFormEvents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TestFormEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Private Sub UserForm_Activate()
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
End Sub

' Command Button Events
Private Sub Emitter_CommandButtonMouseOver(CommandButton As MSForms.CommandButton)
    CommandButton.Backcolor = 9029664
End Sub

Private Sub Emitter_CommandButtonMouseOut(CommandButton As MSForms.CommandButton)
    CommandButton.Backcolor = 8435998
End Sub

' Textbox Events
Private Sub Emitter_TextboxBlur(Textbox As MSForms.Textbox)
    RendorEventLabel Textbox, Blur
    
    ' CHANGE BORDER COLOR BACK TO A LIGHT GREY
    Textbox.BorderColor = 12434877
    Textbox.BorderStyle = fmBorderStyleNone
    Textbox.BorderStyle = fmBorderStyleSingle
End Sub

Private Sub Emitter_TextboxFocus(Textbox As MSForms.Textbox)
    RendorEventLabel Textbox, Focus
    
    ' CHANGE BORDER COLOR FOR TEXTBOX TO A LIGHT BLUE
    Textbox.BorderColor = 16034051
    Textbox.BorderStyle = fmBorderStyleNone
    Textbox.BorderStyle = fmBorderStyleSingle
End Sub

' Mouse Over/out events
Private Sub Emitter_MouseOut(control As Object)
    RendorEventLabel control, MouseOut
End Sub

Private Sub Emitter_MouseOver(control As Object)
    RendorEventLabel control, MouseOver
End Sub

' Update form to demo what events are happening
Private Sub RendorEventLabel(control As Object, EventName As EmittedEvent)
    Select Case EventName
        Case MouseOver
            MouseOverLabel.Caption = "MouseOver: " & control.name
        
        Case MouseOut
            MouseOutLabel.Caption = "MouseOut: " & control.name
        
        Case Focus
            FocusLabel.Caption = "Focus: " & control.name
            
        Case Blur
            BlurLabel.Caption = "Blur: " & control.name
    End Select
End Sub




