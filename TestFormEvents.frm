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
    RunUnitTests
End Sub


Private Sub RunUnitTests()
    
    'CHECK TO SEE IF SETTING FOCUS WORKS ON CONTROLS
    Emitter.SetFocusToControl TextBox1
    Emitter.SetFocusToControl TextBox2
    Emitter.SetFocusToControl CommandButton1
    Emitter.SetFocusToControl CommandButton2
    
End Sub


'EXAMPLE SHOWING A BASIC WAY OF DOING A HOVER EFFECT
Private Sub Emitter_EmittedEvent(Control As Object, ByVal EventName As EmittedEvent, EventParameters As Scripting.Dictionary)
    
    'Select statements are really handy working with these events
    Select Case True

        'Change color when mouseover, for a fun hover effect :)
        Case EventName = MouseOver And TypeName(Control) = "CommandButton"
            Control.Backcolor = 9029664

        'Don't forget to change it back!
        Case EventName = MouseOut And TypeName(Control) = "CommandButton"
            Control.Backcolor = 8435998

    End Select

End Sub

Private Sub Emitter_Focus(Control As Object)
    
    RendorEventLabel Control, Focus
    
    'CHANGE BORDER COLOR FOR TEXTBOX TO A LIGHT BLUE
    If TypeName(Control) = "TextBox" Then
        Control.BorderColor = 16034051
        Control.BorderStyle = fmBorderStyleNone
        Control.BorderStyle = fmBorderStyleSingle
    End If
    
End Sub

Private Sub Emitter_Blur(Control As Object)
    
    RendorEventLabel Control, Blur
    
    'CHANGE BORDER COLOR BACK TO A LIGHT GREY
    If TypeName(Control) = "TextBox" Then
        Control.BorderColor = 12434877
                Control.BorderStyle = fmBorderStyleNone
        Control.BorderStyle = fmBorderStyleSingle
    End If
    
End Sub

Private Sub Emitter_MouseOut(Control As Object)
    RendorEventLabel Control, MouseOut
End Sub

Private Sub Emitter_MouseOver(Control As Object)
    RendorEventLabel Control, MouseOver
End Sub

Private Sub RendorEventLabel(Control As Object, EventName As EmittedEvent)
    
    Select Case EventName
        
        Case MouseOver
            MouseOverLabel.Caption = "MouseOver: " & Control.Name
        
        Case MouseOut
            MouseOutLabel.Caption = "MouseOut: " & Control.Name
        
        Case Focus
            FocusLabel.Caption = "Focus: " & Control.Name
            
        Case Blur
            BlurLabel.Caption = "Blur: " & Control.Name
            
    End Select
    
End Sub




