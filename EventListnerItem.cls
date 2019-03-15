VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListnerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'SET FROM AddEventListner - NEEDED TO EMIT EVENT BACK TO IT.
Private WithEvents pEmitter As EventListnerEmitter
Attribute pEmitter.VB_VarHelpID = -1

'CONTROLS THAT HAVE THE EVENTS
Private WithEvents Form As MSForms.UserForm
Attribute Form.VB_VarHelpID = -1
Private WithEvents Txt As MSForms.Textbox
Attribute Txt.VB_VarHelpID = -1
Private WithEvents Lbl As MSForms.Label
Attribute Lbl.VB_VarHelpID = -1
Private WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Private WithEvents Cmb As MSForms.ComboBox
Attribute Cmb.VB_VarHelpID = -1
Private WithEvents Frm As MSForms.Frame
Attribute Frm.VB_VarHelpID = -1

'PROPERTIES OF CONTROL
Private pControl As Object
Private IsHoveredControl As Boolean
Private IsFocusedControl As Boolean

'TODO
' - ADD EVENTS FOR ALL USERFORM CONTROLS
' - ADD MORE EVENTS THAN THE LIST BELOW
' - DOUBLE CHECK THAT EMITTS ARE ALL CURRENT

'CURRENT EVENTS (WILL ADD MORE AS I HAVE NEED)
' - MouseOver
' - MouseOut
' - MouseMove
' - Click
' - DblClick
' - Focus
' - Blur
' - KeyUp
' - KeyDown

'***********************************************************************************
' PUBLIC FUNCTIONS
'***********************************************************************************

'ONLY PUBLIC METHOD. CALLED FROM EVENTLISTNER CLASS MODULE
Public Sub AddEventListner(ByRef Control As Object, ByRef Emitter As EventListnerEmitter)
    
    Set pEmitter = Emitter
    Set pControl = Control
    
    Select Case TypeName(Control)
        Case "CommandButton"
            Set Btn = Control
        
        Case "ComboBox"
            Set Cmb = Control
            
        Case "Frame"
            Set Frm = Control
        
        Case "Label"
            Set Lbl = Control
        
        Case "TextBox"
            Set Txt = Control
            
        Case Else
 
            If IsUserform(Control) Then
                Set Form = Control
            End If

    End Select
    
End Sub

'***********************************************************************************
' PRIVATE HELPER FUNCTIONS
'***********************************************************************************

'CALLED ON MOUSEMOVE EVENT, THIS IS A WAY OF CREATING A MOUSEOVER AND MOUSEOUT EVENT
Private Sub CheckIfHoveredControl()
    
    If Not IsHoveredControl Then
        IsHoveredControl = True
        pEmitter.EmitEvent pControl, "MouseOver", Array()
    End If
    
End Sub

'CALLED ON MOUSEMOVE EVENT, THIS IS A WAY OF CREATING A MOUSEOVER AND MOUSEOUT EVENT
Private Sub CheckIfFocusedControl()
    
    If Not IsFocusedControl Then
    
        If TypeName(pControl) = "Frame" Then
             pEmitter.SetFocusToControl pControl.ActiveControl
        Else
            IsFocusedControl = True
            pEmitter.EmitEvent pControl, "Focus", Array()
        End If
        
    End If
    
End Sub

'CHECK TO SEE IF OBJ IS A USERFORM
Function IsUserform(ByRef obj As Object) As Boolean
  
    Dim UForm As Object
    
    For Each UForm In VBA.UserForms
      If UForm.Name = obj.Name Then
        IsUserform = True
        Exit For
      End If
    Next
  
End Function


'***********************************************************************************
' EVENTS
'***********************************************************************************

'-------------------------------------------------------------------------
' ONCE AN EVENT HAS EMMITED, EACH EVENT LISTNER WILL LISTEN FOR THAT EVENT
'-------------------------------------------------------------------------
Private Sub pEmitter_EmittedEvent(ByRef Control As Object, ByVal EventName As String, ByRef EventValue As Variant)
    
    'CREATE A MOUSEOVER MOUSEOUT EVENTS
    Select Case EventName
    
        Case "MouseOver"
            If pControl.Name <> Control.Name And IsHoveredControl Then
                IsHoveredControl = False
                pEmitter.EmitEvent pControl, "MouseOut", Array()
            End If
            
        Case "Focus"
            If pControl.Name <> Control.Name And IsFocusedControl Then
                IsFocusedControl = False
                pEmitter.EmitEvent pControl, "Blur", Array()
            ElseIf pControl.Name = Control.Name And IsFocusedControl = False Then
                IsFocusedControl = True
            End If
    
    End Select

End Sub


'------------------------------------------------------------------------
' USERFORM
'------------------------------------------------------------------------
Private Sub Form_Click()
    pEmitter.EmitEvent pControl, "Click", Array()
End Sub

Private Sub Form_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", Array(Cancel)
End Sub

Private Sub Form_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   CheckIfHoveredControl
   pEmitter.EmitEvent pControl, "MouseMove", Array(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseDown", Array(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseUp", Array(Button, Shift, X, Y)
End Sub


'------------------------------------------------------------------------
' COMMAND BUTTON
'------------------------------------------------------------------------
Private Sub Btn_Click()
    pEmitter.EmitEvent pControl, "Click", Array()
End Sub

Private Sub Btn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", Array(Cancel)
End Sub

Private Sub Btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, "MouseMove", Array(Button, Shift, X, Y)
End Sub

Private Sub Btn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, "MouseUp", Array(Button, Shift, X, Y)
End Sub

Private Sub Btn_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseDown", Array(Button, Shift, X, Y)
End Sub

Private Sub Btn_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, "KeyUp", Array(KeyCode, Shift)
End Sub

Private Sub Btn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    pEmitter.EmitEvent pControl, "KeyDown", Array(KeyCode, Shift)
End Sub


'------------------------------------------------------------------------
' LABEL
'------------------------------------------------------------------------
Private Sub Lbl_Click()
    pEmitter.EmitEvent pControl, "Click", Array()
End Sub

Private Sub Lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", Array(Cancel)
End Sub

Private Sub Lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, "MouseMove", Array(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseDown", Array(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseUp", Array(Button, Shift, X, Y)
End Sub


'------------------------------------------------------------------------
' Frame
'------------------------------------------------------------------------
Private Sub Frm_Click()
    pEmitter.EmitEvent pControl, "Click", Array()
End Sub

Private Sub Frm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", Array(Cancel)
End Sub

Private Sub Frm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, "MouseMove", Array(Button, Shift, X, Y)
End Sub

Private Sub Frm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl 'FRAME DOESN'T TAKE FOCUS BUT ACTIVE CONTROL IN FRAME DOES
    pEmitter.EmitEvent pControl, "MouseDown", Array()
End Sub

Private Sub Frm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseUp", Array(Button, Shift, X, Y)
End Sub


'------------------------------------------------------------------------
' Textbox
'------------------------------------------------------------------------
Private Sub Txt_Click()
    pEmitter.EmitEvent pControl, "Click", Array()
End Sub

Private Sub Txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", Array(Cancel)
End Sub

Private Sub Txt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, "MouseMove", Array(Button, Shift, X, Y)
End Sub

Private Sub Txt_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, "MouseUp", Array(Button, Shift, X, Y)
End Sub

Private Sub Txt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseDown", Array(Button, Shift, X, Y)
End Sub

Private Sub Txt_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, "KeyUp", Array(KeyCode, Shift)
End Sub

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    pEmitter.EmitEvent pControl, "KeyDown", Array(KeyCode, Shift)
End Sub


'------------------------------------------------------------------------
' Combobox
'------------------------------------------------------------------------
Private Sub Cmb_Click()
    pEmitter.EmitEvent pControl, "Click", Array()
End Sub

Private Sub Cmb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", Array(Cancel)
End Sub

Private Sub Cmb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, "MouseMove", Array(Button, Shift, X, Y)
End Sub


