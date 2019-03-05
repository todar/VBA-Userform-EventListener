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
Private WithEvents Txt As MSForms.TextBox
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
Private IsActiveControl As Boolean


'CURRENT EVENTS (WILL ADD MORE AS I HAVE NEED)
' - MouseOver
' - MouseOut
' - Click
' - DblClick


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
        
    End Select
    
    'CHECK IF IT IS A USERFORM
    If IsUserform(Control) Then
        Set Form = Control
    End If
    
End Sub


'***********************************************************************************
' PRIVATE HELPER FUNCTIONS
'***********************************************************************************

'CALLED ON MOUSEMOVE EVENT, THIS IS A WAY OF CREATING A MOUSEOVER AND MOUSEOUT EVENT
Private Sub CheckIfActiveControl()
    
    If Not IsActiveControl Then
        IsActiveControl = True
        pEmitter.EmitEvent pControl, "MouseOver"
    End If
    
End Sub

'CHECK TO SEE IF OBJ IS A USERFORM
Function IsUserform(ByRef Obj As Object) As Boolean
  
    Dim UForm As Object
    
    For Each UForm In VBA.UserForms
      If UForm.Name = Obj.Name Then
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
Private Sub pEmitter_EmittedEvent(ByRef Control As Object, ByVal EventName As String)
    
    'CREATE A MOUSEOVER MOUSEOUT EVENTS
    If EventName = "MouseOver" Then

        If pControl.Name <> Control.Name And IsActiveControl Then
            IsActiveControl = False
            pEmitter.EmitEvent pControl, "MouseOut"
        End If

    End If
    
End Sub


'------------------------------------------------------------------------
' COMMAND BUTTON
'------------------------------------------------------------------------
Private Sub Btn_Click()
    pEmitter.EmitEvent pControl, "Click"
End Sub

Private Sub Btn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick"
End Sub

Private Sub Btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfActiveControl
End Sub


'------------------------------------------------------------------------
' USERFORM
'------------------------------------------------------------------------
Private Sub Form_Click()
    pEmitter.EmitEvent pControl, "Click"
End Sub

Private Sub Form_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick"
End Sub

Private Sub Form_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   CheckIfActiveControl
End Sub


'------------------------------------------------------------------------
' LABEL
'------------------------------------------------------------------------
Private Sub Lbl_Click()
    pEmitter.EmitEvent pControl, "Click"
End Sub

Private Sub Lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick"
End Sub

Private Sub Lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfActiveControl
End Sub


'------------------------------------------------------------------------
' Frame
'------------------------------------------------------------------------
Private Sub Frm_Click()
    pEmitter.EmitEvent pControl, "Click"
End Sub

Private Sub Frm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick"
End Sub

Private Sub Frm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfActiveControl
End Sub


'------------------------------------------------------------------------
' Textbox
'------------------------------------------------------------------------
Private Sub Txt_Click()
    pEmitter.EmitEvent pControl, "Click"
End Sub

Private Sub Txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick"
End Sub

Private Sub Txt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfActiveControl
End Sub


'------------------------------------------------------------------------
' Combobox
'------------------------------------------------------------------------
Private Sub Cmb_Click()
    pEmitter.EmitEvent pControl, "Click"
End Sub

Private Sub Cmb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick"
End Sub

Private Sub Cmb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfActiveControl
End Sub
