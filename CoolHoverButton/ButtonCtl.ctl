VERSION 5.00
Begin VB.UserControl CoolHoverButton 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   ForeColor       =   &H80000008&
   ScaleHeight     =   315
   ScaleWidth      =   375
   Begin VB.Timer GoTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   2040
   End
   Begin VB.Label Go 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Go!"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   375
   End
   Begin VB.Shape GoHover 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   315
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "CoolHoverButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINT_API) As Long
Private Type POINT_API
    X As Long
    Y As Long
End Type

'Default Property Values:
Const m_def_Enabled = True
'Property Variables:
Dim m_Enabled As Boolean
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Private Sub Go_Click()
If Enabled = True Then
    RaiseEvent Click
End If
End Sub

Private Sub Go_DblClick()
If Enabled = True Then
    RaiseEvent DblClick
End If
End Sub

Private Sub Go_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub Go_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Go_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
If Enabled = True Then
    If GoTimer.Enabled = False Then
        GoTimer.Enabled = True
    End If
End If
End Sub

Private Sub GoTimer_Timer()
If Enabled = True Then
    Dim pnt As POINT_API
    UserControl.ScaleMode = 3
    GetCursorPos pnt
    ScreenToClient UserControl.hWnd, pnt
    If pnt.X < UserControl.ScaleLeft Or _
            pnt.Y < UserControl.ScaleTop Or _
            pnt.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
            pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        GoTimer.Enabled = False
        GoHover.Visible = False
    Else
        GoHover.Visible = True
    End If
End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
If Enabled = True Then
    RaiseEvent Click
End If
End Sub

Private Sub UserControl_DblClick()
If Enabled = True Then
    RaiseEvent DblClick
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled = True Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
    GoHover.BackColor = &HE0E0E0
    GoHover.BorderColor = &H404040
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled = True Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If
    GoHover.BackColor = &HFFC0C0
    GoHover.BorderColor = &H800000
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Go,Go,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Go.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Go.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Go.Caption = PropBag.ReadProperty("Caption", "Go!")
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 315
GoHover.Width = UserControl.ScaleWidth
GoHover.Left = 0
GoHover.Top = 0
GoHover.Height = 315
Go.Height = 195
Go.Width = UserControl.ScaleWidth
Go.Top = 60
Go.Left = 0
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Caption", Go.Caption, "Go!")
End Sub

