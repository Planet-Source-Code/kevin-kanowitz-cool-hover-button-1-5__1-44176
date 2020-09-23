VERSION 5.00
Object = "*\ACoolHoverButtonLib.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin CoolHoverButtonLib.CoolHoverButton CoolHoverButton1 
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      Caption         =   "Un-toched"
   End
   Begin CoolHoverButtonLib.CoolHoverButton CoolHoverButton1 
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      Enabled         =   0   'False
      Caption         =   "Hover"
   End
   Begin CoolHoverButtonLib.CoolHoverButton CoolHoverButton1 
      Height          =   315
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      Caption         =   "Clicked"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CoolHoverButton1_Click(Index As Integer)
MsgBox "Test"
End Sub
