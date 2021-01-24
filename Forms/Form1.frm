VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   600
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "How's your battery doing?"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'https://www.nuget.org/packages/Microsoft-WindowsAPICodePack-Core/
'https://www.nuget.org/packages/Microsoft-WindowsAPICodePack-Shell/
'Microsoft.WindowsAPICodePack.ApplicationServices.BatteryState


Private Sub Command1_Click()
    Dim BatS As BatteryState: Set BatS = MPowerManager.GetCurrentBatteryState
    Text1.Text = BatS.ToStr
    
    Dim obj As Object: Set obj = MPower.GetSystemPowerCapabilities
    
    
    'MPower.
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub
