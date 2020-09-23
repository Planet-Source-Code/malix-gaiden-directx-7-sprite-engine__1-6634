VERSION 5.00
Begin VB.Form ResolutionForm 
   Caption         =   "Select Resolution"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   2325
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "ResolutionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.AddItem "320 X 240"
List1.AddItem "640 X 480"
List1.AddItem "800 X 600"
List1.AddItem "1024 X 768"
End Sub

Private Sub List1_DblClick()
If List1.Text = "320 X 240" Then
    CurResX = 320
    CurResY = 240
ElseIf List1.Text = "640 X 480" Then
    CurResX = 640
    CurResY = 480
ElseIf List1.Text = "800 X 600" Then
    CurResX = 800
    CurResY = 600
ElseIf List1.Text = "1024 X 768" Then
    CurResX = 1024
    CurResY = 768
Else
    MsgBox "Select A Resolution!"
    Exit Sub
End If
Me.Hide
Load VisibleForm
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 List1_DblClick
End If
End Sub
