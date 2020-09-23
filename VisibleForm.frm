VERSION 5.00
Begin VB.Form VisibleForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ProjectName"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "VisibleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Running = True
SpriteEng.InitializeSystem CurResX, CurResY
DInput.Initialize
MainLoop
End Sub

Sub MainLoop()
On Error GoTo errout
Running = False 'rem this line out when you start programming
Do While Running
    'Game info
Loop

errout:
SpriteEng.TerminateSystem
DInput.Terminate
    If Err.Description <> "" Then
        MsgBox Err.Description
    End If
End

End Sub
