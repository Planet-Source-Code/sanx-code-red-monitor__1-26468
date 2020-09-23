VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save Log"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton butSave 
      Caption         =   "&Save"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butSave_Click()

frmSave.Hide
frmMain.SaveLog txtPath.Text

End Sub

Private Sub Form_Load()

Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub
