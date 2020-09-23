VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Listening on port 80"
   ClientHeight    =   645
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox lstLog 
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox lstDisplay 
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   5040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popop"
      Visible         =   0   'False
      Begin VB.Menu mnuClearLog 
         Caption         =   "&Clear Log"
      End
      Begin VB.Menu mnuSaveLog 
         Caption         =   "&Save Log"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This application is copyright Sanx, 2001. This application is
'offered as freeware, and as such, you may copy, modify and
'distribute it without conditions, provided this copyright
'notice remains.
'http://www.sanx.org

Private Sub Form_Load()

Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
sckMain.LocalPort = 80
sckMain.Listen
lstDisplay.AddItem "Listening ..."

End Sub

Private Sub Form_Resize()

lstDisplay.Height = frmMain.Height - 360
lstDisplay.Width = frmMain.Width - 120

End Sub

Private Sub lstDisplay_DblClick()

If lstDisplay.ListIndex > 0 Then
    If lstLog.List(lstDisplay.ListIndex - 1) <> "" Then
        MsgBox lstLog.List(lstDisplay.ListIndex - 1), vbOKOnly, "Data Received"
    Else
        MsgBox "No data received", vbOKOnly, "Data Received"
    End If
End If

End Sub

Private Sub lstDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    PopupMenu mnuPopup
End If

End Sub

Private Sub mnuClearLog_Click()

lstDisplay.Clear
lstLog.Clear

End Sub

Private Sub mnuSaveLog_Click()

frmSave.Show vbModal

End Sub

Private Sub sckMain_ConnectionRequest(ByVal requestID As Long)

lstDisplay.AddItem Str$(Now) & " - Connection request from: " & sckMain.RemoteHostIP
lstDisplay.Selected(lstDisplay.ListCount - 1) = True
sckMain.Close
sckMain.Accept requestID

End Sub

Public Sub SaveLog(savePath As String)

Dim count As Integer

On Error GoTo ErrHit

Open savePath For Output As #1

For count = 1 To (lstDisplay.ListCount - 1)
    Print #1, lstDisplay.List(count)
    Print #1, lstLog.List(count - 1)
Next count

Close #1

Exit Sub
ErrHit:
MsgBox "Error occured during Save operation. Please check the path and try again.", vbApplicationModal + vbCritical + vbOKOnly, "Code Red Watcher"
frmSave.txtPath.Text = savePath
frmSave.Show vbModal

End Sub

Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)

Dim tempStr As String

sckMain.GetData tempStr

CloseSocket tempStr

End Sub

Private Sub CloseSocket(tempStr As String)

If tempStr <> "" Then
    lstLog.AddItem Str$(Now) & vbCrLf & "Data received from: " & sckMain.RemoteHostIP & vbCrLf & tempStr
    sckMain.SendData "HTTP/1.1 200 OK" & vbCrLf
    sckMain.SendData "Server: Sanx/1.0" & vbCrLf
    sckMain.SendData Now & vbCrLf
    sckMain.SendData "Keep-Alive: timeout=2, max=100" & vbCrLf
    sckMain.SendData "Connection: keep-alive" & vbCrLf
    sckMain.SendData "Content-Location: http://www.why.dont.you.just.sod.off/now.html" & vbCrLf
    sckMain.SendData "Content-Type: text/html" & vbCrLf & vbCrLf
    sckMain.SendData "<HTML><BODY><H1>Sod Off!</H1></BODY></HTML>" & vbCrLf
    sckMain.SendData vbCrLf & vbCrLf
End If

tmrMain.Enabled = True

End Sub

Private Sub tmrMain_Timer()

If sckMain.State <> sckClosed Then
    sckMain.Close
    sckMain.Listen
End If

tmrMain.Enabled = False

End Sub
