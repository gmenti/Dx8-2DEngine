VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carregando servidor..."
   ClientHeight    =   1245
   ClientLeft      =   6375
   ClientTop       =   4110
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   83
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrNotifications 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar pbarLoading 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblProg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando..."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   8535
   End
   Begin VB.Label lblNotifications 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   9015
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Long, x As Long
   On Error GoTo errorhandler

    Me.Show
    InitServer


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   On Error GoTo errorhandler

    End


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_QueryUnload", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblExitServer_Click()
    

   On Error GoTo errorhandler
    End
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblExitServer_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblExitServer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    

   On Error GoTo errorhandler
   
   lblExitServer.Font.Bold = True
   lblNewUser.Font.Bold = False
   lblExistingUser.Font.Bold = False
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblExitServer_MouseMove", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub tmrNotifications_Timer()
    Static LastNotification As String
    Static TimeShown As Long
    

   On Error GoTo errorhandler

    If lblNotifications.Caption <> "" Then
        If lblNotifications.Caption = LastNotification Then
            If TimeShown >= 6 Then
                LastNotification = ""
                lblNotifications.Caption = ""
                TimeShown = 0
            Else
                TimeShown = TimeShown + 1
            End If
        Else
            LastNotification = lblNotifications.Caption
            TimeShown = 0
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "tmrNotifications_Timer", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

