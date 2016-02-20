VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Press ESC to open server menu!"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   9000
   ForeColor       =   &H80000010&
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSendGetData.frx":3332
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "DSSDASDASDASADDSASDASDA"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   1635
      Width           =   8760
   End
   Begin VB.Label lblStatus2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "DSSDASDASDASADDSASDASDA"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   135
      TabIndex        =   1
      Top             =   1650
      Width           =   8760
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   On Error GoTo errorhandler

    'Me.Caption = Servers(ServerIndex).Game_Name & " (esc to cancel)"


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmLoad", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

   On Error GoTo errorhandler

    If KeyAscii = vbKeyEscape Then
        Call DestroyTCP
        frmLoad.Hide
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmLoad", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' When the form close button is pressed
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   On Error GoTo errorhandler

    If UnloadMode = vbFormControlMenu Then
        Call DestroyTCP
        frmLoad.Hide
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_QueryUnload", "frmLoad", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub picLoad_Resize()


   On Error GoTo errorhandler
    If IsConnected Then
        frmLoad.BorderStyle = 0
    Else
        frmLoad.BorderStyle = 1
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "picLoad_Resize", "frmLoad", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
