VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Autenticação"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogar 
      Caption         =   "Entrar"
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   3300
      Width           =   3735
   End
   Begin VB.Frame Frame 
      Caption         =   "Painel de login"
      Height          =   2175
      Left            =   180
      TabIndex        =   1
      Top             =   1020
      Width           =   3735
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   180
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label 
         Caption         =   "Senha"
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label 
         Caption         =   "Login"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   420
         Width           =   1035
      End
   End
   Begin VB.Label lblErroLogin 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   300
      TabIndex        =   7
      Top             =   750
      Width           =   3435
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "RINGEX ONLINE"
      BeginProperty Font 
         Name            =   "Source Code Pro"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   4125
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogar_Click()

    If txtLogin.Text <> "mentifg" Then
        lblErroLogin.Caption = "Login inválido"
        Exit Sub
    End If
    
    If txtSenha.Text <> "menti1921" Then
        lblErroLogin.Caption = "Senha inválida"
        Exit Sub
    End If
    
   
    txtLogin.Text = ""
    txtSenha.Text = ""
    lblErroLogin.Caption = ""
    
    frmLogin.Hide
    
    frmServer.WindowState = vbNormal
    frmServer.Show
    
End Sub

Private Sub Form_Resize()

    If frmLogin.WindowState = vbMinimized Then
        frmLogin.Hide
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Cancel = True
    Call DestroyServer

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long, i As Long

    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmLogin.WindowState = vbNormal
            frmLogin.Show
    End Select

End Sub

