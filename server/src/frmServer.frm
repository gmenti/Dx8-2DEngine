VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13305
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   360
      TabIndex        =   69
      Top             =   3960
      Width           =   495
   End
   Begin VB.Timer tmrNotifications 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Server Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   12120
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fraControlPanel 
      Caption         =   "Painel de Controle"
      Height          =   6735
      Left            =   2760
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   9855
      Begin VB.PictureBox picMapName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   5280
         ScaleHeight     =   2505
         ScaleWidth      =   2745
         TabIndex        =   60
         Top             =   2880
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CommandButton cmdReserveMaps 
            Caption         =   "Alterar"
            Height          =   315
            Left            =   360
            TabIndex        =   64
            Top             =   2040
            Width           =   1815
         End
         Begin VB.TextBox txtMapName 
            Height          =   285
            Left            =   120
            TabIndex        =   63
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox txtRMaps2 
            Height          =   285
            Left            =   120
            TabIndex        =   62
            Text            =   "0"
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtRMap 
            Height          =   285
            Left            =   120
            TabIndex        =   61
            Text            =   "0"
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Ultimo:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Primeiro:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdShutDown 
         Caption         =   "Desligar servidor"
         Height          =   495
         Left            =   5280
         TabIndex        =   59
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Alterar nome dos mapas"
         Height          =   375
         Left            =   1800
         TabIndex        =   43
         Top             =   3120
         Width           =   2775
      End
      Begin VB.CheckBox chkServerLog 
         BackColor       =   &H80000004&
         Caption         =   "Server Log"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         MaskColor       =   &H00000000&
         TabIndex        =   42
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkDisableRestart 
         BackColor       =   &H80000004&
         Caption         =   "Desativar Reinicio Automatico"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   37
         Top             =   720
         Width           =   3255
      End
      Begin VB.CheckBox chkStaffOnly 
         BackColor       =   &H80000004&
         Caption         =   "Somente Staff?"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   36
         Top             =   480
         Width           =   2655
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Caption         =   "Map Report"
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   1800
         TabIndex        =   33
         Top             =   360
         Width           =   2775
         Begin VB.ListBox lstMaps 
            Height          =   2205
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraDatabase 
         BackColor       =   &H80000004&
         Caption         =   "Recarregar"
         ForeColor       =   &H00000000&
         Height          =   3495
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1455
         Begin VB.CommandButton cmdLoadOptions 
            Caption         =   "Options"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   3120
            Width           =   1155
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Classes"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   2400
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2760
            Width           =   1215
         End
      End
   End
   Begin VB.Frame fraHousing 
      Caption         =   "Configurações"
      Height          =   6735
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   9855
      Begin VB.Frame fraNews 
         BackColor       =   &H80000004&
         Caption         =   "Notícias do menu"
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   6720
         TabIndex        =   56
         Top             =   3720
         Width           =   2895
         Begin VB.TextBox txtNews 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   58
            Text            =   "frmServer.frx":1708A
            Top             =   240
            Width           =   2655
         End
         Begin VB.CommandButton cmdSaveNews 
            Caption         =   "Salvar notícias"
            Height          =   315
            Left            =   240
            TabIndex        =   57
            Top             =   1200
            Width           =   2415
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000004&
         Caption         =   "Configurações (Padrão)"
         ForeColor       =   &H00000000&
         Height          =   1935
         Left            =   240
         TabIndex        =   49
         Top             =   3720
         Width           =   3255
         Begin VB.CommandButton cmdSaveDataFolder 
            Caption         =   "Salvar"
            Height          =   315
            Left            =   1560
            TabIndex        =   52
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtDataFolder 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   51
            Text            =   "txtDataFolder"
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox txtUpdateUrl 
            Height          =   285
            Left            =   120
            TabIndex        =   50
            Text            =   "txtUpdateUrl"
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Label lblDataFolder 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Pasta padrão do jogo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label12 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Update.ini URL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblUpdateHelp 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuda"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000004&
         Caption         =   "Créditos"
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   3600
         TabIndex        =   46
         Top             =   3720
         Width           =   3015
         Begin VB.CommandButton cmdSaveCredits 
            Caption         =   "Salvar créditos"
            Height          =   315
            Left            =   240
            TabIndex        =   48
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtCredits 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   47
            Text            =   "frmServer.frx":17092
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "House Setup"
         Height          =   2895
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Width           =   8535
         Begin VB.ListBox lstHouses 
            Height          =   2400
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtHouseName 
            Height          =   255
            Left            =   4200
            TabIndex        =   15
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtBaseMap 
            Height          =   285
            Left            =   4200
            TabIndex        =   14
            Top             =   645
            Width           =   2655
         End
         Begin VB.TextBox txtHouseFurniture 
            Height          =   285
            Left            =   4200
            TabIndex        =   13
            Top             =   2085
            Width           =   2655
         End
         Begin VB.CommandButton cmdSaveHouse 
            Caption         =   "Save Changes"
            Height          =   855
            Left            =   7200
            TabIndex        =   12
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtHousePrice 
            Height          =   285
            Left            =   4200
            TabIndex        =   11
            Top             =   1725
            Width           =   2655
         End
         Begin VB.TextBox txtXEntrance 
            Height          =   285
            Left            =   4200
            TabIndex        =   10
            Top             =   1005
            Width           =   2655
         End
         Begin VB.TextBox txtYEntrance 
            Height          =   285
            Left            =   4200
            TabIndex        =   9
            Top             =   1365
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Name of House:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   22
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblHouseMap 
            BackStyle       =   0  'Transparent
            Caption         =   "Base Map:"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2520
            TabIndex        =   21
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Max Pieces of Furniture (0 for no max):"
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   2520
            TabIndex        =   20
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblHousePrice 
            BackStyle       =   0  'Transparent
            Caption         =   "Price:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            TabIndex        =   19
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrance X:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            TabIndex        =   18
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrance Y:"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2520
            TabIndex        =   17
            Top             =   1320
            Width           =   1575
         End
      End
   End
   Begin VB.Frame fraConsole 
      Caption         =   "Console"
      Height          =   6735
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   9855
      Begin VB.TextBox txtText 
         Height          =   2055
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3840
         Width           =   9255
      End
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   6000
         Width           =   8655
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2415
         Left            =   240
         TabIndex        =   44
         Top             =   840
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4260
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Jogadores"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblCPS 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "CPS: 0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label lblCpsLock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "[Unlock]"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   3480
         Width           =   720
      End
   End
   Begin VB.Label lblServerMenuOpt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "EXTRA"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblServerMenuOpt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   41
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblServerMenuOpt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Painel de Controle"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   40
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblServerMenuOpt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Configurações"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   39
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblServerMenuOpt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Console"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblDisableAutoLogin 
      BackStyle       =   0  'Transparent
      Caption         =   "Disable Auto Login!"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblNotifications 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Notification: ................"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   11655
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuEditPlayer 
         Caption         =   "Edit Player"
      End
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkStaffOnly_Click()
Dim i As Long

   On Error GoTo errorhandler
    If chkStaffOnly.Value = 1 Then
        Options.StaffOnly = 1
        SaveOptions
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If GetPlayerAccess(i) = 0 Then
                    AlertMsg i, "Sorry, the server was switched to staff-only mode. Please check back later!"
                End If
            End If
        Next
    Else
        Options.StaffOnly = 0
        SaveOptions
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkStaffOnly_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub



Private Sub cmdAccess_Click()
    Dim help As String

   On Error GoTo errorhandler

    help = "Access is defined by 5 numbers..."
    help = help & vbNewLine & "0: Normal Player"
    help = help & vbNewLine & "1: Moderator - Can kick/warp and simple admin functions."
    help = help & vbNewLine & "2: Mapper - Mod Powers + Mapping Abilities"
    help = help & vbNewLine & "3: Developer - Mapper Powers and ability to edit all game content."
    help = help & vbNewLine & "4: Creator - All Powers, for owner(s) of the game)"
    
    MsgBox help


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAccess_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Private Sub cmdCancelLogin_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancelLogin_Click", "frumServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseAccount_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseAccount_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseConsole_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseConsole_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseControlPanel_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseControlPanel_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseHousing_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseHousing_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdClosePlayers_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdClosePlayers_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub


Private Sub cmdLoadOptions_Click()
    LoadOptions
    Call TextAdd("All options reloaded.")
End Sub

Private Sub cmdReserveMaps_Click()
Dim map1 As Long, map2 As Long, i As Long

   On Error GoTo errorhandler

    If IsNumeric(txtRMap.Text) Then
        If IsNumeric(txtRMaps2.Text) Then
            map1 = Val(txtRMap.Text)
            map2 = Val(txtRMaps2.Text)
            If map1 > map2 Or map1 < 1 Or map1 > MAX_MAPS Or map2 < 1 Or map2 > MAX_MAPS Then
                lblNotifications.Caption = "An error occured. One of the map values are invalid. The first value must be the smaller one."
                lblNotifications.ForeColor = &HFF&
                Exit Sub
            Else
                For i = map1 To map2
                    Map(i).Name = txtMapName.Text
                    Map(i).Revision = Map(i).Revision + 1
                    MapCache_Create i
                    SaveMap i
                Next
                picMapName.Visible = False
                lblNotifications.Caption = "Maps reserved."
                lblNotifications.ForeColor = &HC000&
                UpdateMapReport
                Exit Sub
            End If
        End If
    End If
    lblNotifications.Caption = "Non-numeric value entered for map number... maps not reserved."
    lblNotifications.ForeColor = &HFF&


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReserveMaps_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveCredits_Click()


   On Error GoTo errorhandler
   
    Credits = txtCredits.Text
    Dim iFileNumber As Integer
    iFileNumber = FreeFile
    Open App.path & "\data\credits.txt" For Output As #iFileNumber
    Print #iFileNumber, Credits
    Close #iFileNumber
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveCredits_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveDataFolder_Click()


   On Error GoTo errorhandler
   
   SaveOptions
    
    If Len(Trim$(txtDataFolder.Text)) > 0 And Trim$(LCase(txtDataFolder.Text)) <> "default" Then
        If IsValidFileName(Trim$(txtDataFolder.Text)) Then
            Options.DataFolder = txtDataFolder.Text
            SaveOptions
            lblNotifications.Caption = "Saved new data folder and update.ini URL!"
        Else
            lblNotifications.Caption = "Data folder not valid! (Saved URL)"
        End If
    Else
        lblNotifications.Caption = "Using 'Default' data folder."
        Options.DataFolder = "default"
        txtDataFolder.Text = "default"
        SaveOptions
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveDataFolder_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Private Function IsValidFileName(strName As String) As Boolean
    IsValidFileName = True
    
    If InStrB(1, strName, "\", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, "/", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, ":", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, "?", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, "*", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, "|", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, Chr(34), vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, "<", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, ">", vbBinaryCompare) Then IsValidFileName = False
End Function

Private Sub cmdSaveHouse_Click()

   On Error GoTo errorhandler

    If Val(txtBaseMap.Text) <= 0 Or Val(txtBaseMap.Text) > MAX_MAPS Then
        lblNotifications.Caption = "Base Map value invalid. Must be a number between 1 and " & MAX_MAPS
        lblNotifications.ForeColor = &HFF&
        Exit Sub
    End If
    If Val(txtHouseFurniture.Text) < 0 Or Val(txtHouseFurniture.Text) > 1000 Then
        lblNotifications.Caption = "Value of max furnitures invalid. Must be a number between 0 (infinite) and 1000"
        lblNotifications.ForeColor = &HFF&
        Exit Sub
    End If
    
    If Val(txtXEntrance.Text) < 0 Or Val(txtXEntrance.Text) > Map(txtBaseMap.Text).MaxX Then
        lblNotifications.Caption = "Value of x coordinate is invalid. Must be a number between 0  and map max x value."
        lblNotifications.ForeColor = &HFF&
        Exit Sub
    End If
    
    If Val(txtYEntrance.Text) < 0 Or Val(txtYEntrance.Text) > Map(txtBaseMap.Text).MaxY Then
        lblNotifications.Caption = "Value of y coordinate is invalid. Must be a number between 0  and map max y value."
        lblNotifications.ForeColor = &HFF&
        Exit Sub
    End If
    
    
    If frmServer.lstHouses.ListIndex > -1 And frmServer.lstHouses.ListIndex < MAX_HOUSES Then
        HouseConfig(frmServer.lstHouses.ListIndex + 1).BaseMap = Val(txtBaseMap.Text)
        HouseConfig(frmServer.lstHouses.ListIndex + 1).ConfigName = txtHouseName.Text
        HouseConfig(frmServer.lstHouses.ListIndex + 1).MaxFurniture = Val(txtHouseFurniture.Text)
        HouseConfig(frmServer.lstHouses.ListIndex + 1).price = Val(txtHousePrice.Text)
        HouseConfig(frmServer.lstHouses.ListIndex + 1).x = Val(txtXEntrance.Text)
        HouseConfig(frmServer.lstHouses.ListIndex + 1).y = Val(txtYEntrance.Text)
        SaveHouse frmServer.lstHouses.ListIndex + 1
        lblNotifications.Caption = "House Saved."
        lblNotifications.ForeColor = &HC000&
    Else
        lblNotifications.Caption = "Error: No house configuration selected in lst config. House not saved."
        lblNotifications.ForeColor = &HFF&
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveHouse_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveNews_Click()


   On Error GoTo errorhandler

    News = txtNews.Text
    Dim iFileNumber As Integer
    iFileNumber = FreeFile
    Open App.path & "\data\news.txt" For Output As #iFileNumber
    Print #iFileNumber, News
    Close #iFileNumber


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveNews_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Command_Click()
    
End Sub

Private Sub Command1_Click()
picMapName.Visible = Not picMapName.Visible
End Sub

Private Sub fraMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long

   On Error GoTo errorhandler

    For i = 0 To 4
        lblServerMenuOpt(i).Font.Underline = False
    Next
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "fraMenu_MouseMove", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblCPSLock_Click()

   On Error GoTo errorhandler

    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblCPSLock_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblServerMenuOpt_Click(Index As Integer)
    

   On Error GoTo errorhandler

    Select Case Index
        Case 0
            ClearServerWindows False
            fraConsole.Visible = True
        Case 1
            ClearServerWindows False
        Case 2
            ClearServerWindows False
            fraHousing.Visible = True
        Case 3
            ClearServerWindows False
            fraControlPanel.Visible = True
            txtNews.Text = News
            txtCredits.Text = Credits
            txtDataFolder.Text = Options.DataFolder
        Case 4
            DestroyServer
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblServerMenuOpt_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblServerMenuOpt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long

   On Error GoTo errorhandler

    For i = 0 To 4
        If Index <> i Then
            lblServerMenuOpt(i).Font.Underline = False
        Else
            lblServerMenuOpt(i).Font.Underline = True
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblServerMenuOpt_MouseMove", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblServerMsg_Click()

End Sub

Private Sub lblUpdateHelp_Click()


   On Error GoTo errorhandler
    
    Call ShellExecute(0, vbNullString, "http://eclipseorigins.com/smf1/index.php?topic=51", vbNullString, vbNullString, vbNormalFocus)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblUpdateHelp_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lstHouses_Click()

   On Error GoTo errorhandler

    If lstHouses.ListIndex > -1 And lstHouses.ListIndex < MAX_HOUSES Then
        txtBaseMap.Text = HouseConfig(lstHouses.ListIndex + 1).BaseMap
        txtHouseName.Text = HouseConfig(lstHouses.ListIndex + 1).ConfigName
        txtHouseFurniture.Text = HouseConfig(lstHouses.ListIndex + 1).MaxFurniture
        txtHousePrice.Text = HouseConfig(lstHouses.ListIndex + 1).price
        txtXEntrance.Text = HouseConfig(lstHouses.ListIndex + 1).x
        txtYEntrance.Text = HouseConfig(lstHouses.ListIndex + 1).y
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstHouses_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub mnuEditPlayer_Click()
    Dim Name As String
    Dim i As Long

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        If Len(Trim$(Name)) <= 0 Then Exit Sub
        i = FindPlayer(Trim$(Name))
        EditingPlayer = i
        For i = 1 To MAX_INV
            EditInv(i) = Player(EditingPlayer).characters(TempPlayer(EditingPlayer).CurChar).Inv(i)
        Next
        For i = 1 To MAX_PLAYER_SPELLS
            EditSpell(i) = Player(EditingPlayer).characters(TempPlayer(EditingPlayer).CurChar).Spell(i)
        Next
        frmEditPlayer.LoadEditPlayer EditingPlayer
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuEditPlayer_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)

   On Error GoTo errorhandler

    Call AcceptConnection(Index, requestID)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Socket_ConnectionRequest", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)

   On Error GoTo errorhandler

    Call AcceptConnection(Index, SocketId)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Socket_Accept", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)


   On Error GoTo errorhandler

    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Socket_Close(Index As Integer)

   On Error GoTo errorhandler

    Call CloseSocket(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Socket_Close", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ********************
Private Sub chkServerLog_Click()

    ' if its not 0, then its true

   On Error GoTo errorhandler

    If Not chkServerLog.Value Then
        ServerLog = True
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkServerLog_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdExit_Click()

   On Error GoTo errorhandler

    Call DestroyServer


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdExit_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadClasses_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadClasses
    Call TextAdd("All classes reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendClasses i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadClasses_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadItems_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadItems
    Call TextAdd("All items reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendItems i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadItems_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadMaps_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadMaps
    Call TextAdd("All maps reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadMaps_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadNPCs_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendNpcs i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadNPCs_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadShops_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadShops
    Call TextAdd("All shops reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendShops i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadShops_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadSpells_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadSpells_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadResources_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadResources
    Call TextAdd("All Resources reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendResources i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadResources_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadAnimations_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadAnimations
    Call TextAdd("All Animations reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendAnimations i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadAnimations_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdShutDown_Click()

   On Error GoTo errorhandler

    If isShuttingDown Then
        isShuttingDown = False
        shutDownType = 0
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
        shutDownType = 0
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdShutDown_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_Load()

   On Error GoTo errorhandler

    lblNotifications.Caption = ""
    lblNotifications.ForeColor = &HFF&


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearServerWindows(Optional showMenu As Boolean = True)
    fraConsole.Visible = False
    fraHousing.Visible = False
    fraControlPanel.Visible = False
    
    If showMenu Then
        fraMenu.Visible = True
    Else
        fraMenu.Visible = False
    End If
End Sub

Private Sub Form_Resize()


   On Error GoTo errorhandler

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Resize", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error GoTo errorhandler

    Cancel = True
    Call DestroyServer


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.

   On Error GoTo errorhandler

    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.Index - 1
    lvwInfo.Sorted = True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lvwInfo_ColumnClick", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
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
                lblNotifications.ForeColor = &HFF&
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
    HandleError "tmrNotifications_Timer", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtText_GotFocus()

   On Error GoTo errorhandler

    txtChat.SetFocus


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtText_GotFocus", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)


   On Error GoTo errorhandler

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtChat_KeyPress", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub UsersOnline_Start()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (i)

        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If

        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UsersOnline_Start", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)


   On Error GoTo errorhandler

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lvwInfo_MouseDown", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server owner!")
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuKickPlayer_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        CloseSocket (FindPlayer(Name))
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuDisconnectPlayer_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        If Ban(FindPlayer(Name), "", True, "Banned by server console.") = False Then
            frmServer.lblNotifications.Caption = Trim$(Name) & " is already banned!"
        Else
            frmServer.lblNotifications.Caption = Trim$(Name) & " and his IP has been banned from this server!"
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuBanPlayer_click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted administrator access.", BrightCyan)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuAdminPlayer_click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have had your administrator access revoked.", BrightRed)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuRemoveAdmin_click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long, i As Long

   On Error GoTo errorhandler

    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select
    
    For i = 0 To 4
        lblServerMenuOpt(i).Font.Underline = False
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
