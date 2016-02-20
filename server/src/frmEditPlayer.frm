VERSION 5.00
Begin VB.Form frmEditPlayer 
   Caption         =   "Editar jogador"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraEditPlayer 
      Caption         =   "Editar Jogador"
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9000
      Begin VB.TextBox txtFace 
         Height          =   285
         Left            =   3840
         TabIndex        =   66
         Text            =   "0"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.ComboBox cmbAccess 
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":0000
         Left            =   1200
         List            =   "frmEditPlayer.frx":0013
         TabIndex        =   39
         Text            =   "cmbAccess"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.ComboBox cmbDir 
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":004F
         Left            =   6480
         List            =   "frmEditPlayer.frx":005F
         TabIndex        =   38
         Text            =   "cmbDir"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelEditPlayer 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   7440
         TabIndex        =   37
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditPlayerOk 
         Caption         =   "Save and Close"
         Height          =   255
         Left            =   5400
         TabIndex        =   36
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   6480
         TabIndex        =   35
         Text            =   "0"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   6480
         TabIndex        =   34
         Text            =   "0"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtMap 
         Height          =   285
         Left            =   6480
         TabIndex        =   33
         Text            =   "0"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Spell Editing"
         Height          =   2055
         Left            =   7320
         TabIndex        =   28
         Top             =   1800
         Width           =   1455
         Begin VB.ComboBox cmbSpellSlot 
            Height          =   315
            ItemData        =   "frmEditPlayer.frx":007A
            Left            =   120
            List            =   "frmEditPlayer.frx":007C
            TabIndex        =   30
            Text            =   "cmbSpellSlot"
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cmbSpells 
            Height          =   315
            ItemData        =   "frmEditPlayer.frx":007E
            Left            =   120
            List            =   "frmEditPlayer.frx":0080
            TabIndex        =   29
            Text            =   "cmbSpells"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Slot:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Spell:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Inventory Editing"
         Height          =   2055
         Left            =   5400
         TabIndex        =   21
         Top             =   1800
         Width           =   1695
         Begin VB.TextBox txtItemQuantity 
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Text            =   "0"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.ComboBox cmbItems 
            Height          =   315
            ItemData        =   "frmEditPlayer.frx":0082
            Left            =   120
            List            =   "frmEditPlayer.frx":0084
            TabIndex        =   23
            Text            =   "cmbItems"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox cmbInvSlot 
            Height          =   315
            ItemData        =   "frmEditPlayer.frx":0086
            Left            =   120
            List            =   "frmEditPlayer.frx":0088
            TabIndex        =   22
            Text            =   "cmbInvSlot"
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Item:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Slot:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbShield 
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":008A
         Left            =   3840
         List            =   "frmEditPlayer.frx":008C
         TabIndex        =   20
         Text            =   "cmbShield"
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox cmbHelmet 
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":008E
         Left            =   3840
         List            =   "frmEditPlayer.frx":0090
         TabIndex        =   19
         Text            =   "cmbHelmet"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.ComboBox cmbArmor 
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":0092
         Left            =   3840
         List            =   "frmEditPlayer.frx":0094
         TabIndex        =   18
         Text            =   "cmbArmor"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.ComboBox cmbWeapon 
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":0096
         Left            =   3840
         List            =   "frmEditPlayer.frx":0098
         TabIndex        =   17
         Text            =   "cmbWeapon"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtPoints 
         Height          =   285
         Left            =   3840
         TabIndex        =   16
         Text            =   "0"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtWillPower 
         Height          =   285
         Left            =   3840
         TabIndex        =   15
         Text            =   "0"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtAgility 
         Height          =   285
         Left            =   3840
         TabIndex        =   14
         Text            =   "0"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtIntelligence 
         Height          =   285
         Left            =   3840
         TabIndex        =   13
         Text            =   "0"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtEndurance 
         Height          =   285
         Left            =   3840
         TabIndex        =   12
         Text            =   "0"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtStrength 
         Height          =   285
         Left            =   3840
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtMP 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   "0"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "0"
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox cmbPK 
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":009A
         Left            =   1200
         List            =   "frmEditPlayer.frx":00A4
         TabIndex        =   8
         Text            =   "cmbPK"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtExp 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "0"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":00B1
         Left            =   1200
         List            =   "frmEditPlayer.frx":00B8
         TabIndex        =   5
         Text            =   "cmbClass"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cmbSex 
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":00C5
         Left            =   1200
         List            =   "frmEditPlayer.frx":00CF
         TabIndex        =   4
         Text            =   "cmbSex"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtCharName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   3
         Text            =   "Character Name"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtPassword 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "Password"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtLogin 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "Login Name"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Face:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   67
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Dir:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   30
         Left            =   5400
         TabIndex        =   64
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   29
         Left            =   5400
         TabIndex        =   63
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   28
         Left            =   5400
         TabIndex        =   62
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Map:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   25
         Left            =   5400
         TabIndex        =   61
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Shield:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   21
         Left            =   2760
         TabIndex        =   60
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Helmet:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   20
         Left            =   2760
         TabIndex        =   59
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Armor:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   19
         Left            =   2760
         TabIndex        =   58
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Weapon:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   2760
         TabIndex        =   57
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Stat Points:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   2760
         TabIndex        =   56
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Willpower:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   2760
         TabIndex        =   55
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Agility:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   2760
         TabIndex        =   54
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Intelligence:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   2760
         TabIndex        =   53
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Endurance:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   2760
         TabIndex        =   52
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Strength:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   2760
         TabIndex        =   51
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "MP:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   50
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   49
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "PKer:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   48
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   47
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   45
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   44
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Char Name:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
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
      Left            =   -480
      TabIndex        =   65
      Top             =   4680
      Width           =   11655
   End
End
Attribute VB_Name = "frmEditPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEditPlayerOk_Click()
Dim i As Long

On Error GoTo errorhandler

    If IsPlaying(EditingPlayer) Then
        'Check Everything First.
        If Val(txtLevel.Text) > MAX_LEVELS Then
            lblNotifications.Caption = "Player Saving Failed: Level is greater than " & MAX_LEVELS
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        If Val(txtStrength.Text) > 255 Then
            lblNotifications.Caption = "Player Saving Failed: Strength is greater than 255"
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        If Val(txtEndurance.Text) > 255 Then
            lblNotifications.Caption = "Player Saving Failed: Endurance is greater than 255"
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        If Val(txtIntelligence.Text) > 255 Then
            lblNotifications.Caption = "Player Saving Failed: Intelligence is greater than 255"
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        If Val(txtAgility.Text) > 255 Then
            lblNotifications.Caption = "Player Saving Failed: Agility is greater than 255"
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        If Val(txtWillPower.Text) > 255 Then
            lblNotifications.Caption = "Player Saving Failed: Willpower is greater than 255"
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        Player(EditingPlayer).Password = Trim$(txtPassword.Text)
        With Player(EditingPlayer).characters(TempPlayer(EditingPlayer).CurChar)
            .Name = Trim$(txtCharName.Text)
            .Sex = cmbSex.ListIndex
            .Class = cmbClass.ListIndex + 1
            .Face(2) = Val(txtFace.Text)
            .Level = Val(txtLevel.Text)
            .Exp = Val(txtExp.Text)
            .access = cmbAccess.ListIndex
            .PK = cmbPK.ListIndex
            .Vital(Vitals.HP) = Val(txtHP.Text)
            .Vital(Vitals.MP) = Val(txtMP.Text)
            .stat(Stats.Strength) = Val(txtStrength.Text)
            .stat(Stats.Endurance) = Val(txtEndurance.Text)
            .stat(Stats.Intelligence) = Val(txtIntelligence.Text)
            .stat(Stats.Agility) = Val(txtAgility.Text)
            .stat(Stats.Willpower) = Val(txtWillPower.Text)
            .Points = Val(txtPoints.Text)
            .Equipment(Equipment.Weapon) = cmbWeapon.ListIndex
            .Equipment(Equipment.armor) = cmbArmor.ListIndex
            .Equipment(Equipment.Helmet) = cmbHelmet.ListIndex
            .Equipment(Equipment.Shield) = cmbShield.ListIndex
            .Map = Val(txtMap.Text)
            .x = Val(txtX.Text)
            .y = Val(txtY.Text)
            .Dir = cmbDir.ListIndex
            
            For i = 1 To MAX_INV
                Player(EditingPlayer).characters(TempPlayer(EditingPlayer).CurChar).Inv(i) = EditInv(i)
            Next
            For i = 1 To MAX_PLAYER_SPELLS
                Player(EditingPlayer).characters(TempPlayer(EditingPlayer).CurChar).Spell(i) = EditSpell(i)
            Next
            SavePlayer EditingPlayer
            ' send vitals, exp + stats
            For i = 1 To Vitals.Vital_Count - 1
                Call SendVital(EditingPlayer, i)
            Next
            SendEXP EditingPlayer
            Call SendStats(EditingPlayer)
            Call SendInventory(EditingPlayer)
            SendDataToMap GetPlayerMap(EditingPlayer), PlayerData(EditingPlayer)
        End With
        PlayerWarp EditingPlayer, GetPlayerMap(EditingPlayer), GetPlayerX(EditingPlayer), GetPlayerY(EditingPlayer), False
    Else
        EditingPlayer = 0
        fraEditPlayer.Visible = False
        lblNotifications.Caption = "Player Saving Failed: Player not Found Online"
        lblNotifications.ForeColor = &HFF&
    End If
    EditingPlayer = 0
    fraEditPlayer.Visible = False
    lblNotifications.Caption = "Player Saved Sucessfully!"
    lblNotifications.ForeColor = &HC000&
    Exit Sub
    
errorhandler:
    lblNotifications.Caption = "Player Saving Failed: Unknown Error!"
    lblNotifications.ForeColor = &HFF&
    HandleError "cmdEditPlayerOk_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub LoadEditPlayer(Index As Long)
Dim i As Long

   On Error GoTo errorhandler

    If IsPlaying(Index) Then
        frmEditPlayer.Visible = True
        
        'Load all of the players info :D
        With Player(Index)
            txtLogin.Text = Trim$(.login)
            txtPassword.Text = Trim$(.Password)
        End With
        
        With Player(Index).characters(TempPlayer(Index).CurChar)
            txtCharName.Text = Trim$(.Name)
            cmbSex.ListIndex = .Sex
            MsgBox .Class
            cmbClass.ListIndex = .Class
            txtFace.Text = .Face(2)
            txtLevel.Text = Val(.Level)
            txtExp.Text = Val(.Exp)
            cmbAccess.ListIndex = .access
            cmbPK.ListIndex = .PK
            txtHP.Text = Val(.Vital(Vitals.HP))
            txtMP.Text = Val(.Vital(Vitals.MP))
            txtStrength.Text = Val(.stat(Stats.Strength))
            txtEndurance.Text = Val(.stat(Stats.Endurance))
            txtIntelligence.Text = Val(.stat(Stats.Intelligence))
            txtAgility.Text = Val(.stat(Stats.Agility))
            txtWillPower.Text = Val(.stat(Stats.Willpower))
            txtPoints.Text = Val(.Points)
            cmbWeapon.ListIndex = .Equipment(Equipment.Weapon)
            cmbArmor.ListIndex = .Equipment(Equipment.armor)
            cmbHelmet.ListIndex = .Equipment(Equipment.Helmet)
            cmbShield.ListIndex = .Equipment(Equipment.Shield)
            txtMap.Text = Val(.Map)
            txtX.Text = Val(.x)
            txtY.Text = Val(.y)
            cmbDir.ListIndex = .Dir
            
            cmbInvSlot.Clear
            For i = 1 To MAX_INV
                cmbInvSlot.AddItem i
            Next
            cmbInvSlot.ListIndex = 0
            
            cmbSpellSlot.Clear
            For i = 1 To MAX_PLAYER_SPELLS
                cmbSpellSlot.AddItem i
            Next
            cmbSpellSlot.ListIndex = 0
        
        End With
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadEditPlayer", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbInvSlot_Click()

   On Error GoTo errorhandler

    cmbItems.ListIndex = EditInv(cmbInvSlot.ListIndex + 1).Num
    txtItemQuantity.Text = Val(EditInv(cmbInvSlot.ListIndex + 1).value)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbInvSlot_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbItems_Click()

   On Error GoTo errorhandler

    EditInv(cmbInvSlot.ListIndex + 1).Num = cmbItems.ListIndex


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbItems_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSpells_Change()

   On Error GoTo errorhandler

    EditSpell(cmbSpellSlot.ListIndex + 1) = cmbSpells.ListIndex


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSpells_Change", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSpellSlot_Click()

   On Error GoTo errorhandler

    cmbSpells.ListIndex = EditSpell(cmbSpellSlot.ListIndex + 1)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSpellSlot_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCancelEditPlayer_Click()

   On Error GoTo errorhandler

    EditingPlayer = 0
    fraEditPlayer.Visible = False
    lblNotifications.Caption = "Player Editing Canceled!"
    lblNotifications.ForeColor = &HFF&


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancelEditPlayer_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub


Private Sub txtItemQuantity_Change()

   On Error GoTo errorhandler

    EditInv(cmbInvSlot.ListIndex + 1).value = Val(txtItemQuantity.Text)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtItemQuantity_Change", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
