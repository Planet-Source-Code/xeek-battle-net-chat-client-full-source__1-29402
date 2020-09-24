VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3135
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin MSWinsockLib.Winsock ws 
      Left            =   5520
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab daTab 
      Height          =   3135
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "pbox(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "bAnonymous"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "tUser"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "tMp3Dir"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Picture3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "tIdleMsg"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "tIdleInterval"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "tPass"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "pbox(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "pbox(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "bSaveUP"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "tOwner"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "tmrIdle"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Buddies"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frame"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Chat"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tSend"
      Tab(2).Control(1)=   "lstWho"
      Tab(2).Control(2)=   "tChat"
      Tab(2).Control(3)=   "lblChat"
      Tab(2).ControlCount=   4
      Begin VB.Timer tmrIdle 
         Interval        =   60
         Left            =   2640
         Top             =   2040
      End
      Begin VB.TextBox tOwner 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   2595
         Width           =   1455
      End
      Begin VB.CheckBox bSaveUP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "save"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4710
         TabIndex        =   28
         Top             =   630
         Width           =   735
      End
      Begin VB.PictureBox pbox 
         Height          =   315
         Index           =   2
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   1695
         TabIndex        =   27
         Top             =   2640
         Width           =   1755
         Begin VB.CommandButton cAbout 
            Caption         =   "Credits"
            Height          =   255
            Left            =   840
            TabIndex        =   11
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cHelp 
            Caption         =   "Help"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox pbox 
         Height          =   270
         Index           =   1
         Left            =   4440
         ScaleHeight     =   210
         ScaleWidth      =   255
         TabIndex        =   26
         Top             =   1080
         Width           =   315
         Begin VB.CommandButton cHelpLogin 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox tPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   585
         Width           =   1335
      End
      Begin VB.TextBox tIdleInterval 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Text            =   "60"
         ToolTipText     =   "Specify '0' for no idle"
         Top             =   2205
         Width           =   255
      End
      Begin VB.TextBox tIdleMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   2205
         Width           =   2055
      End
      Begin VB.PictureBox Picture3 
         Height          =   270
         Left            =   3000
         ScaleHeight     =   210
         ScaleWidth      =   375
         TabIndex        =   21
         Top             =   1620
         Width           =   435
         Begin VB.CommandButton cMp3Dir 
            Caption         =   "…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.TextBox tMp3Dir 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1605
         Width           =   2535
      End
      Begin VB.TextBox tUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   585
         Width           =   1335
      End
      Begin VB.Frame frame 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   5460
         Begin VB.PictureBox pbox 
            Height          =   315
            Index           =   3
            Left            =   4530
            ScaleHeight     =   255
            ScaleWidth      =   735
            TabIndex        =   38
            Top             =   2235
            Width           =   795
            Begin VB.CommandButton cRefresh 
               Caption         =   "refresh"
               Height          =   255
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   24
            Left            =   210
            TabIndex        =   59
            Top             =   1875
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   23
            Left            =   1200
            TabIndex        =   58
            Top             =   1875
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   22
            Left            =   2220
            TabIndex        =   57
            Top             =   1875
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   21
            Left            =   3225
            TabIndex        =   56
            Top             =   1875
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   20
            Left            =   4230
            TabIndex        =   55
            Top             =   1875
            Width           =   975
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   19
            Left            =   210
            TabIndex        =   54
            Top             =   1605
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   18
            Left            =   1215
            TabIndex        =   53
            Top             =   1605
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   17
            Left            =   2220
            TabIndex        =   52
            Top             =   1605
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   16
            Left            =   3225
            TabIndex        =   51
            Top             =   1605
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   15
            Left            =   4230
            TabIndex        =   50
            Top             =   1605
            Width           =   975
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   14
            Left            =   210
            TabIndex        =   49
            Top             =   1350
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   13
            Left            =   1215
            TabIndex        =   48
            Top             =   1350
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   12
            Left            =   2220
            TabIndex        =   47
            Top             =   1350
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   11
            Left            =   3225
            TabIndex        =   46
            Top             =   1350
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   10
            Left            =   4230
            TabIndex        =   45
            Top             =   1350
            Width           =   975
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   9
            Left            =   210
            TabIndex        =   44
            Top             =   1095
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   8
            Left            =   1215
            TabIndex        =   43
            Top             =   1095
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   7
            Left            =   2220
            TabIndex        =   42
            Top             =   1095
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   6
            Left            =   3225
            TabIndex        =   41
            Top             =   1095
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   4230
            TabIndex        =   40
            Top             =   1095
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "green - user is online"
            ForeColor       =   &H00008000&
            Height          =   210
            Left            =   360
            TabIndex        =   37
            Top             =   480
            Width           =   1620
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "red - user is offline"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   360
            TabIndex        =   36
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   4230
            TabIndex        =   35
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   3225
            TabIndex        =   34
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   2220
            TabIndex        =   33
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   1215
            TabIndex        =   32
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lblBuddy 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   210
            TabIndex        =   31
            Top             =   840
            Width           =   1020
         End
      End
      Begin VB.CheckBox bAnonymous 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Log in anonymously"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2280
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.PictureBox pbox 
         Height          =   315
         Index           =   0
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   1455
         TabIndex        =   16
         Top             =   1035
         Width           =   1515
         Begin VB.CheckBox bLogin 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cDropSrvrs 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   2
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox tSend 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -74880
         TabIndex        =   15
         Top             =   2715
         Width           =   3975
      End
      Begin VB.ListBox lstWho 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2325
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":0054
         Left            =   -70800
         List            =   "frmMain.frx":0056
         TabIndex        =   14
         Top             =   705
         Width           =   1335
      End
      Begin RichTextLib.RichTextBox tChat 
         Height          =   2235
         Left            =   -74880
         TabIndex        =   13
         Top             =   450
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3942
         _Version        =   393217
         BackColor       =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":0058
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblChat 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -70800
         TabIndex        =   60
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bot owner:"
         Height          =   210
         Index           =   7
         Left            =   600
         TabIndex        =   30
         Top             =   2640
         Width           =   795
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "interval"
         Height          =   210
         Index           =   1
         Left            =   2400
         TabIndex        =   29
         Top             =   1965
         Width           =   525
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seconds"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   2760
         TabIndex        =   25
         ToolTipText     =   "Specify '0' for no idle"
         Top             =   2205
         Width           =   675
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "idle message"
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   24
         Top             =   1965
         Width           =   945
      End
      Begin VB.Label lbl 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   " To use the mp3 feature, you need to start playing mp3 music before you start whichever game you're using."
         ForeColor       =   &H00808080&
         Height          =   1050
         Index           =   5
         Left            =   3645
         TabIndex        =   23
         ToolTipText     =   "click to play a random mp3"
         Top             =   1560
         Width           =   1830
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mp3 directory"
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   1365
         Width           =   990
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "password:"
         Height          =   210
         Index           =   2
         Left            =   2520
         TabIndex        =   19
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "username:"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   765
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   2535
         Left            =   120
         Top             =   480
         Width           =   5490
      End
   End
   Begin VB.Menu mnuSrvrs 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSrv 
         Caption         =   "uswest.battle.net"
         Index           =   0
      End
      Begin VB.Menu mnuSrv 
         Caption         =   "useast.battle.net"
         Index           =   1
      End
      Begin VB.Menu mnuSrv 
         Caption         =   "europe.battle.net"
         Index           =   2
      End
      Begin VB.Menu mnuSrv 
         Caption         =   "asia.battle.net"
         Index           =   3
      End
      Begin VB.Menu mnuSrv 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSrv 
         Caption         =   "custom server"
         Index           =   5
      End
   End
   Begin VB.Menu mnuBuddys 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAddBuddy 
         Caption         =   "Add Buddy"
      End
      Begin VB.Menu mnuRemoveBuddy 
         Caption         =   "Remove Buddy"
      End
      Begin VB.Menu mnuDv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWhisperBuddy 
         Caption         =   "Whisper Buddy"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bLogin_Click()
    'if there's any error go to offline
    On Error GoTo offline
    'check the login caption
    Select Case bLogin.Caption
        'if it says connect
        Case "connect"
            'make it say cancel
            bLogin.Caption = "cancel"
            Dim i As Integer
            'check the server menu, see if anything is checked or not
            For i = 0 To 6
                Debug.Print i
                If i = 6 Then
                    'if nothing is checked, tell the user what to do
                    MsgBox "Please click the down-arrow button next to the [log in] button" & vbCrLf & "to choose a server to connect to"
                    'raise the login button (click it while it says 'cancel')
                    bLogin.Value = 0
                    Exit Sub
                End If
                'if it finds one checked set the login tag with the server to use
                If mnuSrv(i).Checked = True Then _
                    bLogin.Tag = mnuSrv(i).Caption: Exit For
            Next
            'set the caption of the app so the user knows it's trying to connect
            Cap "connecting..."
            'start the connection to whichever server was selected
            ws.Connect bLogin.Tag, 6112
        'if the button says cancel
        Case "cancel"
            'make the caption say 'connect' again
            bLogin.Caption = "connect"
            'if it's not closed by the time the user hits cancel, close it
            If ws.State <> sckClosed Then ws.Close
            Cap
        'if it says disconnect
        Case "disconnect"
            'check the connection, if it's open, close it
            If ws.State <> sckClosed Then ws.Close
            'raise the button right away (so it doesn't stay down)
            bLogin.Value = 0
            Disc
    End Select
    Exit Sub
offline:
    'it errored, so let the user know
    If Err.Number = 10065 Then MsgBox "You are not connected to the Internet"
End Sub

Private Sub cAbout_Click()
    Dim s As String
    'show the credits
    s = vbCrLf & "               Û"
    s = s & vbCrLf & "   Battle.net Û²Û"
    s = s & vbCrLf & "     Mp3 Bot Û²±²Û"
    s = s & vbCrLf & "            Û²±°²Û"
    s = s & vbCrLf & "           Û²±°°±²Û"
    s = s & vbCrLf & "          Û²±°v2±²Û"
    s = s & vbCrLf & "           Û²±°±²Û"
    s = s & vbCrLf & "            Û²±²Û"
    s = s & vbCrLf & "             Û²Û"
    s = s & vbCrLf & "             ôõ"
    s = s & vbCrLf & "            _|_"
    s = s & vbCrLf & "        (ßßß   ßßß)"
    s = s & vbCrLf & "       C_\_______/"
    s = s & vbCrLf
    s = s & vbCrLf & " |\_\_\_ Credits _/_/_/|"
    s = s & vbCrLf
    s = s & vbCrLf & " programmer ------ xeek"
    s = s & vbCrLf & " tester/editor sandrock"
    s = s & vbCrLf & " tester/editor - xphyle"
    s = s & vbCrLf & " email xeek@hotmail.com"
    s = s & vbCrLf & " copyright ------ ©2001"
    Display "About", s
End Sub

Private Sub cDropSrvrs_Click()
    'pop up the menu just below the picture box that holds the login button
    PopupMenu mnuSrvrs, , pbox(0).Left, pbox(0).Top + pbox(0).Height
End Sub

Private Sub cHelp_Click()
    Dim s As String
    'show the help
    s = vbCrLf & "        {Help}"
    s = s & vbCrLf
    s = s & vbCrLf & " logging on ----------"
    s = s & vbCrLf & "- please click the question mark button on the [General] tab"
    s = s & vbCrLf
    s = s & vbCrLf & " buddies -------------"
    s = s & vbCrLf & "- battle.net only allows you to have a maximum of 25 buddies, so that's why there's only 25 boxes"
    s = s & vbCrLf & "- right-click or double-click the boxes on the [buddies] tab for options"
    s = s & vbCrLf & "- since it's really no use to take up bandwidth constantly checking your buddies, we put a button [refresh] to reveal your buddies. As the labels say... buddies that appear in red are offline, and the ones that appear in green are online."
    s = s & vbCrLf & "- some servers other than battle.net might not have the buddy feature"
    s = s & vbCrLf
    s = s & vbCrLf & " chat ----------------"
    s = s & vbCrLf & "- if you type '/help' in the chatroom it will show you what commands there are on the server-side."
    s = s & vbCrLf & "- the chat will also show you any information coming from the server other than what people are saying"
    s = s & vbCrLf & "- if you want to retrieve a user's stats, the server will tell you to use a valid program id. The only games that give a user stats are StarCraft and WarCraft, but just for you to know, here are the valid id's:"
    s = s & vbCrLf & " SEXP - broodwar"
    s = s & vbCrLf & " STAR - starcraft"
    s = s & vbCrLf & " DRTL - diablo"
    s = s & vbCrLf & " D2DV - diablo 2"
    s = s & vbCrLf & " D2XP - diablo 2 exp."
    s = s & vbCrLf & " W2BN - warcraft 2"
    s = s & vbCrLf & " CHAT - chat bot"
    s = s & vbCrLf
    s = s & vbCrLf & " idle ----------------"
    s = s & vbCrLf & "- the idle is used to send a message to the chatroom every so many seconds. You can use this to tell people to go to a different private channel or whatever."
    s = s & vbCrLf & "- setting the interval to 0 turns off the idle"
    s = s & vbCrLf & "- the textbox where you type the amount of seconds for the interval must lose focus before the interval is set."
    s = s & vbCrLf
    s = s & vbCrLf & " bot owner -----------"
    s = s & vbCrLf & "- this is important so that the bot knows who will be whispering commands to it. If you leave it blank, the bot will not accept commands from anyone."
    s = s & vbCrLf & "- the bot owner is the username battle.net gives you after you log on from within the game (not this bot). sometimes battle.net will give you a different username, usually it just appends a number to your username you used"
    s = s & vbCrLf
    s = s & vbCrLf & " mp3 options ---------"
    s = s & vbCrLf & "- if you left-click the [...] button, you'll be able to choose the directory (folder) that holds your mp3."
    s = s & vbCrLf & "- if you right-click the [...] button you'll be presented with mp3 options (play, stop, etc.)"
    s = s & vbCrLf & "- if you whisper '/help' to your bot, it will reply with the available commands."
    s = s & vbCrLf & "- as it says, the sound card must be in use by mp3 before you begin playing a game."
    s = s & vbCrLf & "- you won't be able to hear the game's sounds (unless your soundcard is phat like that) while you're listening to mp3."
    s = s & vbCrLf & "- you can't start playing mp3 while the game is running (unless your sound card is pimp)."
    s = s & vbCrLf & "- if you wish to hear the game's sounds again, you must completely stop playing mp3 and exit & restart the game."
    Display "Help", s
End Sub

Private Sub cHelpLogin_Click()
    Dim s As String
    'show connecting info
    s = vbCrLf & "      {Connecting}"
    s = s & vbCrLf
    s = s & vbCrLf & " accounts ------------"
    s = s & vbCrLf & "- accounts can only be made through a user of a game. The account is bound by the terms of service of whichever server you're using."
    s = s & vbCrLf & "- for game info go to blizzard.com"
    s = s & vbCrLf & " anonymous logins ----"
    s = s & vbCrLf & "- logging in anonymously only allows you to issue server queries. query commands can be retrieved by going to the [chat] tab and typing /help. for more chat info, click the [help] button on the 'general' tab"
    Display "Login Help", s
End Sub

Private Sub cMp3Dir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbRightButton
            'popupmenu mp3ops
        Case vbLeftButton
            Dim s As String, i As Integer
            s = getFolder(frmMain.hWnd)
            If s <> "" Then tMp3Dir.Text = s
            i = countMp3(True)
            MsgBox i & " mp3 were found"
    End Select
End Sub

Private Sub cRefresh_Click()
    On Error GoTo offline
    Dim i As Integer
    'for every label there is
    For i = 0 To 24
        'clear them all
        lblBuddy(i).Caption = ""
    Next
    'reset our buddy count
    buddy = 0
    'send the server query to retrieve our buddies
    ws.SendData "/friends list" & vbCrLf
    Exit Sub
offline:
    'if disconnected, let the user know
    If Err.Number = 40006 Then MsgBox "You were disconnected."
    Disc
End Sub

Private Sub Form_Load()
    'set the default state
    Disc
    Set mp3 = New clsMP3
    Set list = New mp3List
End Sub
Private Sub lblBuddy_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if the user isn't connected, tell them
    If ws.State <> sckConnected Then
        MsgBox "You are not connected"
        Disc
    End If
    'if they right-clicked
    If Button = vbRightButton Then
        'if there's no buddy in the caption
        If lblBuddy(Index).Caption = "" Then
            'only make the 'add buddy' menu visible
            mnuRemoveBuddy.Visible = False
            mnuAddBuddy.Visible = True
            mnuDv1.Visible = False
            mnuWhisperBuddy.Visible = False
        'if there's a buddy there
        Else
            'only make the 'add buddy' menu not visible
            mnuRemoveBuddy.Visible = True
            mnuAddBuddy.Visible = False
            mnuDv1.Visible = True
            mnuWhisperBuddy.Visible = True
        End If
        mnuBuddys.Tag = Index
        'pop up the menu
        PopupMenu mnuBuddys
    End If
End Sub

Private Sub mnuAddBuddy_Click()
    On Error GoTo offline
    Dim s As String
    'get the username of the buddy they wish to add
    s = InputBox("Enter the username of your buddy" & vbCrLf & "(remember, usernames don't have spaces)", "Add Buddy")
    'remove any spaces
    s = Replace(s, " ", "")
    'if the user didn't input anything, don't do anything
    If s = "" Then Exit Sub
    'send the command to add the buddy
    ws.SendData "/f a " & s & vbCrLf
    Exit Sub
offline:
    'let the user know they got disconnected
    If Err.Number = 40006 Then MsgBox "You were disconnected"
    Disc
End Sub

Private Sub mnuRemoveBuddy_Click()
    On Error GoTo offline
    'remove the buddy
    ws.SendData "/f r " & lblBuddy(mnuBuddys.Tag) & vbCrLf
    Exit Sub
offline:
    'tell the user they got disconnected
    If Err.Number = 40006 Then MsgBox "You were disconnected"
    Disc
End Sub

Private Sub mnuSrv_Click(Index As Integer)
    'if the 5th one was clicked (the custom server one)
    If Index = 5 Then
        Dim s As String
        'ask the user for the server to use, using the menu's caption as the default
        s = InputBox(vbCrLf & vbCrLf & "Enter in the custom server location.", , mnuSrv(5).Caption)
        'if nothing was entered, don't do nothing
        If Trim(s) = "" Then Exit Sub
        'set the caption of the menu to what the user entered
        mnuSrv(5).Caption = s
    End If
    Dim i As Integer
    'uncheck all the server menus
    For i = 0 To 5
        mnuSrv(i).Checked = False
    Next
    'check the one that was clicked
    mnuSrv(Index).Checked = True
    'if the user selects a non-custom server, reset the caption of the 5th menu
    If Index < 5 Then mnuSrv(5).Caption = "custom server"
End Sub

Private Sub mnuWhisperBuddy_Click()
    'switch to the chat tab
    daTab.Tab = 2
    'enter in the server query for them
    tSend.Text = "/w " & lblBuddy(mnuBuddys.Tag).Caption & " "
    'set the cursor at the end, so they can type their message right away
    tSend.SelStart = Len(tSend.Text)
End Sub

Private Sub tIdleInterval_KeyPress(KeyAscii As Integer)
    'if the user hits delete, don't do nothing
    If KeyAscii = 8 Then Exit Sub
    'if the user tries typing anything that's not numeric, don't let nothing happen
    If Not IsNumeric(Chr(CLng(KeyAscii))) Then KeyAscii = 0
End Sub

Private Sub tIdleInterval_LostFocus()
    'when the idle interval loses focus,
    'make sure the interval is between 0 and 60 seconds
    With tIdleInterval
        If .Text = "" Then .Text = 0
        If .Text > 60 Then .Text = 60
        If .Text < 0 Then .Text = 0
    End With
    'set the timer's interval
    tmrIdle.Interval = tIdleInterval.Text * 1000
End Sub

Private Sub tmrIdle_Timer()
    On Error GoTo offline
    'if there's no text, don't do anything
    If tIdleMsg.Text = "" Then Exit Sub
    'send the idle message
    ws.SendData tIdleMsg.Text & vbCrLf
    Exit Sub
offline:
    'if there's no connection, don't do anything
    If Err.Number = 40006 Then Exit Sub
End Sub

Private Sub tOwner_KeyPress(KeyAscii As Integer)
    'usernames don't have spaces, so if the user hits space, don't do anything
    If KeyAscii = 32 Then _
        KeyAscii = 0: Exit Sub
End Sub

Private Sub tSend_KeyPress(KeyAscii As Integer)
    On Error GoTo offline
    'check which key is being pushed
    Select Case KeyAscii
        'if it's the enter/return key
        Case 13
            'make sure our app doesn't beep
            KeyAscii = 0
            'send our data to the server
            ws.SendData tSend.Text & vbCrLf
            'if we're not issuing a server query, the server won't send anything back to us
            'so that means we make sure our chat gets displayed also
            If Left(tSend.Text, 1) <> "/" Then
                'set the color to a light blue
                tChat.SelColor = RGB(51, 153, 255)
                'using our username, display who's sending what
                tChat.SelText = vbCrLf & "<" & un & "> "
                'set the color to white
                tChat.SelColor = RGB(255, 255, 255)
                'show what we're chatting
                tChat.SelText = tSend.Text
                'scroll to the end of the chat area, so we can 'follow' the chat
                tChat.SelStart = Len(tChat.Text)
            End If
            'set our textbox to nothing so we can type more right away
            tSend.Text = ""
        'if the user hit ctrl + a
        Case 1
            'do a few things to select the text in the textbox
            KeyAscii = 0
            tSend.SelStart = 0
            tSend.SelLength = Len(tSend.Text)
    End Select
    Exit Sub
offline:
    'if we were offline, let the user know
    If Err.Number = 40006 Then MsgBox "You were disconnected"
    Disc
End Sub

Private Sub ws_Close()
    'if there was a remote closure, make sure it's completely closed
    If ws.State <> sckClosed Then ws.Close
    'let the person know they were disconnected
    Cap "remotely disconnected"
    Disc
End Sub

Private Sub ws_Connect()
    Conn
    'check if we're logging in anonymously or not
    If bAnonymous.Value = 0 Then
        'if we're not, use the username/password supplied by the user
        ws.SendData Chr(3) & Chr(4) & tUser.Text & Chr(13) & Chr(10) & tPass.Text & Chr(13) & Chr(10)
    Else
        'otherwise use "anonymous" as the username and password
        ws.SendData Chr(3) & Chr(4) & "anonymous" & Chr(13) & Chr(10) & "anonymous" & Chr(13) & Chr(10)
    End If
End Sub

Private Sub parse(what As String)
    Dim l As Long
    On Error GoTo nod
    'check the first 9 letters of what the server is sending us
    Select Case Left(what, 9)
        
        'the server occasionally sends this, to check for dead connections, we can just ignore it
        Case "2000 NULL"
            'nothing
        
        'the server sends this to tell us the exact username it's identifying us with
        Case "2010 NAME"
            un = Mid(what, 11)
            Cap "connected ( " & un & " )"
        
        'this is the same as join
        'this is sent when you join a channel, and there's already users there
        Case "1001 USER"
            'find the position of the first space and two zeros
            l = InStr(what, " 00")
            'grab who the user is from the middle of the string
            what = Mid(what, 11, l - 1 - 10)
            'for every user in our list of who's in the channel
            For l = 0 To (lstWho.ListCount - 1)
                'check if the user is already in the list (servers can be weird sometimes)
                If lstWho.list(l) = what Then Exit Sub
            Next
            'add the user to the list
            lstWho.AddItem what
            'set the caption with the channel name and the amount of users in it
            lblChat.Caption = chan & " (" & lstWho.ListCount & ")"
        
        'this tells you when a user has entered the channel you're in
        Case "1002 JOIN"
            l = InStr(what, " 00")
            what = Mid(what, 11, l - 1 - 10)
            For l = 0 To (lstWho.ListCount - 1)
                If lstWho.list(l) = what Then Exit Sub
            Next
            lstWho.AddItem what
            lblChat.Caption = chan & " (" & lstWho.ListCount & ")"
        
        'this tells you when a user has left the room
        Case "1003 LEAV"
            l = InStr(what, " 00")
            what = Mid(what, 12, l - 1 - 11)
            For l = 0 To (lstWho.ListCount - 1)
                If lstWho.list(l) = what Then lstWho.RemoveItem l
                lblChat.Caption = chan & " (" & lstWho.ListCount & ")"
            Next
            
        'this tells you when someone is whispering to you
        Case "1004 WHIS"
            'find out who's whispering you
            'and using it's own set of colors display the whisper
            what = Mid(what, 1, Len(what) - 1)
            l = InStr(what, " 00")
            'if the person whispering our bot, is the owner
            If Mid(what, 14, l - 1 - 13) = tOwner.Text Then
                'find out what the owner said and process it
                Owner Mid(what, l + 7)
            End If
            tChat.SelColor = RGB(255, 255, 0)
            tChat.SelText = vbCrLf & "<From: " & Mid(what, 14, l - 1 - 13) & "> "
            tChat.SelColor = RGB(192, 192, 192)
            tChat.SelText = Mid(what, l + 7)
            
        'this lets you know when someone else is talking in the chat
        Case "1005 TALK"
            'find out who it is and display what they're saying using it's own set of colors
            what = Mid(what, 1, Len(what) - 1)
            l = InStr(what, " 00")
            tChat.SelColor = RGB(255, 255, 0)
            tChat.SelText = vbCrLf & "<" & Mid(what, 11, l - 1 - 10) & "> "
            tChat.SelColor = RGB(255, 255, 255)
            tChat.SelText = Mid(what, l + 7)
            
        'this is when the server tells you what channel you've joined
        Case "1007 CHAN"
            'clear our list of people
            lstWho.Clear
            'find out the channel name, set the globals, and display it as green text in the chat
            what = Mid(what, 1, Len(what) - 1)
            l = InStr(what, Chr(34))
            chan = Mid(what, l + 1)
            tChat.SelColor = RGB(0, 255, 0)
            tChat.SelText = vbCrLf & "Joining Channel: " & chan
            lblChat.Caption = chan & " (" & lstWho.ListCount & ")"
        
        'when you whisper someone,
        'the server will either tell you the user is not online (1019 ERROR),
        'or let you know the whisper went through
        Case "1010 WHIS"
            'find out who you whispered, and your message, and display it in the chat
            what = Mid(what, 1, Len(what) - 1)
            l = InStr(what, " 00")
            tChat.SelColor = RGB(51, 153, 255)
            tChat.SelText = vbCrLf & "<To: " & Mid(what, 14, l - 14) & "> "
            tChat.SelColor = RGB(192, 192, 192)
            tChat.SelText = Mid(what, l + 7)
        
        'the server is sending you information
        Case "1018 INFO"
            'find out the exact message it sent us
            what = Mid(what, 1, Len(what) - 1)
            l = InStr(what, Chr(34))
            what = Mid(what, l + 1)
            'when we issue a friend list query, the server sends data back in the *: *, * format,
            'and if it's sending us a reply...
            If what Like "*: *, *" Then
                'so if we're receiving that kind of info, find out if the user is offline
                'if the user is offline the server sends the last part as ", offline"
                If Right(what, 9) = ", offline" Then
                    'if the user is offline, set the current buddy label to red
                    lblBuddy(buddy).ForeColor = RGB(128, 0, 0)
                'when the user is online, what the server sends is always different
                Else
                    'if the user is online, set the current buddy label to green
                    lblBuddy(buddy).ForeColor = RGB(0, 128, 0)
                End If
                'find out who the user is, and display it in the caption
                lblBuddy(buddy).Caption = Mid(what, InStr(what, ":") + 2, InStrRev(what, ", ") - InStr(what, ":") - 2)
                'if the user's name is longer than what's displayed, our user can hold the cursor over the labels and see the full username
                lblBuddy(buddy).ToolTipText = lblBuddy(buddy).Caption
                'increment our buddy variable, so we know where to stick our next buddy
                buddy = buddy + 1
                If buddy > 24 Then buddy = 0
            End If
            'let the user know using yellow, what the server is telling us
            tChat.SelColor = RGB(255, 255, 0)
            tChat.SelText = vbCrLf & what
            
        'when you issue queries, this is what you get when you issue a bad command or whatever
        Case "1019 ERRO"
            'find out the servers message and display it in red
            what = Mid(what, 1, Len(what) - 1)
            l = InStr(what, Chr(34))
            tChat.SelColor = RGB(255, 0, 0)
            tChat.SelText = vbCrLf & Mid(what, l + 1)
        
        'when a user types an EMOTE, it's like the user is talking in the third person, or the server is talking about the user
        Case "1023 EMOT"
            'find out who it's about, and what the message is,
            'and display it in the chat
            what = Mid(what, 1, Len(what) - 1)
            l = InStr(what, " 00")
            tChat.SelColor = RGB(255, 255, 0)
            tChat.SelText = vbCrLf & "<" & Mid(what, 12, l - 1 - 11) & " " & Mid(what, l + 7) & ">"
        
        'if the server sends anything else it's ignored, but displayed in the immediate window so we can see what's going on (in development)
        Case Else
            Debug.Print what
    End Select
nod:
    'scroll the chat so we can watch the chat 'flow'
    tChat.SelStart = Len(tChat.Text)
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim b As String, l As Long
    'get our data
    ws.GetData b, vbString
    'destroy the last vbcrlf (it's 2 characters)
    b = Mid(b, 1, Len(b) - 2)
    'we're going to keep looping until there are no linebreaks left in our string
    'so we use DoEvents to keep our app from freezing up
    Do: DoEvents
        'find the first linebreak (crlf)
        l = InStr(1, b, vbCrLf)
        'if there's a line break
        If l <> 0 Then
            'parse everything before the line break
            parse Mid(b, 1, l - 1)
            'set the string to everything after the first line break
            b = Mid(b, l + 2)
        'if there's no line break
        Else
            'just parse the string itself
            parse b
            'since the loop isn't waiting for anything specific we exit it
            Exit Do
        End If
    Loop
End Sub

