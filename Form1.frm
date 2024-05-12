VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"Form1.frx":0000
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13650
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6480
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Command25"
      Height          =   615
      Left            =   240
      TabIndex        =   92
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Command25"
      Height          =   615
      Left            =   12480
      TabIndex        =   91
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Command25"
      Height          =   615
      Left            =   11400
      TabIndex        =   90
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Command25"
      Height          =   615
      Left            =   10320
      TabIndex        =   89
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Command25"
      Height          =   615
      Left            =   9240
      TabIndex        =   88
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Command25"
      Height          =   615
      Left            =   8160
      TabIndex        =   87
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Command25"
      Height          =   615
      Left            =   7080
      TabIndex        =   86
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Command25"
      Height          =   615
      Left            =   5520
      TabIndex        =   85
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Command25"
      Height          =   615
      Left            =   4440
      TabIndex        =   84
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Command25"
      Height          =   615
      Left            =   3360
      TabIndex        =   83
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Command25"
      Height          =   615
      Left            =   2280
      TabIndex        =   82
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Command25"
      Height          =   615
      Left            =   1200
      TabIndex        =   81
      Top             =   6120
      Width           =   855
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   9720
      TabIndex        =   73
      Text            =   "PILIH FILE"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H0000FFFF&
      Caption         =   "SCEMATIC/BOARDVIEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9600
      TabIndex        =   72
      Top             =   4320
      Width           =   3855
      Begin VB.DirListBox Dir2 
         Height          =   315
         Left            =   120
         TabIndex        =   80
         Top             =   1200
         Width           =   3375
      End
      Begin VB.CommandButton Command24 
         BackColor       =   &H0080FF80&
         Caption         =   "DOWNLOAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H0080C0FF&
         Caption         =   "CARI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H000080FF&
         Height          =   285
         Left            =   120
         TabIndex        =   76
         Text            =   "TULIS KODE BOARD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H0000FFFF&
      Caption         =   "BIOS LAPTOP/PC "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9600
      TabIndex        =   68
      Top             =   2520
      Width           =   3855
      Begin VB.DirListBox Dir1 
         Height          =   315
         Left            =   120
         TabIndex        =   79
         Top             =   1440
         Width           =   3375
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   120
         TabIndex        =   75
         Text            =   "PILIH SERVER"
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   120
         TabIndex        =   74
         Text            =   "PILIH FILE"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H0080FF80&
         Caption         =   "DOWNLOAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H0080C0FF&
         Caption         =   "CARI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H000080FF&
         Height          =   285
         Left            =   120
         TabIndex        =   69
         Text            =   "TULIS KODE BOARD"
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H0000FFFF&
      Caption         =   "RESETER EPSON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11640
      TabIndex        =   63
      Top             =   1440
      Width           =   1815
      Begin VB.CommandButton Command20 
         BackColor       =   &H00C0E0FF&
         Caption         =   "OPEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H0000FFFF&
      Caption         =   "RESETER CANON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9600
      TabIndex        =   62
      Top             =   1440
      Width           =   1815
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "OPEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00808080&
      Caption         =   "DIRECX 12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00808080&
      Caption         =   "DIRECX 9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   1920
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2013"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   58
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00808080&
      Caption         =   "DIRECX 10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00808080&
      Caption         =   "DIRECX 11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFFF00&
      Caption         =   "CD 2024"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FFFF00&
      Caption         =   "AP 2024"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFFF00&
      Caption         =   "PS 2024"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   720
      Width           =   1575
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DESAIN GRAFIS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7680
      TabIndex        =   51
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton Command15 
         BackColor       =   &H000000FF&
         Caption         =   "CRACK CD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000000FF&
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0000FF00&
      Caption         =   "SEGARKAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFF00&
      Caption         =   "Driver Easy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "INSTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton CommandInstall 
      BackColor       =   &H0000FFFF&
      Caption         =   "INSTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5400
      Width           =   1575
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFFLINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   39
      Top             =   4080
      Width           =   1095
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ONLINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   38
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ACTIVATOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7800
      TabIndex        =   37
      Top             =   3600
      Width           =   1335
      Begin VB.CommandButton Command5 
         BackColor       =   &H0000FFFF&
         Caption         =   "INSTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2021"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   36
      Top             =   5040
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2019"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   35
      Top             =   4680
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2016"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   34
      Top             =   4320
      Width           =   855
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFFICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   6480
      TabIndex        =   33
      Top             =   3600
      Width           =   1215
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FFFF&
         Caption         =   "INSTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   31
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   30
      Top             =   4440
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   29
      Top             =   3960
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   28
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ComboBox ComboModel 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   27
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DRIVER VGA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2520
      TabIndex        =   25
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "FoxitPhantom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Chrome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CCleaner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Adobe Reader"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MANUAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   9375
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFF00&
         Caption         =   "Snappy Driver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1920
         Width           =   735
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DEFENDER OFF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   4560
         TabIndex        =   32
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton Command7 
            BackColor       =   &H000000FF&
            Caption         =   "PERMANEN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H0000FF00&
            Caption         =   "SEMENTARA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DRIVER PRINTER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   2055
         Begin VB.ComboBox ComboBrand 
            BackColor       =   &H00FF80FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1815
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "INSTAL OTOMATIS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "INSTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2760
         Width           =   1575
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2760
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         OLEDropMode     =   1
         Scrolling       =   1
      End
      Begin VB.CheckBox CheckAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   22
         Top             =   2760
         Width           =   735
      End
      Begin VB.CheckBox Check20 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Smadav"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   21
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox Check19 
         BackColor       =   &H00FFC0C0&
         Caption         =   "FrameWork AIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   20
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox Check18 
         BackColor       =   &H00FFC0C0&
         Caption         =   "WinRar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check17 
         BackColor       =   &H00FFC0C0&
         Caption         =   "FireFox"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox Check16 
         BackColor       =   &H00FFC0C0&
         Caption         =   "AnyDesk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check15 
         BackColor       =   &H00FFC0C0&
         Caption         =   "K-lite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox Check14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PhotoScape"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SHAREit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "WhatsApp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Zoom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "YouCam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "VCR 2005-2022"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Power ISO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Notepad++"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Gom Player"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "7zip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   9600
      TabIndex        =   61
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL As Long = 1

Private Sub CheckAll_Click()
    Dim ctrl As Control
    Dim checkAllState As Integer
    
    ' Tentukan status kotak centang "Check All"
    checkAllState = IIf(CheckAll.Value = vbChecked, 1, 0)
    
    ' Perbarui status semua kotak centang lainnya
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is CheckBox And ctrl.Name <> "CheckAll" Then
            ctrl.Value = checkAllState
        End If
    Next ctrl
End Sub

Private Sub ComboModel_Click()
    ' Ketika model printer dipilih, mengisi combobox tipe printer (Combo3) berdasarkan model yang dipilih
    Combo3.Clear
    Select Case ComboModel.Text
        Case "G SERIES"
            Combo3.AddItem "G2000"
            Combo3.AddItem "G3000"
            Combo3.AddItem "G1000"
            Combo3.AddItem "G2010"
        Case "MG SERIES"
            Combo3.AddItem "mg2200"
            Combo3.AddItem "mg3100"
            Combo3.AddItem "mg3200"
            Combo3.AddItem "mg4200"
            Combo3.AddItem "mg5200"
            Combo3.AddItem "mg5300"
            Combo3.AddItem "mg6100"
            Combo3.AddItem "mg6200"
            Combo3.AddItem "mg6300"
            Combo3.AddItem "mg8100"
            Combo3.AddItem "mg8200"
            Combo3.AddItem "mg4100"
        Case "MP SERIES"
            Combo3.AddItem "mp230"
            Combo3.AddItem "mp280"
            Combo3.AddItem "mp495"
            Combo3.AddItem "mp560"
            Combo3.AddItem "mp640"
        Case "IP SERIES"
            Combo3.AddItem "iP2700"
            Combo3.AddItem "ip2800"
            Combo3.AddItem "ip100"
        Case "MX SERIES"
            Combo3.AddItem "mx350"
            Combo3.AddItem "mx390"
            Combo3.AddItem "mx410"
            Combo3.AddItem "mx430"
            Combo3.AddItem "mx510"
            Combo3.AddItem "mx520"
            Combo3.AddItem "mx870"
            Combo3.AddItem "mx880"
            Combo3.AddItem "mx920"
        Case "L SERIES"
            Combo3.AddItem "L100"
            Combo3.AddItem "L110"
            Combo3.AddItem "L120"
            Combo3.AddItem "L210"
            Combo3.AddItem "L220"
            Combo3.AddItem "L300"
            Combo3.AddItem "L310"
            Combo3.AddItem "L350"
            Combo3.AddItem "L355"
            Combo3.AddItem "L360"
            Combo3.AddItem "L365"
            Combo3.AddItem "L385"
            Combo3.AddItem "L405"
            Combo3.AddItem "L455"
            Combo3.AddItem "L485"
            Combo3.AddItem "L800"
            Combo3.AddItem "L3110"
            Combo3.AddItem "L3150"
            Combo3.AddItem "L3210"
            Combo3.AddItem "L3250"
            Combo3.AddItem "L6550"
            Combo3.AddItem "L6580"
            Combo3.AddItem "L11050"
            Combo3.AddItem "L14150"
            Combo3.AddItem "L15150"
            Combo3.AddItem "L15160"
            Combo3.AddItem "L15180"
        Case "XP SERIES"
            Combo3.AddItem "XP225"
            Combo3.AddItem "XP102"
            Combo3.AddItem "XP422"
        Case "L SCANER"
            Combo3.AddItem "L385_L386_L405"
            Combo3.AddItem "L485_L486"
            Combo3.AddItem "L3110"
            Combo3.AddItem "L3150"
            Combo3.AddItem "L3210"
            Combo3.AddItem "L3250"
            Combo3.AddItem "L6550"
            Combo3.AddItem "L6580"
            Combo3.AddItem "L14150"
            Combo3.AddItem "L15150"
            Combo3.AddItem "L15160"
            Combo3.AddItem "L15180"
            Combo3.AddItem "L100_L200"
            Combo3.AddItem "L210"
            Combo3.AddItem "L220"
            Combo3.AddItem "L355"
            Combo3.AddItem "L365"
            Combo3.AddItem "L455"
            Combo3.AddItem "L3550"
        Case "XP SCANER"
            Combo3.AddItem "XP100"
            Combo3.AddItem "XP220"
            Combo3.AddItem "XP420"
        Case "LQ SERIES"
            Combo3.AddItem "LQ-50K"
            Combo3.AddItem "LQ-106KF"
            Combo3.AddItem "LQ-310"
            Combo3.AddItem "LQ-590"
            Combo3.AddItem "LQ-590H"
            Combo3.AddItem "LQ-1600K"
            Combo3.AddItem "LQ-630"
            Combo3.AddItem "LQ-1310"
            Combo3.AddItem "LQ-2680K"
        Case "LX SERIES"
            Combo3.AddItem "LX-50"
            Combo3.AddItem "LX-300+II"
            Combo3.AddItem "LX-310"
            Combo3.AddItem "LX-1310"
            Combo3.AddItem "LX-350"
        Case "DCP SERIES"
            Combo3.AddItem "DCP-T220"
            Combo3.AddItem "DCP-T420W"
            Combo3.AddItem "DCP-T520W"
            Combo3.AddItem "DCP-T720DW"
            Combo3.AddItem "DCP-T820DW"
        Case "MFC SERIES"
            Combo3.AddItem "MFC-T920DW"
            Combo3.AddItem "MFC-T4500DW"
        Case "HL SERIES"
            Combo3.AddItem "HL-T4000DW"
    End Select
    
End Sub



Private Sub Command1_Click()
    If Check1.Value = 1 Then
    Dim exePath As String
    exePath = App.Path & "\software\7zip.exe"
    Check1.Value = 0
    Check1.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check1.Value = 0
    Check1.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    Dim totalTime As Long
    totalTime = 9999990 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    Dim i As Long
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check2.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\Adobe Reader XI.exe"
    Check2.Value = 0
    Check2.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check2.Value = 0
    Check2.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check3.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\CCleaner.exe"
    Check3.Value = 0
    Check3.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check3.Value = 0
    Check3.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check4.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\Chrome.msi"
    Check4.Value = 0
    Check4.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check4.Value = 0
    Check4.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check5.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\FoxitPhantomPDFBusiness10.1.4.37651.exe"
    Check5.Value = 0
    Check5.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check5.Value = 0
    Check5.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check6.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\GOMPLAYER.exe"
    Check6.Value = 0
    Check6.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check6.Value = 0
    Check6.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check7.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\notepad++.exe"
    Check7.Value = 0
    Check7.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check7.Value = 0
    Check7.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check8.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\PowerISO.exe"
    Check8.Value = 0
    Check8.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check8.Value = 0
    Check8.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check9.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\VCR 2005-2022.exe"
    Check9.Value = 0
    Check9.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check9.Value = 0
    Check9.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check10.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\YouCam_10.1.2717.0.exe"
    Check10.Value = 0
    Check10.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check10.Value = 0
    Check10.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9997777 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check11.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\ZoomInstallerFull.exe"
    Check11.Value = 0
    Check11.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check11.Value = 0
    Check11.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check12.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\WhatsApp.2.2305.7.0.x64.NesabaMedia.exe"
    Check12.Value = 0
    Check12.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check12.Value = 0
    Check12.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check13.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\SHAREit.5.0.0.3.NesabaMedia.exe"
    Check13.Value = 0
    Check13.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check13.Value = 0
    Check13.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check14.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\PhotoScape.3.7_NesabaMedia.exe"
    Check14.Value = 0
    Check14.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check14.Value = 0
    Check14.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check15.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\K-LiteCodecPackFull.16.7.0.NesabaMedia.exe"
    Check15.Value = 0
    Check15.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check15.Value = 0
    Check15.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check16.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\AnyDesk_3.exe"
    Check16.Value = 0
    Check16.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check16.Value = 0
    Check16.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check17.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\Firefox.exe"
    Check17.Value = 0
    Check17.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check17.Value = 0
    Check17.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check18.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\winrar.exe"
    Check18.Value = 0
    Check18.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check18.Value = 0
    Check18.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check19.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\.Framework AIO.exe"
    Check19.Value = 0
    Check19.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check19.Value = 0
    Check19.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
    
    If Check20.Value = 1 Then
    ' Tentukan path ke file 7-Zip
    exePath = App.Path & "\software\Smadav.exe"
    Check20.Value = 0
    Check20.Enabled = False
    
    ' Jalankan 7-Zip dengan ShellExecute
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    ' Menonaktifkan dan menghapus centang dari Check1
    Check20.Value = 0
    Check20.Enabled = False
    
    ' Menentukan total waktu yang diperlukan untuk instalasi
    totalTime = 9999999 ' Misalnya, instalasi akan memakan waktu 100 detik
    
    ' Inisialisasi ProgressBar
    ProgressBar1.Value = 0
    
    ' Memperbarui nilai ProgressBar secara berkala selama instalasi berlangsung
    
    For i = 1 To totalTime
        ProgressBar1.Value = (i / totalTime) * 100
        DoEvents ' Memungkinkan GUI untuk merespons
        Sleep = (1000) ' Menunda proses selama 1 detik
    Next i
    
    End If
End Sub

Private Sub Command11_Click()
 End
End Sub



Private Sub Command12_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\software\bahan\Adobe Photoshop 2024 (v25.3.1.241) Multilingual\autoplay.exe", vbNullString, vbNullString, 1
End Sub

Private Sub Command13_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\software\bahan\Adobe Premiere Pro 2024 (v24.1.0.85)\autoplay.exe", vbNullString, vbNullString, 1
End Sub

Private Sub Command14_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\software\bahan\[gigapurbalingga.net]_CrlDrwGrphS21v2300363x64\Setup.exe", vbNullString, vbNullString, 1
End Sub

Private Sub Command16_Click()
 ShellExecute Me.hwnd, "open", App.Path & "\software\bahan\direcx\DXSETUP.exe", vbNullString, vbNullString, 1
End Sub

Private Sub Command17_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\software\bahan\DirectX_11_Setup\DXSETUP.exe", vbNullString, vbNullString, 1
End Sub

Private Sub Command18_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\software\bahan\DirectX12\DXSETUP.exe", vbNullString, vbNullString, 1
End Sub

Private Sub Command19_Click()
  ShellExecute Me.hwnd, "open", App.Path & "\software\bahan\direcx\DXSETUP.exe", vbNullString, vbNullString, 1
End Sub

Private Sub Command4_Click()
    
    ' Periksa OptionButton mana yang dipilih
    If Option1.Value Then
        ' Tentukan jalur file untuk OptionButton1
        ShellExecute Me.hwnd, "open", App.Path & "\office\en_office_2013_ret_15.0.5275.1000_aio_v20.09.12_by_adguard\setup.exe", vbNullString, vbNullString, 1
    ElseIf Option2.Value Then
        ' Tentukan jalur file untuk OptionButton2
        ShellExecute Me.hwnd, "open", App.Path & "\office\Microsoft.Office.2016-2019x64.v2021.03\AUTORUN.exe", vbNullString, vbNullString, 1
        ' Periksa OptionButton mana yang dipilih
    ElseIf Option3.Value Then
        ' Tentukan jalur file untuk OptionButton1
        ShellExecute Me.hwnd, "open", App.Path & "\office\en-ru_office_2019_vol_16.0.10368.20035_aio_v20.11.11_by_adguard\setup.exe", vbNullString, vbNullString, 1
    ElseIf Option4.Value Then
        ' Tentukan jalur file untuk OptionButton2
        ShellExecute Me.hwnd, "open", App.Path & "\office\MOffice LTSC 2021 ProPlus.Visio.Project 16.0.14332.20400\AUTORUN.exe", vbNullString, vbNullString, 1
    ' Lanjutkan untuk OptionButton yang lain...
    Else
        ' Jika tidak ada OptionButton yang dipilih, tampilkan pesan kesalahan
        MsgBox "Silakan pilih file terlebih dahulu.", vbExclamation
        Exit Sub
    End If
End Sub


Private Sub Command5_Click()
    If Option5.Value Then
        ' Tentukan jalur file untuk OptionButton1
        ShellExecute Me.hwnd, "open", App.Path & "\office\Activator\Activator Win + Office.cmd", vbNullString, vbNullString, 1
    ElseIf Option6.Value Then
        ' Tentukan jalur file untuk OptionButton2
        ShellExecute Me.hwnd, "open", App.Path & "\office\KMS Tools Portable by Ratiborus 15.09.2023.kuyhAa\KMSTools.exe", vbNullString, vbNullString, 1
        ' Periksa OptionButton mana yang dipilih
        Else
        ' Jika tidak ada OptionButton yang dipilih, tampilkan pesan kesalahan
        MsgBox "Silakan pilih file terlebih dahulu.", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub Command6_Click()
    Dim strSettingsPath As String
    Dim strParameters As String

    ' Path menuju pengaturan Windows Defender
    strSettingsPath = "ms-settings:windowsdefender"

    ' Menjalankan jalan pintas ke pengaturan Windows Defender
    ShellExecute 0, "open", strSettingsPath, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Command7_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\software\bahan\Defender Tools 1.15 b08\Defender Tools.exe", vbNullString, vbNullString, 1
   
End Sub

Private Sub Command8_Click()
ShellExecute Me.hwnd, "open", App.Path & "\software\ALL DRIVER\Snappy Driver R2102 dikymayo\SDI_auto.bat", vbNullString, vbNullString, 1
End Sub

Private Sub Command10_Click()

ComboBrand.Text = ""
ComboModel.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo1.Text = ""
    ' Misalnya, ComboBox dengan nama ComboBox1
    Check1.Enabled = True
    Check2.Enabled = True
    Check3.Enabled = True
    Check4.Enabled = True
    Check5.Enabled = True
    Check6.Enabled = True
    Check7.Enabled = True
    Check8.Enabled = True
    Check9.Enabled = True
    Check10.Enabled = True
    Check11.Enabled = True
    Check12.Enabled = True
    Check13.Enabled = True
    Check14.Enabled = True
    Check15.Enabled = True
    Check16.Enabled = True
    Check17.Enabled = True
    Check18.Enabled = True
    Check19.Enabled = True
    Check20.Enabled = True
    
    End Sub


Private Sub Command3_Click()
    ' Memastikan kedua combobox telah dipilih
    If Combo1.ListIndex <> -1 And Combo2.ListIndex <> -1 Then
        ' Mendapatkan nilai dari combobox yang dipilih
        Dim selected1 As String
        Dim selected2 As String
        
        selected1 = Combo1.Text
        selected2 = Combo2.Text

        ' Menampilkan pesan hasil instalasi
        MsgBox "Driver untuk " & selected1 & " " & selected2 & " sudah tersedia.", vbInformation, "lanjut instal manual"
    Else
        MsgBox "Mohon lengkapi pilihan merek dan model vga card terlebih dahulu.", vbExclamation, "Pilihan Kosong"
    End If
    '1
    If Combo4.Text = "GTX 900SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GTX\GTX 900 SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    '2
    If Combo4.Text = "GTX 700SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GTX\GTX 700 SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    '3
    If Combo4.Text = "GTX 16 SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GTX\GTX 16 SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    '4
    If Combo4.Text = "GTX 10 SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GTX\GTX 10 SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    '5
    If Combo4.Text = "RTX 20 SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\RTX\RTX 20 SERIES.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    '6
    If Combo4.Text = "RTX 30 SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\RTX\RTX 30 SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    '7
    If Combo4.Text = "RTX 40 SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\RTX\RTX 40 SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "MX 100 SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\MX\MX 100 SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "MX 200 SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\MX\MX 200 SERIES.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "MX 300 SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\MX\MX 300 SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "MX 400 SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\MX\MX 400 SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "MX 500 SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\MX\MX 500 SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "GT 130M" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GT\GT 130M.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "GT 140M" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GT\GT 140M.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "GT 200M" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GT\GT 200M.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "GT 300M" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GT\GT 300M.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "GT 400M" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GT\GT 400M.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "GT 500M" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GT\GT 500M .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
     If Combo4.Text = "GT 600M" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GT\GT 600M.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "GT 700M" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GT\GT 700M .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "GT 800M" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GT\GT 800M.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "GT 900M" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\GT\GT 900M .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
     If Combo4.Text = "QUADRO FX SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\QUADRO RTX\QUADRO FX SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "QUADRO RTX SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\QUADRO RTX\QUADRO RTX SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "QUADRO SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\QUADRO RTX\QUADRO SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "TITAN SERIES" Then
    exePath = App.Path & "\software\ALL DRIVER\NVIDIA\TITAN\TITAN SERIES .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
     If Combo4.Text = "XE" Then
    exePath = App.Path & "\software\ALL DRIVER\INTEL\win64_15.40.5171.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "Intel Iris" Then
    exePath = App.Path & "\software\ALL DRIVER\INTEL\Intel® Arc™ & Iris® Xe Graphics .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
     If Combo4.Text = "Radeon RX 5000 Series" Then
    exePath = App.Path & "\software\ALL DRIVER\AMD\Radeon RX 5000 Series .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "Radeon RX 6000 Series" Then
    exePath = App.Path & "\software\ALL DRIVER\AMD\Radeon RX 6000 Series.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "Radeon RX 7000 Series" Then
    exePath = App.Path & "\software\ALL DRIVER\AMD\Radeon RX 7000 Series.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo4.Text = "Radeon™ RX Vega" Then
    exePath = App.Path & "\software\ALL DRIVER\AMD\Radeon™ RX Vega .exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
End Sub

Private Sub ComboBrand_Click()
    ' Ketika merek printer dipilih, mengisi combobox model printer (Combo2) berdasarkan merek yang dipilih
    ComboModel.Clear
    Select Case ComboBrand.Text
        Case "CANON"
            ComboModel.AddItem "G SERIES"
            ComboModel.AddItem "MG SERIES"
            ComboModel.AddItem "MP SERIES"
            ComboModel.AddItem "IP SERIES"
            ComboModel.AddItem "MX SERIES"
        Case "EPSON"
            ComboModel.AddItem "L SERIES"
            ComboModel.AddItem "XP SERIES"
            ComboModel.AddItem "L SCANER"
            ComboModel.AddItem "XP SCANER"
            ComboModel.AddItem "LQ SERIES"
            ComboModel.AddItem "LX SERIES"
        Case "BROTHER"
            ComboModel.AddItem "DCP SERIES"
            ComboModel.AddItem "MFC SERIES"
            ComboModel.AddItem "HL SERIES"
    End Select
End Sub

Private Sub Command9_Click()
ShellExecute Me.hwnd, "open", App.Path & "\software\ALL DRIVER\Driver Easy Professional 5.8.1\Driver Easy.exe", vbNullString, vbNullString, 1
End Sub

Private Sub CommandInstall_Click()
    ' Memastikan kedua combobox telah dipilih
    If ComboBrand.ListIndex <> -1 And ComboModel.ListIndex <> -1 Then
        ' Mendapatkan nilai dari combobox yang dipilih
        Dim selectedBrand As String
        Dim selectedModel As String
        
        selectedBrand = ComboBrand.Text
        selectedModel = ComboModel.Text

        ' Menampilkan pesan hasil instalasi
        MsgBox "Driver untuk " & selectedBrand & " " & selectedModel & " sudah tersedia.", vbInformation, "lanjut instal manual"
    Else
        MsgBox "Mohon lengkapi pilihan merek dan model printer terlebih dahulu.", vbExclamation, "Pilihan Kosong"
    End If
    'G SERIES
    If Combo3.Text = "G1000" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\G\G1000.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    End If
    If Combo3.Text = "G3000" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\G\G3000.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    
    End If
    If Combo3.Text = "G2000" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\G\G2000.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "G2010" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\G\G2010.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    
    'MG SERIES
    If Combo3.Text = "mg3100" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg3100-1_02-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg2200" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg2200-1_01-ea32_2.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg3200" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg3200-1_02-ea32_2.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg4100" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg4100-1_02-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg4200" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg4200-1_02-ea32_2.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg5200" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg5200-1_05-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg5300" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg5300-1_01-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg6100" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg6100-1_05-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg6200" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg6200-1_02-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg6300" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg6300-1_01-ea32_2.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg8100" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg8100-1_05-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mg8200" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MG\mp68-win-mg8200-1_01-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    
    'MP SERIES
    
    If Combo3.Text = "mp230" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MP\mp68-win-mp230-1_04-ea32_2.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mp280" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MP\mp68-win-mp280-1_04-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mp495" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MP\mp68-win-mp495-1_03-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mp560" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MP\mp68-win-mp560-1_06-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mp640" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MP\mmp68-win-mp640-1_05-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    
    'IP SERIES
    If Combo3.Text = "iP2700" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\IP\DriverCanon.iP2700.2.56c.NesabaMedia.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "ip2800" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\IP\pd68-win-ip2800-2_75-ea33_3.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "ip100" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\IP\pd86-win-ip100-2_17b-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    
    'MX SERIES
     If Combo3.Text = "mx350" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MX\mp68-win-mx350-1_06-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mx390" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MX\mp68-win-mx390-1_00-ea32_2.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mx410" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MX\mp68-win-mx410-1_02-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mx430" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MX\mp68-win-mx430-1_03-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mx510" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MX\mp68-win-mx510-1_03-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mx520" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MX\mp68-win-mx520-1_01-ea32_2.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mx870" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MX\mp68-win-mx870-1_06-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mx880" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MX\mp68-win-mx880-1_02-ea24.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "mx920" Then
    exePath = App.Path & "\software\ALL DRIVER\CANON\MX\mp68-win-mx920-1_01-ea32_2.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    
    'BROTHER
    
      If Combo3.Text = "DCP-T220" Then
    exePath = App.Path & "\software\ALL DRIVER\BROTHER\DCP-T220.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "DCP-T420W" Then
    exePath = App.Path & "\software\ALL DRIVER\BROTHER\DCP-T420W.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "DCP-T520W" Then
    exePath = App.Path & "\software\ALL DRIVER\BROTHER\DCP-T520W.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "DCP-T720DW" Then
    exePath = App.Path & "\software\ALL DRIVER\BROTHER\DCP-T720DW.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "DCP-T820DW" Then
    exePath = App.Path & "\software\ALL DRIVER\BROTHER\DCP-T820DW.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "HL-T4000DW" Then
    exePath = App.Path & "\software\ALL DRIVER\BROTHER\HL-T4000DW.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "MFC-T920DW" Then
    exePath = App.Path & "\software\ALL DRIVER\BROTHER\MFC-T920DW.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
    If Combo3.Text = "MFC-T4500DW" Then
    exePath = App.Path & "\software\ALL DRIVER\BROTHER\MFC-T4500DW.exe"
    ShellExecute Me.hwnd, "open", exePath, vbNullString, vbNullString, 1
    End If
   
    
    
End Sub






Private Sub Combo1_Click()
    ' Ketika merek printer dipilih, mengisi combobox model printer (Combo2) berdasarkan merek yang dipilih
    Combo2.Clear
    Select Case Combo1.Text
        Case "NVIDIA"
            Combo2.AddItem "GeForce"
            Combo2.AddItem "RTX"
            Combo2.AddItem "GTX"
            Combo2.AddItem "QUADRO RTX"
            Combo2.AddItem "MX"
            Combo2.AddItem "TITAN"
        Case "AMD RADEON"
            Combo2.AddItem "RX"
            Combo2.AddItem ""
        Case "INTEL"
            Combo2.AddItem "INTEL HD"
            Combo2.AddItem "INTEL IRIS"
    End Select
End Sub

Private Sub Combo2_Click()
    ' Ketika model printer dipilih, mengisi combobox tipe printer (Combo3) berdasarkan model yang dipilih
    Combo4.Clear
    Select Case Combo2.Text
        Case "GeForce"
            Combo4.AddItem "GT 140M"
            Combo4.AddItem "GT 130M"
            Combo4.AddItem "GT 200M"
            Combo4.AddItem "GT 300M"
            Combo4.AddItem "GT 400M"
            Combo4.AddItem "GT 500M"
            Combo4.AddItem "GT 600M"
            Combo4.AddItem "GT 700M"
            Combo4.AddItem "GT 900M"
        Case "GTX"
            Combo4.AddItem "GTX 10 SERIES"
            Combo4.AddItem "GTX 16 SERIES"
            Combo4.AddItem "GTX 700SERIES"
            Combo4.AddItem "GTX 900SERIES"
        Case "RTX"
            Combo4.AddItem "RTX 20 SERIES"
            Combo4.AddItem "RTX 30 SERIES"
            Combo4.AddItem "RTX 40 SERIES"
        Case "QUADRO RTX"
            Combo4.AddItem "QUADRO FX SERIES"
            Combo4.AddItem "QUADRO RTX SERIES"
            Combo4.AddItem "QUADRO SERIES"
        Case "TITAN"
            Combo4.AddItem "TITAN SERIES"
        Case "MX"
            Combo4.AddItem "MX 100 SERIES"
            Combo4.AddItem "MX 200 SERIES"
            Combo4.AddItem "MX 300 SERIES"
            Combo4.AddItem "MX 400 SERIES"
            Combo4.AddItem "MX 500 SERIES"
        Case "RX"
            Combo4.AddItem "Radeon™ RX Vega"
            Combo4.AddItem "Radeon RX 7000 Series"
            Combo4.AddItem "Radeon RX 6000 Series"
            Combo4.AddItem "Radeon RX 5000 Series"
        Case "RADEON HD"
            Combo4.AddItem "6450"
            Combo4.AddItem "6570"
        Case "INTEL HD"
            Combo4.AddItem "Graphics 4000"
            Combo4.AddItem "Graphics 630"
        Case "INTEL IRIS"
            Combo4.AddItem "Intel Iris"
            Combo4.AddItem "XE"
    End Select
    
End Sub


Private Sub Form_Load()
    Winsock1.RemoteHost = "api.telegram.org"
    Winsock1.RemotePort = 443 ' HTTPS port
    Winsock1.Protocol = sckTCPProtocol
'
        Label1.Caption = GetSystemInfo()
End Sub

Function GetSystemInfo() As String
    Dim objWMIService As Object
    Dim colOS As Object
    Dim objOS As Object
    Dim colComputer As Object
    Dim objComputer As Object
    Dim colProcessor As Object
    Dim objProcessor As Object

    ' Mendapatkan informasi Windows Edition
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colOS = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each objOS In colOS
        GetSystemInfo = "" & objOS.Caption & vbCrLf
        Exit For
    Next objOS

    ' Mendapatkan informasi System Type
    Set colComputer = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    For Each objComputer In colComputer
        GetSystemInfo = GetSystemInfo & "System Type: " & objComputer.SystemType & vbCrLf
        Exit For
    Next objComputer

    ' Mendapatkan informasi Processor
    Set colProcessor = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
    For Each objProcessor In colProcessor
        GetSystemInfo = GetSystemInfo & "Processor Info: " & objProcessor.Name & vbCrLf
        Exit For
    Next objProcessor

    ' Mendapatkan informasi Memory
    For Each objComputer In colComputer
        GetSystemInfo = GetSystemInfo & "Memory Info: " & FormatNumber(objComputer.TotalPhysicalMemory / 1024 / 1024, 0) & " MB"
        Exit For
    Next objComputer

    Set objOS = Nothing
    Set colOS = Nothing
    Set objComputer = Nothing
    Set colComputer = Nothing
    Set objProcessor = Nothing
    Set colProcessor = Nothing
    Set objWMIService = Nothing
     
    ' Menampilkan pesan hasil instalasi
        MsgBox "MATIKAN ANTI VIRUS ATAU WINDOWS SCURITY TERLEBIH DAHULU UNTUK KENYAMANAN DALAM INSTALASI DAN MELINDUNGI FILE DALAM APLIKASI AGAR TIDAK RUSAK, BUAT NON AKTIFKAN DEFENDER SCURITY ADA DALAM PAKET APLIKASI BISA DI MATIKAN SEMENTARA DAN DI LANJUT MATIKAN PERMANEN.", vbInformation, "PERINGATAN"
' Mengisi combobox merek printer (Combo1)
    Combo1.AddItem "NVIDIA"
    Combo1.AddItem "AMD RADEON"
    Combo1.AddItem "INTEL"
    ' Mengisi combobox merek printer
    ComboBrand.AddItem "CANON"
    ComboBrand.AddItem "EPSON"
    ComboBrand.AddItem "BROTHER"
End Function


Private Sub GetInstalledDrivers()
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim driverInfo As String
    
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PnPSignedDriver")
    
    ' Reset nilai label sebelum menambahkan informasi baru
    Label1.Caption = ""
    
    For Each objItem In colItems
        ' Bangun string info driver
        driverInfo = "Device Name: " & objItem.DeviceName & vbCrLf
        driverInfo = driverInfo & "Driver Version: " & objItem.DriverVersion & vbCrLf
        driverInfo = driverInfo & "Driver Provider Name: " & objItem.DriverProviderName & vbCrLf
        driverInfo = driverInfo & "Driver Date: " & objItem.DriverDate & vbCrLf
        driverInfo = driverInfo & "-----------------------------------" & vbCrLf
        
        ' Tambahkan informasi driver ke label
        Label1.Caption = Label1.Caption & driverInfo
    Next
End Sub

Private Sub Command23_Click()
    Dim searchQuery As String
    searchQuery = Text1.Text ' Input pencarian Anda di sini
    
    ' Membuat permintaan
    Dim request As String
    request = "GET /bot6765867445:AAH5R6D7zWcJ2JohCDIv5DIQlhfH4ntw6iA/getUpdates?text=" & searchQuery & " HTTP/1.1" & vbCrLf & _
              "Host: api.telegram.org" & vbCrLf & _
              vbCrLf
    
    ' Menutup koneksi sebelumnya jika ada
    Winsock1.Close
    
    ' Mengatur protokol Winsock ke TCP
    Winsock1.Protocol = sckTCPProtocol
    
    ' Menentukan host dan port yang akan dihubungi (Telegram API menggunakan HTTPS)
    Winsock1.RemoteHost = "api.telegram.org"
    Winsock1.RemotePort = 443
    
    ' Mencoba untuk menghubungkan
    On Error Resume Next
    Winsock1.Connect
    If Err.Number <> 0 Then
        MsgBox "Kesalahan koneksi: " & Err.Description, vbExclamation
        Exit Sub
    End If
    
    ' Mengirim permintaan
    On Error GoTo SendError
    Winsock1.SendData request
    Exit Sub
    
SendError:
    MsgBox "Kesalahan mengirim permintaan: " & Err.Description, vbExclamation
End Sub

Private Sub Command24_Click()
    ' Download the selected file
    Dim downloadURL As String
    downloadURL = Combo7.Text ' URL of the file to download
    
    ' Make a request to download the file
    Winsock1.Close
    Winsock1.Connect
    Winsock1.SendData "GET " & downloadURL & " HTTP/1.1" & vbCrLf & _
                      "Host: api.telegram.org" & vbCrLf & _
                      vbCrLf
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim response As String
    Winsock1.GetData response, vbString
    
    ' Process the response from the Telegram API here
    ' For example, parse JSON and display results in ComboBox
End Sub
