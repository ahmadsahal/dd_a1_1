VERSION 5.00
Begin VB.Form frmSulap 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   Icon            =   "frmSulap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   168
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "YES"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   52
      Top             =   6240
      Width           =   1095
   End
   Begin VB.PictureBox Pic 
      Height          =   2895
      Index           =   1
      Left            =   2978
      ScaleHeight     =   2835
      ScaleWidth      =   5955
      TabIndex        =   53
      Top             =   3173
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   160
         Left            =   2280
         TabIndex        =   198
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   99
         Left            =   120
         TabIndex        =   77
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   98
         Left            =   840
         TabIndex        =   76
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   97
         Left            =   1560
         TabIndex        =   75
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   96
         Left            =   3000
         TabIndex        =   74
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   95
         Left            =   3720
         TabIndex        =   73
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   94
         Left            =   4440
         TabIndex        =   72
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   93
         Left            =   5160
         TabIndex        =   71
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   92
         Left            =   120
         TabIndex        =   70
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   91
         Left            =   840
         TabIndex        =   69
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   90
         Left            =   1560
         TabIndex        =   68
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   89
         Left            =   2280
         TabIndex        =   67
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   88
         Left            =   3000
         TabIndex        =   66
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   87
         Left            =   3720
         TabIndex        =   65
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   86
         Left            =   4440
         TabIndex        =   64
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   85
         Left            =   5160
         TabIndex        =   63
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   84
         Left            =   120
         TabIndex        =   62
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   83
         Left            =   840
         TabIndex        =   61
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   82
         Left            =   1560
         TabIndex        =   60
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   81
         Left            =   2280
         TabIndex        =   59
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   80
         Left            =   3000
         TabIndex        =   58
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "43"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   79
         Left            =   3720
         TabIndex        =   57
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "45"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   78
         Left            =   4440
         TabIndex        =   56
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "47"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   77
         Left            =   5160
         TabIndex        =   55
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "49"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   76
         Left            =   120
         TabIndex        =   54
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.PictureBox Pic 
      Height          =   2895
      Index           =   0
      Left            =   2978
      ScaleHeight     =   2835
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   3173
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   49
         Left            =   840
         TabIndex        =   50
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "49"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   48
         Left            =   120
         TabIndex        =   49
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "48"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   47
         Left            =   5160
         TabIndex        =   48
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "47"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   46
         Left            =   4440
         TabIndex        =   47
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "46"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   45
         Left            =   3720
         TabIndex        =   46
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "45"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   44
         Left            =   3000
         TabIndex        =   45
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "44"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   43
         Left            =   2280
         TabIndex        =   44
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "43"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   42
         Left            =   1560
         TabIndex        =   43
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   41
         Left            =   840
         TabIndex        =   42
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   40
         Left            =   120
         TabIndex        =   41
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   39
         Left            =   5160
         TabIndex        =   40
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   38
         Left            =   4440
         TabIndex        =   39
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   37
         Left            =   3720
         TabIndex        =   38
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   36
         Left            =   3000
         TabIndex        =   37
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   2280
         TabIndex        =   36
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   34
         Left            =   1560
         TabIndex        =   35
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   33
         Left            =   840
         TabIndex        =   34
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   32
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   5160
         TabIndex        =   32
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   4440
         TabIndex        =   31
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   3720
         TabIndex        =   30
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   3000
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   2280
         TabIndex        =   28
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   1560
         TabIndex        =   27
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   840
         TabIndex        =   26
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   5160
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   4440
         TabIndex        =   23
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   3720
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   3000
         TabIndex        =   21
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   2280
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   1560
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   840
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   5160
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   4440
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   3720
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   3000
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   2280
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   1560
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   840
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   5160
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   4440
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3720
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3000
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2280
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox Pic 
      Height          =   2895
      Index           =   2
      Left            =   2978
      ScaleHeight     =   2835
      ScaleWidth      =   5955
      TabIndex        =   173
      Top             =   3173
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   209
         Left            =   120
         TabIndex        =   197
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   208
         Left            =   840
         TabIndex        =   196
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   207
         Left            =   1560
         TabIndex        =   195
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   206
         Left            =   2280
         TabIndex        =   194
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   205
         Left            =   3000
         TabIndex        =   193
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   204
         Left            =   3720
         TabIndex        =   192
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   203
         Left            =   4440
         TabIndex        =   191
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   202
         Left            =   5160
         TabIndex        =   190
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   201
         Left            =   120
         TabIndex        =   189
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   200
         Left            =   840
         TabIndex        =   188
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   199
         Left            =   1560
         TabIndex        =   187
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   198
         Left            =   2280
         TabIndex        =   186
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   197
         Left            =   3000
         TabIndex        =   185
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   196
         Left            =   3720
         TabIndex        =   184
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   195
         Left            =   4440
         TabIndex        =   183
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   194
         Left            =   5160
         TabIndex        =   182
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   193
         Left            =   120
         TabIndex        =   181
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   192
         Left            =   840
         TabIndex        =   180
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   191
         Left            =   1560
         TabIndex        =   179
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   190
         Left            =   2280
         TabIndex        =   178
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "43"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   189
         Left            =   3000
         TabIndex        =   177
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "46"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   188
         Left            =   3720
         TabIndex        =   176
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "47"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   187
         Left            =   4440
         TabIndex        =   175
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   186
         Left            =   5160
         TabIndex        =   174
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.PictureBox picNilai 
      Height          =   2895
      Left            =   2978
      ScaleHeight     =   2835
      ScaleWidth      =   5955
      TabIndex        =   169
      Top             =   3173
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Timer tmrUlang 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   120
         Top             =   1800
      End
      Begin VB.Timer tmrProgress 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   1320
      End
      Begin VB.PictureBox picBar 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   480
         ScaleHeight     =   375
         ScaleMode       =   0  'User
         ScaleWidth      =   101.173
         TabIndex        =   171
         Top             =   2400
         Visible         =   0   'False
         Width           =   5175
         Begin VB.PictureBox picProgress 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   0
            ScaleHeight     =   1
            ScaleMode       =   0  'User
            ScaleWidth      =   0.111
            TabIndex        =   172
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.CommandButton cmdNilai 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   72
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   720
         TabIndex        =   170
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.PictureBox Pic 
      Height          =   2895
      Index           =   6
      Left            =   2978
      ScaleHeight     =   2835
      ScaleWidth      =   5955
      TabIndex        =   148
      Top             =   3173
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   159
         Left            =   120
         TabIndex        =   167
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   158
         Left            =   840
         TabIndex        =   166
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   157
         Left            =   1560
         TabIndex        =   165
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   65
         Left            =   2280
         TabIndex        =   164
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   64
         Left            =   3000
         TabIndex        =   163
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   63
         Left            =   3720
         TabIndex        =   162
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   62
         Left            =   4440
         TabIndex        =   161
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   61
         Left            =   5160
         TabIndex        =   160
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   60
         Left            =   120
         TabIndex        =   159
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   59
         Left            =   840
         TabIndex        =   158
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   58
         Left            =   1560
         TabIndex        =   157
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "43"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   57
         Left            =   2280
         TabIndex        =   156
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "44"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   56
         Left            =   3000
         TabIndex        =   155
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "45"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   55
         Left            =   3720
         TabIndex        =   154
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "46"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   54
         Left            =   4440
         TabIndex        =   153
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "47"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   53
         Left            =   5160
         TabIndex        =   152
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "48"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   52
         Left            =   120
         TabIndex        =   151
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "49"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   51
         Left            =   840
         TabIndex        =   150
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   50
         Left            =   1560
         TabIndex        =   149
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.PictureBox Pic 
      Height          =   2895
      Index           =   5
      Left            =   2978
      ScaleHeight     =   2835
      ScaleWidth      =   5955
      TabIndex        =   128
      Top             =   3173
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   156
         Left            =   120
         TabIndex        =   147
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   155
         Left            =   840
         TabIndex        =   146
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   154
         Left            =   1560
         TabIndex        =   145
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   153
         Left            =   2280
         TabIndex        =   144
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   152
         Left            =   3000
         TabIndex        =   143
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   151
         Left            =   3720
         TabIndex        =   142
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   150
         Left            =   4440
         TabIndex        =   141
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   149
         Left            =   5160
         TabIndex        =   140
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   148
         Left            =   120
         TabIndex        =   139
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   147
         Left            =   840
         TabIndex        =   138
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   146
         Left            =   1560
         TabIndex        =   137
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   145
         Left            =   2280
         TabIndex        =   136
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   144
         Left            =   3000
         TabIndex        =   135
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   143
         Left            =   3720
         TabIndex        =   134
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   142
         Left            =   4440
         TabIndex        =   133
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   141
         Left            =   5160
         TabIndex        =   132
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "48"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   68
         Left            =   120
         TabIndex        =   131
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "49"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   67
         Left            =   840
         TabIndex        =   130
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   66
         Left            =   1560
         TabIndex        =   129
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.PictureBox Pic 
      Height          =   2895
      Index           =   4
      Left            =   2978
      ScaleHeight     =   2835
      ScaleWidth      =   5955
      TabIndex        =   103
      Top             =   3173
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   140
         Left            =   120
         TabIndex        =   127
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   139
         Left            =   840
         TabIndex        =   126
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   138
         Left            =   1560
         TabIndex        =   125
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   137
         Left            =   2280
         TabIndex        =   124
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   136
         Left            =   3000
         TabIndex        =   123
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   135
         Left            =   3720
         TabIndex        =   122
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   134
         Left            =   4440
         TabIndex        =   121
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   133
         Left            =   5160
         TabIndex        =   120
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   132
         Left            =   120
         TabIndex        =   119
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   131
         Left            =   840
         TabIndex        =   118
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   130
         Left            =   1560
         TabIndex        =   117
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   129
         Left            =   2280
         TabIndex        =   116
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   128
         Left            =   3000
         TabIndex        =   115
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   127
         Left            =   3720
         TabIndex        =   114
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   126
         Left            =   4440
         TabIndex        =   113
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   125
         Left            =   5160
         TabIndex        =   112
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   124
         Left            =   120
         TabIndex        =   111
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   75
         Left            =   840
         TabIndex        =   110
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   74
         Left            =   1560
         TabIndex        =   109
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "43"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   73
         Left            =   2280
         TabIndex        =   108
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "44"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   72
         Left            =   3000
         TabIndex        =   107
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "45"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   71
         Left            =   3720
         TabIndex        =   106
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "46"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   70
         Left            =   4440
         TabIndex        =   105
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "47"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   69
         Left            =   5160
         TabIndex        =   104
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.PictureBox Pic 
      Height          =   2895
      Index           =   3
      Left            =   2978
      ScaleHeight     =   2835
      ScaleWidth      =   5955
      TabIndex        =   78
      Top             =   3173
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   123
         Left            =   120
         TabIndex        =   102
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   122
         Left            =   840
         TabIndex        =   101
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   121
         Left            =   1560
         TabIndex        =   100
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   120
         Left            =   2280
         TabIndex        =   99
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   119
         Left            =   3000
         TabIndex        =   98
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   118
         Left            =   3720
         TabIndex        =   97
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   117
         Left            =   4440
         TabIndex        =   96
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   116
         Left            =   5160
         TabIndex        =   95
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   115
         Left            =   120
         TabIndex        =   94
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   114
         Left            =   840
         TabIndex        =   93
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   113
         Left            =   1560
         TabIndex        =   92
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   112
         Left            =   2280
         TabIndex        =   91
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   111
         Left            =   3000
         TabIndex        =   90
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   110
         Left            =   3720
         TabIndex        =   89
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   109
         Left            =   4440
         TabIndex        =   88
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   108
         Left            =   5160
         TabIndex        =   87
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   107
         Left            =   120
         TabIndex        =   86
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   106
         Left            =   840
         TabIndex        =   85
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   105
         Left            =   1560
         TabIndex        =   84
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   104
         Left            =   2280
         TabIndex        =   83
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "44"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   103
         Left            =   3000
         TabIndex        =   82
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "45"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   102
         Left            =   3720
         TabIndex        =   81
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "46"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   101
         Left            =   4440
         TabIndex        =   80
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "47"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   100
         Left            =   5160
         TabIndex        =   79
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Label lblKeterangan 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Silakan pilih sebuah angka, kemudian ingat-ingat dalam hati."
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   570
      Left            =   3848
      TabIndex        =   51
      Top             =   2453
      Width           =   4275
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSulap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bagian As Byte, Nilai%

Private Sub cmdNo_Click()
    If Bagian = 1 Then
        Unload Me
    ElseIf Bagian = 2 Then
        Bagian = 3
        Pic(1).Visible = False
        Pic(2).Visible = True
    ElseIf Bagian = 3 Then
        Bagian = 4
        Pic(2).Visible = False
        Pic(3).Visible = True
    ElseIf Bagian = 4 Then
        Bagian = 5
        Pic(3).Visible = False
        Pic(4).Visible = True
    ElseIf Bagian = 5 Then
        Bagian = 6
        Pic(4).Visible = False
        Pic(5).Visible = True
    ElseIf Bagian = 6 Then
        Bagian = 7
        Pic(5).Visible = False
        Pic(6).Visible = True
    ElseIf Bagian = 7 Then
        Bagian = 8
        Pic(6).Visible = False
        picNilai.Visible = True
        picBar.Visible = True
        tmrProgress.Enabled = True
    ElseIf Bagian = 8 Then
        Unload Me
    End If
End Sub

Private Sub cmdYes_Click()
    If Bagian = 1 Then
        Bagian = 2
        Pic(0).Visible = False
        Pic(1).Visible = True
        lblKeterangan.Caption = "Apakah " & _
        "angka yang Anda pilih ada " & _
        "di kotak? Jika ada klik Yes"
    ElseIf Bagian = 2 Then
        Bagian = 3: Nilai = 1
        Pic(1).Visible = False
        Pic(2).Visible = True
    ElseIf Bagian = 3 Then
        Bagian = 4: Nilai = Nilai + 2
        Pic(2).Visible = False
        Pic(3).Visible = True
    ElseIf Bagian = 4 Then
        Bagian = 5: Nilai = Nilai + 4
        Pic(3).Visible = False
        Pic(4).Visible = True
    ElseIf Bagian = 5 Then
        Pic(4).Visible = False
        Bagian = 6: Nilai = Nilai + 8
        Pic(5).Visible = True
    ElseIf Bagian = 6 Then
        Bagian = 7: Nilai = Nilai + 16
        Pic(5).Visible = False
        Pic(6).Visible = True
    ElseIf Bagian = 7 Then
        Bagian = 8: Nilai = Nilai + 32
        Pic(6).Visible = False
        picNilai.Visible = True
        picBar.Visible = True
        tmrProgress.Enabled = True
    ElseIf Bagian = 8 Then
        Pic(0).Visible = True
        Me.tmrProgress.Enabled = False
        Me.tmrUlang.Enabled = False
        Bagian = 1: Nilai = 0
    End If
End Sub

Private Sub Form_Load()
    Bagian = 1: Nilai = 0
End Sub

Private Sub Form_Resize()
    Me.Move 0, 0, Screen.Width, Screen.Height
End Sub

Private Sub tmrProgress_Timer()
    Me.cmdYes.Enabled = False
    Me.cmdNo.Enabled = False
    Me.picBar.Visible = True
    Me.picProgress.Move 0, 0, picProgress.Width + 1, picProgress.Height
    Me.cmdNilai.Caption = Int(Rnd * 63)
    Me.lblKeterangan.Caption = "Angka yang Anda pilih..."
    If picProgress.Width >= picBar.ScaleWidth Then
        Me.picBar.Visible = False
        Me.picProgress.Width = 0
        Me.tmrProgress.Enabled = False
        Me.cmdNilai.Caption = Nilai
        Me.tmrUlang.Enabled = True
    End If
End Sub

Private Sub tmrUlang_Timer()
    Me.picNilai.Visible = False
    lblKeterangan.Caption = "Mau mengulang?"
    Me.cmdYes.Enabled = True
    Me.cmdNo.Enabled = True
End Sub
