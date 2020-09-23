VERSION 5.00
Begin VB.Form fComDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Common Dialogs"
   ClientHeight    =   5805
   ClientLeft      =   420
   ClientTop       =   2130
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9375
   Begin Project1.ucFolderDialog ucFolderDialog1 
      Height          =   4335
      Left            =   600
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7646
   End
   Begin Project1.ucPageSetupDialog ucPageSetupDialog1 
      Height          =   4575
      Left            =   1080
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
   End
   Begin Project1.ucPrintDialog ucPrintDialog1 
      Height          =   4095
      Left            =   1800
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7223
   End
   Begin Project1.ucColorDialog ucColorDialog1 
      Height          =   3735
      Left            =   1200
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6588
   End
   Begin Project1.ucFontDialog ucFontDialog1 
      Height          =   3855
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6800
   End
   Begin Project1.ucFileDialog ucFileDialog1 
      Height          =   4935
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8705
   End
   Begin VB.OptionButton opt 
      Caption         =   "Page Setup"
      Height          =   375
      Index           =   5
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.OptionButton opt 
      Caption         =   "Print"
      Height          =   375
      Index           =   4
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.OptionButton opt 
      Caption         =   "Color"
      Height          =   375
      Index           =   3
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.OptionButton opt 
      Caption         =   "Font"
      Height          =   375
      Index           =   2
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.OptionButton opt 
      Caption         =   "Folder"
      Height          =   375
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.OptionButton opt 
      Caption         =   "File"
      Height          =   375
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "fComDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    opt(0).Value = True
End Sub

Private Sub opt_Click(Index As Integer)
ucFileDialog1.Visible = False
ucFontDialog1.Visible = False
ucColorDialog1.Visible = False
ucFolderDialog1.Visible = False
ucPrintDialog1.Visible = False
ucPageSetupDialog1.Visible = False
Select Case Index
Case 0
    ucFileDialog1.Visible = True
Case 1
    ucFolderDialog1.Visible = True
Case 2
    ucFontDialog1.Visible = True
Case 3
    ucColorDialog1.Visible = True
Case 4
    ucPrintDialog1.Visible = True
Case 5
    ucPageSetupDialog1.Visible = True
End Select
End Sub
