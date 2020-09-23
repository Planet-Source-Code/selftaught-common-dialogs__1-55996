VERSION 5.00
Begin VB.UserControl ucFontDialog 
   ClientHeight    =   3825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   ScaleHeight     =   3825
   ScaleWidth      =   8295
   Begin VB.ListBox lst 
      Height          =   1425
      Index           =   1
      ItemData        =   "ucFontDialog.ctx":0000
      Left            =   5160
      List            =   "ucFontDialog.ctx":0002
      TabIndex        =   5
      Top             =   2280
      Width           =   2655
   End
   Begin VB.ListBox lst 
      Height          =   1410
      Index           =   0
      ItemData        =   "ucFontDialog.ctx":0004
      Left            =   120
      List            =   "ucFontDialog.ctx":0090
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   2280
      Width           =   2655
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   7935
      TabIndex        =   3
      Top             =   960
      Width           =   7995
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Show Dialog"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Text            =   "8"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   1
      Left            =   5760
      TabIndex        =   0
      Text            =   "28"
      Top             =   90
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Max Size:"
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Min Size:"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Events: (Must have dlgFontEnableHook)"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   7
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label lbl 
      Caption         =   "Flags:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "ucFontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents moFontDialog As cFontDialog
Attribute moFontDialog.VB_VarHelpID = -1

Private Sub cmd_Click()
    With moFontDialog
        .MinSize = Val(txt(0).Text)
        .MaxSize = Val(txt(1).Text)
        .Flags = OrItemData(lst(0))
    End With
    
    If moFontDialog.Show(pic.Font) Then
        pic.ForeColor = moFontDialog.FontColor
        pRenderSampleText
    End If
End Sub

Private Sub moFontDialog_DialogClose(ByVal hDlg As Long)
    pIndicate "DialogClose"
End Sub

Private Sub moFontDialog_DialogInit(ByVal hDlg As Long)
    pIndicate "DialogInit"
End Sub

Private Sub moFontDialog_WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)
    pIndicate "WMCommand" & wParam & ", " & lParam
End Sub

Private Sub pIndicate(ByRef s As String)
    lst(1).AddItem s
    lst(1).ListIndex = lst(1).NewIndex
End Sub


Private Sub UserControl_Initialize()
    Set moFontDialog = New cFontDialog
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    moFontDialog.hWndOwner = UserControl.Parent.hWnd
    moFontDialog.FontColor = pic.ForeColor
    moFontDialog.hdc = pic.hdc
    On Error GoTo 0
End Sub

Private Sub pRenderSampleText()
    pic.Cls
    pic.Print "This is sample text."
    pic.Refresh
End Sub

