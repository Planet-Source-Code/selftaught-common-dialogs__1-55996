VERSION 5.00
Begin VB.UserControl ucColorDialog 
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   3705
   ScaleWidth      =   6765
   Begin VB.ListBox lst 
      Height          =   1410
      Index           =   0
      ItemData        =   "ucColorDialog.ctx":0000
      Left            =   120
      List            =   "ucColorDialog.ctx":0018
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
   End
   Begin VB.ListBox lst 
      Height          =   1425
      Index           =   1
      ItemData        =   "ucColorDialog.ctx":0077
      Left            =   3720
      List            =   "ucColorDialog.ctx":0079
      TabIndex        =   3
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Show Dialog"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox pic 
      Height          =   555
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   1080
      Width           =   6555
   End
   Begin VB.Label lbl 
      Caption         =   "Flags:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Events: (Must have dlgColorEnableHook)"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblInfo 
      Caption         =   "Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6555
   End
End
Attribute VB_Name = "ucColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents moColorDialog As cColorDialog
Attribute moColorDialog.VB_VarHelpID = -1

Private Sub cmd_Click()
    moColorDialog.Flags = OrItemData(lst(0))
    If moColorDialog.Show() Then pic.BackColor = moColorDialog.Color
End Sub

Private Sub moColorDialog_DialogClose(ByVal hDlg As Long)
    pIndicate "DialogClose"
End Sub

Private Sub moColorDialog_DialogInit(ByVal hDlg As Long)
    pIndicate "DialogInit"
End Sub

Private Sub moColorDialog_DialogOK(ByVal hDlg As Long, bCancel As Boolean)
    bCancel = MsgBox("Are you sure you want to choose this color?", vbYesNo) = vbNo
End Sub

Private Sub moColorDialog_WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)
    pIndicate "WMCommand" & wParam & ", " & lParam
End Sub

Private Sub pIndicate(ByRef s As String)
    lst(1).AddItem s
    lst(1).ListIndex = lst(1).NewIndex
End Sub

Private Sub UserControl_Initialize()
    Set moColorDialog = New cColorDialog
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    moColorDialog.hWndOwner = UserControl.Parent.hWnd
    moColorDialog.Color = pic.BackColor
    On Error GoTo 0
End Sub
