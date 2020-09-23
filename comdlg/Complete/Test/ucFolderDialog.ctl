VERSION 5.00
Begin VB.UserControl ucFolderDialog 
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   ScaleHeight     =   4380
   ScaleWidth      =   8055
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   3
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   4575
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Text            =   "Choose a folder"
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Show Dialog"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
   Begin VB.ListBox lst 
      Height          =   1410
      Index           =   0
      ItemData        =   "ucFolderDialog.ctx":0000
      Left            =   240
      List            =   "ucFolderDialog.ctx":003A
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   2760
      Width           =   2655
   End
   Begin VB.ListBox lst 
      Height          =   1425
      Index           =   1
      ItemData        =   "ucFolderDialog.ctx":0146
      Left            =   5280
      List            =   "ucFolderDialog.ctx":0148
      TabIndex        =   0
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lbl 
      Caption         =   "Events:"
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   12
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label lbl 
      Caption         =   "Flags:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label lbl 
      Caption         =   "Return Path:"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Initial Folder:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Root Folder:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Title:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "ucFolderDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents moFolderDialog As cFolderDialog
Attribute moFolderDialog.VB_VarHelpID = -1

Private Sub cmd_Click()
    moFolderDialog.Flags = OrItemData(lst(0))
    moFolderDialog.Title = txt(0).Text
    moFolderDialog.RootPath = txt(1).Text
    moFolderDialog.InitialPath = txt(2).Text
    
    If moFolderDialog.Show() Then txt(3).Text = moFolderDialog.Path
    
End Sub

Private Sub moFolderDialog_DialogInit(ByVal hDlg As Long)
    pIndicate "DialogInit"
End Sub

Private Sub moFolderDialog_FolderChanged(ByVal hDlg As Long, ByVal sPath As String, bCancel As Boolean)
    pIndicate "FolderChanged"
End Sub

Private Sub moFolderDialog_ValidationFailed(ByVal hDlg As Long, ByVal sPath As String, bKeepOpen As Boolean)
    pIndicate "Validation failed"
End Sub

Private Sub pIndicate(ByRef s As String)
    lst(1).AddItem s
    lst(1).ListIndex = lst(1).NewIndex
End Sub

Private Sub UserControl_Initialize()
    Set moFolderDialog = New cFolderDialog
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    moFolderDialog.hWndOwner = UserControl.Parent.hWnd
    On Error GoTo 0
End Sub
