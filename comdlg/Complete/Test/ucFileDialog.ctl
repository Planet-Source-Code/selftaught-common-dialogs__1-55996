VERSION 5.00
Begin VB.UserControl ucFileDialog 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   ScaleHeight     =   4950
   ScaleWidth      =   8565
   Begin VB.CheckBox chk 
      Caption         =   "Confirm Selection (when hooked)"
      Height          =   255
      Left            =   3120
      TabIndex        =   31
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   480
      Width           =   2655
   End
   Begin VB.ListBox lst 
      Height          =   3765
      Index           =   2
      ItemData        =   "ucFileDialog.ctx":0000
      Left            =   6000
      List            =   "ucFileDialog.ctx":0002
      TabIndex        =   25
      Top             =   960
      Width           =   2415
   End
   Begin VB.OptionButton opt 
      Caption         =   "Save"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   24
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton opt 
      Caption         =   "Open"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   23
      Top             =   840
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.ListBox lst 
      Height          =   840
      Index           =   1
      ItemData        =   "ucFileDialog.ctx":0004
      Left            =   3120
      List            =   "ucFileDialog.ctx":0006
      TabIndex        =   16
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Dialog"
      Height          =   615
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Text            =   "Choose a File"
      Top             =   4440
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Text            =   "ChooseThisFile.vbp"
      Top             =   3240
      Width           =   2655
   End
   Begin VB.ListBox lst 
      Height          =   1410
      Index           =   0
      ItemData        =   "ucFileDialog.ctx":0008
      Left            =   3120
      List            =   "ucFileDialog.ctx":0068
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Text            =   "0"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Text            =   "All Files (*.*)|*.*|VB Projects (*.vbp)|*.vbp"
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Text            =   "vbp"
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label lbl 
      Caption         =   "Level 2: +dlgFileExplorerStyle"
      Height          =   255
      Index           =   19
      Left            =   6000
      TabIndex        =   30
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lbl 
      Caption         =   "Level 1: +dlgFileEnableHook"
      Height          =   255
      Index           =   18
      Left            =   6000
      TabIndex        =   29
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "Events:"
      Height          =   255
      Index           =   17
      Left            =   6000
      TabIndex        =   28
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "Folder Returned:"
      Height          =   255
      Index           =   16
      Left            =   3120
      TabIndex        =   27
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   14
      Left            =   4560
      TabIndex        =   22
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   13
      Left            =   4560
      TabIndex        =   21
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   12
      Left            =   4560
      TabIndex        =   20
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Ext. Error:"
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   19
      Top             =   2475
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Filter Index:"
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   18
      Top             =   2205
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Files Returned:"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   17
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Return Flags:"
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Title:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Initial Path:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Default File Name:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Flags:"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Default Index:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Filter String:                      Separator: |"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label lbl 
      Caption         =   "Default Extension:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "ucFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents moDialog As cFileDialog
Attribute moDialog.VB_VarHelpID = -1

Private Sub Command1_Click()
    With moDialog
        .DefaultExt = txt(0).Text
        .Filter = txt(1).Text
        .FilterIndex = Val(txt(2).Text)
        .InitialFile = txt(3).Text
        .InitialPath = txt(4).Text
        .Title = txt(5).Text
        .Flags = OrItemData(lst(0))
        pShowRet False
        
        If opt(0).Value _
            Then pShowRet .ShowOpen() _
            Else pShowRet .ShowSave()
        
    End With
End Sub

Private Sub pShowRet(ByVal bVal As Boolean)
    If bVal Then
        Dim lsFiles() As String
        Dim lsFolder As String
        Dim i As Long
        
        For i = 0& To moDialog.GetMultiFileNames(lsFolder, lsFiles) - 1&
            lst(1).AddItem lsFiles(i)
        Next
        txt(6).Text = lsFolder
        lbl(12).Caption = moDialog.ReturnFlags
        lbl(13).Caption = moDialog.ReturnFilterIndex
        lbl(14).Caption = moDialog.ReturnExtendedError
    Else
        lst(1).Clear
        txt(6).Text = vbNullString
        lbl(12).Caption = vbNullString
        lbl(13).Caption = vbNullString
        lbl(14).Caption = vbNullString
    End If
End Sub

Private Sub moDialog_DialogClose(ByVal hDlg As Long)
    pIndicate "Dialog Close"
End Sub

Private Sub moDialog_DialogInit(ByVal hDlg As Long)
    pIndicate "Dialog Init"
End Sub

Private Sub moDialog_DialogOK(ByVal hDlg As Long, bCancel As Boolean)
    If chk.Value Then
        bCancel = MsgBox("Are you sure you want to choose this file?", vbYesNo) = vbNo
    End If
End Sub

Private Sub moDialog_FileChange(ByVal hDlg As Long)
    pIndicate "FileChange"
End Sub

Private Sub moDialog_FolderChange(ByVal hDlg As Long)
    pIndicate "FolderChange"
End Sub

Private Sub moDialog_HelpClicked(ByVal hDlg As Long)
    pIndicate "HelpClicked"
End Sub

Private Sub moDialog_TypeChange(ByVal hDlg As Long)
    pIndicate "TypeChange"
End Sub

Private Sub moDialog_WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)
    pIndicate "WMCommand" & wParam & ", " & lParam
End Sub

Private Sub pIndicate(ByRef s As String)
    lst(2).AddItem s
    lst(2).ListIndex = lst(2).NewIndex
End Sub

Private Sub UserControl_Initialize()
    Set moDialog = New cFileDialog
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    moDialog.hWndOwner = UserControl.Parent.hWnd
    On Error GoTo 0
End Sub
