VERSION 5.00
Begin VB.UserControl ucPageSetupDialog 
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   ScaleHeight     =   4575
   ScaleWidth      =   7065
   Begin VB.ListBox lst 
      Height          =   1410
      Index           =   0
      ItemData        =   "ucPageSetupDialog.ctx":0000
      Left            =   120
      List            =   "ucPageSetupDialog.ctx":003B
      Style           =   1  'Checkbox
      TabIndex        =   24
      Top             =   3000
      Width           =   2655
   End
   Begin VB.ListBox lst 
      Height          =   1425
      Index           =   1
      ItemData        =   "ucPageSetupDialog.ctx":0126
      Left            =   4320
      List            =   "ucPageSetupDialog.ctx":0128
      TabIndex        =   21
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Show Dialog"
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   4
      Left            =   3600
      TabIndex        =   10
      Text            =   "0.25"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   5
      Left            =   3600
      TabIndex        =   9
      Text            =   "0.25"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   3600
      TabIndex        =   8
      Text            =   "0.25"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   7
      Left            =   3600
      TabIndex        =   7
      Text            =   "0.25"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   8
      Left            =   5940
      TabIndex        =   6
      Text            =   "0.10"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   9
      Left            =   5940
      TabIndex        =   5
      Text            =   "0.10"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   10
      Left            =   5940
      TabIndex        =   4
      Text            =   "0.10"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   11
      Left            =   5940
      TabIndex        =   3
      Text            =   "0.10"
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      ItemData        =   "ucPageSetupDialog.ctx":012A
      Left            =   3600
      List            =   "ucPageSetupDialog.ctx":0315
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.OptionButton opt 
      Caption         =   "Landscape"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.OptionButton opt 
      Caption         =   "Portrait"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Flags:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lbl 
      Caption         =   "Events:"
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   22
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Top Margin:"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   20
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Left Margin:"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   19
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Right Margin:"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   18
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Bottom Margin:"
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   17
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Min. Top Margin:"
      Height          =   255
      Index           =   7
      Left            =   4620
      TabIndex        =   16
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Min. Left Margin:"
      Height          =   255
      Index           =   8
      Left            =   4620
      TabIndex        =   15
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Min. Right Margin:"
      Height          =   255
      Index           =   9
      Left            =   4500
      TabIndex        =   14
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Min. Bottom Margin:"
      Height          =   255
      Index           =   10
      Left            =   4380
      TabIndex        =   13
      Top             =   2160
      Width           =   1515
   End
   Begin VB.Label lblInfo 
      Caption         =   "Paper:"
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   12
      Top             =   600
      Width           =   675
   End
End
Attribute VB_Name = "ucPageSetupDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents moPageSetupDialog As cPageSetupDialog
Attribute moPageSetupDialog.VB_VarHelpID = -1

Private Sub cmd_Click(Index As Integer)
    Dim fLeftMargin As Single
    Dim fMinLeftMargin As Single
    Dim fRightMargin As Single
    Dim fMinRightMargin As Single
    Dim fTopMargin As Single
    Dim fMinTopMargin As Single
    Dim fBottomMargin As Single
    Dim fMinBottomMargin As Single
    
    With moPageSetupDialog.DeviceMode
        If cmb.ListIndex > -1& Then
            .PaperSize = cmb.ItemData(cmb.ListIndex)
        Else
            .PaperSize = dlgPrintPaperSizeLetter
        End If
        .Orientation = IIf(opt(1).Value, dlgPrintPortrait, dlgPrintLandscape)
        .Fields = dmfOrientation Or dmfPaperSize
    End With
    
    fTopMargin = CSng(txt(4).Text)
    fLeftMargin = CSng(txt(5).Text)
    fRightMargin = CSng(txt(6).Text)
    fBottomMargin = CSng(txt(7).Text)
    
    fMinTopMargin = CSng(txt(8).Text)
    fMinLeftMargin = CSng(txt(9).Text)
    fMinRightMargin = CSng(txt(10).Text)
    fMinBottomMargin = CSng(txt(11).Text)
    
    With moPageSetupDialog
        If .Show(OrItemData(lst(0)), dlgPrintInches, fLeftMargin, fMinLeftMargin, fRightMargin, fMinRightMargin, fTopMargin, fMinTopMargin, fBottomMargin, fMinBottomMargin) Then
            Const ThisFormat = "0.00"
            With txt
                .Item(4).Text = Format$(fTopMargin, ThisFormat)
                .Item(5).Text = Format$(fLeftMargin, ThisFormat)
                .Item(6).Text = Format$(fRightMargin, ThisFormat)
                .Item(7).Text = Format$(fBottomMargin, ThisFormat)
                
                .Item(8).Text = Format$(fMinTopMargin, ThisFormat)
                .Item(9).Text = Format$(fMinLeftMargin, ThisFormat)
                .Item(10).Text = Format$(fMinRightMargin, ThisFormat)
                .Item(11).Text = Format$(fMinBottomMargin, ThisFormat)
            End With
            Dim i As Long
            Dim j As Long
            j = .DeviceMode.PaperSize
            For i = 0 To cmb.ListCount - 1
                If cmb.ItemData(i) = j Then
                    cmb.ListIndex = i
                    Exit For
                End If
            Next
            If i = cmb.ListCount Then cmb.ListIndex = -1
            If .DeviceMode.Orientation = dlgPrintLandscape Then opt(0).Value = True Else opt(1).Value = True
        End If
    End With

End Sub

Private Sub moPageSetupDialog_DialogClose()
    pIndicate "Close"
End Sub

Private Sub moPageSetupDialog_DialogInit(ByVal hDlg As Long)
    pIndicate "Init"
End Sub

Private Sub moPageSetupDialog_WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)
    pIndicate "WMCommand" & wParam & ", " & lParam
End Sub

Private Sub pIndicate(ByRef s As String)
    lst(1).AddItem s
    lst(1).ListIndex = lst(1).NewIndex
End Sub

Private Sub UserControl_Initialize()
    Set moPageSetupDialog = New cPageSetupDialog
    cmb.ListIndex = 0

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    moPageSetupDialog.hWndOwner = UserControl.Parent.hWnd
    On Error GoTo 0
End Sub


