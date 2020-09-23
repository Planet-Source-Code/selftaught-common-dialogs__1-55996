VERSION 5.00
Begin VB.UserControl ucPrintDialog 
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   ScaleHeight     =   4065
   ScaleWidth      =   5700
   Begin VB.ListBox lst 
      Height          =   1425
      Index           =   1
      ItemData        =   "ucPrintDialog.ctx":0000
      Left            =   2880
      List            =   "ucPrintDialog.ctx":0002
      TabIndex        =   14
      Top             =   2400
      Width           =   2655
   End
   Begin VB.ListBox lst 
      Height          =   1410
      Index           =   0
      ItemData        =   "ucPrintDialog.ctx":0004
      Left            =   0
      List            =   "ucPrintDialog.ctx":005B
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   2400
      Width           =   2655
   End
   Begin VB.OptionButton optPrint 
      Caption         =   "Pages:              From:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Show Dialog"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.OptionButton optPrint 
      Caption         =   "All Pages"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   12
      Left            =   1920
      TabIndex        =   5
      Text            =   "1"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   13
      Left            =   2880
      TabIndex        =   4
      Text            =   "5"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   14
      Left            =   1920
      TabIndex        =   3
      Text            =   "1"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   15
      Left            =   2880
      TabIndex        =   2
      Text            =   "5"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   16
      Left            =   2880
      TabIndex        =   1
      Text            =   "5"
      Top             =   960
      Width           =   615
   End
   Begin VB.OptionButton optPrint 
      Caption         =   "Selection"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "To print a test page, select the dlgPrintReturnDC flag."
      Height          =   495
      Left            =   2520
      TabIndex        =   17
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lbl 
      Caption         =   "Flags:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   16
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label lbl 
      Caption         =   "Events:"
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   15
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label lblInfo 
      Caption         =   "Copies:"
      Height          =   255
      Index           =   12
      Left            =   2880
      TabIndex        =   12
      Top             =   720
      Width           =   555
   End
   Begin VB.Label lblInfo 
      Caption         =   "To:"
      Height          =   255
      Index           =   13
      Left            =   2520
      TabIndex        =   11
      Top             =   1380
      Width           =   315
   End
   Begin VB.Label lblInfo 
      Caption         =   "Min:"
      Height          =   255
      Index           =   14
      Left            =   1560
      TabIndex        =   10
      Top             =   1725
      Width           =   390
   End
   Begin VB.Label lblInfo 
      Caption         =   "Max:"
      Height          =   255
      Index           =   15
      Left            =   2550
      TabIndex        =   9
      Top             =   1710
      Width           =   390
   End
End
Attribute VB_Name = "ucPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As tRect, ByVal wFormat As Long) As Long
Private Const DT_CALCRECT As Long = &H400
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private Type tDocInfo
    cbSize As Long
    lpszDocName As String
    lpszOutput As String
End Type

Private Declare Function StartDoc Lib "gdi32.dll" Alias "StartDocA" (ByVal hdc As Long, lpdi As tDocInfo) As Long
Private Declare Function StartPage Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function EndPage Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function EndDoc Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private WithEvents moPrintDialog As cPrintDialog
Attribute moPrintDialog.VB_VarHelpID = -1

Private Sub cmd_Click()
    Dim iRange As ePrintRange
    Dim liFrom As Long
    Dim liTo As Long
    Dim liCopies As Integer
    Dim hdc As Long
    Dim liFlags As ePrintDialog
    
    If optPrint(0).Value Then
        iRange = dlgPrintRangeAll
    ElseIf optPrint(1).Value Then
        iRange = dlgPrintRangePageNumbers
    Else
        iRange = dlgPrintRangeSelection
    End If
    
    liFrom = txt(12).Text
    liTo = txt(13).Text
    liCopies = txt(16).Text
    
    liFlags = OrItemData(lst(0))
    moPrintDialog.DeviceMode.Copies = liCopies
    If moPrintDialog.Show(hdc, liFlags, iRange, liFrom, liTo, txt(14).Text, txt(15).Text) Then
        liCopies = moPrintDialog.DeviceMode.Copies
        With txt
            .Item(12).Text = liFrom
            .Item(13).Text = liTo
            .Item(16).Text = liCopies
        End With
        
        If iRange = dlgPrintRangeAll Then
            optPrint(0).Value = True
        ElseIf iRange = dlgPrintRangePageNumbers Then
            optPrint(1).Value = True
        Else
            optPrint(2).Value = True
        End If
        
        If CBool(liFlags And dlgPrintReturnDc) Then
            pPrintTest hdc, liCopies, liFrom, liTo
        ElseIf CBool(liFlags And dlgPrintReturnIc) Then
            DeleteDC hdc
        End If
        
    End If
End Sub

Private Sub moPrintDialog_DialogClose(ByVal hDlg As Long)
    pIndicate "DialogClose"
End Sub

Private Sub moPrintDialog_DialogInit(ByVal hDlg As Long)
    pIndicate "DialogInit"
End Sub

Private Sub moPrintDialog_WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)
    pIndicate "WMCOMMAND" & wParam & ", " & lParam
End Sub

Private Sub pIndicate(ByRef s As String)
lst(1).AddItem s
lst(1).ListIndex = lst(1).NewIndex
End Sub

Private Sub UserControl_Initialize()
    Set moPrintDialog = New cPrintDialog
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    moPrintDialog.hWndOwner = UserControl.Parent.hWnd
    On Error GoTo 0
End Sub

Private Sub pPrintTest(ByVal hdc As Long, ByVal iCopies As Long, ByVal iFrom As Long, ByVal iTo As Long)
    If hdc Then
        Dim tDI As tDocInfo
        Dim hJob As Long
        
        tDI.cbSize = Len(tDI)
        tDI.lpszDocName = "Test Print Page"
        
        If StartDoc(hdc, tDI) Then
            StartPage hdc
            pDrawText hdc, "Copies: " & iCopies
            pDrawText hdc, "From Page: " & iFrom
            pDrawText hdc, "To Page: " & iTo
            pDrawText hdc, "To Page: " & iTo
            pDrawText hdc, "Device Name: " & moPrintDialog.DeviceMode.DeviceName
            EndPage hdc
            EndDoc hdc
            
        End If
        
        DeleteDC hdc
    End If
End Sub

Private Sub pDrawText(ByVal hdc As Long, ByVal sText As String)
    
    Static tR As tRect
    tR.Top = tR.Bottom
    DrawText hdc, sText, Len(sText), tR, DT_CALCRECT
    DrawText hdc, sText, Len(sText), tR, 0&

    
End Sub
