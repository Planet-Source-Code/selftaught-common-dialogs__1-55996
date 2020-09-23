VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   3345
   ClientTop       =   2115
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   4215
   Begin VB.CheckBox Check1 
      Caption         =   "Use ""Indirect"" calls"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Font"
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Color"
      Height          =   495
      Index           =   2
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Save"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Open"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click(Index As Integer)
    If Check1.Value = vbUnchecked Then
        
        Dim lsName As String
        Select Case Index
        Case 0
            If File_ShowOpen(lsName) Then Text1(0).Text = lsName
        Case 1
            If File_ShowSave(lsName) Then Text1(1).Text = lsName
        Case 2
            Dim liColor As OLE_COLOR
            liColor = Command1(Index).BackColor
            If Color_Show(liColor) Then Command1(Index).BackColor = liColor
        Case 3
            Font_Show Command1(Index).Font
        End Select
        
    Else
    
        Dim ltFileDialog As tFileDialog
        Select Case Index
        Case 0 'show open
            With ltFileDialog
                .hWndOwner = hWnd
                .iFilterIndex = 1
                .iFlags = dlgFileExplorerStyle Or dlgFileMustExist Or dlgFilePathMustExist Or dlgFileHideReadOnly
                .sFilter = "All Files (*.*)|*.*"
            End With
            If File_ShowOpenIndirect(ltFileDialog) Then Text1(0).Text = ltFileDialog.sReturnFileName
        Case 1 'show save
            With ltFileDialog
                .hWndOwner = hWnd
                .iFilterIndex = 1
                .iFlags = dlgFileExplorerStyle
                .sFilter = "All Files (*.*)|*.*"
            End With
            If File_ShowSaveIndirect(ltFileDialog) Then Text1(1).Text = ltFileDialog.sReturnFileName
        Case 2 'show color
            Dim ltColorDialog As tColorDialog
            With ltColorDialog
                .hWndOwner = hWnd
                .iColor = Command1(Index).BackColor
                .iFlags = dlgColorAny
                Dim i As Long
                For i = LBound(.iColors) To UBound(.iColors)
                    .iColors(i) = QBColor(i)
                Next
            End With
            If Color_ShowIndirect(ltColorDialog) Then Command1(Index).BackColor = ltColorDialog.iColor
        Case 3 'show font
            Dim ltFontDialog As tFontDialog
            With ltFontDialog
                .hWndOwner = hWnd
                Set .oFont = Command1(Index).Font
                .iMaxSize = 72
                .iMinSize = 6
                .iFlags = dlgFontScreenFonts
            End With
            Font_ShowIndirect ltFontDialog
        End Select
        
    End If
End Sub
