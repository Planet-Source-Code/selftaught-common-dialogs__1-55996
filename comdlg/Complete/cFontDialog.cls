VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cFontDialog.cls                  8/25/04
'
'           GENERAL PURPOSE:
'               Show a font dialog, optionally hooking it and raising events for useful messages.
'
'           LINEAGE:
'               CommonDialogDirect6 from www.vbaccelerator.com
'
'           DEPENDENCIES:
'               mCommonDialog.bas (with the bFontDialog compiler switch turned on)
'
'               if the compiler switch bHookDialogs is set, iComDlgHook.cls
'               if the compiler switch is not set, comment out the implements statement.
'               see the implements statement for info
'
'==================================================================================================

#Const bHookDialogs = True

'Due to a bug in VB, you will (generally ???) have to comment out the implements
'statement if you change the bHookDialogs compiler switch to False and remove
'the iComDlgHook class from the project, even if you have the implements statement
'in a conditional compiler block statement. Actually, if you leave the conditional
'statement in, it will  only stop VB from selecting the implements line upon the
'compile error, just giving you a messagebox saying "User-defined type not defined".
'#If bHookDialogs Then
    Implements iComDlgHook
'#End If

Private mtDialog As tFontDialog

Public Event DialogInit(ByVal hDlg As Long)
Public Event DialogClose(ByVal hDlg As Long)
Public Event WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long

Const ClassName = "cFontDialog"

'Caps!
#Const Flags = 0

Public Function Show( _
                Optional ByVal oFont As Object, _
                Optional ByVal iFlags As eFontDialog = NegOneL, _
                Optional ByVal hdc As Long = NegOneL, _
                Optional ByVal hWndOwner As Long = NegOneL, _
                Optional ByVal iMinSize As Long = NegOneL, _
                Optional ByVal iMaxSize As Long = NegOneL, _
                Optional ByRef iColor As OLE_COLOR, _
                Optional ByRef iReturnFlags As eFontDialog, _
                Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                    As Boolean
                    
    Dim ltDialog As tFontDialog
    
    pGetUDT ltDialog, oFont, iFlags, hdc, hWndOwner, iMinSize, iMaxSize, iColor
    
    Show = mCommonDialog.Font_ShowIndirect(ltDialog)
    
    If Show Then
        mtDialog.iColor = ltDialog.iColor
        mtDialog.iReturnFlags = ltDialog.iReturnFlags
        
        iColor = ltDialog.iColor
        iReturnFlags = ltDialog.iReturnFlags
    End If
    
    Set mtDialog.oFont = ltDialog.oFont
    mtDialog.iReturnExtendedError = ltDialog.iReturnExtendedError
    iReturnExtendedError = ltDialog.iReturnExtendedError
    
    If Not Show Then
        If CBool(ltDialog.iFlags And dlgFontRaiseError) Then
            If ltDialog.iReturnExtendedError _
                Then gErr dlgExtendedError, ClassName _
                Else gErr dlgUserCanceled, ClassName
        End If
    End If
    
End Function

Private Sub pGetUDT( _
            ByRef tUDT As tFontDialog, _
            ByVal oFont As Object, _
            ByVal iFlags As eFontDialog, _
            ByVal hdc As Long, _
            ByVal hWndOwner As Long, _
            ByVal iMinSize As Long, _
            ByVal iMaxSize As Long, _
            ByVal iColor As Long)
    
    With tUDT
        If oFont Is Nothing _
            Then Set .oFont = mtDialog.oFont _
            Else Set .oFont = oFont
        
        If iFlags <> NegOneL _
            Then .iFlags = iFlags _
            Else .iFlags = mtDialog.iFlags
            
        If hdc <> NegOneL _
            Then .hdc = hdc _
            Else .hdc = mtDialog.hdc
            
        If hWndOwner <> NegOneL _
            Then .hWndOwner = hWndOwner _
            Else .hWndOwner = mtDialog.hWndOwner
            
        If iMinSize <> NegOneL _
            Then .iMinSize = iMinSize _
            Else .iMinSize = mtDialog.iMinSize
            
        If iMaxSize <> NegOneL _
            Then .iMaxSize = iMaxSize _
            Else .iMaxSize = mtDialog.iMaxSize
            
        If iColor <> NegOneL _
            Then OleTranslateColor iColor, ZeroL, .iColor _
            Else .iColor = mtDialog.iColor
        
        #If bHookDialogs Then
            Set .oEventSink = Me
        #End If
    End With
End Sub

Public Property Get Font() As Object
    Set Font = mtDialog.oFont
End Property

Public Property Set Font(ByVal oNew As Object)
    Set mtDialog.oFont = oNew
End Property

Public Property Get Flags() As eFontDialog
    Flags = mtDialog.iFlags
End Property
Public Property Let Flags(ByVal iNew As eFontDialog)
    mtDialog.iFlags = iNew
End Property
Public Sub SetFlags(Optional ByVal iFlagsOr As eFontDialog, Optional ByVal iFlagsAndNot As eFontDialog)
    mtDialog.iFlags = (mtDialog.iFlags Or iFlagsOr) And Not iFlagsAndNot
End Sub

Public Property Get hdc() As Long
    hdc = mtDialog.hdc
End Property
Public Property Let hdc(ByVal iNew As Long)
    mtDialog.hdc = iNew
End Property

Public Property Get hWndOwner() As Long
    hWndOwner = mtDialog.hWndOwner
End Property
Public Property Let hWndOwner(ByVal iNew As Long)
    mtDialog.hWndOwner = iNew
End Property

Public Property Get FontColor() As Long
    FontColor = mtDialog.iColor
End Property
Public Property Let FontColor(ByVal iNew As Long)
    mtDialog.iColor = iNew
End Property

Public Property Get MinSize() As Long
    MinSize = mtDialog.iMinSize
End Property
Public Property Let MinSize(ByVal iNew As Long)
    mtDialog.iMinSize = iNew
End Property

Public Property Get MaxSize() As Long
    MaxSize = mtDialog.iMaxSize
End Property
Public Property Let MaxSize(ByVal iNew As Long)
    mtDialog.iMaxSize = iNew
End Property

Public Property Get ReturnFlags() As eFontDialog
    ReturnFlags = mtDialog.iReturnFlags
End Property

Public Property Get ExtendedError() As eComDlgExtendedError
    ExtendedError = mtDialog.iReturnExtendedError
End Property

Public Sub CenterDialog(ByVal hDlg As Long, ByVal vCenterTo As Variant)
    CenterWindow hDlg, vCenterTo, "cFontDialog"
End Sub

#If bHookDialogs Then
    Private Sub iComDlgHook_Proc(lReturn As Long, ByVal hDlg As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
        Const WM_INITDIALOG As Long = &H110
        Const WM_COMMAND As Long = &H111
        Const WM_DESTROY As Long = &H2
        
        Select Case iMsg
        Case WM_INITDIALOG
            RaiseEvent DialogInit(hDlg)
            lReturn = ZeroL
        Case WM_COMMAND
            RaiseEvent WMCommand(hDlg, wParam, lParam)
            lReturn = ZeroL
        Case WM_DESTROY
            RaiseEvent DialogClose(hDlg)
            lReturn = ZeroL
        End Select
    End Sub
#End If
