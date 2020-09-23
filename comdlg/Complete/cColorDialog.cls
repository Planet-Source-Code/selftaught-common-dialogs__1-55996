VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cColorDialog.cls                    8/25/04
'
'           GENERAL PURPOSE:
'               Display a common color dialog, optionally hooking it to
'               raise events for useful messages.
'
'           LINEAGE:
'               CommonDialogDirect6 from vbaccelerator.com
'
'           DEPENDENCIES:
'               mComClsGeneral.bas
'               mCommonDialog.bas (with the bColorDialog compiler switch turned on)
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


Private mtDialog As tColorDialog

Const ClassName = "cColorDialog"

Public Event DialogInit(ByVal hDlg As Long)
Public Event DialogOK(ByVal hDlg As Long, ByRef bCancel As Boolean)
Public Event DialogClose(ByVal hDlg As Long)
Public Event WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Sub Class_Initialize()                              'init the custom colors
    Dim i As Integer
    For i = 0 To 15
        mtDialog.iColors(i) = QBColor(i)
    Next
End Sub

Private Sub pGetUDT(ByRef tDialog As tColorDialog, _
                    ByVal iColor As OLE_COLOR, _
                    ByVal iFlags As eColorDialogFlag, _
                    ByVal sTitle As String, _
                    ByVal hWndOwner As Long)

    With tDialog
        
        CopyMemory .iColors(0), mtDialog.iColors(0), 64&    'store the color table
        
        If iColor <> NegOneL _
            Then .iColor = iColor _
            Else .iColor = mtDialog.iColor                  'store the color
            
        If iFlags <> NegOneL _
            Then tDialog.iFlags = iFlags _
            Else tDialog.iFlags = mtDialog.iFlags           'store the flags
        
        If hWndOwner <> NegOneL _
            Then tDialog.hWndOwner = hWndOwner _
            Else tDialog.hWndOwner = mtDialog.hWndOwner     'store the owner
        
        #If bHookDialogs Then
            Set .oEventSink = Me
        #End If
    End With
End Sub

Public Function Show( _
                Optional ByRef iColor As OLE_COLOR = NegOneL, _
                Optional ByVal iFlags As eColorDialogFlag = NegOneL, _
                Optional ByVal hWndOwner As Long = NegOneL, _
                Optional ByVal sTitle As String, _
                Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                    As Boolean
    
    Dim ltDialog As tColorDialog
    pGetUDT ltDialog, iColor, iFlags, sTitle, hWndOwner         'get a new udt with default or provided values
    
    Show = mCommonDialog.Color_ShowIndirect(ltDialog)           'show dialog
    
    If Show Then
        CopyMemory mtDialog.iColors(0), ltDialog.iColors(0), 64&    'get the return color table
        iColor = ltDialog.iColor                                    'get the rest of the return values
        mtDialog.iColor = ltDialog.iColor
    End If
    
    iReturnExtendedError = ltDialog.iReturnExtendedError
    mtDialog.iReturnExtendedError = ltDialog.iReturnExtendedError
    
    If CBool(ltDialog.iFlags And dlgColorRaiseError) And Not Show Then
        If mtDialog.iReturnExtendedError _
            Then gErr dlgExtendedError, ClassName _
            Else gErr dlgUserCanceled, ClassName
            
    End If
    
End Function

Public Sub CenterDialog(ByVal hDlg As Long, ByVal vCenterTo As Variant)
    CenterWindow hDlg, vCenterTo, ClassName
End Sub

Public Property Get Color() As OLE_COLOR
    Color = mtDialog.iColor
End Property
Public Property Let Color(ByVal iNew As OLE_COLOR)
    mtDialog.iColor = iNew
End Property

Public Property Get Flags() As eColorDialogFlag
    Flags = mtDialog.iFlags
End Property
Public Property Let Flags(ByVal iNew As eColorDialogFlag)
    mtDialog.iFlags = iNew
End Property
Public Sub SetFlags(Optional ByVal iFlagsOr As eColorDialogFlag, Optional ByVal iFlagsAndNot As eColorDialogFlag)
    mtDialog.iFlags = (mtDialog.iFlags Or iFlagsOr) And Not iFlagsAndNot
End Sub

Public Property Get hWndOwner() As Long
    hWndOwner = mtDialog.hWndOwner
End Property
Public Property Let hWndOwner(ByVal iNew As Long)
    mtDialog.hWndOwner = iNew
End Property

Public Property Get ReturnExtendedError() As eComDlgExtendedError
    ReturnExtendedError = mtDialog.iReturnExtendedError
End Property

#If bHookDialogs Then

    Private Sub iComDlgHook_Proc(lReturn As Long, ByVal hDlg As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
        'callback wndproc for the hook
        
        Const WM_INITDIALOG As Long = &H110
        Const WM_COMMAND As Long = &H111
        Const WM_DESTROY As Long = &H2
        Const WM_NOTIFY As Long = &H4E
        
        Select Case iMsg
        Case WM_INITDIALOG
            RaiseEvent DialogInit(hDlg)
            lReturn = ZeroL
        Case Color_OKMsg
            Dim lbCancel As Boolean
            RaiseEvent DialogOK(hDlg, lbCancel)
            lReturn = Abs(lbCancel)
        Case WM_COMMAND
            RaiseEvent WMCommand(hDlg, wParam, lParam)
            lReturn = ZeroL
        Case WM_DESTROY
            RaiseEvent DialogClose(hDlg)
            lReturn = ZeroL
        End Select
    End Sub

#End If
