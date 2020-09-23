VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPageSetupDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cPageSetupDialog.cls                  8/25/04
'
'           GENERAL PURPOSE:
'               Show a page setup dialog, optionally hooking it and raising event for useful messages.
'
'           LINEAGE:
'               CommonDialogDirect6 from www.vbaccelerator.com
'
'           DEPENDENCIES:
'               mCommonDialog.bas (with the bPageSetupDialog compiler switch turned on)
'
'               cDeviceMode.cls
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

Private mtDialog As tPageSetupDialog

Public Event DialogInit(ByVal hDlg As Long)
'Public Event DialogOK(ByRef bCancel As Boolean)
Public Event DialogClose()
Public Event WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)

Const ClassName = "cPageSetupDialog"

Public Function Show( _
            Optional ByVal iFlags As ePrintPageSetup = NegOneL, _
            Optional ByVal iUnits As ePrintPageSetupUnits = NegOneL, _
            Optional ByRef fLeftMargin As Single = NegOneF, _
            Optional ByVal fMinLeftMargin As Single = NegOneF, _
            Optional ByRef fRightMargin As Single = NegOneF, _
            Optional ByVal fMinRightMargin As Single = NegOneF, _
            Optional ByRef fTopMargin As Single = NegOneF, _
            Optional ByVal fMinTopMargin As Single = NegOneF, _
            Optional ByRef fBottomMargin As Single = NegOneF, _
            Optional ByVal fMinBottomMargin As Single = NegOneF, _
            Optional ByRef oDeviceMode As cDeviceMode, _
            Optional ByVal hWndOwner As Long = NegOneL, _
            Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                As Boolean

    Dim ltDialog As tPageSetupDialog
    pGetUDT ltDialog, iFlags, iUnits, fLeftMargin, fMinLeftMargin, fRightMargin, fMinRightMargin, fTopMargin, fMinTopMargin, fBottomMargin, fMinBottomMargin, oDeviceMode, hWndOwner
    
    Show = mCommonDialog.PageSetup_ShowIndirect(ltDialog)
    
    If Show Then
        With ltDialog
            fLeftMargin = .fLeftMargin
            fRightMargin = .fRightMargin
            fTopMargin = .fTopMargin
            fBottomMargin = .fBottomMargin
            Set oDeviceMode = .oDeviceMode
            iReturnExtendedError = .iReturnExtendedError
        End With
    Else
        If CBool(ltDialog.iFlags And dlgPPSRaiseError) Then
            If ltDialog.iReturnExtendedError _
                Then gErr dlgExtendedError, ClassName _
                Else gErr dlgUserCanceled, ClassName
        End If
    End If
    
End Function

Private Sub pGetUDT( _
            ByRef tDialog As tPageSetupDialog, _
            ByVal iFlags As ePrintPageSetup, _
            ByVal iUnits As ePrintPageSetupUnits, _
            ByVal fLeftMargin As Single, _
            ByVal fMinLeftMargin As Single, _
            ByVal fRightMargin As Single, _
            ByVal fMinRightMargin As Single, _
            ByVal fTopMargin As Single, _
            ByVal fMinTopMargin As Single, _
            ByVal fBottomMargin As Single, _
            ByVal fMinBottomMargin As Single, _
            ByVal oDeviceMode As cDeviceMode, _
            ByVal hWndOwner As Long)
    
    With tDialog
        
        If fLeftMargin <> NegOneF _
            Then .fLeftMargin = fLeftMargin _
            Else .fLeftMargin = mtDialog.fLeftMargin
        
        If fMinLeftMargin <> NegOneF _
            Then .fMinLeftMargin = fMinLeftMargin _
            Else .fMinLeftMargin = mtDialog.fMinLeftMargin
        
        If fRightMargin <> NegOneF _
            Then .fRightMargin = fRightMargin _
            Else .fRightMargin = mtDialog.fRightMargin
        
        If fMinRightMargin <> NegOneF _
            Then .fMinRightMargin = fMinRightMargin _
            Else .fMinRightMargin = mtDialog.fMinRightMargin
        
        If fTopMargin <> NegOneF _
            Then .fTopMargin = fTopMargin _
            Else .fTopMargin = mtDialog.fTopMargin
        
        If fMinTopMargin <> NegOneF _
            Then .fMinTopMargin = fMinTopMargin _
            Else .fMinTopMargin = mtDialog.fMinTopMargin
        
        If fBottomMargin <> NegOneF _
            Then .fBottomMargin = fBottomMargin _
            Else .fBottomMargin = mtDialog.fBottomMargin
        
        If fMinBottomMargin <> NegOneF _
            Then .fMinBottomMargin = fMinBottomMargin _
            Else .fMinBottomMargin = mtDialog.fMinBottomMargin
        
        If iFlags <> NegOneL _
            Then .iFlags = iFlags _
            Else .iFlags = mtDialog.iFlags
                
        If iUnits <> iUnits _
            Then .iUnits = iUnits _
            Else .iUnits = mtDialog.iUnits
        
        If hWndOwner <> NegOneL _
            Then .hWndOwner = hWndOwner _
            Else .hWndOwner = mtDialog.hWndOwner
        
        If oDeviceMode Is Nothing _
            Then Set .oDeviceMode = mtDialog.oDeviceMode _
            Else Set .oDeviceMode = oDeviceMode
        
        #If bHookDialogs Then
            Set .oEventSink = Me
        #End If
    End With
    
End Sub

Public Property Get hWndOwner() As Long
    hWndOwner = mtDialog.hWndOwner
End Property
Public Property Let hWndOwner(ByVal iNew As Long)
    mtDialog.hWndOwner = iNew
End Property

Public Property Get Flags() As ePrintPageSetup
    Flags = mtDialog.iFlags
End Property
Public Property Let Flags(ByVal iNew As ePrintPageSetup)
    mtDialog.iFlags = iNew
End Property
Public Sub SetFlags(ByVal iFlagsOr As ePrintPageSetup, ByVal iFlagsAndNot As ePrintPageSetup)
    mtDialog.iFlags = (mtDialog.iFlags Or iFlagsOr) And Not iFlagsAndNot
End Sub


Public Property Get LeftMargin() As Single
    LeftMargin = mtDialog.fLeftMargin
End Property
Public Property Let LeftMargin(ByRef rNew As Single)
    mtDialog.fLeftMargin = rNew
End Property

Public Property Get MinLeftMargin() As Single
    MinLeftMargin = mtDialog.fMinLeftMargin
End Property
Public Property Let MinLeftMargin(ByRef rNew As Single)
    mtDialog.fMinLeftMargin = rNew
End Property

Public Property Get RightMargin() As Single
    RightMargin = mtDialog.fRightMargin
End Property
Public Property Let RightMargin(ByRef rNew As Single)
    mtDialog.fRightMargin = rNew
End Property

Public Property Get MinRightMargin() As Single
    MinRightMargin = mtDialog.fMinRightMargin
End Property
Public Property Let MinRightMargin(ByRef rNew As Single)
    mtDialog.fMinRightMargin = rNew
End Property

Public Property Get TopMargin() As Single
    TopMargin = mtDialog.fTopMargin
End Property
Public Property Let TopMargin(ByRef rNew As Single)
    mtDialog.fTopMargin = rNew
End Property

Public Property Get MinTopMargin() As Single
    MinTopMargin = mtDialog.fMinTopMargin
End Property
Public Property Let MinTopMargin(ByRef rNew As Single)
    mtDialog.fMinTopMargin = rNew
End Property

Public Property Get BottomMargin() As Single
    BottomMargin = mtDialog.fBottomMargin
End Property
Public Property Let BottomMargin(ByRef rNew As Single)
    mtDialog.fBottomMargin = rNew
End Property

Public Property Get MinBottomMargin() As Single
    MinBottomMargin = mtDialog.fMinBottomMargin
End Property
Public Property Let MinBottomMargin(ByRef rNew As Single)
    mtDialog.fMinBottomMargin = rNew
End Property

Public Property Get Units() As ePrintPageSetupUnits
    Units = mtDialog.iUnits
End Property
Public Property Let Units(ByVal iNew As ePrintPageSetupUnits)
    mtDialog.iUnits = Units
End Property

Public Property Get DeviceMode() As cDeviceMode
    Set DeviceMode = mtDialog.oDeviceMode
End Property

Private Sub Class_Initialize()
    Set mtDialog.oDeviceMode = New cDeviceMode
End Sub

Public Sub CenterDialog(ByVal hDlg As Long, ByVal vCenterTo As Variant)
    CenterWindow hDlg, vCenterTo, "cPrintDialog"
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
            RaiseEvent DialogClose
            lReturn = ZeroL
        End Select
    End Sub
#End If
