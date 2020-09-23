VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cPrintDialog.cls                  8/25/04
'
'           GENERAL PURPOSE:
'               Show a print dialog, optionally hooking it and raising event for useful messages.
'
'           LINEAGE:
'               CommonDialogDirect6 from www.vbaccelerator.com
'
'           DEPENDENCIES:
'               mCommonDialog.bas (with the bPrintDialog compiler switch turned on)
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

Private mtDialog As tPrintDialog

Public Event DialogInit(ByVal hDlg As Long)
'Public Event DialogOK(ByVal hDlg as long, ByRef bCancel As Boolean)
Public Event DialogClose(ByVal hDlg As Long)
Public Event WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)

Const ClassName As String = "cPrintDialog"

Public Function Show( _
                Optional ByRef hdc As Long, _
                Optional ByVal iFlags As ePrintDialog = NegOneL, _
                Optional ByRef iRange As ePrintRange = NegOneL, _
                Optional ByRef iFromPage As Long = NegOneL, _
                Optional ByRef iToPage As Long = NegOneL, _
                Optional ByVal iMinPage As Long = NegOneL, _
                Optional ByVal iMaxPage As Long = NegOneL, _
                Optional ByRef oDeviceMode As cDeviceMode, _
                Optional ByRef sDevice As String, _
                Optional ByRef sDriver As String, _
                Optional ByRef sOutputPort As String, _
                Optional ByRef bCollate As Boolean, _
                Optional ByRef bPrintToFile As Boolean, _
                Optional ByVal hWndOwner As Long = NegOneL, _
                Optional ByRef iReturnFlags As ePrintDialog, _
                Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                    As Boolean

    
    Dim ltDialog As tPrintDialog
    
    pGetUDT ltDialog, iFlags, iRange, iFromPage, iToPage, iMinPage, iMaxPage, oDeviceMode, hWndOwner
    
    Show = mCommonDialog.Print_ShowIndirect(ltDialog)
    
    If Show Then
        With ltDialog
            hdc = .hdc
            iRange = .iRange
            iFromPage = .iFromPage
            iToPage = .iToPage
            iReturnFlags = .iReturnFlags
            iReturnExtendedError = .iReturnExtendedError
            bCollate = .bCollate
            bPrintToFile = .bPrintToFile
            sDevice = .sDevice
            sDriver = .sDriver
            sOutputPort = .sOutputPort
            Set oDeviceMode = .oDeviceMode
        End With
    Else
        If CBool(ltDialog.iFlags And dlgPrintRaiseError) Then
            If ltDialog.iReturnExtendedError _
                Then gErr dlgExtendedError, ClassName _
                Else gErr dlgUserCanceled, ClassName
        End If
    End If
    
End Function

Private Sub pGetUDT( _
            ByRef tUDT As tPrintDialog, _
            ByVal iFlags As ePrintDialog, _
            ByVal iRange As ePrintRange, _
            ByVal iFromPage As Long, _
            ByVal iToPage As Long, _
            ByVal iMinPage As Long, _
            ByVal iMaxPage As Long, _
            ByVal oDeviceMode As cDeviceMode, _
            ByVal hWndOwner As Long)
    
    With tUDT
        
        If iFlags <> NegOneL _
            Then .iFlags = iFlags _
            Else .iFlags = mtDialog.iFlags
        
        If iRange <> NegOneL _
            Then .iRange = iRange _
            Else .iRange = mtDialog.iRange
            
        If iFromPage <> NegOneL _
            Then .iFromPage = iFromPage _
            Else .iFromPage = mtDialog.iFromPage
        
        If iToPage <> NegOneL _
            Then .iToPage = iToPage _
            Else .iToPage = mtDialog.iToPage
            
        If iMinPage <> NegOneL _
            Then .iMinPage = iMinPage _
            Else .iMinPage = mtDialog.iMinPage
            
        If iMaxPage <> NegOneL _
            Then .iMaxPage = iMaxPage _
            Else .iMaxPage = mtDialog.iMaxPage
            
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

Public Property Get Flags() As ePrintDialog
    Flags = mtDialog.iFlags
End Property
Public Property Let Flags(ByVal iNew As ePrintDialog)
    mtDialog.iFlags = iNew
End Property
Public Sub SetFlags(Optional ByVal iFlagsOr As ePrintDialog, Optional ByVal iFlagsAndNot As ePrintDialog)
    mtDialog.iFlags = (mtDialog.iFlags Or iFlagsOr) And Not iFlagsAndNot
End Sub

Public Property Get Range() As ePrintRange
    Range = mtDialog.iRange
End Property
Public Property Let Range(ByVal iNew As ePrintRange)
    mtDialog.iRange = iNew
End Property

Public Property Get FromPage() As Long
    FromPage = mtDialog.iFromPage
End Property
Public Property Let FromPage(ByVal iNew As Long)
    mtDialog.iFromPage = iNew
End Property

Public Property Get ToPage() As Long
    ToPage = mtDialog.iToPage
End Property
Public Property Let ToPage(ByVal iNew As Long)
    mtDialog.iToPage = iNew
End Property

Public Property Get MinPage() As Long
    MinPage = mtDialog.iMinPage
End Property
Public Property Let MinPage(ByVal iNew As Long)
    mtDialog.iMinPage = iNew
End Property

Public Property Get MaxPage() As Long
    MaxPage = mtDialog.iMaxPage
End Property
Public Property Let MaxPage(ByVal iNew As Long)
    mtDialog.iMaxPage = iNew
End Property

Public Property Get ExtendedError() As Long
    ExtendedError = mtDialog.iReturnExtendedError
End Property

Public Property Get Collate() As Boolean
    Collate = mtDialog.bCollate
End Property

Public Property Get PrintToFile() As Boolean
    PrintToFile = mtDialog.bPrintToFile
End Property

Public Property Get DriverName() As String
    DriverName = mtDialog.sDriver
End Property

Public Property Get DeviceName() As String
    DeviceName = mtDialog.sDevice
End Property

Public Property Get OutputPort() As String
    OutputPort = mtDialog.sOutputPort
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
            RaiseEvent DialogClose(hDlg)
            lReturn = ZeroL
        End Select
    End Sub

#End If
