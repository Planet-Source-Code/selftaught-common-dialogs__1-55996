VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cFileDialog.cls                  8/25/04
'
'           GENERAL PURPOSE:
'               Show a file dialog, optionally hooking it and raising events for useful messages.
'
'           LINEAGE:
'               CommonDialogDirect6 from www.vbaccelerator.com
'
'           DEPENDENCIES:
'               mCommonDialog.bas  (with the bFileDialog compiler switch on)
'               mComClsGeneral.bas
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

Private mtFileDialog As tFileDialog

Event DialogInit(ByVal hDlg As Long)
Event DialogOK(ByVal hDlg As Long, ByRef bCancel As Boolean)
Event DialogClose(ByVal hDlg As Long)
Event FileChange(ByVal hDlg As Long)
Event FolderChange(ByVal hDlg As Long)
Event TypeChange(ByVal hDlg As Long)
Event HelpClicked(ByVal hDlg As Long)
Event WMCommand(ByVal hDlg As Long, wParam As Long, lParam As Long)

Const ClassName = "cFileDialog"

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Function ShowOpen( _
   Optional ByRef sReturnFileName As String, _
   Optional ByVal iFlags As eFileDialog = NegOneL, _
   Optional ByVal sFilter, _
   Optional ByVal iFilterIndex As Long = NegOneL, _
   Optional ByVal sDefaultExt, _
   Optional ByVal sInitPath, _
   Optional ByVal sInitFile, _
   Optional ByVal sTitle, _
   Optional ByVal hWndOwner As Long = NegOneL, _
   Optional ByVal vTemplate, _
   Optional ByVal hInstance As Long, _
   Optional ByRef iReturnValue As Long, _
   Optional ByRef iReturnFlags As Long, _
   Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
   Optional ByRef iReturnFilterIndex As Long) _
                As Boolean
                
    Dim ltDialog As tFileDialog
    pGetUDT ltDialog, sTitle, iFlags, sFilter, iFilterIndex, sDefaultExt, sInitPath, sInitFile, hWndOwner, vTemplate, hInstance
    ShowOpen = mCommonDialog.File_ShowOpenIndirect(ltDialog)
    
    iReturnExtendedError = ltDialog.iReturnExtendedError
    mtFileDialog.iReturnExtendedError = ltDialog.iReturnExtendedError
    
    If ShowOpen Then
    
    With ltDialog
        
        iReturnFilterIndex = .iReturnFilterIndex
        iReturnFlags = .iReturnFlags
        sReturnFileName = .sReturnFileName
        
        
        mtFileDialog.iReturnFilterIndex = .iReturnFilterIndex
        mtFileDialog.iReturnFlags = .iReturnFlags
        mtFileDialog.sReturnFileName = .sReturnFileName
    End With

    Else
    
        If CBool(mtFileDialog.iFlags And dlgFileRaiseError) Then
            If ltDialog.iReturnExtendedError = 0& _
                Then gErr dlgUserCanceled, ClassName _
                Else gErr dlgExtendedError, ClassName
        End If
    End If

End Function

Public Function ShowSave( _
   Optional ByRef sReturnFileName As String, _
   Optional ByVal iFlags As eFileDialog = NegOneL, _
   Optional ByVal sFilter, _
   Optional ByVal iFilterIndex As Long = NegOneL, _
   Optional ByVal sDefaultExt, _
   Optional ByVal sInitPath, _
   Optional ByVal sInitFile, _
   Optional ByVal sTitle, _
   Optional ByVal hWndOwner As Long = NegOneL, _
   Optional ByVal vTemplate, _
   Optional ByVal hInstance As Long, _
   Optional ByRef iReturnValue As Long, _
   Optional ByRef iReturnFlags As Long, _
   Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
   Optional ByRef iReturnFilterIndex As Long) _
                As Boolean
    
    Dim ltDialog As tFileDialog
    pGetUDT ltDialog, sTitle, iFlags, sFilter, iFilterIndex, sDefaultExt, sInitPath, sInitFile, hWndOwner, vTemplate, hInstance
    ShowSave = mCommonDialog.File_ShowSaveIndirect(ltDialog)
    
    With ltDialog
        iReturnExtendedError = .iReturnExtendedError
        iReturnFilterIndex = .iReturnFilterIndex
        iReturnFlags = .iReturnFlags
        sReturnFileName = .sReturnFileName
        
        mtFileDialog.iReturnExtendedError = .iReturnExtendedError
        mtFileDialog.iReturnFilterIndex = .iReturnFilterIndex
        mtFileDialog.iReturnFlags = .iReturnFlags
        mtFileDialog.sReturnFileName = .sReturnFileName
    End With
    
    If ShowSave = False And CBool(mtFileDialog.iFlags And dlgFileRaiseError) Then
        If ltDialog.iReturnExtendedError = 0& _
            Then gErr dlgUserCanceled, ClassName _
            Else gErr dlgExtendedError, ClassName
    End If
    
End Function

Public Function GetFilter(ParamArray vFilters() As Variant) As String
    
    Dim i As Long
    
    For i = LBound(vFilters) To UBound(vFilters)
        If Not IsMissing(vFilters(i)) Then
            GetFilter = GetFilter & vFilters(i) & dlgFileFilterDelim
        End If
    Next
End Function

Public Function GetMultiFileNames(ByRef sReturnPath As String, ByRef sFileNames() As String) As Long
    GetMultiFileNames = mCommonDialog.File_GetMultiFileNames(mtFileDialog.sReturnFileName, sReturnPath, sFileNames)
End Function

Public Sub CenterDialog(ByVal hDlg As Long, ByVal vCenterTo As Variant)
    CenterWindow hDlg, vCenterTo, ClassName, True
End Sub

Public Property Get Title() As String
    Title = mtFileDialog.sTitle
End Property
Public Property Let Title(ByRef sNew As String)
    mtFileDialog.sTitle = sNew
End Property

Public Property Get Flags() As eFileDialog
    Flags = mtFileDialog.iFlags
End Property
Public Property Let Flags(ByVal iNew As eFileDialog)
    mtFileDialog.iFlags = iNew
End Property
Public Sub SetFlags(Optional ByVal iFlagsOr As eFileDialog, Optional ByVal iFlagsAndNot As eFileDialog)
    mtFileDialog.iFlags = (mtFileDialog.iFlags Or iFlagsOr) And Not iFlagsAndNot
End Sub

Public Property Get Filter() As String
    Filter = mtFileDialog.sFilter
End Property
Public Property Let Filter(ByRef sNew As String)
    mtFileDialog.sFilter = sNew
End Property

Public Property Get FilterIndex() As Long
    FilterIndex = mtFileDialog.iFilterIndex
End Property
Public Property Let FilterIndex(ByVal iNew As Long)
    mtFileDialog.iFilterIndex = iNew
End Property

Public Property Get DefaultExt() As String
    DefaultExt = mtFileDialog.sDefaultExt
End Property
Public Property Let DefaultExt(ByRef sNew As String)
    mtFileDialog.sDefaultExt = sNew
End Property

Public Property Get InitialPath() As String
    InitialPath = mtFileDialog.sInitPath
End Property
Public Property Let InitialPath(ByRef sNew As String)
    mtFileDialog.sInitPath = sNew
End Property

Public Property Get InitialFile() As String
    InitialFile = mtFileDialog.sInitFile
End Property
Public Property Let InitialFile(ByRef sNew As String)
    mtFileDialog.sInitFile = sNew
End Property

Public Property Get hWndOwner() As Long
    hWndOwner = mtFileDialog.hWndOwner
End Property
Public Property Let hWndOwner(ByVal iNew As Long)
    mtFileDialog.hWndOwner = iNew
End Property

Public Property Get Template() As Variant
    Template = mtFileDialog.vTemplate
End Property
Public Property Let Template(ByRef vNew As Variant)
    mtFileDialog.vTemplate = vNew
End Property

Public Property Get hInstanceTemplate() As Long
    hInstanceTemplate = mtFileDialog.hInstance
End Property
Public Property Let hInstanceTemplate(ByVal iNew As Long)
    mtFileDialog.hInstance = iNew
End Property


Public Property Get ReturnFlags() As eFileDialog
    ReturnFlags = mtFileDialog.iReturnFlags
End Property

Public Property Get ReturnFilterIndex() As Long
    ReturnFilterIndex = mtFileDialog.iReturnFilterIndex
End Property

Public Property Get ReturnExtendedError() As eComDlgExtendedError
    ReturnExtendedError = mtFileDialog.iReturnExtendedError
End Property

Public Property Get ReturnFileName() As String
    ReturnFileName = mtFileDialog.sReturnFileName
End Property


Private Sub Class_Initialize()
    mtFileDialog.iFlags = dlgFileExplorerStyle Or dlgFileEnableSizing Or dlgFileMustExist Or dlgFilePathMustExist
    mtFileDialog.sFilter = "All Files (*.*)" & dlgFileFilterDelim & "*.*"
End Sub

Private Sub pGetUDT( _
            ByRef tFileDialog As tFileDialog, _
            ByRef sTitle As Variant, _
            ByVal iFlags As eFileDialog, _
            ByRef sFilter As Variant, _
            ByVal iFilterIndex As Long, _
            ByRef sDefaultExt As Variant, _
            ByRef sInitPath As Variant, _
            ByRef sInitFile As Variant, _
            ByVal hWndOwner As Long, _
            ByRef vTemplate As Variant, _
            ByVal hInstance As Long)
        
    With tFileDialog
        
        If iFlags <> NegOneL _
            Then .iFlags = iFlags _
            Else .iFlags = mtFileDialog.iFlags
        
        If hWndOwner <> NegOneL _
            Then .hWndOwner = hWndOwner _
            Else .hWndOwner = mtFileDialog.hWndOwner
        
        If IsMissing(sTitle) _
            Then .sTitle = mtFileDialog.sTitle _
            Else .sTitle = CStr(sTitle)
        
        If IsMissing(sInitFile) _
            Then .sInitFile = mtFileDialog.sInitFile _
            Else .sInitFile = CStr(sInitFile)
        
        If IsMissing(sInitPath) _
            Then .sInitPath = mtFileDialog.sInitPath _
            Else .sInitPath = CStr(sInitPath)
        
        If IsMissing(sDefaultExt) _
            Then .sDefaultExt = mtFileDialog.sDefaultExt _
            Else .sDefaultExt = CStr(sDefaultExt)
        
        If hInstance <> NegOneL _
            Then .hInstance = hInstance _
            Else .hInstance = mtFileDialog.hInstance
        
        If IsMissing(vTemplate) _
            Then .vTemplate = mtFileDialog.vTemplate _
            Else .vTemplate = vTemplate
        
        If IsMissing(sFilter) _
            Then .sFilter = mtFileDialog.sFilter _
            Else .sFilter = CStr(sFilter)
        
        If iFilterIndex <> NegOneL _
            Then .iFilterIndex = iFilterIndex _
            Else .iFilterIndex = mtFileDialog.iFilterIndex
        
        #If bHookDialogs Then
            Set .oEventSink = Me
        #End If
    End With
End Sub

#If bHookDialogs Then

    Private Sub iComDlgHook_Proc(lReturn As Long, ByVal hDlg As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
        
        Const WM_INITDIALOG As Long = &H110
        Const WM_COMMAND As Long = &H111
        Const WM_DESTROY As Long = &H2
        Const WM_NOTIFY As Long = &H4E
        
        Const DWL_MSGRESULT = ZeroL
        
        Const H_MAX As Long = &HFFFF + 1
        Const CDN_FIRST As Long = (H_MAX - 601)
        Const CDN_INITDONE As Long = (CDN_FIRST - &H0)
        Const CDN_SELCHANGE As Long = (CDN_FIRST - &H1)
        Const CDN_FOLDERCHANGE As Long = (CDN_FIRST - &H2)
        Const CDN_SHAREVIOLATION As Long = (CDN_FIRST - &H3)
        Const CDN_HELP As Long = (CDN_FIRST - &H4)
        Const CDN_FILEOK As Long = (CDN_FIRST - &H5)
        Const CDN_TYPECHANGE As Long = (CDN_FIRST - &H6)
        
        Select Case iMsg
        Case WM_INITDIALOG
            RaiseEvent DialogInit(hDlg)
            lReturn = ZeroL
        
        Case WM_NOTIFY
            'offset to NMHDR.code
            
            CopyMemory iMsg, ByVal UnsignedAdd(lParam, 8&), 4&
            Select Case iMsg
            Case CDN_SELCHANGE
                RaiseEvent FileChange(hDlg)
            Case CDN_FOLDERCHANGE
                RaiseEvent FolderChange(hDlg)
            Case CDN_FILEOK
                Dim lbCancel As Boolean
                RaiseEvent DialogOK(hDlg, lbCancel)
                lReturn = Abs(lbCancel)
                SetWindowLong hDlg, DWL_MSGRESULT, lReturn
                
            Case CDN_HELP
                RaiseEvent HelpClicked(hDlg)
            Case CDN_TYPECHANGE
                RaiseEvent TypeChange(hDlg)
            End Select
            
        Case WM_COMMAND
            RaiseEvent WMCommand(hDlg, wParam, lParam)
            
        Case WM_DESTROY
            RaiseEvent DialogClose(hDlg)
            
        End Select
    End Sub

#End If
