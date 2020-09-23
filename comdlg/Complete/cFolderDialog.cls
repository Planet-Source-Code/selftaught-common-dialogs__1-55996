VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFolderDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cFolderDialog.cls                  8/25/04
'
'           GENERAL PURPOSE:
'               Show a Folder dialog, optionally hooking it and raising event for useful messages.
'
'           LINEAGE:
'               www.vbaccelerator.com
'
'           DEPENDENCIES:
'               vbBaseTlb.tlb
'
'               mCommonDialog.bas (with the bFolderDialog compiler switch on)
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

Private mtDialog As tFolderDialog
Private mtTempDialog As tFolderDialog

Private msReturnPath As String

Public Event DialogInit(ByVal hDlg As Long)
Public Event FolderChanged(ByVal hDlg As Long, ByVal sPath As String, ByRef bCancel As Boolean)
Public Event ValidationFailed(ByVal hDlg As Long, ByVal sPath As String, ByRef bKeepOpen As Boolean)

Private Const ClassName = "cFolderDialog"

Private Sub Class_Initialize()
   mtDialog.sTitle = "Choose Folder"
   mtDialog.iFlags = dlgFolderUseNewUI Or dlgFolderStatusText Or dlgFolderReturnOnlyFSDirs
End Sub

Private Sub pGetUDT( _
            ByRef tUDT As tFolderDialog, _
            ByVal iFlags As Long, _
            ByRef sTitle, _
            ByRef sInitialPath, _
            ByRef sRootPath, _
            ByVal hWndOwner As Long)
    
    With tUDT
        If iFlags <> NegOneL _
            Then .iFlags = iFlags _
            Else .iFlags = mtDialog.iFlags
        
        If hWndOwner <> NegOneL _
            Then .hWndOwner = hWndOwner _
            Else .hWndOwner = mtDialog.hWndOwner
        
        If IsMissing(sTitle) _
            Then .sTitle = mtDialog.sTitle _
            Else .sTitle = CStr(sTitle)
        
        If IsMissing(sInitialPath) _
            Then .sInitialPath = mtDialog.sInitialPath _
            Else .sInitialPath = CStr(sInitialPath)
        
        If IsMissing(sRootPath) _
            Then .sRootPath = mtDialog.sRootPath _
            Else .sRootPath = CStr(sRootPath)
        
        #If bHookDialogs Then
            Set .oEventSink = Me
        #End If
    End With
    
End Sub

Public Property Get Flags() As eFolderDialog
    Flags = mtDialog.iFlags
End Property
Public Property Let Flags(ByVal iNew As eFolderDialog)
    mtDialog.iFlags = iNew
End Property
Public Sub SetFlags(Optional ByVal iFlagsOr As eFolderDialog, Optional ByVal iFlagsAndNot As eFolderDialog)
    mtDialog.iFlags = (mtDialog.iFlags Or iFlagsOr) And Not iFlagsAndNot
End Sub

Public Property Get EditBox() As Boolean
   EditBox = CBool(mtDialog.iFlags And dlgFolderEditBox)
End Property
Public Property Let EditBox(ByVal bState As Boolean)
   If bState _
        Then SetFlags dlgFolderEditBox _
        Else SetFlags , dlgFolderEditBox
End Property

Public Property Get StatusText() As Boolean
   StatusText = CBool(mtDialog.iFlags And dlgFolderStatusText)
End Property
Public Property Let StatusText(ByVal bState As Boolean)
   If bState _
        Then SetFlags dlgFolderStatusText _
        Else SetFlags , dlgFolderStatusText
End Property

Public Property Get FileSystemOnly() As Boolean
   FileSystemOnly = CBool(mtDialog.iFlags And dlgFolderReturnOnlyFSDirs)
End Property
Public Property Let FileSystemOnly(ByVal bState As Boolean)
   If bState _
        Then SetFlags dlgFolderReturnOnlyFSDirs Or dlgFolderReturnFSAncestors _
        Else SetFlags , dlgFolderReturnOnlyFSDirs Or dlgFolderReturnFSAncestors
End Property

Public Property Get ValidateEditBox() As Boolean
   ValidateEditBox = CBool(mtDialog.iFlags And dlgFolderValidate)
End Property
Public Property Let ValidateEditBox(ByVal bState As Boolean)
   If bState _
        Then SetFlags dlgFolderValidate _
        Else SetFlags , dlgFolderValidate
End Property

Public Property Get UseNewUI() As Boolean
   UseNewUI = (mtDialog.iFlags And dlgFolderUseNewUI)
End Property
Public Property Let UseNewUI(ByVal bState As Boolean)
    If bState _
        Then SetFlags dlgFolderUseNewUI _
        Else SetFlags , dlgFolderUseNewUI
End Property

Public Property Get Title() As String
   Title = mtDialog.sTitle
End Property
Public Property Let Title(ByVal sTitle As String)
   mtDialog.sTitle = sTitle
End Property

Public Property Get hWndOwner() As Long
   hWndOwner = mtDialog.hWndOwner
End Property
Public Property Let hWndOwner(ByVal lhWnd As Long)
   mtDialog.hWndOwner = lhWnd
End Property

Public Property Get InitialPath() As String
   InitialPath = mtDialog.sInitialPath
End Property
Public Property Let InitialPath(ByVal sDir As String)
   mtDialog.sInitialPath = sDir
End Property

Public Property Get RootPath() As String
   RootPath = mtDialog.sRootPath
End Property
Public Property Let RootPath(ByVal sDir As String)
   mtDialog.sRootPath = sDir
End Property

Public Property Get Path() As String
    Path = msReturnPath
End Property

Public Property Get SpecialFolder(ByVal iFolder As eSpecialFolders) As String
    SpecialFolder = mCommonDialog.Folder_SpecialFolder(mtDialog.hWndOwner, iFolder)
End Property
    
Public Function Show( _
   Optional ByRef sReturnPath As String, _
   Optional ByVal iFlags As Long = NegOneL, _
   Optional ByVal sTitle, _
   Optional ByVal sInitialPath, _
   Optional ByVal sRootPath, _
   Optional ByVal hWndOwner As Long = NegOneL) _
                As Boolean
    
    pGetUDT mtTempDialog, iFlags, sTitle, sInitialPath, sRootPath, hWndOwner
    
    Show = mCommonDialog.Folder_ShowIndirect(mtTempDialog)
    msReturnPath = mtTempDialog.sReturnPath
    
    If Not Show And CBool(mtTempDialog.iFlags And dlgFolderRaiseError) Then
        gErr dlgUserCanceled, ClassName
    End If
End Function

Public Sub SetFolder(ByVal hDlg As Long, ByVal sPath As String)
    mCommonDialog.Folder_SetFolder hDlg, sPath
End Sub
Public Sub SetStatus(ByVal hDlg As Long, ByVal sText As String)
    mCommonDialog.Folder_SetStatus hDlg, sText
End Sub

Public Sub CenterDialog(ByVal hDlg As Long, ByVal vCenterTo As Variant)
    CenterWindow hDlg, vCenterTo, ClassName
End Sub

#If bHookDialogs Then

    Private Sub iComDlgHook_Proc(lReturn As Long, ByVal hDlg As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
        
        Const BFFM_INITIALIZED As Long = 1
        Const BFFM_SELCHANGED  As Long = 2
        Const BFFM_VALIDATEFAILEDA = 3      '// lParam:szPath ret:1(cont),0(EndDialog)
        
        Dim sPath As String
        Dim lR As Long
        Dim bCancel As Boolean
        Dim bKeepOpen As Boolean
        
        Select Case iMsg
        ' Browse dialog box has finished initializing (lParam is NULL)
        Case BFFM_INITIALIZED
            If mtTempDialog.iPidlInitial > 0 Then
                mCommonDialog.Folder_SetFolderPidl hDlg, mtTempDialog.iPidlInitial
            End If
            RaiseEvent DialogInit(hDlg)
            lReturn = ZeroL
    
        ' Selection has changed (lParam contains pidl of selected folder)
        Case BFFM_SELCHANGED
            'Debug.Print "BFFM_SELCHANGED"
            ' Display full path if status area if enabled
            
            sPath = mCommonDialog.Folder_PathFromPidl(lParam)
            mCommonDialog.Folder_SetStatus hDlg, sPath
                   
            bCancel = False
            RaiseEvent FolderChanged(hDlg, sPath, bCancel)
            
            Debug.Assert Not bCancel
            If bCancel Then mCommonDialog.Folder_EnableOK hDlg, False
    
            lReturn = ZeroL
                
        ' Invalid name in edit box (lParam parameter has invalid name string)
        Case BFFM_VALIDATEFAILEDA
            'Debug.Print "BFFM_VALIDATEFAILED"
            ' Return zero to dismiss dialog or nonzero to keep it displayed
            ' Disable the OK button
            mCommonDialog.Folder_EnableOK hDlg, False
            sPath = mCommonDialog.Folder_PtrToString(lParam)
            sPath = "Path invalid: " & sPath
            mCommonDialog.Folder_SetStatus hDlg, sPath
            
            RaiseEvent ValidationFailed(hDlg, sPath, bKeepOpen)
            lReturn = Abs(bKeepOpen)
            
       End Select
    
    End Sub

#End If
