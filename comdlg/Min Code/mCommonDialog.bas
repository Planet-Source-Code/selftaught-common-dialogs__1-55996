Attribute VB_Name = "mCommonDialog"
Option Explicit

'==================================================================================================
'mCommonDialog.bas              8/25/04
'
'           GENERAL PURPOSE:
'               Implement all COMDLG common dialogs, and shell browse for folder.
'
'               Can be used as a stand alone module for calling the dialogs, and/or
'               to provide functionality for these dialog classes.
'                   cColorDialog.cls
'                   cFileDialog.cls
'                   cFolderDialog.cls
'                   cFontDialog.cls
'                   cPrintDialog.cls
'                   cPageSetupDialog.cls
'
'               Functionality can be enabled/disabled separately for each dialog,
'               using compiler switches.  If you include any of the classes listed above,
'               then you must turn on the corresponding compiler switch from below.
'                    #Const bColorDialog
'                    #Const bFileDialog
'                    #Const bFolderDialog
'                    #Const bFontDialog
'                    #Const bPrintDialog
'                    #Const bPageSetupDialog
'
'               Disabling dialog hooks if you don't need them will eliminate one
'               member from every dialog udt and one argument from each ShowXxxx
'               procedure, in addition to the code required for creating
'               the assembly code and the code that responds to the hook procedure
'               messages in the classes listed above.
'                    #Const bHookDialogs
'               If this compiler switch is false, then the bHookDialog compiler switch
'               must be false in the classes listed above to avoid a compilation error.
'
'           LINEAGE:
'               CommonDialogDirect6 from www.vbaccelerator.com
'
'           DEPENDENCIES:
'               If the bHookDialogs compiler switch is set, then iComDlgHook.cls
'               if the bVBVMTypeLib compiler switch is set, then VBVM6Lib.tlb
'
'               If bFontDialog and bDependcFont then cFont.cls
'               If bPageSetupDialog or bPrintDialog then cDeviceMode.cls
'
'
'==================================================================================================

'The compiler switch bHookDialogs must also be changed to match this
'module's value in all of these types of classes in the same project:
'
'                   cFileDialog.cls
'                   cFolderDialog.cls
'                   cFontDialog.cls
'                   cColorDialog.cls
'                   cPrintDialog.cls
'                   cPageSetupDialog.cls
'
'   A True value for the bHookDialogs compiler constant in any of the classes above
'   combined with a False value in this module will result in a compilation error.

Public Const NoDate        As Date = #12:00:00 AM#
Public Const TenThouC      As Currency = 10000@
Public Const ZeroC         As Currency = 0@
Public Const NegOneC       As Currency = -1@
Public Const ZeroL         As Long = 0&
Public Const NegOneL       As Long = -1&
Public Const OneL          As Long = 1&
Public Const TwoL          As Long = 2&
Public Const NegOneF       As Single = -1!
Public Const ZeroY         As Byte = 0

#Const bHookDialogs = False

#Const bVBVMTypeLib = False

#Const bColorDialog = True
#Const bFileDialog = True
#Const bFolderDialog = False
#Const bFontDialog = True
#Const bPrintDialog = False
#Const bPageSetupDialog = False

#Const bDependcFont = False

Private Type tRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum eComDlgError
    dlgUserCanceled = 32755
    dlgExtendedError = 39512
    dlgTypeMismatch = 13
End Enum

Public Enum eComDlgExtendedError
    dlgErrDialogFailure = &HFFFF

    dlgErrGeneralCodes = &H0&
    dlgErrStructsize = &H1&
    dlgErrInitialization = &H2&
    dlgErrNoTemplate = &H3&
    dlgErrNoHInstance = &H4&
    dlgErrLoadStrFailure = &H5&
    dlgErrFindResFailure = &H6&
    dlgErrLoadResFailure = &H7&
    dlgErrLockResFailure = &H8&
    dlgErrMmeAllocFailure = &H9&
    dlgErrMmelockFailure = &HA&
    dlgErrNoHook = &HB&
    dlgErrRegisterMsgFail = &HC&

    dlgPrintErrPrinterCodes = &H1000&
    dlgPrintErrSetupFailure = &H1001&
    dlgPrintErrParseFailure = &H1002&
    dlgPrintErrRetDefFailure = &H1003&
    dlgPrintErrLoadDrvFailure = &H1004&
    dlgPrintErrGetDevModeFail = &H1005&
    dlgPrintErrInitFailure = &H1006&
    dlgPrintErrNoDevices = &H1007&
    dlgPrintErrNoDefaultPrn = &H1008&
    dlgPrintErrDNDMMismatch = &H1009&
    dlgPrintErrCreateICFailure = &H100A&
    dlgPrintErrPrinterNotFound = &H100B&
    dlgPrintErrDefaultDifferent = &H100C&

    dlgFontErrChooseFontCodes = &H2000&
    dlgFontErrNoFonts = &H2001&
    dlgFontErrMaxLessThanMin = &H2002&

    dlgFileErrFileNameCodes = &H3000&
    dlgFileErrSubclassFailure = &H3001&
    dlgFileErrInvalidFilename = &H3002&
    dlgFileErrBufferTooSmall = &H3003&

    dlgColorErrChooseColorCodes = &H5000&
End Enum

#If bFileDialog Then

    Public Enum eFileDialog
        dlgFileExplorerStyle = &H80000
        dlgFileMustExist = &H1000&
        dlgFilePathMustExist = &H800&
        dlgFileMultiSelect = &H200
        dlgFilePromptToCreate = &H2000
        dlgFileEnableSizing = &H800000
        dlgFileNoDereferenceLinks = &H100000
        dlgFileHideNetworkButton = &H20000
        dlgFileHideReadOnly = &H4
        dlgFileNoReadOnlyReturn = &H8000&
        dlgFileNoTestFileCreate = &H10000
        dlgFileOverwritePrompt = &H2&
        dlgFileInitToReadOnly = &H1&
        dlgFileShowHelpButton = &H10&
        dlgFileEnableHook = &H20&
        dlgFileEnableTemplate = &H40&
        dlgFileRaiseError = &H10000000
    End Enum

#End If

#If bColorDialog Then
    Public Enum eColorDialogFlag
        'dlgColorRGBInit = &H1
        dlgColorFullOpen = &H2
        dlgColorPreventFullOpen = &H4
    ' Win95 only
        dlgColorSolid = &H80
        dlgColorAny = &H100
    ' End Win95 only
        dlgColorEnableHook = &H10
        'dlgColorEnableTemplate = &H20
        'dlgColorEnableTemplateHandle = &H40
        dlgColorRaiseError = &H10000000
    End Enum
#End If

#If bFontDialog Then
    
    Public Enum eFontDialog
        dlgFontScreenFonts = &H1
        dlgFontPrinterFonts = &H2
        dlgFontScreenAndPrinterFonts = &H3
        dlgFontUseStyle = &H80
        dlgFontEffects = &H100
        dlgFontAnsiOnly = &H400
        dlgFontNoVectorFonts = &H800
        dlgFontNoOemFonts = dlgFontNoVectorFonts
        dlgFontNoSimulations = &H1000
        dlgFontFixedPitchOnly = &H4000
        dlgFontWysiwyg = &H8000&  ' Must also have ScreenFonts And PrinterFonts
        dlgFontForceExist = &H10000
        dlgFontScalableOnly = &H20000
        dlgFontTTOnly = &H40000
        dlgFontNoFaceSel = &H80000
        dlgFontNoStyleSel = &H100000
        dlgFontNoSizeSel = &H200000
        ' Win95 only
        dlgFontSelectScript = &H400000
        dlgFontNoScriptSel = &H800000
        dlgFontNoVertFonts = &H1000000
    
        
        dlgFontApply = &H200
        dlgFontEnableHook = &H8
        'dlgFontEnableTemplate = &H10
        'dlgFontEnableTemplateHandle = &H20
        'dlgFontNotSupported = &H238
        dlgFontRaiseError = &H10000000
    End Enum
    
#End If

#If bColorDialog Or _
    bFileDialog Or _
    bFolderDialog Or _
    bFontDialog Or _
    bPrintDialog Or _
    bPageSetupDialog _
        Then
    
    Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
    Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As tRect) As Long
    Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    
    Private Const SPI_GETWORKAREA = 48&
    
#End If

#If bFileDialog Or bFolderDialog Or bPrintDialog Then
    Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal pszStr As Long, ByVal lLenB As Long) As Long
    Private Const MAX_PATH      As Long = 260
#End If

#If bFileDialog Then

    Public Const dlgFileFilterDelim = "|"
    
    Public Type tFileDialog
        iFlags                  As eFileDialog
        sFilter                 As String
        iFilterIndex            As Long
        sDefaultExt             As String
        sInitPath               As String
        sInitFile               As String
        sTitle                  As String
        hWndOwner               As Long
        vTemplate               As Variant
        hInstance               As Long
        
        #If bHookDialogs Then
        
        oEventSink              As iComDlgHook
        
        #End If
    
        sReturnFileName         As String
        iReturnFlags            As Long
        iReturnExtendedError    As eComDlgExtendedError
        iReturnFilterIndex      As Long
    End Type
    
    Private Type OPENFILENAME
        lStructSize                     As Long
        hWndOwner                       As Long
        hInstance                       As Long
        lpstrFilter                     As String
        lpstrCustomFilter               As String
        nMaxCustFilter                  As Long
        nFilterIndex                    As Long
        lpstrFile                       As String
        nMaxFile                        As Long
        lpstrFileTitle                  As String
        nMaxFileTitle                   As Long
        lpstrInitialDir                 As String
        lpstrTitle                      As String
        Flags                           As Long
        nFileOffset                     As Integer
        nFileExtension                  As Integer
        lpstrDefExt                     As String
        lCustData                       As Long
        lpfnHook                        As Long
        lpTemplateName                  As Long
    End Type
    
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (File As OPENFILENAME) As Long
    Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (File As OPENFILENAME) As Long
    
#End If

#If bFolderDialog Then
    
    Public Type tFolderDialog
        hWndOwner               As Long
        sTitle                  As String
        sInitialPath            As String
        sRootPath               As String
        sReturnPath             As String
        iPidlInitial            As Long
        iFlags                  As eFolderDialog
        #If bHookDialogs Then
        oEventSink              As iComDlgHook
        #End If
    End Type

    Private Const WM_USER = &H400
    Private Const BFFM_INITIALIZED As Long = 1
    Private Const BFFM_SELCHANGED  As Long = 2
    Private Const BFFM_VALIDATEFAILEDA = 3      '// lParam:szPath ret:1(cont),0(EndDialog)
    
    Private Const BFFM_SETSTATUSTEXTA = (WM_USER + 100)
    Private Const BFFM_ENABLEOK = (WM_USER + 101)
    Private Const BFFM_SETSELECTIONA = (WM_USER + 102)
    Private Const BFFM_SETSELECTIONW = (WM_USER + 103)
    Private Const BFFM_SETSTATUSTEXTW = (WM_USER + 104)
    
    Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
    
    Private Type BrowseInfo
        hWndOwner As Long
        pIDLRoot As Long
        pszDisplayName As Long ';// Return display name of item selected.
        lpszTitle As Long ';      // text to go in the banner over the tree.
        ulFlags As Long ';       // Flags that control the return stuff
        lpfn As Long
        lParam As Long         '// extra info that's passed back in callbacks
        iImage As Long ';      // output var: where to return the Image index.
    End Type
    
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function SendMessageLong Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hWnd As Long) As Long
    
    Private Declare Function lstrlenptr Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
    Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
    Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
    
    Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpbi As BrowseInfo) As Long
    Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
    Private Declare Function SHGetDesktopFolder Lib "shell32.dll" (ppshf As IShellFolder) As Long
    Private Declare Function SHGetMalloc Lib "shell32.dll" (ppMalloc As IMalloc) As Long
    
    Private Declare Function CoInitialize Lib "ole32.dll" Alias "CoInitializeEx" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long
    
#End If

#If bFolderDialog Or bPrintDialog Then
    Private Declare Sub CopyMemoryLpToStr Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal lpvDest As String, lpvSource As Long, ByVal cbCopy As Long)
#End If


#If bColorDialog Then
    
    Public Type tColorDialog
        hWndOwner               As Long
        iColor                  As OLE_COLOR
        iFlags                  As eColorDialogFlag
        iReturnExtendedError    As eComDlgExtendedError
        #If bHookDialogs Then
        oEventSink              As iComDlgHook
        #End If
        iColors(0 To 15)        As Long
    End Type
    
    Private Type TCHOOSECOLOR
        lStructSize As Long
        hWndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        Flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As Long
    End Type
    
    Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long
    Private Declare Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
    Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
    
#End If


#If bFontDialog Then
   
    Private Const FW_BOLD As Long = 700
    Private Const FW_NORMAL As Long = 400
   
    Private Const LF_FACESIZE As Long = 32
    Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(0 To LF_FACESIZE - 1) As Byte
    End Type
    
    Public Type tFontDialog
        'in:
        iFlags                  As eFontDialog
        hdc                     As Long
        hWndOwner               As Long
        iMinSize                As Long
        iMaxSize                As Long
        oFont                   As Object
        #If bHookDialogs Then
        oEventSink              As iComDlgHook
        #End If
        'out
        iColor                  As OLE_COLOR
        iReturnFlags            As eFontDialog
        iReturnExtendedError    As eComDlgExtendedError
    End Type
    
    
    Const dlgFontLimitSize = &H2000&
    Const dlgFontInitToLogFontStruct = &H40&
    
    Private Type TCHOOSEFONT
        lStructSize As Long         ' Filled with UDT size
        hWndOwner As Long           ' Caller's window handle
        hdc As Long                 ' Printer DC/IC or NULL
        lpLogFont As Long           ' Pointer to LOGFONT
        iPointSize As Long          ' 10 * size in points of font
        Flags As Long               ' Type flags
        rgbColors As Long           ' Returned text color
        lCustData As Long           ' Data passed to hook function
        lpfnHook As Long            ' Pointer to hook function
        lpTemplateName As Long      ' Custom template name
        hInstance As Long           ' Instance handle for template
        lpszStyle As String         ' Return style field
        nFontType As Integer        ' Font type bits
        iAlign As Integer           ' Filler
        nSizeMin As Long            ' Minimum point size allowed
        nSizeMax As Long            ' Maximum point size allowed
    End Type
    
    Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long
    Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    
#End If


#If bPageSetupDialog Then

    Public Type tPageSetupDialog
        iFlags                  As ePrintPageSetup
        hWndOwner               As Long
        iUnits                  As ePrintPageSetupUnits
        
        fLeftMargin             As Single
        fMinLeftMargin          As Single
        fRightMargin            As Single
        fMinRightMargin         As Single
        fTopMargin              As Single
        fMinTopMargin           As Single
        fBottomMargin           As Single
        fMinBottomMargin        As Single
        oDeviceMode             As cDeviceMode
        #If bHookDialogs Then
        oEventSink              As iComDlgHook
        #End If
        iReturnExtendedError    As eComDlgExtendedError
    End Type
    
    'Private Const dlgPrintInWinIniIntlMeasure = &H0&
    Private Const dlgPrintInThousandthsOfInches = &H4&
    Private Const dlgPrintInHundredthsofMillimeters = &H8&
    
    Private Type tPoint
        x As Long
        y As Long
    End Type
    
    Private Type tRect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    
    Private Type TPAGESETUPDLG
        lStructSize                 As Long
        hWndOwner                   As Long
        hDevMode                    As Long
        hDevNames                   As Long
        Flags                       As Long
        ptPaperSize                 As tPoint
        rtMinMargin                 As tRect
        rtMargin                    As tRect
        hInstance                   As Long
        lCustData                   As Long
        lpfnPageSetupHook           As Long
        lpfnPagePaintHook           As Long
        lpPageSetupTemplateName     As Long
        hPageSetupTemplate          As Long
    End Type
    
    Private Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (ByRef pPagesetupdlg As TPAGESETUPDLG) As Long
    
#End If


#If bPrintDialog Then
    
    Public Type tPrintDialog
        hdc                     As Long
        hWndOwner               As Long
        iFlags                  As ePrintDialog
        iRange                  As ePrintRange
        iMinPage                As Long
        iMaxPage                As Long
        
        oDeviceMode             As cDeviceMode
            
        iFromPage               As Long
        iToPage                 As Long
        iReturnFlags            As ePrintRange
    
        iReturnExtendedError          As eComDlgExtendedError
        bCollate                As Boolean
        bPrintToFile            As Boolean
    
        sDriver                 As String
        sDevice                 As String
        sOutputPort             As String
        #If bHookDialogs Then
        oEventSink              As iComDlgHook
        #End If
    End Type

    Private Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
    End Type

    Private Type TPRINTDLG
        lStructSize As Long
        hWndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        Flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        ' right, now we have serious alignment problems.
        ' We're two bytes off for the next 8 fields...
        b(0 To 31) As Byte
        'hInstance As Long              '  0 - 3
        'lCustData As Long              '  4 - 7
        'lpfnPrintHook As Long          '  8 - 11
        'lpfnSetupHook As Long          ' 12 - 15
        'lpPrintTemplateName As Long    ' 16 - 19
        'lpSetupTemplateName As Long    ' 20 - 23
        'hPrintTemplate As Long         ' 24 - 27
        'hSetupTemplate As Long         ' 28 - 31
    End Type
    
    Private Const tPrintDlghInstanceOffset As Long = 0
    Private Const tPrintDlglCustDataOffset As Long = 4
    Private Const tPrintDlglpfnPrintHookOffset As Long = 8
    Private Const tPrintDlglpfnSetupHookOffset As Long = 12
    Private Const tPrintDlglpPrintTemplateNameOffset As Long = 16
    Private Const tPrintDlglpSetupTemplateNameOffset As Long = 20
    Private Const tPrintDlgPrintTemplateOffset As Long = 24
    Private Const tPrintDlgSetupTemplateOffset As Long = 28
    
    Private Const dlgPrintUseDevModeCopiesAndCollate = &H40000

    '  DEVMODE collation selections
    Private Const DMCOLLATE_FALSE = 0
    Private Const DMCOLLATE_TRUE = 1

    Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (ByRef pPrintdlg As TPRINTDLG) As Long

    Private Declare Function lstrlenByLong Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long

#End If

#If bPrintDialog Or bPageSetupDialog Or bHookDialogs Then
    
    Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
    
#End If

#If bPrintDialog Or bPageSetupDialog Or bFontDialog Then

    Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
    
#End If


#If bHookDialogs Then
    
    Private Function pAllocateHookProc(ByVal oWho As iComDlgHook) As Long

        Const ASM As String = "E8000000005A8F823400000031C089823800000081C23800000052B8ABCABC01508B00FF501CE8000000005AFFB20E0000008B8212000000C30000000000000000"
        Const ASMLength As Long = 65
        
        Const PATCH_OBJPTR = &H1C&
       
        Static yCodeBuffer(0 To ASMLength - 1) As Byte
        Static bInit As Boolean
        
        If Not oWho Is Nothing Then                                             'Must have valid objptr
            
            pAllocateHookProc = GlobalAlloc(0&, ASMLength)                      'allocate memory
            
            Debug.Assert pAllocateHookProc                                      'out of memory!
            
            If pAllocateHookProc <> 0& Then                                     'if succeeded
                
                Dim i As Long
                
                If Not bInit Then                                               'if we haven't initialized
                    bInit = True
                    
                    For i = 0 To ASMLength - 1&                                 'build the asm in memory
                        yCodeBuffer(i) = CByte("&H" & Mid$(ASM, (i + i) + 1&, 2&))
                    Next
                    
                End If
                
                CopyMemory ByVal pAllocateHookProc, yCodeBuffer(0), ASMLength   'copy the asm to the allocated memory
                
                'patch the owner object address
                #If bVBVMTypeLib Then               'use the illegal type lib
                    MemLong(ByVal UnsignedAdd(pAllocateHookProc, PATCH_OBJPTR)) = ObjPtr(oWho)
                #Else                               'use the block moving sledgehammer
                    CopyMemory ByVal UnsignedAdd(pAllocateHookProc, PATCH_OBJPTR), ObjPtr(oWho), 4&
                #End If
                
            End If
            
        End If
        
    End Function
    
#End If

#If bColorDialog Or _
    bFileDialog Or _
    bFolderDialog Or _
    bFontDialog Or _
    bPrintDialog Or _
    bPageSetupDialog _
        Then

    Public Function UnsignedAdd(ByVal iStart As Long, ByVal iInc As Long) As Long
    ' This function is useful when doing pointer arithmetic,
    ' but note it only works for positive values of Incr
        Const VBLongSignBit = &H80000000
        
        If iStart And VBLongSignBit Then
            UnsignedAdd = iStart + iInc
        ElseIf (iStart Or VBLongSignBit) < -iInc Then
            UnsignedAdd = iStart + iInc
        Else
            UnsignedAdd = (iStart + VBLongSignBit) + _
                            (iInc + VBLongSignBit)
        End If
    End Function
    
    Public Sub gErr(ByVal iNum As eComDlgError, ByRef sSource As String, Optional ByRef sDesc As String)
        Dim lsDesc As String
        If LenB(sDesc) = ZeroL Then
            Select Case iNum
                Case dlgTypeMismatch
                    lsDesc = "Type mismatch."
                Case dlgUserCanceled
                    lsDesc = "Operation was canceled by the user."
                Case dlgExtendedError
                    lsDesc = "Extended common dialog error."
            End Select
        Else
            lsDesc = sDesc
        End If
        Err.Raise iNum, sSource, lsDesc
    End Sub
    
    Public Sub CenterWindow(ByVal hWnd As Long, ByRef vCenterTo As Variant, ByRef sClassName As String, Optional ByVal bParent As Boolean)
    
    Dim hWndCenter As Long
    Dim ltRectDialog As tRect
    Dim ltRectCenterTo As tRect
    Dim ltRectWorkArea As tRect
    
        If (SystemParametersInfo(SPI_GETWORKAREA, ZeroL, ltRectWorkArea, ZeroL) = ZeroL) Then
           ' Call failed - just use standard screen:
            With Screen
                ltRectWorkArea.Right = .Width \ .TwipsPerPixelX
                ltRectWorkArea.Bottom = .Height \ .TwipsPerPixelY
            End With
        End If
    
        If bParent Then hWnd = GetParent(hWnd)
        Debug.Assert hWnd <> ZeroL
        
        If GetWindowRect(hWnd, ltRectDialog) Then
            
            If VarType(vCenterTo) = vbObject Then
                If TypeOf vCenterTo Is Screen Then
                    LSet ltRectCenterTo = ltRectWorkArea
                Else
                    On Error Resume Next
                    hWndCenter = vCenterTo.hWnd
                    On Error GoTo 0
                    If GetWindowRect(hWndCenter, ltRectCenterTo) = ZeroL Then gErr dlgTypeMismatch, sClassName
                End If
            ElseIf VarType(vCenterTo) = vbLong Then
                hWndCenter = vCenterTo
                If GetWindowRect(hWndCenter, ltRectCenterTo) = ZeroL Then gErr dlgTypeMismatch, sClassName
            End If
            
            Dim x As Long
            Dim y As Long
            Dim cx As Long
            Dim cy As Long
            
            With ltRectCenterTo
                x = .Left + ((.Right - .Left) \ 2)
                y = .Top + ((.Bottom - .Top) \ 2)
            End With
    
            With ltRectDialog
                cx = .Right - .Left
                cy = .Bottom - .Top
            End With
            
            x = x - (cx \ 2)
            y = y - (cy \ 2)
            
            With ltRectWorkArea
                If x + cx > .Right Then x = .Right - cx
                If y + cy > .Bottom Then y = .Bottom - cy
                If x < .Left Then x = .Left
                If y < .Top Then y = .Top
            End With
                
            MoveWindow hWnd, x, y, cx, cy, 1
        
        End If
    End Sub

#End If

#If bFileDialog Or bFolderDialog Or bPrintDialog Then
    
    Private Function pAllocateString(ByVal iLen As Long, Optional ByVal bBytes As Boolean) As String
        #If bVBVMTypeLib Then
            If bBytes _
                Then StringPtr(pAllocateString) = SysAllocStringByteLen(ZeroL, iLen) _
                Else StringPtr(pAllocateString) = SysAllocStringByteLen(ZeroL, iLen + iLen)
        #Else
            If bBytes _
                Then CopyMemory ByVal VarPtr(pAllocateString), SysAllocStringByteLen(ZeroL, iLen), 4& _
                Else CopyMemory ByVal VarPtr(pAllocateString), SysAllocStringByteLen(ZeroL, iLen + iLen), 4&
        #End If
    End Function
    
#End If

#If bFileDialog Then
    
    #If bHookDialogs Then
        Public Function File_ShowSave( _
           Optional ByRef sReturnFileName As String, _
           Optional ByVal iFlags As eFileDialog = dlgFileExplorerStyle, _
           Optional ByVal sFilter As String = "All Files (*.*)|*.*", _
           Optional ByVal iFilterIndex As Long = 1, _
           Optional ByVal sDefaultExt As String, _
           Optional ByVal sInitPath As String, _
           Optional ByVal sInitFile As String, _
           Optional ByVal sTitle As String, _
           Optional ByVal hWndOwner As Long, _
           Optional ByVal vTemplate As Variant, _
           Optional ByVal hInstance As Long, _
           Optional ByVal oEventSink As iComDlgHook, _
           Optional ByRef iReturnFlags As Long, _
           Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
           Optional ByRef iReturnFilterIndex As Long) _
                        As Boolean
            
            Dim ltDialog As tFileDialog
            pFile_GetUDT ltDialog, sTitle, iFlags, sFilter, iFilterIndex, sDefaultExt, sInitPath, sInitFile, hWndOwner, vTemplate, hInstance, oEventSink
            
    #Else
        Public Function File_ShowSave( _
           Optional ByRef sReturnFileName As String, _
           Optional ByVal iFlags As eFileDialog = dlgFileExplorerStyle, _
           Optional ByVal sFilter As String = "All Files (*.*)|*.*", _
           Optional ByVal iFilterIndex As Long = 1, _
           Optional ByVal sDefaultExt As String, _
           Optional ByVal sInitPath As String, _
           Optional ByVal sInitFile As String, _
           Optional ByVal sTitle As String, _
           Optional ByVal hWndOwner As Long, _
           Optional ByVal vTemplate As Variant, _
           Optional ByVal hInstance As Long, _
           Optional ByRef iReturnFlags As Long, _
           Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
           Optional ByRef iReturnFilterIndex As Long) _
                        As Boolean
    
            Dim ltDialog As tFileDialog
            pFile_GetUDT ltDialog, sTitle, iFlags, sFilter, iFilterIndex, sDefaultExt, sInitPath, sInitFile, hWndOwner, vTemplate, hInstance
            
    #End If
       
        File_ShowSave = pFile_Show(ltDialog, True)
        
        With ltDialog
            iReturnExtendedError = .iReturnExtendedError
            iReturnFilterIndex = .iReturnFilterIndex
            iReturnFlags = .iReturnFlags
            sReturnFileName = .sReturnFileName
        End With
        
    End Function
    
    
    #If bHookDialogs Then
        Public Function File_ShowOpen( _
           Optional ByRef sReturnFileName As String, _
           Optional ByVal iFlags As eFileDialog = dlgFileExplorerStyle Or dlgFileMustExist Or dlgFilePathMustExist Or dlgFileHideReadOnly, _
           Optional ByVal sFilter As String = "All Files (*.*)|*.*", _
           Optional ByVal iFilterIndex As Long = 1, _
           Optional ByVal sDefaultExt As String, _
           Optional ByVal sInitPath As String, _
           Optional ByVal sInitFile As String, _
           Optional ByVal sTitle As String, _
           Optional ByVal hWndOwner As Long, _
           Optional ByVal vTemplate As Variant, _
           Optional ByVal hInstance As Long, _
           Optional ByVal oEventSink As iComDlgHook, _
           Optional ByRef iReturnFlags As Long, _
           Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
           Optional ByRef iReturnFilterIndex As Long) _
                        As Boolean
            
            Dim ltDialog As tFileDialog
            pFile_GetUDT ltDialog, sTitle, iFlags, sFilter, iFilterIndex, sDefaultExt, sInitPath, sInitFile, hWndOwner, vTemplate, hInstance, oEventSink
            
    #Else
        Public Function File_ShowOpen( _
           Optional ByRef sReturnFileName As String, _
           Optional ByVal iFlags As eFileDialog = dlgFileExplorerStyle Or dlgFileMustExist Or dlgFilePathMustExist Or dlgFileHideReadOnly, _
           Optional ByVal sFilter As String = "All Files (*.*)|*.*", _
           Optional ByVal iFilterIndex As Long = 1, _
           Optional ByVal sDefaultExt As String, _
           Optional ByVal sInitPath As String, _
           Optional ByVal sInitFile As String, _
           Optional ByVal sTitle As String, _
           Optional ByVal hWndOwner As Long, _
           Optional ByVal vTemplate As Variant, _
           Optional ByVal hInstance As Long, _
           Optional ByRef iReturnFlags As Long, _
           Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
           Optional ByRef iReturnFilterIndex As Long) _
                        As Boolean
    
            Dim ltDialog As tFileDialog
            pFile_GetUDT ltDialog, sTitle, iFlags, sFilter, iFilterIndex, sDefaultExt, sInitPath, sInitFile, hWndOwner, vTemplate, hInstance
            
    #End If
       
        File_ShowOpen = pFile_Show(ltDialog, False)
        
        With ltDialog
            iReturnExtendedError = .iReturnExtendedError
            iReturnFilterIndex = .iReturnFilterIndex
            iReturnFlags = .iReturnFlags
            sReturnFileName = .sReturnFileName
        End With
        
    End Function
        
    
    Public Function File_ShowOpenIndirect( _
                ByRef tFileDialog As tFileDialog) _
                    As Boolean
        File_ShowOpenIndirect = pFile_Show(tFileDialog, False)
    End Function
    
    Public Function File_ShowSaveIndirect( _
                ByRef tFileDialog As tFileDialog) _
                    As Boolean
        File_ShowSaveIndirect = pFile_Show(tFileDialog, True)
    End Function
    
    Public Function File_GetFilter(ParamArray vFilters() As Variant) As String
        Dim i As Long, j As Long
        Dim lvFilters As Variant
        
        For i = LBound(vFilters) To UBound(vFilters)
            lvFilters = Split(vFilters(i), dlgFileFilterDelim)
            For j = ZeroL To UBound(lvFilters)
                File_GetFilter = File_GetFilter & lvFilters(i) & vbNullChar
            Next
        Next
        
        File_GetFilter = File_GetFilter & vbNullChar
    End Function
    
    Public Function File_GetMultiFileNames(ByRef sReturnFileName As String, ByRef sReturnPath As String, ByRef sFileNames() As String) As Long
        Dim liTemp As Long
        Erase sFileNames
        sReturnPath = vbNullString
        
        liTemp = InStr(sReturnFileName, vbNullChar)
        If liTemp > ZeroL Then
            sReturnPath = Left$(sReturnFileName, liTemp - 1)
            If Right$(sReturnPath, 1) <> "\" Then sReturnPath = sReturnPath & "\"
            sFileNames = Split(Mid$(sReturnFileName, liTemp + 1), vbNullChar)
            File_GetMultiFileNames = UBound(sFileNames) + 1&
        Else
            liTemp = InStrRev(sReturnFileName, "\")
            If liTemp > ZeroL Then
                sReturnPath = Left$(sReturnFileName, liTemp)
                ReDim sFileNames(0 To 0)
                sFileNames(0) = Mid$(sReturnFileName, liTemp + 1&)
                File_GetMultiFileNames = 1&
            End If
        End If
    End Function
    
    
    #If bHookDialogs Then
        Private Sub pFile_GetUDT( _
                    ByRef tDialog As tFileDialog, _
                    ByRef sTitle As String, _
                    ByVal iFlags As eFileDialog, _
                    ByRef sFilter As String, _
                    ByVal iFilterIndex As Long, _
                    ByRef sDefaultExt As String, _
                    ByRef sInitPath As String, _
                    ByRef sInitFile As String, _
                    ByVal hWndOwner As Long, _
                    ByRef vTemplate As Variant, _
                    ByVal hInstance As Long, _
                    ByVal oEventSink As iComDlgHook)
    
    #Else
        Private Sub pFile_GetUDT( _
                    ByRef tDialog As tFileDialog, _
                    ByRef sTitle As String, _
                    ByVal iFlags As eFileDialog, _
                    ByRef sFilter As String, _
                    ByVal iFilterIndex As Long, _
                    ByRef sDefaultExt As String, _
                    ByRef sInitPath As String, _
                    ByRef sInitFile As String, _
                    ByVal hWndOwner As Long, _
                    ByRef vTemplate As Variant, _
                    ByVal hInstance As Long)
                    
    #End If
    
        With tDialog
            .iFlags = iFlags
            .hWndOwner = hWndOwner
            .sTitle = CStr(sTitle)
            .sInitFile = CStr(sInitFile)
            .sInitPath = CStr(sInitPath)
            .sDefaultExt = CStr(sDefaultExt)
            .hInstance = hInstance
            .vTemplate = vTemplate
            .sFilter = CStr(sFilter)
            .iFilterIndex = iFilterIndex
            
            #If bHookDialogs Then
                Set .oEventSink = oEventSink
            #End If
            
        End With
        
    End Sub
    
    Private Function pFile_GetFilterStr(ByRef sFilter As String) As String
        pFile_GetFilterStr = sFilter & vbNullChar & vbNullChar
        
        Dim i As Long
        
        Do
            i = InStr(1, pFile_GetFilterStr, dlgFileFilterDelim)
            If i > ZeroL _
                Then Mid$(pFile_GetFilterStr, i, 1) = vbNullChar _
                Else Exit Do
        Loop
    End Function
    
    Private Function pFile_Show( _
                ByRef tFileDialog As tFileDialog, _
                ByVal bSave As Boolean) _
                    As Boolean
    
    'in         iFlags                  As eFileDialog
    'in         sFilter                 As String
    'in         iFilterIndex            As Long
    'in         sDefaultExt             As String
    'in         sInitPath               As String
    'in         sInitFile               As String
    'in         sTitle                  As String
    'in         hWndOwner               As Long
    'in         vTemplate               As Variant
    'in         hInstance               As Long
    'in         oEventSink              As iComDlgHook
    '
    'out        sReturnFileName         As String
    'out        iReturnFlags            As Long
    'out        iReturnExtendedError    As eComDlgExtendedError
    'out        iReturnFilterIndex      As Long
    
    Const ValidFlags As Long = dlgFileExplorerStyle Or _
                               dlgFileMustExist Or _
                               dlgFilePathMustExist Or _
                               dlgFileMultiSelect Or _
                               dlgFilePromptToCreate Or _
                               dlgFileEnableSizing Or _
                               dlgFileNoDereferenceLinks Or _
                               dlgFileHideNetworkButton Or _
                               dlgFileHideReadOnly Or _
                               dlgFileNoReadOnlyReturn Or _
                               dlgFileNoTestFileCreate Or _
                               dlgFileOverwritePrompt Or _
                               dlgFileInitToReadOnly Or _
                               dlgFileShowHelpButton Or _
                               dlgFileEnableHook Or _
                               dlgFileEnableTemplate
       

        
        Dim ltOFN As OPENFILENAME
        Dim liTemp As Long
    
        With tFileDialog
            .iReturnExtendedError = ZeroL                                       'initialize return values
            .iReturnFlags = ZeroL
            .sReturnFileName = vbNullString
            .iReturnFilterIndex = NegOneL
        End With
        
        With ltOFN
            .lStructSize = Len(ltOFN)                                           'initialize api structure
            .Flags = tFileDialog.iFlags And ValidFlags
            .hWndOwner = tFileDialog.hWndOwner
            .lpstrInitialDir = tFileDialog.sInitPath & vbNullChar
            .lpstrDefExt = tFileDialog.sDefaultExt & vbNullChar
            .lpstrTitle = tFileDialog.sTitle & vbNullChar
            
            .lpfnHook = 0&
            
            #If bHookDialogs Then                                               'if hooking is supported
                Dim liHookProc As Long
                       
                If CBool(tFileDialog.iFlags And dlgFileEnableHook) Then         'if hooking is requested
                    liHookProc = pAllocateHookProc(tFileDialog.oEventSink)      'allocate the machine code
                    .lpfnHook = liHookProc                                      'point to it
                    'no callback interface or out of memory
                    Debug.Assert liHookProc
                End If
                
            #End If
            
            If .lpfnHook = 0& Then .Flags = .Flags And Not dlgFileEnableHook
            
            If CBool(tFileDialog.iFlags And dlgFileEnableTemplate) Then         'if template is requested
                Dim lsTemplate As String
                If VarType(tFileDialog.vTemplate) = vbString Then               'if it's a template name
                    lsTemplate = CStr(tFileDialog.vTemplate) & vbNullChar       'store the name
                    .lpTemplateName = StrPtr(lsTemplate)                        'point to it
                Else
                    On Error Resume Next                                        'else-assume it's a numeric id
                    .lpTemplateName = CLng(tFileDialog.vTemplate)
                    On Error GoTo 0
                End If
                
                .hInstance = tFileDialog.hInstance                              'store the instance handle
                
                If .hInstance = 0& Or .lpTemplateName = 0& Then                 'make sure we have valid data
                    'flag specified, but no template data
                    Debug.Assert False                                          'if invalid, don't use the template
                    .hInstance = 0&
                    .lpTemplateName = 0&
                    .Flags = .Flags And Not dlgFileEnableTemplate
                End If

            End If
            
            .lpstrFilter = pFile_GetFilterStr(tFileDialog.sFilter)              'get the api filter
            .nFilterIndex = tFileDialog.iFilterIndex
            
            If CBool(tFileDialog.iFlags And dlgFileMultiSelect) Then            'if multiselect
                .lpstrFile = pAllocateString(8192&, True)                       'allocate a large string
                .nMaxFile = 8192&
            Else                                                                'otherwise
                .lpstrFile = pAllocateString(MAX_PATH, True)                    'allocate MAX_PATH
                .nMaxFile = MAX_PATH
            End If
            
            liTemp = Len(tFileDialog.sInitFile)
            
            If liTemp Then                                                      'store the initial file
                If liTemp > .nMaxFile Then liTemp = .nMaxFile
                Mid$(.lpstrFile, 1, liTemp) = tFileDialog.sInitFile
            End If
            
            If liTemp < .nMaxFile Then                                          'null terminate the initial file
                Mid$(.lpstrFile, liTemp + OneL, OneL) = ChrW$(ZeroL)
            End If
            
            .lpstrCustomFilter = vbNullString                                   'make it obvious we don't use these members
            .nMaxCustFilter = ZeroL
            .lpstrFileTitle = vbNullString
            .nMaxFileTitle = ZeroL
            .nFileOffset = ZeroL
            .nFileExtension = ZeroL
            .lCustData = ZeroL
            
            Dim liReturn As Long
            
            If bSave _
                Then liReturn = GetSaveFileName(ltOFN) _
                Else liReturn = GetOpenFileName(ltOFN)                          'call the api
            
            #If bHookDialogs Then                                               'if hooking is supported
                If liHookProc Then GlobalFree liHookProc                        'free the machine code, if any
            #End If
            
            If liReturn = 1& Then                                               'if success
    
                pFile_Show = True                                               'return true
                tFileDialog.iReturnFlags = .Flags                               'return the flags
                tFileDialog.iReturnFilterIndex = .nFilterIndex                  'return the filter index
                
                If CBool(tFileDialog.iFlags And dlgFileMultiSelect) Then        'if multiselect
                    If CBool(tFileDialog.iFlags And dlgFileExplorerStyle) Then  'if explorer style
                        liTemp = InStr(1, .lpstrFile, vbNullChar & vbNullChar)  'return the file names (already null separated)
                        If liTemp > ZeroL Then
                            tFileDialog.sReturnFileName = Left$(.lpstrFile, liTemp - OneL)
                        Else
                            'should never happen!
                            Debug.Assert False
                            tFileDialog.sReturnFileName = .lpstrFile
                        End If
                    Else                                                        'if multiselect and not explorerstyle
                        liTemp = InStr(1, .lpstrFile, vbNullChar)               'return the file names, (space separated -> null separated)
                        If liTemp > ZeroL Then
                            tFileDialog.sReturnFileName = Left$(.lpstrFile, liTemp - OneL)
                        Else
                            'should never happen!
                            Debug.Assert False
                            tFileDialog.sReturnFileName = .lpstrFile
                        End If
                        tFileDialog.sReturnFileName = Replace$(tFileDialog.sReturnFileName, " ", vbNullChar)
                    End If
                    
                Else                                                            'if not multiselect
                    liTemp = InStr(1, .lpstrFile, vbNullChar)                   'return the file name
                    If liTemp > ZeroL _
                        Then tFileDialog.sReturnFileName = Left$(.lpstrFile, liTemp - OneL) _
                        Else tFileDialog.sReturnFileName = .lpstrFile
                End If
            
            Else                                                                'if api failed (cancel or error)
                
                If liReturn <> ZeroL Then                                       'if errored
                    tFileDialog.iReturnExtendedError = CommDlgExtendedError()
                End If
                
            End If
            
        End With
        
    End Function
    
#End If

#If bFolderDialog Then

    #If bHookDialogs Then
        Public Function Folder_Show( _
           Optional ByRef sReturnPath As String, _
           Optional ByVal iFlags As eFolderDialog = dlgFolderUseNewUI Or dlgFolderStatusText, _
           Optional ByVal sTitle As String, _
           Optional ByVal sInitialPath As String, _
           Optional ByVal sRootPath As String, _
           Optional ByVal hWndOwner As Long, _
           Optional ByVal oEventSink As iComDlgHook) _
                As Boolean

    #Else
        Public Function Folder_Show( _
           Optional ByRef sReturnPath As String, _
           Optional ByVal iFlags As eFolderDialog = dlgFolderUseNewUI Or dlgFolderStatusText, _
           Optional ByVal sTitle, _
           Optional ByVal sInitialPath, _
           Optional ByVal sRootPath, _
           Optional ByVal hWndOwner As Long) _
                As Boolean
    
    #End If
    
        Dim tFolder As tFolderDialog
        With tFolder
            .iFlags = iFlags
            .sTitle = sTitle
            .sInitialPath = sInitialPath
            .sRootPath = sRootPath
            .hWndOwner = hWndOwner
            #If bHookDialogs Then
            Set .oEventSink = oEventSink
            #End If
        End With
        
        Folder_Show = Folder_ShowIndirect(tFolder)
        
        sReturnPath = tFolder.sReturnPath
    
    End Function
    
    Public Function Folder_ShowIndirect(ByRef tFolderDialog As tFolderDialog) As Boolean
    
'        in         hWndOwner               As Long
'        in         sTitle                  As String
'        in         sInitialPath            As String
'        in         sRootPath               As String
'        out        sReturnPath             As String
'        reserved   iPidlInitial            As Long
'        reserved   hDlg                    As Long
'        in         iFlags                  As eFolderDialog
'                   #If bHookDialogs Then
'        in         oEventSink              As iComDlgHook
'                   #End If
    
    Const ValidFlags As Long = _
            dlgFolderReturnOnlyFSDirs Or _
            dlgFolderDontGoBelowDomain Or _
            dlgFolderStatusText Or _
            dlgFolderReturnFSAncestors Or _
            dlgFolderEditBox Or _
            dlgFolderValidate Or _
            dlgFolderUseNewUI Or _
            dlgFolderBrowseForComputer Or _
            dlgFolderBrowseForPrinter Or _
            dlgFolderBrowseIncludeFiles
        
        Dim tBI As BrowseInfo
        Dim lsTitle As String
        Dim pIDLRoot As Long
        Dim pidlInitial As Long
        Dim pidlOut As Long
        Dim sPath As String
        
        lsTitle = StrConv(tFolderDialog.sTitle, vbFromUnicode)                          'store an ansi version of the title
        
        tBI.hWndOwner = tFolderDialog.hWndOwner                                         'initialize the api structure
        tBI.lpszTitle = StrPtr(lsTitle)
        tBI.ulFlags = tFolderDialog.iFlags And ValidFlags
        
        If Len(tFolderDialog.sRootPath) <> ZeroL Then                                   'if there is a root path
            ' Get a PIDL for the selected path:
            pIDLRoot = pFolder_PathToPidl(tFolderDialog.sRootPath)                       'get the pidl
        End If
        
        tBI.pIDLRoot = pIDLRoot                                                         'store the root pidl if any
        
        If Len(tFolderDialog.sInitialPath) <> ZeroL Then                                'if there is an initial path
            pidlInitial = pFolder_PathToPidl(tFolderDialog.sInitialPath)
        End If
        
        tFolderDialog.iPidlInitial = pidlInitial                                        'store the initial pidl if any
        
        tBI.lpfn = 0&
        
        #If bHookDialogs Then                                                           'if hooking is supported
            Dim liHookProc As Long
            
            If CBool(tFolderDialog.iFlags And dlgFolderEnableHook) Then                 'if hooking is requested
                liHookProc = pAllocateHookProc(tFolderDialog.oEventSink)                'allocate the machine code
                tBI.lpfn = liHookProc
                'no callback interface or out of memory
                Debug.Assert liHookProc
            End If
        #End If
        
        If tBI.lpfn = 0& Then tBI.ulFlags = tBI.ulFlags And Not dlgFolderEnableHook
        
        If CBool(tFolderDialog.iFlags And dlgFolderUseNewUI) Then CoInitialize ZeroL, ZeroL
    
        tBI.lParam = ZeroL                                                              'make it obvious we don't use these members
        tBI.iImage = ZeroL
        tBI.pszDisplayName = 0&
        
        pidlOut = SHBrowseForFolder(tBI)                                                'call the api
        
        #If bHookDialogs Then
            If liHookProc Then GlobalFree liHookProc                                    'if we allocated a hook proc, free it
        #End If
        
        tFolderDialog.sReturnPath = Folder_PathFromPidl(pidlOut)                        'return the selected path
        Folder_ShowIndirect = CBool(Len(tFolderDialog.sReturnPath))                     'indicate success or failure
        
        With pFolder_Allocator                                                          'Free the pidls we create
            If pIDLRoot <> ZeroL Then
                .Free ByVal pIDLRoot
            End If
            If pidlInitial <> ZeroL Then
                .Free ByVal pidlInitial
            End If
            If pidlOut <> ZeroL Then
                .Free ByVal pidlOut
            End If
        End With
        
        tFolderDialog.iPidlInitial = ZeroL
        
    End Function
    
    Public Function Folder_SpecialFolder(ByVal hWndOwner As Long, ByVal iFolder As eSpecialFolders) As String
        Dim pidl As Long
        On Error Resume Next
        ' Get pidl of special folder:
        SHGetSpecialFolderLocation hWndOwner, iFolder, pidl
        If Err = 0 Then
            ' Convert it to a path:
            Folder_SpecialFolder = Folder_PathFromPidl(pidl)
            ' Free the pidl:
            pFolder_Allocator.Free ByVal pidl
    
        End If
    End Function
    
    Public Function Folder_SetFolder(ByVal hDlg As Long, ByVal sPath As String) As Boolean
        Dim pidl As Long
        If hDlg Then
            pidl = pFolder_PathToPidl(sPath)
            SendMessageLong hDlg, BFFM_SETSELECTIONA, 0, pidl
            pFolder_Allocator.Free ByVal pidl
            SetFocusAPI hDlg
            Folder_SetFolder = True
        End If
    End Function
    
    Public Function Folder_SetFolderPidl(ByVal hDlg As Long, ByVal iPidl As Long) As Boolean
        If hDlg Then
            SendMessageLong hDlg, BFFM_SETSELECTIONA, 0, iPidl
            SetFocusAPI hDlg
            Folder_SetFolderPidl = True
        End If
    End Function
    
    Private Function pFolder_PathToPidl(ByRef sPath As String) As Long
    Dim Folder As IShellFolder
    Dim pidlMain As Long
    Dim cParsed As Long
    Dim afItem As Long
    Dim lFilePos As Long
    Dim lR As Long
    Dim sRet As String
    
       ' Make sure the file name is fully qualified
       sRet = pAllocateString(MAX_PATH, False)
       lR = GetFullPathName(sPath, MAX_PATH, sRet, lFilePos)
       If lR Then
          ' debug.Assert c <= cMaxPath
          sPath = Left$(sRet, lR)
    
          ' Convert the path name into a pointer to an item ID list (pidl)
          Set Folder = pFolder_GetDesktopFolder
          ' Will raise an error if path cannot be found:
          If (Folder.ParseDisplayName(ZeroL, ZeroL, StrConv(sPath, vbUnicode), cParsed, pidlMain, afItem)) <= 0& Then
             pFolder_PathToPidl = pidlMain
          End If
       End If
    
    End Function
    
    Public Function Folder_PathFromPidl(ByVal pidl As Long) As String
    Dim sPath As String * MAX_PATH
    Dim lR As Long
       lR = SHGetPathFromIDList(pidl, sPath)
       If lR Then
          Folder_PathFromPidl = Left$(sPath, lstrlen(sPath))
       End If
    End Function
    
    Public Function Folder_PtrToString(ByVal lptr As Long) As String
    Dim lLen As Long
        ' Get length of Unicode string to first null
        lLen = lstrlenptr(lptr)
        ' Allocate a string of that length
        Folder_PtrToString = pAllocateString(lLen)
        ' Copy the pointer data to the string
        CopyMemoryLpToStr Folder_PtrToString, ByVal lptr, lLen
    End Function
    
    Public Function Folder_SetStatus(ByVal hDlg As Long, ByRef sText As String) As Boolean
        If hDlg Then
            SendMessage hDlg, BFFM_SETSTATUSTEXTA, ZeroL, ByVal sText
            Folder_SetStatus = True
        End If
    End Function
    
    Public Function Folder_EnableOK(ByVal hDlg As Long, ByVal bVal As Boolean) As Boolean
        If hDlg Then
            SendMessageLong hDlg, BFFM_ENABLEOK, 0, Abs(bVal)
            Folder_EnableOK = True
        End If
    End Function
    
    
    Private Property Get pFolder_Allocator() As IMalloc
        Static oAlloc As IMalloc
        If oAlloc Is Nothing Then SHGetMalloc oAlloc
    
        Set pFolder_Allocator = oAlloc
    End Property
    
    Private Function pFolder_GetDesktopFolder() As IShellFolder
        SHGetDesktopFolder pFolder_GetDesktopFolder
    End Function



#End If

#If bColorDialog Then
    
    #If bHookDialogs Then
    
        Public Function Color_Show( _
                        ByRef iColor As Long, _
               Optional ByVal iFlags As eColorDialogFlag = dlgColorAny, _
               Optional ByVal hWndOwner As Long, _
               Optional ByVal sTitle As String = "Choose Color", _
               Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
               Optional ByVal oEventSink As iComDlgHook) _
                            As Boolean
    #Else
    
        Public Function Color_Show( _
                        ByRef iColor As OLE_COLOR, _
               Optional ByVal iFlags As eColorDialogFlag = dlgColorAny, _
               Optional ByVal hWndOwner As Long, _
               Optional ByVal sTitle As String = "Choose Color", _
               Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                            As Boolean
    
    #End If
                        
        Dim ltDialog As tColorDialog
        With ltDialog
            .iColor = iColor
            .iFlags = iFlags
            .hWndOwner = hWndOwner
            
            #If bHookDialogs Then
                Set .oEventSink = oEventSink
            #End If
            
            Dim i As Long
            For i = 0 To 15
                .iColors(i) = QBColor(i)
            Next
            
            Color_Show = Color_ShowIndirect(ltDialog)
            
            If Color_Show Then
                iColor = .iColor
            End If
            
            iReturnExtendedError = .iReturnExtendedError
            
        End With
            
    End Function

    Public Function Color_ShowIndirect(ByRef tColorDialog As tColorDialog) As Boolean

'       in          hWndOwner               As Long
'       in/out      iColor                  As OLE_COLOR
'       in          iFlags                  As eColorDialogFlag
'       out         iReturnExtendedError          As eComDlgExtendedError
'                   #If bHookDialogs Then
'       in          oEventSink              As iComDlgHook
'                   #End If
'       in/out      iColors(0 To 15)        As Long
    
    Const ValidFlags As Long = dlgColorFullOpen Or dlgColorPreventFullOpen Or dlgColorSolid Or dlgColorAny Or dlgColorEnableHook
    
        Dim ltChooseColor As TCHOOSECOLOR
    
        tColorDialog.iReturnExtendedError = ZeroL
    
        With ltChooseColor
            .lStructSize = Len(ltChooseColor)                                           'initialize the api structure
            .hWndOwner = tColorDialog.hWndOwner
            If OleTranslateColor(tColorDialog.iColor, 0&, .rgbResult) Then .rgbResult = NegOneL
            .Flags = (tColorDialog.iFlags And ValidFlags) Or Abs(CBool(.rgbResult <> NegOneL))
            .lpCustColors = VarPtr(tColorDialog.iColors(0))
            
            .lpfnHook = 0&
            
            #If bHookDialogs Then                                                       'if hooking is supported
                Dim liHookProc As Long
                If CBool(.Flags And dlgColorEnableHook) Then                            'if hooking is requested
                    liHookProc = pAllocateHookProc(tColorDialog.oEventSink)             'allocate the machine code
                    'no callback interface or out of memory
                    Debug.Assert liHookProc
                    .lpfnHook = liHookProc                                              'store the pointer
                End If
            #End If
            
            If .lpfnHook = 0& Then .Flags = .Flags And Not dlgColorEnableHook           'if no hook then don't enable it
            
            .hInstance = ZeroL                                                          'make it obvious we don't use these members
            .lCustData = ZeroL
            .lpTemplateName = ZeroL
            
            Dim liReturn As Long
            
            liReturn = ChooseColor(ltChooseColor)                                       'call te api
            
            #If bHookDialogs Then
                If liHookProc Then GlobalFree liHookProc                                'if we allocated the procedure, free it
            #End If
            
            If liReturn = OneL Then                                                     'if succeeded
                Color_ShowIndirect = True
                tColorDialog.iColor = .rgbResult                                        'return the color
            Else
                If liReturn <> ZeroL Then                                               'if errored
                    tColorDialog.iReturnExtendedError = CommDlgExtendedError()                'store the error value
                End If
            End If
        End With
    
    End Function
    
    Public Property Get Color_OKMsg() As Long
        Static b As Boolean
        Static iMsg As Long
        Const COLOROKSTRING As String = "commdlg_ColorOK"
    
        If Not b Then
            iMsg = RegisterWindowMessage(COLOROKSTRING)
            b = True
        End If
        Color_OKMsg = iMsg
    End Property
    
#End If

#If bFontDialog Then
    
    #If bHookDialogs Then
        Public Function Font_Show( _
                    Optional ByVal oFont As Object, _
                    Optional ByVal iFlags As eFontDialog = dlgFontScreenFonts Or dlgFontEffects, _
                    Optional ByVal hdc As Long, _
                    Optional ByVal hWndOwner As Long, _
                    Optional ByVal iMinSize As Long = 6, _
                    Optional ByVal iMaxSize As Long = 72, _
                    Optional ByRef iColor As Long, _
                    Optional ByRef iReturnFlags As eFontDialog, _
                    Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
                    Optional ByVal oEventSink As iComDlgHook) _
                        As Boolean

    #Else
        Public Function Font_Show( _
                    Optional ByVal oFont As Object, _
                    Optional ByVal iFlags As eFontDialog = dlgFontScreenFonts Or dlgFontEffects, _
                    Optional ByVal hdc As Long, _
                    Optional ByVal hWndOwner As Long, _
                    Optional ByVal iMinSize As Long = 6, _
                    Optional ByVal iMaxSize As Long = 72, _
                    Optional ByRef iColor As Long, _
                    Optional ByRef iReturnFlags As eFontDialog, _
                    Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                        As Boolean
        
    #End If
        
        Dim ltDialog As tFontDialog
        
        With ltDialog
            #If bHookDialogs Then
                Set .oEventSink = oEventSink
            #End If
            Set .oFont = oFont
            .iFlags = iFlags
            .hdc = hdc
            .hWndOwner = hWndOwner
            .iMinSize = iMinSize
            .iMaxSize = iMaxSize
            .iColor = iColor
        End With
        
        Font_Show = Font_ShowIndirect(ltDialog)
                
        If Font_Show Then
            iColor = ltDialog.iColor
            iReturnFlags = ltDialog.iReturnFlags
        End If
        
        Set oFont = ltDialog.oFont
        
        iReturnExtendedError = ltDialog.iReturnExtendedError
        
    End Function
    
    Public Function Font_ShowIndirect(ByRef tFontDialog As tFontDialog) As Boolean

'        in         iFlags                  As eFontDialog
'        in         hdc                     As Long
'        in         hWndOwner               As Long
'        in         iMinSize                As Long
'        in         iMaxSize                As Long
'        in/out     oLogFont                as cLogFont
'        #If bHookDialogs Then
'        in         oEventSink              As iComDlgHook
'        #End If
'        out        iColor                  As OLE_COLOR
'        out        iReturnFlags            As eFontDialog
'        out        iReturnExtendedError          As eComDlgExtendedError
        
        Const ValidFlags As Long = _
            dlgFontScreenFonts Or _
            dlgFontPrinterFonts Or _
            dlgFontScreenAndPrinterFonts Or _
            dlgFontUseStyle Or _
            dlgFontEffects Or _
            dlgFontAnsiOnly Or _
            dlgFontNoVectorFonts Or _
            dlgFontNoSimulations Or _
            dlgFontFixedPitchOnly Or _
            dlgFontWysiwyg Or _
            dlgFontForceExist Or _
            dlgFontScalableOnly Or _
            dlgFontTTOnly Or _
            dlgFontNoFaceSel Or _
            dlgFontNoStyleSel Or _
            dlgFontNoSizeSel Or _
            dlgFontSelectScript Or _
            dlgFontNoScriptSel Or _
            dlgFontNoVertFonts Or _
            dlgFontApply Or _
            dlgFontEnableHook
        
        Dim ltChooseFont As TCHOOSEFONT
        
        Dim ltLF As LOGFONT
        
        With tFontDialog
            If .oFont Is Nothing Then Set .oFont = New StdFont
            
            If TypeOf .oFont Is StdFont Then
                pFont_PutStdFont ltLF, .oFont
        #If bDependcFont Then
            ElseIf TypeOf .oFont Is cFont Then
                pFont_PutFont ltLF, .oFont
        #End If
            End If
            
            .iReturnFlags = ZeroL                                                                   'initialize the return values
            .iReturnExtendedError = ZeroL
        End With
    
        With ltChooseFont
            .lStructSize = Len(ltChooseFont)                                                        'initialize the api structure
            .hWndOwner = tFontDialog.hWndOwner
            .hdc = tFontDialog.hdc
            .lpLogFont = VarPtr(ltLF)
            .Flags = (tFontDialog.iFlags And ValidFlags) Or dlgFontInitToLogFontStruct
            If tFontDialog.iMinSize > ZeroL Or tFontDialog.iMaxSize > ZeroL Then .Flags = .Flags Or dlgFontLimitSize
            .rgbColors = tFontDialog.iColor
            .nSizeMin = tFontDialog.iMinSize
            .nSizeMax = tFontDialog.iMaxSize
        End With
    
        ltChooseFont.lpfnHook = 0&
        
        #If bHookDialogs Then                                                                       'if hooking is supported
            Dim liHookProc As Long
            If CBool(tFontDialog.iFlags And dlgFontEnableHook) Then                                 'if hooking is requested
                liHookProc = pAllocateHookProc(tFontDialog.oEventSink)                              'allocate the machine code
                'no callback interface or out of memory
                Debug.Assert liHookProc
                ltChooseFont.lpfnHook = liHookProc                                                  'store the pointer
            End If
        #End If
        
        If ltChooseFont.lpfnHook = 0& _
            Then ltChooseFont.Flags = ltChooseFont.Flags And Not dlgFontEnableHook                  'if no procedure, don't enable the hook
        
        With ltChooseFont                                                                           'make it obvious we don't use these members
            .iPointSize = ZeroL
            .lCustData = ZeroL
            .lpTemplateName = ZeroL
            .hInstance = ZeroL
            .lpszStyle = vbNullString
            .nFontType = 0
            .iAlign = 0
        End With
        
        Dim liReturn As Long
        
        liReturn = ChooseFont(ltChooseFont)
        
        #If bHookDialogs Then
            If liHookProc Then GlobalFree liHookProc                                                'if we allocated a procedure, free it
        #End If

        If liReturn = 1& Then                                                                       'if we succeeded
            Font_ShowIndirect = True
            With tFontDialog
                If TypeOf .oFont Is StdFont Then
                    pFont_GetStdFont ltLF, .oFont
            #If bDependcFont Then
                ElseIf TypeOf .oFont Is cFont Then
                    pFont_GetFont ltLF, .oFont
            #End If
                Else
                    'unknown font object, can't return the font chosen!
                    Debug.Assert False
                End If
                
                .iReturnFlags = ltChooseFont.Flags                                                  'return the flags
                .iColor = ltChooseFont.rgbColors                                                    'return the color
            End With
        Else
            If liReturn <> ZeroL Then
                tFontDialog.iReturnExtendedError = CommDlgExtendedError()                                 'return the extended error
            End If
        End If
        
    End Function
    
    Private Property Let pFont_FaceName(ByRef tLogFont As LOGFONT, ByRef sName As String)
        
        Dim ls As String
        Dim iLen As Long
        
        ls = StrConv(sName, vbFromUnicode)
        
        iLen = LenB(ls)
        If iLen > LF_FACESIZE Then iLen = LF_FACESIZE
        
        If iLen > 0& Then CopyMemory tLogFont.lfFaceName(0), ByVal StrPtr(ls), iLen
        If iLen < LF_FACESIZE _
            Then ZeroMemory tLogFont.lfFaceName(iLen), (LF_FACESIZE - iLen) _
            Else tLogFont.lfFaceName(LF_FACESIZE - OneL) = ZeroY
        
    End Property
    
    Private Property Get pFont_FaceName(ByRef tLogFont As LOGFONT) As String
        pFont_FaceName = StrConv(tLogFont.lfFaceName, vbUnicode)
        Dim i As Long
        i = InStr(1&, pFont_FaceName, vbNullChar)
        If i Then pFont_FaceName = Left$(pFont_FaceName, i - 1&)
    End Property
    
    Private Sub pFont_PutStdFont(ByRef tLogFont As LOGFONT, ByVal oFont As StdFont)
        On Error Resume Next
        pFont_FaceName(tLogFont) = oFont.Name
        With tLogFont
            .lfHeight = -MulDiv(oFont.Size, 1440& / Screen.TwipsPerPixelY, 72&)
            .lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
            .lfItalic = Abs(oFont.Italic)
            .lfUnderline = Abs(oFont.Underline)
            .lfStrikeOut = Abs(oFont.Strikethrough)
            .lfCharSet = oFont.Charset And &HFF
            .lfEscapement = 0&
            .lfOrientation = 0&
            .lfOutPrecision = 0
            .lfClipPrecision = 0
            .lfQuality = 0
            .lfPitchAndFamily = 0
        End With
        On Error GoTo 0
    End Sub
    
    Private Sub pFont_GetStdFont(ByRef tLogFont As LOGFONT, ByVal oFont As StdFont)
        On Error Resume Next
        With oFont
            .Name = pFont_FaceName(tLogFont)
            If tLogFont.lfHeight Then
                .Size = MulDiv(72&, Abs(tLogFont.lfHeight), (1440& / Screen.TwipsPerPixelY))
            End If
            .Charset = tLogFont.lfCharSet
            .Italic = CBool(tLogFont.lfItalic)
            .Underline = CBool(tLogFont.lfUnderline)
            .Strikethrough = CBool(tLogFont.lfStrikeOut)
            .Bold = CBool(tLogFont.lfWeight > FW_NORMAL)
        End With
        On Error GoTo 0
    End Sub
    
    #If bDependcFont Then
        Private Sub pFont_GetFont(ByRef tLogFont As LOGFONT, ByVal oFont As cFont)
            oFont.fPutLogFontLong VarPtr(tLogFont)
        End Sub
        
        Private Sub pFont_PutFont(ByRef tLogFont As LOGFONT, ByVal oFont As cFont)
            oFont.fGetLogFontLong VarPtr(tLogFont)
        End Sub
    #End If
   
#End If

#If bPrintDialog Then

    #If bHookDialogs Then

        Public Function Print_Show( _
                    Optional ByRef hdc As Long, _
                    Optional ByVal iFlags As ePrintDialog, _
                    Optional ByVal hWndOwner As Long, _
                    Optional ByRef iRange As ePrintRange, _
                    Optional ByRef iFromPage As Long = 1, _
                    Optional ByRef iToPage As Long = 1, _
                    Optional ByVal iMinPage As Long = 1, _
                    Optional ByVal iMaxPage As Long = 1, _
                    Optional ByRef oDeviceMode As cDeviceMode, _
                    Optional ByRef sDevice As String, _
                    Optional ByRef sDriver As String, _
                    Optional ByRef sOutputPort As String, _
                    Optional ByRef iReturnFlags As ePrintDialog, _
                    Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
                    Optional ByRef bCollate As Boolean, _
                    Optional ByRef bPrintToFile As Boolean, _
                    Optional ByVal oEventSink As iComDlgHook) _
                        As Boolean
    #Else
        Public Function Print_Show( _
                    Optional ByRef hdc As Long, _
                    Optional ByVal iFlags As ePrintDialog, _
                    Optional ByVal hWndOwner As Long, _
                    Optional ByRef iRange As ePrintRange, _
                    Optional ByRef iFromPage As Long = 1, _
                    Optional ByRef iToPage As Long = 1, _
                    Optional ByVal iMinPage As Long = 1, _
                    Optional ByVal iMaxPage As Long = 1, _
                    Optional ByRef oDeviceMode As cDeviceMode, _
                    Optional ByRef sDevice As String, _
                    Optional ByRef sDriver As String, _
                    Optional ByRef sOutputPort As String, _
                    Optional ByRef iReturnFlags As ePrintDialog, _
                    Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
                    Optional ByRef bCollate As Boolean, _
                    Optional ByRef bPrintToFile As Boolean) _
                        As Boolean
    
    
    #End If
        Dim ltDialog As tPrintDialog
        
        With ltDialog
            .hWndOwner = hWndOwner
            .iFlags = iFlags
            .iRange = iRange
            .iMinPage = iMinPage
            .iMaxPage = iMaxPage
            If oDeviceMode Is Nothing Then Set oDeviceMode = New cDeviceMode
            Set .oDeviceMode = oDeviceMode
            .iFromPage = iFromPage
            .iToPage = iToPage
            #If bHookDialogs Then
                Set .oEventSink = oEventSink
            #End If
            
            Print_Show = Print_ShowIndirect(ltDialog)
            
            If Print_Show Then
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
            End If
            
            
        End With

    End Function

    Public Function Print_ShowIndirect(ByRef tDialog As tPrintDialog) As Boolean
    
'        out        hdc                     As Long
'        in         hWndOwner               As Long
'        in         iFlags                  As ePrintDialog
'        in/out     iRange                  As ePrintRange
'        in         iMinPage                As Long
'        in         iMaxPage                As Long
'
'        in/out     oDeviceMode             As cDeviceMode
'
'        in/out     iFromPage               As Long
'        in/out     iToPage                 As Long
'        out        iReturnFlags            As ePrintDialog
'
'        out        iReturnExtendedError    As eComDlgExtendedError
'        out        bCollate                As Boolean
'        out        bPrintToFile            As Boolean
'
'        out        sDriver                 As String
'        out        sDevice                 As String
'        out        sOutputPort             As String
'                   #If bHookDialogs Then
'        in         oEventSink              As iComDlgHook
'                   #End If
        
        Const ValidFlags As Long = _
                    dlgPrintAllPages Or _
                    dlgPrintSelection Or _
                    dlgPrintPageNums Or _
                    dlgPrintNoSelection Or _
                    dlgPrintNoPageNums Or _
                    dlgPrintCollate Or _
                    dlgPrintToFile Or _
                    dlgPrintSetup Or _
                    dlgPrintNoWarning Or _
                    dlgPrintReturnDc Or _
                    dlgPrintReturnIc Or _
                    dlgPrintReturnDefault Or _
                    dlgPrintShowHelp Or _
                    dlgPrintEnablePrintHook Or _
                    dlgPrintEnableSetupHook Or _
                    dlgPrintDisablePrintToFile Or _
                    dlgPrintHidePrintToFile Or _
                    dlgPrintNoNetworkButton
        
        Dim liOr As ePrintDialog
        Dim hDevMode As Long
    
        tDialog.iReturnExtendedError = ZeroL                                                    'initialize return values
        tDialog.iReturnFlags = ZeroL
    
        ' Set PRINTDLG flags
        If tDialog.iRange = dlgPrintRangePageNumbers Then                                       'get the print range flag
            liOr = dlgPrintPageNums
        ElseIf tDialog.iRange = dlgPrintRangeSelection Then
            liOr = dlgPrintSelection
        End If
        
        liOr = liOr Or dlgPrintUseDevModeCopiesAndCollate
    
        ' Fill in PRINTDLG structure
        Dim ltPrintDialog As TPRINTDLG
        With ltPrintDialog
            .lStructSize = Len(ltPrintDialog)                                                   'init the api structure
            .hWndOwner = tDialog.hWndOwner
            .Flags = (tDialog.iFlags And ValidFlags) Or liOr
            .nFromPage = tDialog.iFromPage
            .nToPage = tDialog.iToPage
            .nMinPage = tDialog.iMinPage
            .nMaxPage = tDialog.iMaxPage
            If tDialog.oDeviceMode Is Nothing Then Set tDialog.oDeviceMode = New cDeviceMode
            
            hDevMode = tDialog.oDeviceMode.NewHandle
            .hDevMode = hDevMode
            .hDevNames = ZeroL
            ZeroMemory .b(0), 32&
            .hdc = ZeroL
            
            #If bHookDialogs Then                                                               'if hooks are supported
                Dim liHookProc As Long
                If CBool(tDialog.iFlags And (dlgPrintEnableSetupHook Or dlgPrintEnablePrintHook)) Then
                    liHookProc = pAllocateHookProc(tDialog.oEventSink)                          'if hooks are requested, allocate machine code
                    'no callback interface or out of memory
                    Debug.Assert liHookProc
                    #If bVBVMTypeLib Then                                                       'if using the illegal type lib
                        If CBool(tDialog.iFlags And dlgPrintEnablePrintHook) Then               'set the requested hook(s)
                            MemLong(.b(tPrintDlglpfnPrintHookOffset)) = liHookProc
                        End If
                        If CBool(tDialog.iFlags And dlgPrintEnableSetupHook) Then
                            MemLong(.b(tPrintDlglpfnSetupHookOffset)) = liHookProc
                        End If
                    #Else                                                                       'else-use block moving sledge-hammer
                        If CBool(tDialog.iFlags And dlgPrintEnablePrintHook) Then
                            CopyMemory .b(tPrintDlglpfnPrintHookOffset), liHookProc, 4&
                        End If
                        If CBool(tDialog.iFlags And dlgPrintEnableSetupHook) Then
                            CopyMemory .b(tPrintDlglpfnSetupHookOffset), liHookProc, 4&
                        End If
                    #End If
                End If
                If liHookProc = 0& Then _
                    .Flags = .Flags And Not (dlgPrintEnableSetupHook Or dlgPrintEnablePrintHook) 'if no hook proc, don't enable it
            #Else
                .Flags = .Flags And Not (dlgPrintEnableSetupHook Or dlgPrintEnablePrintHook)    'if no hook proc, don't enable it
            #End If
            
        End With
    
        ' Show Print dialog
        Dim liReturn As Integer
        liReturn = PrintDlg(ltPrintDialog)
        
        #If bHookDialogs Then
            If liHookProc Then GlobalFree liHookProc                                            'if we allocated a proc, release it
        #End If
        
        If liReturn = 1 Then                                                                    'if we succeeded
            Print_ShowIndirect = True
            ' Return dialog values in parameters
            With ltPrintDialog
                tDialog.hdc = .hdc                                                              'set the return values
                tDialog.iReturnFlags = .Flags
                If CBool(.Flags And dlgPrintPageNums) Then
                    tDialog.iRange = dlgPrintRangePageNumbers
                ElseIf CBool(.Flags And dlgPrintSelection) Then
                    tDialog.iRange = dlgPrintRangeSelection
                Else
                    tDialog.iRange = dlgPrintRangeAll
                End If
                tDialog.iFromPage = .nFromPage
                tDialog.iToPage = .nToPage
                tDialog.bPrintToFile = (.Flags And dlgPrintToFile)
    
                ' Get DEVMODE structure from PRINTDLG
                tDialog.oDeviceMode.SetByHandle .hDevMode                                       'return the chosen device mode
    
                If CBool(.Flags And dlgPrintCollate) Then
                    ' User selected collate option but printer driver
                    ' does not support collation.
                    ' Collation option must be set from the
                    ' PRINTDLG structure:
                    tDialog.bCollate = True
                Else
                    ' Print driver supports collation or collation
                    ' not switched on.
                    ' DEVMODE structure contains Collation and copy
                    ' information
                    ' Get Copies and Collate settings from DEVMODE structure
                    tDialog.bCollate = (tDialog.oDeviceMode.Collate = DMCOLLATE_TRUE)
                End If
    
                Dim tDevNames As DEVNAMES
                Dim ptrDevNames As Long
    
                If .hDevNames Then
                    ptrDevNames = GlobalLock(.hDevNames)
                    If ptrDevNames Then
                        CopyMemory tDevNames, ByVal ptrDevNames, Len(tDevNames)                     'get the devmode structure
        
                        tDialog.sDriver = pPrint_GetDevNameString(ptrDevNames, tDevNames.wDriverOffset)   'extract the strings
                        tDialog.sDevice = pPrint_GetDevNameString(ptrDevNames, tDevNames.wDeviceOffset)
                        tDialog.sOutputPort = pPrint_GetDevNameString(ptrDevNames, tDevNames.wOutputOffset)
            
                        GlobalUnlock .hDevNames
                    End If
                End If
                
            End With
        Else
            If liReturn Then
                tDialog.iReturnExtendedError = CommDlgExtendedError()                                 'return the extended error
            End If
        End If
        
        With ltPrintDialog
            If .hDevNames Then GlobalFree .hDevNames                                            'free the returned devmode
            If .hDevMode <> hDevMode And .hDevMode <> 0& Then GlobalFree .hDevMode              'free the return device mode
            If hDevMode Then GlobalFree hDevMode                                                'free the allocated device mode
        End With
    
    End Function
    
    Private Function pPrint_GetDevNameString( _
            ByVal ptrDevNames As Long, _
            ByVal ptrOffset As Long) _
                As String
       Dim ptr As Long
       Dim lSize As Long
    
       ptr = UnsignedAdd(ptrDevNames, ptrOffset)
    
       lSize = lstrlenByLong(ptr)
       If (lSize > 0) Then
          pPrint_GetDevNameString = pAllocateString(lSize, False)
          CopyMemoryLpToStr pPrint_GetDevNameString, ByVal ptr, lSize
       End If
    End Function
    
#End If

#If bPageSetupDialog Then

    #If bHookDialogs Then
        Public Function PageSetup_Show( _
            Optional ByVal iFlags As ePrintPageSetup, _
            Optional ByVal iUnits As ePrintPageSetupUnits = dlgPrintInches, _
            Optional ByRef fLeftMargin As Single = 1, _
            Optional ByVal fMinLeftMargin As Single = 1, _
            Optional ByRef fRightMargin As Single = 1, _
            Optional ByVal fMinRightMargin As Single = 1, _
            Optional ByRef fTopMargin As Single = 0.5, _
            Optional ByVal fMinTopMargin As Single = 0.5, _
            Optional ByRef fBottomMargin As Single = 0.5, _
            Optional ByVal fMinBottomMargin As Single = 0.5, _
            Optional ByRef oDeviceMode As cDeviceMode, _
            Optional ByVal hWndOwner As Long, _
            Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
            Optional ByVal oEventSink As iComDlgHook) _
                As Boolean
                
    #Else
    
        Public Function PageSetup_Show( _
            Optional ByVal iFlags As ePrintPageSetup, _
            Optional ByVal iUnits As ePrintPageSetupUnits = dlgPrintInches, _
            Optional ByRef fLeftMargin As Single = 1, _
            Optional ByVal fMinLeftMargin As Single = 1, _
            Optional ByRef fRightMargin As Single = 1, _
            Optional ByVal fMinRightMargin As Single = 1, _
            Optional ByRef fTopMargin As Single = 0.5, _
            Optional ByVal fMinTopMargin As Single = 0.5, _
            Optional ByRef fBottomMargin As Single = 0.5, _
            Optional ByVal fMinBottomMargin As Single = 0.5, _
            Optional ByRef oDeviceMode As cDeviceMode, _
            Optional ByVal hWndOwner As Long, _
            Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                As Boolean
    
    #End If
        Dim ltDialog As tPageSetupDialog
        
        With ltDialog
            .iFlags = iFlags
            .hWndOwner = hWndOwner
            .iUnits = iUnits
            .fLeftMargin = fLeftMargin
            .fMinLeftMargin = fMinLeftMargin
            .fRightMargin = fRightMargin
            .fMinRightMargin = fMinRightMargin
            .fTopMargin = fTopMargin
            .fMinTopMargin = fMinTopMargin
            .fBottomMargin = fBottomMargin
            .fMinBottomMargin = fMinBottomMargin
            If oDeviceMode Is Nothing Then Set oDeviceMode = New cDeviceMode
            Set .oDeviceMode = oDeviceMode
            
            #If bHookDialogs Then
                Set .oEventSink = oEventSink
            #End If
            
            PageSetup_Show = PageSetup_ShowIndirect(ltDialog)
            
            If PageSetup_Show Then
                fLeftMargin = .fLeftMargin
                fRightMargin = .fRightMargin
                fBottomMargin = .fBottomMargin
                fTopMargin = .fTopMargin
                iReturnExtendedError = .iReturnExtendedError
            End If
            
        End With
    
    End Function

    Public Function PageSetup_ShowIndirect(ByRef tDialog As tPageSetupDialog) As Boolean
    
'        in             iFlags                  As ePrintPageSetup
'        in             hWndOwner               As Long
'        in             iUnits                  As ePrintPageSetupUnits
'
'        in/out         fLeftMargin             As Single
'        in             fMinLeftMargin          As Single
'        in/out         fRightMargin            As Single
'        in             fMinRightMargin         As Single
'        in/out         fTopMargin              As Single
'        in             fMinTopMargin           As Single
'        in/out         fBottomMargin           As Single
'        in             fMinBottomMargin        As Single
'        in/out         oDeviceMode             As cDeviceMode
'                       #If bHookDialogs Then
'        in             oEventSink              As iComDlgHook
'                       #End If
'        out            iReturnExtendedError    As eComDlgExtendedError
    
    Const ValidFlags As Long = dlgPPSDefaultMinMargins Or _
        dlgPPSMinMargins Or _
        dlgPPSMargins Or _
        dlgPPSDisableMargins Or _
        dlgPPSDisablePrinter Or _
        dlgPPSNoWarning Or _
        dlgPPSDisableOrientation Or _
        dlgPPSReturnDefault Or _
        dlgPPSDisablePaper Or _
        dlgPPSShowHelp Or _
        dlgPPSEnablePageSetupHook Or _
        dlgPPSDisablePagePainting
    
        Dim ltPageSetup As TPAGESETUPDLG
        Dim hDevMode As Long
        
        tDialog.iReturnExtendedError = ZeroL                                                        'initialize return extended error
    
        ' Fill in PRINTDLG structure
        With ltPageSetup
            .lStructSize = Len(ltPageSetup)                                                         'initialize the api structure
            .hWndOwner = tDialog.hWndOwner
    
            Dim liOr As ePrintPageSetup
            Dim iUnits As Long
    
            If tDialog.iUnits = dlgPrintInches Then                                                 'initialize the scale factor
                liOr = dlgPrintInThousandthsOfInches
                iUnits = 1000
            Else
                liOr = dlgPrintInHundredthsofMillimeters
                iUnits = 100
            End If
    
            With .rtMargin
                .Top = tDialog.fTopMargin * iUnits                                                  'set the default margins
                .Left = tDialog.fLeftMargin * iUnits
                .Bottom = tDialog.fBottomMargin * iUnits
                .Right = tDialog.fRightMargin * iUnits
                If CBool(.Top Or .Left Or .Right Or .Bottom) Then liOr = liOr Or dlgPPSMargins
            End With
    
            With .rtMinMargin
                .Top = tDialog.fMinTopMargin * iUnits                                               'set the default min margins
                .Left = tDialog.fMinLeftMargin * iUnits
                .Bottom = tDialog.fMinBottomMargin * iUnits
                .Right = tDialog.fMinRightMargin * iUnits
                If CBool(.Top Or .Left Or .Right Or .Bottom) Then liOr = liOr Or dlgPPSMinMargins
            End With
    
            .Flags = (tDialog.iFlags And ValidFlags) Or liOr                                        'set the flags
            
            .lpfnPageSetupHook = 0&
            
            #If bHookDialogs Then                                                                   'if hooking is supported
                Dim liHookProc As Long
                If CBool(.Flags And dlgPPSEnablePageSetupHook) Then                                 'if hooking is requested
                    liHookProc = pAllocateHookProc(tDialog.oEventSink)                              'allocate the machine code
                    'no callback interface or out of memory
                    Debug.Assert liHookProc
                    .lpfnPageSetupHook = liHookProc                                                 'store the procedure pointer
                End If
            #End If
            
            If .lpfnPageSetupHook = 0& Then .Flags = .Flags And Not dlgPPSEnablePageSetupHook       'if no hook procedure, don't enable it
            
            hDevMode = tDialog.oDeviceMode.NewHandle                                                'init the devmode
            .hDevMode = hDevMode
    
            ZeroMemory .ptPaperSize, Len(.ptPaperSize)                                              'make it obvious that we don't use these members
            .hInstance = ZeroL
            .lCustData = ZeroL
            .lpfnPagePaintHook = ZeroL
            .lpPageSetupTemplateName = ZeroL
            .hPageSetupTemplate = ZeroL
                
        End With
        
        ' Show Print dialog
        If PageSetupDlg(ltPageSetup) Then                                                           'if dialog succeeded
            PageSetup_ShowIndirect = True
            ' Return dialog values in parameters
            With ltPageSetup
                With .rtMargin                                                                      'return the selected margins
                    tDialog.fTopMargin = .Top / iUnits
                    tDialog.fLeftMargin = .Left / iUnits
                    tDialog.fBottomMargin = .Bottom / iUnits
                    tDialog.fRightMargin = .Right / iUnits
                End With
    
                ' Get DEVMODE structure from PRINTDLG
                tDialog.oDeviceMode.SetByHandle .hDevMode                                           'return the devmode
    
            End With
            
        Else                                                                                        'if dialog failed
            tDialog.iReturnExtendedError = CommDlgExtendedError()                                   'return the extended error
        End If
        
        If ltPageSetup.hDevMode Then GlobalFree ltPageSetup.hDevMode                                'free the allocated memory
        If ltPageSetup.hDevNames Then GlobalFree ltPageSetup.hDevNames
        If hDevMode <> ZeroL And hDevMode <> ltPageSetup.hDevMode Then GlobalFree hDevMode
        
        #If bHookDialogs Then
            If liHookProc Then GlobalFree liHookProc                                                'if we allocated a procedure, free it
        #End If
        
    End Function
    
#End If

'Public Sub testPrintDlg()
''(VB or ByteAlignment) == [Type Mismatch]
'
'Dim t As TPRINTDLG
'Dim i As Long
'i = VarPtr(t)
'
'Debug.Print VarPtr(t.lStructSize) - i, 0
'Debug.Print VarPtr(t.hWndOwner) - i, 4
'Debug.Print VarPtr(t.hDevMode) - i, 8
'Debug.Print VarPtr(t.hDevNames) - i, 12
'Debug.Print VarPtr(t.hDc) - i, 16
'Debug.Print VarPtr(t.Flags) - i, 20
'Debug.Print VarPtr(t.nFromPage) - i, 24
'Debug.Print VarPtr(t.nToPage) - i, 26
'Debug.Print VarPtr(t.nMinPage) - i, 28
'Debug.Print VarPtr(t.nMaxPage) - i, 30
'Debug.Print VarPtr(t.nCopies) - i, 32
'Debug.Print VarPtr(t.b(tPrintDlghInstanceOffset)) - i, 34
'Debug.Print VarPtr(t.b(tPrintDlglCustDataOffset)) - i, 38
'Debug.Print VarPtr(t.b(tPrintDlglpfnPrintHookOffset)) - i, 42
'Debug.Print VarPtr(t.b(tPrintDlglpfnSetupHookOffset)) - i, 46
'Debug.Print VarPtr(t.b(tPrintDlglpPrintTemplateNameOffset)) - i, 50
'Debug.Print VarPtr(t.b(tPrintDlglpSetupTemplateNameOffset)) - i, 54
'Debug.Print VarPtr(t.b(tPrintDlgPrintTemplateOffset)) - i, 58
'Debug.Print VarPtr(t.b(tPrintDlgSetupTemplateOffset)) - i, 62
'
'End Sub
'
'Public Sub testDevMode()
''(VB or ByteAlignment) == [Type Mismatch]
'
'Dim t As DEVMODE
'Dim i As Long
'i = VarPtr(t)
'
'Debug.Print VarPtr(t.dmDeviceName(0)) - i, 0
'Debug.Print VarPtr(t.dmSpecVersion) - i, 32
'Debug.Print VarPtr(t.dmDriverVersion) - i, 34
'Debug.Print VarPtr(t.dmSize) - i, 36
'Debug.Print VarPtr(t.dmDriverExtra) - i, 38
'Debug.Print VarPtr(t.dmFields) - i, 40
'Debug.Print VarPtr(t.dmOrientation) - i, 44
'Debug.Print VarPtr(t.dmPaperSize) - i, 46
'Debug.Print VarPtr(t.dmPaperLength) - i, 48
'Debug.Print VarPtr(t.dmPaperWidth) - i, 50
'Debug.Print VarPtr(t.dmScale) - i, 52
'Debug.Print VarPtr(t.dmCopies) - i, 54
'Debug.Print VarPtr(t.dmDefaultSource) - i, 56
'Debug.Print VarPtr(t.dmPrintQuality) - i, 58
'Debug.Print VarPtr(t.dmColor) - i, 60
'Debug.Print VarPtr(t.dmDuplex) - i, 62
'Debug.Print VarPtr(t.dmYResolution) - i, 64
'Debug.Print VarPtr(t.dmTTOption) - i, 66
'Debug.Print VarPtr(t.dmCollate) - i, 68
'Debug.Print VarPtr(t.dmFormName(0)) - i, 70
'Debug.Print VarPtr(t.dmUnusedPadding) - i, 102
'Debug.Print VarPtr(t.dmBitsPerPel) - i, 104
'Debug.Print VarPtr(t.dmPelsWidth) - i, 106
'Debug.Print VarPtr(t.dmPelsHeight) - i, 110
'Debug.Print VarPtr(t.dmDisplayFlags) - i, 114
'Debug.Print VarPtr(t.dmDisplayFrequency) - i, 118
'
''Private Type DEVMODE
''    dmDeviceName As String * CCHDEVICENAME
''    dmSpecVersion As Integer
''    dmDriverVersion As Integer
''    dmSize As Integer
''    dmDriverExtra As Integer
''    dmFields As Long
''    dmOrientation As Integer
''    dmPaperSize As Integer
''    dmPaperLength As Integer
''    dmPaperWidth As Integer
''    dmScale As Integer
''    dmCopies As Integer
''    dmDefaultSource As Integer
''    dmPrintQuality As Integer
''    dmColor As Integer
''    dmDuplex As Integer
''    dmYResolution As Integer
''    dmTTOption As Integer
''    dmCollate As Integer
''    dmFormName As String * CCHFORMNAME
''    dmUnusedPadding As Integer
''    dmBitsPerPel As Integer
''    dmPelsWidth As Long
''    dmPelsHeight As Long
''    dmDisplayFlags As Long
''    dmDisplayFrequency As Long
''End Type
'
'End Sub
