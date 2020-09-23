VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#Const bColorDialog = True
#Const bFileDialog = True
#Const bFolderDialog = True
#Const bFontDialog = True
#Const bPrintDialog = True
#Const bPageSetupDialog = True

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

#If bFolderDialog Then

    Public Enum eFolderDialog
        dlgFolderReturnOnlyFSDirs = &H1      'Only returns file system directories
        dlgFolderDontGoBelowDomain = &H2     'Does not include network folders below domain level
        dlgFolderStatusText = &H4            'Includes status area in the dialog for use with callback
        dlgFolderReturnFSAncestors = &H8     'Only returns file system ancestors.
        dlgFolderEditBox = &H10              'allows user to rename selection
        dlgFolderValidate = &H20             'insist on valid editbox result (or CANCEL)
        dlgFolderUseNewUI = &H40             'Version 5.0. Use the new user-interface. Setting
                                                      'this flag provides the user with a larger dialog box
                                                      'that can be resized. It has several new capabilities
                                                      'including: drag and drop capability within the
                                                      'dialog box, reordering, context menus, new folders,
                                                      'delete, and other context menu commands. To use
                                                      'this flag, you must call OleInitialize or
                                                      'CoInitialize before calling SHBrowseForFolder.
        dlgFolderBrowseForComputer = &H1000  'Only returns computers.
        dlgFolderBrowseForPrinter = &H2000   'Only returns printers.
        dlgFolderBrowseIncludeFiles = &H4000 'Browse for everything
        dlgFolderRaiseError = &H10000000
        dlgFolderEnableHook = &H20000000
    End Enum
    
    Public Enum eSpecialFolders
        dlgFolderDesktop = &H0                  '(desktop)
        dlgFolderInternet = &H1                 'Internet Explorer (icon on desktop)
        dlgFolderPrograms = &H2                 'Start Menu\Programs
        dlgFolderControls = &H3                 'My Computer\Control Panel
        dlgFolderPrinters = &H4                 'My Computer\Printers
        dlgFolderPersonal = &H5                 'My Documents
        dlgFolderFavorites = &H6                '(user name)\Favorites
        dlgFolderStartup = &H7                  'Start Menu\Programs\Startup
        dlgFolderRecent = &H8                   '(user name)\Recent
        dlgFolderSendTo = &H9                   '(user name)\SendTo
        dlgFolderBitBucket = &HA                '(desktop)\Recycle Bin
        dlgFolderStartMenu = &HB                '(user name)\Start Menu
        dlgFolderDesktopDirectory = &H10        '(user name)\Desktop
        dlgFolderDrives = &H11                  'My Computer
        dlgFolderNetwork = &H12                 'Network Neighborhood
        dlgFolderNethood = &H13                 '(user name)\nethood
        dlgFolderFonts = &H14                   'windows\fonts
        dlgFolderTemplates = &H15
        dlgFolderCommonStartMenu = &H16         'All Users\Start Menu
        dlgFolderCommonPrograms = &H17          'All Users\Programs
        dlgFolderCommonStartup = &H18           'All Users\Startup
        dlgFolderCommonDesktopDirectory = &H19  'All Users\Desktop
        dlgFolderAppData = &H1A                 '(user name)\Application Data
        dlgFolderPrinthood = &H1B               '(user name)\PrintHood
        dlgFolderLocalAppData = &H1C            '(user name)\Local Settings\Applicaiton Data (non roaming)
        dlgFolderAltStartup = &H1D              'non localized startup
        dlgFolderCommonAltStartup = &H1E        'non localized common startup
        dlgFolderCommonFavorites = &H1F
        dlgFolderInternetCache = &H20
        dlgFolderCookies = &H21
        dlgFolderHistory = &H22
        dlgFolderCommonAppData = &H23           'All Users\Application Data
        dlgFolderWindows = &H24                 'GetWindowsDirectory()
        dlgFolderSystem = &H25                  'GetSystemDirectory()
        dlgFolderProgramFiles = &H26            'C:\Program Files
        dlgFolderMyPictures = &H27              'C:\Program Files\My Pictures
        dlgFolderProfile = &H28                 'USERPROFILE
        dlgFolderProgramFilesCommon = &H2B      'C:\Program Files\Common
        dlgFolderCommonTemplates = &H2D         'All Users\Templates
        dlgFolderCommonDocuments = &H2E         'All Users\Documents
        dlgFolderCommonAdminTools = &H2F        'All Users\Start Menu\Programs\Administrative Tools
        dlgFolderAdminTools = &H30              '(user name)\Start Menu\Programs\Administrative Tools
        
        dlgFolderFlagCreate = &H8000&            'combine with dlgFolder value to force create on SHGetSpecialFolderLocation()
        dlgFolderFlagDontVerify = &H4000        'combine with dlgFolder value to force create on SHGetSpecialFolderLocation()
        dlgFolderFlagMask = &HFF00              'mask for all possible flag values
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

#If bPrintDialog Then
    
    Public Enum ePrintDialog
        dlgPrintAllPages = &H0
        dlgPrintSelection = &H1
        dlgPrintPageNums = &H2
        dlgPrintNoSelection = &H4
        dlgPrintNoPageNums = &H8
        dlgPrintCollate = &H10
        dlgPrintToFile = &H20
        dlgPrintSetup = &H40
        dlgPrintNoWarning = &H80
        dlgPrintReturnDc = &H100
        dlgPrintReturnIc = &H200
        dlgPrintReturnDefault = &H400
        dlgPrintShowHelp = &H800
        dlgPrintEnablePrintHook = &H1000
        dlgPrintEnableSetupHook = &H2000
        'dlgPrintEnablePrintTemplate = &H4000
        'dlgPrintEnableSetupTemplate = &H8000&
        'dlgPrintEnablePrintTemplateHandle = &H10000
        'dlgPrintEnableSetupTemplateHandle = &H20000
        'dlgPrintUseDevModeCopies = &H40000
        'dlgPrintUseDevModeCopiesAndCollate = &H40000
        dlgPrintDisablePrintToFile = &H80000
        dlgPrintHidePrintToFile = &H100000
        dlgPrintNoNetworkButton = &H200000
        'dlgPrintEnableSetupTemplate_C = &H8000&
        dlgPrintRaiseError = &H10000000
    End Enum
    
    Public Enum ePrintRange
        dlgPrintRangeAll
        dlgPrintRangePageNumbers
        dlgPrintRangeSelection
    End Enum
    
#End If

#If bPageSetupDialog Then

    Public Enum ePrintPageSetup
        dlgPPSDefaultMinMargins = &H0 ' Default (printer's)
        dlgPPSMinMargins = &H1
        dlgPPSMargins = &H2
        dlgPPSDisableMargins = &H10
        dlgPPSDisablePrinter = &H20
        dlgPPSNoWarning = &H80
        dlgPPSDisableOrientation = &H100
        dlgPPSReturnDefault = &H400
        dlgPPSDisablePaper = &H200
        dlgPPSShowHelp = &H800
        dlgPPSEnablePageSetupHook = &H2000
        'dlgPrintEnablePageSetupTemplate = &H8000&
        'dlgPrintEnablePageSetupTemplateHandle = &H20000
        'dlgPrintEnablePagePaintHook = &H40000
        dlgPPSDisablePagePainting = &H80000
        dlgPPSRaiseError = &H10000000
    End Enum
    
    Public Enum ePrintPageSetupUnits
        dlgPrintInches
        dlgPrintMillimeters
    End Enum

#End If

#If bPageSetupDialog Or bPrintDialog Then
    Private Const DM_ORIENTATION As Long = &H1&
    Private Const DM_PAPERSIZE As Long = &H2&
    Private Const DM_PAPERLENGTH As Long = &H4&
    Private Const DM_PAPERWIDTH As Long = &H8&
    Private Const DM_SCALE As Long = &H10&
    Private Const DM_COPIES As Long = &H100&
    Private Const DM_DEFAULTSOURCE As Long = &H200&
    Private Const DM_PRINTQUALITY As Long = &H400&
    Private Const DM_POSITION As Long = &H20&
    Private Const DM_COLOR As Long = &H800&
    Private Const DM_DUPLEX As Long = &H1000&
    Private Const DM_YRESOLUTION As Long = &H2000&
    Private Const DM_TTOPTION As Long = &H4000&
    Private Const DM_COLLATE As Long = &H8000
    Private Const DM_FORMNAME As Long = &H10000
    Private Const DM_LOGPIXELS As Long = &H20000
    Private Const DM_BITSPERPEL As Long = &H40000
    Private Const DM_PELSWIDTH As Long = &H80000
    Private Const DM_PELSHEIGHT As Long = &H100000
    Private Const DM_NUP As Long = &H40&
    
    Public Enum eDeviceModeFields
        dmfOrientation = DM_ORIENTATION
        dmfPaperSize = DM_PAPERSIZE
        dmfPaperLength = DM_PAPERLENGTH
        dmfPaperWidth = DM_PAPERWIDTH
        dmfScale = DM_SCALE
        dmfCopies = DM_COPIES
        dmfDefaultSource = DM_DEFAULTSOURCE
        dmfPrintQuality = DM_PRINTQUALITY
        dmfPosition = DM_POSITION
        dmfColor = DM_COLOR
        dmfDuplex = DM_DUPLEX
        dmfYResolution = DM_YRESOLUTION
        dmfTTOption = DM_TTOPTION
        dmfCollate = DM_COLLATE
        dmfFormName = DM_FORMNAME
        dmfLogPixels = DM_LOGPIXELS
        dmfBitsPerPel = DM_BITSPERPEL
        dmfPelsWidth = DM_PELSWIDTH
        dmfPelsHeight = DM_PELSHEIGHT
        dmfNUP = DM_NUP
    End Enum

    Public Enum ePrintOrientation
        dlgPrintPortrait = 1
        dlgPrintLandscape
    End Enum

    Public Enum ePrintPaperSize
        dlgPrintPaperSize10x11 = 45
        dlgPrintPaperSize10x14 = 16
        dlgPrintPaperSize11x17 = 17
        dlgPrintPaperSize12x11 = 90
        dlgPrintPaperSize15x11 = 46
        dlgPrintPaperSize9x11 = 44
        dlgPrintPaperSizeA_Plus = 57
        dlgPrintPaperSizeA2 = 66
        dlgPrintPaperSizeA3 = 8                                 ' A3, 297 x 420 mm
        dlgPrintPaperSizeA3Extra = 63
        dlgPrintPaperSizeA3ExtraTransverse = 68
        dlgPrintPaperSizeA3Rotated = 76
        dlgPrintPaperSizeA3Transverse = 67
        dlgPrintPaperSizeA4 = 9                                 ' A4, 210 x 297 mm
        dlgPrintPaperSizeA4Extra = 53
        dlgPrintPaperSizeA4Plus = 60
        dlgPrintPaperSizeA4Rotated = 77
        dlgPrintPaperSizeA4Transverse = 55
        dlgPrintPaperSizeA4Small = 10                           ' A4 Small, 210 x 297 mm
        dlgPrintPaperSizeA5 = 11                                ' A5, 148 x 210 mm
        dlgPrintPaperSizeA5Extra = 64
        dlgPrintPaperSizeA5Rotated = 78
        dlgPrintPaperSizeA5Transverse = 61
        dlgPrintPaperSizeA6 = 70
        dlgPrintPaperSizeA6Rotated = 83
        dlgPrintPaperSizeBPlus = 58
        dlgPrintPaperSizeB4 = 12                                ' B4, 250 x 354 mm
        dlgPrintPaperSizeB4JisRotated = 79
        dlgPrintPaperSizeB5 = 13                                ' B5, 182 x 257 mm
        dlgPrintPaperSizeB5Extra = 65
        dlgPrintPaperSizeB5JisRotated = 80
        dlgPrintPaperSizeB5Transverse = 62
        dlgPrintPaperSizeB6Jis = 88
        dlgPrintPaperSizeB6JisRotated = 89
        dlgPrintPaperSizeCSheet = 24
        dlgPrintPaperSizeDblJapanesePostcard = 69
        dlgPrintPaperSizeDblJapanesePostcartRotated = 82
        dlgPrintPaperSizeDSheet = 25
        dlgPrintPaperSizeEnv10 = 20                             ' Envelope #10, 4 1/8 x 9 1/2 in.
        dlgPrintPaperSizeEnv11 = 21                             ' Envelope #11, 4 1/2 x 10 3/8 in.
        dlgPrintPaperSizeEnv12 = 22                             ' Envelope #12, 4 1/2 x 11 in.
        dlgPrintPaperSizeEnv14 = 23                             ' Envelope #14, 5 x 11 1/2 in.
        dlgPrintPaperSizeEnv9 = 19                              ' Envelope #9, 3 7/8 x 8 7/8 in.
        dlgPrintPaperSizeEnvB4 = 33                             ' Envelope B4, 250 x 353 mm
        dlgPrintPaperSizeEnvB5 = 34                             ' Envelope B5, 176 x 250 mm
        dlgPrintPaperSizeEnvB6 = 35                             ' Envelope B6, 176 x 125 mm
        dlgPrintPaperSizeEnvC3 = 29                             ' Envelope C3, 324 x 458 mm
        dlgPrintPaperSizeEnvC4 = 30                             ' Envelope C4, 229 x 324 mm
        dlgPrintPaperSizeEnvC5 = 28                             ' Envelope C5, 162 x 229 mm
        dlgPrintPaperSizeEnvC6 = 31                             ' Envelope C6, 114 x 162 mm
        dlgPrintPaperSizeEnvC65 = 32                            ' Envelope C65, 114 x 229 mm
        dlgPrintPaperSizeEnvDL = 27                             ' Envelope DL, 110 x 220 mm
        dlgPrintPaperSizeEnvInvite = 47
        dlgPrintPaperSizeEnvItaly = 36                          ' Envelope, 110 x 230 mm
        dlgPrintPaperSizeEnvMonarch = 37                        ' Envelope Monarch, 3 7/8 x 7 1/2 in.
        dlgPrintPaperSizeEnvPersonal = 38                       ' Envelope, 3 5/8 x 6 1/2 in.
        dlgPrintPaperSizeESheet = 26
        dlgPrintPaperSizeExecutive = 7                          ' Executive, 7 1/2 x 10 1/2 in.
        dlgPrintPaperSizeFanfoldLglGerman = 41                  ' German Legal Fanfold, 8 1/2 x 13 in.
        dlgPrintPaperSizeFanfoldStdGerman = 40                  ' German Standard Fanfold, 8 1/2 x 12 in.
        dlgPrintPaperSizeFanfoldUS = 39                         ' U.S. Standard Fanfold, 14 7/8 x 11 in.
        dlgPrintPaperSizeLetter = 1                             ' Letter, 8 1/2 x 11 in.
        dlgPrintPaperSizeFolio = 14                             ' Folio, 8 1/2 x 13 in.
        dlgPrintPaperSizeIso_B4 = 42
        dlgPrintPaperSizeJapanesePostcard = 43
        dlgPrintPaperSizeJapanesePostcardRotated = 81
        dlgPrintPaperSizeJEnvChou3 = 73
        dlgPrintPaperSizeJEnvChou3Rotated = 86
        dlgPrintPaperSizeJEnvChou4 = 74
        dlgPrintPaperSizeJEnvChou4_Rotated = 87
        dlgPrintPaperSizeJEnvKaku2 = 71
        dlgPrintPaperSizeJEnvKaku2_Rotated = 84
        dlgPrintPaperSizeJEnvKaku3 = 72
        dlgPrintPaperSizeJEnvKaku3_Rotated = 85
        dlgPrintPaperSizeJEnvYou4 = 91
        dlgPrintPaperSizeJEnvYou4_Rotated = 92
        dlgPrintPaperSizeLedger = 4                             ' Ledger, 17 x 11 in.
        dlgPrintPaperSizeLegal = 5                              ' Legal, 8 1/2 x 14 in.
        dlgPrintPaperSizeLegalExtra = 51
        dlgPrintPaperSizeLetterExtra = 50
        dlgPrintPaperSizeLetterExtraTransverse = 56
        dlgPrintPaperSizeLetterPlus = 59
        dlgPrintPaperSizeLetterRotated = 75
        dlgPrintPaperSizeLetterTransverse = 54
        dlgPrintPaperSizeLetterSmall = 2                        ' Letter Small, 8 1/2 x 11 in.
        dlgPrintPaperSizeNote = 18                              ' Note, 8 1/2 x 11 in.
        dlgPrintPaperSizeP16K = 93
        dlgPrintPaperSizeP16KRotated = 106
        dlgPrintPaperSizeP32K = 94
        dlgPrintPaperSizeP32KRotated = 107
        dlgPrintPaperSizeP32KBig = 95
        dlgPrintPaperSizeP32KBigRotated = 108
        dlgPrintPaperSizePEnv1 = 96
        dlgPrintPaperSizePEnv1Rotated = 109
        dlgPrintPaperSizePEnv10 = 105
        dlgPrintPaperSizePEnv10Rotated = 118
        dlgPrintPaperSizePEnv2 = 97
        dlgPrintPaperSizePEnv2Rotated = 110
        dlgPrintPaperSizePEnv3 = 98
        dlgPrintPaperSizePEnv3Rotated = 111
        dlgPrintPaperSizePEnv4 = 99
        dlgPrintPaperSizePEnv4Rotated = 112
        dlgPrintPaperSizePEnv5 = 100
        dlgPrintPaperSizePEnv5Rotated = 113
        dlgPrintPaperSizePEnv6 = 101
        dlgPrintPaperSizePEnv6Rotated = 114
        dlgPrintPaperSizePEnv7 = 102
        dlgPrintPaperSizePEnv7Rotated = 115
        dlgPrintPaperSizePEnv8 = 103
        dlgPrintPaperSizePEnv8Rotated = 116
        dlgPrintPaperSizePEnv9 = 104
        dlgPrintPaperSizePEnv9Rotated = 117
        dlgPrintPaperSizeQuarto = 15                            ' Quarto, 215 x 275 mm
        dlgPrintPaperSizeReserved48 = 48
        dlgPrintPaperSizeReserved49 = 49
        dlgPrintPaperSizeStatement = 6                          ' Statement, 5 1/2 x 8 1/2 in.
        dlgPrintPaperSizeTabloid = 3                            ' Tabloid, 11 x 17 in.
        dlgPrintPaperSizeTabloid_Extra = 52
        dlgPrintPaperSizeUserDefined = 256
    End Enum
    
    ' EPrintQuality constants same as vbPRPQ constants
    Public Enum ePrintQuality
        dlgPrintDraft = -1
        dlgPrintLow = -2
        dlgPrintMedium = -3
        dlgPrintHigh = -4
    End Enum

#End If
