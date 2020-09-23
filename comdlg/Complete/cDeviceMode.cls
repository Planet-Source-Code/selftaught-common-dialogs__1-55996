VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeviceMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cDeviceMode.cls                      8/25/04
'
'           GENERAL PURPOSE:
'               Wraps a DEVMODE structure
'
'           LINEAGE:
'               N/A
'
'           DEPENDENCIES:
'               NONE
'
'==================================================================================================


Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Type DEVMODE
    dmDeviceName    As String * CCHDEVICENAME
    dmSpecVersion   As Integer
    dmDriverVersion As Integer
    dmSize          As Integer
    dmDriverExtra   As Integer
    dmFields        As Long
    dmOrientation   As Integer
    dmPaperSize     As Integer
    dmPaperLength   As Integer
    dmPaperWidth    As Integer
    dmScale         As Integer
    dmCopies        As Integer
    dmDefaultSource As Integer
    dmPrintQuality  As Integer
    dmColor         As Integer
    dmDuplex        As Integer
    dmYResolution   As Integer
    dmTTOption      As Integer
    dmCollate       As Integer
    dmFormName      As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel    As Integer
    'off two bytes starting here
    b(0 To 15)      As Byte
    'dmPelsWidth     As Long
    'dmPelsHeight    As Long
    'dmDisplayFlags  As Long
    'dmDisplayFrequency As Long
End Type

Private Const GMEM_MOVEABLE = &H2

Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private mtDevMode As DEVMODE

Public Property Get DeviceName() As String
    Dim i As Long
    i = InStr(1, mtDevMode.dmDeviceName, vbNullChar)
    If i Then DeviceName = Left$(mtDevMode.dmDeviceName, i - 1) Else DeviceName = mtDevMode.dmDeviceName
End Property
Public Property Let DeviceName(ByRef sNew As String)
    If Len(sNew) < CCHDEVICENAME Then
        mtDevMode.dmDeviceName = sNew & String$(CCHDEVICENAME - Len(sNew), 0)
    Else
        mtDevMode.dmDeviceName = sNew
    End If
End Property

Public Property Get SpecVersion() As Integer
    SpecVersion = mtDevMode.dmSpecVersion
End Property
Public Property Let SpecVersion(ByVal iNew As Integer)
    mtDevMode.dmSpecVersion = iNew
End Property

Public Property Get DriverVersion() As Integer
    DriverVersion = mtDevMode.dmDriverVersion
End Property
Public Property Let DriverVersion(ByVal iNew As Integer)
    mtDevMode.dmDriverVersion = iNew
End Property

Public Property Get DriverExtra() As Integer
    DriverExtra = mtDevMode.dmDriverExtra
End Property
Public Property Let DriverExtra(ByVal iNew As Integer)
    mtDevMode.dmDriverExtra = iNew
End Property

Public Property Get Fields() As eDeviceModeFields
    Fields = mtDevMode.dmFields
End Property
Public Property Let Fields(ByVal iNew As eDeviceModeFields)
    mtDevMode.dmFields = iNew
End Property

Public Property Get Orientation() As ePrintOrientation
    Orientation = mtDevMode.dmOrientation
End Property
Public Property Let Orientation(ByVal iNew As ePrintOrientation)
    mtDevMode.dmOrientation = iNew And &HFFFF
End Property

Public Property Get PaperSize() As ePrintPaperSize
    PaperSize = mtDevMode.dmPaperSize
End Property
Public Property Let PaperSize(ByVal iNew As ePrintPaperSize)
    mtDevMode.dmPaperSize = iNew And &HFFFF
End Property

Public Property Get PaperLength() As Integer
    PaperLength = mtDevMode.dmPaperLength
End Property
Public Property Let PaperLength(ByVal iNew As Integer)
    mtDevMode.dmPaperLength = iNew
End Property

Public Property Get PaperWidth() As Integer
    PaperWidth = mtDevMode.dmPaperWidth
End Property
Public Property Let PaperWidth(ByVal iNew As Integer)
    mtDevMode.dmPaperWidth = iNew
End Property

Public Property Get OutputScaleFactor() As Integer
    OutputScaleFactor = mtDevMode.dmScale
End Property
Public Property Let OutputScaleFactor(ByVal iNew As Integer)
    mtDevMode.dmScale = iNew
End Property

Public Property Get Copies() As Integer
    Copies = mtDevMode.dmCopies
End Property
Public Property Let Copies(ByVal iNew As Integer)
    mtDevMode.dmCopies = iNew
End Property

Public Property Get DefaultSource() As Integer
    DefaultSource = mtDevMode.dmDefaultSource
End Property
Public Property Let DefaultSource(ByVal iNew As Integer)
    mtDevMode.dmDefaultSource = iNew
End Property

Public Property Get PrintQuality() As ePrintQuality
    PrintQuality = mtDevMode.dmPrintQuality
End Property
Public Property Let PrintQuality(ByVal iNew As ePrintQuality)
    mtDevMode.dmPrintQuality = iNew And &HFFFF
End Property

Public Property Get Color() As Integer
    Color = mtDevMode.dmColor
End Property
Public Property Let Color(ByVal iNew As Integer)
    mtDevMode.dmColor = iNew
End Property

Public Property Get Duplex() As Integer
    Duplex = mtDevMode.dmDuplex
End Property
Public Property Let Duplex(ByVal iNew As Integer)
    mtDevMode.dmDuplex = iNew
End Property

Public Property Get YResolution() As Integer
    YResolution = mtDevMode.dmYResolution
End Property
Public Property Let YResolution(ByVal iNew As Integer)
    mtDevMode.dmYResolution = iNew
End Property

Public Property Get TTOption() As Integer
    TTOption = mtDevMode.dmTTOption
End Property
Public Property Let TTOption(ByVal iNew As Integer)
    mtDevMode.dmTTOption = iNew
End Property

Public Property Get Collate() As Integer
    Collate = mtDevMode.dmCollate
End Property
Public Property Let Collate(ByVal iNew As Integer)
    mtDevMode.dmCollate = iNew
End Property

Public Property Get FormName() As String
    Dim i As Long
    i = InStr(1, mtDevMode.dmFormName, vbNullChar)
    If i Then FormName = Left$(mtDevMode.dmFormName, i - 1) Else FormName = mtDevMode.dmFormName
End Property
Public Property Let FormName(ByRef sNew As String)
    If Len(sNew) < CCHFORMNAME Then
        mtDevMode.dmFormName = sNew & String$(CCHFORMNAME - Len(sNew), 0)
    Else
        mtDevMode.dmFormName = sNew
    End If
End Property

Public Property Get BitsPerPel() As Integer
    BitsPerPel = mtDevMode.dmBitsPerPel
End Property
Public Property Let BitsPerPel(ByVal iNew As Integer)
    mtDevMode.dmBitsPerPel = iNew
End Property

Private Sub Class_Initialize()
    mtDevMode.dmSize = Len(mtDevMode)
End Sub

Friend Property Get NewHandle() As Long
    NewHandle = GlobalAlloc(GMEM_MOVEABLE, Len(mtDevMode))
    If NewHandle Then
        CopyMemory ByVal GlobalLock(NewHandle), mtDevMode, Len(mtDevMode)
        GlobalUnlock NewHandle
    End If
End Property

Friend Sub SetByHandle(ByVal hDevMode As Long)
    If hDevMode Then
        CopyMemory mtDevMode, ByVal GlobalLock(hDevMode), Len(mtDevMode)
        GlobalUnlock hDevMode
    End If
End Sub
