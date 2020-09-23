VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iComDlgHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
'iComDlgHook.cls                9/1/04
'
'           GENERAL PURPOSE:
'               Callback interface for COMDLG.dll dialog hooks.
'
'           LINEAGE:
'               N/A
'
'           DEPENDENCIES:
'               NONE
'
'==================================================================================================


Public Sub Proc( _
            ByRef lReturn As Long, _
            ByVal hDlg As Long, _
            ByVal iMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long)
    
End Sub
