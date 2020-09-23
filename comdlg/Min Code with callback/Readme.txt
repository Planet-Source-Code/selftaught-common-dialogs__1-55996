In this example, Dialog hooks are enabled by setting the bHookDialogs compiler constant to TRUE in mCommonDialog.bas.  This requires the iComDlgHook class to be included in the project, and requires the object on which callbacks are to be made to implement the iComDlgHook interface.

The Print and Page Setup Dialogs require the cDeviceMode.cls file, and are not included in the minimum code examples.

The Folder dialog requires the vbBase.tlb file, and also is not included in the minimum code examples.