;Common dialog hook procedure
;uses iComDlgHook.cls for the owner object

%define _patch1_    01BCCAABh               ;address of owner object

[bits 32]

    call    L1                              ;Push this address onto the stack
L1: pop     edx                             ;Pop this address into edx
    pop     dword [edx+(_ret-L1)]           ;Pop the calling function's return address off the stack to the save location
    xor     eax,  eax                       ;Clear eax
    mov     dword [edx+(_lRet-L1)], eax     ;Clear lReturn
    add     edx,  (_lRet-L1)                ;get the address of _lRet
    push    edx                             ;Push ByRef _lRet
    
    mov     eax,  _patch1_                  ;Address of the owner object, patched at runtime
    push    eax                             ;Push address of the owner object
    mov     eax,  [eax]                     ;Get the address of the vTable
    call    dword [eax+1Ch]                 ;Make the call, vTable offset 1Ch
    
    call    L2                              ;Call the next instruction
L2: pop     edx                             ;Pop the return address (this address!) into edx
    push    dword [edx+(_ret-L2)]           ;Push the saved return address
    mov     eax,  [edx+(_lRet-L2)]          ;Return the function value in eax
    ret                                     ;Return to caller
    _ret    dd 0h                           ;Return address of the caller saved here
    _lRet   dd 0h                           ;function return value saved here