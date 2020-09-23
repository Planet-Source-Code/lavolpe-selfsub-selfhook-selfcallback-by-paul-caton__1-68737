;*****************************************************************************************
;** Hook.asm - hooking thunk. Assemble with nasm.
;**
;** Paul_Caton@hotmail.com
;** Copyright free, use and abuse as you see fit.
;**
;** By LaVolpe: Created from Paul Caton's subclassing ASM....................... 20070623
;*****************************************************************************************

; *********** JUMP TO BOTTOM FOR MORE INFORMATION ABOUT THIS ASM DOCUMENT ****************

;API definitions
M_RELEASE	equ	8000h	    ;VirtualFree memory release flag

;******************************
;Stack frame access definitions
%define lParam          [ebp + 44]  ;hook lParam
%define wParam          [ebp + 40]  ;hook wParam
%define nCode           [ebp + 36]  ;hook nCode
%define lRetAddr        [ebp + 32]  ;Return address of the code that called us
%define lReturn         [ebp + 28]  ;lReturn local, restored to eax after popad
%define bHandled        [ebp + 20]  ;bHandled local, restored to edx after popad

;***********************
;Data access definitions
%define nCallStack      [ebx]       ;Hook call stack counter
%define bBypassing      [ebx +  4]  ;Shutdown flag
%define lHookType       [ebx +  8]  ;type of hook
%define fnEbMode        [ebx + 12]  ;EbMode function address
%define fnCallHookProc  [ebx + 16]  ;CallNextHook function address
%define fnUnhookWindow  [ebx + 20]  ;UnhookWindowsEx function address
%define ideObjSig   	[ebx + 24]  ;VTableRef that should change if objOwner changed due to VB END
%define fnIsBadCodePtr  [ebx + 28]  ;IsBadCodePtr function address
%define objOwner        [ebx + 32]  ;Owner object address
%define addrHookProc    [ebx + 36]  ;Original hook address
%define addrCallback    [ebx + 40]  ;Callback address
%define addrTableB      [ebx + 44]  ;non null if "before" messages desired
%define addrTableA      [ebx + 48]  ;non null if "after" message desired
%define lParamUser      [ebx + 52]  ;User defined callback parameter

use32
;************
;Data storage
    dd_nCallStack		dd 0	    ;HookProc call stack counter
    dd_bBypassing		dd 0	    ;Shutdown flag
    dd_lHookType		dd 0	    ;Hook type
    dd_fnEbMode 		dd 0	    ;EbMode function address
    dd_fnCallHookProc	dd 0	    ;CallNextHook function address
    dd_fnUnhookWindow	dd 0	    ;UnhookWindowsEx function address
    dd_ideObjSig		dd 0	    ;VTableRef that should change if objOwner changed due to VB END
    dd_fnIsBadCodePtr	dd 0	    ;IsBadCodePtr function address
    dd_objOwner 		dd 0	    ;Owner object address
    dd_addrHookProc	dd 0	    ;Original Hook address
    dd_addrCallback	dd 0	    ;Callback address
    dd_addrTableB		dd 0	    ;non null if "before" messages desired
    dd_addrTableA		dd 0	    ;non null if "after" messages desired
    dd_lParamUser		dd 0	    ;User defined callback parameter
    
;***********
;Thunk start    
    xor     eax, eax		    	;Zero eax, lReturn in the ebp stack frame
    xor     edx, edx		    	;Zero edx, bHandled in the ebp stack frame
    pushad			    		;Push all the cpu registers on to the stack
    mov     ebp, esp		    	;Setup the ebp stack frame
    mov     ebx, 012345678h	    	;dummy address of the data, patched from VB
 
    xor     esi, esi		    	;Zero esi

    inc     dword nCallStack	    	;Increment the WndProc call counter

    cmp	bBypassing, esi		;bypassing if not zero
    jnz	_bypass
						;Check for dead callback address
    push    dword addrCallback	;Push the callback address
    call    near fnIsBadCodePtr     ;Check the code is live
    jnz	_set_bypass			;if non-zero, set bBypassing flag & bypass

    cmp     fnEbMode, eax	    	;Check if the EbMode address is set
    jz     _before			;if zero, user not in IDE

						;more IDE validation
    mov	edi, objOwner 		;get ObjPtr(callback) memory address
    mov	eax, [edi + 8]		;get the value at the next 8 bytes
    cmp	ideObjSig, eax		;compare against our validation token
    jnz	_set_bypass			;if not same, don't allow callbacks

    call    near fnEbMode	    	;Determine the IDE state
    cmp     eax, dword 1	    	;If EbMode = 1
    je	_before		    	;Running normally
    
    test    eax, eax		    	;If EbMode = 0 then shut down
    jnz	_bypass			;else on breakpoint

Align 4
_set_bypass:
    mov	bBypassing, dword M_RELEASE	;set bypass flag indicating result of END

Align 4
_bypass:
    call    _hookproc		    	;EbMode = 2, breakpoint... call original WndProc
    jmp     _return		    	;Return

Align 4
_before:			    		;Before the next hook proc
    cmp	addrTableB, esi
    je	_after

    dec     edx 		    		;edx <> 0, bBefore callback parameter = True
    mov	lReturn, esi
    call    _callback	    		;Attempt the VB callback

    cmp     bHandled, esi	    	;If bHandled <> False
    jne     _return		    	;The callback entirely handled the message
    
Align 4
_after: 			    		;After the next hook

    call    _hookproc		    	;Call the next hook

    cmp	addrTableA, esi		;Is the after message flag set?
    je	_return			;if not, have thunk return

    xor     edx, edx		    	;Zero edx, bBefore callback parameter = False
    call    _callback	    		;Attempt the VB callback

Align 4
_return:			    		;Clean up and return to caller
    dec     dword nCallStack	    	;Decrement the HookProc call counter

    cmp	nCallStack, esi
    jnz     _doReturn			;hookproc call stack is recursed
    
    cmp     bBypassing, esi	    	;If bShutdown flag = 0
    jz	_doReturn		    	;Return
    
_unhook:			    		;Remove the hook
    push    dword addrHookProc	;Push the next hook proc
    call    dword fnUnhookWindow	;Call UnhookWindowsEx

Align 4
_doReturn:
    popad			    		;Pop all registers. lReturn is popped into eax
    ret     12			    	;Return with a 12 byte stack release


Align 4
_hookproc:			    		;Call the next hookproc
    push    dword lParam	    	;ByVal lParam
    push    dword wParam		;ByVal wParam
    push    dword nCode		    	;ByVal nCode
    push    dword addrHookProc	;ByVal Address of the next hook proc
    call    near fnCallHookProc	;Call CallNextHook
    mov     lReturn, eax	    	;Save the return value
    ret

Align 4
_callback:				    	;Callback required, do it...
    push    edx 		    		;Preserve edx (bBefore)
    push    dword addrCallback	;Push the callback address
    call    dword fnIsBadCodePtr    ;Check the code is live
    pop     edx 		    		;Preserve edx (bBefore)
    jnz     _generic_ret	    	;If not, skip callback
    
    lea     eax, lParamUser	    	;Address of lParamUser
    lea     ecx, bHandled	    	;Address of bHandled
    push    eax 		    		;ByRef lParamUser
    lea     eax, lReturn	    	;Address of lReturn
    push    dword lHookType	    	;ByVal hook type
    push    dword lParam	    	;ByVal lParam
    push    dword wParam	    	;ByVal wParam
    push    dword nCode		    	;ByVal nCode
    push    eax 		    		;ByRef lReturn
    push    ecx 		    		;ByRef bHandled
    push    edx 		    		;ByVal bBefore
    push    dword objOwner	    	;ByVal the owner object
    call    near addrCallback	    	;Call the zWndProc callback procedure
    
Align 4
_generic_ret:			    	;Shared return
    ret

;*****************************************************************************************

; FYI: Hook and thunk behaviour when END is executed in IDE

;  By terminating your project using END, there are no messages sent to the hook that lets
;  the thunk know that IDE just terminated. The mouse hook will probably fire next time you 
;  move the mouse, but other hooks won't because they are waiting on their next message. 
;  And END prevented normal cleanup code which would have unhooked the hook and released 
;  thunk memory.

;  Well, when you re-run your project, next time hook gets a message it will check to see if
;  IDE is dead (it isn't) and whether or not the callback address pointer is still a valid
;  pointer. It most likely is, but the pointer is now for a different object because the
;  original object was destroyed when END occurred. These are the safety checks, but when
;  the message is forwarded to the callback -- CRASH because the thunk's object doesn't
;  exist any longer, it was destroyed with END.

;  To get around this problem, the thunk now has a value which is a VTable pointer that
;  changes whenever a new object is created (result of END). This value is checked when
;  in IDE only and can determine if callback address is valid or not. If not, the thunk
;  does not terminate but it will not do the callbacks either. Once END is executed on a
;  project that has an active hook, the only way to release the memory used by the thunk
;  is to simply close VB.

;*****************************************************************************************

;To compile this asm file into a bin file, you can create batch file with....
;Suggest naming batch file:  NASMtoBIN.bat
;Replace path below with path to where your nasmw.exe file exists

;	ECHO ON
;	C:\nasm\nasmw -f bin %1.asm -o %1.bin -l %1_List.txt -E %1_Err.txt
;	echo off
;	echo :
;	echo  CLOSE THIS WINDOW NOW
;	echo  THEN CHECK Err.txt FILE To Validate Compilation

; Copy the batch file into same folder as the asm file
; Then Open a DOS window, command prompt. Change directories to where the asm file is
; Finally type at the command prompt:  NASMtoBIN hook

; Download NASM compiler from following website
; http://sourceforge.net/project/showfiles.php?group_id=6208

;*****************************************************************************************
