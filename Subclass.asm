;*****************************************************************************************
;** Subclass.asm - subclassing thunk. Assemble with nasm.
;** Copyright free, use and abuse as you see fit.
;**
;** v2.0 Re-write by LaVolpe, based mostly on Paul Caton's original thunks....... 20070720
;** .... Reorganized & provided following additional logic
;** ....... Unsubclassing only occurs after thunk is no longer recursed
;** ....... Flag used to bypass callbacks until unsubclassing can occur
;** ....... Timer used as delay mechanism to free thunk memory afer unsubclassing occurs
;** .............. Prevents crash when one window subclassed multiple times
;** .............. More END safe, even if END occurs within the subclass procedure
;** ....... Added ability to destroy API windows when IDE terminates
;** ....... Added auto-unsubclass when WM_NCDESTROY processed 
;**
;** v1.x by Paul Caton can be found at following link:
;** 	www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=64867&lngWId=1
;*****************************************************************************************

; *********** JUMP TO BOTTOM FOR MORE INFORMATION ABOUT THIS ASM DOCUMENT ****************

;API definitions
GWL_WNDPROC		equ	dword -4	    ;SetWindowsLong WndProc parameter
M_RELEASE		equ	dword 8000h	    ;VirtualFree memory release flag
WM_NCDESTROY 	equ   dword 82h  	    ;the last message a window gets

;******************************
;Stack frame access definitions
%define lParam          [ebp + 48]  ;WndProc lParam or TickCount if own timer callback
%define wParam          [ebp + 44]  ;WndProc wParam or TimerID if own timer callback
%define uMsg            [ebp + 40]  ;WndProc uMsg
%define hWnd            [ebp + 36]  ;WndProc hWnd
%define lRetAddr        [ebp + 32]  ;Return address of the code that called us
%define lReturn         [ebp + 28]  ;lReturn local, restored to eax after popad
%define bHandled        [ebp + 20]  ;bHandled local, restored to edx after popad

;***********************
;Data access definitions
%define nCallStack      [ebx]       ;WndProc call stack counter
%define bBypassing      [ebx +  4]  ;flag preventing callbacks to user
%define lWnd            [ebx +  8]  ;Window handle
%define fnEbMode        [ebx + 12]  ;EbMode function address
%define fnCallWinProc   [ebx + 16]  ;CallWindowProc function address
%define fnSetWinLong    [ebx + 20]  ;SetWindowsLong function address
%define fnVirtualFree   [edx + 24]  ;VirtualFree function address (deliberately edx 
%define fnIsBadCodePtr  [ebx + 28]  ;IsBadCodePtr function address
%define objOwner        [ebx + 32]  ;Owner object address
%define addrWndProc     [ebx + 36]  ;Original WndProc address 
%define addrCallback    [ebx + 40]  ;Callback address
%define addrTableB      [ebx + 44]  ;Address of before original WndProc message table
%define addrTableA      [ebx + 48]  ;Address of after original WndProc message table
%define lParamUser      [ebx + 52]  ;User defined callback parameter or TimerID when own callback
%define fnDestroyWindow [ebx + 56]	;DestroyWindow function address
%define fnSetTimer	[ebx + 60]	;SetTimer function address
%define fnKillTimer	[ebx + 64]	;KillTimer function address
%define timerID		[ebx + 68]	;local use when SetTimer called

; WARNING:  If any parameters/definitions are added/deleted, ensure
;		the 'add' instruction offset in _SetTimer is correct.
;		The value must be the number of storage items * 4. 
;		With 18 items, 18*4=72 or 48h

use32
;************
;Data storage
    dd_nCallStack		dd 0	    ;WndProc call stack counter
    dd_bBypassing		dd 0	    ;Shutdown flag
    dd_lWnd			dd 0	    ;Window handle
    dd_fnEbMode 		dd 0	    ;EbMode function address
    dd_fnCallWinProc	dd 0	    ;CallWindowProc function address
    dd_fnSetWinLong	dd 0	    ;SetWindowsLong function address
    dd_fnVirtualFree	dd 0	    ;VirtualFree function address
    dd_fnIsBadCodePtr	dd 0	    ;ISBadCodePtr function address
    dd_objOwner 		dd 0	    ;Owner object address
    dd_addrWndProc	dd 0	    ;Original WndProc address
    dd_addrCallback	dd 0	    ;Callback address
    dd_addrTableB		dd 0	    ;Address of before original WndProc message table
    dd_addrTableA		dd 0	    ;Address of after original WndProc message table
    dd_lParamUser		dd 0	    ;User defined callback parameter
    dd_fnDestroyWindow  dd 0	    ;DestroyWindow function address
    dd_SetTimer		dd 0	    ;SetTimer function address
    dd_KillTimer		dd 0      ;KillTimer function address
    dd_timerID		dd 0	    ;TimerID returned by SetTimer
    
;***********
;Thunk start    
Align 4
    xor     eax, eax		    	;Zero eax, lReturn in the ebp stack frame
    xor     edx, edx		    	;Zero edx, bHandled in the ebp stack frame
    pushad			    		;Push all the cpu registers on to the stack
    mov     ebp, esp		    	;Setup the ebp stack frame
    mov     ebx, 012345678h	    	;dummy Address of the data, patched from VB

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
    call    _wndproc		    	;EbMode = 2, breakpoint... call original WndProc
    jmp     _return		    	;Return

Align	4
_before:			    		;Before the original WndProc
    dec     edx 		    		;edx <> 0, bBefore callback parameter = True
    mov     edi, addrTableB	    	;Get the before message table
    mov	lReturn, esi
    call    _callback		    	;Attempt the VB callback
    
Align 4
_handle_before:
    cmp     bHandled, esi	    	;If bHandled <> False
    jne     _return		    	;The callback entirely handled the message
    call    _wndproc		    	;Call the original WndProc
    
_after: 			    		;After the original WndProc
    xor     edx, edx		    	;Zero edx, bBefore callback parameter = False
    mov     edi, addrTableA	    	;Get the after message table
    call    _callback		    	;Attempt the VB callback

Align	4
_return:			    		;Clean up and return to caller
    dec     dword nCallStack	    	;Decrement the WndProc call counter

    cmp	uMsg, dword WM_NCDESTROY ;is window being destroyed?
    jne	_check_recursion		; if so, set the bypass flag
    mov	bBypassing, dword 1

Align 4
_check_recursion:
    cmp	nCallStack, esi		;done recursing?
    jne	_do_return			;if not then return
    
    cmp	bBypassing, esi		;is bypass flag set?
    je	_do_return			;if not then return
						;otherwise not recursing & bypass set
						;unsubclass if not already done
_unsubclass:
    push    dword addrWndProc	    	;Address of the original WndProc
    push    dword GWL_WNDPROC	    	;WndProc index
    push    dword lWnd		    	;Push the window handle
    call    near fnSetWinLong		;Call SetWindowsLong

_do_we_destroy:
    cmp	fnDestroyWindow, esi	;option to destroywindow when objOwner is gone
    jz	_SetTimer			;if option not set, set the delay timer
    cmp	bBypassing, dword M_RELEASE	;   as a result of objOwner dying
    jnz	_SetTimer			; skip destruction option
    push	dword hWnd			;push the handle to destroy
    call	near fnDestroyWindow	;destroy the window

Align 4    
_SetTimer:					;delay memory release for a few ms
	mov	timerID, ebx 		;put our memory address in TimerID
	add	timerID, dword 48h	;now add offset to procedure
	push	dword timerID		;push that address
	push	dword 100			;push the timer interval
	push	esi				;push the timer id, will be assigned by windows
	push	esi				;push hWnd, don't use because window is probably destroyed
	call	near fnSetTimer		;set the timer
	mov	timerID, eax		;save the timerID

Align 4
_do_return:
    popad			    		;Pop all registers. lReturn is popped into eax
    ret     16			    	;Return with a 16 byte stack release

    
Align	4
_wndproc:			    		;Call the original WndProc
    call	_killTimer			;kill timer if it was set
    cmp	hWnd, esi			;no hWnd passed? Then this is a timer message
    jz	_free_memory		

    push    dword lParam	    	;ByVal lParam
    push    dword wParam	    	;ByVal wParam
    push    dword uMsg		    	;ByVal uMsg
    push    dword hWnd		    	;ByVal hWnd
    push    dword addrWndProc	    	;ByVal Address of the original WndProc

    call    near fnCallWinProc	;Call CallWindowProc
    mov     lReturn, eax	    	;Save the return value
    ret

Align 4
_free_memory:			    	;Free the memory this code is running in.... tricky

    pop     eax 		    		;Eat the call return address

    mov     uMsg, ebx		    	;VirtualFree param #1, start address of this memory
    mov     wParam, esi 	    	;VirtualFree param #2, 0
    mov     lParam, dword M_RELEASE ;VirtualFree param #3, memory release flag
    mov     eax, lRetAddr	    	;Return address of the code that called this thunk
    mov     bHandled, ebx	    	;ebx popped to edx after the popad instruction
    mov     hWnd, eax		    	;Return address to the code that called this thunk
    popad			    		;Restore the registers
    add     esp, 4		    	;Adjust the stack to point to the new return address
    jmp     dword fnVirtualFree     ;Jump to VirtualFree, ret to the caller of this thunk
    
Align	4
_callback:			    		;Validate the callback
    mov     ecx, [edi]		    	;ecx = table entry count
    jecxz   _generic_ret	    	;ecx = 0, table is empty

    test    ecx, ecx		    	;Set the flags as per ecx
    js	_call		    		;Table entry count is negative, all messages callback

    add     edi, 4		    	;Inc edi to point to the start of the callback table
    mov     eax, uMsg		    	;eax = the value to scan for
    repne   scasd		    		;Scan the callback table for uMsg
    jne     _generic_ret	    	;uMsg not in the callback table

Align	4
_call:				    	;Callback required, do it...

    lea     ecx, bHandled	    	;Address of bHandled
    lea     eax, lParamUser	    	;Address of lParamUser
    push    eax 		    		;ByRef lParamUser
    lea     eax, lReturn	    	;Address of lReturn
    push    dword lParam	    	;ByVal lParam
    push    dword wParam	    	;ByVal wParam
    push    dword uMsg		    	;ByVal uMsg
    push    dword hWnd		    	;ByVal hWnd
    push    eax 		    		;ByRef lReturn
    push    ecx 		    		;ByRef bHandled
    push    edx 		    		;ByVal bBefore
    push    dword objOwner	    	;ByVal the owner object
    call    near addrCallback	    	;Call the zWndProc callback procedure

Align	4
_generic_ret:			    	;Shared return (local only)
    ret

Align 4
_killTimer:
      cmp	timerID, esi		;is timer set? Called each time message is received
	jz	_generic_ret		;if not we just return
	push 	dword timerID		;else push the timer ID
	push	esi				;push hWnd
	call	near fnKillTimer		;kill the timer & return
	mov	timerID, esi		;zero the timerID
	ret

;*****************************************************************************************

; FYI: Usercontrol (UC) issues when subclassing same parent window multiple times & how timer prevents crash

; The general logic for unsubclassing goes like this:
;   bBypassing flag is set one of several ways:
;		in VB code in the ssc_UnThunk routine
;		result of receiving WM_NCDestroy
;		subclassing procedure dying (END or object set to Nothing)
;		IDE dies (END)
;   Flag checked every time thunk exits
;   When set and thunk is not being recursed, then 
;	thunk unsubclasses its hWnd after first forwarding any message on
;	if option to call DestroyWindow, then it is called if IDE died
;	thunk sets a timer to call back to itself
;	if thunk is re-entered, timer is killed and Flag checks begin again
;	when timer received, timer is killed & then thunk memory is released

; A = subclassed parent's original window procedure
; ucB loads & subclasses parent, B now parent's procedure, A is previous
; ucC loads & subclasses parent, C now parent's procedure, B is previous
; ucD loads & subclasses parent, D now parent's procedure, C is previous
; When unloading occurs in VB, the first loaded is first unloaded
;	and children are unloaded before the parent.....
; ucB unloads and changes parent's procedure to A - good, is original procedure
; ucC unloads and changes parent's procedure to B - bad because B is dead
; ucD unloads and changes parent's procedure to C - bad because C is dead
; parent unloads and message sent to C -- crash because C is dead

; Delay timer prevents this from occurring by keeping thunks alive and, within
;     this thunk, the bBypassing flag is set when UC was destroyed. So to
;     continue on, here is what happens when hWnd unloads using delay mechanism
; .... parent unloads and message sent to C (kept alive until timer fires)
; .... C gets message, Bypass is active, sets parent's procedure to B; message sent to B
; .... B gets message, Bypass is active, sets parent's procedure to A; message sent to A
; .... Timers fire and clear memory for C & B in that order

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
; Finally type at the command prompt:  NASMtoBIN subclass

; Download NASM compiler from following website
; http://sourceforge.net/project/showfiles.php?group_id=6208

;*****************************************************************************************