;*****************************************************************************************
;** Callback.asm - Generic Class/Form/UserControl callback thunk. Assemble with nasm.
;**
;** Paul_Caton@hotmail.com
;** Copyright free, use and abuse as you see fit.
;**
;** v1.0 The original............................................................ 20060408
;** v1.1 IDE safe................................................................ 20060409
;** v1.1 Validate that the callback address is live code......................... 20060413
;** By Ralph Eastwood (tcmreastwood@ntlworld.com) - Modified for FASM assembler.. 20060901
;*****************************************************************************************

;***********************
;Definitions
%define 	objOwner		[ebx]	    	;Owner object address
%define	addrCallback	[ebx + 4]  	;Callback address
%define 	fnEbMode		[ebx + 8]  	;EbMode address
%define	fnIsBadCodePtr	[ebx + 12]  ;EbMode address
%define	fnKillTimer		[ebx + 16]
%define	RetValue		[ebp - 4]  	;Callback/thunk return value

[Bits 32]
;************
;Data storage
    dd_objOwner 		dd 0	    ;Owner object address
    dd_addrCallback	dd 0	    ;Callback address
    dd_fnEbmode 		dd 0	    ;EbMode address
    dd_fnIsBadCodePtr	dd 0	    ;IsBadCodePtr address
    dd_fnKillTimer	dd 0	    ;
    
;***********
;Thunk start    
    mov     eax, esp		    	;Get a copy of esp as it is now
    pushad			    		;Push all the cpu registers on to the stack
    mov     ebx, 012345678h	    	;Address of the data, patched from VB
    mov     ebp, eax		    	;Set ebp to point to the return address
    
    push    dword addrCallback	;Callback address
    call    fnIsBadCodePtr    	;Call IsBadCodePtr
    jz     _check_ide		    	;If the callback code isn't live, return

    call 	_kill_timer
    jmp	_return

Align 4
_check_ide:    
    cmp     fnEbMode, dword 0	    ;Are we running in the IDE?
    jnz     _ide_state		    ;Check the IDE state

Align 4
_ide_running:	 
    mov     eax, ebp                ;Copy the stack frame pointer
    sub     eax, 4                  ;Address of the callback/thunk return value
    push    eax                     ;Push the return value address
nop
nop
nop    
    mov     ecx, 012345678h         ;Parameter count into ecx, patched from VB
    jecxz   _callback               ;If parameter count = 0, skip _parameter_loop

Align 4
_parameter_loop:
    push    dword [ebp + ecx * 4]   ;Push parameter
    loop    _parameter_loop	    ;Decrement ecx, if <> 0 jump to _parameter_loop

Align 4
_callback:    
    push    dword objOwner	    ;Owning object
    call    dword addrCallback	    ;Make the callback

Align 4
_return:	

    popad			    ;Restore registers
    ret     01234h		    ;Return, the number of esp stack bytes to release is patched from VB
    
Align 4
_ide_state:			    ;Running under the VB IDE
    call    near fnEbMode	    ;Determine the IDE state

    cmp     eax, dword 1           ;If EbMode = 1
    je	    _ide_running	    ;Running normally

    cmp     eax, dword 2            ;If EbMode = 2
    je      _breakpoint             ;Breakpoint

    call _kill_timer

Align 4
_breakpoint:    
    xor     eax, eax		    ;Zero eax
    mov     RetValue, eax	    ;Set the return value
    jmp     _return		    ;Outta here

Align 4
_kill_timer:
    push	dword [ebp + 12]
    push	dword [ebp + 4]
    call    fnKillTimer
    ret