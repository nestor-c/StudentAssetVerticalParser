; hello_world.asm

; Author: Nestor Cortes
; Date: 02-May-2021

section .text:

section .data:
	message: db "Hello World!", 0xA
	message_length equ $-message
