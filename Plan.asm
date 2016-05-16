COMMENT #
	Planilha desenvolvida em Assembly na disciplina 
	Laboratório de Arquitetura e Organizacao de Computadores 2
    Autor: Leandro Novak
        
	// Funcao:	// Estado atual:
	cel 		Basicamente Implementada
	clr			Basicamente Implementada
	clt			Basicamente Implementada
	cop			Basicamente Implementada
	cut			Basicamente Implementada
	int			Basicamente Implementada
	max			Não Implementada
	med			Não Implementada
	min			Não Implementada
	mul			Não Implementada
	opn			Não Implementada
	sav			Não Implementada
	str			Basicamente Implementada
	sub			Basicamente Implementada
	sum			Basicamente Implementada
	ext			Basicamente Implementada
#

INCLUDE Irvine32.inc

CELL STRUCT
	typeData	BYTE 0
	integer		SDWORD 11111111111111111111111111111111b
	string		BYTE "Loremipsumdolorsit", 0
	formula		BYTE "123456789012345", 0
CELL ENDS

SIZECELL	EQU 40
SIZELINE	EQU 240
CELLINT		EQU 1
CELLSTR		EQU 5
CELLFOR		EQU 24

; Data segment
.data
readBuffer	BYTE 20 DUP(0), 0
tempBuffer	BYTE 20 DUP(0), 0
currentCell	DWORD 0
tempCell	DWORD 0
intString1	BYTE 3 DUP(0), 0
intString2	BYTE 3 DUP(0), 0
cellName	BYTE 4 DUP(0), 0
cell01		DWORD 0
cell02		DWORD 0

separator	BYTE 121 DUP (" "), 0
titlePlan	BYTE "                                                   Planilha ASM                                                       ", 0
columnBar	BYTE "|        A         |        B         |        C         |        D         |        E         |        F         |", 0
functions	BYTE "CEL CLR", 00h, "CLT", 00h, "COP CUT INT MAX MED MIN OPN SAV STR SUB SUM EXT", 00h, 0
selectdCell	BYTE "Celula selecionada: ", 0
viewFunc	BYTE "Funcao: ", 0
inputArrow	BYTE ">> ", 0
clearSpace	BYTE "     ", 0
whiteLine	BYTE "                                                                                                        ", 0
inputError	BYTE "Invalid input! Please type again: ", 0

line01 CELL 6 DUP(<>)
line02 CELL 6 DUP(<>)
line03 CELL 6 DUP(<>)
line04 CELL 6 DUP(<>)
line05 CELL 6 DUP(<>)
line06 CELL 6 DUP(<>)
line07 CELL 6 DUP(<>)
line08 CELL 6 DUP(<>)
line09 CELL 6 DUP(<>)
line10 CELL 6 DUP(<>)
line11 CELL 6 DUP(<>)
line12 CELL 6 DUP(<>)
line13 CELL 6 DUP(<>)
line14 CELL 6 DUP(<>)
line15 CELL 6 DUP(<>)
line16 CELL 6 DUP(<>)
line17 CELL 6 DUP(<>)
line18 CELL 6 DUP(<>)
line19 CELL 6 DUP(<>)
line20 CELL 6 DUP(<>)

teste BYTE "PERIGO!!!", 0
; Code segment
.code
main PROC
_mainLoop:
    mov ecx, 1
	call _DrawPlan
    call _DrawData
    call _ReadInput
    jecxz _exitPlan
    jmp _mainLoop
_exitPlan:
    exit
main ENDP

; Funcoes de uso geral:
; Le a entrada fornecida pelo usuario
_ReadInput PROC
	call _setColor1
	
	mov edx, offset readBuffer
	mov ecx, sizeof readBuffer
	call ReadString
	call _ClearInput		;Limpa a entrada do usuario
	invoke Str_copy, ADDR readBuffer, ADDR tempBuffer
	invoke Str_ucase, ADDR readBuffer

	
	mov ecx, [edx]			;Inicia a verificacao da funcao digitada
	mov ebx, offset functions

	cmp ecx, [ebx]			;CEL
	jnz _l0
	call _CelFunc
_l0:	
	cmp ecx, [ebx + 4]		;CLR
	jnz _l1
	call _ClrFunc
_l1:	
	cmp ecx, [ebx + 8]		;CLT
	jnz _l2
	call _CltFunc
_l2:	
	cmp ecx, [ebx + 12]		;COP
	jnz _l3
	call _CopFunc
_l3:	
	cmp ecx, [ebx + 16]		;CUT
	jnz _l4
	call _CutFunc
_l4:	
	cmp ecx, [ebx + 20]		;INT
	jnz _l5
	call _IntFunc
_l5:	
	cmp ecx, [ebx + 24]		;MAX
	jnz _l6
	call _MaxFunc
_l6:	
	cmp ecx, [ebx + 28]		;MED
	jnz _l7
	call _MedFunc
_l7:	
	cmp ecx, [ebx + 32]		;MIN
	jnz _l8
	call _MinFunc
_l8:	
	cmp ecx, [ebx + 36]		;OPN
	jnz _l9
	call _OpnFunc
_l9:	
	cmp ecx, [ebx + 40]		;SAV
	jnz _lA
	call _SavFunc
_lA:	
	cmp ecx, [ebx + 44]		;STR
	jnz _lB
	call _StrFunc
_lB:	
	cmp ecx, [ebx + 48]		;SUB
	jnz _lC
	call _SubFunc
_lC:	
	cmp ecx, [ebx + 52]		;SUM
	jnz _lD
	call _SumFunc
_lD:	
	cmp ecx, [ebx + 56]		;EXT
	jnz _lE
	mov ecx, 0
_lE:
	
	ret
_ReadInput ENDP

; Limpa a linha de entrada de usuário
_ClearInput PROC uses edx
	mov dh, 28
	mov dl, 4
	call Gotoxy
	mov edx, offset WhiteLine
	call WriteString
	ret
_ClearInput ENDP

; CELFUNC: Seleciona a celula informada pelo usuario
; Recebe em edx a entrada fornecida pelo usuario
_CelFunc PROC uses eax ebx ecx edx
	mov eax, [edx + 5]				;move para eax o nome da celula. ex: 01
	mov ecx, offset intString1
	mov [ecx], eax
	push edx
	mov edx, offset intString1		;transforma o conteudo de intString1 em um decimal e retorna em eax
	mov ecx, 3
	call ParseDecimal32
	pop edx
	dec eax
	cmp eax, 0
	jb _error1
	cmp eax, 19
	ja _error1
	push eax	
	mov ecx, 0
	movzx cx, BYTE PTR [edx + 4]
	cmp cx, 'A'
	jb _notValid
	cmp cx, 'Z'
	jb _continue1
_error1:
	jmp _notValid
_continue1:
	sub cx, 41h
	mov eax, 0
	mov ax, SIZECELL
	mul cx							;retorna em ax a coluna da celula como um inteiro binario
	movzx ecx, dx
	shl ecx, 16
	mov cx, ax						; o valor esta salvo em ecx
	pop eax							;traz de volta para eax a posicao da linha
	push ecx
	mov dx, SIZELINE
	mul dx
	movzx ecx, dx
	shl ecx, 16
	mov cx, ax
	pop eax
	add eax, ecx					;deslocamento total para a celula selecionada
	mov ecx, offset line01
	add ecx, eax
	mov ebx, offset currentCell
	mov [ebx], ecx
_notValid:
	ret
_CelFunc ENDP

; CLRFUNC: Limpa o conteudo da celula previamente selecionada
; Não recebe nada como parametro
_ClrFunc PROC uses eax ebx ecx edx
	mov eax, currentCell
	mov ebx, 0
	cmp ebx, eax
	jz _error
	mov ecx, 36
_clear:
	mov [eax], ebx 
	inc eax
	loop _clear
_error:
	ret
_ClrFunc ENDP

; CLTFUNC: Limpa todas as celulas da tabela
; Não recebe nada como parametro
_CltFunc PROC uses eax ebx ecx edx
	mov eax, offset line01
	mov ecx, 4800
	;mov ecx, 120
	mov ebx, 0
_clear:
	mov [eax], bl
	inc eax
	loop _clear
	ret
_CltFunc ENDP

; COPFUNC: Copia o conteudo de uma celula para outra
; Recebe a funcao de copia como parametro em edx
_CopFunc PROC uses eax ebx ecx edx
	mov eax, currentCell			; Se a celula de destino estiver vazia, salta como erro
	mov ebx, 0
	cmp ebx, eax
	jz _error1
	mov eax, [edx + 5]				;move para eax o nome da celula. ex: 01
	mov ecx, offset intString2
	mov [ecx], eax
	push edx
	mov edx, offset intString2		;transforma o conteudo de intString1 em um decimal e retorna em eax
	mov ecx, 3
	call ParseDecimal32
	pop edx
	dec eax
	cmp eax, 0
	jb _error1
	cmp eax, 19
	ja _error1
	push eax	
	mov ecx, 0
	movzx cx, BYTE PTR [edx + 4]
	cmp cx, 'A'
	jb _notValid
	cmp cx, 'Z'
	jb _continue1
_error1:
	jmp _notValid
_continue1:
	sub cx, 41h
	mov eax, 0
	mov ax, SIZECELL
	mul cx							;retorna em ax a coluna da celula como um inteiro binario
	movzx ecx, dx
	shl ecx, 16
	mov cx, ax						; o valor esta salvo em ecx
	pop eax							;traz de volta para eax a posicao da linha
	push ecx
	mov dx, SIZELINE
	mul dx
	movzx ecx, dx
	shl ecx, 16
	mov cx, ax
	pop eax
	add eax, ecx					;deslocamento total para a celula selecionada
	mov ecx, offset line01
	add ecx, eax
	mov ebx, offset tempCell
	mov [ebx], ecx

	mov eax, currentCell
	mov ebx, tempCell
	cmp ebx, eax
	jz _notValid
	mov ecx, SIZECELL - 1
	mov edx, 0
_move:								; Copia byte a byte o conteudo de uma celula para outra
	mov dl, [ebx]
	mov [eax], dl
	inc ebx
	inc eax
	loop _move
	mov ecx, 0
	mov eax, offset currentCell
	mov [eax], ecx
	mov ebx, offset tempCell
	mov [ebx], ecx
_notValid:	
	ret
_CopFunc ENDP

; COPFUNC: Copia o conteudo de uma celula para outra, limpando a celula de origem
; Recebe a funcao de recorte como parametro em edx
_CutFunc PROC uses eax ebx ecx edx
	mov eax, currentCell			; Se a celula de destino estiver vazia, salta como erro
	mov ebx, 0
	cmp ebx, eax
	jz _error1
	mov eax, [edx + 5]				;move para eax o nome da celula. ex: 01
	mov ecx, offset intString2
	mov [ecx], eax
	push edx
	mov edx, offset intString2		;transforma o conteudo de intString1 em um decimal e retorna em eax
	mov ecx, 3
	call ParseDecimal32
	pop edx
	dec eax
	cmp eax, 0
	jb _error1
	cmp eax, 19
	ja _error1
	push eax	
	mov ecx, 0
	movzx cx, BYTE PTR [edx + 4]
	cmp cx, 'A'
	jb _notValid
	cmp cx, 'Z'
	jb _continue1
_error1:
	jmp _notValid
_continue1:
	sub cx, 41h
	mov eax, 0
	mov ax, SIZECELL
	mul cx							;retorna em ax a coluna da celula como um inteiro binario
	movzx ecx, dx
	shl ecx, 16
	mov cx, ax						; o valor esta salvo em ecx
	pop eax							;traz de volta para eax a posicao da linha
	push ecx
	mov dx, SIZELINE
	mul dx
	movzx ecx, dx
	shl ecx, 16
	mov cx, ax
	pop eax
	add eax, ecx					;deslocamento total para a celula selecionada
	mov ecx, offset line01
	add ecx, eax
	mov ebx, offset tempCell
	mov [ebx], ecx
	mov eax, currentCell
	mov ebx, tempCell
	cmp ebx, eax
	jz _notValid
	mov ecx, SIZECELL - 1
	mov edx, 0
_move:								; Copia byte a byte o conteudo de uma celula para outra
	mov dl, [ebx]
	mov [eax], dl
	inc ebx
	inc eax
	loop _move
	mov ecx, 0
	mov eax, offset currentCell
	mov [eax], ecx
	mov eax, tempCell
	mov ebx, 0
	cmp ebx, eax
	jz _notValid
	mov ecx, 36
_clear:								; Limpa a celula de origem
	mov [eax], ebx 
	inc eax
	loop _clear
_notValid:
	ret
_CutFunc ENDP

; INTFUNC: Armazena um numero inteiro na celula previamente selecionada pelo usuario
; Recebe a funcao de insercao de inteiro como parametro em edx
_IntFunc PROC uses eax ebx ecx edx
	mov eax, currentCell			; Se a celula de destino estiver vazia, salta como erro
	mov ebx, 0
	cmp ebx, eax
	jz _error1
	mov ebx, 1
	mov [eax], bl
	add eax, CELLINT
	mov ebx, eax
	mov edx, offset readBuffer
	add edx, 4
	mov ecx, 10
	call ParseInteger32
	mov [ebx], eax
_error1:
	ret
_IntFunc ENDP

_MaxFunc PROC
	ret
_MaxFunc ENDP

_MedFunc PROC
	ret
_MedFunc ENDP

_MinFunc PROC
	ret
_MinFunc ENDP

_OpnFunc PROC
	ret
_OpnFunc ENDP

_SavFunc PROC
	ret
_SavFunc ENDP

; STRFUNC: Armazena uma string na celula previamente selecionada pelo usuario
; Obtem a string a partir da copia da entrada de usuario
_StrFunc PROC uses eax ebx ecx edx esi
	mov eax, currentCell			; Se a celula de destino estiver vazia, salta como erro
	mov ebx, 0
	cmp ebx, eax
	jz _error1
	mov ebx, 2
	mov [eax], ebx
	add eax, CELLSTR
	mov edx, offset tempBuffer
	add edx, 4
	mov ecx, 18
	mov ebx, 0
	mov esi, 0Ah
_copy:
	cmp [edx], esi
	jz _error1
	mov bl, [edx]
	mov [eax], bl
	inc eax
	inc edx
	loop _copy
_error1:
	ret
_StrFunc ENDP

; SUBFUNC: Subtrai da celula destino o valor da celula origem
; Recebe a funcao de subtracao como parametro em edx
_SubFunc PROC uses eax ebx ecx edx
mov eax, currentCell			; Se a celula de destino estiver vazia, salta como erro
	mov ebx, 0
	cmp ebx, eax
	jz _error1
	mov eax, [edx + 5]				;move para eax o nome da celula. ex: 01
	mov ecx, offset intString2
	mov [ecx], eax
	push edx
	mov edx, offset intString2		;transforma o conteudo de intString1 em um decimal e retorna em eax
	mov ecx, 3
	call ParseDecimal32
	pop edx
	dec eax
	cmp eax, 0
	jb _error1
	cmp eax, 19
	ja _error1
	push eax	
	mov ecx, 0
	movzx cx, BYTE PTR [edx + 4]
	cmp cx, 'A'
	jb _notValid
	cmp cx, 'Z'
	jb _continue1
_error1:
	jmp _notValid
_continue1:
	sub cx, 41h
	mov eax, 0
	mov ax, SIZECELL
	mul cx							;retorna em ax a coluna da celula como um inteiro binario
	movzx ecx, dx
	shl ecx, 16
	mov cx, ax						; o valor esta salvo em ecx
	pop eax							;traz de volta para eax a posicao da linha
	push ecx
	mov dx, SIZELINE
	mul dx
	movzx ecx, dx
	shl ecx, 16
	mov cx, ax
	pop eax
	add eax, ecx					;deslocamento total para a celula selecionada
	mov ecx, offset line01
	add ecx, eax
	mov ebx, offset tempCell
	mov [ebx], ecx
	mov ebx, tempCell
	mov eax, currentCell
	mov edx, 1
	cmp [eax], dl
	jne _notValid
	cmp [ebx], dl
	jne _notValid
	add eax, CELLINT
	add ebx, CELLINT
	mov edx, [ebx]
	sub [eax], edx
_notValid:
	ret
_SubFunc ENDP

; SUBFUNC: Soma da celula destino o valor da celula origem
; Recebe a funcao de soma como parametro em edx
_SumFunc PROC uses eax ebx ecx edx
	mov eax, currentCell			; Se a celula de destino estiver vazia, salta como erro
	mov ebx, 0
	cmp ebx, eax
	jz _error1
	mov eax, [edx + 5]				;move para eax o nome da celula. ex: 01
	mov ecx, offset intString2
	mov [ecx], eax
	push edx
	mov edx, offset intString2		;transforma o conteudo de intString1 em um decimal e retorna em eax
	mov ecx, 3
	call ParseDecimal32
	pop edx
	dec eax
	cmp eax, 0
	jb _error1
	cmp eax, 19
	ja _error1
	push eax	
	mov ecx, 0
	movzx cx, BYTE PTR [edx + 4]
	cmp cx, 'A'
	jb _notValid
	cmp cx, 'Z'
	jb _continue1
_error1:
	jmp _notValid
_continue1:
	sub cx, 41h
	mov eax, 0
	mov ax, SIZECELL
	mul cx							;retorna em ax a coluna da celula como um inteiro binario
	movzx ecx, dx
	shl ecx, 16
	mov cx, ax						; o valor esta salvo em ecx
	pop eax							;traz de volta para eax a posicao da linha
	push ecx
	mov dx, SIZELINE
	mul dx
	movzx ecx, dx
	shl ecx, 16
	mov cx, ax
	pop eax
	add eax, ecx					;deslocamento total para a celula selecionada
	mov ecx, offset line01
	add ecx, eax
	mov ebx, offset tempCell
	mov [ebx], ecx
	mov ebx, tempCell
	mov eax, currentCell
	mov edx, 1
	cmp [eax], dl
	jne _notValid
	cmp [ebx], dl
	jne _notValid
	add eax, CELLINT
	add ebx, CELLINT
	mov edx, [ebx]
	add [eax], edx
_notValid:
	ret
_SumFunc ENDP

; Funcoes de interface
; Exibe a base da planilha
_DrawPlan PROC uses eax ebx ecx edx esi
	mov dh, 0
	mov dl, 0
	call Gotoxy
	call _setColor1
	call Clrscr
	call _setColor2
	mov edx, offset separator		; Desenha uma linha em "branco" abaixo do titulo
	call WriteString
	mov edx, offset titlePlan		; Desenha o titulo da planilha
	call WriteString
	mov edx, offset separator		; Desenha uma linha em "branco" abaixo do titulo
	call WriteString
	call _setColor1					; Desenha a saida da funcao presente na celula
	call Crlf
	mov edx, offset viewFunc
	call WriteString
	call Crlf
	call _setColor2
	mov edx, offset clearSpace		; Desenha as a base das linhas da tabela
	call WriteString
	mov edx, offset columnBar
	call WriteString
	mov ecx, 20
	mov edx, offset clearSpace
_drawSideBar:						; Desenha as colunas da tabela
	call WriteString
	call Crlf
	loop _drawSideBar
	mov dh, 6
	mov dl, 1
	mov eax, 1b
	mov ecx, 20	
_drawLines:							; Desenha as linhas da tabela
	call GotoXY
	call WriteDec
	call Crlf
	inc eax
	inc dh
	loop _drawLines
	mov eax, 00010111b				; Exibe as setas de entrada de usuario
	call SetTextColor
	mov dh, 28
	mov dl, 1
	call GotoXY
	mov edx, offset inputArrow
	call WriteString
	ret
_DrawPlan ENDP

; Exibe nas celulas da planilha o conteudo existente em cada uma delas
; Verifica para cada celula o tipo de conteudo existente nesta
_DrawData PROC uses eax ebx ecx edx esi
	call _setColor1
	mov dh, 3
	mov dl, 0
	call gotoxy
	mov edx, offset selectdCell
	call WriteString
	mov edx, offset cellName
	call WriteString
	mov dh, 6
	mov dl, 6
	mov eax, offset line01
	mov ecx, 20
_drawLines:
	push ecx
	mov ecx, 6
_drawColumns:
	push ecx
	call gotoxy
_verifyType:
	movzx ecx, BYTE PTR [eax]
	jecxz _drawNothing
	cmp ecx, 1
	jz _drawIntegers
	cmp ecx, 2
	jz _drawStrings
	jmp _drawNothing
_drawIntegers:
	push eax
	pop ebx
	mov eax, [ebx + CELLINT]
	call WriteInt
	push ebx
	pop eax	
	jmp _drawNothing
_drawStrings:
	push edx
	mov edx, eax
	add edx, CELLSTR
	call WriteString
	pop edx
_drawNothing:
	pop ecx
	add dl, 19
	cmp dl, 78h
	jnz _continue
	mov dl, 6
	inc dh
_continue:
	add eax, SIZECELL
	loop _drawColumns
	pop ecx
	loop _drawLines
	mov dh, 28
	mov dl, 4
	call gotoxy
	ret
_DrawData ENDP

; Funcao para atribuir cor a planilha
_SetColor1 PROC
	mov eax, 17h
	call SetTextColor
	ret
_SetColor1 ENDP

; Outra funcao para atribuir cor a planilha
_SetColor2 PROC
	mov eax, 71h
	call SetTextColor
	ret
_SetColor2 ENDP
END main
