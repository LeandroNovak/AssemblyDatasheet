COMMENT #
	Planilha desenvolvida em Assembly na disciplina v.3
	Laboratório de Arquitetura e Organizacao de Computadores 2
    Autor: Leandro Novak
        
	Funcao:		Estado atual:
	CEL			Funcional
	CLR			Funcional
	CLT			Funcional
	COP			Funcional
	CUT			Funcional
	EXT			Funcional
	FLT			Funcional
	FNC			Funcional
	HLP			Funcional
	INT			Funcional
	MAX			Em implementacao
	MED			Em implementacao
	MIN			Em implementacao
	OPN			Em implementacao
	SAV			Em implementacao
	STR			Funcional
	SUB			Funcional
	SUM			Funcional

	Obs. Aplicacao feita para janelas com tamanho 120 x 30.
#

INCLUDE Irvine32.inc				; Biblioteca padrão do Irvine
INCLUDE Macros.inc					; Biblioteca para permitir o uso da macro mWrite

CELL STRUCT
	_type		BYTE 0				; 0 = empty, 1 = integer, 2 = real, 3 = string, 4 = formulaInt, 5 = formulaReal
	_int		SDWORD 2			; Area de inteiros
	_float		REAL4 5.0			; Area de reais
	_string		BYTE 18 DUP(0), 0	; Area de string
	_formula	BYTE 18 DUP(0), 0	; Area de funcoes
CELL ENDS

SIZECELL		EQU 47				; Tamanho de uma celula em bytes
SIZELINE		EQU 282				; Tamanho de uma linha (6 celulas)
INTPOS			EQU 1				; Tamanho do deslocamento para a area de inteiros
FLTPOS			EQU 5				; Tamanho do deslocamento para a area de reais
STRPOS			EQU 9				; Tamanho do deslocamento para a area de strings
FNCPOS			EQU 28				; Tamanho do deslocamento para a area de funcoes
TOTALCELLS		EQU 5640			; Total de bytes ocupados por todas as celulas da planilha

.data
; Variaveis para string de entrada, uma para string e outra para backup da string sem uppercase
inputBuffer		BYTE 18 DUP(0), 0
tempBuffer		BYTE 18 DUP(0), 0

; Variaveis relacionadas a celula selecionada (offset e flag de celula selecionada)
offsetCell		DWORD 0
cellName		DWORD 0
isSelectedCell	BYTE 0

; Variaveis para manipulacao de arquivos
filename  		BYTE 16 DUP(0)
filehandle  	DWORD 0

; Variavel para lixo gerado pela exibicao de valores reais
trashReal		REAL4 0.0

; Area de dados para celulas da planilha
line1	CELL 6 DUP(<>)
line2	CELL 6 DUP(<>)
line3	CELL 6 DUP(<>)
line4	CELL 6 DUP(<>)
line5	CELL 6 DUP(<>)
line6	CELL 6 DUP(<>)
line7	CELL 6 DUP(<>)
line8	CELL 6 DUP(<>)
line9	CELL 6 DUP(<>)
line10	CELL 6 DUP(<>)
line11	CELL 6 DUP(<>)
line12	CELL 6 DUP(<>)
line13	CELL 6 DUP(<>)
line14	CELL 6 DUP(<>)
line15	CELL 6 DUP(<>)
line16	CELL 6 DUP(<>)
line17	CELL 6 DUP(<>)
line18	CELL 6 DUP(<>)
line19	CELL 6 DUP(<>)
line20	CELL 6 DUP(<>)

; Strings para UI
separator		BYTE 120 DUP(" "), 0
titlePlan		BYTE 54 DUP(" "), "Planilha ASM", 54 DUP(" "), 0
columnBar		BYTE "|        A         |        B         |        C         |        D         |        E         |        F         |", 0
whitecell		BYTE 18 DUP(" ")
functions		BYTE "CEL CLR", 00h, "CLT", 00h, "COP CUT EXT", 00h, "FLT", 00h, "FNC", 00h, "HLP", 00h, "INT", 00h, "MAX MED MIN OPN SAV STR", 00H, "SUB SUM ", 00h, 0
selectdCell		BYTE "Celula selecionada: ", 0
viewFunc		BYTE "Funcao: ", 0
inputArrow		BYTE ">> ", 0
clearSpace		BYTE 5 DUP(" "), 0


.code

; Prototipos das funcoes
ReadInput		PROTO stringPtr: PTR DWORD
ProcessString	PROTO stringPtr: PTR DWORD

Func_CEL		PROTO stringPTR: PTR DWORD
Func_CLR		PROTO
Func_CLT		PROTO
Func_COP		PROTO stringPTR: PTR DWORD
Func_CUT 		PROTO StringPTR: PTR DWORD
Func_EXT		PROTO
Func_HLP		PROTO
Func_FNC		PROTO
Func_FLT		PROTO
Func_INT 		PROTO
Func_MAX 		PROTO
Func_MED 		PROTO
Func_MIN 		PROTO
Func_OPN 		PROTO
Func_SAV 		PROTO
Func_STR 		PROTO
Func_SUB 		PROTO StringPTR: PTR DWORD
Func_SUM 		PROTO StringPTR: PTR DWORD

ClearSheet		PROTO
DrawBase		PROTO
DrawData		PROTO
GetOffsetCell	PROTO cellAsString: DWORD
GotoPosXY		PROTO posx: DWORD, posy: DWORD
SetColor		PROTO color: DWORD
Update			PROTO


; Funcao main
main PROC
	finit
	invoke SetColor, (lightGray + (16 * blue))
	call Clrscr
	call DrawBase
 _run:
	mov ecx, 1
	invoke SetColor, (blue + (16 * lightGray))
	invoke ClearSheet
	invoke DrawData
	invoke ReadInput, addr inputBuffer
	jecxz _exit
	jmp _run
 _exit:
	exit
main ENDP

; Le o comando digitado pelo usuario
ReadInput PROC USES EDX, stringPtr: PTR DWORD
	mov ecx, 18
	invoke SetColor, (lightGray + (16 * blue))
	invoke GoToPosXY, 0, 28
	mov edx, offset separator
	call WriteString
	invoke GoToPosXY, 1, 28
	mov edx, offset inputArrow
	call WriteString
	invoke GoToPosXY, 4, 28

	mov edx, stringPtr
	call ReadString
	invoke Str_copy, stringPtr, addr tempBuffer
	invoke Str_ucase, stringPtr
	invoke ProcessString, stringPtr
	call Crlf
	ret
ReadInput ENDP

; Processa a entrada do usuario
ProcessString PROC USES EAX EDX, stringPtr: PTR DWORD
	;functions		CEL CLR", 00h, "CLT", 00h, "COP CUT EXT", 00h, "FLT", 00h, "FNC", 00h, "HLP", 00h, "INT", 00h, "MAX MED MIN OPN SAV STR", 00H, "SUB", 00h, "SUM", 00h, 0
	mov edx, stringPtr
	mov eax, [edx]

	mov edx, offset functions
	cmp eax, [edx]
	jnz _clr
	invoke Func_CEL, stringPtr
	jmp _end
 _clr:
	add edx, 4
	cmp eax, [edx]
	jnz _clt
	invoke Func_CLR
	jmp _end
 _clt:
	add edx, 4
	cmp eax, [edx]
	jnz _cop
	invoke Func_CLT
	jmp _end
 _cop:
	add edx, 4
	cmp eax, [edx]
	jnz _cut
	invoke Func_COP, stringPtr
	jmp _end
 _cut:
	add edx, 4
	cmp eax, [edx]
	jnz _ext
	invoke Func_CUT, stringPtr
	jmp _end
 _ext:
	add edx, 4
	cmp eax, [edx]
	jnz _fnc
	invoke Func_EXT
	jmp _end
 _fnc:
 	add edx, 4
	cmp eax, [edx]
	jnz _flt
	invoke Func_FNC
	jmp _end
 _flt:
 	add edx, 4
 	jnz _hlp
	invoke Func_FLT
	jmp _end
 _hlp:
	add edx, 4
	cmp eax, [edx]
	jnz _int
	invoke Func_HLP
	jmp _end
 _int:
	add edx, 4
	cmp eax, [edx]
	jnz _max
	invoke Func_INT
	jmp _end
 _max:
	add edx, 4
	cmp eax, [edx]
	jnz _med
	invoke Func_MAX
	jmp _end
 _med:
	add edx, 4
	cmp eax, [edx]
	jnz _min
	invoke Func_MED
	jmp _end
 _min:
	add edx, 4
	cmp eax, [edx]
	jnz _opn
	invoke Func_MIN
	jmp _end
 _opn:
	add edx, 4
	cmp eax, [edx]
	jnz _sav
	invoke Func_OPN
	jmp _end
 _sav:
	add edx, 4
	cmp eax, [edx]
	jnz _str
	invoke Func_SAV
	jmp _end
 _str:
	add edx, 4
	cmp eax, [edx]
	jnz _sub
	invoke Func_STR
	jmp _end
 _sub:
	add edx, 4
	cmp eax, [edx]
	jnz _sum
	invoke Func_SUB, StringPTR
	jmp _end
 _sum:
	add edx, 4
	cmp eax, [edx]
	jnz _error
	invoke Func_SUM, StringPTR
	jmp _end
 _error:
	invoke SetColor, (lightGray + (16 * blue))
	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "OPERACAO INVALIDA!"		; Mensagem de erro

 _end:
	ret
ProcessString ENDP

; Seleciona a celula digitada salvando em offsetCell o offset da celula desejada
Func_CEL PROC USES EBX ECX EDX, StringPTR: PTR DWORD
	invoke SetColor, (lightGray + (16 * blue))
	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString

	mov edx, stringPTR
	add edx, 4										; "Aponta" EDX para o inicio da area da string que representa a celula
	xor eax, eax									; Limpa EAX
	mov al, [edx]									; Move para AL o caractere que representa a coluna
	sub al, 41h										; Subtrai 41h para converter A em decimal, onde A = 0, B = 1...

	cmp al, -1										; Verifica a validade do valor em EAX
	jle _error										; Fora da area de celulas
	cmp al, 6										; Verifica a validade do valor em EAX
	jge _error										; Fora da area de celulas

	mov ebx, SIZECELL								; Move para EBX o tamanho de um struct CELL
	push edx										; Salva o valor de EDX antes da multiplicacao
	mul ebx											; Multiplica EAX pelo tamanho de uma celula, deslocando para a coluna desejada
	pop edx											; Restitui o valor de EDX
	inc edx											; "Aponta" EDX para o trecho da string que representa a linha
	mov ecx, 2										; Atribui a ECX 2, que eh o numero de digitos que representam a linha
	push eax										; Salva o valor de EAX
	call ParseInteger32								; Transforma o numero da string em um inteiro e atribui a EAX
	dec eax											; Ajusta o indice da celula que inicia a contagem em 0

	cmp eax, -1										; Verifica a validade do valor em EAX
	jle _error										; Fora da area de celulas
	cmp eax, 20										; Verifica a validade do valor em EAX
	jge _error										; Fora da area de celulas

	mov ebx, SIZELINE								; Move para EDX o tamanho em bytes de uma linha
	mul ebx											; Multiplica eax gerando o deslocamento para a linha correta
	pop ebx											; Retorna o valor anteriormente em EAX para EBX
	add eax, ebx									; Adiciona o deslocamento de columas ao deslocamento de linhas
	mov edx, offset line1							; Obtem o offset da primeira celula da primeira linha
	add eax, edx									; Adiciona o deslocamento ao offset obtido
	
	mov edx, stringPTR							; "Aponta" EDX para o inicio da area da string que representa a celula
	add edx, 4
	mov ecx, [edx]
	mov edx, offset cellName						; Salva o nome da celula selecionada
	mov [edx], ecx
	
	mov edx, offset offsetCell
	mov [edx], eax									; Salva o offset da celula selecionada em offsetCell

	mov edx, offset isSelectedCell					; Seta o "FLAG" de celula selecionada
	push eax
	mov al, 1
	mov [edx], al
	pop eax

	invoke GotoPosXY, 1, 27
	mWrite "CELULA SELECIONADA COM SUCESSO!"		; Mensagem de sucesso
	jmp _end
 _error:
	invoke GotoPosXY, 1, 27
	mWrite "ERRO AO SELECIONAR CELULA!"				; Mensagem de erro
 _end:
	ret
Func_CEL ENDP

; Limpa todo o conteudo de uma celula
Func_CLR PROC USES EAX EBX ECX EDX
	invoke SetColor, (lightGray + (16 * blue))
	invoke GoToPosXY, 0, 26
	mov edx, offset separator
	call WriteString
	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString

	xor eax, eax
	mov al, isSelectedCell
	cmp eax, 0
	jz _error
	mov ebx, 0
	mov ecx, 10
	mov eax, offsetCell
 _clear:
	mov [eax], ebx
	add eax, 4
	loop _clear
	jmp _sucess
 _error:
	invoke GotoPosXY, 1, 27
	mWrite "SELECIONE UMA CELULA PARA LIMPAR!"
	jmp _end
 _sucess:
	invoke GotoPosXY, 1, 27
	mWrite "CELULA LIMPA COM SUCESSO!"
 _end:
	ret
Func_CLR ENDP

; Limpa todo o conteudo da planilha
Func_CLT PROC USES EAX ECX EDX
	invoke GotoPosXY, 1, 26
	mWrite "DESEJA REALMENTE LIMPAR TODA A PLANILHA? ESTA OPERACAO NAO PODE SER DESFEITA."
	invoke GotoPosXY, 1, 27
	mWrite "Pressione 'S' para continuar ou 'N' para cancelar: "
	call ReadChar
	cmp al, 'n'
	jz _cancel
	cmp al, 'N'
	jz _cancel
	cmp al, 's'
	jz _continue
	cmp al, 'S'
	jz _continue
 _error:
	mov edx, offset separator
	invoke GotoPosXY, 0, 26
	call WriteString
	invoke GotoPosXY, 0, 27
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "OPCAO DIGITADA INVALIDA! A planilha nao sera limpa."
	jmp _end
 _cancel:
 	invoke GotoPosXY, 0, 26
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 0, 27
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "OPERACAO CANCELADA. A PLANILHA NAO FOI LIMPA."
	jmp _end
 _continue:
	xor eax, eax
	mov edx, offset line1
	mov ecx, 1410
	mov eax, 0
 _clear:
	mov [edx], eax
	add edx, 4
	loop _clear
 _sucess:
	mov edx, offset separator
	invoke GotoPosXY, 0, 26
	call WriteString
	invoke GotoPosXY, 0, 27
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "PLANILHA LIMPA COM SUCESSO!"
 _end:
	ret
Func_CLT ENDP

; Copia o conteudo da celula origem para a celula destino.
; Ex. COP A10 A20 copia o conteudo da celula A10 para a celula A20
Func_COP PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offset1: DWORD
	LOCAL offset2: DWORD
	mov edx, StringPTR

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx
	cmp eax, 0
	jz _error
	lea ebx, offset1
	mov [ebx], eax

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx
	cmp eax, 0
	jz _error
	lea ebx, offset2
	mov [ebx], eax

	mov eax, offset1
	mov ebx, offset2
	mov ecx, 10
 _move:
	mov edx, [eax]
	mov [ebx], edx
	add eax, 4
	add ebx, 4
	loop _move
	jmp _sucess
 _error:
	invoke GotoPosXY, 1, 27
	mWrite "Alguma das celulas selecionadas eh invalida! Tente novamente."
	jmp _exit
 _sucess:
 	invoke GotoPosXY, 1, 27
	mWrite "Copia efetuada com sucesso"
 _exit:
	ret
Func_COP ENDP

; Move o conteudo da celula origem para a celula destino.
; Ex. CUT A10 A20 copia o conteudo da celula A10 para a celula A20 e apaga o conteudo de A10
Func_CUT PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offset1: DWORD
	LOCAL offset2: DWORD
	mov edx, StringPTR

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx
	cmp eax, 0
	jz _error
	lea ebx, offset1
	mov [ebx], eax

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx
	cmp eax, 0
	jz _error
	lea ebx, offset2
	mov [ebx], eax

	mov eax, offset1
	mov ebx, offset2
	mov ecx, 10
 _move:
	mov edx, [eax]
	mov [ebx], edx
	add eax, 4
	add ebx, 4
	loop _move
	jmp _sucess
 _error:
	invoke GotoPosXY, 1, 27
	mWrite "Alguma das celulas selecionadas eh invalida! Tente novamente."
	jmp _exit
 _sucess:
 	invoke GotoPosXY, 1, 27
	mWrite "Copia efetuada com sucesso"

	mov ebx, 0
	mov ecx, 10
	mov eax, offset1
 _clear:
	mov [eax], ebx
	add eax, 4
	loop _clear
 _exit:
	ret
Func_CUT ENDP

; Encerra a execucao da planilha
; Nao recebe parametro
; Retorna ECX = 0 para sair ou ECX = 1 para continuar na planilha
Func_EXT PROC USES EAX
	invoke GotoPosXY, 1, 26
	mWrite "DESEJA REALMENTE FECHAR A PLANILHA?"
	invoke GotoPosXY, 1, 27
	mWrite "Pressione 'S' para sair ou 'N' para continuar na planilha: "
	call ReadChar
	cmp al, 'n'
	jz _cancel
	cmp al, 'N'
	jz _cancel
	cmp al, 's'
	jz _continue
	cmp al, 'S'
	jz _continue
	invoke GotoPosXY, 1, 26
	mWrite "                       "
	invoke GotoPosXY, 1, 27
	mWrite "Opcao invalida!"
	jmp _cancel
 _cancel:
	mov ecx, 1
	jmp _exit
 _continue:
 	mov ecx, 0
 _exit:
 	invoke SetColor, (lightGray + (16 * blue))
	invoke GoToPosXY, 0, 26
	mov edx, offset separator
	call WriteString
	invoke GoToPosXY, 0, 27
	call WriteString
	ret
Func_EXT ENDP

; Atribui um valor real para a celula previamente selecionada
Func_FLT PROC USES EAX EBX ECX EDX
	xor ecx, ecx
	mov cl, isSelectedCell
	jecxz _error

	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27
	mWrite "Digite o numero real inteiro desejado: "

	mov edx, offsetCell
	mov cl, 2
	mov [edx], cl
	add edx, FLTPOS
	call ReadFloat
	fstp REAL4 PTR [edx]
	jmp _sucess
 _error:
 	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "SELECIONE UMA CELULA PRIMEIRO!"
	jmp _end
 _sucess:
  	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString
 	invoke GotoPosXY, 1, 27
	mWrite "Atribuicao de inteiro efetuada com sucesso!"
 _end:
	ret
Func_FLT ENDP

; Atribui uma funcao para a celula previamente selecionada (a funcao deve ser valida para atribuicao)
Func_FNC PROC USES EAX ECX EDX
	xor ecx, ecx
	mov cl, isSelectedCell
	jecxz _error
	
	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27
	mWrite "Digite a funcao desejada: "

	mov edx, offsetCell
	mov cl, 1
	mov [edx], cl
	mov ecx, 18
	add edx, FNCPOS
	call ReadString
	mov [edx], eax
	jmp _sucess
 _error:
  	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "SELECIONE UMA CELULA PRIMEIRO!"
	jmp _end
 _sucess:
  	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString
 	invoke GotoPosXY, 1, 27
	mWrite "Atribuicao de funcao efetuada com sucesso!"
 _end:
	ret
Func_FNC ENDP

; Exibe a tela de ajuda para o usuario
Func_HLP PROC USES EAX EBX ECX EDX
	invoke SetColor, (lightGray + (16 * blue))
	call Clrscr
	invoke gotoPosXY, 0, 0
	invoke SetColor, (blue + (16 * lightGray))
	mov edx, offset separator
	call WriteString
	call WriteString
	call WriteString
	invoke gotoPosXY, 58, 1
	mWrite "AJUDA"
	invoke SetColor, (lightGray + (16 * blue))
	call Crlf
	call Crlf
	invoke gotoPosXY, 1, 4
	mWrite "CEL - Seleciona uma celula da planilha. Ex. CEL A01 (Seleciona a celula A01)."
	invoke gotoPosXY, 1, 5
	mWrite "CLR - Limpa o conteudo da celula previamente selecionada."
	invoke gotoPosXY, 1, 6
	mWrite "CLT - Limpa o conteudo de todas as celulas da tabela."
	invoke gotoPosXY, 1, 7
	mWrite "COP - Copia todo o conteudo de uma celula para outra (incluindo funcoes)."
	invoke gotoPosXY, 1, 8
	mWrite "CUT - Move o conteudo de uma celula para outra. Semelhante a COP mas limpa a celula de origem."
	invoke gotoPosXY, 1, 9
	mWrite "INT - Prefixo para atribuir um valor inteiro a uma celula."
	invoke gotoPosXY, 1, 10
	mWrite "MAX - Retorna o maior valor entre duas celulas."
	invoke gotoPosXY, 1, 11
	mWrite "MED - Calcula a media de valores das celulas."
	invoke gotoPosXY, 1, 12
	mWrite "MIN - Retorna o menor valor de duas celulas."
	invoke gotoPosXY, 1, 13
	mWrite "OPN - Abre uma planilha salva no disco rigido."
	invoke gotoPosXY, 1, 14
	mWrite "SAV - Salva a planilha atual no disco rigido."
	invoke gotoPosXY, 1, 15
	mWrite "STR - Prefixo para atribuir uma string a uma celula"
	invoke gotoPosXY, 1, 16
	mWrite "SUB - Subtrai o valor de uma celula de outra (inteiro ou racional)."
	invoke gotoPosXY, 1, 17
	mWrite "SUM - Soma o valor de uma celula com outra (inteiro ou racional)."
	invoke gotoPosXY, 1, 18
	mWrite "EXT - Encerra a execucao da planilha"
	invoke gotoPosXY, 1, 19
	mWrite "HLP - Exibe a tela de ajuda."
 _input:
	invoke GotoPosXY, 1, 27
	mWrite "Pressione 'E' para retornar a planilha: "
	call ReadChar
	cmp al, 'E'
	jz _exit
	cmp al, 'e'
	jz _exit
	invoke GotoPosXY, 1, 26
	mWrite "A tecla pressionada eh invalida."
	jmp _input
 _exit:
	invoke SetColor, (lightGray + (16 * blue))
	call Clrscr
	call DrawBase
	ret
Func_HLP ENDP

; Atribui um valor inteiro com sinal para a celula previamente selecionada
Func_INT PROC USES EAX EBX ECX EDX
	xor ecx, ecx
	mov cl, isSelectedCell
	jecxz _error
	
	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27
	mWrite "Digite o numero inteiro desejado: "

	mov edx, offsetCell
	mov cl, 1
	mov [edx], cl
	add edx, INTPOS
	call ReadInt
	mov [edx], eax
	jmp _sucess
 _error:
  	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "SELECIONE UMA CELULA PRIMEIRO!"
	jmp _end
 _sucess:
  	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString
 	invoke GotoPosXY, 1, 27
	mWrite "Atribuicao de inteiro efetuada com sucesso!"
 _end:
	ret
Func_INT ENDP

Func_MAX PROC
	ret
Func_MAX ENDP

Func_MED PROC
	ret
Func_MED ENDP

Func_MIN PROC
	ret
Func_MIN ENDP

Func_OPN PROC
	ret
Func_OPN ENDP

Func_SAV PROC

	ret
Func_SAV ENDP

; Atribui uma string para a celula previamente selecionada
Func_STR PROC USES ECX EDX
	xor ecx, ecx							; Limpa ECX
	mov cl, isSelectedCell					; Verifica se ha alguma celula selecionada (flag != 0)
	jecxz _error

	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	
	invoke GotoPosXY, 1, 27
	mWrite "Digite a string desejada: "

	mov edx, offsetCell						; Move para ECX o offset da celula selecionada
	mov cl, 3								; Valor que seta a flag de string da celula
	mov [edx], cl							; Seta a flag de string
	add edx, STRPOS
	push edx
	mov ecx, 18
 _clearString:
	mov [edx], cl
	inc edx
	loop _clearString
	pop edx
	mov ecx, 19
	call ReadString
	jmp _sucess
 _error:
  	invoke GotoPosXY, 1, 27
 	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "SELECIONE UMA CELULA PRIMEIRO!"
	jmp _end
 _sucess:
 	invoke GotoPosXY, 1, 27
 	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "String atribuida com sucesso!"
 _end:
	ret
Func_STR ENDP

; Subtrai o valor de duas celulas e atribui ah celula selecionada
Func_SUB PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offset1: DWORD
	LOCAL offset2: DWORD
	LOCAL flagval: BYTE

	mov edx, StringPTR

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx			; Obtem o offset da primeira celula
	cmp eax, 0
	jz _error
	lea ebx, offset1
	mov [ebx], eax						; Salva o offset da primeira celula em offset1

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx			; Obtem o offset da segunda celula
	cmp eax, 0
	jz _error
	lea ebx, offset2
	mov [ebx], eax						; Salva o offset da segunda celula em offset2

	; 0 = empty, 1 = integer, 2 = real, 3 = string, 4 = formulaInt, 5 = formulaReal
	mov eax, offset1
	mov cl, [eax]
	cmp cl, 3							; Erro caso a celula contenha uma string
	jz _error
	cmp cl, 0							; Erro caso a celula não contenha valor
	jz _error
	cmp cl, 1							; Valor inteiro
	jz _intv1
	cmp cl, 4							; Valor inteiro decorrente de funcao
	jz _intv1
	cmp cl, 2							; Valor real
	jz _floatv1
	cmp cl, 5							; Valor real decorrente de funcao
	jz _floatv1
 _intv1:
 	add edx, INTPOS
	fild REAL4 PTR [edx]
	lea ecx, flagval
	mov edx, 1							; Valor eh inteiro
	mov [ecx], edx
 	jmp _continueToV2
 _floatv1:
 	add edx, FLTPOS
	fld REAL4 PTR [edx]
	lea ecx, flagval
	mov edx, 2							; Valor eh real
	mov [ecx], edx
 _continueToV2:
	mov eax, offset2
	mov cl, [eax]
	cmp cl, 3							; Erro caso a celula contenha uma string
	jz _error
	cmp cl, 0							; Erro caso a celula não contenha valor
	jz _error
	cmp cl, 1							; Valor inteiro
	jz _intv2
	cmp cl, 4							; Valor inteiro decorrente de funcao
	jz _intv2
	cmp cl, 2							; Valor real
	jz _floatv2
	cmp cl, 5							; Valor real decorrente de funcao
	jz _floatv2
 _intv2:
 	add edx, INTPOS
	fild REAL4 PTR [edx]
	lea ecx, flagval
	mov edx, 1							; Valor eh inteiro
	mov [ecx], edx
 	jmp _sub
 _floatv2:
 	add edx, FLTPOS
	fld REAL4 PTR [edx]
	lea ecx, flagval
	mov edx, 2							; Valor eh real
	mov [ecx], edx
 _sub:
 	fsub st(0), st(1)
	mov cl, flagval
	cmp cl, 1
	jz _popint
	cmp cl, 2
	jz _popreal
	jmp _error
 _popint:
 	mov edx, offsetCell
	add edx, INTPOS
	fistp DWORD PTR [edx]
	jmp _sucess
 _popreal:
	mov edx, offsetCell
	add edx, FLTPOS
	fstp REAL4 PTR [edx]
 	jmp _sucess
 _error:
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Alguma das celulas selecionadas eh invalida ou possui valores invalidos! Tente novamente."
	jmp _exit
 _sucess:
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
 	invoke GotoPosXY, 1, 27
	mWrite "Copia efetuada com sucesso"
 _exit:
 	fstp trashReal
	ret
Func_SUB ENDP

; Soma o valor de duas celulas e atribui ah celula selecionada
Func_SUM PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offset1: DWORD
	LOCAL offset2: DWORD
	LOCAL flagval: BYTE
	mov edx, StringPTR

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx			; Obtem o offset da primeira celula
	cmp eax, 0
	jz _error
	lea ebx, offset1
	mov [ebx], eax						; Salva o offset da primeira celula em offset1

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx			; Obtem o offset da segunda celula
	cmp eax, 0
	jz _error
	lea ebx, offset2
	mov [ebx], eax						; Salva o offset da segunda celula em offset2

	; 0 = empty, 1 = integer, 2 = real, 3 = string, 4 = formulaInt, 5 = formulaReal
	mov edx, offset1
	mov cl, [eax]
	cmp cl, 3							; Erro caso a celula contenha uma string
	jz _error
	cmp cl, 0							; Erro caso a celula não contenha valor
	jz _error
	cmp cl, 1							; Valor inteiro
	jz _intv1
	cmp cl, 4							; Valor inteiro decorrente de funcao
	jz _intv1
	cmp cl, 2							; Valor real
	jz _floatv1
	cmp cl, 5							; Valor real decorrente de funcao
	jz _floatv1
 _intv1:
 	add edx, INTPOS
	fild REAL4 PTR [edx]
	lea ecx, flagval
	mov dl, 1							; Valor eh inteiro
	mov [ecx], dl
 	jmp _continueToV2
 _floatv1:
 	add edx, FLTPOS
	fld REAL4 PTR [edx]
	lea ecx, flagval
	mov edx, 2							; Valor eh real
	mov [ecx], edx
 _continueToV2:
	mov edx, offset2
	mov cl, [eax]
	cmp cl, 3							; Erro caso a celula contenha uma string
	jz _error
	cmp cl, 0							; Erro caso a celula não contenha valor
	jz _error
	cmp cl, 1							; Valor inteiro
	jz _intv2
	cmp cl, 4							; Valor inteiro decorrente de funcao
	jz _intv2
	cmp cl, 2							; Valor real
	jz _floatv2
	cmp cl, 5							; Valor real decorrente de funcao
	jz _floatv2
 _intv2:
 	add edx, INTPOS
	fild REAL4 PTR [edx]
	lea ecx, flagval
	mov dl, 1							; Valor eh inteiro
	mov [ecx], dl
 	jmp _sum
 _floatv2:
 	add edx, FLTPOS
	fld REAL4 PTR [edx]
	lea ecx, flagval
	mov edx, 2							; Valor eh real
	mov [ecx], edx
 _sum:
 	fadd st(0), st(1)
	mov cl, flagval
	cmp cl, 1
	jz _popint
	cmp cl, 2
	jz _popreal
	jmp _error
 _popint:
 	mov edx, offsetCell
	mov cl, 4
	mov [edx], cl
	add edx, INTPOS
	fistp DWORD PTR [edx]
	sub edx, INTPOS
	add edx, FNCPOS
	invoke Str_copy, edx, addr inputBuffer
	jmp _sucess
 _popreal:
 	mov edx, offsetCell
	mov cl, 5
	mov [edx], cl
	mov edx, offsetCell
	add edx, FLTPOS
	fstp REAL4 PTR [edx]
	sub edx, FLTPOS
	add edx, FNCPOS
	invoke Str_copy, edx, addr inputBuffer
 	jmp _sucess
 _error:
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Alguma das celulas selecionadas eh invalida ou possui valores invalidos! Tente novamente."
	jmp _exit
 _sucess:
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
 	invoke GotoPosXY, 1, 27
	mWrite "Copia efetuada com sucesso"
 _exit:
 	fstp trashReal
	ret
Func_SUM ENDP

; Retorna o offset de uma celula a partir do nome dessa celula.
; Recebe uma DWORD que eh uma string que representa a celula e um espaço em brando. ex: "A01 "
; Retorna em EAX o offset da celula desejada
GetOffsetCell PROC USES EBX ECX EDX, cellAsString: DWORD
	lea edx, cellAsString
	xor eax, eax									; Limpa EAX
	mov al, [edx]									; Move para AL o caractere que representa a coluna
	sub al, 41h										; Subtrai 41h para converter A em decimal, onde A = 0, B = 1...

	cmp al, -1										; Verifica a validade do valor em EAX
	jle _error										; Fora da area de celulas
	cmp al, 6										; Verifica a validade do valor em EAX
	jge _error										; Fora da area de celulas

	mov ebx, SIZECELL								; Move para EBX o tamanho de um struct CELL
	push edx										; Salva o valor de EDX antes da multiplicacao
	mul ebx											; Multiplica EAX pelo tamanho de uma celula, deslocando para a coluna desejada
	pop edx											; Restitui o valor de EDX
	inc edx											; "Aponta" EDX para o trecho da string que representa a linha
	mov ecx, 2										; Atribui a ECX 2, que eh o numero de digitos que representam a linha
	push eax										; Salva o valor de EAX
	call ParseInteger32								; Transforma o numero da string em um inteiro e atribui a EAX
	dec eax											; Ajusta o indice da celula que inicia a contagem em 0

	cmp eax, -1										; Verifica a validade do valor em EAX
	jle _error										; Fora da area de celulas
	cmp eax, 20										; Verifica a validade do valor em EAX
	jge _error										; Fora da area de celulas

	mov ebx, SIZELINE								; Move para EDX o tamanho em bytes de uma linha
	mul ebx											; Multiplica eax gerando o deslocamento para a linha correta
	pop ebx											; Retorna o valor anteriormente em EAX para EBX
	add eax, ebx									; Adiciona o deslocamento de columas ao deslocamento de linhas
	mov edx, offset line1							; Obtem o offset da primeira celula da primeira linha
	add eax, edx									; Adiciona o deslocamento ao offset obtido
	jmp _end
 _error:
 	mov eax, 0
 _end:
	ret
GetOffsetCell ENDP

; Atribui a cor de fundo e da fonte do texto
SetColor PROC USES EAX, color: DWORD
	mov eax, color
	call SetTextColor
	ret
SetColor ENDP

; Gera a base da interface da planila
DrawBase PROC USES EAX ECX EDX
	invoke GoToPosXY, 0, 0							; Move o cursor para o inicio da planilha
	invoke SetColor, (blue + (16 * lightGray))		; Atribui a cor ao texto

	mov edx, offset separator						; Escreve o cabecalho da planilha
	call WriteString
	mov edx, offset titlePlan
	call WriteString
	mov edx, offset separator
	call WriteString

	invoke SetColor, (lightGray + (16 * blue))

	mov edx, offset selectdCell						; Escreve a legenda da area que exibira a celula selecionada
	call WriteString
	call Crlf

	mov edx, offset viewFunc						; Escreve a legenda da area que exibira a funcao presente na celula selecionada
	call WriteString
	call Crlf

	invoke SetColor, (blue + (16 * lightGray))

	mWrite "     "									; Preenche o deslocamento do cabecalho das colunas
	mov edx, offset columnBar						; Marca o cabecalho das colunas
	call WriteString
	mov ecx, 20

 _drawSidebar:										; Preenche o fundo da barra lateral
	mWrite "     "
	call Crlf
	loop _drawSidebar

	mov dh, 6
	mov dl, 1
	mov eax, 1
	mov ecx, 20
 _drawLines:										; Numera as linhas
	call GotoXY
	call WriteDec
	call Crlf
	inc eax
	inc dh
	loop _drawLines

	invoke SetColor, (lightGray + (16 * blue))
	invoke GotoPosXY, 1, 28
	mov edx, offset inputArrow
	call WriteString								; Inicializa a area de entrada de usuario
	call Crlf
	ret
DrawBase ENDP

; Exibe o conteudo das celulas em cada uma das posicoes da tela
DrawData PROC USES EAX EBX ECX EDX ESI
	LOCAL posx: DWORD 
	LOCAL posy: DWORD

	mov dl, isSelectedCell
	cmp dl, 0
	jz _noCellSelected
	invoke GotoPosXY, 8, 4
	mov edx, offsetCell
	add edx, FNCPOS
	call WriteString
	invoke GotoPosXY, 20, 3							; Move o cursor para a posicao 20,3
	mov edx, offset cellName
	call WriteString
 _noCellSelected:
	
	lea edx, posx
	mov eax, 6
	mov [edx], eax									; Atribui a posicao x da primeira celula como 6
	lea edx, posy
	mov [edx], eax									; Atribui a posicao y da primeira celula como 6

	mov edx, offset line1							; Move para edx o offset da primeira celula
	mov ecx, 120									; Move para ecx o numero de celulas
 _draw:
 	mov eax, posx
	mov ebx, posy
	invoke GotoPosXY, eax, ebx						; Move o cursor para a posicao xy da celula atual

	; Inicio da verificacao do tipo de dado contido na celula
	; 0 = empty, 1 = integer, 2 = real, 3 = string, 4 = formulaInt, 5 = formulaReal
	; 4 e 5 sao, respectivamente, inteiros ou reais decorrentes da funcao presente na celula 
	; (Diferenca necessaria para a funcao de atualizacao de valores)
 	mov al, [edx]
	cmp al, 0
	jz _nextCell
	cmp al, 1						; Inteiro
	jz _hasInt
	cmp al, 2						; Real
	jz _hasReal
	cmp al, 3						; String
	jz _hasString
	cmp al, 4						; Inteiro de funcao
	jz _hasInt
	cmp al, 5						; Real de funcao
	jz _hasReal
	mWrite "                  "
	jmp _nextCell
 _hasInt:							; Exibe o inteiro presente na celula
	mov esi, INTPOS					; Para isso atribui ao ESI o deslocamento do inicio da celula ateh o valor inteiro
	mov eax, [edx + esi]	
	call WriteInt
	jmp _nextCell					; Salta para a proxima celula
 _hasReal:
 	mov esi, FLTPOS
	fld REAL4 ptr [edx + esi]
	call WriteFloat
	fstp trashReal
 	jmp _nextCell
 _hasString:						; Exibe a string presente na celula
 	add edx, STRPOS					; Para isso adiciona EDX o deslocamento do inicio da celula ateh a string
	call WriteString
	sub edx, STRPOS					; Retorna edx para o valor anterior
 _nextCell:
 	mov eax, posx
	mov ebx, posy
	cmp eax, 101					; Verifica se estah na ultima celula da linha
	jnz _keepLine
 _changeLine:						; Se estiver:
	mov eax, 6						; Move x para o inicio da linha
	inc ebx							; Salta para a proxima linha
	jmp _continue
 _keepLine:
 	add eax, 19						; Senao, salta para a proxima celula na linha atual
 _continue:
	lea esi, posx					; Atualiza a variavel posx com a nova posicao
	mov [esi], eax
	lea esi, posy					; Atualiza a variavel posy com a nova posicao
	mov [esi], ebx
	add edx, SIZECELL
 	loop _gotoDraw
	jmp _end
 _gotoDraw:
	jmp _draw
 _end:
	ret
DrawData ENDP

; Move o cursor para uma determinada posicao, funcao gotoxy facilitada pelo uso de invoke
GotoPosXY PROC USES EDX, posx: DWORD, posy: DWORD
	mov dl, BYTE PTR posX
	mov dh, BYTE PTR posY
	call Gotoxy
	ret
GotoPosXY ENDP

; Limpa o conteudo da planilha sem apagar a base da planilha
ClearSheet	PROC USES EAX EBX ECX EDX
	invoke SetColor, (lightGray + (16 * blue))		; Atribui a cor de fundo
	mov eax, 6										; Posicao inicial em X
	mov ebx, 6										; Posicao inicial em Y
	mov ecx, 20										; 20 linhas
 _clearLines:
	push ecx										; Salva o ecx para linhas
	mov ecx, 6										; 6 columas
 _clearColumns:
	push ecx										; Salva o ecx para columas
	invoke GotoPosXY, eax, ebx						; Move o cursor para o inicio da celula a ser limpa
	mWrite "                  "						; Limpa cada uma das celulas da linha, preenchendo com espacos
	add eax, 19										; Mve incrementa 19 na posicao do cursor para a proxima celula
	pop ecx											; restitui o valor de ecx para colunas
	loop _clearColumns

	mov eax, 6										; volta eax para o inicio da linha
	add ebx, 1										; move ebx para a proxima linha
	pop ecx											; restitui o valor de ecx para linhas
	loop _clearLines
	ret
ClearSheet ENDP

END main
