COMMENT #
	Planilha desenvolvida em Assembly na disciplina v.6
	Laboratório de Arquitetura e Organizacao de Computadores 2
    	Autor: Leandro Novak

	Funcao:			Estado atual:
	CEL			Finalizada
	CLR			Finalizada
	CLT			Finalizada
	COP			Finalizada
	CUT			Finalizada
	EXT			Finalizada
	FLT			Finalizada
	HLP			Finalizada
	INT			Finalizada
	MAX			Finalizada
	MED			Finalizada
	MIN			Finalizada
	OPN			Finalizada
	SAV			Finalizada
	STR			Finalizada
	SUB			Finalizada
	SUM			Finalizada

	As demais funcoes presentes na planilha sao de uso interno da planilha, cuja utilizacao nao é solicitada pelo usuario, mas ocorre de forma automatica quando necessario.

	Obs. Aplicacao feita para janelas com tamanho 120 x 30 (Windows 10). 
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

; Variavel para verificacao de menor valor
infinity		DWORD 7F800000h
infinityneg		DWORD 0FF800000h

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
functions		BYTE "CEL CLR", 00h, "CLT", 00h, "COP CUT EXT", 00h, "FLT", 00h, "HLP", 00h, "INT", 00h, "MAX MED MIN OPN", 00h, "SAV", 00h, "STR", 00H, "SUB SUM ", 00h, 0
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
Func_FLT		PROTO
Func_INT 		PROTO
Func_MAX 		PROTO StringPTR: PTR DWORD
Func_MED 		PROTO StringPTR: PTR DWORD
Func_MIN 		PROTO StringPTR: PTR DWORD
Func_OPN 		PROTO
Func_SAV 		PROTO
Func_STR 		PROTO
Func_SUB 		PROTO StringPTR: PTR DWORD
Func_SUM 		PROTO StringPTR: PTR DWORD

; Funcoes para atualizacao da planilha
UpdatePlan		PROTO
ProcessUpdate	PROTO stringPtr: PTR DWORD, actualCell: PTR DWORD
Update_MAX		PROTO stringPTR: PTR DWORD, actualCell: PTR DWORD
Update_MED		PROTO stringPTR: PTR DWORD, actualCell: PTR DWORD
Update_MIN		PROTO stringPTR: PTR DWORD, actualCell: PTR DWORD
Update_SUB		PROTO stringPTR: PTR DWORD, actualCell: PTR DWORD
Update_SUM		PROTO stringPTR: PTR DWORD, actualCell: PTR DWORD

; Funcoes gerais
ClearSheet		PROTO
DrawBase		PROTO
DrawData		PROTO
GetOffsetCell	PROTO cellAsString: DWORD
GotoPosXY		PROTO posx: DWORD, posy: DWORD
SetColor		PROTO color: DWORD

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Funcao main
; ---------------------------------------------------------------------------------------------------------------------------------------------------
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

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Le o comando digitado pelo usuario
; Recebe stringPtr que eh um ponteiro para a variavel que armazenara a string digitada pelo usuario.
; ---------------------------------------------------------------------------------------------------------------------------------------------------
ReadInput PROC USES EDX, stringPtr: PTR DWORD
	mov ecx, 18										; Trecho que prepara a interface para a leitura da entrada de usuario, limpando e desenhando certos itens.
	invoke SetColor, (lightGray + (16 * blue))
	invoke GoToPosXY, 0, 28
	mov edx, offset separator
	call WriteString
	invoke GoToPosXY, 1, 28
	mov edx, offset inputArrow
	call WriteString
	invoke GoToPosXY, 4, 28

	mov edx, stringPtr
	call ReadString									; Le a entrada do usuario como uma string.
	invoke Str_copy, stringPtr, addr tempBuffer		; Salva uma copia da entrada digitada.
	invoke Str_ucase, stringPtr						; Deixa a entrada toda em letras maiusculas.
	invoke ProcessString, stringPtr					; Chama a funcao que processara a entrada e definira qual a funcao escolhida.
	call Crlf
	ret
ReadInput ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Processa a entrada do usuario
; Funcao chamada automaticamente
; Recebe stringPtr que eh um ponteiro para a string digitada pelo usuario
; ---------------------------------------------------------------------------------------------------------------------------------------------------
ProcessString PROC USES EAX EBX EDX, stringPtr: PTR DWORD
	mov edx, stringPtr
	mov eax, [edx]

	mov edx, offset functions						; Neste trecho eh feita a verificacao da parte inicial entrada fornecida pelo usuario
	cmp eax, [edx]									; se for CEL
	jnz _clr
	invoke Func_CEL, stringPtr
	jmp _end
 _clr:
	add edx, 4
	cmp eax, [edx]									; CLR
	jnz _clt
	invoke Func_CLR
	jmp _end
 _clt:
	add edx, 4
	cmp eax, [edx]									; CLT
	jnz _cop
	invoke Func_CLT
	jmp _end
 _cop:
	add edx, 4
	cmp eax, [edx]									; COP
	jnz _cut
	invoke Func_COP, stringPtr
	jmp _end
 _cut:
	add edx, 4
	cmp eax, [edx]									; CUT
	jnz _ext
	invoke Func_CUT, stringPtr
	jmp _end
 _ext:
	add edx, 4
	cmp eax, [edx]									; EXT
	jnz _flt
	invoke Func_EXT
	jmp _end
 _flt:
 	add edx, 4
	cmp eax, [edx]									; FLT
	jnz _hlp
	invoke Func_FLT
	jmp _end
 _hlp:
	add edx, 4
	cmp eax, [edx]									; HLP
	jnz _int
	invoke Func_HLP
	jmp _end
 _int:
	add edx, 4
	cmp eax, [edx]									; INT
	jnz _max
	invoke Func_INT
	jmp _end
 _max:
	add edx, 4
	cmp eax, [edx]									; MAX
	jnz _med
	invoke Func_MAX, StringPTR
	jmp _end
 _med:
	add edx, 4
	cmp eax, [edx]									; MED
	jnz _min
	invoke Func_MED, StringPTR
	jmp _end
 _min:
	add edx, 4
	cmp eax, [edx]									; MIN
	jnz _opn
	invoke Func_MIN, StringPTR
	jmp _end
 _opn:
	add edx, 4
	cmp eax, [edx]									; OPN
	jnz _sav
	invoke Func_OPN
	jmp _end
 _sav:
	add edx, 4
	cmp eax, [edx]									; SAV
	jnz _str
	invoke Func_SAV
	jmp _end
 _str:
	add edx, 4
	cmp eax, [edx]									; STR
	jnz _sub
	invoke Func_STR
	jmp _end
 _sub:
	add edx, 4
	cmp eax, [edx]									; SUB
	jnz _sum
	invoke Func_SUB, StringPTR
	jmp _end
 _sum:
	add edx, 4
	cmp eax, [edx]									; SUM
	jnz _error
	invoke Func_SUM, StringPTR
	jmp _end
 _error:											; Caso onde a entrada eh invalida
	invoke SetColor, (lightGray + (16 * blue))
	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString								; Limpa a area de feedback
	invoke GotoPosXY, 1, 27
	mWrite "OPERACAO INVALIDA!"						; Exibe uma mensagem de erro

 _end:
	ret
ProcessString ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Seleciona a celula digitada salvando em offsetCell o offset da celula desejada
; Recebe stringPtr que eh um ponteiro para a string digitada pelo usuario
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_CEL PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
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

	mov edx, stringPTR								; "Aponta" EDX para o inicio da area da string que representa a celula
	add edx, 4
	mov ecx, [edx]
	mov edx, offset cellName						; Salva o nome da celula selecionada
	mov [edx], ecx

	mov edx, offset offsetCell
	mov [edx], eax									; Salva o offset da celula selecionada em offsetCell

	mov edx, offset isSelectedCell					; Seta o "FLAG" de celula selecionada
	mov al, 1
	mov [edx], al

	invoke GotoPosXY, 1, 27
	mWrite "CELULA SELECIONADA COM SUCESSO!"		; Mensagem de sucesso
	jmp _end
 _error:
	invoke GotoPosXY, 1, 27
	mWrite "ERRO AO SELECIONAR CELULA!"				; Mensagem de erro
 _end:
	invoke GotoPosXY, 7, 4
	mWrite "                               "		; Limpa a area que exibe a funcao da celula selecionada
	ret
Func_CEL ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Limpa todo o conteudo de uma celula
; Nao recebe parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_CLR PROC USES EAX EBX ECX EDX
	invoke SetColor, (lightGray + (16 * blue))		; Limpa a area de feedback 
	invoke GoToPosXY, 0, 26
	mov edx, offset separator
	call WriteString
	invoke GoToPosXY, 0, 27
	call WriteString

	xor eax, eax									; Limpa o conteudo de eax
	mov al, isSelectedCell							; Move para al a flag de celula selecionada
	cmp eax, 0										; Verifica se ha alguma celula selecionada (flag diferente de 0)
	jz _error
	mov ebx, 0										; Move para ebx 0, que eh o valor presente em celulas limpas
	mov ecx, 10										; Move para ecx 10, que eh o tamanho de uma celula em dwords (40 bytes)
	mov eax, offsetCell								; Move para eax o offset da celula
 _clear:											; Executa a limpeza de 4 em 4 bytes
	mov [eax], ebx
	add eax, 4
	loop _clear
	jmp _sucess
 _error:
	invoke GotoPosXY, 1, 27
	mWrite "SELECIONE UMA CELULA PARA LIMPAR!"		; Mensagem de erro caso nao haja uma celula selecionada
	jmp _end
 _sucess:
	invoke GotoPosXY, 1, 27
	mWrite "CELULA LIMPA COM SUCESSO!"				; Mensagem confirmando que a celula foi limpa
 _end:
	ret
Func_CLR ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Limpa todo o conteudo da planilha
; Nao recebe parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_CLT PROC USES EAX ECX EDX
	invoke GotoPosXY, 1, 26							; O trecho abaixo exibe uma mensagem pedindo a confirmacao de limpeza das celulas.
	mWrite "DESEJA REALMENTE LIMPAR TODA A PLANILHA? ESTA OPERACAO NAO PODE SER DESFEITA."		
	invoke GotoPosXY, 1, 27
	mWrite "Pressione 'S' para continuar ou 'N' para cancelar: "								
	call ReadChar									; Recebe a resposta do usuario
	cmp al, 'n'
	jz _cancel
	cmp al, 'N'
	jz _cancel
	cmp al, 's'
	jz _continue
	cmp al, 'S'
	jz _continue
 _error:											; Caso o usuario digite uma opcao invalida
	mov edx, offset separator
	invoke GotoPosXY, 0, 26
	call WriteString
	invoke GotoPosXY, 0, 27
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "OPCAO DIGITADA INVALIDA! A planilha nao sera limpa."
	jmp _end
 _cancel:											; Caso o usuario cancele a limpeza
 	invoke GotoPosXY, 0, 26
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 0, 27
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "OPERACAO CANCELADA. A PLANILHA NAO FOI LIMPA."
	jmp _end
 _continue:											; Limpa as celulas de 4 em 4 bytes (semelhante ao processo de limpeza de uma unica celula)
	xor eax, eax
	mov edx, offset line1							; Inicio da area a ser limpa
	mov ecx, 1410									; Tamanho total das celulas em dwords
	mov eax, 0										; Valor padrao para as celulas
 _clear:
	mov [edx], eax
	add edx, 4
	loop _clear
 _sucess:											; Mensagem informando que as celulas foram limpas
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

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Copia o conteudo da celula origem para a celula destino.
; Recebe stringPtr que eh um ponteiro para a string digitada pelo usuario
; Ex. COP A10 A20 copia o conteudo da celula A10 para a celula A20
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_COP PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offset1: DWORD
	LOCAL offset2: DWORD
	mov edx, StringPTR								; Move para edx o endereco da string que contem a entrada informada pelo usuario

	add edx, 4										; Faz edx apontar para o nome da primeira celula na string
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da primeira celula
	cmp eax, 0										; Salta se nao for possivel obter o offset
	jz _error
	lea ebx, offset1
	mov [ebx], eax  								; Salva o offset da primeira celula em offset1

	add edx, 4										; Faz edx apontar para o nome da segunda celula na string
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da primeira celula
	cmp eax, 0										; Salta se nao for possivel obter o offset
	jz _error
	lea ebx, offset2
	mov [ebx], eax  								; Salva o offset da segunda celula em offset2

	mov eax, offset1								; Move para eax o offset da primeira celula
	mov ebx, offset2								; Move para ebx o offset da segunda celula
	mov ecx, 10										; Move para edx o tamanho do dado a ser movido (em dwords)
 _move:
	mov edx, [eax]									; Move 4 bytes da primeira celula para edx
	mov [ebx], edx									; Move edx para a segunda celula
	add eax, 4										; Avanca o endereco das duas celulas
	add ebx, 4
	loop _move
	jmp _sucess										; Salta ao fim da copia
 _error:											; Exibe uma mensagem de erro caso não seja possivel copiar uma das celulas
	invoke GotoPosXY, 1, 27
	mWrite "Alguma das celulas selecionadas eh invalida! Tente novamente."
	jmp _exit
 _sucess:											; Exibe uma mensagem informando que a celula foi copiada com sucesso
 	invoke GotoPosXY, 1, 27
	mWrite "Copia efetuada com sucesso"
 _exit:
 	invoke UpdatePlan								; Chama a funcao de atualizacao da planilha (necessaria para casos onde ha alguma celula cujo valor dependa da celula destino).
	ret
Func_COP ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Move o conteudo da celula origem para a celula destino.
; Recebe stringPtr que eh um ponteiro para a string digitada pelo usuario
; Ex. CUT A10 A20 copia o conteudo da celula A10 para a celula A20 e apaga o conteudo de A10
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_CUT PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offset1: DWORD
	LOCAL offset2: DWORD
	mov edx, StringPTR								; Move para edx o endereco da string que contem a entrada informada pelo usuario

	add edx, 4										; Faz edx apontar para o nome da primeira celula na string
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da primeira celula
	cmp eax, 0										; Salta se nao for possivel obter o offset
	jz _error
	lea ebx, offset1
	mov [ebx], eax  								; Salva o offset da primeira celula em offset1

	add edx, 4										; Faz edx apontar para o nome da segunda celula na string
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da segunda celula
	cmp eax, 0										; Salta se nao for possivel obter o offset
	jz _error
	lea ebx, offset2
	mov [ebx], eax  								; Salva o offset da segunda celula em offset2

	mov eax, offset1								; Move para eax o offset da primeira celula
	mov ebx, offset2								; Move para eax o offset da segunda celula
	mov ecx, 10
 _move:
	mov edx, [eax]									; Move 4 bytes da primeira celula para edx
	mov [ebx], edx									; Move edx para a segunda celula
	add eax, 4										; Avanca o endereco das duas celulas
	add ebx, 4
	loop _move
	jmp _sucess										; Salta ao fim da copia (Recorte eh uma copia seguida da limpeza da celula origem)
 _error:											; Exibe uma mensagem de erro caso não seja possivel copiar uma das celulas
	invoke GotoPosXY, 1, 27	
	mWrite "Alguma das celulas selecionadas eh invalida! Tente novamente."
	jmp _exit
 _sucess:											; Exibe uma mensagem informando que a celula foi copiada com sucesso
 	invoke GotoPosXY, 1, 27
	mWrite "Recorte efetuado com sucesso"

	mov ebx, 0										; Move para ebx 0, que eh o valor padrao das celulas
	mov ecx, 10										; Move para ecx 10, que eh o tamanho de uma celula em dwords (40 bytes)
	mov eax, offset1								; Move para eax o offset da primeira celula
 _clear:											; Efetua a limpeza da celula de 4 em 4 bytes
	mov [eax], ebx
	add eax, 4
	loop _clear
 _exit:
	invoke UpdatePlan								; Chama a funcao de atualizacao da planilha (necessaria para casos onde ha alguma celula cujo valor dependa da celula origem).
	ret
Func_CUT ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Encerra a execucao da planilha
; Nao recebe parametro
; Retorna ECX = 0 para sair ou ECX = 1 para continuar na planilha
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_EXT PROC USES EAX
	invoke GotoPosXY, 1, 26							; Pede para o usuario confirmar se deseja mesmo fechar a planilha
	mWrite "DESEJA REALMENTE FECHAR A PLANILHA?"
	invoke GotoPosXY, 1, 27
	mWrite "Pressione 'S' para sair ou 'N' para continuar na planilha: "
	call ReadChar									; Recebe a resposta do usuario
	cmp al, 'n'
	jz _cancel
	cmp al, 'N'
	jz _cancel
	cmp al, 's'
	jz _continue
	cmp al, 'S'
	jz _continue
 _cancel:											; Move para ecx, 1, em caso de cancelamento ou opcao invalida
	mov ecx, 1
	jmp _exit										; Encerra a funcao de fechamento
 _continue:
 	mov ecx, 0										; Move para ecx, 0, caso a planilha seja fechada
 _exit:												; Limpa a area de feedback independente do resultado da funcao
 	invoke SetColor, (lightGray + (16 * blue))
	invoke GoToPosXY, 0, 26
	mov edx, offset separator
	call WriteString
	invoke GoToPosXY, 0, 27
	call WriteString
	ret
Func_EXT ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Atribui um valor real para a celula previamente selecionada
; Nao recebe nenhum parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_FLT PROC USES EAX EBX ECX EDX
	xor ecx, ecx
	mov cl, isSelectedCell							; Verifica se ha uma celula selecionada
	jecxz _error									; Exibe uma mensagem de erro caso nao haja

	invoke GotoPosXY, 1, 27							; Limpa a area onde sera solicitada a entrada do numero real
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27							; Solicita que o usuario digite um numero real
	mWrite "Digite o numero real desejado: "

	mov edx, offsetCell								; Move para edx o offset da celula selecionada
	mov cl, 2										
	mov [edx], cl									; Modifica a flag de tipo da celula para 2 (Float)
	add edx, FLTPOS									; Move edx para a posicao do valor float na celula
	call ReadFloat
	fstp REAL4 PTR [edx]							; Retira o valor float da pilha e armazena na celula
							
	invoke GoToPosXY, 0, 27
	mov edx, offset separator						; Exibe uma mensagem informando que o valor foi atribuido com sucesso
	call WriteString
 	invoke GotoPosXY, 1, 27
	mWrite "Atribuicao de real efetuada com sucesso!"
	jmp _end										; Salta para o fim da funcao
 _error:
 	invoke GoToPosXY, 0, 27							; Exibe uma mensagem de erro caso nao haja nenhuma celula selecionada
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "SELECIONE UMA CELULA PRIMEIRO!"
	jmp _end

 _end:
 	invoke UpdatePlan								; Chama a funcao de atualizacao da planilha (necessaria para casos onde ha alguma celula cujo valor dependa da celula modificada).
	ret
Func_FLT ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Exibe a tela de ajuda para o usuario
; Nao recebe nenhum parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_HLP PROC USES EAX EBX ECX EDX
	invoke SetColor, (lightGray + (16 * blue))		; Este trecho o cabecalho da tela contendo o titulo "AJUDA"
	call Clrscr
	invoke gotoPosXY, 0, 0
	invoke SetColor, (blue + (16 * lightGray))
	mov edx, offset separator
	call WriteString
	call WriteString
	call WriteString
	invoke gotoPosXY, 58, 1
	mWrite "AJUDA"

	invoke SetColor, (lightGray + (16 * blue))		; Este trecho exibe uma breve descricao das funcoes presentes na planilha, bem como uma legenda para facilitar a leitura.
	call Crlf
	call Crlf
	call Crlf
	call Crlf
	mWrite "   CEL - Seleciona uma celula da planilha. CEL A01"
	call Crlf
	mWrite "   CLR - Limpa o conteudo da celula previamente selecionada com a funcao CEL. Sem parametros."
	call Crlf
	mWrite "   CLT - Limpa o conteudo de todas as celulas da tabela. Sem parametros."
	call Crlf
	mWrite "   COP - Copia todo o conteudo de uma celula para outra. COP ORI DES"
	call Crlf
	mWrite "   CUT - Move o conteudo de uma celula para outra. COP ORI DES"
	call Crlf
	mWrite "   EXT - Encerra a execucao da planilha."
	call Crlf
	mWrite "   FLT - Funcao para atribuir um valor real a uma celula. Sem parametros, valor digitado posteriormente."
	call Crlf
	mWrite "   INT - Funcao para atribuir um valor inteiro a uma celula. Sem parametros, valor digitado posteriormente."
	call Crlf
	mWrite "   MAX - Salva o maior valor entre duas ou mais celulas na celula selecionada. MAX CL1 CL2. Ordem insignificante."
	call Crlf
	mWrite "   MED - Calcula a media de valores das celulas e salva na celula selecionada. MED CL1 CL2. Ordem insignificante."
	call Crlf
	mWrite "   MIN - Salva o menor valor de duas ou mais celulas e na na celula selecionada. MIN CL1 CL2. Ordem insignificante."
	call Crlf
	mWrite "   OPN - Abre uma planilha salva no disco rigido. Sem parametros, nome do arquivo digitado posteriormente."
	call Crlf
	mWrite "   SAV - Salva a planilha atual no disco rigido. Sem parametros, nome do arquivo digitado posteriormente."
	call Crlf
	mWrite "   STR - Funcao para atribuir um texto a uma celula. Sem parametros, texto digitado posteriormente."
	call Crlf
	mWrite "   SUB - Subtrai o valor de uma celula de outra. SUB CL1 CL2 (CL1 - CL2)."
	call Crlf
	mWrite "   SUM - Soma o valor de uma celula com outra. SUM CL1 CL2 (CL1 + CL2)."
	call Crlf
	call Crlf
	call Crlf
	mWrite " Legenda: ORI = ORIGEM, DES = DESTINO, CL1 = Celula como primeiro parametro, CL2 = Celula como segundo parametro."
	call Crlf
	mWrite " Para mais detalhes leia o arquivo TXT que se encontra no mesmo diretorio do executavel da planilha."
 _input:
	invoke GotoPosXY, 1, 27							; Pede que o usuario pressione a tecla enter caso deseje retornar para a planilha
	mWrite "Pressione a tecla enter para retornar a planilha: "
	xor eax, eax
	call ReadChar
	cmp al, 0Dh
	jz _exit										; Salta para o fim da funcao
	invoke GotoPosXY, 1, 26						
	mWrite "A tecla pressionada eh invalida."		; Mensagem de erro caso a tecla seja invalida. Pede para pressionar novamente.
	jmp _input
 _exit:
	invoke SetColor, (lightGray + (16 * blue))	
	call Clrscr										; Limpa a tela
	call DrawBase									; Desenha a base da planilha
	ret
Func_HLP ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Atribui um valor inteiro com sinal para a celula previamente selecionada
; Nao recebe nenhum parametro
; Exige que a celula de destino tenha sido previamente selecionada
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_INT PROC USES EAX EBX ECX EDX
	xor ecx, ecx
	mov cl, isSelectedCell							; Verifica se ha alguma celula selecionada
	jecxz _error									; Salta para _error caso não haja

	invoke GotoPosXY, 1, 27							; Este trecho exibe uma mensagem pedindo para o usuario digitar o valor inteiro desejado.
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Digite o numero inteiro desejado: "
	call ReadInt									; Le o inteiro digitado pelo usuario

	mov edx, offsetCell								; Move para edx o offset da celula
	mov cl, 1
	mov [edx], cl									; Muda o tipo valor da celula para inteiro
	add edx, INTPOS									; Desloca edx para a area de inteiros da celula
	mov [edx], eax									; Salva o inteiro digitado na celula
	jmp _sucess										; Salta para area que exibe a mensagem de sucesso

 _error:											; Exibe uma mensagem de erro caso nao haja nenhuma celula selecionada
  	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "SELECIONE UMA CELULA PRIMEIRO!"
	jmp _end
 _sucess:											; Exibe uma mensagem informando que o numero foi atribuido com sucesso
  	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString
 	invoke GotoPosXY, 1, 27
	mWrite "Atribuicao de inteiro efetuada com sucesso!"
 _end:
 	invoke UpdatePlan								; Chama a funcao de atualizacao da planilha (necessaria para casos onde ha alguma celula cujo valor dependa da celula modificada).
	ret
Func_INT ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Calcula o maior valor dentre um conjunto de celulas
; Recebe stringPtr que eh um ponteiro para a string digitada pelo usuario
; A string deve conter a primeira e a ultima celula do intervalo (mesma linha ou mesma coluna/ vizinhas ou nao)
; Exige que a celula de destino tenha sido previamente selecionada
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_MAX PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offsetBegin: DWORD
	LOCAL offsetEnd: DWORD
	LOCAL flagval: BYTE
	LOCAL max: REAL4

	mov bl, isSelectedCell							; Verifica se ha alguma celula selecionada
	cmp bl, 0
	jnz _cellSelected								; Salta para _error caso não haja
	jmp _error
 _cellSelected:

	lea edx, flagval
	mov ebx, 0
	mov [edx], ebx

	fild DWORD PTR infinityneg
	fstp max										; Atribui o maior valor como -infinito

	lea edx, flagval
	mov ebx, 1
	mov [edx], ebx									; Seta a flag de tipo como 1

	mov edx, StringPTR								; Obtem o offset da entrada
	add edx, 4	
	mov eax, [edx]									; Obtem o nome da primera celula
	and eax, 00FFFFFFh								; Limpa o byte mais significativo
	add edx, 4
 _isTheSameColumn:	
	mov ebx, [edx]									; Obtem o nome da segunda celula
	cmp al, bl										; Verifica se a coluna da primeira celula eh igual a da segunda
	jz _sameColumn									; Salta se for a mesma coluna
	jmp _differentColumn							; Salta se a coluna for diferente (espera-se que seja mesma linha)

 _sameColumn:
 	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8
	cmp eax, ebx									; Compara as linhas para verificar se a segunda eh maior que a primeira
	jb _scSecondGreater								; Se a segunda for maior que a primeira
 _scFirstGreater:
	xchg eax, ebx									; Corrige caso a primeira seja maior q a segunda, invertendo as celulas
 _scSecondGreater:
	rol eax, 8										; Rotaciona para a forma inicial
	rol ebx, 8
	invoke getOffsetCell, eax						; Obtem o offset da celula inicial
	xchg eax, ebx									; Troca para salvar os valor
	invoke getOffsetCell, eax						; Obtem o offset da celula final
	lea edx, offsetEnd
	mov [edx], eax									; Salva a celula final em offsetEnd
	lea edx, offsetBegin
	mov [edx], ebx									; Salva a celula inicial em offsetBegin
	mov edx, offsetBegin
	mov ecx, 0
 _scdoingComparison:
	push edx										; Salva o offset atual
	fld max
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _scjmpToError
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _scjmpToError
	cmp cl, 1										; Valor inteiro
	jz _scvalInt
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _scvalInt
	cmp cl, 2										; Valor real
	jz _scvalFloat	
	cmp cl, 5										; Valor real decorrente de funcao
	jz _scvalFloat
 _scjmpToError:
	jmp _error
 _scvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	jb _continueInColumn
	fild DWORD PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp max
	jmp _continueInColumn
 _scvalFloat:
	push edx
 	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2
	mov [edx], bl									; Seta a flag de tipo com 2 (float)
	pop edx
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	jb _continueInColumn	
	fld REAL4 PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp max
 _continueInColumn:
	fstp trashreal
	pop edx											; Retorna para edx o offset atual
	add edx, SIZELINE								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _scjmpEndMax									; Se ultrapassou, todos os valores foram verificados
	jmp _scdoingComparison							; Senao, continua verificando

 _scjmpEndMax:
	jmp _endMax

 _differentColumn:
	cmp al, bl
	ja _dcFirstGreater
 _dcSecondGreater:
	xchg eax, ebx
 _dcFirstGreater:
  	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8
	cmp ax, bx
	jnz _dcjmpToError
	rol eax, 8
	rol ebx, 8
	invoke getOffsetCell, eax						; Obtem o offset da ultima celula
	lea edx, offsetEnd
	mov [edx], eax
	invoke getOffsetCell, ebx						; Obtem o offset da primeira celula
	lea edx, offsetBegin
	mov [edx], eax
	mov edx, offsetBegin							; Salva em edx o offset da primeira celula
	mov ecx, 0
 _dcdoingComparison:
	push edx										; Salva o offset atual
	fld max
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _dcjmpToError
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _dcjmpToError
	cmp cl, 1										; Valor inteiro
	jz _dcvalInt
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _dcvalInt
	cmp cl, 2										; Valor real
	jz _dcvalFloat	
	cmp cl, 5										; Valor real decorrente de funcao
	jz _dcvalFloat
 _dcjmpToError:
	jmp _error
 _dcvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	jb _continueInLine
	fild DWORD PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp max
	jmp _continueInLine
 _dcvalFloat:
	push edx
 	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2	
	mov [edx], bl									; Seta a flag de tipo com 2 (float)
	pop edx
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	jb _continueInLine
	fld REAL4 PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp max
 _continueInLine:
	fstp trashReal
	pop edx											; Retorna para edx o offset atual
	add edx, SIZECELL								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _endMax										; Se ultrapassou, todos as celulas foram verificadas
	jmp _dcdoingComparison							; Senao, continua verificando

 _endMax:
	mov edx, offsetCell								; Salva em edx o offset da celula destino
	mov bl, flagVal									; Move para bl o valor da flag de tipo
	cmp bl, 1
	jz _valint										; Salta se o resultado for decorrente apenas de inteiros
	cmp bl, 2
	jz _valfloat									; Salta se algum dos numeros for float
	jmp _error
 _valint:
	mov al, 4										; Move 4 para al para indicar o tipo de dado contido na celula
	mov [edx], al									; Salva 4 na celula
	add edx, INTPOS									; Move edx para a posicao inteira
	fld REAL4 PTR max
	fistp DWORD PTR[edx]							; Salva o resultado na celula
	add edx, FNCPOS - INTPOS						; Move edx para o inicio da area que contem a funcao
	invoke Str_copy, addr inputBuffer, edx			; Copia a funcao para a celula
	jmp _sucess
 _valfloat:
	mov al, 5										; Move 5 para al para indicar o tipo de dado contido na celula
	mov [edx], al	
	add edx, FLTPOS									; Move edx para a posicao float
	fld REAL4 PTR max
	fstp REAL4 PTR[edx]								; Salva o resultado na celula
	add edx, FNCPOS - FLTPOS						; Move edx para o inicio da area que contem a funcao
	invoke Str_copy, addr inputBuffer, edx			; Copia a funcao para a celula
	jmp _sucess
 _error:											; Exibe uma mensagem de erro
	pop eax
	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27
	mWrite "Nao foi possivel obter o menor valor."
	jmp _end
 _sucess:											; Exibe uma mensagem informando que a operacao foi concluida com sucesso
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27
	mWrite "Maior valor obtido com sucesso."
 _end:
 	invoke UpdatePlan
	ret
Func_MAX ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Calcula a media dos valores presentes em um conjunto de celulas.
; Recebe stringPtr que eh um ponteiro para a string digitada pelo usuario
; A string deve conter a primeira e a ultima celula do intervalo (mesma linha ou mesma coluna/ vizinhas ou nao)
; Exige que a celula de destino tenha sido previamente selecionada
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_MED PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offsetBegin: DWORD
	LOCAL offsetEnd: DWORD
	LOCAL flagval: BYTE

	mov bl, isSelectedCell							; Verifica se ha alguma celula selecionada
	cmp bl, 0
	jnz _cellSelected							; Salta para _error caso não haja
	jmp _error
 _cellSelected:

	lea edx, flagVal
	mov ebx, 0
	mov [edx], ebx									; Seta a flag de tipo como 1
	fild DWORD PTR flagVal							; Inicializa a pilha com 0

	lea edx, flagVal
	mov ebx, 1
	mov [edx], ebx									; Seta a flag de tipo como 1

	mov edx, StringPTR								; Obtem o offset da entrada
	add edx, 4
	mov eax, [edx]									; Obtem o nome da primera celula
	and eax, 00FFFFFFh								; Limpa o byte mais significativo
	add edx, 4
 _isTheSameColumn:
	mov ebx, [edx]									; Obtem o nome da segunda celula
	cmp al, bl										; Verifica se a coluna da primeira celula eh igual a da segunda
	jz _sameColumn									; Salta se for a mesma coluna
	jmp _differentColumn							; Salta se a coluna for diferente (espera-se que seja mesma linha)

 _sameColumn:
 	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8
	cmp eax, ebx									; Compara as linhas para verificar se a segunda eh maior que a primeira
	jb _scSecondGreater								; Se a segunda for maior que a primeira
 _scFirstGreater:
	xchg eax, ebx									; Corrige caso a primeira seja maior q a segunda, invertendo as celulas
 _scSecondGreater:
	rol eax, 8										; Rotaciona para a forma inicial
	rol ebx, 8
	invoke getOffsetCell, eax						; Obtem o offset da celula inicial
	xchg eax, ebx									; Troca para salvar os valor
	invoke getOffsetCell, eax						; Obtem o offset da celula final
	lea edx, offsetEnd
	mov [edx], eax									; Salva a celula final em offsetEnd
	lea edx, offsetBegin
	mov [edx], ebx									; Salva a celula inicial em offsetBegin
	mov edx, offsetBegin
	mov ecx, 0
 _scmakingSum:
	push edx										; Salva o offset atual
	push ecx
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _scjmpToError
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _scjmpToError
	cmp cl, 1										; Valor inteiro
	jz _scvalInt
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _scvalInt
	cmp cl, 2										; Valor real
	jz _scvalFloat
	cmp cl, 5										; Valor real decorrente de funcao
	jz _scvalFloat
 _scjmpToError:
	jmp _error
 _scvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fadd ST(1), ST(0)								; Soma o valor atual com a soma anterior
	fstp trashReal									; Deixa so um valor na pilha
	jmp _continueInColumn
 _scvalFloat:
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fadd ST(1), ST(0)								; Soma o valor atual com a soma anterior
	fstp trashReal									; Deixa so um valor na pilha
	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2
	mov [edx], ebx									; Seta a flag de tipo com 2 (float)
 _continueInColumn:
	pop ecx
	inc ecx
	pop edx											; Retorna para edx o offset atual
	add edx, SIZELINE								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _scjmpEndSum									; Se ultrapassou, a soma foi realizada com sucesso
	jmp _scmakingSum								; Senao, continua somando
 _scjmpEndSum:
	jmp _endSum

 _differentColumn:
	cmp al, bl
	ja _dcFirstGreater
 _dcSecondGreater:
	xchg eax, ebx
 _dcFirstGreater:
  	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8	
	cmp ax, bx
	jnz _dcjmpToError
	rol eax, 8
	rol ebx, 8
	invoke getOffsetCell, eax						; Obtem o offset da ultima celula
	lea edx, offsetEnd
	mov [edx], eax
	invoke getOffsetCell, ebx						; Obtem o offset da primeira celula
	lea edx, offsetBegin
	mov [edx], eax
	mov edx, offsetBegin							; Salva em edx o offset da primeira celula
	mov ecx, 0
 _dcmakingSum:
	push edx										; Salva o offset atual
	push ecx
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _dcjmpToError
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _dcjmpToError
	cmp cl, 1										; Valor inteiro
	jz _dcvalInt
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _dcvalInt
	cmp cl, 2										; Valor real
	jz _dcvalFloat
	cmp cl, 5										; Valor real decorrente de funcao
	jz _dcvalFloat
 _dcjmpToError:
	jmp _error
 _dcvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fadd ST(1), ST(0)								; Soma o valor atual com a soma anterior
	fstp trashReal									; Deixa so um valor na pilha
	jmp _continueInLine
 _dcvalFloat:
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fadd ST(1), ST(0)								; Soma o valor atual com a soma anterior
	fstp trashReal									; Deixa so um valor na pilha
	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2
	mov [edx], bl									; Seta a flag de tipo com 2 (float)
 _continueInLine:
	pop ecx
	inc ecx
	pop edx											; Retorna para edx o offset atual
	add edx, SIZECELL								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _jmpEndSum									; Se ultrapassou, a soma foi realizada com sucesso
	jmp _dcmakingSum								; Senao, continua somando
 _jmpEndSum:
	jmp _endSum

 _endSum:
	mov edx, offsetCell								; Salva em edx o offset da celula destino
	mov bl, flagVal									; Move para bl o valor da flag de tipo
	cmp bl, 1
	jz _valint										; Salta se o resultado for decorrente apenas de inteiros
	cmp bl, 2
	jz _valfloat									; Salta se algum dos numeros for float
	jmp _error
 _valint:
	mov al, 4										; Move 4 para al para indicar o tipo de dado contido na celula
	mov [edx], al									; Salva 4 na celula
	add edx, INTPOS									; Move edx para a posicao inteira
	lea ebx, flagval								; Obtem o endereco de flagval, usado agora como auxiliar
	mov [ebx], ecx									; Move ecx que contem o numero de celulas utilizadas para flagval
	fild DWORD PTR flagval							; Carrega flagval na pilha de floats
	fdiv ST(1), ST(0)								; Efetua a divisao da soma total pelo numero de valores
	fstp trashreal									; Remove valor lixo da pilha
	fistp DWORD PTR[edx]							; Salva o resultado na celula
	add edx, FNCPOS - INTPOS						; Move edx para o inicio da area que contem a funcao
	invoke Str_copy, addr inputBuffer, edx			; Copia a funcao para a celula
	jmp _sucess
 _valfloat:
	mov al, 5										; Move 5 para al para indicar o tipo de dado contido na celula
	mov [edx], al
	add edx, FLTPOS									; Move edx para a posicao float
	lea ebx, flagval								; Obtem o endereco de flagval, usado agora como auxiliar
	mov [ebx], ecx									; Move ecx que contem o numero de celulas utilizadas para flagval
	fild DWORD PTR flagval							; Carrega flagval na pilha de floats
	fdiv ST(1), ST(0)								; Efetua a divisao da soma total pelo numero de valores
	fstp trashreal									; Remove valor lixo da pilha
	fstp REAL4 PTR[edx]								; Salva o resultado na celula
	add edx, FNCPOS - FLTPOS						; Move edx para o inicio da area que contem a funcao
	invoke Str_copy, addr inputBuffer, edx			; Copia a funcao para a celula
	jmp _sucess
 _error:											; Exibe uma mensagem de erro
	pop eax
	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Nao foi possivel calcular a media."
	jmp _end

 _sucess:											; Exibe uma mensagem informando que a operacao foi concluida com sucesso
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Media calculada com sucesso."
 _end:
 	invoke UpdatePlan
	ret
Func_MED ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Calcula o menor valor dentre um conjunto de celulas
; Recebe stringPtr que eh um ponteiro para a string digitada pelo usuario
; a string deve conter a primeira e a ultima celula do intervalo (mesma linha ou mesma coluna/ vizinhas ou nao)
; Exige que a celula de destino tenha sido previamente selecionada
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_MIN PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offsetBegin: DWORD
	LOCAL offsetEnd: DWORD
	LOCAL flagval: BYTE
	LOCAL min: REAL4

	mov bl, isSelectedCell							; Verifica se ha alguma celula selecionada
	cmp bl, 0
	jnz _cellSelected								; Salta para _error caso não haja
	jmp _error
 _cellSelected:

	lea edx, flagval
	mov ebx, 1									
	mov [edx], ebx									; Seta a flag de tipo de valor como 1

	fld REAL4 PTR infinity
	fstp min										; Salva o menor valor como infinito
	fld REAL4 PTR min								; Carrega o menor valor na pilha

	mov edx, StringPTR								; Obtem o offset da entrada
	add edx, 4
	mov eax, [edx]									; Obtem o nome da primera celula
	and eax, 00FFFFFFh								; Limpa o byte mais significativo
	add edx, 4
 _isTheSameColumn:
	mov ebx, [edx]									; Obtem o nome da segunda celula
	cmp al, bl										; Verifica se a coluna da primeira celula eh igual a da segunda
	jz _sameColumn									; Salta se for a mesma coluna
	jmp _differentColumn							; Salta se a coluna for diferente (espera-se que seja mesma linha)

 _sameColumn:
 	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8
	cmp eax, ebx									; Compara as linhas para verificar se a segunda eh maior que a primeira
	jb _scSecondGreater								; Se a segunda for maior que a primeira
 _scFirstGreater:
	xchg eax, ebx									; Corrige caso a primeira seja maior q a segunda, invertendo as celulas
 _scSecondGreater:
	rol eax, 8										; Rotaciona para a forma inicial
	rol ebx, 8
	invoke getOffsetCell, eax						; Obtem o offset da celula inicial
	xchg eax, ebx									; Troca para salvar os valor
	invoke getOffsetCell, eax						; Obtem o offset da celula final
	lea edx, offsetEnd
	mov [edx], eax									; Salva a celula final em offsetEnd
	lea edx, offsetBegin
	mov [edx], ebx									; Salva a celula inicial em offsetBegin
	mov edx, offsetBegin

 _scdoingComparison:
	push edx										; Salva o offset atual
	fld min
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _scjmpToError	
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _scjmpToError
	cmp cl, 1										; Valor inteiro
	jz _scvalInt
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _scvalInt
	cmp cl, 2										; Valor real
	jz _scvalFloat
	cmp cl, 5										; Valor real decorrente de funcao
	jz _scvalFloat
 _scjmpToError:
	jmp _error
 _scvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	ja _continueInColumn
	fild DWORD PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp min

	jmp _continueInColumn
 _scvalFloat:
	push edx
 	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2
	mov [edx], ebx									; Seta a flag de tipo com 2 (float)
	pop edx	
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	ja _continueInColumn
	fld REAL4 PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp min

 _continueInColumn:
	fstp trashReal
	pop edx											; Retorna para edx o offset atual
	add edx, SIZELINE								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _scjmpEndMin									; Se ultrapassou, todos os valores foram verificados
	jmp _scdoingComparison							; Senao, continua verificando

 _scjmpEndMin:
	jmp _endMin

 _differentColumn:
	cmp al, bl
	ja _dcFirstGreater
 _dcSecondGreater:
	xchg eax, ebx
 _dcFirstGreater:
  	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8
	cmp ax, bx
	jnz _dcjmpToError
	rol eax, 8
	rol ebx, 8
	invoke getOffsetCell, eax						; Obtem o offset da ultima celula
	lea edx, offsetEnd
	mov [edx], eax
	invoke getOffsetCell, ebx						; Obtem o offset da primeira celula
	lea edx, offsetBegin
	mov [edx], eax	
	mov edx, offsetBegin							; Salva em edx o offset da primeira celula
	mov ecx, 0
 _dcdoingComparison:
	push edx										; Salva o offset atual
	fld min
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _dcjmpToError
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _dcjmpToError
	cmp cl, 1										; Valor inteiro
	jz _dcvalInt
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _dcvalInt
	cmp cl, 2										; Valor real
	jz _dcvalFloat
	cmp cl, 5										; Valor real decorrente de funcao
	jz _dcvalFloat
 _dcjmpToError:
	jmp _error
 _dcvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	ja _continueInLine
	fild DWORD PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp min
	jmp _continueInLine
 _DcvalFloat:
	push edx
 	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2
	mov [edx], ebx									; Seta a flag de tipo com 2 (float)
	pop edx
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	ja _continueInLine
	fld REAL4 PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp min
 _continueInLine:
	fstp trashReal

	pop edx											; Retorna para edx o offset atual
	add edx, SIZECELL								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _endMin										; Se ultrapassou, todos as celulas foram verificadas
	jmp _dcdoingComparison							; Senao, continua verificando

 _endMin:
	mov edx, offsetCell								; Salva em edx o offset da celula destino
	mov bl, flagVal									; Move para bl o valor da flag de tipo
	cmp bl, 1
	jz _valint										; Salta se o resultado for decorrente apenas de inteiros
	cmp bl, 2
	jz _valfloat									; Salta se algum dos numeros for float
	jmp _error
 _valint:
	mov al, 4										; Move 4 para al para indicar o tipo de dado contido na celula
	mov [edx], al									; Salva 4 na celula
	add edx, INTPOS									; Move edx para a posicao inteira
	fld REAL4 PTR min
	fistp DWORD PTR[edx]							; Salva o resultado na celula
	add edx, FNCPOS - INTPOS						; Move edx para o inicio da area que contem a funcao
	invoke Str_copy, addr inputBuffer, edx			; Copia a funcao para a celula
	jmp _sucess
 _valfloat:
	mov al, 5										; Move 5 para al para indicar o tipo de dado contido na celula
	mov [edx], al
	add edx, FLTPOS									; Move edx para a posicao float
	fld REAL4 PTR min
	fstp REAL4 PTR[edx]								; Salva o resultado na celula
	add edx, FNCPOS - FLTPOS						; Move edx para o inicio da area que contem a funcao
	invoke Str_copy, addr inputBuffer, edx			; Copia a funcao para a celula
	jmp _sucess

 _error:											; Exibe uma mensagem de erro
	pop eax
	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Nao foi possivel obter o menor valor."
	jmp _end
 _sucess:											; Exibe uma mensagem informando que a operacao foi concluida com sucesso
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Menor valor obtido com sucesso."
 _end:
 	invoke UpdatePlan
	ret
Func_MIN ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Abre a planilha a partir de um arquivo
; Nao recebe nenhum parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_OPN PROC USES EAX EBX ECX EDX
	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27							; Pede para o usuario digitar o nome do arquivo desejado
	mWrite "Digite o nome desejado ja com a extensao. ex: planilha.pln: "
	mov edx, offset filename
	mov ecx, 15
	call ReadString									; Recebe o nome do arquivo informado pelo usuario
	mov edx, offset filename
	invoke CreateFile,								; Tenta abrir o arquivo
		edx,
		GENERIC_READ,
		DO_NOT_SHARE,
		NULL,
		OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL,
		0
	push eax										; Salva eax
	cmp eax, INVALID_HANDLE_VALUE
	jz _error										; Salta para _error caso nao seja possivel abrir o arquivo
	mov edx, offset line1							; Move para edx o offset da primeira celula
	mov ecx, TOTALCELLS								; Move para ecx o numero total de celulas
	call ReadFromFile								; Le todas as celulas salvas no arquivo
	jmp _sucess
 _error:											; Exibe uma mensagem de erro caso nao seja possivel abrir o arquivo
	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Nao foi possivel abrir a planilha."
	jmp _end
 _sucess:											; Exibe uma mensagem informando que a planilha foi aberta com sucesso
	call CloseFile
	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Planilha aberta com sucesso."
 _end:												
	pop eax
	invoke UpdatePlan								; Atualiza a planilha
	invoke GotoPosXY, 0, 0
	ret
Func_OPN ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Salva a planilha em um arquivo de texto
; Nao recebe nenhum parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_SAV PROC
	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27							; Pede para o usuario digitar o nome do arquivo desejado
	mWrite "Digite o nome desejado ja com a extensao. ex: planilha.pln: "
	mov edx, offset filename
	mov ecx, 15
	call ReadString									; Recebe o nome do arquivo informado pelo usuario
	mov edx, offset filename
	call CreateOutputFile							; Tenta criar um arquivo com o nome digitado pelo usuario
	push eax
	cmp eax, INVALID_HANDLE_VALUE					; Verifica se foi possivel criar o arquivo
	jz _error										; Salta para _error caso nao seja possivel
	mov edx, offset line1							; Move para edx o offset da primeira celula
	mov ecx, TOTALCELLS								; Move para ecx o numero total de celulas
	call WriteToFile								; Salva o conteudo de todas as celulas no arquivo
	jmp _sucess
 _error:											; Exibe uma mensagem de erro caso nao seja possivel criar o arquivo para salvar
	pop eax
	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27
	mWrite "Nao foi possivel salvar a planilha."
	jmp _end
 _sucess:											; Exibe uma mensagem informando que a planilha foi aberta com sucesso
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27
	mWrite "Planilha salva com sucesso."
 	pop eax
	call CloseFile									; Fecha o arquivo
 _end:
	ret
Func_SAV ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Atribui uma string para a celula selecionada
; Nao recebe nenhum parametro
; Exige que a celula destino tenha sido previamente selecionada
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_STR PROC USES ECX EDX
	xor ecx, ecx									; Limpa ECX
	mov cl, isSelectedCell							; Verifica se ha alguma celula selecionada (flag != 0)
	jecxz _error	

	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString

	invoke GotoPosXY, 1, 27
	mWrite "Digite a string desejada: "				; Pede para o usuario informar a string desejada

	mov edx, offsetCell								; Move para ECX o offset da celula selecionada
	mov cl, 3										; Valor que seta a flag de string da celula
	mov [edx], cl									; Seta a flag de string
	add edx, STRPOS
	push edx
	mov ecx, 18
 _clearString:										; Limpa a area de string da celula
	mov [edx], cl
	inc edx
	loop _clearString
	pop edx
	mov ecx, 19	
	call ReadString									; Le a string digitada pelo usuario
	jmp _sucess
 _error:											; Exibe uma mensagem de erro caso nenhuma celula tenha sido selecionada
  	invoke GotoPosXY, 1, 27
 	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "SELECIONE UMA CELULA PRIMEIRO!"
	jmp _end
 _sucess:											; Exibe uma mensagem informando que a atribuicao de string foi realizada com sucesso
 	invoke GotoPosXY, 1, 27
 	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "String atribuida com sucesso!"
 _end:
 	invoke UpdatePlan								; Atualiza a planilha para garantir a consistencia de valores em celulas que dependam da celula alterada
	ret
Func_STR ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Subtrai o valor de duas celulas e salva na celula selecionada
; Recebe stringPtr que eh um ponteiro para a string digitada pelo usuario
; A string deve conter o nome de duas celulas
; A ordem da operacao eh primeiro valor menos o segundo
; Exige que a celula destino tenha sido previamente selecionada
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_SUB PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offset1: DWORD
	LOCAL offset2: DWORD
	LOCAL flagval: BYTE
	mov edx, StringPTR

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da primeira celula
	cmp eax, 0
	jz _error
	lea ebx, offset1
	mov [ebx], eax									; Salva o offset da primeira celula em offset1

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da segunda celula
	cmp eax, 0
	jz _error
	lea ebx, offset2
	mov [ebx], eax									; Salva o offset da segunda celula em offset2

 _val1:
	mov edx, offset1
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _error
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _error
	cmp cl, 1										; Valor inteiro
	jz _val1Int
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _val1Int
	cmp cl, 2										; Valor real
	jz _val1Float
	cmp cl, 5										; Valor real decorrente de funcao
	jz _val1Float
	jmp _error
 _val1Int:											; Caso o valor da primeira celula seja inteiro
 	add edx, INTPOS
	fild DWORD PTR[edx]
	lea edx, flagval
	mov ecx, 1
	mov [edx], cl
	jmp _val2
 _val1Float:										; Caso o valor da primeira celula seja float
 	add edx, FLTPOS
	fld REAL4 PTR[edx]
	lea edx, flagval
	mov ecx, 2
	mov [edx], cl
	jmp _val2
 _val2:
 	mov edx, offset2
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _error
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _error
	cmp cl, 1										; Valor inteiro
	jz _val2Int
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _val2Int
	cmp cl, 2										; Valor real
	jz _val2Float
	cmp cl, 5										; Valor real decorrente de funcao
	jz _val2Float
	jmp _error		
 _val2Int:											; Caso o valor da segunda celula seja inteiro
 	add edx, INTPOS	
	fild DWORD PTR[edx]
	lea edx, flagval
	mov ecx, 1
	mov [edx], cl
	jmp _doSub
 _val2Float:										; Caso o valor da segunda celula seja float
 	add edx, FLTPOS
	fld REAL4 PTR[edx]
	lea edx, flagval
	mov ecx, 2
	mov [edx], cl	
 _doSub:											; Efetua a subtracao
 	fsub st(1), st(0)
	fstp trashReal
 _verifyType:										; Verifica se o resultado da subtracao eh decorrente de inteiros ou contem algum float
 	mov dl, flagval
	cmp dl, 1
	jz _saveAsInt
	cmp dl, 2
	jz _saveAsFloat
	jmp _error			
 _saveAsInt:										; Salva como inteiro
 	mov edx, offsetCell	
	mov cl, 4
	mov [edx], cl
	push edx
	add edx, INTPOS
	fistp DWORD PTR[edx]
	add edx, FNCPOS - INTPOS
	invoke Str_copy, addr inputBuffer, edx
	jmp _sucess
 _saveAsFloat:										; Salva como float
  	mov edx, offsetCell
	mov cl, 5
	mov [edx], cl
	push edx
	add edx, FLTPOS
	fstp DWORD PTR[edx]
	add edx, FNCPOS - FLTPOS
	invoke Str_copy, addr inputBuffer, edx
	jmp _sucess
 _error:											; Exibe uma mensagem de erro caso uma das celulas contenha valores invalidos
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Alguma das celulas selecionadas eh invalida ou possui valores invalidos! Tente novamente."
	jmp _exit
 _sucess:											; Exibe uma mensagem informando que a operacao foi realizada com sucesso
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
 	invoke GotoPosXY, 1, 27
	mWrite "Subtracao efetuada com sucesso"
 _exit:
 	invoke UpdatePlan
	ret
Func_SUB ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Soma o valor de duas celulas e atribui ah celula selecionada
; Recebe stringPtr que eh um ponteiro para a string digitada pelo usuario
; A string deve conter o nome de duas celulas
; Exige que a celula destino tenha sido previamente selecionada
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Func_SUM PROC USES EAX EBX ECX EDX, StringPTR: PTR DWORD
	LOCAL offset1: DWORD
	LOCAL offset2: DWORD
	LOCAL flagval: BYTE
	lea edx, flagval	
	mov ebx, 1
	mov [edx], bl									; Inicializa a flag de tipo como 1

	mov edx, StringPTR

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da primeira celula
	cmp eax, 0
	jz _error
	lea ebx, offset1
	mov [ebx], eax									; Salva o offset da primeira celula em offset1

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da segunda celula
	cmp eax, 0
	jz _error
	lea ebx, offset2
	mov [ebx], eax									; Salva o offset da segunda celula em offset2

 _val1:
	mov edx, offset1
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _error
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _error
	cmp cl, 1										; Valor inteiro
	jz _val1Int	
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _val1Int
	cmp cl, 2										; Valor real
	jz _val1Float
	cmp cl, 5										; Valor real decorrente de funcao
	jz _val1Float		
	jmp _error
 _val1Int:
 	add edx, INTPOS
	fild DWORD PTR[edx]								; Carrega o inteiro nos registradores de ponto flutuante
	jmp _val2
 _val1Float:
 	add edx, FLTPOS
	fld REAL4 PTR[edx]								; Carrega o valor real nos registradores de ponto flutuante
	lea edx, flagval
	mov ecx, 2
	mov [edx], cl
	jmp _val2
 _val2:
 	mov edx, offset2
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _error
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _error
	cmp cl, 1										; Valor inteiro
	jz _val2Int
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _val2Int
	cmp cl, 2										; Valor real
	jz _val2Float
	cmp cl, 5										; Valor real decorrente de funcao
	jz _val2Float
	jmp _error
 _val2Int:
 	add edx, INTPOS
	fild DWORD PTR[edx]								; Carrega o inteiro nos registradores de ponto flutuante
	jmp _doSum
 _val2Float:
 	add edx, FLTPOS
	fld REAL4 PTR[edx]								; Carrega o valor real nos registradores de ponto flutuante
	lea edx, flagval
	mov ecx, 2
	mov [edx], cl
 _doSum:
 	fadd st(0), st(1)								; Salva o valor da primeira celula com o da segunda
 _verifyType:										; Verifica o valor da flag de tipo para definir como o resultado sera salvo
 	mov dl, flagval
	cmp dl, 1
	jz _saveAsInt
	cmp dl, 2
	jz _saveAsFloat
	jmp _error
 _saveAsInt:										; Salva o valor na celula destino como um inteiro
 	mov edx, offsetCell
	mov cl, 4
	mov [edx], cl
	add edx, INTPOS
	fistp DWORD PTR[edx]
	add edx, FNCPOS - INTPOS
	invoke Str_copy, addr inputBuffer, edx			; Copia a funcao na celula destino
	fstp trashReal
	jmp _sucess
 _saveAsFloat:										; Salva o valor na celula destino como um float
  	mov edx, offsetCell
	mov cl, 5
	mov [edx], cl
	add edx, FLTPOS
	fstp DWORD PTR[edx]
	add edx, FNCPOS - FLTPOS
	invoke Str_copy, addr inputBuffer, edx			; Copia a funcao na celula destino
	fstp trashReal
	jmp _sucess
 _error:											; Exibe uma mensagem de erro caso algum dos valores seja invalido ou nenhuma celula tenha sido selecionada
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
	invoke GotoPosXY, 1, 27
	mWrite "Alguma das celulas selecionadas eh invalida ou possui valores invalidos, ou nenhuma celula foi selecionada!"
	jmp _exit
 _sucess:											; Exibe uma mensagem informando que a soma foi efetuada com sucesso
 	invoke GotoPosXY, 1, 27
	mov edx, offset separator
	call WriteString
 	invoke GotoPosXY, 1, 27
	mWrite "Soma efetuada com sucesso"
 _exit:
 	invoke UpdatePlan
	ret
Func_SUM ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Atualiza toda a planilha
; Atualiza o conteudo das celulas 5 vezes para garantir que todas as operacoes foram realizadas
; Funcao chamada automaticamente
; Nao recebe nada como parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
UpdatePlan PROC USES EAX EBX ECX EDX
	mov edx, offset line1
	mov ecx, 3
 _reupdate:											; Loop para repetir a atualizacao 3 vezes
 	push ecx
	mov ecx, 120
 _update:											; Atualiza o conteudo da celula chamando a funcao processUpdate, que recebe a funcao presente na celula e o endereco da celula
 	mov al, [edx]									; So atualiza o conteudo da celula se a celula conter uma funcao
	cmp al, 4
	jz _continue
	cmp al, 5
	jz _continue
	jmp _noFormula
 _continue:
	mov eax, edx
 	add eax, FNCPOS
 	invoke ProcessUpdate, eax, edx
 _noFormula:
 	add edx, SIZECELL
	loop _update
	pop ecx
	loop _reupdate
	ret
UpdatePlan ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Processa a funcao presente na celula para atualizar
; Funcao chamada automaticamente
; Recebe como parametros stringPTR que eh um ponteiro para a funcao presente na celula e cellPtr que eh um ponteiro para a celula
; ---------------------------------------------------------------------------------------------------------------------------------------------------
ProcessUpdate PROC USES EAX EBX ECX EDX, stringPtr: PTR DWORD, cellPtr: PTR DWORD
	mov edx, stringPtr
	mov eax, [edx]
	
	; Verifica qual a funcao presente na cleula e chama a funcao referente

	mov edx, offset functions					
	add edx, 36
	cmp eax, [edx]								; MEX
	jz _maxt
	jmp _med
 _maxt:
	invoke Update_MAX, stringPtr, cellPtr
	jmp _end
 _med:
	add edx, 4
	cmp eax, [edx]								; MED
	jz _medt
 	jmp _min
 _medt:
	invoke Update_MED, stringPtr, cellPtr
	jmp _end
 _min:
	add edx, 4
	cmp eax, [edx]								; MIN
	jz _mint
	jmp _sub
 _mint:
	invoke Update_MIN, stringPtr, cellPtr
	jmp _end
 _sub:
	add edx, 16
	cmp eax, [edx]								; SUB
	jz _subt
	jmp _sum
 _subt:
	invoke Update_SUB, stringPtr, cellPtr
	jmp _end
 _sum:
	add edx, 4
	cmp eax, [edx]								; SUM
	jnz _error
	invoke Update_SUM, stringPtr, cellPtr
	jmp _end
 _error:
	invoke SetColor, (lightGray + (16 * blue))
	invoke GoToPosXY, 0, 27
	mov edx, offset separator
	call WriteString
 _end:
	ret
ProcessUpdate ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Faz o mesmo que a funcao de Maximo, porem, sem informar ao usuario se a operacao foi realizada com sucesso ou nao.
; Funcao chamada automaticamente
; Recebe como parametros stringPTR que eh um ponteiro para a funcao presente na celula e cellPtr que eh um ponteiro para a celula
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Update_MAX PROC USES EAX EBX ECX EDX, stringPTR: PTR DWORD, actualCell: PTR DWORD
	LOCAL offsetBegin: DWORD
	LOCAL offsetEnd: DWORD
	LOCAL flagval: BYTE
	LOCAL max: REAL4

	lea edx, flagval
	mov ebx, 0
	mov [edx], ebx

	fild infinityneg
	fstp max										

	mov ebx, 1
	mov [edx], ebx

	mov edx, StringPTR								; Obtem o offset da entrada
	add edx, 4
	mov eax, [edx]									; Obtem o nome da primera celula
	and eax, 00FFFFFFh								; Limpa o byte mais significativo
	add edx, 4
 _isTheSameColumn:
	mov ebx, [edx]									; Obtem o nome da segunda celula
	cmp al, bl										; Verifica se a coluna da primeira celula eh igual a da segunda
	jz _sameColumn									; Salta se for a mesma coluna
	jmp _differentColumn							; Salta se a coluna for diferente (espera-se que seja mesma linha)

 _sameColumn:
 	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8
	cmp eax, ebx									; Compara as linhas para verificar se a segunda eh maior que a primeira
	jb _scSecondGreater								; Se a segunda for maior que a primeira
 _scFirstGreater:
	xchg eax, ebx									; Corrige caso a primeira seja maior q a segunda, invertendo as celulas
 _scSecondGreater:
	rol eax, 8										; Rotaciona para a forma inicial
	rol ebx, 8
	invoke getOffsetCell, eax						; Obtem o offset da celula inicial
	xchg eax, ebx									; Troca para salvar os valor
	invoke getOffsetCell, eax						; Obtem o offset da celula final
	lea edx, offsetEnd
	mov [edx], eax									; Salva a celula final em offsetEnd
	lea edx, offsetBegin
	mov [edx], ebx									; Salva a celula inicial em offsetBegin
	mov edx, offsetBegin
	mov ecx, 0

 _scdoingComparison:
	push edx										; Salva o offset atual
	fld max
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _scjmpToError
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _scjmpToError
	cmp cl, 1										; Valor inteiro
	jz _scvalInt
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _scvalInt
	cmp cl, 2										; Valor real
	jz _scvalFloat	
	cmp cl, 5										; Valor real decorrente de funcao
	jz _scvalFloat
 _scjmpToError:
	jmp _end
 _scvalInt:	
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	jb _continueInColumn
	fild DWORD PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp max

	jmp _continueInColumn
 _scvalFloat:
	push edx
 	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2
	mov [edx], bl									; Seta a flag de tipo com 2 (float)
	pop edx
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	jb _continueInColumn		
	fld REAL4 PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp max

 _continueInColumn:
	fstp trashReal
	pop edx											; Retorna para edx o offset atual
	add edx, SIZELINE								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _scjmpEndMax									; Se ultrapassou, todos os valores foram verificados
	jmp _scdoingComparison							; Senao, continua verificando

 _scjmpEndMax:
	jmp _endMax

 _differentColumn:
	cmp al, bl
	ja _dcFirstGreater
 _dcSecondGreater:
	xchg eax, ebx
 _dcFirstGreater:
  	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8
	cmp ax, bx
	jnz _dcjmpToEnd
	rol eax, 8
	rol ebx, 8
	invoke getOffsetCell, eax						; Obtem o offset da ultima celula
	lea edx, offsetEnd
	mov [edx], eax
	invoke getOffsetCell, ebx						; Obtem o offset da primeira celula
	lea edx, offsetBegin	
	mov [edx], eax
	mov edx, offsetBegin							; Salva em edx o offset da primeira celula
	mov ecx, 0
 _dcdoingComparison:
	push edx										; Salva o offset atual
	fld max
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _dcjmpToEnd
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _dcjmpToEnd
	cmp cl, 1										; Valor inteiro
	jz _dcvalInt
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _dcvalInt
	cmp cl, 2										; Valor real
	jz _dcvalFloat	
	cmp cl, 5										; Valor real decorrente de funcao
	jz _dcvalFloat
 _dcjmpToEnd:
	jmp _end
 _dcvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	jb _continueInLine
	fild DWORD PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp max
	jmp _continueInLine
 _dcvalFloat:
	push edx
 	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2
	mov [edx], bl									; Seta a flag de tipo com 2 (float)
	pop edx
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	jb _continueInLine
	fld REAL4 PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp max	
 _continueInLine:
	fstp trashReal

	pop edx											; Retorna para edx o offset atual
	add edx, SIZECELL								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _endMax										; Se ultrapassou, todos as celulas foram verificadas
	jmp _dcdoingComparison							; Senao, continua verificando

 _endMax:
	mov edx, actualCell								; Salva em edx o offset da celula destino
	mov bl, flagVal									; Move para bl o valor da flag de tipo
	cmp bl, 1
	jz _valint										; Salta se o resultado for decorrente apenas de inteiros
	cmp bl, 2
	jz _valfloat									; Salta se algum dos numeros for float
	jmp _end
 _valint:
	mov al, 4										; Move 4 para al para indicar o tipo de dado contido na celula
	mov [edx], al									; Salva 4 na celula
	add edx, INTPOS									; Move edx para a posicao inteira
	fld REAL4 PTR max
	fistp DWORD PTR[edx]							; Salva o resultado na celula
	jmp _end
 _valfloat:	
	mov al, 5										; Move 5 para al para indicar o tipo de dado contido na celula
	mov [edx], al
	add edx, FLTPOS									; Move edx para a posicao float
	fld REAL4 PTR max
	fstp REAL4 PTR[edx]								; Salva o resultado na celula
	jmp _end
 _end:
	ret
Update_MAX ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Faz o mesmo que a funcao de Media, porem, sem informar ao usuario se a operacao foi realizada com sucesso ou nao.
; Funcao chamada automaticamente
; Recebe como parametros stringPTR que eh um ponteiro para a funcao presente na celula e cellPtr que eh um ponteiro para a celula
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Update_MED PROC USES EAX EBX ECX EDX, stringPTR: PTR DWORD, actualCell: PTR DWORD
	LOCAL offsetBegin: DWORD
	LOCAL offsetEnd: DWORD
	LOCAL flagval: BYTE

	lea edx, flagVal
	mov ebx, 0
	mov [edx], ebx									; Seta a flag de tipo como 1
	fild DWORD PTR flagVal							; Inicializa a pilha com 0

	lea edx, flagVal
	mov ebx, 1
	mov [edx], ebx									; Seta a flag de tipo como 1
		
	mov edx, StringPTR								; Obtem o offset da entrada
	add edx, 4	
	mov eax, [edx]									; Obtem o nome da primera celula
	and eax, 00FFFFFFh								; Limpa o byte mais significativo
	add edx, 4	
 _isTheSameColumn:
	mov ebx, [edx]									; Obtem o nome da segunda celula
	cmp al, bl										; Verifica se a coluna da primeira celula eh igual a da segunda
	jz _sameColumn									; Salta se for a mesma coluna
	jmp _differentColumn							; Salta se a coluna for diferente (espera-se que seja mesma linha)

 _sameColumn:
 	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8
	cmp eax, ebx									; Compara as linhas para verificar se a segunda eh maior que a primeira
	jb _scSecondGreater								; Se a segunda for maior que a primeira
 _scFirstGreater:
	xchg eax, ebx									; Corrige caso a primeira seja maior q a segunda, invertendo as celulas
 _scSecondGreater:
	rol eax, 8										; Rotaciona para a forma inicial
	rol ebx, 8	
	invoke getOffsetCell, eax						; Obtem o offset da celula inicial
	xchg eax, ebx									; Troca para salvar os valor
	invoke getOffsetCell, eax						; Obtem o offset da celula final
	lea edx, offsetEnd
	mov [edx], eax									; Salva a celula final em offsetEnd
	lea edx, offsetBegin
	mov [edx], ebx									; Salva a celula inicial em offsetBegin
	mov edx, offsetBegin
	mov ecx, 0
 _scmakingSum:
	push edx										; Salva o offset atual
	push ecx
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _scjmpToError
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _scjmpToError
	cmp cl, 1										; Valor inteiro
	jz _scvalInt	
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _scvalInt
	cmp cl, 2										; Valor real
	jz _scvalFloat
	cmp cl, 5										; Valor real decorrente de funcao
	jz _scvalFloat
 _scjmpToError:
	jmp _error
 _scvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fadd ST(1), ST(0)								; Soma o valor atual com a soma anterior
	fstp trashReal									; Deixa so um valor na pilha
	jmp _continueInColumn
 _scvalFloat:
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fadd ST(1), ST(0)								; Soma o valor atual com a soma anterior
	fstp trashReal									; Deixa so um valor na pilha
	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2
	mov [edx], ebx									; Seta a flag de tipo com 2 (float)
 _continueInColumn:
	pop ecx
	inc ecx
	pop edx											; Retorna para edx o offset atual
	add edx, SIZELINE								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _scjmpEndSum									; Se ultrapassou, a soma foi realizada com sucesso
	jmp _scmakingSum								; Senao, continua somando
 _scjmpEndSum:
	jmp _endSum

 _differentColumn:
	cmp al, bl
	ja _dcFirstGreater
 _dcSecondGreater:
	xchg eax, ebx
 _dcFirstGreater:
  	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8
	cmp ax, bx
	jnz _dcjmpToError
	rol eax, 8
	rol ebx, 8
	invoke getOffsetCell, eax						; Obtem o offset da ultima celula
	lea edx, offsetEnd
	mov [edx], eax
	invoke getOffsetCell, ebx						; Obtem o offset da primeira celula
	lea edx, offsetBegin
	mov [edx], eax
	mov edx, offsetBegin							; Salva em edx o offset da primeira celula
	mov ecx, 0
 _dcmakingSum:
	push edx										; Salva o offset atual
	push ecx
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _dcjmpToError
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _dcjmpToError
	cmp cl, 1										; Valor inteiro
	jz _dcvalInt	
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _dcvalInt
	cmp cl, 2										; Valor real
	jz _dcvalFloat	
	cmp cl, 5										; Valor real decorrente de funcao
	jz _dcvalFloat
 _dcjmpToError:
	jmp _error
 _dcvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fadd ST(1), ST(0)								; Soma o valor atual com a soma anterior
	fstp trashReal									; Deixa so um valor na pilha
	jmp _continueInLine
 _dcvalFloat:
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fadd ST(1), ST(0)								; Soma o valor atual com a soma anterior
	fstp trashReal									; Deixa so um valor na pilha
	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2
	mov [edx], bl									; Seta a flag de tipo com 2 (float)
 _continueInLine:
	pop ecx
	inc ecx
	pop edx											; Retorna para edx o offset atual
	add edx, SIZECELL								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _jmpEndSum									; Se ultrapassou, a soma foi realizada com sucesso
	jmp _dcmakingSum								; Senao, continua somando
 _jmpEndSum:
	jmp _endSum

 _endSum:
	mov edx, actualCell								; Salva em edx o offset da celula destino
	mov bl, flagVal									; Move para bl o valor da flag de tipo
	cmp bl, 1
	jz _valint										; Salta se o resultado for decorrente apenas de inteiros
	cmp bl, 2
	jz _valfloat									; Salta se algum dos numeros for float
	jmp _error
 _valint:
	mov al, 4										; Move 4 para al para indicar o tipo de dado contido na celula
	mov [edx], al									; Salva 4 na celula
	add edx, INTPOS									; Move edx para a posicao inteira
	lea ebx, flagval								; Obtem o endereco de flagval, usado agora como auxiliar
	mov [ebx], ecx									; Move ecx que contem o numero de celulas utilizadas para flagval
	fild DWORD PTR flagval							; Carrega flagval na pilha de floats
	fdiv ST(1), ST(0)								; Efetua a divisao da soma total pelo numero de valores
	fstp trashreal									; Remove valor lixo da pilha
	fistp DWORD PTR[edx]							; Salva o resultado na celula
	jmp _end
 _valfloat:
	mov al, 5										; Move 5 para al para indicar o tipo de dado contido na celula
	mov [edx], al
	add edx, FLTPOS									; Move edx para a posicao float
	lea ebx, flagval								; Obtem o endereco de flagval, usado agora como auxiliar
	mov [ebx], ecx									; Move ecx que contem o numero de celulas utilizadas para flagval
	fild DWORD PTR flagval							; Carrega flagval na pilha de floats
	fdiv ST(1), ST(0)								; Efetua a divisao da soma total pelo numero de valores
	fstp trashreal									; Remove valor lixo da pilha
	fstp REAL4 PTR[edx]								; Salva o resultado na celula
	jmp _end
 _error:
	pop eax
 _end:
	ret
Update_MED ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Faz o mesmo que a funcao de minimo, porem, sem informar ao usuario se a operacao foi realizada com sucesso ou nao.
; Funcao chamada automaticamente
; Recebe como parametros stringPTR que eh um ponteiro para a funcao presente na celula e cellPtr que eh um ponteiro para a celula
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Update_MIN PROC USES EAX EBX ECX EDX, stringPTR: PTR DWORD, actualCell: PTR DWORD
	LOCAL offsetBegin: DWORD
	LOCAL offsetEnd: DWORD
	LOCAL flagval: BYTE
	LOCAL min: REAL4

	fld REAL4 PTR infinity							; Carrega um valor infinito na pilha de floats
	fstp min										; Salva o menor valor inicialmente como infinito

	lea edx, flagval								; Inicializa a variavel de flag como 1 (inteiro)						
	mov ebx, 1
	mov [edx], bl

	mov edx, StringPTR								; Obtem o offset da entrada
	add edx, 4
	mov eax, [edx]									; Obtem o nome da primera celula
	and eax, 00FFFFFFh								; Limpa o byte mais significativo
	add edx, 4
 _isTheSameColumn:	
	mov ebx, [edx]									; Obtem o nome da segunda celula
	cmp al, bl										; Verifica se a coluna da primeira celula eh igual a da segunda
	jz _sameColumn									; Salta se for a mesma coluna
	jmp _differentColumn							; Salta se a coluna for diferente (espera-se que seja mesma linha)
		
 _sameColumn:
 	ror eax, 8										; Rotaciona para comparar a linha
	ror ebx, 8
	cmp eax, ebx									; Compara as linhas para verificar se a segunda eh maior que a primeira
	jb _scSecondGreater								; Se a segunda for maior que a primeira
 _scFirstGreater:
	xchg eax, ebx									; Corrige caso a primeira seja maior q a segunda, invertendo as celulas
 _scSecondGreater:
	rol eax, 8										; Rotaciona para a forma inicial
	rol ebx, 8
	invoke getOffsetCell, eax						; Obtem o offset da celula inicial
	xchg eax, ebx									; Troca para salvar os valor
	invoke getOffsetCell, eax						; Obtem o offset da celula final
	lea edx, offsetEnd
	mov [edx], eax									; Salva a celula final em offsetEnd
	lea edx, offsetBegin
	mov [edx], ebx									; Salva a celula inicial em offsetBegin
	mov edx, offsetBegin

 _scdoingComparison:
	push edx										; Salva o offset atual
	fld min
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _dcjmpToEnd
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _dcjmpToEnd
	cmp cl, 1										; Valor inteiro
	jz _scvalInt
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _scvalInt
	cmp cl, 2										; Valor real
	jz _scvalFloat
	cmp cl, 5										; Valor real decorrente de funcao
	jz _scvalFloat
 _scjmpToEnd:
	jmp _end
 _scvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	ja _continueInColumn
	fild DWORD PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp min	

	jmp _continueInColumn
 _scvalFloat:
	push edx
 	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2
	mov [edx], bl									; Seta a flag de tipo com 2 (float)
	pop edx
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	ja _continueInColumn	
	fld REAL4 PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp min

 _continueInColumn:
	fstp trashReal
	pop edx											; Retorna para edx o offset atual
	add edx, SIZELINE								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _scjmpEndMin									; Se ultrapassou, todos os valores foram verificados
	jmp _scdoingComparison							; Senao, continua verificando
 _scjmpEndMin:	
	jmp _endMin
 _differentColumn:
	cmp al, bl										; Verifica se as celulas estao em ordem crescente
	ja _dcFirstGreater	
 _dcSecondGreater:									; Ajusta a ordem das celulas caso necessario
	xchg eax, ebx
 _dcFirstGreater:
	invoke getOffsetCell, eax						; Obtem o offset da ultima celula
	lea edx, offsetEnd
	mov [edx], eax
	invoke getOffsetCell, ebx						; Obtem o offset da primeira celula
	lea edx, offsetBegin
	mov [edx], eax
	mov edx, offsetBegin							; Salva em edx o offset da primeira celula
	mov ecx, 0
 _dcdoingComparison:
	push edx										; Salva o offset atual
	fld min
	xor ecx, ecx
	mov cl, [edx]
	cmp cl, 3										; Erro caso a celula contenha uma string
	jz _dcjmpToEnd
	cmp cl, 0										; Erro caso a celula não contenha valor
	jz _dcjmpToEnd
	cmp cl, 1										; Valor inteiro
	jz _dcvalInt
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _dcvalInt
	cmp cl, 2										; Valor real
	jz _dcvalFloat
	cmp cl, 5										; Valor real decorrente de funcao
	jz _dcvalFloat
 _dcjmpToEnd:
	jmp _end
 _dcvalInt:
	add edx, INTPOS									; Move para a posicao inteira
	fild DWORD PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	ja _continueInLine
	fild DWORD PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp min
	jmp _continueInLine
 _DcvalFloat:
	push edx
 	lea edx, flagVal								; Obtem o offset da flag de tipo
	mov ebx, 2	
	mov [edx], bl									; Seta a flag de tipo com 2 (float)
	pop edx
	add edx, FLTPOS									; Move para a posicao float
	fld REAL4 PTR[edx]								; Carrega na pilha
	fcomip ST(0), ST(1)								; Compara o valor da celula atual com o maior valor atual
	ja _continueInLine	
	fld REAL4 PTR[edx]								; Atribui a maior valor o valor da celula atual caso esse seja maior
	fstp min
 _continueInLine:
	fstp trashReal
	pop edx											; Retorna para edx o offset atual
	add edx, SIZECELL								; Move para a celula na mesma coluna na linha seguinte
	cmp edx, offsetEnd								; Verifica se edx ultrapassou o offset da ultima celula
	ja _endMin										; Se ultrapassou, todos as celulas foram verificadas
	jmp _dcdoingComparison							; Senao, continua verificando
 _endMin:
	mov edx, actualCell								; Salva em edx o offset da celula destino
	mov bl, flagVal									; Move para bl o valor da flag de tipo
	cmp bl, 1	
	jz _valint										; Salta se o resultado for decorrente apenas de inteiros
	cmp bl, 2
	jz _valfloat									; Salta se algum dos numeros for float
	jmp _end
 _valint:
	mov al, 4										; Move 4 para al para indicar o tipo de dado contido na celula
	mov [edx], al									; Salva 4 na celula
	add edx, INTPOS									; Move edx para a posicao inteira
	fld REAL4 PTR min
	fistp DWORD PTR[edx]							; Salva o resultado na celula
	jmp _end
 _valfloat:
	mov al, 5										; Move 5 para al para indicar o tipo de dado contido na celula
	mov [edx], al
	add edx, FLTPOS									; Move edx para a posicao float
	fld REAL4 PTR min
	fstp REAL4 PTR[edx]								; Salva o resultado na celula
	jmp _end
 _end:
	ret
Update_MIN ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Faz o mesmo que a funcao de Subtracao, porem, sem informar ao usuario se a operacao foi realizada com sucesso ou nao.
; Funcao chamada automaticamente
; Recebe como parametros stringPTR que eh um ponteiro para a funcao presente na celula e cellPtr que eh um ponteiro para a celula
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Update_SUB PROC USES EAX EBX ECX EDX, stringPTR: PTR DWORD, actualCell: PTR DWORD
	LOCAL offset1: DWORD
	LOCAL offset2: DWORD
	LOCAL flagval: BYTE
	lea edx, flagval
	mov ebx, 1
	mov [edx], bl

	mov edx, StringPTR

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da primeira celula
	lea ebx, offset1
	mov [ebx], eax									; Salva o offset da primeira celula em offset1

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da segunda celula
	lea ebx, offset2
	mov [ebx], eax									; Salva o offset da segunda celula em offset2

 _val1:
	mov edx, offset1
	mov cl, [edx]
	cmp cl, 1										; Valor inteiro
	jz _val1Int	
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _val1Int
	cmp cl, 2										; Valor real
	jz _val1Float
	cmp cl, 5										; Valor real decorrente de funcao
	jz _val1Float
	jmp _exit
 _val1Int:
 	add edx, INTPOS
	fild DWORD PTR[edx]								; Carrega o inteiro nos registradores de ponto flutuante
	jmp _val2
 _val1Float:
 	add edx, FLTPOS
	fld REAL4 PTR[edx]								; Carrega um real na pilha de registradores de ponto flutuante
	lea edx, flagval
	mov ecx, 2
	mov [edx], cl									; Seta o flag de tipo de valor como 2
	jmp _val2
 _val2:
 	mov edx, offset2
	mov cl, [edx]
	cmp cl, 1										; Valor inteiro
	jz _val2Int
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _val2Int
	cmp cl, 2										; Valor real
	jz _val2Float	
	cmp cl, 5										; Valor real decorrente de funcao
	jz _val2Float
	jmp _exit
 _val2Int:
 	add edx, INTPOS
	fild DWORD PTR[edx]								; Carrega um valor inteiro na pilha
	jmp _doSum
 _val2Float:
 	add edx, FLTPOS
	fld REAL4 PTR[edx]								; Carrega um valor float na pilha
	lea edx, flagval
	mov ecx, 2
	mov [edx], cl
 _doSum:
 	fsub st(1), st(0)								; Efetua a subtracao
	fstp trashReal									; Remove lixo da pilha
 _verifyType:
 	mov dl, flagval									; Verifica o tipo de valor obtido para salvar na area certa da celula
	cmp dl, 1
	jz _saveAsInt
	cmp dl, 2
	jz _saveAsFloat
	jmp _exit
 _saveAsInt:
 	mov edx, actualCell								; Salva como inteiro
	mov ecx, 4
	mov [edx], cl
	add edx, INTPOS
	fistp DWORD PTR[edx]
	jmp _exit
 _saveAsFloat:
  	mov edx, actualCell								; Salva como float
	mov ecx, 5
	mov [edx], cl
	add edx, FLTPOS
	fstp DWORD PTR[edx]
 _exit:
	ret
Update_SUB ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Faz o mesmo que a funcao de Soma, porem, sem informar ao usuario se a operacao foi realizada com sucesso ou nao.
; Funcao chamada automaticamente
; Recebe como parametros stringPTR que eh um ponteiro para a funcao presente na celula e cellPtr que eh um ponteiro para a celula
; ---------------------------------------------------------------------------------------------------------------------------------------------------
Update_SUM PROC USES EAX EBX ECX EDX, stringPTR: PTR DWORD, actualCell: PTR DWORD
	LOCAL offset1: DWORD
	LOCAL offset2: DWORD
	LOCAL flagval: BYTE
	lea edx, flagval
	mov ebx, 1
	mov [edx], bl									; Inicializa a flag de tipo como 1

	mov edx, StringPTR

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da primeira celula
	lea ebx, offset1
	mov [ebx], eax									; Salva o offset da primeira celula em offset1

	add edx, 4
	mov ebx, [edx]
	invoke GetOffsetCell, ebx						; Obtem o offset da segunda celula
	lea ebx, offset2
	mov [ebx], eax									; Salva o offset da segunda celula em offset2

 _val1:
	mov edx, offset1
	mov cl, [edx]
	cmp cl, 1										; Valor inteiro
	jz _val1Int
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _val1Int
	cmp cl, 2										; Valor real
	jz _val1Float
	cmp cl, 5										; Valor real decorrente de funcao
	jz _val1Float
	jmp _exit										; Encerra a operacao caso não houver um valor valido
 _val1Int:
 	add edx, INTPOS									; Move edx para a area de inteiro
	fild DWORD PTR[edx]								; Carrega o primeiro valor como inteiro na pilha
	jmp _val2										; Salta para a verificacao do segundo valor
 _val1Float:
 	add edx, FLTPOS									; Move edx para a area de float
	fld REAL4 PTR[edx]								; Carrega o primeiro valor como float na pilha
	lea edx, flagval								; Muda o valor da flag de tipo para 2 (float)
	mov ecx, 2
	mov [edx], cl
	jmp _val2										; Salta para a verificacao do segundo valor
 _val2:
 	mov edx, offset2								; Salva em edx o offset da segunda celula
	mov cl, [edx]
	cmp cl, 1										; Valor inteiro
	jz _val2Int
	cmp cl, 4										; Valor inteiro decorrente de funcao
	jz _val2Int
	cmp cl, 2										; Valor real
	jz _val2Float
	cmp cl, 5										; Valor real decorrente de funcao
	jz _val2Float
	jmp _exit										; Encerra a operacao caso não houver um valor valido
 _val2Int:
 	add edx, INTPOS									; Move edx para a area de inteiro da celula
	fild DWORD PTR[edx]								; Carrega o segundo valor como inteiro na pilha
	jmp _doSum										; Salta para o trecho onde é realizada a operação de soma
 _val2Float:
 	add edx, FLTPOS									; Move edx para a area de float
	fld REAL4 PTR[edx]								; Carrega o segundo como float na pilha
	lea edx, flagval								; Muda o valor da flag de tipo para 2 (float)
	mov ecx, 2
	mov [edx], cl
 _doSum:	
 	fadd st(0), st(1)								; Soma o primeiro valor com o segundo
 _verifyType:
 	mov dl, flagval									; Verifica o tipo de resultado da soma
	cmp dl, 1							
	jz _saveAsInt									; Salva como inteiro se for 1
	cmp dl, 2
	jz _saveAsFloat									; Salva como float se for 2
	jmp _exit	
 _saveAsInt:
 	mov edx, actualCell								; Move para edx o offset da celula destino
	mov ecx, 4										; Seta o tipo de valor da celula como sendo inteiro decorrente de funcao
	mov [edx], cl
	add edx, INTPOS									; Move edx para a area de inteiro da celula
	fistp DWORD PTR[edx]							; Retira o valor da pilha e salva na celula
	fstp trashReal									; Limpa a pilha
	jmp _exit										; Encerra a operacao
 _saveAsFloat:
  	mov edx, actualCell								; Move para edx o offset da celula destino
	mov ecx, 5										; Seta o tipo de valor da celula como sendo float decorrente de funcao
	mov [edx], cl
	add edx, FLTPOS									; Move edx para a area de float da celula
	fstp DWORD PTR[edx]								; Retira o valor da pilha e salva na celula
	fstp trashReal									; Limpa a pilha
 _exit:	
	ret
Update_SUM ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Retorna o offset de uma celula a partir do nome dessa celula.
; Recebe uma DWORD que eh uma string que representa a celula e um espaço em branco. ex: "A01 "
; Retorna em EAX o offset da celula desejada ou 0 em caso de erro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
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
 	mov eax, 0										; Retorna 0 em eax caso não for possivel obter o offset da celula
 _end:
	ret
GetOffsetCell ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Atribui a cor de fundo e da fonte do texto
; Funcao chamada automaticamente
; Nao recebe nenhum parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
SetColor PROC USES EAX, color: DWORD
	mov eax, color
	call SetTextColor
	ret
SetColor ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Gera a base da interface da planila
; Funcao chamada automaticamente
; Nao recebe nenhum parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
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

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Exibe o conteudo das celulas em cada uma das posicoes da tela
; Funcao chamada automaticamente
; Nao recebe nenhum parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
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
	cmp al, 1										; Inteiro
	jz _hasInt
	cmp al, 2										; Real
	jz _hasReal
	cmp al, 3										; String
	jz _hasString
	cmp al, 4										; Inteiro de funcao
	jz _hasInt
	cmp al, 5										; Real de funcao
	jz _hasReal
	mWrite "                  "
	jmp _nextCell
 _hasInt:											; Exibe o inteiro presente na celula
	mov esi, INTPOS									; Para isso atribui ao ESI o deslocamento do inicio da celula ateh o valor inteiro
	mov eax, [edx + esi]
	call WriteInt
	jmp _nextCell									; Salta para a proxima celula
 _hasReal:
 	mov esi, FLTPOS
	fld REAL4 ptr [edx + esi]
	call WriteFloat
	fstp trashReal
 	jmp _nextCell
 _hasString:										; Exibe a string presente na celula
 	add edx, STRPOS									; Para isso adiciona EDX o deslocamento do inicio da celula ateh a string
	call WriteString
	sub edx, STRPOS									; Retorna edx para o valor anterior
 _nextCell:
 	mov eax, posx
	mov ebx, posy
	cmp eax, 101									; Verifica se estah na ultima celula da linha
	jnz _keepLine
 _changeLine:										; Se estiver:
	mov eax, 6										; Move x para o inicio da linha
	inc ebx											; Salta para a proxima linha
	jmp _continue
 _keepLine:
 	add eax, 19										; Senao, salta para a proxima celula na linha atual
 _continue:	
	lea esi, posx									; Atualiza a variavel posx com a nova posicao
	mov [esi], eax
	lea esi, posy									; Atualiza a variavel posy com a nova posicao
	mov [esi], ebx
	add edx, SIZECELL
 	loop _gotoDraw
	jmp _end
 _gotoDraw:
	jmp _draw
 _end:
	ret
DrawData ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Move o cursor para uma determinada posicao, funcao gotoxy facilitada pelo uso de invoke
; Funcao chamada automaticamente
; Recebe como parametros duas DWORDs que sao, respectivamente, as posicoes x e y desejadas
; ---------------------------------------------------------------------------------------------------------------------------------------------------
GotoPosXY PROC USES EDX, posx: DWORD, posy: DWORD
	mov dl, BYTE PTR posX
	mov dh, BYTE PTR posY
	call Gotoxy
	ret
GotoPosXY ENDP

; ---------------------------------------------------------------------------------------------------------------------------------------------------
; Limpa o conteudo da planilha sem apagar a base da planilha
; Funcao chamada automaticamente
; Nao recebe nenhum parametro
; ---------------------------------------------------------------------------------------------------------------------------------------------------
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
