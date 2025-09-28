#include 'tbiconn.ch'
#include "rwmake.ch"
#include "protheus.ch"
#include "topconn.ch"
#INCLUDE 'FONT.CH'
#INCLUDE 'COLORS.CH'
#Include "TOTVS.ch"
#Include 'totvs.ch'
#INCLUDE "DBTREE.CH"


//???????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????
//? Fonte para criação de tabelas e campo direto no dicionario sem SQL                                                              ?                                                                                                              
//? Ao efetuar o login, o fonte realiza mostra opções de importação via XLSX                                                        ?                                                                                                              
//? O excel DEVE ser uma copia completa das query das tabelas SX2,SIX,SX3,SX6,SX7 para ter funcionamento completo                   ?
//? Cabe o usuario Tecnico decidir se vai ser inteiro ou só os customizados                                                         ?
//???????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

Static __lAdmin:= .F.
Static __cUser := ""
Static __cNome := ""
User Function XTmpDevMan()
	Local cEmp     := ""
	Local cFil     := ""
	Private oSelWnd

	SelUser(@cEmp, @cFil, oSelWnd ) //Chama o fonte SelUser que vai validar o login do usuario que acessar

	if lOk  //Se o login for admin ele chama o fonte de migração de base
		U_MigraDic()
	endif

Return
Static Function SelUser(cEmp, cFil, oWnd, lLogin)

	Static lOk	:= .T.

	Default lLogin  := .t.

	If lLogin
		lOk	:=	Login(cEmp, cFil, oWnd)  //Chama a funções de Login enviando Empresa e Filial
	EndIf

Return lOk

Static Function Login(cEmp, cFil, oWnd)
	Local oModal
	Local oFont
	Local lOk				:= .F.
	Local cUserName := Space(100)
	Local cSenha    := Space(250)
	Local cUserID   := ""

	//Montagem da tela de Login
	oFont := TFont():New('Arial',, -11, .T., .T.)

	oModal  := FWDialogModal():New()
	oModal:SetEscClose(.f.)
	oModal:setTitle("TOTVS TIDEV - Autenticação Administrativa")
	oModal:setSize(120, 140)
	oModal:createDialog()   //Chama o VldLogin() posterior ao usuario inserir as informações
	oModal:AddButton("OK", {|| lOk := VldLogin(cUserName, cSenha, @cUserID),  IF(lOk, oModal:DeActivate(), NIL)}     , "OK",,.T.,.F.,.T.,)
	oModal:AddButton("Cancelar",{|| lOk := .F., oModal:DeActivate()}     , "Cancelar",,.T.,.F.,.T.,)

	@ 010,005 Say "Usuario:" PIXEL of oModal:getPanelMain()   FONT oFont
	@ 018,005 GET oUsuario  VAR cUserName  SIZE 130, 9 OF oModal:getPanelMain() PIXEL FONT oFont

	@ 032,005 Say "Senha:" PIXEL of oModal:getPanelMain()  FONT oFont
	@ 040,005 GET oSenha VAR cSenha PASSWORD  SIZE 130, 9 OF oModal:getPanelMain() PIXEL FONT oFont

	oModal:Activate()

	If lOk
		__cUserId := cUserID
	Else
		Return
	EndIf

Return lOk
Static Function VldLogin(cUser, cSenha, cUserID)
	Local lRet    := .F.
	Local aUser   := {}
	Local nRetPsw := 0
	__lAdmin := .F.

	//Recebe usuario e senha
	cUser     := Alltrim(cUser)
	cSenha    := Alltrim(cSenha)
	nRetPsw   := PswAdmin(cUser, cSenha) //A função retorna 0 se o usuário for do grupo de admin, 1  não admin e 2 se for senha inválida.

	If nRetPsw == 0 .and. ! cUser $ "ATENDCONSULTA,SYSTEM"   //Usuario admin
		lRet     := .T.
		__lAdmin := .T.
	Elseif nRetPsw == 2  //Senha invalida
		FWAlertWarning("Senha e/ou usuario invalidos!")
		lRet:= .F.
	Elseif nRetPsw == 1 .or. cUser == "ATENDCONSULTA" .Or. cUser == "SYSTEM"  // Usuario Não admin
		FWAlertWarning('O usu rio Não   admin!')
		lRet        := .F.
	Endif

	If lRet
		aUser	:= PswRet() //Retorna vetor contendo informações do Usuário ou do Grupo de Usuários.
		cUserID := aUser[1, 1]
		__cNome := aUser[1, 4]
		__cUser := cUserID
	EndIf

Return lRet


User Function MigraDic()

	Static lSX2 := .F.
	Static lSIX := .F.
	Static lSX3 := .F.
	Static lSX6 := .F.
	Static lSX7 := .F.

	SetPrvt("oDlg1","oBSX2","oBSIX","oBSX3","oBSX6","oBSX7")

	oDlg1  := MSDialog():New( 177,297,642,1069,"Migra base",,,.F.,,,,,,.T.,,,.T. )
	oBSX2  := TButton():New( 016,139,"Cria TABELA"   ,oDlg1,{ || avisotab(oBSX2:LPROCESSING,oBSIX:LPROCESSING,oBSX3:LPROCESSING,oBSX6:LPROCESSING,oBSX7:LPROCESSING)},108,024,,,,.T.,,"",,,,.F. )
	oBSIX  := TButton():New( 053,139,"Cria INDICE"   ,oDlg1,{ || avisotab(oBSX2:LPROCESSING,oBSIX:LPROCESSING,oBSX3:LPROCESSING,oBSX6:LPROCESSING,oBSX7:LPROCESSING)},108,024,,,,.T.,,"",,,,.F. )
	oBSX3  := TButton():New( 089,139,"Cria CAMPO"    ,oDlg1,{ || avisotab(oBSX2:LPROCESSING,oBSIX:LPROCESSING,oBSX3:LPROCESSING,oBSX6:LPROCESSING,oBSX7:LPROCESSING)},108,024,,,,.T.,,"",,,,.F. )
	oBSX6  := TButton():New( 123,139,"Cria PARAMETRO",oDlg1,{ || avisotab(oBSX2:LPROCESSING,oBSIX:LPROCESSING,oBSX3:LPROCESSING,oBSX6:LPROCESSING,oBSX7:LPROCESSING)},108,024,,,,.T.,,"",,,,.F. )
	oBSX7  := TButton():New( 156,139,"Cria GATILHO"  ,oDlg1,{ || avisotab(oBSX2:LPROCESSING,oBSIX:LPROCESSING,oBSX3:LPROCESSING,oBSX6:LPROCESSING,oBSX7:LPROCESSING)},108,024,,,,.T.,,"",,,,.F. )
	oBCanc := TButton():New( 196,259,"Cancelar"      ,oDlg1,{ || oDlg1:End()},048,024,,,,.T.,,"",,,,.F. )
	oBOk   := TButton():New( 196,319,"Ok"            ,oDlg1,{ || U_ChamaTab()},048,024,,,,.T.,,"",,,,.F. )
	oDlg1:Activate(,,,.T.)

Return
Static Function avisotab(lSX2,lSIX,lSX3,lSX6,lSX7)

	if lSX2.AND.!lSIX .AND.!lSX3.AND.!lSX6.AND.!lSX7

		if fwalertYesNo("Selecione o Excel com a SX2","Cria SX2")
			CriaArrExcel(lSX2,lSIX,lSX3,lSX6,lSX7)
		else
			Return .F.
		endif

	endif
	if !lSX2.AND.lSIX .AND. !lSX3.AND.!lSX6.AND.!lSX7

		if fwalertYesNo("Selecione o Excel com a SIX","Cria SIX")
			CriaArrExcel(lSX2,lSIX,lSX3,lSX6,lSX7)
		else
			Return .F.
		endif

	endif
	if !lSX2 .AND.!lSIX .AND.lSX3.AND.!lSX6.AND.!lSX7
		if fwalertYesNo("Selecione o Excel com a SX3","Cria SX3")
			CriaArrExcel(lSX2,lSIX,lSX3,lSX6,lSX7)
		else
			Return .F.
		endif

	endif
	if !lSX2 .AND.!lSIX.AND.!lSX3.AND.lSX6.AND.!lSX7
		if fwalertYesNo("Selecione o Excel com a SX6","Cria SX6")
			CriaArrExcel(lSX2,lSIX,lSX3,lSX6,lSX7)
		else
			Return .F.
		endif

	endif
	if !lSX2 .AND.!lSIX.AND.!lSX3.AND.!lSX6.AND.lSX7
		if fwalertYesNo("Selecione o Excel com a SX7","Cria SX7")
			CriaArrExcel(lSX2,lSIX,lSX3,lSX6,lSX7)
		else
			Return .F.
		endif

	endif

	if !lSX2 .AND.!lSIX.AND.!lSX3.AND.!lSX6.AND.!lSX7
		if fwalertYesNo("Selecione o Excel com a SX2","Chama Tab")
			CriaArrExcel(lSX2,lSIX,lSX3,lSX6,lSX7)
		else
			Return .F.
		endif

	endif

return(.T.)

Static Function CriaArrExcel(lSX2,lSIX,lSX3,lSX6,lSX7)

//	RpcSetEnv("99","01")
	Local aArea     := FWGetArea()
	Local cDirIni   := GetTempPath()
	Local cTipArq   := 'Arquivos Excel (*.xlsx) | Arquivos Excel 97-2003 (*.xls)'
	Local cTitulo   := 'Seleção de Arquivos para Processamento'
	Local lSalvar   := .F.
	Local cArqSel   := ''
	Private cArqCSV := ""

	//Chama a função para buscar arquivos
	cArqSel := tFileDialog(;
		cTipArq,;  // Filtragem de tipos de arquivos que serão selecionados
	cTitulo,;  // Título da Janela para seleção dos arquivos
	,;         // Compatibilidade
	cDirIni,;  // Diretório inicial da busca de arquivos
	lSalvar,;  // Se for .T., será uma Save Dialog, senão será Open Dialog
	;          // Se não passar parâmetro, irá pegar apenas 1 arquivo; Se for informado GETF_MULTISELECT será possível pegar mais de 1 arquivo; Se for informado GETF_RETDIRECTORY será possível selecionar o diretório
	)

	//Se tiver o arquivo selecionado e ele existir
	If ! Empty(cArqSel) .And. File(cArqSel)
		//Faz a conversão de XLS para CSV
		cArqCSV := fXLStoCSV(cArqSel,lSX2,lSIX,lSX3,lSX6,lSX7)

		//Se o arquivo XLS existir
		If File(cArqCSV)
			Processa({|| fImporta(cArqCSV,lSX2,lSIX,lSX3,lSX6,lSX7) }, 'Importando...')
		EndIf
	EndIf

	FWRestArea(aArea)

return(.T.)
/*/{Protheus.doc} fImporta
Função que processa o arquivo e realiza a importação para o sistema
@author Daniel Atilio
@since 16/07/2022
@version 1.0
@type function
@obs Codigo gerado automaticamente pelo Autumn Code Maker
@see http://autumncodemaker.com
/*/

Static Function fImporta(cArqSel,lSX2,lSIX,lSX3,lSX6,lSX7)

	Local nTotLinhas := 0
	Local cLinAtu    := ''
	Local nLinhaAtu  := 0
	Local nX   := 0
	Local aLinha     := {}
	Local aSXArr     := {}
	Local oArquivo
	Private aDados         := {}
	Private lMSHelpAuto    := .T.
	Private lAutoErrNoFile := .T.
	Private lMsErroAuto    := .F.
	//Variáveis da Importação
	Private cSeparador := ','


	if lSX2 .AND.!lSIX .AND. !lSX3 .AND. !lSX6 .AND. !lSX7
		cSeparador := ','
	endif
	if !lSX2 .AND.!lSIX .AND. !lSX3 .AND. !lSX6 .AND. !lSX7
		cSeparador := ','
	endif
	if !lSX2 .AND. !lSIX .AND. lSX3 .AND. !lSX6 .AND. !lSX7
		cSeparador := '\'
	endif
	if (!lSX2 .AND. !lSIX .AND. !lSX3 .AND. lSX6 .AND. !lSX7) .OR. (!lSX2 .AND. !lSIX .AND. !lSX3 .AND. !lSX6 .AND. lSX7)
		cSeparador := '~'
	endif

	//Definindo o arquivo a ser lido
	oArquivo := FWFileReader():New(cArqSel)

	//Se o arquivo pode ser aberto
	If (oArquivo:Open())

		//Se não for fim do arquivo
		If ! (oArquivo:EoF())

			//Definindo o tamanho da régua
			aLinhas := oArquivo:GetAllLines()
			nTotLinhas := Len(aLinhas)
			ProcRegua(nTotLinhas)

			//Método GoTop não funciona (dependendo da versão da LIB), deve fechar e abrir novamente o arquivo
			oArquivo:Close()
			oArquivo := FWFileReader():New(cArqSel)
			oArquivo:Open()

			for nX := 1 to len(aLinhas)

				//Incrementa na tela a mensagem
				nLinhaAtu++
				IncProc('Analisando linha ' + cValToChar(nLinhaAtu) + ' de ' + cValToChar(nTotLinhas) + '...')

				//Pegando a linha atual e transformando em array
				cLinAtu := oArquivo:GetLine()

				aLinha  := Separa(aLinhas[nX], cSeparador)

				//Se houver posições no array
				If Len(aLinha) > 0
					aadd(aSXArr ,{aLinha})
				EndIf

			next

			if !Empty(aSXArr) .AND. Len(aSXArr) >= 1

				if validasx(aSXArr,lSX2,lSIX,lSX3,lSX6,lSX7)

					if tValSX .AND. lSX2 .AND. !lSX3.AND. !lSX6 .AND.!lSX7
						criatable(aSXArr)
					endif

					if tValSX .AND. !lSX2 .AND. lSIX .AND. !lSX3 .AND. !lSX6 .AND. !lSX7
						criaindice(aSXArr)
					endif

					if tValSX .AND. !lSX2 .AND.!lSIX .AND.lSX3.AND.!lSX6.AND.!lSX7
						criacampo(aSXArr)
					endif

					if tValSX .AND. !lSX2 .AND.!lSIX .AND. !lSX3.AND. lSX6 .AND. !lSX7
						criaparam(aSXArr)
					endif

					if tValSX .AND. !lSX2 .AND.!lSIX .AND. !lSX3 .AND. !lSX6 .AND. lSX7
						criagat(aSXArr)
					endif


					if tValSX .AND. !lSX2 .AND.!lSIX .AND. !lSX3 .AND. !lSX6 .AND. !lSX7
						CriaTB(aSXArr)
					endif

				endif

			Endif

		Else
			MsgStop('Arquivo não tem conteúdo!', 'Atenção')
		EndIf

		//Fecha o arquivo
		oArquivo:Close()
	Else
		MsgStop('Arquivo não pode ser aberto!', 'Atenção')
	EndIf

return(.T.)

//Essa função foi baseada como referência no seguinte link: https://stackoverflow.com/questions/1858195/convert-xls-to-csv-on-command-line
Static Function fXLStoCSV(cArqXLS,lSX2,lSIX,lSX3,lSX6,lSX7)
	Local cArqCSV    := ""
	Local cDirTemp   := GetTempPath()
	Local cArqScript := cDirTemp + "XlsToCsv.vbs"
	Local cScript    := ""
	Local cDrive     := ""
	Local cDiretorio := ""
	Local cNome      := ""
	Local cExtensao  := ""

	if !lSX2 .AND.!lSIX.AND. lSX3.AND.!lSX6.AND.!lSX7

		cScript := 'If WScript.Arguments.Count < 2 Then' + CRLF
		cScript += '    WScript.Quit' + CRLF
		cScript += 'End If' + CRLF
		cScript += 'Dim oExcel' + CRLF
		cScript += 'Set oExcel = CreateObject("Excel.Application")' + CRLF
		cScript += 'oExcel.DisplayAlerts = False' + CRLF
		cScript += 'Dim oBook' + CRLF
		cScript += 'Set oBook = oExcel.Workbooks.Open(WScript.Arguments.Item(0))' + CRLF
		cScript += 'oBook.SaveAs WScript.Arguments.Item(1), 6' + CRLF
		cScript += 'oBook.Close False' + CRLF
		cScript += 'oExcel.Quit' + CRLF
		cScript += 'Set oExcel = Nothing' + CRLF
		cScript += 'Dim fso, csvFile, content, processedContent' + CRLF
		cScript += 'Set fso = CreateObject("Scripting.FileSystemObject")' + CRLF
		cScript += 'Set csvFile = fso.OpenTextFile(WScript.Arguments.Item(1), 1)' + CRLF
		cScript += 'content = csvFile.ReadAll' + CRLF
		cScript += 'csvFile.Close' + CRLF
		cScript += 'Dim lines, line, newLines, quoteOpen, j, char, newLine' + CRLF
		cScript += 'lines = Split(content, vbCrLf)' + CRLF
		cScript += 'ReDim newLines(UBound(lines))' + CRLF
		cScript += 'For i = 0 To UBound(lines)' + CRLF
		cScript += '    line = lines(i)' + CRLF
		cScript += '    newLine = ""' + CRLF
		cScript += '    quoteOpen = False' + CRLF
		cScript += '    For j = 1 To Len(line)' + CRLF
		cScript += '        char = Mid(line, j, 1)' + CRLF
		cScript += '        If char = """" Then' + CRLF
		cScript += '            quoteOpen = Not quoteOpen' + CRLF
		cScript += '        End If' + CRLF
		cScript += '        If char = "," And Not quoteOpen Then' + CRLF
		cScript += '            newLine = newLine & "\"' + CRLF
		cScript += '        Else' + CRLF
		cScript += '            newLine = newLine & char' + CRLF
		cScript += '        End If' + CRLF
		cScript += '    Next' + CRLF
		cScript += '    newLines(i) = newLine' + CRLF
		cScript += 'Next' + CRLF
		cScript += 'processedContent = Join(newLines, vbCrLf)' + CRLF
		cScript += 'Set csvFile = fso.OpenTextFile(WScript.Arguments.Item(1), 2)' + CRLF
		cScript += 'csvFile.Write processedContent' + CRLF
		cScript += 'csvFile.Close' + CRLF
		MemoWrite(cArqScript, cScript)

		//Pega os detalhes do arquivo original em XLS
		SplitPath(cArqXLS, @cDrive, @cDiretorio, @cNome, @cExtensao)

		//Monta o nome do CSV, conforme os detalhes do XLS
		cArqCSV := cDrive + cDiretorio + cNome + ".csv"

		//Executa a conversão, exemplo:
		//   c:\totvs\Testes\XlsToCsv.vbs "C:\Users\danat\Downloads\tste2.xls" "C:\Users\danat\Downloads\tst2_csv.csv"
		ShellExecute("OPEN", cArqScript, ' "' + cArqXLS + '" "' + cArqCSV + '"', cDirTemp, 0 )
	Endif

	if (!lSX2 .AND. !lSIX .AND. !lSX3 .AND. lSX6 .AND. !lSX7) .OR. (!lSX2 .AND. !lSIX .AND. !lSX3 .AND. !lSX6 .AND. lSX7)
		cScript := 'If WScript.Arguments.Count < 2 Then' + CRLF
		cScript += '    WScript.Echo "Erro! Por favor especifique o caminho do arquivo fonte e destino."' + CRLF
		cScript += '    Wscript.Quit' + CRLF
		cScript += 'End If' + CRLF
		cScript += '' + CRLF
		cScript += 'Dim oExcel' + CRLF
		cScript += 'Set oExcel = CreateObject("Excel.Application")' + CRLF
		cScript += 'oExcel.DisplayAlerts = False' + CRLF
		cScript += 'Dim oBook' + CRLF
		cScript += 'Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))' + CRLF
		cScript += 'oBook.SaveAs WScript.Arguments.Item(1), 6' + CRLF
		cScript += 'oBook.Close False' + CRLF
		cScript += 'oExcel.Quit' + CRLF
		cScript += 'Set oExcel = Nothing' + CRLF
		cScript += '' + CRLF
		cScript += 'Dim fso, csvFile, content, processedContent' + CRLF
		cScript += 'Set fso = CreateObject("Scripting.FileSystemObject")' + CRLF
		cScript += 'Set csvFile = fso.OpenTextFile(WScript.Arguments.Item(1), 1)' + CRLF
		cScript += 'content = csvFile.ReadAll' + CRLF
		cScript += 'csvFile.Close' + CRLF
		cScript += '' + CRLF
		cScript += 'processedContent = ""' + CRLF
		cScript += 'Dim i, ch, inQuotes' + CRLF
		cScript += 'inQuotes = False' + CRLF
		cScript += '' + CRLF
		cScript += 'For i = 1 To Len(content)' + CRLF
		cScript += '    ch = Mid(content, i, 1)' + CRLF
		cScript += '    ' + CRLF
		cScript += '    If ch = """" Then' + CRLF
		cScript += '        If inQuotes And i < Len(content) And Mid(content, i+1, 1) = """" Then' + CRLF
		cScript += '            processedContent = processedContent & """"' + CRLF
		cScript += '            i = i + 1' + CRLF
		cScript += '        Else' + CRLF
		cScript += '            inQuotes = Not inQuotes' + CRLF
		cScript += '        End If' + CRLF
		cScript += '    ElseIf ch = "," And Not inQuotes Then' + CRLF
		cScript += '        processedContent = processedContent & "~"' + CRLF
		cScript += '    Else' + CRLF
		cScript += '        processedContent = processedContent & ch' + CRLF
		cScript += '    End If' + CRLF
		cScript += 'Next' + CRLF
		cScript += '' + CRLF
		cScript += 'Set csvFile = fso.OpenTextFile(WScript.Arguments.Item(1), 2)' + CRLF
		cScript += 'csvFile.Write processedContent' + CRLF
		cScript += 'csvFile.Close' + CRLF
		MemoWrite(cArqScript, cScript)

		//Pega os detalhes do arquivo original em XLS
		SplitPath(cArqXLS, @cDrive, @cDiretorio, @cNome, @cExtensao)

		//Monta o nome do CSV, conforme os detalhes do XLS
		cArqCSV := cDrive + cDiretorio + cNome + ".csv"

		//Executa a conversão, exemplo:
		//   c:\totvs\Testes\XlsToCsv.vbs "C:\Users\danat\Downloads\tste2.xls" "C:\Users\danat\Downloads\tst2_csv.csv"
		ShellExecute("OPEN", cArqScript, ' "' + cArqXLS + '" "' + cArqCSV + '"', cDirTemp, 0 )
	Endif

	if (lSX2 .AND.!lSIX.AND. !lSX3.AND.!lSX6.AND.!lSX7) .OR.  (!lSX2 .AND. lSIX .AND. !lSX3.AND.!lSX6.AND.!lSX7 ) .OR.  (!lSX2 .AND. !lSIX .AND. !lSX3.AND.!lSX6.AND.!lSX7 )
		//if lSX2 .AND. !lSX3.AND.!lSX6.AND.!lSX7
		//Monta o Script para converter
		cScript := 'if WScript.Arguments.Count < 2 Then' + CRLF
		cScript += '    WScript.Echo "Error! Please specify the source path and the destination. Usage: XlsToCsv SourcePath.xls Destination.csv"' + CRLF
		cScript += '    Wscript.Quit' + CRLF
		cScript += 'End If' + CRLF
		cScript += 'Dim oExcel' + CRLF
		cScript += 'Set oExcel = CreateObject("Excel.Application")' + CRLF
		cScript += 'Dim oBook' + CRLF
		cScript += 'Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))' + CRLF
		cScript += 'oBook.SaveAs WScript.Arguments.Item(1), 6' + CRLF
		cScript += 'oBook.Close False' + CRLF
		cScript += 'oExcel.Quit' + CRLF
		MemoWrite(cArqScript, cScript)

		//Pega os detalhes do arquivo original em XLS
		SplitPath(cArqXLS, @cDrive, @cDiretorio, @cNome, @cExtensao)

		//Monta o nome do CSV, conforme os detalhes do XLS
		cArqCSV := cDrive + cDiretorio + cNome + ".csv"

		//Executa a convers?o, exemplo:
		//   c:\totvs\Testes\XlsToCsv.vbs "C:\Users\danat\Downloads\tste2.xls" "C:\Users\danat\Downloads\tst2_csv.csv"
		ShellExecute("OPEN", cArqScript, ' "' + cArqXLS + '" "' + cArqCSV + '"', cDirTemp, 0 )
	endif


Return cArqCSV


Static Function criacampo(aSXArr )

	Local nX := 0
	Local nIndex := 0
	Local nReg := 1
	Local cQuery := ""
	Local lReck := .F.

// Percorre todas as posições do array aSXArr

	cQuery += "SELECT * "
	cQuery += "FROM SX3X3199 "
	USE SX3X3199 ALIAS SX3X3199 SHARED NEW VIA "TOPCONN"
	DbGotop()
	For nX := 1 to Len(aSXArr)
		If Alltrim(aSXArr[nX][1][1]) <> "SKIP"
			nReg := nX
			While !SX3X3199->(Eof()) .AND. !lReck
				nReg++
				if Alltrim(aSXArr[nX][1][3]) == alltrim(X3_CAMPO)
					lReck := .F.
					nX++
				endif
				('SX3X3199')->(dbskip())
			enddo

			if nReg > len(aSXArr)
				FWAlertWarning("Campos ja estao em chache, Salve na via CFG","OK")
				('SX3X3199')->(DbClosearea())
				Return
			endif

			if Alltrim(aSXArr[nX][1][3]) <> alltrim(X3_CAMPO) .AND. Empty(X3_CAMPO)
				lReck := .T.
			endif

			if len(aSXArr[nX][1]) < 49
				For nIndex := Len(aSXArr[nX][1]) + 1 to 49
					AADD(aSXArr[nX][1], "")
				Next
			endif

			if len(aSXArr[nX][1]) > 49
				alert("ERRO campo: "+Alltrim(aSXArr[nX][1][3]))
				RETURN
			endif

			if lReck
				RecLock('SX3X3199',.T. )
				X3_ARQUIVO := Alltrim(aSXArr[nX][1][1])
				X3_ORDEM   := Alltrim(aSXArr[nX][1][2])
				X3_CAMPO   := Alltrim(aSXArr[nX][1][3])
				X3_TIPO    := Alltrim(aSXArr[nX][1][4])
				X3_TAMANHO := Val(aSXArr[nX][1][5])
				X3_DECIMAL := Val(aSXArr[nX][1][6])
				X3_TITULO  := Alltrim(aSXArr[nX][1][7])
				X3_TITSPA  := Alltrim(aSXArr[nX][1][8])
				X3_TITENG  := Alltrim(aSXArr[nX][1][9])
				X3_DESCRIC := Alltrim(aSXArr[nX][1][10])
				X3_DESCSPA := Alltrim(aSXArr[nX][1][11])
				X3_DESCENG := Alltrim(aSXArr[nX][1][12])
				X3_PICTURE := Alltrim(aSXArr[nX][1][13])
				X3_VALID   := Alltrim(aSXArr[nX][1][14])
				X3_USADO   := Alltrim(aSXArr[nX][1][15])
				X3_RELACAO := Alltrim(aSXArr[nX][1][16])
				X3_F3      := Alltrim(aSXArr[nX][1][17])
				X3_NIVEL   := Val(aSXArr[nX][1][18])
				X3_RESERV  := Alltrim(aSXArr[nX][1][19])
				X3_CHECK   := Alltrim(aSXArr[nX][1][20])
				X3_TRIGGER := Alltrim(aSXArr[nX][1][21])
				X3_PROPRI  := Alltrim(aSXArr[nX][1][22])
				X3_BROWSE  := Alltrim(aSXArr[nX][1][23])
				X3_VISUAL  := Alltrim(aSXArr[nX][1][24])
				X3_CONTEXT := Alltrim(aSXArr[nX][1][25])
				X3_OBRIGAT := Alltrim(aSXArr[nX][1][26])
				X3_VLDUSER := Alltrim(aSXArr[nX][1][27])
				X3_CBOX    := Alltrim(aSXArr[nX][1][28])
				X3_CBOXSPA := Alltrim(aSXArr[nX][1][29])
				X3_CBOXENG := Alltrim(aSXArr[nX][1][30])
				X3_PICTVAR := Alltrim(aSXArr[nX][1][31])
				X3_WHEN    := Alltrim(aSXArr[nX][1][32])
				X3_INIBRW  := Alltrim(aSXArr[nX][1][33])
				X3_GRPSXG  := Alltrim(aSXArr[nX][1][34])
				X3_FOLDER  := Alltrim(aSXArr[nX][1][35])
				X3_PYME    := Alltrim(aSXArr[nX][1][36])
				X3_CONDSQL := Alltrim(aSXArr[nX][1][37])
				X3_CHKSQL  := Alltrim(aSXArr[nX][1][38])
				X3_IDXSRV  := Alltrim(aSXArr[nX][1][39])
				X3_ORTOGRA := Alltrim(aSXArr[nX][1][40])
				X3_IDXFLD  := Alltrim(aSXArr[nX][1][41])
				X3_TELA    := Alltrim(aSXArr[nX][1][42])
				X3_PICBRV  := Alltrim(aSXArr[nX][1][43])
				X3_AGRUP   := Alltrim(aSXArr[nX][1][44])
				X3_POSLGT  := Alltrim(aSXArr[nX][1][45])
				X3_MODAL   := Alltrim(aSXArr[nX][1][46])
				MsUnlock()
			EndIf
		EndIf
	NEXT

	('SX3X3199')->(DbClosearea())
	FWAlertInfo("Salve na via CFG","OK")

return(.T.)
Static Function criatable(aSXArr)

	Local nX := 0

	DbSelectArea("SX2")
	DbSetOrder(1)
	DbGotop()
	For nX := 1 to Len(aSXArr)
		If !ExisteSX2(alltrim(aSXArr[nX][1][1]))

			RecLock('SX2',.T. )
			X2_CHAVE     := alltrim(aSXArr[nX][1][1])
			X2_PATH      := alltrim(aSXArr[nX][1][2])
			X2_ARQUIVO   := alltrim(aSXArr[nX][1][3])
			X2_NOME      := alltrim(aSXArr[nX][1][4])
			X2_NOMESPA   := alltrim(aSXArr[nX][1][5])
			X2_NOMEENG   := alltrim(aSXArr[nX][1][6])
			X2_ROTINA    := alltrim(aSXArr[nX][1][7])
			X2_MODO      := alltrim(aSXArr[nX][1][8])
			X2_MODOUN    := alltrim(aSXArr[nX][1][9])
			X2_MODOEMP   := alltrim(aSXArr[nX][1][10])
			X2_DELET     := Val(aSXArr[nX][1][11])
			X2_TTS       := alltrim(aSXArr[nX][1][12])
			X2_UNICO     := alltrim(aSXArr[nX][1][13])
			X2_PYME      := alltrim(aSXArr[nX][1][14])
			X2_MODULO    := Val(aSXArr[nX][1][15])
			X2_DISPLAY   := alltrim(aSXArr[nX][1][16])
			X2_SYSOBJ    := alltrim(aSXArr[nX][1][17])
			X2_USROBJ    := alltrim(aSXArr[nX][1][18])
			X2_POSLGT    := alltrim(aSXArr[nX][1][19])
			X2_CLOB      := alltrim(aSXArr[nX][1][20])
			X2_AUTREC    := alltrim(aSXArr[nX][1][21])
			X2_TAMFIL    := Val(aSXArr[nX][1][22])
			X2_TAMUN     := Val(aSXArr[nX][1][23])
			X2_TAMEMP    := Val(aSXArr[nX][1][24])
			X2_STAMP     := alltrim(aSXArr[nX][1][28])
			MsUnlock()

		Else
			FWAlertError("Tabela "+alltrim(aSXArr[nX][1][1])+" encontrada, não criará ela novamente", "Tabela inexistente")
		EndIf
	NEXT

	('SX2')->(DbClosearea())
	FWAlertInfo("Salve na via CFG","OK")

return(.T.)

Static Function criaindice(aSXArr )

	Local nX := 0

	DbSelectArea("SIX")
	DbSetOrder(1)
	DbGotop()

	For nX := 1 to Len(aSXArr)

		if !dbseek(Alltrim(aSXArr[nX][1][1])+Alltrim(aSXArr[nX][1][2]))
			RecLock('SIX',.T. )
			INDICE    := Alltrim(aSXArr[nX][1][1])
			ORDEM     := Alltrim(aSXArr[nX][1][2])
			CHAVE     := Alltrim(aSXArr[nX][1][3])
			DESCICAO  := Alltrim(aSXArr[nX][1][4])
			DESCSPA   := Alltrim(aSXArr[nX][1][5])
			DESCENG   := Alltrim(aSXArr[nX][1][6])
			PROPRI    := Alltrim(aSXArr[nX][1][7])
			F3        := Alltrim(aSXArr[nX][1][8])
			NICKNAME  := Alltrim(aSXArr[nX][1][9])
			SHOWPESQ  := Alltrim(aSXArr[nX][1][10])
			IX_VIRTUAL:= Alltrim(aSXArr[nX][1][11])
			IX_VIRCUST:= Alltrim(aSXArr[nX][1][12])
			MsUnlock()
		endif
	NEXT

	('SIX')->(DbClosearea())
	FWAlertInfo("Salve na via CFG","OK")

return(.T.)
Static Function criaparam(aSXArr )

	Local nX := 0
	Local cVar := ""
	Local cX6Fil := ""

	DbSelectArea("SX6")
	DbSetOrder(1)
	For nX := 1 to Len(aSXArr)

		if Empty(alltrim(aSXArr[nX][1][1]))
			cX6Fil := "  "
		else
			cX6Fil := alltrim(aSXArr[nX][1][1])
		endif

		IF !DbSeek(cX6Fil +alltrim(aSXArr[nX][1][2]))
			RecLock('SX6',.T. )
			X6_FIL          := Alltrim(aSXArr[nX][1][1])
			X6_VAR	        := Alltrim(aSXArr[nX][1][2])
			X6_TIPO	        := Alltrim(aSXArr[nX][1][3])
			X6_DESCRIC    	:= Alltrim(aSXArr[nX][1][4])
			X6_DSCSPA  	    := Alltrim(aSXArr[nX][1][5])
			X6_DSCENG  	    := Alltrim(aSXArr[nX][1][6])
			X6_DESC1   	    := Alltrim(aSXArr[nX][1][7])
			X6_DSCSPA1   	:= Alltrim(aSXArr[nX][1][8])
			X6_DSCENG1    	:= Alltrim(aSXArr[nX][1][9])
			X6_DESC2        := Alltrim(aSXArr[nX][1][10])
			X6_DSCSPA2    	:= Alltrim(aSXArr[nX][1][11])
			X6_DSCENG2    	:= Alltrim(aSXArr[nX][1][12])
			X6_CONTEUD    	:= Alltrim(aSXArr[nX][1][13])
			X6_CONTSPA     	:= Alltrim(aSXArr[nX][1][14])
			X6_CONTENG      := Alltrim(aSXArr[nX][1][15])
			X6_PROPRI	    := Alltrim(aSXArr[nX][1][16])
			X6_PYME	        := Alltrim(aSXArr[nX][1][17])
			X6_VALID        := Alltrim(aSXArr[nX][1][18])
			X6_INIT	        := Alltrim(aSXArr[nX][1][19])
			X6_DEFPOR       := Alltrim(aSXArr[nX][1][20])
			X6_DEFSPA       := Alltrim(aSXArr[nX][1][21])
			X6_DEFENG	    := Alltrim(aSXArr[nX][1][22])
			X6_EXPDEST     	:= Alltrim(aSXArr[nX][1][23])
			//X6_ACTIVE       := Alltrim(aSXArr[nX][1][27])
			MsUnlock()
		else
			cVar += Alltrim(aSXArr[nX][1][2])
		endif
	NEXT

	if !Empty(cVar)
		Alert("Os segunites parametros ja estão em sistema: "+ cVar)
	endif

	('SX6')->(DbClosearea())
	FWAlertInfo("Salve na via CFG","OK")

return(.T.)

Static Function criagat(aSXArr )

	Local nX := 0
	Local cCampo := ""

	DbSelectArea("SX7")
	DbSetOrder(1)
	For nX := 1 to Len(aSXArr)
		IF !DbSeek(alltrim(aSXArr[nX][1][1])+alltrim(aSXArr[nX][1][2]))
			RecLock('SX7',.T. )
			X7_CAMPO    := Alltrim(aSXArr[nX][1][1])
			X7_SEQUENC  := Alltrim(aSXArr[nX][1][2])
			X7_REGRA    := Alltrim(aSXArr[nX][1][3])
			X7_CDOMIN   := Alltrim(aSXArr[nX][1][4])
			X7_TIPO	    := Alltrim(aSXArr[nX][1][5])
			X7_SEEK	    := Alltrim(aSXArr[nX][1][6])
			X7_ALIAS	:= Alltrim(aSXArr[nX][1][7])
			X7_ORDEM	:= Val(aSXArr[nX][1][8])
			X7_CHAVE	:= Alltrim(aSXArr[nX][1][9])
			X7_CONDIC	:= Alltrim(aSXArr[nX][1][10])
			X7_PROPRI	:= Alltrim(aSXArr[nX][1][11])
			MsUnlock()
		else
			cCampo += Alltrim(aSXArr[nX][1][1])
		Endif
	NEXT

	if !Empty(cCampo)
		Alert("Os segunites gatilhos ja estão em sistema para os campos: "+cCampo)
	endif

	('SX7')->(DbClosearea())
	FWAlertInfo("Salve na via CFG","OK")

return(.T.)

Static Function validasx(aSXArr,lSX2,lSIX,lSX3,lSX6,lSX7)

	Local nX := 0
	Local aValSX := {}
	Local cExistSX := ""
	Local cNExistSX := ""
	Static tValSX := .F.

	if lSX2
		For nX := 1 to Len(aSXArr)
			aAdd(aValSX, alltrim(aSXArr[nX][1][3]))

			DbSelectArea("SX2")
			DbSetOrder(1)
			//Faz a validação se os campos existem
			If dbseek(aValSX[1])
				if nX <> Len(aSXArr)
					cExistSX += aValSX[1] + "; "
				else
					cExistSX += aValSX[1]
				endif
			else
				tValSX := .T.
			endif

			aValSX := {}

			("SX2")->(DbClosearea())

		next
	endif

	if lSIX
		For nX := 1 to Len(aSXArr)
			aAdd(aValSX, alltrim(aSXArr[nX][1][1]))
			aAdd(aValSX, alltrim(aSXArr[nX][1][2]))

			DbSelectArea("SIX")
			DbSetOrder(1)
			//Faz a validação se os campos existem
 			If dbseek(aValSX[1]+aValSX[2])
				if nX <> Len(aSXArr)
					cExistSX += aValSX[1] + "; "
				else
					cExistSX += aValSX[1]
				endif
			else
				tValSX := .T.
			endif

			aValSX := {}

			("SIX")->(DbClosearea())

		next
	endif

	//Adiciona os campos que serão verificados
	if lSX3
		For nX := 1 to Len(aSXArr)
			aAdd(aValSX, alltrim(aSXArr[nX][1][3]))

			DbSelectArea("SX3")
			DbSetOrder(2)
			//Faz a validação se os campos existem
			If dbseek(aValSX[1])
				if nX <> Len(aSXArr)
					cExistSX += aValSX[1] + "; "
					adel(aSXArr[nx],1)
					aSize(aSXArr[nx], 0)
					aadd(aSXArr[nx],{"SKIP"})
				else
					cExistSX += aValSX[1]
					adel(aSXArr[nx],1)
					aSize(aSXArr[nx], 0)
					aadd(aSXArr[nx],{"SKIP"})
				endif
			else
				tValSX := .T.
			endif

			aValSX := {}

			("SX3")->(DbClosearea())

		next
	endif

	if lSX6
		For nX := 1 to Len(aSXArr)

			if empty(alltrim(aSXArr[nX][1][1]))
				cX6FIL := "  "
			else
				cX6FIL := alltrim(aSXArr[nX][1][1])
			endif

			DbSelectArea("SX6")
			DbSetOrder(1)
			If DbSeek(cX6FIL+alltrim(aSXArr[nX][1][2]))
				if nX <> Len(aSXArr)
					cExistSX += alltrim(aSXArr[nX][1][2])+ "; "
				else
					cExistSX += alltrim(aSXArr[nX][1][2])
				endif
			else
				tValSX := .T.
			endif

			("SX6")->(DbClosearea())

		next
	endif

	if lSX7
		For nX := 1 to Len(aSXArr)
			DbSelectArea("SX7")
			DbSetOrder(1)
			//Faz a validação se os campos existem
			If DbSeek(alltrim(aSXArr[nX][1][1])+alltrim(aSXArr[nX][1][2]))
				if nX <> Len(aSXArr)
					cExistSX += alltrim(aSXArr[nX][1][2]) + "; "
				else
					cExistSX += alltrim(aSXArr[nX][1][2])
				endif
			else
				tValSX := .T.
			endif

			("SX7")->(DbClosearea())

		next
	endif

	if !lSX2.AND.!lSIX.AND.!lSX3.AND.!lSX6.AND.!lSX7
		For nX := 1 to Len(aSXArr)
			aAdd(aValSX, alltrim(aSXArr[nX][1][3]))

			DbSelectArea("SX2")
			DbSetOrder(1)
			//Faz a validação se os campos existem
			If !dbseek(aValSX[1])
				if nX <> Len(aSXArr)
					cNExistSX += aValSX[1] + "; "
				else
					cNExistSX += aValSX[1]
				endif
			else
				tValSX := .T.
			endif

			aValSX := {}

			("SX2")->(DbClosearea())

		next

		if !Empty(cNExistSX)
			fwalertwaring("Os seguintes dados ja existem: "+cNExistSX,"Alerta")
		endif


	endif

	if !Empty(cExistSX)
 		fwalertwaring("Os seguintes dados ja existem: "+cExistSX,"Alerta")
	endif

Return tValSX

user function chamatab()

	CFGX031()

	fwalertwaring("Precisamos fazer a chamada das tabelas customizadas para serem criadas no dicionario, selecione o EXCEL da SX2 para seguir o processo", "AVISO!")

	if fwalertYesNo("Selecione o Excel com a SX2","Chama Tabela")
		CriaArrExcel(lSX2,lSIX,lSX3,lSX6,lSX7)
	else
		Return .F.
	endif

return(.T.)
Static function CriaTB(aSXArr)
	local ll := 1

	For ll:=1 to len(aSXArr)
		CHKFILE(alltrim(aSXArr[ll][1][1]))
		DbSelectarea(alltrim(aSXArr[ll][1][1]))
		DbGotop()
		FWAlertInfo("Tabela " + alltrim(aSXArr[ll][1][1]) + ' OK !',"Aviso!")
	Next local

	FWAlertSuccess("Fim do Processo!", "Sucesso")
return(.T.)

