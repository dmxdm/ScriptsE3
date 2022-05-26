'GERADOR DE LISTA CIM-TEAM - André de Oliveira
'
'

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------

	Set e3 = CreateObject("CT.Application") 'Seta objeto para o E3
	Set Job = e3.CreateJobObject			'Seta objeto para o Projeto no E3
	Set Con = Job.CreateConnectionObject	'Seta objeto para as Conexões
	Set Wire = Job.CreatePinObject			'Seta objeto para os Fios
	Set Dev = Job.CreateDeviceObject		'Seta objeto para Dipositivo da Primeira Conexão
	Set Dev2 = Job.CreateDeviceObject		'Seta objeto para Dipositivo da Segunda Conexão
	Set Pin = Job.CreatePinObject			'Seta objeto para os Pinos dos Dipositivo da Primeira Conexão
	Set Pin2 = Job.CreatePinObject 			'Seta objeto para os Pinos dos Dipositivo da Segunda Conexão
	dim lista()								'Variavel Global para Lista em Excel
	dim listaTXT							'Variavel Global para Lista em TXT
	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------	
	
	EXT=1 ' 1 - PARA EXCEL 2 - PARA TXT
	
	'---------------------SEPARADORES LISTA DE-PARA---------------------
	SEPARADOR1 = ":" 	' SEPARADOR DESIGNAÇÃO-PINO1
	SEPARADOR2 = "/" 	' SEPARADOR DE PARA
	SEPARADOR3 = ":" 	' SEPARADOR DESIGNAÇÃO-PINO2
	SEPARADOR4 = " - " 	' SEPARADOR DE PARA-NOME DO FIO
	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------	
	
	
	'Apresenta a tela de escolha entre tipos de lista e armazena a escolha
	MODE = InputBox("INFORME A OPCAO DESEJADA" & vbCrlf & "1 - Lista de Tag" & vbCrlf & "2 - Lista De/Para-Fio")
	Select Case MODE 														'Filtra a seleção escolhida
		Case "1"															
			TagList															'Case 1, chama a lista de TAG
		Case "2"															
			DeParaList														'Case 2, chama a lista de DE/PARA
		Case ""																
			MsgBox "Saindo."												'Case cancele
		Case Else															
			MsgBox "Opcao invalida" & ", saindo."							'Case qualquer outra informação
	End Select
	
	
Sub TagList																	'Subrotina que busca e armazena as TAGs
	
	Job.GetAllDeviceIds connIds												'Pega os ID de todos Dispositivos no Projeto
	ReDim lista(UBound(connIds))											'Redimensiona o vetor de acordo com o número de Dispositivos
	
	For i = 1 to UBound(connIds)											'Percore todos Dispositivos 
		Dev.SetId connIds(i)												'Seta o Dispositivo do projeto para o objeto do script a ser lido 
		If InStr(Dev.GetName, "Fios") > 0 Then								'Se o nome do Dipositivo for Fios evauca do looping sem armazenar nada -A pasta fios é um dispositivo com pinos FIOS-
		Exit for			
		End IF	
			Select Case EXT													'Verifica a escolha por 1-Excel ou 2-Txt para armazenar corretamente
				Case 1
					lista(i) = " " & Dev.GetName							'Case 1, armazena na Varivavel Vertor Posição i o nome do Dispositivo para formato Excel do item setado
				Case 2
					listaTXT = listaTXT & Dev.GetName & vbCrlf				'Case 2, armazena na Varivavel o nome do Dispositivo e quebra linha para formato Txt do item setado
			End Select
	Next
	
	txtFileName  = Job.GetPath & Job.GetName  & ".xlsx"                     'Dá o nome para o arquvio final
			Select Case EXT													'Verifica a escolha por 1-Excel ou 2-Txt para chamada da Função de Escrita Correta
				Case 1														
					WriteFileExcel lista, txtFileName						'Case 1, chama função de escrita para Excel
				Case 2
					WriteFileTxt listaTXT, txtFileName						'Case 2, chama função de escrita para Txt
			End Select
	
End Sub
	
	
Sub DeParaList																'Subrotina que busca e armazena os De/Para
	
	Job.GetAllConnectionIds connIds											'Pega is ID de todas as Conexões do Projeto
	ReDim lista(UBound(connIds))											'Redimensiona o vetor de acordo com o número de Conexões
	
	For i = 1 to UBound(connIds)											'Percorre cada Conexão
		Con.SetId connIds(i)												'Seta uma conexão do projeto para o objeto do script a ser lido 
		Con.GetCoreIds coreIds												'Pega os ID de todos os fios dessa Conexão do Projeto
		For j = 1 to UBound(coreIds)										'Percorre cada Fio da Conexão
			Wire.SetId coreIds(j)											'Seta cada fio da conexão do projeto para o objeto do script a ser lido 
			Pin.SetId Wire.GetEndPinId (1,ret)								'Seta o Pino do Dipositivo a partir do ponto inicial de conexão do fio para o objeto do script a ser lido 
			Pin2.SetId Wire.GetEndPinId (2,ret)								'Seta o Pino final do fio no objeto do script a ser lido 
			Dev.SetId Pin.GetId												'Seta o Dispositivo conectado ao Pino incial do fio no objeto do script a ser lido 
			Dev2.SetId Pin2.GetId											'Seta o Dispositivo conectado ao Pino Final do fio no objeto do script a ser lido 
			Select Case EXT													'Verifica a escolha por 1-Excel ou 2-Txt para armazenar corretamente
				Case 1														'Case 1, armazena na Varivavel Vertor Posição i o De/Para para formato Excel dos itens setados
					lista(i) = " " & Dev.GetName & SEPARADOR1 & Pin.GetName & SEPARADOR2 & Dev2.GetName & SEPARADOR3 & Pin2.GetName & SEPARADOR4 & Wire.GetName			
				Case 2														'Case 2, armazena na Varivavel do De/Para para formato Txt dos itens setados
					listaTXT = listaTXT & Dev.GetName & SEPARADOR1 & Pin.GetName & SEPARADOR2 & Dev2.GetName & SEPARADOR3 & Pin2.GetName & SEPARADOR4 & Wire.GetName	& vbCrlf
			End Select
	Next
	Next
	txtFileName  = Job.GetPath & Job.GetName  & ".xlsx"						'Dá o nome para o arquvio final
			Select Case EXT													'Verifica a escolha por 1-Excel ou 2-Txt para chamada da Função de Escrita Correta
				Case 1
					WriteFileExcel lista, txtFileName						'Case 1, chama função de escrita para Excel
				Case 2
					WriteFileTxt listaTXT, txtFileName						'Case 2, chama função de escrita para Txt
			End Select

	
End Sub
	
Function WriteFileExcel(List, FileToSave)									'Função para criação do Excel

	Set objExcel = CreateObject("Excel.Application")						'Seta objeto para o Excel
	objExcel.Visible = False												'Mantem excel fechado enquanto escreve
	Set WorkBook = objExcel.Workbooks.Add()									'Seta objeto para WorkBook
	
	currentLine = 1															'Contator de linhas do Excel
	For line = 1 To UBound(List)											'Contador para Preecher as linhas 
		objExcel.Cells(currentLine,1).Value = List(line)					'Define o valor da celula linha,coluna com o valor da posição correponde da lista Excel recebida
		objExcel.Columns.Autofit											'Ajusta largura da coluna de acordo com o Texto
		currentLine = currentLine + 1										'Incrementa a linha
	Next
	
	objExcel.Visible = True													'Abre excel após escrever

	Set WorkBook = Nothing													'Encerrando Workbook
	Set objExcel = Nothing													'Encerrando Excel

End Function

Function WriteFileTxt(list, file)											'Função para criação do Txt

	Set ObjFSO = CreateObject("Scripting.FileSystemObject")					'Seta objeto para arquivo de script
	Set MyFile = ObjFSO.CreateTextFile(file, True)							'Seta objeto para criação de arquvio txt
	Set WSHShell = WScript.CreateObject("WScript.Shell")					'Seta objeto para conversar com comando do windows na shell-CMD-
	
	MyFile.WriteLine(list)													'Insere a lista no arquivo
	MyFile.Close															'Fecha o arquivio
	
	result = MsgBox ("Deseja abrir o arquvio?", vbYesNo + vbQuestion)		'Oferece para abrir o arquvio ao usuario

	Select Case result														'Filta escolha entre sim e não
		Case vbYes	
			WshShell.Run "notepad.exe " & file								'Caso Sim, abre o notepad
		Case vbNo
			MsgBox "Saindo." 												'Caso Não, encerra criação sem abrir
		End Select
	
End Function


