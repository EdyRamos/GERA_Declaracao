Gui, Show, w430 h430, Declaração de Atendimento
CreateGui()


ShowWaitMessage() {
    Gui, WaitMsg: New, +AlwaysOnTop
    Gui, WaitMsg: Add, Text, x20 y20, Aguarde, a declaração está sendo impressa...
    Gui, WaitMsg: Show, NoActivate
}

CloseWaitMessage() {
    Gui, WaitMsg: Destroy
}

Vertur:
Gui, Submit, NoHide
    turno := ""
    If (Matut = 1) {
        turno .= "Matutino "
    }
    If (Vesp = 1) {
        turno .= "Vespertino"
    }
Return

Ativar:
Gui, Submit, NoHide
    If (RadioGroup = 1) {
        EnableFields("Paciente")
    }
    Else If (RadioGroup = 2) {
        EnableFields("Acompanhante")
    }
Return

Check:
Gui, Submit, NoHide
    If (RadioGroup = 1) {
        ProcessDeclaration("Paciente", turno)
    }
    Else If (RadioGroup = 2) {
        ProcessDeclaration("Acompanhante", turno)
    }
Return

EnableFields(type) {
    GuiControl, Disable, NomePCT
    GuiControl, Disable, NomeAC
    GuiControl, Disable, Parent
	GuiControl, Disable, Matut
	GuiControl, Disable, Vesp
	If (type = "Paciente") {
	GuiControl, Enable, NomePCT
	GuiControl, Enable, Matut
	GuiControl, Enable, Vesp
	}
	Else If (type = "Acompanhante") {
	GuiControl, Enable, NomePCT
	GuiControl, Enable, NomeAC
	GuiControl, Enable, Parent
	GuiControl, Enable, Matut
	GuiControl, Enable, Vesp
	}
	}
	
	ProcessDeclaration(type, turno) {
	ShowWaitMessage()
	FilePath := A_ScriptDir "\Dec_" type ".dotm"
	wdApp := ComObjCreate("Word.Application") ; Cria uma instância do Word
	MsgBox, [ ,, o valor de: %FilePath%, Timeout]
	wdApp.Visible := false
	ComObjConnect(wdApp, "wd_")
	MyDocNew := wdApp.Documents.Add(FilePath) ; Abre o modelo
	FormatTime, MyTime,, d 'de' MMMM 'de' yyyy
	FormatTime, Datap,, dMMyy
	MyDocNew.Bookmarks("DataCT").Range.Text := MyTime
	MyDocNew.Bookmarks("Turno").Range.Text := turno
	MyDocNew.Bookmarks("NomePCT").Range.Text := NomePCT
	MyDocNew.Bookmarks("DataTime").Range.Text := "Londrina, " MyTime
	If (type = "Acompanhante") {
		MyDocNew.Bookmarks("NomeAC").Range.Text := NomeAC
		MyDocNew.Bookmarks("Parent").Range.Text := Parent
	}
	
	; Imprimir, salvar cópia e fechar
	MyDocNew.PrintOut() ; Imprime o documento
	pathBkp := "Bkp_atend/"
	NomeTR := SubStr(NomePCT, 1, 20)
	MyDocNew.SaveAs(A_ScriptDir "/" (pathbkp) (NomeTR) "_" (Datap)) ; Salva uma cópia de backup
	
	; Fechar a caixa de mensagem "Aguarde"
	CloseWaitMessage()
	
	; Limpar os campos do formulário
	GuiControl, 1:, NomePCT
	GuiControl, 1:, NomeAC
	GuiControl, 1:, Parent
	GuiControl, , Matut, 0
	GuiControl, , Vesp, 0
	MyDocNew.Close(0) ; Fecha o documento sem salvar
	wdApp.Quit() ; Fecha o Word
}

CreateGui() {
	global NomePCT
	global NomeAC
	global Parent
	global Matut
	global Vesp
	global RadioGroup
	global turno

	Gui, Show, w430 h430, Declaração de Atendimento
	Gui, Add, Text , x60 y15 center, SISTEMA DE DECLARAÇÃO DE ATENDIMENTO MANUAL
	Gui, Add, Text , x10 y40 +center, Por favor, selecione o tipo de declaração para ativar os campos de preenchimento!
	Gui, add, groupbox, x10 y250 w175, Selecione turno de atendimento:
	Gui, add, checkbox, x15 y270 vMatut gVertur, Matutino
	Gui, add, checkbox, x15 y290 vVesp gVertur, Vespertino
	Gui, Add, Radio, x10 y60 vRadioGroup gAtivar, Paciente
	Gui, Add, Radio, x10 y80 gAtivar, Acompanhante
	Gui, Add, Text , x10 y100, Nome Paciente:
	Gui, add, edit, x10 y120 w350 vNomePCT ; VEJA QUE IDENTIFICAMOS A VARIÁVEL "NomePCT" COMO ATRELADA AO CAMPO.
	Gui, Add, Text, x10 y150, Nome Acompanhante:
	Gui, add, edit, x10 y170 w350 vNomeAC
	Gui, Add, Text, x10 y200, Parentesco:
	Gui, add, edit, x10 y220 vParent
	Gui, Add, Button, x120 y360 w185 h25 gCheck, IMPRIMIR ; AQUI O BOTÃO IMPRIMIR FICA ATRELADO À ROTINA VERIFICA_SENHA (OPÇÃO g).
	GuiControl, Disable, NomePCT
	GuiControl, Disable, NomeAC
	GuiControl, Disable, Parent
	GuiControl, Disable, Matut
	GuiControl, Disable, Vesp
}