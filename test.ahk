#SingleInstance force

Esc::ExitApp

; Ajout automatique 962
^l::
send, {Right 4}
Loop
{
	WinGetText, contentField, A 
	FoundPos := RegExMatch(contentField, "s)\<(.*962.*)$", SubPattContentField) ; ajouter saut de ligne à la regex
	MsgBox, SubPattContentField1 = %SubPattContentField1%
	If FoundPos != 0
	{
		MsgBox, Il existe un champs 962
		send, {Tab 3}
		WinGetText, 962Content, A
		FoundPos := RegExMatch(962Content, "s)\$a\s(.*)$", SubPatt962Content) ; !!! vérifier si marche; si s) ne fonctionne pas utiliser stringReplace; extrait les sigles de bibliothèques
		FoundPos := RegExMatch(SubPatt962Content1, "s)^.*(nebpun).*$")
		If FoundPos != 0
		{
			msgbox, nebpun est déjà en 962
			Break
		}
		Else
		{
			msgbox, 962 existe mais pas de nebpun
			SubPatt962Content1 =  %SubPatt962Content1% nebpun ; ajout "nebpun" aux autres sigles
			Sort, SubPatt962Content1, CL D%A_SPACE% ; trie par ordre alphabétiques des sigles
			send, ^+{Right 50}{Delete} ; supprimer le contenu de la zone 962 ; A améliorer
			send, %SubPatt962Content1% ; insertion des sigles
			Break
		}
	}
	Else if SubPattContentField1 > 962 ; préciser pour ne pas confondre avec le contenu de la notice
	{
		send, ^+a904{enter}
		send, nebpun
		Break
	}
	Else 
	{
		MsgBox, Cette zone n'est pas une 962 continu à chercher
		send, {Tab}
	} 
}    
return 

^u::
;récupère le numéro de commande depuis l'Etat d'acqu.
	send, {Ctrl Down}{Tab 2}{Ctrl Up} ; A faire
;loop
;{
	send, {Tab 3}
	send, {Enter}
	;WinGetTitle, AcquTitle, A
	;msgbox, AcquTitle = %AcquTitle%
	sleep, 1000
	send, {Enter}
	send, {LShift Down}{Tab 2}{LShift Up}{Down}
	sleep, 5000
	IfWinExist, ahk_class TAcqzPO_HdrForm
	{
		send, {Tab 2}{Enter}
		WinGetText, NumAcq, A
		msgbox, NumAcq = %NumAcq%
		FileAppend, %NumAcq%, C:\Users\mottetj\Desktop\sortietest.txt
		FoundPos := RegExMatch(NumAcq, "s)^.*Etat acqu\.`r`n([1-9]\s[1-9])`r`nStatut commande:.*", SubPattNumAcq) ; marche pas
		msgbox, SubPattNumAcq1 = %SubPattNumAcq1%
	}
	If Else
	{
		
		;msgbox, AcqWindow = %AcqWindow%
		;si 2 à 2 tour de boucle de suite acqwindow = acqwindow alors arrêter la boucle
		;Break
	}
	Else
	{
		;continu la boucle
	}
	
	;WinGetTitle, AcqWindow, 
	;msgbox, AcqWinTitle = %AcqWinTitle%
		
	;Sleep, 1000
	;WinGetTitle, AcqWindow, A
	;msgbox, AcqWindow = %AcqWindow%
	;If AcqWindow = Messages - catalogue RERO
	;{
	;	MsgBox, ok coucou
	;	WinActivate
	;	send, {Enter} 
	;	send, {LShift}{Tab 2}{Down}
	;}
	;Else
	;{	
	;}
;}
return

^o::
WinGetTitle, AcqWindow, A



If AcqWindow = Messages - catalogue RERO
{
	msgbox, ca marche
}
Else
{
	msgbox, marche pas ou mauvaise fenêtre
}
	WinGetClass, class, A
	msgbox, class = %class%
return


^6::
	;send, {Tab 2}{Enter}
	WinGetText, NumAcq, A
	msgbox, NumAcq = %NumAcq%
	FileAppend, %NumAcq%, C:\Users\mottetj\Desktop\sortietest.txt
	FoundPos := RegExMatch(NumAcq, "s)^.*Etat acqu\.`r`n([1-9]\s[1-9])`r`nStatut commande:.*", SubPattNumAcq) ; marche pas
	msgbox, SubPattNumAcq1 = %SubPattNumAcq1%
	
return


;Etat acqu.
;461806 - 18
;461806 - 18
;Statut commande:





