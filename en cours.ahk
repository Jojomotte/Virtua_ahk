#SingleInstance force
Esc::ExitApp


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; en cours

;Ouverture automatique des programmes au d�marrage est pr�paration de Virtua
^Esc::
Run, virtua.exe, C:\Program Files (x86)\Vtls\Virtua 16.0.0.0, max
send, {enter}
;send, AppsKey{Down 2}{enter}AppsKey{Down 3}{enter}AppsKey{Up 2}{enter}AppsKey{Up}{enter}
send, {LWin Down}{Left}
;Run, OUTLOOK.EXE, C:\Program Files (x86)\Microsoft Office\Office14
;Run, firefox.exe, C:\Program Files (x86)\Mozilla Firefox
return


;///////////////////////////////////////////////////////////////////////////

;CODE OK ajout de 994 automatique sauf num�ro de commande
^r::
;R�cup�re la date courante
FormatTime, year, , yyyy

;r�cup�re la cote depuis la notice d'exemplaire
WinGetText, exemplareData, A
FoundPos := RegExMatch(exemplareData, "s)^.*Cote`r`n[1-9]`r`n(.*)`r`nInformations � propos de l'exemplaire", SubPattExemplareData)
;MsgBox, FoundPos=%FoundPos%, SubPattExemplareData=%SubPattExemplareData1%

;revenir � la notice bibliographique
;send, {Alt Down}{e}{n}{Alt Up}

;r�cup�re le num�ro de commande depuis l'Etat d'acqu.
	;send, {Ctrl}{Tab 2} ; A faire
;loop
;{
;	send, {tab 3}{Enter} 
;	if ;wintitle = messages catalogue rero 
;	{
;		send, {Enter} 
;		send, {LShift}{Tab 2}{Down}
;	}
;	ifelse ;wintitle = voir commande
;	{
;		send, {Tab 2}{Enter}
;		WinGetText, orderNum, A
;		;orderNum = Etat acqu.215869 - 1
;		FoundPos := RegExMatch(orderNum, "^.*Etat acqu\.([0-9]{6} -[0-9]{1-3})+++++++++++retour charriot++++++++++++$", SubPattOrderNum)
;		MsgBox, SubPattOrderNum= %SubPattOrderNum%
;		Break		
;	}
;}
;return

;Editer la notice en MARC
;send, {Ctrl Down}{Shift Down}{Tab}
;send, {Tab 3}{Enter}

;d�termine la modalit� d'acquisition
Gui, Add, Text,, D�terminez la modalit� d'acquisition:
gui, font, s11, Arial
Gui, Add, Radio, vRadioChecked, a (achat)
Gui, Add, Radio, , d (don)
Gui, Add, Radio, , as (achat syst�matique)
Gui, Add, Radio, , w (inconnu ou r�vidion)
Gui, Add, Radio, , e (�change)
Gui, Add, Radio, , s (th�ses suisse ou autre �crit acad�mique d�pos� par l'Universit� de Neuch�tel � la BPUN(cotes TS, TSG, et 4Y))
Gui, Add, Radio, , bslc (soci�t� de lectures contemporaines)
Gui, Add, Radio, , rott (biblioth�que Rott (8R))
Gui, Add, Radio, , sgeo (Soci�t� neuch�teloise de g�ographie)
Gui, Add, Radio, , shan (Soci�t� neuch�teloise d'arch�ologie du canton de Neuch�tel)
Gui, Add, Radio, , snat (Soci�t� neuch�teloise des sciences naturelles)
Gui, Add, Radio, , ssch (Soci�t� suisse de chronom�trie)
Gui, Add, Button, , OK
Gui, Add, Button, xp+60 vCancelButton gCancelSelected, Cancel
Gui, Show
Return

; Action lorsque OK est activ�
ButtonOK:
  Gui, Submit 
  If (RadioChecked = 1)
  {
	If %numOrder%
	{
		send, ^+a994{enter}
		send, %SubPattExemplareData1% $x ne/bpun/a/%year%/%numOrder%
		Gui, Destroy 
	}
	Else
	{
		send, ^+a994{enter}
		send, %SubPattExemplareData1% $x ne/bpun/a/%year%
		Gui, Destroy 
	}
  }
  Else If (RadioChecked = 2)
  {
	InputBox, numDonation, Num�ro de donation, Veuillez entrez le num�ro de don:, ,240, 130	; a am�liorer comportement si zone de texte n'est pas frempli pas l'user
	if ErrorLevel
	{
		send, ^+a994{enter}
		send, %SubPattExemplareData1% $x ne/bpun/d/%year%
		Gui, Destroy 
	}
	Else
	{
		send, ^+a994{enter}
		send, %SubPattExemplareData1% $x ne/bpun/d/%year%/%numDonation%
		Gui, Destroy 
	}
  }
  Else If (RadioChecked = 3)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/as/%year% ; � v�rifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 4)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/w/%year% ; � v�rifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 5)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/e/%year% ; � v�rifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 6)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/s/%year% ; � v�rifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 7)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/bslc/%year% ; � v�rifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 8)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/rott/x/%year% ; � v�rifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 9)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/sgeo/x/%year% ; � v�rifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 10)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/shan/x/%year% ; � v�rifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 11)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/snat/x/%year% ; � v�rifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 12)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/ssch/x/%year% ; � v�rifier si juste
	Gui, Destroy 
  }
Return

; Action lorsque Cancel est activ�
CancelSelected:
GuiClose:
  ExitApp
Return
Return







;////////////////////////////////////////////////////


; Ajout automatique 962
^l::
send, {Right 4}
Loop
{
	WinGetText, contentField, A 
	FoundPos := RegExMatch(contentField, "s)\<(.*962.*)$", SubPattContentField) ; ajouter saut de ligne � la regex
	MsgBox, SubPattContentField1 = %SubPattContentField1%
	If FoundPos != 0
	{
		MsgBox, Il existe un champs 962
		send, {Tab 3}
		WinGetText, 962Content, A
		FoundPos := RegExMatch(962Content, "s)\$a\s(.*)$", SubPatt962Content) ; !!! v�rifier si marche; si s) ne fonctionne pas utiliser stringReplace; extrait les sigles de biblioth�ques
		FoundPos := RegExMatch(SubPatt962Content1, "s)^.*(nebpun).*$")
		If FoundPos != 0
		{
			msgbox, nebpun est d�j� en 962
			Break
		}
		Else
		{
			msgbox, 962 existe mais pas de nebpun
			SubPatt962Content1 =  %SubPatt962Content1% nebpun ; ajout "nebpun" aux autres sigles
			Sort, SubPatt962Content1, CL D%A_SPACE% ; trie par ordre alphab�tiques des sigles
			send, ^+{Right 50}{Delete} ; supprimer le contenu de la zone 962 ; A am�liorer en comptant le nombre de mot dans le champs
			send, %SubPatt962Content1% ; insertion des sigles
			Break
		}
	}
	Else if SubPattContentField1 > 962 ; pr�ciser pour ne pas confondre avec le contenu de la notice
	{
		send, ^+a904{enter}
		send, nebpun
		Break
	}
	Else 
	{
		MsgBox, Cette zone n'est pas une 962 continu � chercher
		send, {Tab}
	} 
}    
return 


;///////////////////////////////////////////////////////////////////////////////////////////////

; Propose une liste de sujet 904 � choisir
^6::
list = 1. G�n�ralit�s, dictionnaires, encyclop�dies|2. Biblioth�ques, lecture, livre ancien, manuscrits, affiches, mus�ologie, archives, �critures, iconographie, typographie, �dition, bande dessin�e|3. Edition, presse, journalisme|4. Religions, spiritualit�, mythologie, �sot�risme, mystique, �tudes bibliques|5. Philosophie, sciences humaines, �thique|6. Psychologie, psychanalyse, p�dagogie, probl�mes scolaires, toxicomanie, neurosciences|7. Droit, police, assurances, banques, administration publique|8. Politique, relations internationales, g�opolitique, �conomie, marketing, gestion d�entreprise, comptabilit�|9. Soci�t�, sociologie, probl�mes sociaux, culture, statistiques|10. Linguistique, philologie, m�thodes de langue|11. Histoire et critique litt�raire, interpr�tation de textes|12. Litt�rature fran�aise (livres en fran�ais ou traduits du fran�ais et ouvrages g�n�raux)|13. Litt�rature allemande (livres en allemand ou traduits de l'allemand et ouvrages g�n�raux)|14. Litt�rature anglaise (livres en anglais ou traduits de l'anglais et ouvrages g�n�raux)|15. Litt�rature italienne (livres en italien ou traduits de l'italien et ouvrages g�n�raux)|16. Litt�rature classique (textes grecs et latins), philologie classique|17. Autres litt�ratures (espagnole, japonaise, chinoise, russe, etc.)|18. Pr�histoire, arch�ologie|19. Histoire par grandes p�riodes, par pays ou par th�me, sciences auxiliaires de l�histoire|20. Histoire de l�Antiquit� et du Moyen Age|21. Histoire contemporaine (20e si�cle, 21e si�cle)|22. Biographies|23. Beaux-arts, peinture, sculpture, dessin, arts d�coratifs, histoire de l�art|24. Architecture, urbanisme, sites et monuments|25. Musique, chant, chanson, op�ra|26. Cin�ma, t�l�vision, radio, jeux vid�o|27. Spectacles, danse, mode|28. Photographie|29. Th��tre|30. G�ographie, g�ologie, paysage, sciences de la Terre|31. Anthropologie, ethnologie, traditions populaires, coutumes|32. Voyages, cartes, atlas, guides de voyage, tourisme, alpinisme|33. Sciences exactes, physique, chimie, math�matiques, astronomie, cosmologie|34. Sciences naturelles, zoologie, botanique, biologie, g�n�tique, sciences de la vie|35. Ecologie, climat, environnement, ressources naturelles|36. Technique, industrie, m�tiers, transports, navigation, construction et b�timent, bricolage|37. Horlogerie, chronom�trie|38. Agriculture, viticulture, jardinage, �levage, chasse, p�che, apiculture|39. M�decine, pharmacie, professions m�dicales, sant�, histoire de la m�decine|40. Sports, loisirs, jeux|41. Famille, couple, sexualit�|42. Gastronomie|43. Informatique, Internet, courrier �lectronique, smartphones, tablettes num�riques|44. Neocomensia (auteurs et sujets neuch�telois), histoire locale neuch�teloise
	Gui, font, s11, Verdana
	Gui, Add, Edit, w200 x5 vName gAutoComplete, 
	Gui, Add, ListBox, w1060 x5 vItemChoice r44, %list%
	;Gui, Add, ListBox, w1060 x5 vItemMove r44
	Gui, Add, Button, vOKButton gOKSelected Default, OK
	Gui, Add, Button, xp+60 vCancelButton gCancelSelected, Cancel
	Gui, Add, Button, vRightSelected gRightSelected,->
	Gui, Add, Button, vLeftSelected gLeftSelected,<-
	Gui, Show
	
return

AutoComplete:
	Gui, submit, nohide
	loop, parse, list, | ; parse the list to see if the name is in it
	{
		if A_LoopField contains %name%
			newlist .= A_LoopField . "|" ; populate the new list
	}
	if newlist =
		newlist := list
	GuiControl,, ItemChoice, |%newlist% ; by starting with | it'll replace the list in total
	newlist := ; to clear the variable for population later on
return
return
;ajout recherche non sensible � la casse
; possibilit� de s�lectionner en une fois plusieurs sujets � mettre dans plusieurs 904

;if EVENT enter ou double clique ALORS va dans une variable. Variable est affich� avec � cot� un bouton (supprimer) EVENT clique OU enter ALORS supprimer la variable
;if raccourci pour afficher Gui activ� mais fen�tre d�j� afficher, alors activate GUI mais pas erreur
;utiliser raccourci touche windows + fl�ches sur la GUI (peut-�tre avec DLL-Call)

RightSelected:
 GuiControl,, ItemChoice, |%newlist%


 ;and item focus alors va dans la variable selectionList
 
 
 
return

LeftSelected:
	msgbox, coucou kiki
return

;/////////////////////////////////////////////////////////////////////////////////

;r�cup�rer automatique le statut d'une liste de livre � partir de la cote depuis une liste excel

^7::
	;Depuis le tableau Excel, s�lectionner le champs Titre extraire son contenu dans var, S�lectionner le Champs cote, extraire son contenu dans var
	;dans Virtua Ifwinactive, lancer search par cote, WingetText, Regex sur Statut ex., fermer les fen�tre
	;Ouvrir le tableau Excel de recherche pour l'arri�re pr�t, inscrire le contenu de la var Titre, de la var cote et de la var statut exemplaire dans leur champs respectifs.

return





;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;exemple de code
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;#O::
;IfWinActive, ahk_class ConsoleWindowClass
;{
;Sleep, 500
;SetKeyDelay, 5
;Send, BETALING NETTO CONTANT OP REKENING
;Send, {ENTER}
;Sleep, 100
;Send, ING 320-0683206-08
;Send, {ENTER 2}
;Sleep, 100
;Send, PDW
;Send, {ENTER}
;Send, {Esc 2}{ENTER 8}
;}
;return





;^l::
;send, {Right 3}
;Loop
;{
;	   WinGetText, Title, A ; n�cessit� d'effectuer un trie, certaines info non-pertinentes
;    IfNotInString, %Title%, 962 ; !!!!!!!!!!!!!!! pas totalement fiable, le texte de la notice peut contenir le chiffre 962
;    {
;        send, {Tab}
;    }
;    Else
;    {
;        send, {Tab 2}
;        WinGetText, 962Content, A
;        IfInString, %962Content%, nebpun ; v�rifie si la zone 962 contient le sigle nebpun
;        {
;            Break
;        }
;        Else
;        {
;			;;;;;;962Content = fbdksjahfkjeh897r934zhrtu <> necfbv nezuc nebpup neaej
;			nebpun = nebpun
;			FoundPos := RegExMatch(962Content, "^[a-zA-Z0-9]*[\s]\<\>[\s]?(.*)$", SubPatt) ; extrait les sigles de biblioth�ques
;			subpatt1 =  %subpatt1% %nebpun% ; ajout "nebpun" aux autres sigles
;			Sort, subpatt1, CL D\s ; trie par ordre alphab�tiques des sigles ; !!?! trie fomnctionne pas peut-^'etre � cause du delimiter blank lines
;			MsgBox, The text is:`n%SubPatt1%
;			send, ^+{Right 50}{Delete} ; supprimer le contenu de la zone 962
;			send, %subpatt1% ; insertion des sigles
;        }
;    }
;}
;return 


;cr�er une boucle qui r�cup�re de contenu avec wingettext puis Tab jusqu'� inscription VIRTUA (condition sortie)
;pour ensuite parser le contenu de la notice. Les code 962 avant nebpun: neacae, neacsj, nearcn pui nebpun ...) OU plutot reprendre tous les code entre 962 et zone suivante puis les classer par ordre alphab�tique





;^F3::
;abc = abcXYZ123 <> neace nerjn necfg nebcf
;abc = abcXYZ123 <> neace nerjn necfg nebcf

;FoundPos := RegExMatch(abc, "^[a-zA-Z0-9]*[\s]\<\>[\s]?(.*)$", SubPat) 
;MsgBox, The text is:`n%FoundPos%
;MsgBox, The text is:`n%SubPat1%
;return



;^j::
;Loop
;{
;	IfNotInString, VIRTUA, %TextNotice%
;	WinGetText, TextNotice, A
;	; a mettre dans un array
;	MsgBox, The text is:`n%TextNotice%
;		continue
;
;	IfInString, VIRTUA, %TextNotice%
;	break  ; Terminate the loop
;
;		
;}
;return



;^n::
;	;Gui, Add, ListBox, vColorChoice, Red|Green|Blue|Black|White
;	loop (nbrLoop= nbrDeChoix)
;	{
;	send, ^+a904{enter}
;	send, $a bpunfar2 $b "Change tous les 3 mois 1712" $c "valeurDeChaqueChoixMultiple"fg
;	}
;Return



;^8::
;Gui, Add, Tab2,, First Tab|Second Tab|Third Tab  ; Tab2 vs. Tab requires v1.0.47.05.
;Gui, Add, Checkbox, vMyCheckbox, Sample checkbox
;Gui, Tab, 2
;Gui, Add, Radio, vMyRadio, Sample radio1
;Gui, Add, Radio,, Sample radio2
;Gui, Tab, 3
;Gui, Add, Edit, vMyEdit r5  ; r5 means 5 rows tall.
;Gui, Tab  ; i.e. subsequently-added controls will not belong to the tab control.
;Gui, Add, Button, default xm, OK  ; xm puts it at the bottom left corner.
;Gui, Show
;return

;ButtonOK:
;GuiClose:
;GuiEscape:
;Gui, Submit  ; Save each control's contents to its associated variable.
;MsgBox You entered:`n%MyCheckbox%`n%MyRadio%`n%MyEdit%
;return


^0::
; a little gui with some text from the manual
Gui 1:Add, Button, gFindString, Search
Gui 1:Add, Edit, w100 vfind
Gui 1:Add, Edit, w330 r3 vContent HSCROLL
Gui 1:Add, StatusBar,, Ctrl+F to search
SB_SetParts(300)
GuiControl, Hide, Button1
GuiControl, Hide, Edit1
Gui 1:Show,, FindSample
WinGet ControlID, ID, FindSample

someText=
(
Creating a script
Each script is a plain text file containing commands
to be executed by the program (AutoHotkey.exe).
A script may also contain hotkeys and hotstrings, or
even consist entirely of them. However, in the absence
of hotkeys and hotstrings, a script will perform its
commands sequentially from top to bottom the moment
it is launched.

To create a new script:

Open Windows Explorer and navigate to a folder of
your choice.
Pull down the File menu and choose New >> AutoHotkey
Script (or Text Document).
Type a name for the file, ensuring that it ends in .ahk.
For example: Test.ahk
Right-click the file and choose Edit Script.
On a new blank line, type the following:
#space::Run www.google.com
The symbol # stands for the Windows key, so #space means
holding down the Windows key then pressing the spacebar
to activate a hotkey. The :: means that the subsequent
command should be executed whenever this hotkey is
pressed, in this case to go to the Google web site.
To try out this script, continue as follows:

Save and close the file.
In Windows Explorer, double-click the script to launch
it. A new tray icon appears. Hold down the Windows key
and press the spacebar. A web page opens in the default
browser.
To exit or edit the script, right click its tray icon.
Note: Multiple scripts can be running simultaneously,
each with its own tray icon. Furthermore, each script
can have multiple hotkeys and hotstrings.
)
Guicontrol,,Edit2, %someText%
return

FindString:
   Gui, Submit, Nohide

   if (find != lastFind) {
    offset = 0
    hits = 0
   }

   GuiControl 1:Focus, Content                           ; focus on main help window to show selection
   SendMessage 0xB6, 0, -999, Edit2, ahk_id %ControlID%  ; Scroll to top
   StringGetPos pos, Content, %find% ,,offset            ; find the position of the search string
   if (pos = -1) {
    if (offset = 0) {
      SB_SetText("'" . find . "' not found", 1)
      SB_SetText("", 2)
    }
    else {
      SB_SetText("No more occurrences of '" . find . "'")
      SB_SetText("", 2)
      offset = 0
      hits = 0
    }
    return
   }
   StringLeft __s, Content, %pos%                        ; cut off end to count lines
   StringReplace __s,__s,`n,`n,UseErrorLevel             ; Errorlevel <- line number
   addToPos=%Errorlevel%
   SendMessage 0xB6, 0, ErrorLevel, Edit2, ahk_id %ControlID% ; Scroll to visible
   SendMessage 0xB1, pos + addToPos, pos + addToPos + Strlen(find), Edit2, ahk_id %ControlID% ; Select search text
   ; http://msdn.microsoft.com/en-us/library/bb761637(VS.85).aspx
   ; Scroll the caret into view in an edit control:
   SendMessage, EM_SCROLLCARET := 0xB7, 0, 0, Edit2, ahk_id %ControlID%
   offset := pos + addToPos + Strlen(find)
   lastFind = %find%
   hits++
   SB_SetText("'" . find . "' found in line " . addToPos + 1, 1)
   SB_SetText(hits . (hits = 1 ? " hit" : " hits"), 2)
Return

^f::
GuiControl, Show, Button1
GuiControl, Show, Edit1
Sleep 100
ControlFocus, Edit1, A
Return
	

