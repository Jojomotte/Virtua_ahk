#SingleInstance force
Esc::ExitApp


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; en cours

;Ouverture automatique des programmes au démarrage est préparation de Virtua
^Esc::
Run, virtua.exe, C:\Program Files (x86)\Vtls\Virtua 16.0.0.0, max
send, {enter}
;send, AppsKey{Down 2}{enter}AppsKey{Down 3}{enter}AppsKey{Up 2}{enter}AppsKey{Up}{enter}
send, {LWin Down}{Left}
;Run, OUTLOOK.EXE, C:\Program Files (x86)\Microsoft Office\Office14
;Run, firefox.exe, C:\Program Files (x86)\Mozilla Firefox
return


;///////////////////////////////////////////////////////////////////////////

;CODE OK ajout de 994 automatique sauf numéro de commande
^r::
;Récupère la date courante
FormatTime, year, , yyyy

;récupère la cote depuis la notice d'exemplaire
WinGetText, exemplareData, A
FoundPos := RegExMatch(exemplareData, "s)^.*Cote`r`n[1-9]`r`n(.*)`r`nInformations à propos de l'exemplaire", SubPattExemplareData)
;MsgBox, FoundPos=%FoundPos%, SubPattExemplareData=%SubPattExemplareData1%

;revenir à la notice bibliographique
;send, {Alt Down}{e}{n}{Alt Up}

;récupère le numéro de commande depuis l'Etat d'acqu.
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

;détermine la modalité d'acquisition
Gui, Add, Text,, Déterminez la modalité d'acquisition:
gui, font, s11, Arial
Gui, Add, Radio, vRadioChecked, a (achat)
Gui, Add, Radio, , d (don)
Gui, Add, Radio, , as (achat systématique)
Gui, Add, Radio, , w (inconnu ou révidion)
Gui, Add, Radio, , e (échange)
Gui, Add, Radio, , s (thèses suisse ou autre écrit académique déposé par l'Université de Neuchâtel à la BPUN(cotes TS, TSG, et 4Y))
Gui, Add, Radio, , bslc (société de lectures contemporaines)
Gui, Add, Radio, , rott (bibliothèque Rott (8R))
Gui, Add, Radio, , sgeo (Société neuchâteloise de géographie)
Gui, Add, Radio, , shan (Société neuchâteloise d'archéologie du canton de Neuchâtel)
Gui, Add, Radio, , snat (Société neuchâteloise des sciences naturelles)
Gui, Add, Radio, , ssch (Société suisse de chronométrie)
Gui, Add, Button, , OK
Gui, Add, Button, xp+60 vCancelButton gCancelSelected, Cancel
Gui, Show
Return

; Action lorsque OK est activé
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
	InputBox, numDonation, Numéro de donation, Veuillez entrez le numéro de don:, ,240, 130	; a améliorer comportement si zone de texte n'est pas frempli pas l'user
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
	send, %SubPattExemplareData1% $x ne/bpun/as/%year% ; à vérifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 4)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/w/%year% ; à vérifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 5)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/e/%year% ; à vérifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 6)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/s/%year% ; à vérifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 7)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/bslc/%year% ; à vérifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 8)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/rott/x/%year% ; à vérifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 9)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/sgeo/x/%year% ; à vérifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 10)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/shan/x/%year% ; à vérifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 11)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/snat/x/%year% ; à vérifier si juste
	Gui, Destroy 
  }
  Else If (RadioChecked = 12)
  {
	send, ^+a994{enter}
	send, %SubPattExemplareData1% $x ne/bpun/ssch/x/%year% ; à vérifier si juste
	Gui, Destroy 
  }
Return

; Action lorsque Cancel est activé
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
			send, ^+{Right 50}{Delete} ; supprimer le contenu de la zone 962 ; A améliorer en comptant le nombre de mot dans le champs
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


;///////////////////////////////////////////////////////////////////////////////////////////////

; Propose une liste de sujet 904 à choisir
^6::
list = 1. Généralités, dictionnaires, encyclopédies|2. Bibliothèques, lecture, livre ancien, manuscrits, affiches, muséologie, archives, écritures, iconographie, typographie, édition, bande dessinée|3. Edition, presse, journalisme|4. Religions, spiritualité, mythologie, ésotérisme, mystique, études bibliques|5. Philosophie, sciences humaines, éthique|6. Psychologie, psychanalyse, pédagogie, problèmes scolaires, toxicomanie, neurosciences|7. Droit, police, assurances, banques, administration publique|8. Politique, relations internationales, géopolitique, économie, marketing, gestion d’entreprise, comptabilité|9. Société, sociologie, problèmes sociaux, culture, statistiques|10. Linguistique, philologie, méthodes de langue|11. Histoire et critique littéraire, interprétation de textes|12. Littérature française (livres en français ou traduits du français et ouvrages généraux)|13. Littérature allemande (livres en allemand ou traduits de l'allemand et ouvrages généraux)|14. Littérature anglaise (livres en anglais ou traduits de l'anglais et ouvrages généraux)|15. Littérature italienne (livres en italien ou traduits de l'italien et ouvrages généraux)|16. Littérature classique (textes grecs et latins), philologie classique|17. Autres littératures (espagnole, japonaise, chinoise, russe, etc.)|18. Préhistoire, archéologie|19. Histoire par grandes périodes, par pays ou par thème, sciences auxiliaires de l’histoire|20. Histoire de l’Antiquité et du Moyen Age|21. Histoire contemporaine (20e siècle, 21e siècle)|22. Biographies|23. Beaux-arts, peinture, sculpture, dessin, arts décoratifs, histoire de l’art|24. Architecture, urbanisme, sites et monuments|25. Musique, chant, chanson, opéra|26. Cinéma, télévision, radio, jeux vidéo|27. Spectacles, danse, mode|28. Photographie|29. Théâtre|30. Géographie, géologie, paysage, sciences de la Terre|31. Anthropologie, ethnologie, traditions populaires, coutumes|32. Voyages, cartes, atlas, guides de voyage, tourisme, alpinisme|33. Sciences exactes, physique, chimie, mathématiques, astronomie, cosmologie|34. Sciences naturelles, zoologie, botanique, biologie, génétique, sciences de la vie|35. Ecologie, climat, environnement, ressources naturelles|36. Technique, industrie, métiers, transports, navigation, construction et bâtiment, bricolage|37. Horlogerie, chronométrie|38. Agriculture, viticulture, jardinage, élevage, chasse, pêche, apiculture|39. Médecine, pharmacie, professions médicales, santé, histoire de la médecine|40. Sports, loisirs, jeux|41. Famille, couple, sexualité|42. Gastronomie|43. Informatique, Internet, courrier électronique, smartphones, tablettes numériques|44. Neocomensia (auteurs et sujets neuchâtelois), histoire locale neuchâteloise
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
;ajout recherche non sensible à la casse
; possibilité de sélectionner en une fois plusieurs sujets à mettre dans plusieurs 904

;if EVENT enter ou double clique ALORS va dans une variable. Variable est affiché avec à coté un bouton (supprimer) EVENT clique OU enter ALORS supprimer la variable
;if raccourci pour afficher Gui activé mais fenêtre déjà afficher, alors activate GUI mais pas erreur
;utiliser raccourci touche windows + flèches sur la GUI (peut-être avec DLL-Call)

RightSelected:
 GuiControl,, ItemChoice, |%newlist%


 ;and item focus alors va dans la variable selectionList
 
 
 
return

LeftSelected:
	msgbox, coucou kiki
return

;/////////////////////////////////////////////////////////////////////////////////

;récupérer automatique le statut d'une liste de livre à partir de la cote depuis une liste excel

^7::
	;Depuis le tableau Excel, sélectionner le champs Titre extraire son contenu dans var, Sélectionner le Champs cote, extraire son contenu dans var
	;dans Virtua Ifwinactive, lancer search par cote, WingetText, Regex sur Statut ex., fermer les fenêtre
	;Ouvrir le tableau Excel de recherche pour l'arrière prêt, inscrire le contenu de la var Titre, de la var cote et de la var statut exemplaire dans leur champs respectifs.

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
;	   WinGetText, Title, A ; nécessité d'effectuer un trie, certaines info non-pertinentes
;    IfNotInString, %Title%, 962 ; !!!!!!!!!!!!!!! pas totalement fiable, le texte de la notice peut contenir le chiffre 962
;    {
;        send, {Tab}
;    }
;    Else
;    {
;        send, {Tab 2}
;        WinGetText, 962Content, A
;        IfInString, %962Content%, nebpun ; vérifie si la zone 962 contient le sigle nebpun
;        {
;            Break
;        }
;        Else
;        {
;			;;;;;;962Content = fbdksjahfkjeh897r934zhrtu <> necfbv nezuc nebpup neaej
;			nebpun = nebpun
;			FoundPos := RegExMatch(962Content, "^[a-zA-Z0-9]*[\s]\<\>[\s]?(.*)$", SubPatt) ; extrait les sigles de bibliothèques
;			subpatt1 =  %subpatt1% %nebpun% ; ajout "nebpun" aux autres sigles
;			Sort, subpatt1, CL D\s ; trie par ordre alphabétiques des sigles ; !!?! trie fomnctionne pas peut-^'etre à cause du delimiter blank lines
;			MsgBox, The text is:`n%SubPatt1%
;			send, ^+{Right 50}{Delete} ; supprimer le contenu de la zone 962
;			send, %subpatt1% ; insertion des sigles
;        }
;    }
;}
;return 


;créer une boucle qui récupère de contenu avec wingettext puis Tab jusqu'à inscription VIRTUA (condition sortie)
;pour ensuite parser le contenu de la notice. Les code 962 avant nebpun: neacae, neacsj, nearcn pui nebpun ...) OU plutot reprendre tous les code entre 962 et zone suivante puis les classer par ordre alphabétique





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
	

