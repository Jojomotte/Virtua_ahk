#SingleInstance force

Esc::ExitApp

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Ajout de: 019 Cat. d'après la couv. $9 nebpun/mm.aaaa
^1::
FormatTime, monthYear, , MM.yyyy
signature = $9 nebpun/%monthYear%
	send, ^+a019{enter}
	send, Cat. d'après la couv. %signature%
Return	

; Ajout de: 019 BPUN: dépouillement des contrib. neuch. (auteur) $9 nebpun/mm.aaaa
^2::
FormatTime, monthYear, , MM.yyyy
signature = $9 nebpun/%monthYear%
	send, ^+a019{enter}
	send, BPUN: dépouillement des contrib. neuch. () %signature%
Return

; Ajout de: 019 BPUN: = NE $9 nebpun/mm.aaaa
^3::
FormatTime, monthYear, , MM.yyyy
signature = $9 nebpun/%monthYear%
	send, ^+a019{enter}
	send, BPUN: = NE %signature%
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Ajout de 040 RERO nebpun
^Numpad0::
	send, ^+a040{enter}
	send, RERO nebpun
Return

; Ajout de nebpun
^NumLock::
	send, nebpun
Return

; Ajout de: 904 neaut
^numpad1::
	send, ^+a904{enter}
	send, neaut
return

; Ajout de: 904 nedit
^numpad2::
	send, ^+a904{enter}
	send, nedit
return

; Ajout de: 904 neoco
^numpad3::
	send, ^+a904{enter}
	send, neoco
return

; Ajout de: 904 neimp
^numpad4::				
	send, ^+a904{enter}
	send, neimp
return

; Ajout de: 904 nebpugast
^numpad5::
	send, ^+a904{enter}
	send, nebpugast
return

; Ajout de: 904 nebpuenf
^numpad6::
	send, ^+a904{enter}
	send, nebpuenf
return

; Ajout de: 500 Justification du tirage:
^numpad7::
	send, ^+a500{enter}
	send, Justification du tirage: 
Return

; Ajout de: 500 La couv. porte:
^numpad8::
	send, ^+a500{enter}
	send, La couv. porte: 
Return

; Ajout de: 500 Traduit de:
^numpad9::
	send, ^+a500{enter}
	send, Traduit de: 
Return

; Ajout de: 500 Ouvrage publié à l'occasion de l'exposition à 
^NumpadDiv::
	send, ^+a500{enter}
	send, Ouvrage publié à l'occasion de l'exposition à    
Return

; Ajout de: 500 Texte en français ou en anglais
^NumpadMult::
	send, ^+a500{enter}
	send, Texte en français ou en anglais ;récupérer info de la 049+condition
Return

; Ajout de: 1 vol. (non paginé)
^NumpadAdd::
	send, 1 vol. (non paginé)
return

; Ajout de: 500 A également contribué:
^NumpadSub::
	send, ^+a500{enter}
	send, A également contribué: 
return

; Ajout de: 670 Ancien fichier des auteurs neuchâtelois de la Bibliothèque publique et universitaire de Neuchâtel (BPUN)
^NumpadEnter::
	send, ^+a670{enter}
	send, Ancien fichier des auteurs neuchâtelois de la Bibliothèque publique et universitaire de Neuchâtel (BPUN)
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; lancer recherche par titre de publication en série
^+p::	
	send, !rt
	send, {Tab 2}{Down 6}{Tab}
return



;!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
;!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
;!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


; CODE OK, Date automatique pour les sujets en 904 
^6::
FormatTime, year904, , yy ; -> 904 seulement 2 années précédentes
FormatTime, month904, , MM

if month904 <= 03
{
	month904 = 03
}	
else if month904 <= 06 && month904 > 03
{
	month904 = 06
}
else if month904 <= 09 && month904 > 06
{
	month904 = 09
}
else if month904 <= 12 && month904 > 09
{
	month904 = 12
}
else 
{
	MsgBox, La date ne correspond pas.
}

send, ^+a904{enter}
send, $a bpunfar2 $b %year904%%month904% $c fg
return



; GUI sujets 904
^+y::
list = 1. Généralités, dictionnaires, encyclopédies|2. Bibliothèques, lecture, livre ancien, manuscrits, affiches, muséologie, archives, écritures, iconographie, typographie, édition, bande dessinée|3. Edition, presse, journalisme|4. Religions, spiritualité, mythologie, ésotérisme, mystique, études bibliques|5. Philosophie, sciences humaines, éthique|6. Psychologie, psychanalyse, pédagogie, problèmes scolaires, toxicomanie, neurosciences|7. Droit, police, assurances, banques, administration publique|8. Politique, relations internationales, géopolitique, économie, marketing, gestion d’entreprise, comptabilité|9. Société, sociologie, problèmes sociaux, culture, statistiques|10. Linguistique, philologie, méthodes de langue|11. Histoire et critique littéraire, interprétation de textes|12. Littérature française (livres en français ou traduits du français et ouvrages généraux)|13. Littérature allemande (livres en allemand ou traduits de l'allemand et ouvrages généraux)|14. Littérature anglaise (livres en anglais ou traduits de l'anglais et ouvrages généraux)|15. Littérature italienne (livres en italien ou traduits de l'italien et ouvrages généraux)|16. Littérature classique (textes grecs et latins), philologie classique|17. Autres littératures (espagnole, japonaise, chinoise, russe, etc.)|18. Préhistoire, archéologie|19. Histoire par grandes périodes, par pays ou par thème, sciences auxiliaires de l’histoire|20. Histoire de l’Antiquité et du Moyen Age (jusqu'en 1492)|21. Histoire moderne et contemporaine (à partir de 1492)|22. Biographies|23. Beaux-arts, peinture, sculpture, dessin, arts décoratifs, histoire de l’art|24. Architecture, urbanisme, sites et monuments|25. Musique, chant, chanson, opéra|26. Cinéma, télévision, radio, jeux vidéo|27. Spectacles, danse, mode|28. Photographie|29. Théâtre|30. Géographie, géologie, paysage, sciences de la Terre|31. Anthropologie, ethnologie, traditions populaires, coutumes|32. Voyages, cartes, atlas, guides de voyage, tourisme, alpinisme|33. Sciences exactes, physique, chimie, mathématiques, astronomie, cosmologie|34. Sciences naturelles, zoologie, botanique, biologie, génétique, sciences de la vie|35. Ecologie, climat, environnement, ressources naturelles|36. Technique, industrie, métiers, transports, navigation, construction et bâtiment, bricolage|37. Horlogerie, chronométrie|38. Agriculture, viticulture, jardinage, élevage, chasse, pêche, apiculture|39. Médecine, pharmacie, professions médicales, santé, histoire de la médecine|40. Sports, loisirs, jeux|41. Famille, couple, sexualité|42. Gastronomie|43. Informatique, Internet, courrier électronique, smartphones, tablettes numériques|44. Neocomensia (auteurs et sujets neuchâtelois), histoire locale neuchâteloise
	Gui, font, s10, Verdana
	Gui, Add, Edit, w200 x5 vName gAutoComplete, 
	Gui, Add, ListBox, w965 x5 vItemChoice r44, %list%
	;Gui, Add, ListBox, w965 x5 vItemMove r44
	Gui, Add, Button, vOKButton gOKSelected Default, OK
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

; Actions on OK or close
OKSelected:
GuiClose:
  Gui Destroy
Return

return






