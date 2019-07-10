#include-once
#include "CLOE.au3"
Local $m_CouleurLiens = 0x6666CC
Local $m_oDocHtml, $m_ChildNode, $m_oLinks, $m_oTableFields, $m_oFields, $g_aArrayEdits, $m_aLabels
Local $m_sName, $m_hList32, $m_aBounds, $m_aHandlesFieldsVisibles

Func CLOE_SCREEN_getHTMLChild($sTagname, $sHtmlText)
	$m_oDocHtml = CLOEGetFrameDoc($sHtmlText)  ; "span;PROFILS DE FACTURATION;innertext")
	$m_ChildNode = _IEParentQuerySelectAvanced($m_oDocHtml, $sTagname, $sHtmlText)    ;"form","td;Mes comptes en rejet;innertext" )
	If IsObj($m_ChildNode) = 0 Then Return mySetError(1, 0, 0, "CLOE_SCREEN_getHTMLChild", @ScriptLineNumber, "$m_ChildNode = 0")

	Return SetError(0, 0, $m_ChildNode)
EndFunc   ;==>CLOE_SCREEN_getHTMLChild

Func CLOE_SCREEN_List()
	$m_hList32 = WaitFunc("CRM_EnumWindows", 500, 20000, "#3.*")
	If UBound($m_hList32) = 0 Then Return MySetError(1, 0, 0, "CLOE_SCREEN_getListRecords", @ScriptLineNumber, "UBound($m_hList32) = 0 ")
	_ArraySort($m_hList32, 0, 0, 0, 6)
	Return $m_hList32
EndFunc   ;==>CLOE_SCREEN_List

Func CLOE_SCREEN_Search($sCriteria, $bOpen = True)
	Local $aCriteria = _StringExplode2D($sCriteria, "|", "=")
	Local $i

	For $i = 0 To UBound($aCriteria) - 1
		If Not IsArray($aCriteria[$i]) Then ContinueLoop
		If UBound($aCriteria[$i]) <> 2 Then ContinueLoop

		CLOE_SCREEN_getHTMLChild("tr", "td;" & ($aCriteria[$i])[0] & ";innertext")
		If IsObj($m_ChildNode) > 0 Then
			With $m_ChildNode.getElementsByTagname("input")
				If .length > 0 Then .item(0).value = ($aCriteria[$i])[1]
			EndWith
		EndIf
	Next
	CLOE_SCREEN_Click("Rechercher")
	CLOE_SCREEN_GetFields()

	For $i = 0 To UBound($aCriteria) - 1
		If Not IsArray($aCriteria[$i]) Then ContinueLoop
		If UBound($aCriteria[$i]) <> 2 Then ContinueLoop

		If StringInStr(CLOE_SCREEN_Read(($aCriteria[$i])[0] & ".*"), ($aCriteria[$i])[1]) = 0 Then
			Return MySetError(@error, 0, 0, "CLOE_SCREEN_Search", @ScriptLineNumber, "No match creteria")
		EndIf
	Next

	If $bOpen Then
		Local $coords = [15, 23, 116, 44]
		If Waitfunc("CLOE_SCREEN_ListSelect", 100, 20000, $coords) = 0 Then
			Return MySetError(1, 0, 0, "CLOE_SCREEN_Search", @ScriptLineNumber, "Echec click CLOE_SCREEN_ListSelect")
		EndIf
	EndIf

	Return 1
EndFunc   ;==>CLOE_SCREEN_Search

Func CLOE_SCREEN_ListSelect($coords)
	If IsArray(CLOE_SCREEN_List()) = 0 Then
	Else
		Local $StatutClic = CLOEBoutonClick($IE_CLOE_Hwnd, "", $m_hList32[0][0], $m_CouleurLiens, $coords)
		If $StatutClic = 1 Then Return SetError(0, 0, 1)
	EndIf

	Return SetError(1, 0, 0)
EndFunc   ;==>CLOE_SCREEN_ListSelect

Func CLOE_SCREEN_Click($sCaption)
	CLOE_SCREEN_getHTMLChild("form", "a;" & $sCaption & ";innertext")
	$m_oLinks = _IEGetElementByAttribute($m_oDocHtml, "a", "Rechercher", "innertext")
	If IsObj($m_oLinks) = 0 Then Return MySetError(1, 0, 0, "CLOE_SCREEN_Click", @ScriptLineNumber, "IsObj($m_oLinks)  = 0 ")
	$m_oLinks.click()
	$m_oDocHtml = 0
	Return 1
EndFunc   ;==>CLOE_SCREEN_Click
Local $sText = "Evénement:=Avenant"
CLOE_SCREEN_FieldsEdit($sText)
Func CLOE_SCREEN_FieldsEdit($sText)
	Local $aText = _StringExplode2D($sText, "|", "=")
	Local $i
	CLOE_SCREEN_GetFields()

	For $i = 0 To UBound($aText) - 1
		If Not IsArray($aText[$i]) Then ContinueLoop
		If UBound($aText[$i]) <> 2 Then ContinueLoop

		Local $iFieldIndex = _ArraySearch($g_aArrayEdits, ($aText[$i])[0] & ".*", 0, 0, 0, 3, 1, 10)
		If $iFieldIndex < 0 Then
			Return MySetError(1, 0, 0, "CLOE_SCREEN_FieldWrite", @ScriptLineNumber, "$iFieldIndex <0")
		EndIf
		CLOE_SCREEN_FieldWrite(($aText[$i])[1], ($aText[$i])[1])
	Next
EndFunc

;~ Local $sFieldsCopied = "Raison Sociale:;SIREN:;SIRET:;Code NAF/APE:;Classe de Gestion:;Marquage:;Type de Profil:;Moyen de Paiement:;Délai de Paiement:;Rythme de Facturation:;N° Téléphone:"
;~ CLOE_SCREEN_Duplicate($sFieldsCopied)
Func CLOE_SCREEN_Duplicate($sFieldsCopied)
	Local $aFieldsCopied = StringSplit($sFieldsCopied, ";")
	$m_oDocHtml = 0
	Local $aFieldOld = CLOE_SCREEN_GetFields()

	Do
		WinSetOnTop($IE_CLOE_Hwnd, "", 1)
		ControlSend($IE_CLOE_Hwnd, "", $IEServer_Hwnd, "^b")
		WinSetOnTop($IE_CLOE_Hwnd, "", 0)
		$m_oDocHtml = 0
		Sleep(2000)
		CLOE_SCREEN_GetFields()
	Until $aFieldOld[0][2] <> $g_aArrayEdits[0][2]
	Local $i = 1
	While $i <= $aFieldsCopied[0]
		Local $iFieldIndex = _ArraySearch($g_aArrayEdits, $aFieldsCopied[$i] & '.*', 0, 0, 0, 3, 1, 10)
		If CLOE_SCREEN_FieldWrite($aFieldsCopied[$i] & '.*', $aFieldOld[$iFieldIndex][2]) = $aFieldOld[$iFieldIndex][2] Then
			$i += 1
		EndIf

	WEnd

	If StringInStr(CLOE_SCREEN_Read("Marquage:" & '.*'), "Duplicata") <> 1 Then CLOE_SCREEN_FieldWrite("Marquage:" & '.*', "Duplicata REF: " & $aFieldOld[0][2] & "")

	WinSetOnTop($IE_CLOE_Hwnd, "", 1)
	ControlSend($IE_CLOE_Hwnd, "", $IEServer_Hwnd, "^s")
	WinSetOnTop($IE_CLOE_Hwnd, "", 0)
	Return SetError(0, 0, CLOE_SCREEN_GetFields())
EndFunc   ;==>CLOE_SCREEN_Duplicate

Func CLOE_SCREEN_FieldWrite($sFieldLabel, $sNewValue)

	Opt("SendKeyDelay", 50)
	Opt("SendKeyDownDelay", 20)
	SetRapport("", "", "Saisie '"  & $sFieldLabel & "= " &  $sNewValue & "'" )
	Local $iFieldIndex = _ArraySearch($g_aArrayEdits, $sFieldLabel, 0, 0, 0, 3, 1, 10)
	If $iFieldIndex < 0 Then
		Return MySetError(1, 9, 0, "CLOE_SCREEN_FieldWrite", @ScriptLineNumber, "$iFieldIndex <0")
	EndIf

	If WinGetText($g_aArrayEdits[$iFieldIndex][0]) <> $g_aArrayEdits[$iFieldIndex][11].value Then

	EndIf
	Local $sOldValue = WinGetText($g_aArrayEdits[$iFieldIndex][0])

	While 1
		If $g_aArrayEdits[$iFieldIndex][11].Value <> $sNewValue Then
			$g_aArrayEdits[$iFieldIndex][11].Focus()
			$g_aArrayEdits[$iFieldIndex][11].click()
			Sleep(300)
			$g_aArrayEdits[$iFieldIndex][11].Value = $sNewValue
			Sleep(300)
			ControlClick($IE_CLOE_Hwnd, "", $IEServer_Hwnd)
		Else
			ExitLoop
		EndIf

		If ($sNewValue <> ControlGetText($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0])) Then
			ControlSetText($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0], "")
			ControlClick($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0])
			$g_aArrayEdits[$iFieldIndex][11].value = $sNewValue
			ControlSetText($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0], $sNewValue)
;~ 			WinSetOnTop($IE_CLOE_Hwnd, "", 1)
			ControlSend($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0], " {BS}{tab}")
			Sleep(1000)
		Else
			ExitLoop
		EndIf

		If ($sNewValue <> ControlGetText($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0])) Then
			ControlSetText($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0], "")
			ControlClick($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0])
			WinSetOnTop($IE_CLOE_Hwnd, "", 1)
			ControlSend($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0], $sNewValue)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $sNewValue = ' & $sNewValue & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			ControlSend($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0], " {BS}{tab}")
			Sleep(1000)
		Else
			ExitLoop
		EndIf
	WEnd
	Return ControlGetText($IE_CLOE_Hwnd, "", $g_aArrayEdits[$iFieldIndex][0])
EndFunc   ;==>CLOE_SCREEN_FieldWrite

Func CLOE_SCREEN_GetFields()
	$m_oDocHtml = CLOEGetFrameDoc()
	If IsObj($m_oDocHtml) = 0 Then Return SetError(@error, 9, 0)

	Return ($g_aArrayEdits)
EndFunc   ;==>CLOE_SCREEN_GetFields

Func CLOE_SCREEN_Read($sFieldLabel)
	CLOE_SCREEN_GetFields()
	Local $iFieldIndex = _ArraySearch($g_aArrayEdits, $sFieldLabel, 0, 0, 0, 3, 1, 10)
	If $iFieldIndex < 0 Then
		Return MySetError(1, 0, 0, "CLOE_SCREEN_FieldWrite", @ScriptLineNumber, "$iFieldIndex <0")
	EndIf
	Return SetError(0, 0, $g_aArrayEdits[$iFieldIndex][12])
EndFunc   ;==>CLOE_SCREEN_Read
