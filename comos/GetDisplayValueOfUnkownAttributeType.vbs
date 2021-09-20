Function GetDisplayValueOfUnkownAttributeType(objAttr)
' DESCRIPTION: tries to get attribut displayvalue

' ComosSUIEdit.SUIEdit							Eingabefeld, Edit: [Min Max], Edit: [Min Value Max]	
' ComosSUIImage.SUIImage						Bildauswahl	
' ComosSUICheck.SUICheck						Checkbox	
' ComosSUIFOpen.SUIFOpen						Dateiauswahl	
' ComosSUIDate.SUIDate							Datum	
' ComosSUIList.SUIList							Liste	
' ComosSUIMemo.SUIMemo							Memofeld (ASCII)	
' ComosSUIRtf.SUIRtf								Memofeld (RTF)	
' ComosSUIQuery.SUIQuery						Objektabfrage	
' ComosSUIBorder.SUIBorder					Rahmen	
' ComosSUIRepeater.SUIRepeater			Repeater	
' ComosSUIButton.SUIButton					Schaltfläche	
' ComosSUISignature.SUISignature		Unterschrift	
' ComosSUILink.SUILink							Verweis	
' ComosSUILabel.SUILabel						Beschreibung	

	GetDisplayValueOfUnkownAttributeType = ""
	If objAttr Is Nothing Then Exit Function
	
	strValue = ""
	
	Select Case objAttr.ControlType
	Case "ComosSUIEdit.SUIEdit"
		Select Case objAttr.RangeType
		Case 1 		'min value max
			strValue = objAttr.GetDisplayXValue(0) & " // " & objAttr.DisplayValue & " // " & objAttr.GetDisplayXValue(1)
		Case 2 		'min max
			strValue = objAttr.GetDisplayXValue(0) & " // " & objAttr.GetDisplayXValue(1)
		Case Else 'normal
			strValue = objAttr.DisplayValue
		End Select
		
		If objAttr.Unit <> "" Then
			Set objPhysUnit = objAttr.GetPhysUnit(objAttr.Unit)
			If Not objPhysUnit Is Nothing Then
				strValue = strValue & " [" & objPhysUnit.Label & "]"
			End If
		End If
		
	Case "ComosSUICheck.SUICheck"
		strValue = objAttr.DisplayValue
		If strValue = "" Then strValue = 0
		
	Case Else
		' not defined so far
	End Select
	
	GetDisplayValueOfUnkownAttributeType = strValue

End Function
