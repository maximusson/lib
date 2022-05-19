' EXAMPLE :
Set objAttribute = a
strDisplayValue = GetDisplayValueOfUnknownAttributeType(objAttribute)

Function GetDisplayValueOfUnknownAttributeType(objAttribute)
' DESCRIPTION: tries to get attribut displayvalue

' SCRIPT REVISIONS :
' (1) 01-Feb-2022 : Created
' (2) 19-May-2022 : beautify script
	
' INPUT :
' (1) objAttribute: attribut - [comos attribut object]

' OUTPUT :
' (1) GetDisplayValueOfUnknownAttributeType: returns displayvalue if possible [string] 	
		
	GetDisplayValueOfUnknownAttributeType = ""
	If objAttribute Is Nothing Then Exit Function
	If objAttribute.SystemType <> 10 Then Exit Function
			
	strValue = ""
	
	Select Case objAttribute.ControlType
	Case "ComosSUIEdit.SUIEdit"
		Select Case objAttribute.RangeType
		Case 1 		'min value max
			strValue = objAttribute.GetDisplayXValue(0) & " // " & objAttribute.DisplayValue & " // " & objAttribute.GetDisplayXValue(1)
		Case 2 		'min max
			strValue = objAttribute.GetDisplayXValue(0) & " // " & objAttribute.GetDisplayXValue(1)
		Case Else 'normal
			strValue = objAttribute.DisplayValue
		End Select
		
		If objAttribute.Unit <> "" Then
			Set objPhysUnit = objAttribute.GetPhysUnit(objAttr.Unit)
			If Not objPhysUnit Is Nothing Then
				strValue = strValue & " [" & objPhysUnit.Label & "]"
			End If
		End If
		
	Case "ComosSUICheck.SUICheck"
		strValue = objAttribute.DisplayValue
		If strValue = "" Then strValue = 0
		
	Case Else
		' not defined so far
	End Select
	
	GetDisplayValueOfUnknownAttributeType = strValue

End Function
		
' ComosSUIEdit.SUIEdit			Eingabefeld, Edit: [Min Max], Edit: [Min Value Max]	
' ComosSUIImage.SUIImage		Bildauswahl	
' ComosSUICheck.SUICheck		Checkbox	
' ComosSUIFOpen.SUIFOpen		Dateiauswahl	
' ComosSUIDate.SUIDate			Datum	
' ComosSUIList.SUIList			Liste	
' ComosSUIMemo.SUIMemo			Memofeld (ASCII)	
' ComosSUIRtf.SUIRtf			Memofeld (RTF)	
' ComosSUIQuery.SUIQuery		Objektabfrage	
' ComosSUIBorder.SUIBorder		Rahmen	
' ComosSUIRepeater.SUIRepeater		Repeater	
' ComosSUIButton.SUIButton		Schaltfläche	
' ComosSUISignature.SUISignature	Unterschrift	
' ComosSUILink.SUILink			Verweis	
' ComosSUILabel.SUILabel		Beschreibung
