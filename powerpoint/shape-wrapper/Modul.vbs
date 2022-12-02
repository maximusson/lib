Sub test()
    
    ' define manager
    Dim objShapeManager As clsShapeManager
    Set objShapeManager = New clsShapeManager

    ' get slide
    Set objSlide = ActivePresentation.Slides(1)
    
    ' get rectangle
    For i = 1 To 10
        Dim objRectangle As clsRectangle
        Set objRectangle = objShapeManager.CreateRectangle(objSlide)
        objRectangle.y = 50 * i
    Next
    
    'objRectangle.x = 50
    'objRectangle.y = 100
    'objRectangle.height = 50.2
    'objRectangle.width = 100
    'objRectangle.backgroundColor = RGB(255, 50, 50)
    'objRectangle.text = "Jan '21"
    'objRectangle.textAnchor = 4
    'objRectangle.font = "Segoe UI"
    'objRectangle.fontColor = RGB(255, 255, 255)
    'objRectangle.fontSize = 20
    'objRectangle.fontItalic = False
    'objRectangle.fontBold = False
    
End Sub

Sub getBlabla()
    Call getclassprops("text", "str", "String")
    Call getclassprops("textAnchor", "int", "Integer")
    Call getclassprops("font", "str", "String")
    Call getclassprops("fontSize", "int", "Integer")
    Call getclassprops("fontColor", "rgb", "RGBColor")
    Call getclassprops("fontBold", "bol", "boolean")
    Call getclassprops("fontItalic", "bol", "boolean")

End Sub

Sub getclassprops(strVariableName, strType, strTypeDescription)

'strVariableName = "text"
'strType = "str"
'strTypeDescription = "String"

' from here automatic
strVariableName2 = strType & UCase(Left(strVariableName, 1)) & Right(strVariableName, Len(strVariableName) - 1)

l = "'" & UCase(Left(strVariableName, 1)) & Right(strVariableName, Len(strVariableName) - 1) & vbCrLf
l = l & "Public Property Let " & strVariableName & "(" & strVariableName2 & " As " & strTypeDescription & ")" & vbCrLf
l = l & "    m_" & strVariableName2 & " = " & strVariableName2 & vbCrLf
l = l & "    If Not m_objShape Is Nothing Then" & vbCrLf
l = l & "        " & vbCrLf
l = l & "    End If" & vbCrLf
l = l & "End Property" & vbCrLf
l = l & "" & vbCrLf
l = l & "Public Property Get " & strVariableName & "() As " & strTypeDescription & vbCrLf
l = l & "    " & strVariableName & " = m_" & strVariableName2 & vbCrLf
l = l & "End Property" & vbCrLf
l = l & "" & vbCrLf
l = l & ""
Debug.Print l

End Sub