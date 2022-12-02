Dim m_objShape As shape
Dim m_dblX As Double
Dim m_dblY As Double
Dim m_dblHeight As Double
Dim m_dblWidth As Double

Dim m_rgbBackgroundColor As Long

Dim m_strText As String
Dim m_intTextAnchor As String
Dim m_rgbTextColor As Long
Dim m_strFont As String
Dim m_intFontsize As Integer

Dim m_bolFontBold As Boolean
Dim m_bolFontItalic As Boolean


' Constructor
Public Sub InitiateProperties(objShape) '

    Set m_objShape = objShape
    Me.x = 50
    Me.y = 10
    Me.width = 100
    Me.height = 50
    
    Me.backgroundColor = RGB(255, 255, 255)
    
    Me.text = "afds"
    Me.textAnchor = 4
    Me.fontColor = RGB(0, 0, 0)
    Me.font = "Segoe UI"
    Me.fontSize = 12
    
    Me.fontBold = False
    Me.fontItalic = False
    

    
End Sub


'shape object
Public Property Get shape() As shape
    Set shape = m_objShape
End Property


'X
Public Property Let x(dblX As Double)
    m_dblX = dblX
    If Not m_objShape Is Nothing Then
        m_objShape.Left = dblX
    End If
End Property

Public Property Get x() As Double
    x = m_dblX
End Property


'Y
Public Property Let y(dblY As Double)
    m_dblY = dblY
    If Not m_objShape Is Nothing Then
        m_objShape.Top = dblY
    End If
End Property

Public Property Get y() As Double
    y = m_dblY
End Property


'Width
Public Property Let width(dblWidth As Double)
    m_dblWidth = dblWidth
    If Not m_objShape Is Nothing Then
        m_objShape.width = dblWidth
    End If
End Property

Public Property Get width() As Double
    width = m_dblWidth
End Property


'Height
Public Property Let height(dblHeight As Double)
    m_dblHeight = dblHeight
    If Not m_objShape Is Nothing Then
        m_objShape.height = dblHeight
    End If
End Property

Public Property Get height() As Double
    height = m_dblHeight
End Property


'BackgroundColor
Public Property Let backgroundColor(rgbBackgroundColor)
    m_rgbBackgroundColor = rgbBackgroundColor
    If Not m_objShape Is Nothing Then
        m_objShape.Fill.ForeColor.RGB = rgbBackgroundColor
    End If
End Property

Public Property Get backgroundColor() As RGBColor
    backgroundColor = m_rgbBackgroundColor
End Property


'Text
Public Property Let text(strText As String)
    m_strText = strText
    If Not m_objShape Is Nothing Then
        m_objShape.TextFrame2.TextRange.text = strText
    End If
End Property

Public Property Get text() As String
    text = m_strText
End Property


'TextAnchor
Public Property Let textAnchor(intTextAnchor As Integer)
    m_intTextAnchor = intTextAnchor
    If Not m_objShape Is Nothing Then
        Select Case intTextAnchor
        Case 1
            intVAnchor = msoAnchorTop
            intHAnchor = msoAlignLeft
        Case 2
            intVAnchor = msoAnchorTop
            intHAnchor = msoAlignCenter
        Case 3
            intVAnchor = msoAnchorTop
            intHAnchor = msoAlignRight
        Case 4
            intVAnchor = msoAnchorMiddle
            intHAnchor = msoAlignLeft
        Case 5
            intVAnchor = msoAnchorMiddle
            intHAnchor = msoAlignCenter
        Case 6
            intVAnchor = msoAnchorMiddle
            intHAnchor = msoAlignRight
        Case 7
            intVAnchor = msoAnchorBottom
            intHAnchor = msoAlignLeft
        Case 8
            intVAnchor = msoAnchorBottom
            intHAnchor = msoAlignCenter
        Case 9
            intVAnchor = msoAnchorBottom
            intHAnchor = msoAlignRight
        End Select
    
        ' define alignment
        With m_objShape.TextFrame2
            'rotation
            '.Orientation = msoTextOrientationHorizontal
            '.HorizontalAnchor = msoAnchorCenter
            .VerticalAnchor = intVAnchor
            .TextRange.ParagraphFormat.Alignment = intHAnchor
        End With
    End If
End Property

Public Property Get textAnchor() As Integer
    textAnchor = m_intTextAnchor
End Property


'Font
Public Property Let font(strFont As String)
    m_strFont = strFont
    If Not m_objShape Is Nothing Then
        ' define font
        With m_objShape.TextFrame2.TextRange.font
            ' font type
            .name = strFont
        End With
    End If
End Property

Public Property Get font() As String
    font = m_strFont
End Property


'FontSize
Public Property Let fontSize(intFontSize As Integer)
    m_intFontsize = intFontSize
    If Not m_objShape Is Nothing Then
        ' define font
        With m_objShape.TextFrame2.TextRange.font
            .Size = intFontSize
        End With
    End If
End Property

Public Property Get fontSize() As Integer
    fontSize = m_intFontsize
End Property


'fontColor
Public Property Let fontColor(rgbFontColor)
    m_rgbFontColor = rgbFontColor
    If Not m_objShape Is Nothing Then
        ' define font
        With m_objShape.TextFrame2.TextRange.font
            .Fill.ForeColor.RGB = rgbFontColor
        End With
    End If
End Property

Public Property Get fontColor() As RGBColor
    fontColor = m_rgbFontColor
End Property


'FontBold
Public Property Let fontBold(bolFontBold As Boolean)
    m_bolFontBold = bolFontBold
    If Not m_objShape Is Nothing Then
        ' define font
        With m_objShape.TextFrame2.TextRange.font
            .Bold = bolFontBold
        End With
    End If
End Property

Public Property Get fontBold() As Boolean
    fontBold = m_bolFontBold
End Property


'FontItalic
Public Property Let fontItalic(bolFontItalic As Boolean)
    m_bolFontItalic = bolFontItalic
    If Not m_objShape Is Nothing Then
        ' define font
        With m_objShape.TextFrame2.TextRange.font
            .Italic = bolFontItalic
        End With
    End If
End Property

Public Property Get fontItalic() As Boolean
    fontItalic = m_bolFontItalic
End Property
