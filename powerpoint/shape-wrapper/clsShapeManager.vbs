Dim m_colRectangles As Collection

Private Sub Class_Initialize()
    Set m_colRectangles = New Collection
End Sub
 
Private Sub Class_Terminate()
    Set m_colRectangles = Nothing
End Sub

Public Function CreateRectangle(objSlide)
    
    ' create real shape
    If objSlide Is Nothing Then Exit Function
    Set objShape = objSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0)
    
    ' create object rectangle
    Set objRectangle = New clsRectangle
    Call objRectangle.InitiateProperties(objShape)
    
    m_colRectangles.Add objRectangle
    Set CreateRectangle = objRectangle
    
End Function

Public Function colRectangles()
    Set colRectangles = m_colRectangles
End Function
