' EXAMPLE :
Set objComos = a
Set objNewOwner = b
bMoved = MoveObject(objComos, objNewOwner)

Function MoveObject(objComos, objNewOwner)
' DESCRIPTION : moves a comos object to a new owner

' REVISIONS :
' (1) 30-Apr-2020 : created
' (2) 19-May-2022: beautify script
   
' INPUT :
' (1) objComos: object to be moved [Comos object]
' (2) objNewOwner: new owner [Comos object]

' OUTPUT :
' (1) MoveObject: returns true or false depending if script was succesful [boolean]

   MoveObject = false

   Set pg = Createobject("COMOSPPGeneral.GFunctions")
   pg.Move objNewOwner, objComos

   objComos.SaveAll

   MoveObject = true

End Function
