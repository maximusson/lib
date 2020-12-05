'Example:
output MoveObject(a, b)


Function MoveObject(objComos, objNewOwner)
'DESCRIPTION: moves a comos object to a new owner

'REVISIONS:
'(1) - 30-April-2020 - created

'INPUT:
'(1) objComos: object to be moved [Comos object]
'(2) objNewOwner: new owner [Comos object]

'OUTPUT:
'(1) MoveObject: returns true or false depending if script was succesful [boolean]

   MoveObject = false

   Set pg = Createobject("COMOSPPGeneral.GFunctions")
   pg.Move objNewOwner, objComos

   objComos.SaveAll

   MoveObject = true

End Function
