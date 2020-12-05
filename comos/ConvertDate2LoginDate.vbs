'Example:
strDate = "09-May-2020"
intDate = CDbl(CDate(strDate))
intLoginDate = ConvertDate2LoginDate(intDate)

Output "Date: " & vbTab & vbTab & vbTab & strDate
Output "Microsoft Date (input): " & vbTab & intDate
Output "Login Date (output): " & vbTab & intLoginDate


Function ConvertDate2LoginDate(intDate)
'DESCRIPTION: Converts a microsoft date to comos date (special COMOS date type in sql table)
'no guarantee for correctness =)

'REVISIONS:
'(1) 30-April-2020 - created

'INPUT:
'(1) intDate: date as integer [integer]

'OUTPUT:
'(1) ConvertDate2LoginDate: converted login date [date]

   intDate = Int(intDate)
   
   intDay = Day(CDate(intDate))
   intMonth = Month(CDate(intDate))
   intYear = Year(CDate(intDate))
   
   ' Reference year
   intReferenceYear = 2000

   ' COMOS Year has 600 days
   intDaysPerYearInComos = 600

   ' COMOS Month has 40 days
   intDaysPerMonthInComos = 40

   ' First of january is 41
   intFirstOfJanuary = 41
   
   intCountDaysInYear = (intYear - intReferenceYear) * intDaysPerYearInComos
   intCountDaysInMonth = (intMonth - 1) * intDaysPerMonthInComos
   intCountDays = intDay + intFirstOfJanuary - 1
   
   ConvertDate2LoginDate = intCountDaysInYear + intCountDaysInMonth + intCountDays

End Function
