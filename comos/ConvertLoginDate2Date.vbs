'Example:
intLoginDate = 12208
intDate = ConvertLoginDate2Date(intLoginDate)
Output "intLoginDate: " & vbTab & intLoginDate:
Output "intDate: " & vbTab & vbTab & intDate
Output "Year: " & vbTab & vbTab & Year(CDate(intDate))
Output "Month:" & vbTab & vbTab & Month(CDate(intDate))
Output "Day:" & vbTab & vbTab & Day(CDate(intDate))


Function ConvertLoginDate2Date(intLoginDate)
'DESCRIPTION: Converts a login date (special COMOS date type in sql table) to corresponding Microsoft date. No guarantee for correctness of script!

'INPUT:
'(1) intLoginDate: date as integer [integer]

'OUTPUT:
'(1) ConvertLoginDate2Date: converted login date [date]

   intLoginDate = Int(intLoginDate)

   ' Reference year
   intReferenceYear = 2000

   ' COMOS Year has 600 days
   intDaysPerYearInComos = 600
   
   ' COMOS Month has 40 days
   intDaysPerMonthInComos = 40
   
   ' First of january is 41
   intFirstOfJanuary = 41
   
   intCounterYear = Int((intLoginDate - intFirstOfJanuary) / intDaysPerYearInComos)
   intYear = intReferenceYear + intCounterYear
   intDayInYear = intLoginDate - intCounterYear * intDaysPerYearInComos
   intMonth = Int((intDayInYear - intDaysPerMonthInComos) / intDaysPerMonthInComos) + 1
   intDay = intDayInYear - intMonth * intDaysPerMonthInComos
   
   ConvertLoginDate2Date = CDbl(CDate(intDay & "-" & MonthName(intMonth) & "-" & intYear))
    
End Function
