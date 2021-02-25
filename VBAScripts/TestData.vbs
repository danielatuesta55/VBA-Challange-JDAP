Sub TestData ()
    
    'PART 1: I need to create a Loop that works through out all the worksheets in the workbook. For this I will use For Each ws In Worksheets command.
    For Each ws In Worksheets
        'I need to declare the variable susing the Dim command 
        'WorksheetName as string
        Dim WorksheetName As String
        'Declaring the current row as long
        Dim i As long
        'Declare the value of J that will point out the start of Ticker block section for interpretation 
        Dim j As long
        'Declare a counter for the 'Ticker' 
        Dim Tickercount As long
        'Declare variable to be able to get the last row in column A 'ticker'
        Dim LastRowColumnA As long
        'Declare variable to be ablet to get the las row in column J 'Yearly Change (open-close)
        Dim LastRowColumnJ As long
        'Decalre a varibale to be able to calculate the percentage calculation 
        Dim PercentageChange as double  
        


 
End Sub
