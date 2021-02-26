Sub MultiYearStockData_JDAP ()
    
    'PART 1: I need to create a Loop that works through out all the worksheets in the workbook. For this I will use For Each ws In Worksheets command.
    For Each ws In Worksheets
    'PART 2: Declare all the variables that are going to be used 
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
        'Declare a varibale to be able to calculate the percentage change  calculation 
        Dim PercentageChange as Double  
        'Declare a variable to calculate greatest increase calculation
        Dim GreatestIncrease as Double
        'Declare a variable to calculate greatest decrease calculation
        Dim GreatestDecrease as Double
        'Declare variable for Greatest volume, must declare as doule not long or else it will generate a overflow error
        dim GreatestVolume as Double

    'PART 3: Getting the WorksheetName
        WorksheetName = ws.Name
        'Msgbox stating the Worksheetname to verify the code ran
        'Msgbox(WorksheetName)
    'PART 4: For the homwwork excercise I have to create the column headers.
        'Creating the column headers 
        ws.Cells(1,10).value="Ticker"
        ws.Cells(1,11).value="Yearly Change"
        ws.Cells(1,12).value="Percentage Change"
        ws.Cells(1,13).value="Total Stock Volumen"
        ws.cells(1,17).value="Ticker"
        ws.Cells(1,18).value = "Value"
        ws.Cells(2,16).value = "Greatest % Increase"
        ws.cells(3,16).value = "Greates % Decrease"
        ws.Cells(4,16).Value = "Greatest Total Volumen"

    'PART 5: Setting the parameters for the loop and ticker counter variable to 2. I need to set values for previously declared variables
        Tickercount = 2
        'Set it to row 2
        j = 2
        'To find the last empty cell in column A I will use the row cunt built in formula I learned this in class and I got the formula again on the following webpage: https://www.excelcampus.com/vba/find-last-row-column-cell/#:~:text=To%20find%20the%20last%20used,the%20rows%20in%20the%20worksheet.
         LastRowColumnA = ws.Cells(Rows.Count,1).End(xlUp).Row
         'I can create a msgbox to check if code ran properly
         'Msgbox("Last row in column A = " & LastRowColumnA)

    'PART 6: Know I need to be able to loop through all of the rows creating a For Loop
            For i = 2 to LastRowColumnA
                'Now I need to know if the ticker name changed from the one on the previouse cell. I will do this using a conditional IF i .. Then
                if ws.Cells(i + 1, 1).Value <> ws.Cells(i,1).Value Then
                    'Print the Ticker name on Range J2:J
                    ws.Cells(Tickercount,10).Value = ws.Cells(i,1).Value
    'PART 7: Calculate the yearly Change inside the for loop that I am currently on and the condition that is currently in place
                    'Calculate the Yearly Change and print the value on cell Range K2:K
                    ws.Cells(Tickercount,11).value = ws.Cells(i,6).Value - ws.cells(j,3).Value
    'PART8: I need to code conditional formating by using a nested conditional If statement inside the current loop and conditional
                        'Coding conditional formating
                        If ws.Cells(Tickercount,11).Value < 0 Then 
                            'I need to set the cells interiro color to red I got the color number for coding on VBA on the following link that was shared by out TA Erin: https://www.automateexcel.com/excel-formatting/color-reference-for-color-index/
                            ws.Cells(Tickercount,11).Interior.ColorIndex = 3
                        else    
                            'For all the cells above 0 I need to change the interior of the cell to the color green
                            ws.Cells(Tickercount,11).Interior.ColorIndex = 4 
                        'I need to end this conditional and continue on the parent conditional
                        End If

    'PART 9: I need to calculate the percentage change and print the value on Range L2:L 
                        'I need to created another nested conditional
                        If ws.Cells(j,3).Value <> 0 Then 
                            PercentageChange = ((ws.Cells(i,6).Value - ws.Cells(j,3).Value)/ws.Cells(j,3).Value)
                            'We need to change the format of the cell so its on percentage value type I found out how to do this on the following webpage: https://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/
                            ws.Cells(Tickercount,12).Value = Format(PercentageChange, "Percent")

                        else
                            ws.Cells(Tickercount,12).Value = Format(0,"Percent")
                        End If
    'PART 10: I need to calculate the total stock volumen in Column M 
                    ' using the worksheetFunction.Sum I will get the value for the total stock volume. i got this code from the following webpage: https://www.automateexcel.com/vba/sum-function/
                    ws.Cells(Tickercount,13).Value = worksheetFunction.Sum(Range(ws.Cells(j,7), ws.Cells(i,7)))
                    
                    ' I need to increase the Ticker counter by 1 so the loops goes "down" on the column 
                    Tickercount = Tickercount + 1
                    ' I need to reset or set a new start row for the Tickercount block
                    j = i + 1

                    'I need to close this conditional with the End If statment
                    End If
                'I need to close the for loop with the next i statment
                Next i 
            
    'PART 11: I need to be able to know what is the last non-blank cell in column J to star my Challange table 
        LastRowColumnJ = ws.Cells(Rows.Count,10).End(xlUp).Row
        'I can crete a msgbox to see if its ok
        'Msgbox("The last row in column J is" & LastRowColumnJ)

    'PART 12: I am adding the variables withe the cells value
        GreatestVolume = ws.Cells(2,13).Value
        GreatestIncrease = ws.Cells(2,12).Value
        GreatestDecrease = ws.Cells(2,12).Value
    'PART 13: I need to create a Loop for the challange table
        for i = 2 to LastRowColumnJ
    'PART 14: Create conditional If statment to calculate the greatest total volume. To do this I am creating a loop that will check if in the column Total Stock Volumen the value on the next cell (i +1) is greater than the next value, if it is it will take over the value and populate ws.cells
            If ws.Cells(i,13).Value> GreatestVolume Then    
                GreatestVolume = ws.Cells(i,13).Value
                ws.Cells(4,17).Value = ws.Cells(i,10).Value
            'finish the conditional with an else statement
            Else 
                GreatestVolume = GreatestVolume
            'close conditional with End If statement
            End If
    'PART 15: Getting the Greatest increase. I am gointo to repeat step 14 but instead of total volume column I will use the percentage change column Range L2:L
            If ws.Cells(i,12).Value> GreatestIncrease then 
                GreatestIncrease = ws.Cells(i,12).Value
                ws.Cells(2,17).value = ws.Cells(i,10).Value
            'finish the conditional with an else statement
            Else    
                GreatestIncrease = GreatestIncrease
            'close conditional with End If statement
            End If
    'PART 16: Getting the Greatest decrease. I am gointo to repeat step 15 but looking for the decrease in the data
             If ws.Cells(i,12).Value> GreatestDecrease then 
                GreatestDecrease = ws.Cells(i,12).Value
                ws.Cells(3,17).value = ws.Cells(i,10).Value
            'finish the conditional with an else statement
            Else    
                GreatestDecrease = GreatestDecrease
            End If
    'PART 17: Display values on the challange table corresponding ws cells
        ws.Cells(2,18).Value = Format(GreatestIncrease,"Percent")
        ws.Cells(3,18).Value = Format(GreatestDecrease,"Percent")
        ws.Cells(4,18).Value = Format(GreatestVolume, "Scientific")
        
        'Close loop with next i statement
        Next i 
    'PART 18: I want to adjust the columns width automatically 
    Worksheets(WorksheetName).columns("A:Z").AutoFit

    ' I need to close the loop creted to go through all of the worksheets on the workbook with an Next ws statement
    Next ws
    
 
End Sub
