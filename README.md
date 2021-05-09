# DQ SOLAR ENERGY
## Steve’s graduate from finance degree and Steve’s parents are going to be their first client, they are passionate in green energy and they want to invest in DQ solar energy, Steve needs to analyze the green energy stocks, he created an excel file to access all the data. He wants to used VBA to analyze. Steve wants to find the total daily volume and yearly return for each stock. The parents of Steve want to know how actively DQ was traded in 2018, because they believe that if a stock is traded often, the price will reflect the value of the stock.
### An small analisis had been made to understand how the values where sort by and the specifications of the worksheets used in this analysis.
![1](https://user-images.githubusercontent.com/82455263/117584624-a4899480-b0d3-11eb-9199-e805522834b4.png)

I start the code mentioning that the data is going to show in the worksheet “All Stock Analysis Refactored” but is going to be refactored. 
             1.	     Sub AllStocksAnalysisRefactored()
I did activate the code with Dim “starting time and ending time”.
             2.   	Dim startTime As Single
                    Dim endTime  As Single
I did activate the “input box” to show which year I need to run the analysis on.
             3.    	yearvalue = InputBox("What year would you like to run the analysis on?")
I did activate the worksheet “All stocks analysis” worksheet. And mentioned that in the cell A1 is going to show “the year” of all stocks I pick to run on the analysis.
             4.   	Worksheets ("All Stocks Analysis"). Activate
             5.   	Range("A1"). Value = "All Stocks (" + yearvalue + ")"
    
I started mentioning the values to print in the columns 1,2,3 of the row 3 from the “all stocks analysis “worksheet.  
             6.	    Cells(3, 1).Value = "Tickers"
                    Cells(3, 2).Value = "Total Daily Volume"
                    Cells(3, 3).Value = "Return"

 ![2](https://user-images.githubusercontent.com/82455263/117584644-c2ef9000-b0d3-11eb-84a5-9e3d8b2a4442.png)

I started mentioning the values to print the tickers, in total 12, starting from the value (0) until the value (11) in the “All stocks analysis” worksheet.  
            7.	     Dim tickers(12) As String
                     tickers(0) = "AY"
                     tickers(1) = "CSIQ"
                     tickers(2) = "DQ"
                     tickers(3) = "ENPH"
                     tickers(4) = "FSLR"
                     tickers(5) = "HASI"
                     tickers(6) = "JKS"
                     tickers(7) = "RUN"
                     tickers(8) = "SEDG"
                     tickers(9) = "SPWR"
                     tickers(10) = "TERP"
                     tickers(11) = "VSLR"
                     
I did activate the year value “worksheet” to be able to pick the year from the 2017 and 2018 worksheet.
           8.	      Worksheets(yearvalue).Activate

![3](https://user-images.githubusercontent.com/82455263/117584648-c5ea8080-b0d3-11eb-92e3-7501c2f4e52e.png)

I did activate the option to get the number of rows to loop over
           9.    	RowCount = Cells(Rows.Count, "A").End(xlUp).Row
           
I Created a ticker Index that will start from “Zero.”
          10.	   Dim tickerindex As Long
                 tickerindex = 0
                 
I Created three output arrays : Ticker volumes, Ticker starting prices and Ticker ending prices  
and also I created a for loop to initialize the ticker Volumes from zero to 12. 

         11.	   Dim tickervolumes(12) As Double
                 Dim tickerstartingprices(12) As Single
                 Dim tickerendingprices(12) As Single  

         12.     For i = 0 To 11
                 tickervolumes(i) = 0
            
I did created a loop over all the rows in the spreadsheet.
        13.	    'Worksheets(yearvalue).Activate
                 For i = 2 To RowCount

![4](https://user-images.githubusercontent.com/82455263/117584653-cdaa2500-b0d3-11eb-96e6-3055b782396a.png)

I did create a code to increase volume for current ticker and also a conditional for ticker volumes, ticker starting prices and ticker ending prices.

         14.	tickervolumes(tickerindex) = tickervolumes(tickerindex) + Cells(i, 8).Value

         15.	If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
              tickerstartingprices(tickerindex) = Cells(i, 6).Value

              If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
              tickerendingprices(tickerindex) = Cells(i, 6).Value

         16.	tickerindex = tickerindex + 1

![5](https://user-images.githubusercontent.com/82455263/117584662-da2e7d80-b0d3-11eb-9bb7-7e79781ac0f7.png)

I did create a loop through your arrays to output the Ticker, Total Daily Volume, and Return.
         17.	For i = 0 To 11
              Worksheets("All Stocks Analysis").Activate
              
             Cells(4 + i, 1).Value = tickers(i)
             Cells(4 + i, 2).Value = tickervolumes(i)
             Cells(4 + i, 3).Value = tickerendingprices(i) / tickerstartingprices(i) - 1

I did Formatted all the information inside the worksheet “ All Stocks Analysis, that’s why I had to activated first and format the font style, the line style and the fit for the columns 

          18.	Worksheets("All Stocks Analysis").Activate
             Range("A3:C3").Font.FontStyle = "Bold"
             Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
             Range("B4:B15").NumberFormat = "#,##0"
             Range("C4:C15").NumberFormat = "0.0%"
             Columns("B").AutoFit


![6](https://user-images.githubusercontent.com/82455263/117584667-dd296e00-b0d3-11eb-90d1-b6a5c511ecf2.png)
I limited the format from the data row starting at 4 and the data row ending at 15, using conditional to use the colors green and red .

         19.	dataRowStart = 4
                    dataRowEnd = 15

            If Cells(i + 4, 3) > 0 Then
               Cells(i + 4, 3).Interior.Color = vbGreen
            Else
               Cells(i + 4, 3).Interior.Color = vbRed
            End If
            
 I did activate the “Msg Box” to show which year I pick to run the analysis on and how many seconds it took to analyze

         20.	endTime = Timer
              MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)
              
And finally i make sure i " end sub" to finish the code 

## Advantages and disadvantages of refactoring code in general 
The advantage is when the code es refactored, is easier to understand and that gives more options to the code to add other functions. 
The disadvantage is that if one mistake is made while refactoring it will take more time solving the problem.

## Advantages and disadvantages of the original and refactored VBA script 
Both have a clear goal to arrive, but the refactored it was more easier to read and more order.

#VBA_Challenge_2017.png 
![2017 worksheet](https://user-images.githubusercontent.com/82455263/117586752-bbce7f00-b0df-11eb-9994-70fc6fd6ff3d.png)


#VBA_Challenge_2018.png
![2018 worksheet](https://user-images.githubusercontent.com/82455263/117586762-c1c46000-b0df-11eb-8e4e-da78f029a0f5.png)
