# stock-analysis
Week 2 challenge

Written Analysis of Results

Overview of Project
  This project seeks to compare and analyze the different provided green stocks for potential investment. The analysis is based on the total volume traded and price variation of the stocks in the catalog to determine the ones with more potential to invest in.
  
Results
  Comparing the years performance of all the stocks in the list we can convey that the price variation was mostly positive in 2017, since the ending price was higher than the starting price that the year began with. While most of the stocks in 2018 ended the period with an ending price lower than the startting price for the period.

Code examples
  We used the follwoing functions in the VBA code to format the table showing a green shade for positive variation in the stock proce and red shade when the variation of the stock price is negative.
  
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
 Graphic examples
  This results in the follwoing image in the tables that help us identify that 2017 was a much better year for the price variation than 2018.
  
2017

![2017 Table.png](/Users/Erick/Desktop/Bootcamp/Class Folder/Stocks Analysis/Resources/2017 Table.png)

2018

![2018 Table.png](/Users/Erick/Desktop/Bootcamp/Class Folder/Stocks Analysis/Resources/2018 Table.png)

  Here we can see a graphic representation with the colors symbolic values, green for positive and red for negative.

Summary
  The advantage of using a refactored code is the formating and patterns help us and the computer to use similar operations to calculate different values. Essentially with the same code formatting we could loop through the signifficant facts for our analysis and that results in fewer time writing the code and fewer time for the program to run. This is basically the same loop for the three arrays that gave us the total trade volume and return, which we could access by using an index.
  
      For i = 0 To 11
        tickerIndex = ticker(i)
        tickerVolumes = 0
        
      If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerVolumes Then      
        tickeStartingPrices = Cells(j, 6).Value          
      End If
  
      If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerVolumes Then      
        tickeEndingPrices = Cells(j, 6).Value
        tickerIndex = tickerIndex + ticketEndingPrices
      End If
      
  The results in time are perceivable even if this is not such a great amount of data. As we can see the unrefactored code takes nearly four seconds and after refactoring the time takes fewer than one second as can be compared in the below images.
  
Before

![Previous VBA_Challenge_2017.png](/Users/Erick/Desktop/Bootcamp/Class Folder/Stocks Analysis/Resources/Previous VBA_Challenge_2017.png)

After

![VBA_Challenge_2017.png](/Users/Erick/Desktop/Bootcamp/Class Folder/Stocks Analysis/Resources/VBA_Challenge_2017.png)

Disadvantages
  Even though this is a great advantage to consider, when we use refactored code, there are potential bugs that can be difficult to figure out if there is any hard coding or magic number left in the code. Which in result becomes a benefit of the non refactored code. It is clear and more simple to manipulate and preferred for smaller datasets.
  
