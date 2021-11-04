# Stock-Analysis
## Overview of Project
Through the utilization of Visual Basic for Applications, or VBA, we can program and interact with excel to find out the difference between 2017 and 2018 stocks of an energy company and if it is worth to invest. We programmed subroutine, loops, and timer in the original script and refactored script to compare and contrast which program ran the VBA faster.

## Results
The excel sheet for 2017 and 2018 green stock data table was downloaded and compared between the two years. The screenshots of the two years are shown as below. Each data table shows the same columns: "Ticker", "Total Daily Volume", "Return". The "Ticker" is the shortend version of the stock name. The "Total Daily Volume" is the amount of stocks traded daily. The "Return" shows the percentage earned back from the stock every day. The color green is represented as positive earning and the color red is the opposite, which is the percentage loss per day. This was achived through the conditional formatting in VBA, which immediately shows the readers which stock was doing good and bad each year.

![2017 Refactored](https://user-images.githubusercontent.com/92328984/140254414-7c08ed84-b33f-4992-9afe-cf8dfb39c992.jpg)
![2018 Refactored](https://user-images.githubusercontent.com/92328984/140254791-f71b01f5-cac3-4006-8eb8-b1c2a371fd47.JPG)

In 2017, all the energy stocks did well, especially DQ with 199.4% return. However, TERP was fairly under valued with -7.2% return in 2017. After a year, there was a big turnout for RUN and ENPH, only two companies to have positive turnout compared to the rest of the green stocks. All the rest of the companies except for RUN and ENPH with all negative returns.  

The purpose of refactoring a list of codes or scripts is to shorten the time of how the code performs. Below screenshots are the message box of excel telling how long it took to code for each year.
#### 2017 Original vs. Refactored script
![2017 Original Script Time](https://user-images.githubusercontent.com/92328984/140257129-a528e50f-64dd-4860-9c9e-f03ccec833fc.JPG)
![2017 Refactored Script Time](https://user-images.githubusercontent.com/92328984/140257717-f0377026-9fa0-474d-99ee-e5111f14b202.JPG)

#### 2018 Original vs. Refactored script
![2018 Original Script Time](https://user-images.githubusercontent.com/92328984/140257146-311b38cf-ad44-48bb-87d1-8039853da1bb.JPG)
![2018 Refactored Script Time](https://user-images.githubusercontent.com/92328984/140257779-70707fd0-461b-4f80-84d5-8a1fc0390a52.JPG)
As shown in the screenshots, the refactored script times were consistently below 1 second mark while the original script had some above 1 second mark and some below 1 second mark due to depending how the VBA was processing the data that would either shorten or lengthen the time it takes.

## Summary
1) What are the advantages or disadvantages of refactoring code?

   The advantages of refactoring a code are increasing efficency of VBA running to solve the code. Long down the road, the script    will be longer and more complicated that it would take longer than the challenge we had. In order to decrease the timing, the      data analysts would take a look into the original code and figure out a way to make it more efficient and shorten the time to      run the code. Another advantage is that it cleans out any blemishes from complicated code that would make a reader confused but    later making sense since the code works at the end. A disadvantage or a flaw I could think of when refactoring a code is the      cleaning up part. If the code already works, why would anyone want to go back and clean up to make it better? Also, it would      take a lot more time to brainstorm and come up with an efficent way to refactor the already written codes.
   
2)How do these pros and cons apply to refactoring the original VBA script?
  
  The biggest pro was that the code efficently ran without a problem. The con was the part where someone had to clean up and make   the code look nice for anyone to read, like commenting every single command line to tell or instruct what it needs to be done     with the code.



