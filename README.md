# Challenge-2-VBA-
1.	Overview Of the Project-
The purpose of this analysis was to take the information from 2017 and 2018 stocks and use the visual basic application to then apply the Macros to code and then be able to get results from the stocks. I think one of the most critical portions of putting this project together was being able to identify the variables from each year. Once I was able to work on each macro code and identify the data type that stood out in running the code it made a better visual to help with the Integer data. I think for finding the integer and long whole number throughout this code was a challenge. I also feel that when running the codes for both spreadsheets on 2017 and 2018 I came across some confusion in which I thought were Boolean where some values came across as being true and false. Once completing the task of the steps, the main goal of this analysis was to have from both sheets of 2017 and 2018 the high and the low of the stock volume.
2.	Results-
There were many different results that I got from running the codes on the VBA with Macro. With having some trial and error and practice with Macro and running the codes I was able to get results once I was able to look at the loops and the array from each spreadsheet and then at towards the end run the button of both sets of my data. The Ticker was the start of the key component of the codes ran and then having the year analysis. The comparisons on each year data set were to see the total daily volume and then the return percentage. When in macro and running the code the top twelve tickers came up which then help me indicate my key components for the outcomes. For my stocks that were in 2017 all of the numbers came out to be a positive percentage except for one stock went into the negative. That stock was TERP and the return percentage on that was negative 7.20%.  The highest stock was DQ for a return percentage of 199.40%. Those were the key factors for the stock market results in 2017. This is an example of how I started to run my code for 2017, for also developing my ticker data.
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
Now, when I was working on running my macro for all stocks 2018, I followed the same steps and try to repeat what I did to then gather new results for this data set. I hope syntax set of rules was arrange correct and carried over to get the most accurate result. For 2018 ten out of the twelve stocks had more of a negative return on their percentage number. Which concerns me because 2018 was clearly not a good year if there were so many stocks that went in the negative numbers. My lowest stock result was DQ and that was at -62.60% return percentage. My highest stock in 2018 was RUN at 84.0% return percentage. This is an example of how I started to run the code for sheet 2018.
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"
3.	Summary- 
For this VBA project there are advantages and disadvantages of refactoring the codes. The advantages that I found while refactoring the code was the fact that you could understand and grasp the concept and see the whole picture of the data when it is organized and reduce the confusion and information that is not needed to highlight the important key factors. The disadvantages of refactoring codes are it can get very confusing and time consuming to go back and reset to factor in the codes to rerun them. From what I learned in this part of the project is when refactoring it also made my data set seem to become messing on getting the big idea of information that was then supposed to be highlighted so you would know the outcomes that were found. Pros and Cons with visual basic application this goes hand in hand in my opinion from the project because in a way you need to refactor the code several times to get the correct results and be able to clean it up so you know what you are looking at and the highlights of the key factors. The cons are that it does get very messy and confusing maybe I just felt that way from this project because all this information is new to me and is a whole new language. I feel like with anything in time and practice ill be able to understand and run the code faster and know how to achieve getting the best results for the overall of the project. VBA is a very cool platform in my opinion because it tells excel how to calculate in a spreadsheet a large amount of data. 

