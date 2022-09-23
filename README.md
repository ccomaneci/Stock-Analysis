# Stock-Analysis

## Overview of Project

In this project we learned how to use VBA inside of Excel and practiced our new skills on the dataset titled “Green Stocks”. This dataset contained reporting information on stock trading for years 2017 and 2018. With the help of VBA, we were able to condense that data into worksheets and run coding that could yield if there were negative or positive returns for those stocks. So, what is VBA? Visual Basic for Application is a programming language that is used within Microsoft Office [^1]. This programming application allows for us to turn words into code and turn code into “results”.


## Results 

 In order to turn code into outputed data we needed to first learn how to pair words and phrases together that the VBA system understood. For instance, in this project, we wanted to know the return for a stock called DQ. So instead of looking up all the stock information for DQ and calculating it by hand, we used a code to tell us what that return would be. First, we started out by creating a worksheet within excel for the results, or the data output, to be placed. Then we had to tell the program the name of the worksheet we created for those results. We did this by using the formula; worksheets(“DQ Analysis”).Activate [^2]. The word worksheets found in the beginning of the code tells VBA that we are about to pull from or place data on a specified worksheet. The “DQ Analysis” is the name of the worksheet we want to interact with, and the word activate tells it to open that worksheet [^2]. From there we were also able to tell VBA where we wanted to put a title on our page using the formula; Range(“A1”).Value =”DAQO(Ticker:DQ). The Range(“A1”), told the program that I wanted to put data in row 1 column A of the worksheet and that the title or Value, that would be going in that cell is supposed to Ticker: DQ [^2]. These are just a few of the very many ways we were able to use this programing platform to help analysis the dataset given to us. 


## Summary 

Using VBA we learned how to write code, pair codes with others or write codes to stand alone. Within this application, we learned that not only can you pair multiple codes within a module together, but you can also utilize them by stacking them on top of each other. We can stack the data in different subs, however, when we do this, it can cause the run time to be a bit longer. When we refactored our data under one sub it allowed our run time to be reduced by almost 3 whole seconds. To calculate the run time, we used the following code:
startTime = timer
endTime = timer

then we created a message box that would pop up and allow us to select a year and compute how long it took VBA to run all the code and produce results [^3].
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
[^3]

![VBA_Challenge_2018](https://user-images.githubusercontent.com/112769590/191878926-edc43a37-b039-49c9-89aa-16c4ce84d46c.png)

![example of output after running macros](https://user-images.githubusercontent.com/112769590/191878707-a6242477-b626-42a7-a064-c903c532e364.png)

![example of the If statment we learned to write](https://user-images.githubusercontent.com/112769590/191878881-4826716e-978a-4236-a686-d4d03bfa34b8.png)


#### footnotes
[^1] Module 2.0.1 Make your way with VBA
[^2] module 2.2.1 Create a worksheet for your analysis 
[^3] module 2.5.3 Measure code performance 

