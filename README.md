# Adio-Olufemi-Peter

WRITTEN ANALYSIS OF RESULTS

1.	The Overview of the project: 
       At the initial stage of the project, Steve’s parent only wanted to be informed how DQ was actively traded in 2018. Whether it performed well or not. Therefore, we use VBA Analysis to determine some codes that eventually produced the total daily volume and for the return for the year. The project also determined total daily volume for each of the tickers and their yearly returns perhaps Steve may want to look at the entire stocks in the future. Therefore, the codes created for DQ were reused for other tickers.

2.	Results:

Below codes were used to make the headings for each column and we use initialize the array for the entire tickers

Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(11) As String
    
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
    

We also used  
Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) – 1

And for 2017 All stocks Analysis we found out that only the following tickers had more than 100% return values for their investments: DQ = 199.4%, ENPH = 129.5%, FSLR = 101.3%, and SEDG = 184.5%, moreover, we still have others that had 50% and above of yearly returns for the investments. These are: JKS = 53.9%, as well as VSLR = 50%. Only few did not do well for year 2017.

More so, for all Stocks Analysis for the year 2018. It was a disaster year to invest as almost all the stocks recorded too many losses. In fact, only ENPH and RUN recorded very high returns with 81.9% and 84.0% respectively. All others have negative returns. Therefore, all stocks analysis for the year 2017 and 2018.

Moreover, when we compare the original scripts, the execution time for the refactored scripts was faster, whereas the execution time for the original scripts was slower.

3.	#Summary:

(a)	Advantages and disadvantages of Refactoring Code.

The following are the advantages of refactoring code:
1.	It saves time by executing the program faster
2.	It makes it simplified for easy comprehensibility.
3.	Refactoring code also improves the design, structure and implementation of the software.
4.	It also makes it easier to find bugs.

   The disadvantages of refactoring code
1.	When refactoring code is imprecise, it could introduce new errors and bugs into the code.
2.	 The time to complete the process might be too much.
3.	It may also land someone into a situation where he or she has no idea of where to go.
(b)	How do these pros and cons apply to refactoring the original VBA script?
   Pros of refactoring original scripts
Refactoring the original scripts should be done only when:
1.	 The Chances of enhancement are very high.

2.	 Fixing bugs take too muchs efforts.


3.	 When code smell is detected etc.


Cons of refactoring original scripts

Original scripts should not be refactored in the following cases.

1.	When the deadline is very close, and planning is on-going for new development.

2.	We should not do refactor if there is no time to test the refactored code before being released.

3.	We should not also refactor stable code

4.	We should not also delay factoring because it contains big mess.
![image](https://user-images.githubusercontent.com/108506115/200132170-08423754-ded3-426f-82f8-a13e64c08849.png)

