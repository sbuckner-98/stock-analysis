# Stock Analysis with VBA and Excel

## Project Background:
In this challenge, we have edited, or refactored, the Module 2 solution code to loop through all the data one time in order to collect the same information that we did previously in the module, but ideally faster and more efficiently.

This code, using VBA and Excel, analyzes the performance of multiple green energy stocks in the years 2017 and 2018 to help a client determine what to invest in.

## Results
_VBA Analysis, 2017_

In 2017, 11 of the 12 stocks the client wished to analyze returned positve results, some more impressive than others.
![2017_VBA.png](https://raw.githubusercontent.com/sbuckner-98/stock-analysis/main/Resources/2017_VBA.png)

We can see that stocks ENPH, DQ, FSLR and SEDG each saw returns of over 100%. These patterns did not continue into 2018, however.


_VBA Analysis, 2018_

![2018_VBA.png](https://raw.githubusercontent.com/sbuckner-98/stock-analysis/main/Resources/2018_VBA.png)

In this year, only ENPH and RUN returned positive results, with ENPH being the only consistently high-performing stock in each year. For this reason, we suggest that our client invest in ENPH.

## Summary

_**1. What are the advantages and disadvantages of refactoring code?**_

**Advantages:**
1. Refactoring code makes the code easier to read, both for yourself during future edits, and for your clients.
2. Refactored code can execute tasks more time-efficiently and with greater flexibility over different datasets.
3. Logical errors and/or bugs within the code may be easier to spot in refactored code, making it easier to fix.

**Disadvantages**
1. Refactoring code can be a time-consuming process, especially when the logic of the initial code must be changed.
2. Refactoring code can lead to the creation of new errors or bugs that did not exist in the original code.

_**2. How can we see these advantages and disadvantages in our refactoring of the original VBA script?**_

The new, refactored VBA script indeed does run faster than the original script, making the code better suited for being scaled-up to projects analyzing more data given its ability to do so in a timely matter. The refactored code is also easier to read and understand to an observer, given its simpler logic and lack of repetitive code.

However, the non-refactored and refactored code accomplish the exact same tasks, and for developers working on a time crunch, it may not be viable to refactor already successful code.
