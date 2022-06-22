# VBA-challenge
# Overview of the Project

The goal of this project is to help Steve a fresh graduate help his first customers who, also happen to be his parents, research research the stock market data over from 2017-2018 for investment. 
We had previously helped Steve design a VBA code to provide data for total volume and returns for a green energy entity(DAOQ). Our objective now would be to provide a refactored code to cover the entire stock market over 2 years.

## Results

1. In step 1a) we defined a variable "tickerIndex" which, was used to access correct index across multiple different arrays

2. In 1b) We have created three different output arrays

3. 2a) Creates a loop to initialize the tickerVolumes to zero

4. 2b) Creates a loop the goes over all rows of the worksheet

5. 3a-c) Sets a conditional inside of the for loop to increase tickerVolume, set tickerStartingPrice and set tickerEndingPrice

6. 4) Creates a loop through the arrays to determine output for Ticker, TotalDailyVolume and Returns

7. 5) Create the formating guidelines for the worksheet

8. Here's the output of both the original and the refactored code

This is the comparative performance of the original and the refactored code for the two years:
Original 2017: 1.4843 sec, Refactored 2017: 1.3437 sec
Original 2018: 1.7187 sec, Refactored 2018: 1.3984 sec 

### Challenges

For both the analysis the only major challenge was my unfamiliarity with excel functions. For the first analysis generating pivot tables and charts was fairly straight forward. One issue that I face was when filtered using the newly created "Years" column. I would not see months and I had to use the "Date ended conversion" column instead as it breaks down into quaters and months.

For the second analysis, my only complain would be typing the Goals column. I am sure that one could write a script to fill in this column, especially if there were more entries in this column than the 12 listed.

## Result

### Theater Outcome Based on Launch Date

The two conclusions that can be drawn from this analysis are;
1. Months May-July are most favorable to start a crowdfunding project. The success to fail ratio for these three months are quite similar (65%+/-2) and eventhough May would look the better month June would be not a distant second choice.
2. Setting up a crowdfunding project should be avoided in the month of December as there's a equal chance that it may fail. 

### Outcome Based on Goals

The clear conclusion here is not to set a crowdfunding goal between $45,000-$49,999 as it has a 100% chance of failure. 

## Limitation

1. One limitation is that this data is not current and relies on information collected between 2010-2017. To get a more realistic picture data generate from 2015-current should be considered.
