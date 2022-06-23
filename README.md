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
![Original_2017](https://user-images.githubusercontent.com/107159218/175424214-5d56694c-df9e-42a8-8e24-b044b87a07d1.JPG)
![Refactored_2017](https://user-images.githubusercontent.com/107159218/175424228-57796fa6-dc00-4ce4-9dd1-71ed723ef993.JPG)
![Original_2018](https://user-images.githubusercontent.com/107159218/175424245-0a4ffc00-7d90-4598-97ff-1f88e2cf7faf.JPG)
![Refactored_2018](https://user-images.githubusercontent.com/107159218/175424255-ad330050-13ac-4d59-873a-a088da240984.JPG)

### Summary
1.  The obvious advantage for the refactored code is that it saves overall analysis time as it involves less lines of written code and uses less memory. The overall idea of refactoring a code is to make an existing code more efficent. The latter statement could also be considered as a disadavantage as we need to formulate a code or a psuedocode in advance. I would not think of creating a refactored code if this was the proverbial 11th hour as it would not get tested enough and prove bug riddled
2. Comparing the advantages and disadvantages of the original and refactored code. 
    The original code though lenght was simple to write and understand (for a novice like me). This would definately be the way to go till I gain suitable amount of experience. This would be a good advantage of this code and it helps developing the building blocks for writing complex codes. Since it had a lot of code lines it would be slower and would required a good memory investment and would not be ideal for larger data sets. 
    The refactored code, is short and concise and can be run faster due to lesser memory requirements. I however had to see the original code before I can work on refactoring it. This would require some additional time commitment on my end. I see this as a disadvantage but could just be an incovenience. 



