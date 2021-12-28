# A Complete Analysis of the Stock-Analysis Challenge Module 2

## Overview of the Analysis
After helping Steve generate a dataset that will be beneficial for not just himself but for his parents as well, we refactored the VBA code from Module 2 to give Steve the opportunity to do a little more research for his parents. Which, in short, we will expand the dataset to include the entire stock market over the last few years.
## Results of the Refactored Code
Once we refactored the previous code we used earlier in the module, we were able to see the Run-Time difference decreased. 
In addition to the explanations below, I also added Comments within the code to help identify its usage. 

A. The first adjustment we made was set the ‘tickerIndex’ to equal zero before looping through the rows. The ‘tickerIndex’ will be used to access the correct Index for the different arrays in this code:
-	The ticker Arrays 
-	tickers
-	tickerVolumes
-	tickerStartingPrices
-	tickerEndingPrices
<img width="385" alt="Screen Shot 2021-12-27 at 6 48 18 PM" src="https://user-images.githubusercontent.com/95304025/147514220-2dd6e08d-1e0f-40ed-8374-3738fca768c9.png">

B.	Next, we created 3 output arrays.
The ‘tickerVolumes’ array is a Long data type and the ‘tickerStartingPrices’ and ‘tickerEndingPrices’ arrays are Single data types.

<img width="577" alt="Screen Shot 2021-12-27 at 6 59 25 PM" src="https://user-images.githubusercontent.com/95304025/147514570-4f4858da-fbd0-4a4e-af0e-c7b8a3105ff0.png">
C. Once the 3 output arrays were created, we then created a loop to initialize the ‘tickerVolumes’ to zero and if the next row’s ticker doesn’t match, increase the tickerIndex.

<img width="563" alt="Screen Shot 2021-12-27 at 8 07 23 PM" src="https://user-images.githubusercontent.com/95304025/147516994-be923a35-5135-4d27-98a1-0221fa77d2e0.png">

D. Then we went back into the loop to increase the current ‘tickerVolumes’ variable and added the ticker volume for the current stock ticker. We also went in and created an if-then statement to ensure the current row is the first row with the selected ‘tickerIndex’. If yes, then assign the current closing price to the ‘tickerStartingPrice’ and ‘tickerEndingPrice’.

<img width="572" alt="Screen Shot 2021-12-27 at 7 49 43 PM" src="https://user-images.githubusercontent.com/95304025/147517130-38ae6d3b-9dd8-4770-93db-dc8bf43d6748.png">

### We're almost done...I Promise :)

E.	Now we did some formatting to make it easier to visualize which stocks did well and which ones didn’t. In this part of the code, we made the positive returns green and negative returns red.

<img width="479" alt="Screen Shot 2021-12-27 at 8 30 01 PM" src="https://user-images.githubusercontent.com/95304025/147517950-cc63c368-383d-4582-8ee5-f1969ff0f6bf.png">



Once I reviewed the code to make sure it was correct, I went ahead and ran the Stock Analysis for the years if 2017 and 2018. We were successful in matching the refactored code to the original code in this Module! YAY!! 

For 2017, the code ran in 0.3203125 seconds:
<img width="1435" alt="Screen Shot 2021-12-28 at 5 29 56 PM" src="https://user-images.githubusercontent.com/95304025/147611944-03b4b0a4-32d0-4b88-8368-7846aa1a0923.png">


For 2018, the code ran in 0.3125 seconds:
<img width="1438" alt="Screen Shot 2021-12-28 at 5 30 25 PM" src="https://user-images.githubusercontent.com/95304025/147611964-b9a4e376-76b1-497b-9b56-088b9b8b6a87.png">

##


# Summary of Stock Analysis

## 1. Explanation of the Advantages and Disadvantages of refactoring any Code.

**Advantages:**

**Keeps the code Simple and Upgradeable.**  I think when the code is simple and clean, it helps with keeping it up to date and easier to make new changes.

**Uses less Memory**

**Maintainable.**  After refactoring, the code is fresher, easier to understand or read, less complex and easier to maintain.

**Easier to Read for anyone who uses the code after it was originally created.**  Especially for New users to VBA (or code in general) it can make understanding the original code easier to comprehend and make needed alterations.


**Disadvantages:**

**Duplication**  It can be easy to duplicate several procedures and it increases the chance of making mistakes especially using copy & paste.
**Time Consuming**. Especially it comes to efficiency, it can take a lot of time to ensure you are inputting the correct code etc.

## 2. How do these pros and cons apply to refactoring the original VBA script?

Refactoring the original VBA code provided faster run times as compared to the original code. I believe that since the dataset was smaller, the performance increase was minimal and when weighed against the extra time it took to refactor the original code it was not of much value, besides making the code cleaner.

