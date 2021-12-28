# A Complete Analysis of the Stock-Analysis Challenge Module 2

## Overview of the Analysis
After helping Steve generate a dataset that will be beneficial for not just himself but for his parents as well, we refactored the VBA code from Module 2 to give Steve the opportunity to do a little more research for his parents. Which, in short, we will expand the dataset to include the entire stock market over the last few years.
## Results of the Refactored Code
Once we refactored the previous code we used earlier in the module, we were able to see the Run-Time difference decreased. 
In addition to the explainations below, I also added Comments within the code to help identify it's usage. 

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
<img width="1440" alt="Screen Shot 2021-12-27 at 8 56 10 PM" src="https://user-images.githubusercontent.com/95304025/147519391-e7c2ba9f-3302-44d3-a4e2-26b8fd42aeba.png">

For 2018, the code ran in 0.3125 seconds:
<img width="1440" alt="Screen Shot 2021-12-27 at 8 56 36 PM" src="https://user-images.githubusercontent.com/95304025/147519467-2619f33c-05da-41a4-9767-c11f041ac49b.png">


## Summary of Stock Analysis
