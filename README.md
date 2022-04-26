# Stock Analysis
## Overview of Project
 
 During this project, we conducted analysis on data collected on various stocks and their performance in the years 2017 and 2018. The data we received included several points of data for each stock for each date in the respective year, e.g. the opening and closing price of the stock and the volume of the stock on the respective date. The code we wrote used FOR loops in VBA code to loop through the raw data set and collect data for each individual stock for the requested year. The data was then organized into a separate spreadsheet, which we titled "All Stocks Analysis," in which we calculated the total volume of each stock and then calculated the return of the stock using the Total Volume data column. This is very useful information to know for an investor, as it turns raw data that is difficult to make heads or tails of into a small, comprehensive table that shows which stocks had positive/good returns and which had negative/bad returns. We also were able to color code each cell in the "Returns" column to more easily show which stocks were good and bad in the requested year. 
  
## Results

To conduct our refactored analysis, we started by creating an input box to ask the user what year they want to analyze. This is a simple but very important step because it makes our code user-friendly and easy to use to analyze more than one year of data without having to go into the code and rewrite it for another year. 

<img width="480" alt="Screen Shot 2022-04-26 at 9 23 01 AM" src="https://user-images.githubusercontent.com/103055666/165309666-e096ef3a-7f94-4b23-86f3-031dcc13d8b1.png">


We then created an output sheet and formatted it using code. The code we wrote for this step simply assigns several header titles to certain cells for our table to make it easy to read. The next step was to create an "array" of tickers to represent each ticker name numerically. You can see what this step looks like in our code below: 

<img width="231" alt="Screen Shot 2022-04-26 at 8 59 50 AM" src="https://user-images.githubusercontent.com/103055666/165305365-01b90ca2-f5ea-4ee9-affd-38174fd0dd2f.png">

Next, we created a "ticker index" variable to represent the number assigned to each ticker in the array we just created in the last step. We began by setting it equal to zero so that it would start from the first ticker. Later on in the code after looping through the whole table, we will increase the ticker index so that we can move on to the next ticker to collect data. You can see the code to create the tickerIndex and begin it at zero below:

<img width="192" alt="Screen Shot 2022-04-26 at 9 03 16 AM" src="https://user-images.githubusercontent.com/103055666/165305972-283ee3d7-195d-4d4a-b1bf-f3dfad3fbc37.png">

The three variables we want to collect from the data set for each ticker are the volume of the ticker, and the starting and ending price of the ticker. In the first code we wrote, we created individual variables for each of these and looped through the whole data set several times to collect them for each individual ticker. When refactoring, we instead created arrays for each variable that would output the value for each ticker's volume and starting and ending price so that we only have to loop through the data set once for each ticker. See below code in which we created the three output arrays:

<img width="256" alt="Screen Shot 2022-04-26 at 9 07 15 AM" src="https://user-images.githubusercontent.com/103055666/165306708-a7056359-8188-425b-a61c-7cd8ca17ed2d.png">
 We also had to initialize the tickerVolume array to zero, so we could add to it each daily volume in the data set and get the total volume. See below:
<img width="228" alt="Screen Shot 2022-04-26 at 9 09 15 AM" src="https://user-images.githubusercontent.com/103055666/165307069-89d07ea1-2ea0-4c57-a2b2-32c5272664d0.png">

We then looped through the whole data set using a FOR loop to increase the volume of each ticker and calculate total volume for each ticker:

<img width="534" alt="Screen Shot 2022-04-26 at 9 11 39 AM" src="https://user-images.githubusercontent.com/103055666/165307492-19608e12-f502-415d-ab2d-6febe3f0dc86.png">

From there, we used IF statements within the exisitng FOR loop to find the starting price and ending price of each ticker. Once the code hit the ending price, we told it to move on to the next ticker by increasing the tickerIndex variable. See below code that shows this process:

<img width="490" alt="Screen Shot 2022-04-26 at 9 15 07 AM" src="https://user-images.githubusercontent.com/103055666/165308169-b200ce91-ff92-46ea-aee8-f6a47d6af963.png">

Lastly, we used another FOR loop to output the data we calculated for each ticker into the tabel we formatted on the "All Stocks Analysis" sheet we created earlier:

<img width="542" alt="Screen Shot 2022-04-26 at 9 15 19 AM" src="https://user-images.githubusercontent.com/103055666/165308207-4871f942-65a0-41dd-b468-0766278f2f2c.png">

As an extra step, we wrote a few lines of code to format the table and color code it to make it easier to read. We also used a FOR loop and IF statements for this process, coloring the cells of the "Return" column red if the calue was negative and green if the value was positive. This code is below:

<img width="432" alt="Screen Shot 2022-04-26 at 9 17 51 AM" src="https://user-images.githubusercontent.com/103055666/165308681-0b3d23ae-fc6d-4a43-8b28-28f2060c2668.png">

The final results we returned for the year 2017 can be seen in the table below:
#### 2017
<img width="263" alt="Screen Shot 2022-04-25 at 6 49 29 PM" src="https://user-images.githubusercontent.com/103055666/165187138-69c59516-b699-4956-928f-f822889c4dd1.png">


The results we returned for the year 2018 can be seen in the table below:
#### 2018
<img width="252" alt="Screen Shot 2022-04-25 at 6 49 54 PM" src="https://user-images.githubusercontent.com/103055666/165187196-8a621497-c748-4a27-845e-26336855191a.png">

We can very easily see in both tables which stock had a positive and negative return in each year. This will be of great value to anyone who is planning on investing int hese stocks, as they can now develop a better understanding of these stocks' performances in these years. This code can also be used for further analysis on other, more recent years going forward.

## Summary
### What are the main advantages of refactoring existing code?

There are many advantages of refactoring exisiting code, even when a code is already working fine. One benefit is that we can often get a code that is shorter and simpler than our existing code. A simplified code will take up less memory, which is very important when you are dealing with a very long code or have limited storage space on your computer to work with. A shorter code will also take less time to run through the data set you're working with, which is important when you have very large data sets that need to be analyzed. In the long run, refactoring codes will save you time if you have to analyze a very big data set. In the context of our analysis in this project, refactoring the code was very beneficial because it shortened the run time of the code. Steve, our boss, will likely want to use this code for future analysis on data sets other than just the two we were given for this project. If those data sets are comparatively much larger than the ones we analyzed here, it would be very beneficial to have our code run more quickly as it may take a long time for the original code to run through a vast set of data.

### Are there disadvantages of refactoring existing code?

There certainly are disadvantages of refactoring code, mainly that it takes extra effort and that you may not have. As a programmer, you'll have to weigh the benefits described in the last section with how much time you have to refactor the code you have that is already working. If you're running on limited time or have a deadline approaching, it may not be feasible or worth the extra time to refactor a code if it's already working. In this project, our initial analysis code worked quite well, especially because we were working with a relatively small set of data. In the real world, if we were analyzing only a small set of data like this one and had a deadline approaching, it might not be worth it to go back and refactor this code since the run time of the code will be short either way. However, if we knew we were going to be using this code again in the future for a large set of data, this would likely not be our conclusion.
