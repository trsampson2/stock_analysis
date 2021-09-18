# Stock Analysis
Refactoring Stock Analysis Code and Measure Performance

## Overview of Project
After preparing an interative workbook for a client, he now would like to expand the workbook analytical capability.  Using code, I will need to expand the capability of the code to include any dataset he would want to add, i.e. more years, more stocks.

## Results

### Comparison of 2017 and 2018
2017 was a great year for investments. All but one of the stocks had a positive return.  The positive returns were mostly in the triple digits. It was a very good deal. In 2018, on the other hand,  the stocks tanked.  All but two of them had negative returns.  TERP was consistent both years. Their negative returns were under 10% for both years. ENPH and RUN had positive returns both years. RUN did exceptionally well on in 2018.  
### Execution Times
#### Execution of code for 2017 data
The initial code from Stock_Analysis ran in 1.265625 seconds.

![Stock_Analysis_Initial_2017](https://user-images.githubusercontent.com/86331812/133895059-e72fcf5a-726f-4db5-8b97-c28f67cc658b.png)


The refactored code for 2017 ran in 1.132813 seconds. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/86331812/133895067-dca2cdf6-41fe-4350-9018-213a5dbb672d.png)


#### Execution of code for 2018 data
The intial code from the Stock_Analysis ran in 0.9765625 seconds.

![Stock_Analysis_Initial_2018](https://user-images.githubusercontent.com/86331812/133895068-103452ef-29d6-4206-af69-f05f8e515fe4.png)


The refactored code for 2018 ran in 1.164063 seconds. 

![VBA_Challenge_2018](https://user-images.githubusercontent.com/86331812/133895069-9093f64b-e290-4ac7-8b68-b6bdc11cd236.png)


## Summary
### Pros and Cons of Refactoring Code
The pros of refactoring code are:

    1. Makes code more adatptive.  As shown with the request of this client.  He will be able to use the this code across a             broader range of data.
    
    2. Makes code easier to read by organizing the code. 
    
    3. Can help code run faster.
    
The cons of refactoring code are: 
    
    1. Can slow down the development.
    
    2. Can increase the cost of development.
    

### Pros and Cons of Refactoring this Code
The pros of refactoring this code was to make it more adapative. The client would be able to use the worksheet across a broader range of data.  It also grouops together the the data to help it run more efficiently. 
The cons of refactoring this code was making the code more complicated to read.  Though the code groups sections together, it can be harder to read. 
