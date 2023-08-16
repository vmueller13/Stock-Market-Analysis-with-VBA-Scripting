# Excel VBA
<ins>Project Overview</ins>
-----

This project utilizes VBA scripting to analyze generated stock market data. The main ask for this challenge was to creat a script that loops through all the stocks for one year adn outputs the following information:
* The ticker symbol
* Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
* The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
* The total stock volume of the stock.

<ins>Processes and Technologies</ins>
-----


To complete the ask, I started with a for loop that would cycle through each worksheet and print out the unique ticker symbols. To get the other values from the data and have them populate in the proper columns, I started by assigning the variables opening and closing values and use the assigned values in a new count 'j' to determine the change over the year. Next, I compared each of these variables to the current and previous cell and reassigned the value based on which one was larger. Eventually, the largest values were assigned to the variable and the second summary table populated. When run from start to finish, the code loops through each sheet of the workbook.

<ins>Challenges</ins>
-----

The biggest challenge I ran into on this project was setting up the original for loop to calculate the sum of the total stock volume using the `Count = Count + 1` formula. Once I configured this part of the formula correctly, I was able run the code correctly and completely.
