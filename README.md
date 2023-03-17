# vba_challenge


The first ask for this challenge was to go through the ticker list and have each individual ticker populate in a separate column. Then to have the yearly change, percent change and total stock volume populate in the columns as well.

To do this I started with a for loop that would run for each worksheeet. This For loop looked through the 'I' column and looked for the row that wasn't equal to the one before and then it would print out the ticker symbol.

To get the other values from the data and have them populate in the proper columns, I started by assigning the variables opening and closing and use those in a new count 'j' to determine the change over the year.

Next, I named the titles for the second summary table and set up the variables for percent increase, percent decrease and greatest total volume. I compared each of these variables to the current and previous cell and reassigned the value based on which one was larger. Eventually, the largest values were assigned to the variable and the second summary table populated.

For the conditional formatting, if the value was greater than 0 it was green and if it was less than 0, it was red.

Finally, this entire code would then run on the next worksheet.