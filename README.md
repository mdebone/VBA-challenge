# VBA-challenge

I first got hung up on a missreading of the 'change' value that was requested, I thought that it was an average of the difference between open and close for each day of trading rather than a simple measure of the first instance and the last instance of the price thus making an 'annual' value. I had assumed that since it was stock volume they were looking for the average volitility.

So I got hung up on the fact that there are 261 instances of each 'trade date' per unique ticket and went down a rabbithole of how to pull the 'A' values from the 2nd and 261st, and then the 'AA' 262nd and 523rd values to get the start price and the end price, rather than what I should have done which was include them in the table to begin with and reference them when they occurred pulling that as a seperate subroutine. 

Okay took a nap, worked through the Kickstarter excel workbook project and got a handle on it this afternoon, it was interesting to see in the space of a week how when I wanted to do actions in excel I was reverting to VBA rather than formulas and its only been a week. So interested to see where this goes, also, for the life of me I do not understand the color coding portion of excel.
