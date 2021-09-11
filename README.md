# VBA-challenge

I first got hung up on a missreading of the 'change' value that was requested, I thought that it was an average of the difference between open and close for each day of trading rather than a simple measure of the first instance and the last instance of the price thus making an 'annual' value. I had assumed that since it was stock volume they were looking for the average volitility.

So I got hung up on the fact that there are 261 instances of each 'trade date' per unique ticket and went down a rabbithole of how to pull the 'A' values from the 2nd and 261st, and then the 'AA' 262nd and 523rd values to get the start price and the end price, rather than what I should have done which was include them in the table to begin with and reference them when they occurred pulling that as a seperate subroutine. 

Okay took a nap, worked through the Kickstarter excel workbook project and got a handle on it this afternoon, it was interesting to see in the space of a week how when I wanted to do actions in excel I was reverting to VBA rather than formulas and its only been a week. So interested to see where this goes, also, for the life of me I do not understand the color coding portion of excel.

So I had the hardest time with this, but I put in the hours and got it to look reasonably correct from the image posted in GitLab. It was the looping  thru all of the stock names on the sheets, one can't hard code it obviously because while there may be 797711 stock names in sheet C, that doesn't help in sheet A, B, D, and E. Nor can one just tell it to loop thru the max value of all of the sheeets, I looked, i think its in B if I remember correctly, nor can one get their arbitrarily by going to 900,000. So I found it in the 'Starcounter' exercise 'LastRow' was what I was looking for. 

Then there was the issue of finding open and close values, for the Yearly change. I knew how to get to Total Stock Volume in column L easily enough, but for some reason I thought that I remembered order was important and that it had to move Left to Right from 'ticker' to 'annual change' to 'percentage change' to stock volume. worked thru that easily enough, at least the last row value, but it was the check if the  ticker is the same <> line that I didn't realize needed to be included in the LastRow portion. 

After all of that conditional highlighting was easy with a less than, greater than values. And I learned that it wasa either or, I think the first few time I had both <= and => but there only needs to be one that determines the 'either' and the rest is just the 'or', was making it too complicated. The credit card challenge helped with the range and summary row values.

I still cant get the highest and lowest percentage change and then the greatests total stock volume to work but its a bonus and Ill go back to this again when refreshed.
