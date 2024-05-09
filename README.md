This challenge required using VBA scripting to analyze quarterly stock results for a large number of stocks over multiple years.  I used VBA script to calculate quarterly price changes for each individual stock, as well as percentage changes and volumes traded.  The script did this for each of 3 worksheets, each of which held data for one year.


I did not use large chunks of code from any sources but I did google many things to solve problems with trying to make the code work.
One place I used was    excelchamps.com/vba/loop-sheets   which helped me with how to set up looping through the worksheets.  I struggled with the syntax a little and then realized that I needed to include Sheets(J). in front of each Cells or Range assignment.  I missed one which caused problems for a while.  I was tempted to use Find and Replace to do this, and in hindsight, I think I should have.
I used a different site to determine how to count the rows in the worksheet in order to loop through the ticker symbols on a given worksheet.  I lost my notes on which site it was, but it may have been    https://excelchamps.com/vba/rows-count/

I first coded the loop to identify the unique ticker symbols, checking each symbol against the next one in the column.  This meant I had to code to include the final stock's year end price and volume separately.  

I stored the ticker symbols in an array or list.  I wanted to do the same with the opening and closing prices, but had a problem with the syntax or something and I think excel wasn't recoginzing these as numerical values and I needed them for calculations.  So I instead used the subroutine to temporarily store the values in cells alongside the ticker symbols and had the subroutine erase them when finishing.  

I did not work with anyone, but I did get some help debugging from a TA, Danila.  He helped me to recognize that putting in MsgBox commands could help to identify where the code was being executed.  This helped me to find the missing      Sheets(J).    that was preventing the code that worked on the first sheet perfectly from working on the remaining sheets.
He also suggested setting up a button with a subroutine to erase the fields that were populated via the macro so that this could be done easily during debugging.
