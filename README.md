# Commitments-of-Traders-Analysis
Automated Excel Workbook for the analysis of COT data from the CFTC

COMMITMENTS OF TRADERS README FILE
Please read the descriptions of all macros if available in the developer tab. If the Developer tab is unavailable to you then, search the internet for how to enable it.

Note as of January 12,2021
There are plenty of macros within the workbook. Please familiarize yourself with them. 


Do not to do the following actions   {updated October 28, 2019}
•	Re-arrange or delete columns that came with the workbook by default. Otherwise, data won’t be placed in its proper column.
•	Sort the worksheets so that the latest data is at the top.
•	Insert new columns between the columns that came with the workbook by default.
•	Delete Queries, Query tables, non-data Tables and Worksheets.
•	Edit Queries.


Updating Weekly Data updated February 2nd, 2021
If you are on a Windows operating system, then data retrieval should be fully automatic if the corresponding checkbox on the “Data Portal” Worksheet is activated.
If you use a MAC then only the most recent week’s data will be available. If missing more than 1 week, then you will be prompted to download and extract their content and then provide their file paths on a separate worksheet.
-Processing times will vary based on your computer specifications.

Formulas Updated January 21st 2021
-You can add your own columns to the worksheets as long as the added column is to the right of all columns that originally came with the workbook. Please note that there is a macro to mirror column formulas onto all other tables.

Stochastic Calculations:
  There are two primary indexes used in the workbooks. One is a stochastic calculation that uses the net position of a given group while the other (WillCo) uses Net positions/Open Interest to calculate its values.
  Please note that since I have opted to not include the current week’s value within the range of values used to create the index, values less than 0 and greater than 100 will appear. Values like these signify that the current value is the most bearish or bullish value compared to the past N weeks’ worth of data.
  The Movement Index comes from Stephen Breise’s book on COT data. It is calculated by showing the difference between the current weeks 3 Year index value and what it was 6 weeks ago.
    Index naming conventions:
      Example [Group Classification] [Number][Month or Year abbreviated as a single letter(M or Y)][The letter I signifying index]

Charts Updated January 12, 2021
●	Only certain charts made by users will auto-update when placed on the Charts worksheet.
●	 Please be aware that I do not have a complete list of charts that currently work with the imbedded scripts. I will expand on this functionality later.
●	Filtering the data table associated with the contract you wish to view will also filter the charts if it is selected as an option.

Notes on cell Coloring
Cell coloring is not intended to be used as a buy or sell indicator. Make sure you know what it represents in EACH column as certain colors may be reused and represent different things. Sometimes the actual color doesn’t matter as much as the shade does. For example, 3 Std Deviations above or below average usually has a darker background with white letters. 
If you don’t like the formatting, then there is a macro to remove all conditional formatting from the current sheet and another to mirror that formatting onto other worksheets.
