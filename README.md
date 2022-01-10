# Commitments-of-Traders-Analysis
Automated Excel Workbook for the analysis of COT data from the CFTC

Due to sandboxing restrictions this project is currently only available on the Windows Operating System.


Necessary database files.

https://www.dropbox.com/s/p1ac5yd7im9idq5/Disaggregated.accdb?dl=0

https://www.dropbox.com/s/dxzx7hz2zyquo5k/Legacy.accdb?dl=0

https://www.dropbox.com/s/36tzm4msno323i4/TFF.accdb?dl=0


Stochastic Calculations:
  There are two primary indexes used in the workbooks. One is a stochastic calculation that uses the net position of a given group while the other (WillCo) uses Net positions/Open Interest to calculate its values.
  Please note that since I have opted to not include the current week’s value within the range of values used to create the index, values less than 0 and greater than 100 will appear. Values like these signify that the current value is the most bearish or bullish value compared to the past N weeks’ worth of data.
  The Movement Index comes from Stephen Breise’s book on COT data. It is calculated by showing the difference between the current weeks 3 Year index value and what it was 6 weeks ago.
    Index naming conventions:
      Example [Group Classification] [Number][Month or Year abbreviated as a single letter(M or Y)][The letter I signifying index]

Charts Updated January 12, 2021
●	Only certain charts made by users will auto-update when placed on the 3 Charts worksheets.
●	 Please be aware that I do not have a complete list of charts that currently work with the imbedded scripts. I will expand on this functionality later.

Notes on cell Coloring
Cell coloring is not intended to be used as a buy or sell indicator. Make sure you know what it represents in EACH column as certain colors may be reused and represent different things. Sometimes the actual color doesn’t matter as much as the shade does. For example, 3 Std Deviations above or below average usually has a darker background with white letters. 
If you don’t like the formatting, then there is a macro to remove all conditional formatting from the current sheet and another to mirror that formatting onto other worksheets.
