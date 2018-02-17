{\rtf1\ansi\ansicpg1252\cocoartf1561\cocoasubrtf200
{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub volume()\
' Easy\
\
' Create a script that will loop through each year of stock data and grab the total\
' amount of volume each stock had over the year.\
\
' You will also need to display the ticker symbol to coincide with the total volume.\
\
\
    ' creating headers for all my new columns\
    Cells(1, 9).Value = "Ticker"\
    Cells(1, 10).Value = "Stock Volume Total"\
   \
   \
   ' Set an initial variable for holding the ticker\
     Dim Ticker As String\
   \
     ' Set an initial variable for holding the Stock Volume Total\
     Dim Stock_Volume_Total As Double\
     Stock_Volume_Total = 0\
   \
     ' Keep track of the location for each ticker in the summary table\
     Dim Summary_Table_Row As Integer\
     Summary_Table_Row = 2\
     \
     ' Loop through all tickers\
     For i = 2 To 800000\
   \
       ' Check if we are still within the same ticker, if it is not...\
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then\
   \
         ' Set the ticker\
         Ticker = Cells(i, 1).Value\
   \
         ' Add to the stock volume total\
         Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value\
   \
         ' Print the Ticker in the Summary Table\
         Range("I" & Summary_Table_Row).Value = Ticker\
   \
         ' Print the Stock Volume Total to the Summary Table\
         Range("J" & Summary_Table_Row).Value = Stock_Volume_Total\
   \
         ' Add one to the summary table row\
         Summary_Table_Row = Summary_Table_Row + 1\
         \
         ' Reset the stock volume total\
         Stock_Volume_Total = 0\
   \
       ' If the cell immediately following a row is the same ticker...\
       Else\
   \
         ' Add to the stock volume total\
         Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value\
   \
       End If\
   \
     Next i\
     \
   \
   End Sub\
\
}