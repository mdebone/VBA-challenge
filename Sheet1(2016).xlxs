VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub UniqueCounter()

    'Set an initial varaiable for holding the Stock Ticker
    
    Dim Stock_Ticker As String
    
    'Set an initial variable for holding the Total per Stock Ticker
    Dim Stock_Total As Long
    Stock_Total = 0
    
    'Keep track of the location for each Stock Ticker in the summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    'Loop through all Stock Names
    For i = 2 To 797711
    
        'Check if we are still within the same Stock Name, if not..
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set the Stock Name
            Stock_Ticker = Cells(i, 1).Value
            
            'Add to Stock Total
            Stock_Total = Stock_Total + Cells(i, 7).Value
            
            'Print the Stock Ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = Stock_Ticker
            
            'Print the Stock Total Amount to the Summary Table
            Range("L" & Summary_Table_Row).Value = Stock_Total
            
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset the Stock Total
            Stock_Total = 0
            
        'If the cell immediately following a row is the same stock...
        Else
            
            'Add to the Stock Total
            Stock_Total = Stock_Total + Cells(i, 3).Value
        
        End If
        
    Next i
    
End Sub

Sub YearlyChange()

    Dim Opening_Price As Integer
    Dim Closing_Price As Integer
    
End Sub


