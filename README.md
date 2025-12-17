
 # Creating macro with VBA code to automate excel formula across multiple worksheets in an excel workbook 

 In this project, I’m using data from an organization, and the data has been imported to an excel workbook. 
 
The organization has four buildings/divisions; East, West, North & South. The worksheets list the expenditure of each division for three months (Jan, Feb & March) spent on various purposes such as Telephone, Overhead, Technical support, Telemarketing, Advertising etc.

Problem statement:
Prepare a Macro with VBA code summing the Total Expense column. 

Solution VBA code: 
--------
Public Sub AutomateTotalSum()
    'Define a variable to get the address of last cell of the selection
    Dim lastCell As String
    'Define variable to each worksheet
    Dim ws As Worksheet
    
    
    'Loop over the worksheets
    For Each ws In Worksheets
        Worksheets(ws.Name).Select
        'Select the first cell of the reference column, here Total Expense
        Range("F2").Select
        
        'Select the last cell of the F column
        Selection.End(xlDown).Select
        
        'Assign the address of the last cell (relative refence, not absolute reference)
        lastCell = ActiveCell.Address(False, False)
        
        'Bring the cursor down to one row, same column
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "=sum(F2:" & lastCell & " )"
    Next ws
    
End Sub

-----

