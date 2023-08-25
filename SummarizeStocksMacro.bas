Attribute VB_Name = "Module1"
'*******************************************************************************************
 '
 '  File Name:  main.c
 '
 '  File Description:
 '      The file contains the macro, SummarizeStocksMacro, which takes original stock data
 '      in the spreadsheet, reformats the spreadsheet, and generates summary data.
 '
 '
 '  Date                         Description                                                     Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023             Initial Development                                         N. James George
 '
'*******************************************************************************************/

' These are the global enumerations that identify the rows and columns in the original
' and summary data.

Enum RowGlobalEnumeration
    
    ENUM_K_TITLE = 1
    
    ENUM_K_FIRST_DATA = 2
    
    ENUM_K_PERCENT_DECREASE = 3
    
    ENUM_K_GREATEST_TOTAL_VOLUME = 4

End Enum


Enum ColumnGlobalEnumeration
    
    ENUM_K_ORIGINAL_TICKER = 1
    
    ENUM_K_ORIGINAL_DATE = 2
    
    ENUM_K_ORIGINAL_OPEN = 3
    
    ENUM_K_ORIGINAL_HIGH = 4
    
    ENUM_K_ORIGINAL_LOW = 5
    
    ENUM_K_ORIGINAL_CLOSE = 6
    
    ENUM_K_ORIGINAL_VOL = 7
    
    ENUM_K_BLANK_1 = 8

    ENUM_K_SUMMARY_TICKER = 9
    
    ENUM_K_SUMMARY_YEARLY_CHANGE = 10
    
    ENUM_K_SUMMARY_PERCENT_CHANGE = 11
    
    ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME = 12
    
    ENUM_K_BLANK_2 = 13
    
    ENUM_K_BLANK_3 = 14
    
    ENUM_K_CHANGE_ROW_TITLES = 15
    
    ENUM_K_CHANGE_TICKERS = 16
    
    ENUM_K_CHANGE_VALUES = 17

End Enum


' These global constants specify substring lengths in the original data's the date strings:
' the date string format is YYYYMMDD.

Global Const _
    GLOBAL_CONSTANT_YEAR_LENGTH _
        As Integer _
= 4

Global Const _
    GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH _
        As Integer _
= 2


' This global variable holds the value of the number of rows in the original data.

Global _
    lastDataRowGlobalLongVariable As Long

'*******************************************************************************************
 '
 '  Macro Name:  SummarizeStocksMacro
 '
 '  Macro Description:
 '      This macro loops through all the stocks for one year and outputs
 '      the following information:
 '
 '          1) The ticker symbol
 '
 '          2) Yearly change from the opening price at the beginning of a given year
 '              to the closing price at the end of that year.
 '
 '          3) The percentage change from the opening price at the beginning of
 '              a given year to the closing price at the end of that year.
 '
 '          4) The total stock volume of the stock.
 '
 '  Macro Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                      n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Sub _
    SummarizeStocksMacro()

    ' This line of code assigns the last row index to the appropriate global variable.
    
    lastDataRowGlobalLongVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (Rows.Count, _
                         ColumnGlobalEnumeration.ENUM_K_ORIGINAL_TICKER) _
                            .End _
                                (xlUp) _
                                    .Row


    ' These subroutines format the active spreadsheet.

    FormatOriginalDataPrivateSubRoutine

    FormatSummaryDataPrivateSubRoutine
    
    FormatTitlesPrivateSubRoutine
    
    FormatEntireWorkSheetPrivateSubRoutine
        
        
    ' This subroutine summarizes the stock information in the original data
    ' and writes it to the new summary table.
    
    CreateSummaryTablePrivateSubRoutine
    
    
    ' This subroutine creates a third table for the tickers with the greatest changes
    ' in percentage and greatest total stock volume.
    
    CreateChangeTablePrivateSubRoutine

End Sub ' This statement ends the macro,
' SummarizeStocksMacro.

'*******************************************************************************************
 '
 '  Subroutine Name:  FormatEntireWorkSheetPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the font type and font size for the entire worksheet.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                      n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
        FormatEntireWorkSheetPrivateSubRoutine()

    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                .Font.Name _
    = "Garamond"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                .Font.Size _
    = 14

End Sub ' This statement ends the private subroutine,
' FormatEntireWorkSheetPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  FormatOriginalDataPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the formats of the various columns in the original data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                      n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
        FormatOriginalDataPrivateSubRoutine()

    ' If the first value in the date column is a string, the subroutine converts all its values
    ' to a Date type.

    If VarType _
            (Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_ORIGINAL_DATE) _
                            .Value) _
       = vbString Then
        
        ChangeStringToDateInDateColumnPrivateSubRoutine
    
    End If
    

    ' These lines of code change the column formats.
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_ORIGINAL_TICKER) _
                    .NumberFormat _
    = "General"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_ORIGINAL_DATE) _
                    .NumberFormat _
    = "mm/dd/yyyy"
            
    For indexLocalCounterVariable _
            = ColumnGlobalEnumeration.ENUM_K_ORIGINAL_OPEN _
                    To ColumnGlobalEnumeration.ENUM_K_ORIGINAL_CLOSE
    
        Worksheets _
            (ActiveSheet.Name) _
                .Columns _
                    (indexLocalCounterVariable) _
                        .NumberFormat _
        = "$#,##0.00"
    
    Next indexLocalCounterVariable ' This statement ends the first repetition loop.
    
            
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_ORIGINAL_VOL) _
                    .NumberFormat _
    = "#,##0"


    ' These lines of code change the column widths.
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_ORIGINAL_TICKER) _
                    .ColumnWidth _
    = 10
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_ORIGINAL_DATE) _
                    .ColumnWidth _
    = 14
    
    
    For indexLocalCounterVariable _
            = ColumnGlobalEnumeration.ENUM_K_ORIGINAL_OPEN _
                    To ColumnGlobalEnumeration.ENUM_K_ORIGINAL_VOL

        Worksheets _
            (ActiveSheet.Name) _
                .Columns _
                    (indexLocalCounterVariable) _
                        .ColumnWidth _
        = 12
            
    Next indexLocalCounterVariable ' This statement ends the second repetition loop.
    
End Sub ' This stastement ends the private subroutine,
' FormatOriginalDataPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  ChangeStringToDateInDateColumnPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine changes the text strings in the original data date column
 '        to a Date type.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a      n/a                      n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
    ChangeStringToDateInDateColumnPrivateSubRoutine()

    '  These lines of code declare variables for the current date.
    
    Dim _
        dateLocalDateVariable _
            As Date
    
    ' These lines of code declare variables for the current year, month, and day.
    
    Dim _
        yearLocalIntegerVariable _
            As Integer
    
    Dim _
        monthLocalIntegerVariable _
            As Integer
    
    Dim _
        dayLocalIntegerVariable _
            As Integer
    
    ' These lines of code declare variables for the start indexes in the date string.
    
    Dim _
        yearStartIndexIntegerVariable _
            As Integer
    
    Dim _
        monthStartIndexIntegerVariable _
            As Integer
    
    Dim _
        dayStartIndexVariable _
            As Integer
    
    
    ' These lines of code initialize variables for the start indices.
    
    yearStartIndexIntegerVariable _
        = 1
        
    monthStartIndexIntegerVariable _
        = yearStartIndexIntegerVariable _
           + GLOBAL_CONSTANT_YEAR_LENGTH
        
    dayStartIndexVariable _
        = monthStartIndexIntegerVariable _
            + GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH
    
    
    ' These lines of code loop through all the values in the specified column
    ' and converts them to a Date type.
    
    For rowLocalCounterVariable _
            = RowGlobalEnumeration.ENUM_K_FIRST_DATA _
                    To lastDataRowGlobalLongVariable
    
        ' These lines of code parse out the date from the string, YYYYMMDD,
        ' in the current cell and converts it to a Date type.
    
        yearLocalIntegerVariable _
            = Mid _
                    (Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowLocalCounterVariable, _
                                    ColumnGlobalEnumeration.ENUM_K_ORIGINAL_DATE) _
                                        .Value, _
                                 yearStartIndexIntegerVariable, _
                                 GLOBAL_CONSTANT_YEAR_LENGTH)
        
        monthLocalIntegerVariable _
            = Mid _
                    (Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowLocalCounterVariable, _
                                    ColumnGlobalEnumeration.ENUM_K_ORIGINAL_DATE) _
                                        .Value, _
                                 monthStartIndexIntegerVariable, _
                                 GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH)
        
        dayLocalIntegerVariable _
            = Mid _
                    (Worksheets(ActiveSheet.Name) _
                        .Cells _
                            (rowLocalCounterVariable, _
                                ColumnGlobalEnumeration.ENUM_K_ORIGINAL_DATE) _
                                    .Value, _
                             dayStartIndexVariable, _
                             GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH)
        
        ' This line of code takes the values for year, month, and day, converts them
        ' to a Date type, then  assigns them to the appropriate variable
        
        dateLocalDateVariable _
            = DateSerial _
                    (yearLocalIntegerVariable, _
                    monthLocalIntegerVariable, _
                    dayLocalIntegerVariable)
    
    
        ' This line of code assigns the new date value to the current cell.
        
        Worksheets _
            (ActiveSheet.Name) _
                .Cells _
                    (rowLocalCounterVariable, _
                        ColumnGlobalEnumeration.ENUM_K_ORIGINAL_DATE).Value _
        = dateLocalDateVariable
    
    Next rowLocalCounterVariable ' This statement ends the repetition loop.

End Sub ' This statement ends the private subroutine,
' ChangeStringToDateInDateColumnPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  SetUpTitlesForSummaryDataPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets up the titles for the summary data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                     n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
    SetUpTitlesForSummaryDataPrivateSubRoutine()

    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_TITLE, _
                    ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                        .Value _
    = "Ticker"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_TITLE, _
                    ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
                        .Value _
    = "Yearly Change"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_TITLE, _
                    ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                        .Value _
    = "Percent Change"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_TITLE, _
                    ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                        .Value _
    = "Total Stock Volume"
    
End Sub  ' This statement ends the public subroutine,
' SetUpTitlesForSummaryDataPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  FormatSummaryDataPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the format for the columns holding summarized data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                     n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
    FormatSummaryDataPrivateSubRoutine()

    ' These lines of code set the formats for the various columns.
     
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                    .NumberFormat _
    = "General"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
                    .NumberFormat _
    = "#,##0.00"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                    .NumberFormat _
    = "0.00%"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                    .NumberFormat _
    = "#,##0"
    
        
    ' These lines of code set the column widths for the various columns.
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                    .ColumnWidth _
    = 10
    
    
    For indexLocalCounterVariable _
            = ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE _
                    To ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE

        Worksheets _
            (ActiveSheet.Name) _
                .Columns _
                    (indexLocalCounterVariable) _
                        .ColumnWidth _
        = 16
            
    Next indexLocalCounterVariable ' This statement ends the repetition loop.
    
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                    .ColumnWidth _
    = 25

End Sub ' This statement ends the private subroutine,
' FormatSummaryDataPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  FormatTitlesPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the format for the row containing titles for both
 '       the original data and the summarized data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                     n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
    FormatTitlesPrivateSubRoutine()

    Worksheets _
        (ActiveSheet.Name) _
            .Rows _
                (RowGlobalEnumeration.ENUM_K_TITLE) _
                    .NumberFormat _
    = "General"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Rows _
                (RowGlobalEnumeration.ENUM_K_TITLE) _
                    .Font _
                        .Bold _
    = True
    
    Worksheets _
        (ActiveSheet.Name) _
            .Rows _
                (RowGlobalEnumeration.ENUM_K_TITLE) _
                    .HorizontalAlignment _
    = xlCenter

End Sub ' This statement ends the private subroutine,
' FormatTitlesPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  CreateSummaryTablePrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine creates the summary table by analyzing the original data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                     n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
    CreateSummaryTablePrivateSubRoutine()
    
    ' This lines of code declare variables for the original data's first row index
    ' in the repetition loop.
    
    Dim _
        firstRowLocalLongVariable _
            As Long
    

    ' These lines of code declare variables for the original data record.
    ' The program uses these values to calculate the summary table record.
    
    Dim _
        currentTickerNameLocalStringVariable _
            As String
    
    Dim _
        openingPriceLocalCurrencyVariable _
            As Currency
    
    Dim _
        closingPriceLocalCurrencyVariable _
            As Currency
    
    Dim _
        totalStockVolumeLocalVariantVariable _
            As Variant
    
    
    ' This line of code declares the variable for the row index
    ' in the summary table.
    
    Dim _
        summaryTableRowLocalLongVariable _
            As Long
    
    
    ' This subroutine places the titles in the appropriate cells.
    
    SetUpTitlesForSummaryDataPrivateSubRoutine
    
    
    ' These lines of code initialize variables with information from the first row
    ' of the original data.
    
    currentTickerNameLocalStringVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_ORIGINAL_TICKER) _
                            .Value
    
    openingPriceLocalCurrencyVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_ORIGINAL_OPEN) _
                            .Value
                
    totalStockVolumeLocalVariantVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_ORIGINAL_VOL) _
                            .Value
                
    
    ' These lines of code set the initial row indices for the original data
    ' and summary tables.
    
    firstRowLocalLongVariable _
        = RowGlobalEnumeration.ENUM_K_FIRST_DATA + 1
    
    summaryTableRowLocalLongVariable _
        = RowGlobalEnumeration.ENUM_K_FIRST_DATA
 
 
    ' This repetition loop runs through all the rows of the original data
    ' and generates the summary table: the loop starts with the second
    ' row of original data.

    For originalRowCounterVariable _
                = firstRowLocalLongVariable _
                        To lastDataRowGlobalLongVariable
    
        If Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (originalRowCounterVariable, _
                            ColumnGlobalEnumeration.ENUM_K_ORIGINAL_TICKER) _
                                .Value _
            = currentTickerNameLocalStringVariable Then
        
        
            ' If the ticker name is the same, this line of code adds the current stock volume
            ' to the total.
                                    
            totalStockVolumeLocalVariantVariable _
                = totalStockVolumeLocalVariantVariable _
                        + Worksheets _
                                (ActiveSheet.Name) _
                                    .Cells _
                                        (originalRowCounterVariable, _
                                            ColumnGlobalEnumeration.ENUM_K_ORIGINAL_VOL) _
                                                .Value
                            
                            
            ' If the loop has reached the last row the program creates a summary table record.
        
            If originalRowCounterVariable _
                    = lastDataRowGlobalLongVariable Then
                                           
               CreateSummaryTableRowPrivateSubRoutine _
                        currentTickerNameLocalStringVariable, _
                        openingPriceLocalCurrencyVariable, _
                        totalStockVolumeLocalVariantVariable, _
                        summaryTableRowLocalLongVariable, _
                        originalRowCounterVariable, _
                        True
                    
            End If
            
        Else
        
            ' This if statement executes if the loop has not reached the end of the data.
        
            If originalRowCounterVariable _
                    <> lastDataRowGlobalLongVariable Then
            
                ' If the current ticker does not match the previous ticker,
                ' the program creates a record.
            
                CreateSummaryTableRowPrivateSubRoutine _
                        currentTickerNameLocalStringVariable, _
                        openingPriceLocalCurrencyVariable, _
                        totalStockVolumeLocalVariantVariable, _
                        summaryTableRowLocalLongVariable, _
                        originalRowCounterVariable, _
                        False
                    
                
                ' These lines of code assign new values to the original data variables.
                
                currentTickerNameLocalStringVariable _
                    = Worksheets _
                            (ActiveSheet.Name) _
                                .Cells _
                                    (originalRowCounterVariable, _
                                        ColumnGlobalEnumeration.ENUM_K_ORIGINAL_TICKER) _
                                            .Value
                
                openingPriceLocalCurrencyVariable _
                    = Worksheets _
                            (ActiveSheet.Name) _
                                .Cells _
                                    (originalRowCounterVariable, _
                                     ColumnGlobalEnumeration.ENUM_K_ORIGINAL_OPEN) _
                                        .Value
                
                totalStockVolumeLocalVariantVariable _
                    = Worksheets _
                            (ActiveSheet.Name) _
                                .Cells _
                                    (originalRowCounterVariable, _
                                     ColumnGlobalEnumeration.ENUM_K_ORIGINAL_VOL) _
                                        .Value
            
                
               ' This line of code increases the summary table row index
               ' by one for the next record.
            
                summaryTableRowLocalLongVariable _
                    = summaryTableRowLocalLongVariable + 1
                        
            Else
            
                ' These lines of code initialize variables with information
                ' from the last row of the original data.
    
                currentTickerNameLocalStringVariable _
                    = Worksheets _
                            (ActiveSheet.Name) _
                                .Cells _
                                    (originalRowCounterVariable, _
                                     ColumnGlobalEnumeration.ENUM_K_ORIGINAL_TICKER) _
                                        .Value
                    
                openingPriceLocalCurrencyVariable _
                    = Worksheets _
                            (ActiveSheet.Name) _
                                .Cells _
                                    (originalRowCounterVariable, _
                                     ColumnGlobalEnumeration.ENUM_K_ORIGINAL_OPEN) _
                                        .Value
                            
                totalStockVolumeLocalVariantVariable _
                    = totalStockVolumeLocalVariantVariable _
                        + Worksheets _
                                (ActiveSheet.Name) _
                                    .Cells _
                                        (originalRowCounterVariable, _
                                         ColumnGlobalEnumeration.ENUM_K_ORIGINAL_VOL) _
                                            .Value
                                
                                
                ' The program then creates a record with this information.
                                
                CreateSummaryTableRowPrivateSubRoutine _
                        currentTickerNameLocalStringVariable, _
                        openingPriceLocalCurrencyVariable, _
                        totalStockVolumeLocalVariantVariable, _
                        summaryTableRowLocalLongVariable, _
                        originalRowCounterVariable, _
                        True
            
            End If
            
        End If
        
    Next originalRowCounterVariable ' This statement ends the repetition loop.
    
End Sub ' This statement ends the private subroutine,
' CreateSummaryTablePrivateSubRoutine.

'*******************************************************************************************
 '
 '  Function Name:  CalculateYearlyChangePrivateFunction
 '
 '  Function Type: Private
 '
 '  Function Description:
 '      This function calculates the yearly change between the first opening price
 '      and the last closing price of a ticker.
 '
 '  Function Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  ByVal   openingPriceParameterDoubleVariable
 '                                          This parameter holds the first opening price of a ticker.
 '  ByVal   closingPriceParameterDoubleVariable
 '                                          This parameter holds the last closing price of a ticker.
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Function _
    CalculateYearlyChangePrivateFunction _
        (ByVal _
            openingPriceParameterDoubleVariable _
                As Double, _
        ByVal _
            closingPriceParameterDoubleVariable _
                As Double) _
As Double
    
    CalculateYearlyChangePrivateFunction _
        = closingPriceParameterDoubleVariable _
            - openingPriceParameterDoubleVariable
            
End Function ' This statement ends the private function,
' CalculateYearlyChangePrivateFunction.

'*******************************************************************************************
 '
 '  Subroutine Name:  FormatYearlyChangeCellPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This function formats the newly assigned yearly change cell in the summary table
 '      based on the summary table row index.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  ByVal   rowIndexParameterIntegerVariable
 '                                          This parameter holds the row index for the current record
 '                                          in the summary table.
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '******************************************************************************************/

Private Sub _
    FormatYearlyChangeCellPrivateSubRoutine _
        (ByVal _
            rowIndexParameterIntegerVariable _
                As Integer)
    
    ' If the yearly change is zero or positive, the program changes the background color
    ' to green.

    If Worksheets _
            (ActiveSheet.Name) _
                .Cells _
                    (rowIndexParameterIntegerVariable, _
                     ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
                        .Value _
       >= 0 Then
        
        Worksheets _
            (ActiveSheet.Name) _
                .Cells _
                    (rowIndexParameterIntegerVariable, _
                     ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
                        .Interior _
                            .ColorIndex _
        = 4
        
    Else ' If the yearly change is negative, the program changes the background color to red.
    
        Worksheets _
            (ActiveSheet.Name) _
                .Cells _
                    (rowIndexParameterIntegerVariable, _
                     ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
                        .Interior _
                            .ColorIndex _
        = 3
    
    End If
    
End Sub ' This statement ends the private subroutine,
' FormatYearlyChangeCellPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Function Name:  CalculatePercentChangePrivateFunction
 '
 '  Function Type: Private
 '
 '  Function Description:
 '      This function calculates the percent change between the first opening price
 '      and the last closing price of a ticker.
 '
 '  Function Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  ByVal   openingPriceParameterDoubleVariable
 '                                          This parameter holds the first opening price of a ticker.
 '  ByVal   closingPriceParameterDoubleVariable
 '                                          This parameter holds the last closing price of a ticker.
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Function _
    CalculatePercentChangePrivateFunction( _
        ByVal _
            openingPriceParameterDoubleVariable _
                As Double, _
        ByVal _
            closingPriceParameterDoubleVariable _
                As Double) _
As Double

    CalculatePercentChangePrivateFunction _
        = (closingPriceParameterDoubleVariable - openingPriceParameterDoubleVariable) _
            / openingPriceParameterDoubleVariable

End Function ' This statement ends the private function,
' CalculatePercentChangePrivateFunction.

'*******************************************************************************************
 '
 '  Subroutine Name:  CreateSummaryTableRowPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine creates a summary table record based on the parameters.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  ByVal    tickerNameParameterStringVariable
 '                                          This parameter holds the name of the stock ticker.
 '  ByVal    openingPriceParameterCurrencyVariable
 '                                          This parameter is the first opening price for this
 '                                          stock ticker.
 '  ByVal    totalStockVolumeParameterVariantVariable
 '                                          This parameter is the total stock volume for this
 '                                          stock ticker.
 '  ByVal    summaryRowParameterLongVariable
 '                                          This parameter is the current summary table row index.
 '  ByVal    originalRowParameterLongVariable
 '                                          This parameter is the current original data row index.
 '  ByVal    lastRowFlagParameterBooleanVariable
 '                                          This parameter indicates whether the program
 '                                          has reached the last record or not.
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
    CreateSummaryTableRowPrivateSubRoutine _
        (ByVal _
            tickerNameParameterStringVariable _
                As String, _
         ByVal _
            openingPriceParameterCurrencyVariable _
                As Currency, _
         ByVal _
            totalStockVolumeParameterVariantVariable _
                As Variant, _
         ByVal _
            summaryRowParameterLongVariable _
                As Long, _
         ByVal _
            originalRowParameterLongVariable _
                As Long, _
        ByVal _
            lastRowFlagParameterBooleanVariable _
                As Boolean)

    ' This line of code declares a variable for the closing price which is different
    ' based on whether the program has reached the last row or not in the
    ' original data
    
    Dim _
        closingPriceLocalCurrencyVariable _
            As Currency


    ' If the program has not reached the last row, the closing price comes
    ' from the previous row in the original data; otherwise, the closing
    ' price comes from the current row.
                  
    If lastRowFlagParameterBooleanVariable _
       = False Then
            
        closingPriceLocalCurrencyVariable _
            = Worksheets _
                    (ActiveSheet.Name) _
                        .Cells _
                            (originalRowParameterLongVariable - 1, _
                             ColumnGlobalEnumeration.ENUM_K_ORIGINAL_CLOSE) _
                                .Value
            
    Else
            
        closingPriceLocalCurrencyVariable _
            = Worksheets _
                    (ActiveSheet.Name) _
                        .Cells _
                            (originalRowParameterLongVariable, _
                             ColumnGlobalEnumeration.ENUM_K_ORIGINAL_CLOSE) _
                                .Value
            
    End If
            
            
    ' These lines of code create a record in the summary table.
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (summaryRowParameterLongVariable, _
                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                    .Value _
    = tickerNameParameterStringVariable
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (summaryRowParameterLongVariable, _
                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
                    .Value _
    = CalculateYearlyChangePrivateFunction _
            (CDbl _
                (openingPriceParameterCurrencyVariable), _
            CDbl _
                (closingPriceLocalCurrencyVariable))
                            
    FormatYearlyChangeCellPrivateSubRoutine _
        summaryRowParameterLongVariable
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (summaryRowParameterLongVariable, _
                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                    .Value _
    = CalculatePercentChangePrivateFunction _
            (CDbl _
                (openingPriceParameterCurrencyVariable), _
             CDbl _
                (closingPriceLocalCurrencyVariable))
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (summaryRowParameterLongVariable, _
                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                    .Value _
    = totalStockVolumeParameterVariantVariable

End Sub ' This statement ends the private subroutine,
' CreateSummaryTableRowPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  CreateChangeTablePrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine creates a table that lists the tickers with the greatest percent change
 '      and the ticker with the greatest total stock volume.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                      n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023            Initial Development                                         N. James George
 '******************************************************************************************/

Public Sub _
    CreateChangeTablePrivateSubRoutine()

    FormatChangeTablePrivateSubRoutine
    
    SetupChangeTableTitlesPrivateSubRoutine
    
    CalculateAndWriteChangeTableDataPrivateSubRoutine

End Sub ' This statement ends the private subroutine,
' CreateChangeTablePrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  FormatChangeTableTitlesPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine formats the row and column titles in the change table.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                      n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
    FormatChangeTablePrivateSubRoutine()

    ' These lines of code format the columns and cells of the change table.

    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_CHANGE_ROW_TITLES) _
                    .NumberFormat _
    = "General"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_CHANGE_TICKERS) _
                    .NumberFormat _
    = "General"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES) _
                    .NumberFormat _
    = "0.00%"
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_PERCENT_DECREASE, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES) _
                    .NumberFormat _
    = "0.00%"
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES) _
                    .NumberFormat _
    = "#,##0"
            
           
    ' These lines of code set the column widths for the change table.
           
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_CHANGE_ROW_TITLES) _
                    .ColumnWidth _
    = 25
            
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_CHANGE_TICKERS) _
                    .ColumnWidth _
    = 10
            
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES) _
                    .ColumnWidth _
    = 25
            
            
    ' This line of code sets the font style for the row titles to bold.
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_CHANGE_ROW_TITLES) _
                    .Font _
                        .Bold _
    = True

End Sub ' This statement ends the private subroutine,
' FormatChangeTableTitlesPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  SetupChangeTableTitlesPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine writes the column and row titles to the appropriate cells
 '      in the change table.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                      n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
    SetupChangeTableTitlesPrivateSubRoutine()

    ' These lines of code set the column titles in the change table.
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_TITLE, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_TICKERS) _
                    .Value _
    = "Ticker"
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_TITLE, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES) _
                    .Value _
    = "Value"
            
            
    ' These lines of code set the row titles in the change table,
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_ROW_TITLES) _
                    .Value _
    = "Greatest % Increase"
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_PERCENT_DECREASE, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_ROW_TITLES).Value _
    = "Greatest % Decrease"
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_ROW_TITLES) _
                    .Value _
    = "Greatest Total Volume"

End Sub ' This statement ends the private subroutine,
' SetupChangeTableTitlesPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  CalculateChangeTableDataPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine calculates the values for the change table based on data
 '      in the summary table and writes the results to the change table.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name                   Description
 '  -----   -------------   ----------------------------------------------
 '  n/a       n/a                      n/a
 '
 '
 '  Date                        Description                                                      Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023            Initial Development                                         N. James George
 '
 '******************************************************************************************/

Private Sub _
    CalculateAndWriteChangeTableDataPrivateSubRoutine()

    ' These lines of code declare variables for the various tickers.
    
    Dim _
        increaseTickerLocalStringVariable _
            As String
    
    Dim _
        decreaseTickerLocalStringVariable _
            As String
    
    Dim _
        volumeTickerLocalStringVariable _
            As String
    
    
    ' These lines of code declare variables for the associated values.
    
    Dim _
        increasePercentageLocalDoubleVariable _
            As Double
    
    Dim _
        decreasePercentageLocalDoubleVariable _
            As Double
    
    Dim _
        volumeLocalVariantVariable _
            As Variant
    
    
    ' These lines of code declare variables for the first and last index
    ' in the repetition loop.
    
    Dim _
        firstRowLocalLongVariable _
            As Long
    
    Dim _
        lastRowLocalLongVariable _
            As Long
    
    
    ' These lines of code initialize the variables with data from the first
    ' record in the summary table.
    
    increaseTickerLocalStringVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                            .Value
         
    decreaseTickerLocalStringVariable _
        = increaseTickerLocalStringVariable
        
    volumeTickerLocalStringVariable _
        = increaseTickerLocalStringVariable
        
       
    increasePercentageLocalDoubleVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                            .Value
             
    decreaseTickerLocalStringVariable _
        = increasePercentageLocalDoubleVariable
        
    volumeLocalVariantVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                            .Value
    
    
    ' These lines of code initialize the first and last index of the repetition loop.
        
    firstRowLocalLongVariable _
        = RowGlobalEnumeration.ENUM_K_FIRST_DATA + 1
        
    lastRowLocalLongVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (Rows.Count, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                                .End _
                                    (xlUp) _
                                        .Row
        
        
    ' This repetition loop starts at the second record of the summary table and,
    ' through comparisons, finds the tickers with the greatest increase, greatest
    ' decrease, and greatest total stock volume.
         
    For rowIndexLocalCounterVariable _
            = firstRowLocalLongVariable _
                    To lastRowLocalLongVariable
    
        ' If a record has a larger change in percentage than the previous holder,
        ' set it as the new leader in percentage increase.
    
        If Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (rowIndexLocalCounterVariable, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                            .Value _
           > increasePercentageLocalDoubleVariable Then
        
            increaseTickerLocalStringVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexLocalCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                                    .Value
                        
            increasePercentageLocalDoubleVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexLocalCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                                    .Value
        
        End If
        
        
        ' If a record has a smaller change in percentage than the previous holder,
        ' set it as the new leader in percentage decrease.
        
        If Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (rowIndexLocalCounterVariable, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                            .Value _
           < decreasePercentageLocalDoubleVariable Then
        
            decreaseTickerLocalStringVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexLocalCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                                    .Value
                        
            decreasePercentageLocalDoubleVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                    (rowIndexLocalCounterVariable, _
                                     ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                                        .Value
        
        End If
        
        
        ' If a record has a larger total stock volume than the previous holder,
        ' set it as the new leader in total stock volume.
        
        If Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (rowIndexLocalCounterVariable, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                            .Value _
           > volumeLocalVariantVariable Then
        
            volumeTickerLocalStringVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexLocalCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                                    .Value
                        
            volumeLocalVariantVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexLocalCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                                    .Value
        
        End If
    
    Next rowIndexLocalCounterVariable ' This statement ends the repetition loop.
             
    
    ' These lines of code write the results to the change table.
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_TICKERS) _
                    .Value _
    = increaseTickerLocalStringVariable
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES) _
                    .Value _
    = increasePercentageLocalDoubleVariable
            
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_PERCENT_DECREASE, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_TICKERS) _
                    .Value _
    = decreaseTickerLocalStringVariable
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_PERCENT_DECREASE, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES).Value _
    = decreasePercentageLocalDoubleVariable

            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_TICKERS).Value _
    = volumeTickerLocalStringVariable
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES) _
                    .Value _
    = volumeLocalVariantVariable
    
End Sub ' This statement ends the private subroutine,
' CalculateChangeTableDataPrivateSubRoutine.
