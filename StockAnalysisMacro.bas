Attribute VB_Name = "Module1"
'*******************************************************************************************
 '
 '  File Name:  main.c
 '
 '  File Description:
 '      The file contains the macro, StockAnalysisMacro, which formats the active
 '      worksheet then generates summary tables from raw stock data. Here is a list
 '      of the support subroutines and functions:
 '
 '      FormatStockDataPrivateSubRoutine
 '      FormatSummaryDataPrivateSubRoutine
 '      FormatTitlesPrivateSubRoutine
 '      FormatWorkSheetPrivateSubRoutine
 '      CreateSummaryTablePrivateSubRoutine
 '      CreateChangeTablePrivateSubRoutine
 '      ConvertDataAndSummaryRangesToTablesPrivateSubRoutine
 '      ChangeStringToDateInDateColumnPrivateSubRoutine
 '      SetUpTitlesForSummaryDataPrivateSubRoutine
 '      CreateSummaryDataRowPrivateSubRoutine
 '      FormatChangeDataTitlesPrivateSubRoutine
 '      SetupChangeDataTitlesPrivateSubRoutine
 '      CalculateAndWriteChangeDataPrivateSubRoutine
 '      ConvertRangeIntoTablePrivateSubRoutine
 '      FormatYearlyChangeCellPrivateSubRoutine
 '
 '      CalculateYearlyChangePrivateFunction
 '      CalculatePercentChangePrivateFunction
 '
 '
 '  Date               Description                             Programmer
 '  ---------------    ------------------------------------    ------------------
 '  07/19/2023         Initial Development                     N. James George
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
    
    ENUM_K_STOCK_TICKER = 1
    
    ENUM_K_STOCK_DATE = 2
    
    ENUM_K_STOCK_OPEN = 3
    
    ENUM_K_STOCK_HIGH = 4
    
    ENUM_K_STOCK_LOW = 5
    
    ENUM_K_STOCK_CLOSE = 6
    
    ENUM_K_STOCK_VOL = 7
    
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


' These global constants specify substring lengths in the original data's date strings:
' the date string format is YYYYMMDD.

Global Const _
    GLOBAL_CONSTANT_YEAR_LENGTH _
        As Integer _
            = 4

Global Const _
    GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH _
        As Integer _
            = 2


' This global variable holds the value of the number of rows in the raw stock data.

Global _
    lastDataRowGlobalLongVariable _
         As Long

'*******************************************************************************************
 '
 '  Macro Name:  StockAnalysisMacro
 '
 '  Macro Description:
 '      This macro formats the active worksheet then generates summary tables
 '      from raw stock data.
 '
 '  Macro Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Sub _
    StockAnalysisMacro()

    ' This line of code assigns the last row index to the appropriate global variable.
    
    lastDataRowGlobalLongVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (Rows.Count, _
                         ColumnGlobalEnumeration.ENUM_K_STOCK_TICKER) _
                    .End _
                        (xlUp) _
                    .Row


    ' These subroutines format the active worksheet.

    FormatWorkSheetPrivateSubRoutine _
       ActiveWorkbook.ActiveSheet

    FormatStockDataPrivateSubRoutine

    FormatSummaryDataPrivateSubRoutine
    
    FormatTitlesPrivateSubRoutine
    
        
    ' This subroutine summarizes the raw stock data and writes it to the summary table.
    
    CreateSummaryTablePrivateSubRoutine
    
    
    ' This subroutine creates a second summary table for the tickers with the greatest
    ' changes in percentage and greatest total stock volume.
    
    CreateChangeTablePrivateSubRoutine
    
    
    ' This subroutine converts the data and summary ranges to tables.
    
    ConvertDataAndSummaryRangesToTablesPrivateSubRoutine

End Sub ' This statement ends the macro,
' SummarizeStocksMacro.

'*******************************************************************************************
 '
 '  Subroutine Name:  FormatStockDataPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the formats of the stock data's various columns.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
        FormatStockDataPrivateSubRoutine()

    ' If the first value in the date column is a string, the subroutine converts all its values
    ' to a Date type.

    If VarType _
            (Worksheets _
                (ActiveSheet.Name) _
                .Cells _
                    (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                     ColumnGlobalEnumeration.ENUM_K_STOCK_DATE) _
                .Value) _
            = vbString Then
        
        ChangeStringToDateInDateColumnPrivateSubRoutine
    
    End If
    

    ' These lines of code change the column formats.
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_STOCK_TICKER) _
            .NumberFormat _
                = "General"
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_STOCK_DATE) _
            .NumberFormat _
                = "mm/dd/yyyy"
            
    For indexCounterVariable _
            = ColumnGlobalEnumeration.ENUM_K_STOCK_OPEN _
                    To ColumnGlobalEnumeration.ENUM_K_STOCK_CLOSE
    
        Worksheets _
            (ActiveSheet.Name) _
                .Columns _
                    (indexCounterVariable) _
                .NumberFormat _
                    = "$#,##0.00"
    
    Next indexCounterVariable ' This statement ends the first repetition loop.
    
            
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_STOCK_VOL) _
            .NumberFormat _
                = "#,##0"


    ' These lines of code change the column widths.
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_STOCK_TICKER) _
            .ColumnWidth _
                = 10
    
    Worksheets _
        (ActiveSheet.Name) _
            .Columns _
                (ColumnGlobalEnumeration.ENUM_K_STOCK_DATE) _
            .ColumnWidth _
                = 14
    
    
    For indexCounterVariable _
            = ColumnGlobalEnumeration.ENUM_K_STOCK_OPEN _
                    To ColumnGlobalEnumeration.ENUM_K_STOCK_CLOSE

        Worksheets _
            (ActiveSheet.Name) _
                .Columns _
                    (indexCounterVariable) _
                .ColumnWidth _
                    = 12
            
    Next indexCounterVariable ' This statement ends the second repetition loop.
    
    Worksheets _
            (ActiveSheet.Name) _
                .Columns _
                    (ColumnGlobalEnumeration.ENUM_K_STOCK_VOL) _
                .ColumnWidth _
                    = 15
    
End Sub ' This stastement ends the private subroutine,
' FormatStockDataPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  FormatSummaryDataPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the format for the summary table's columns.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                               Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                       N. James George
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
    
    
    For indexCounterVariable _
            = ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE _
                    To ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE

        Worksheets _
            (ActiveSheet.Name) _
                .Columns _
                    (indexCounterVariable) _
                .ColumnWidth _
                    = 16
            
    Next indexCounterVariable ' This statement ends the repetition loop.
    
    
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
 '       the stock and the summary data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
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
 '  Subroutine Name:  FormatWorkSheetPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the font type and font size for a worksheet.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  ByVal
 '          inputWorksheetParameter
 '                          This parameter is the input worksheet.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
    FormatWorkSheetPrivateSubRoutine _
        (ByVal _
            inputWorksheetParameter _
                As Worksheet)
    
    inputWorksheetParameter _
        .Cells _
            .Font _
                .Name _
                    = "Garamond"
    
    inputWorksheetParameter _
        .Cells _
            .Font _
                .Size _
                    = 14
    
End Sub ' This statement ends the private subroutine,
' FormatWorkSheetPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  CreateSummaryTablePrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine creates the summary table by analyzing the raw stock data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
    CreateSummaryTablePrivateSubRoutine()
    
    ' This lines of code declare variables for the stock data's first row index
    ' in the repetition loop.
    
    Dim _
        firstRowLongVariable _
            As Long
    

    ' These lines of code declare variables for the stock record.  The program uses
    ' these values to calculate the summary table record.
    
    Dim _
        currentTickerNameStringVariable _
            As String
    
    Dim _
        openingPriceCurrencyVariable _
            As Currency
    
    Dim _
        closingPriceCurrencyVariable _
            As Currency
    
    Dim _
        totalStockVolumeVariantVariable _
            As Variant
    
    
    ' This line of code declares the variable for the row index
    ' in the summary table.
    
    Dim _
        summaryTableRowLongVariable _
            As Long
    
    
    ' This subroutine places the titles in the appropriate cells.
    
    SetUpTitlesForSummaryDataPrivateSubRoutine
    
    
    ' These lines of code initialize variables with information from the first row
    ' of the raw stock.
    
    currentTickerNameStringVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_STOCK_TICKER) _
                    .Value
    
    openingPriceCurrencyVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_STOCK_OPEN) _
                    .Value
                
    totalStockVolumeVariantVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_STOCK_VOL) _
                    .Value
                
    
    ' These lines of code set the initial row indices for the original data
    ' and summary tables.
    
    firstRowLongVariable _
        = RowGlobalEnumeration.ENUM_K_FIRST_DATA + 1
    
    summaryTableRowLongVariable _
        = RowGlobalEnumeration.ENUM_K_FIRST_DATA
 
 
    ' This repetition loop runs through all the rows of the original data
    ' and generates the summary table: the loop starts with the second
    ' row of original data.

    For originalRowCounterVariable _
                = firstRowLongVariable _
                        To lastDataRowGlobalLongVariable
    
        If Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (originalRowCounterVariable, _
                         ColumnGlobalEnumeration.ENUM_K_STOCK_TICKER) _
                    .Value _
                        = currentTickerNameStringVariable Then
        
        
            ' If the ticker name is the same, this line of code adds the current stock volume
            ' to the total.
                                    
            totalStockVolumeVariantVariable _
                = totalStockVolumeVariantVariable _
                        + Worksheets _
                                (ActiveSheet.Name) _
                                    .Cells _
                                        (originalRowCounterVariable, _
                                         ColumnGlobalEnumeration.ENUM_K_STOCK_VOL) _
                                    .Value
                            
                            
            ' If the loop has reached the last row the program creates a summary table record.
        
            If originalRowCounterVariable _
                    = lastDataRowGlobalLongVariable Then
                                           
               CreateSummaryDataRowPrivateSubRoutine _
                        currentTickerNameStringVariable, _
                        openingPriceCurrencyVariable, _
                        totalStockVolumeVariantVariable, _
                        summaryTableRowLongVariable, _
                        originalRowCounterVariable, _
                        True
                    
            End If
            
        Else
        
            ' This selection statement executes if the repetition loop has not reached
            ' the end of the data.
        
            If originalRowCounterVariable _
                    <> lastDataRowGlobalLongVariable Then
            
                ' If the current ticker does not match the previous ticker,
                ' the script creates a record.
            
                CreateSummaryDataRowPrivateSubRoutine _
                        currentTickerNameStringVariable, _
                        openingPriceCurrencyVariable, _
                        totalStockVolumeVariantVariable, _
                        summaryTableRowLongVariable, _
                        originalRowCounterVariable, _
                        False
                    
                
                ' These lines of code assign new values to the stock data variables.
                
                currentTickerNameStringVariable _
                    = Worksheets _
                            (ActiveSheet.Name) _
                                .Cells _
                                    (originalRowCounterVariable, _
                                     ColumnGlobalEnumeration.ENUM_K_STOCK_TICKER) _
                                .Value
                
                openingPriceCurrencyVariable _
                    = Worksheets _
                            (ActiveSheet.Name) _
                                .Cells _
                                    (originalRowCounterVariable, _
                                     ColumnGlobalEnumeration.ENUM_K_STOCK_OPEN) _
                                .Value
                
                totalStockVolumeVariantVariable _
                    = Worksheets _
                            (ActiveSheet.Name) _
                                .Cells _
                                    (originalRowCounterVariable, _
                                     ColumnGlobalEnumeration.ENUM_K_STOCK_VOL) _
                                .Value
            
                
               ' This line of code increases the summary table row index
               ' by one for the next record.
            
                summaryTableRowLongVariable _
                    = summaryTableRowLongVariable + 1
                        
            Else
            
                ' These lines of code initialize variables with information
                ' from the stock data's last row.
    
                currentTickerNameStringVariable _
                    = Worksheets _
                            (ActiveSheet.Name) _
                                .Cells _
                                    (originalRowCounterVariable, _
                                     ColumnGlobalEnumeration.ENUM_K_STOCK_TICKER) _
                                .Value
                    
                openingPriceCurrencyVariable _
                    = Worksheets _
                            (ActiveSheet.Name) _
                                .Cells _
                                    (originalRowCounterVariable, _
                                     ColumnGlobalEnumeration.ENUM_K_STOCK_OPEN) _
                                .Value
                            
                totalStockVolumeVariantVariable _
                    = totalStockVolumeVariantVariable _
                        + Worksheets _
                                (ActiveSheet.Name) _
                                    .Cells _
                                        (originalRowCounterVariable, _
                                         ColumnGlobalEnumeration.ENUM_K_STOCK_VOL) _
                                    .Value
                                
                                
                ' The program then creates a record with this information.
                                
                CreateSummaryDataRowPrivateSubRoutine _
                        currentTickerNameStringVariable, _
                        openingPriceCurrencyVariable, _
                        totalStockVolumeVariantVariable, _
                        summaryTableRowLongVariable, _
                        originalRowCounterVariable, _
                        True
            
            End If
            
        End If
        
    Next originalRowCounterVariable ' This statement ends the repetition loop.
    
End Sub ' This statement ends the private subroutine,
' CreateSummaryTablePrivateSubRoutine.

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
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Public Sub _
    CreateChangeTablePrivateSubRoutine()

    FormatChangeDataTitlesPrivateSubRoutine
    
    SetupChangeDataTitlesPrivateSubRoutine
    
    CalculateAndWriteChangeDataPrivateSubRoutine

End Sub ' This statement ends the private subroutine,
' CreateChangeTablePrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  ConvertDataAndSummaryRangesToTablesPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine converts the stock data and summary ranges to tables.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
    ConvertDataAndSummaryRangesToTablesPrivateSubRoutine()
    
    ConvertRangeIntoTablePrivateSubRoutine _
        RowGlobalEnumeration.ENUM_K_TITLE, _
        ColumnGlobalEnumeration.ENUM_K_STOCK_TICKER, _
        "StockData"

    ConvertRangeIntoTablePrivateSubRoutine _
        RowGlobalEnumeration.ENUM_K_TITLE, _
        ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER, _
        "Summary"
    
End Sub ' This statement ends the private subroutine,
' ConvertDataAndSummaryRangesToTablesPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  ChangeStringToDateInDateColumnPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine changes the text strings in the stock data's date column
 '        to a Date type.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
    ChangeStringToDateInDateColumnPrivateSubRoutine()

    '  This line of code declares a variable for the current date.
    
    Dim _
        dateDateVariable _
            As Date
    
    ' These lines of code declare variables for the current year, month, and day.
    
    Dim _
        yearIntegerVariable _
            As Integer
    
    Dim _
        monthIntegerVariable _
            As Integer
    
    Dim _
        dayIntegerVariable _
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
    
    For rowCounterVariable _
            = RowGlobalEnumeration.ENUM_K_FIRST_DATA _
                    To lastDataRowGlobalLongVariable
    
        ' These lines of code parse out the date from the string, YYYYMMDD,
        ' in the current cell and converts it to a Date type.
        
        yearIntegerVariable _
            = Mid _
                    (Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_STOCK_DATE) _
                                    .Value, _
                     yearStartIndexIntegerVariable, _
                     GLOBAL_CONSTANT_YEAR_LENGTH)
        
        monthIntegerVariable _
            = Mid _
                    (Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_STOCK_DATE) _
                                    .Value, _
                     monthStartIndexIntegerVariable, _
                     GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH)
        
        dayIntegerVariable _
            = Mid _
                    (Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_STOCK_DATE) _
                                    .Value, _
                         dayStartIndexVariable, _
                         GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH)
        
        ' This line of code takes the values for year, month, and day, converts them
        ' to a Date type, then  assigns them to the appropriate variable
        
        dateDateVariable _
            = DateSerial _
                    (yearIntegerVariable, _
                    monthIntegerVariable, _
                    dayIntegerVariable)
    
    
        ' This line of code assigns the new date value to the current cell.
        
        Worksheets _
            (ActiveSheet.Name) _
                .Cells _
                    (rowCounterVariable, _
                     ColumnGlobalEnumeration.ENUM_K_STOCK_DATE) _
                        .Value _
                            = dateDateVariable
    
    Next rowCounterVariable ' This statement ends the repetition loop.

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
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
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
 '  Subroutine Name:  CreateSummaryDataRowPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine creates a summary data record based on the parameters.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  ByVal
 '          tickerNameStringParameter
 '                          This parameter holds the name of the stock ticker.
 '  ByVal
 '          openingPriceCurrencyParameter
 '                          This parameter is the first opening price for this
 '                          stock ticker.
 '  ByVal
 '          totalStockVolumeVariantParameter
 '                          This parameter is the total stock volume for this
 '                          stock ticker.
 '  ByVal
 '          summaryRowLongParameter
 '                          This parameter is the current summary table row index.
 '  ByVal
 '          originalRowLongParameter
 '                          This parameter is the current original data row index.
 '  ByVal
 '          lastRowFlagBooleanParameter
 '                          This parameter indicates whether the program
 '                          has reached the last record or not.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
    CreateSummaryDataRowPrivateSubRoutine _
        (ByVal _
            tickerNameStringParameter _
                As String, _
         ByVal _
            openingPriceCurrencyParameter _
                As Currency, _
         ByVal _
            totalStockVolumeVariantParameter _
                As Variant, _
         ByVal _
            summaryRowLongParameter _
                As Long, _
         ByVal _
            originalRowLongParameter _
                As Long, _
        ByVal _
            lastRowFlagBooleanParameter _
                As Boolean)

    ' This line of code declares a variable for the closing price which is different
    ' based on whether the program has reached the last row or not in the
    ' raw stock data
    
    Dim _
        closingPriceCurrencyVariable _
            As Currency


    ' If the script has not reached the last row, the closing price comes
    ' from the previous row in the raw stock data; otherwise, the closing
    ' price comes from the current row.
                  
    If lastRowFlagBooleanParameter _
       = False Then
            
        closingPriceCurrencyVariable _
            = Worksheets _
                    (ActiveSheet.Name) _
                        .Cells _
                            (originalRowLongParameter - 1, _
                             ColumnGlobalEnumeration.ENUM_K_STOCK_CLOSE) _
                                .Value
            
    Else
            
        closingPriceCurrencyVariable _
            = Worksheets _
                    (ActiveSheet.Name) _
                        .Cells _
                            (originalRowLongParameter, _
                             ColumnGlobalEnumeration.ENUM_K_STOCK_CLOSE) _
                                .Value
            
    End If
            
            
    ' These lines of code create a record in the summary data.
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (summaryRowLongParameter, _
                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                    .Value _
                        = tickerNameStringParameter
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (summaryRowLongParameter, _
                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
                    .Value _
                        = CalculateYearlyChangePrivateFunction _
                                (CDbl _
                                    (openingPriceCurrencyParameter), _
                                 CDbl _
                                    (closingPriceCurrencyVariable))
                            
    FormatYearlyChangeCellPrivateSubRoutine _
        summaryRowLongParameter
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (summaryRowLongParameter, _
                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                    .Value _
                        = CalculatePercentChangePrivateFunction _
                                (CDbl _
                                    (openingPriceCurrencyParameter), _
                                 CDbl _
                                    (closingPriceCurrencyVariable))
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (summaryRowLongParameter, _
                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                    .Value _
                        = totalStockVolumeVariantParameter

End Sub ' This statement ends the private subroutine,
' CreateSummaryDataRowPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  FormatChangeDataTitlesPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine formats the row and column titles in the change data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
    FormatChangeDataTitlesPrivateSubRoutine()

    ' These lines of code format the columns and cells of the change data.

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
            
           
    ' These lines of code set the column widths for the change data.
           
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
' FormatChangeDataTitlesPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  SetupChangeDataTitlesPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine writes the column and row titles to the appropriate cells
 '      for the change data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
    SetupChangeDataTitlesPrivateSubRoutine()

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
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_ROW_TITLES) _
            .Value _
                = "Greatest % Decrease"
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_ROW_TITLES) _
            .Value _
                = "Greatest Total Volume"

End Sub ' This statement ends the private subroutine,
' SetupChangeDataTitlesPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  CalculateAndWriteChangeDataPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine calculates the values for the change data based on raw stock
 '      data and writes the results to the change table.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
    CalculateAndWriteChangeDataPrivateSubRoutine()
    
    Dim _
        increaseTickerStringVariable _
            As String
    
    Dim _
        decreaseTickerStringVariable _
            As String
    
    Dim _
        volumeTickerStringVariable _
            As String
    
    Dim _
        increasePercentageDoubleVariable _
            As Double
    
    Dim _
        decreasePercentageDoubleVariable _
            As Double
    
    Dim _
        volumeVariantVariable _
            As Variant

    Dim _
        firstRowLongVariable _
            As Long
    
    Dim _
        lastRowLongVariable _
            As Long
    
    
    ' These lines of code initialize the variables with the first record
    ' in the summary data.
    
    increaseTickerStringVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                    .Value
         
    decreaseTickerStringVariable _
        = increaseTickerStringVariable
        
    volumeTickerStringVariable _
        = increaseTickerStringVariable
        
       
    increasePercentageDoubleVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                    .Value
             
    decreaseTickerStringVariable _
        = increasePercentageDoubleVariable
        
    volumeVariantVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                    .Value
    
    
    ' These lines of code initialize the first and last index of the repetition loop.
        
    firstRowLongVariable _
        = RowGlobalEnumeration.ENUM_K_FIRST_DATA + 1
        
    lastRowLongVariable _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (Rows.Count, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                    .End _
                        (xlUp) _
                    .Row
        
        
    ' This repetition loop starts at the second record of the summary data and,
    ' through comparisons, finds the tickers with the greatest increase, greatest
    ' decrease, and greatest total stock volume.
         
    For rowIndexCounterVariable _
            = firstRowLongVariable _
                    To lastRowLongVariable
    
        ' If a record has a larger change in percentage than the previous holder,
        ' set it as the new leader in percentage increase.
    
        If Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (rowIndexCounterVariable, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                    .Value _
            > increasePercentageDoubleVariable Then
        
            increaseTickerStringVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                            .Value
                        
            increasePercentageDoubleVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                            .Value
        
        End If
        
        
        ' If a record has a smaller change in percentage than the previous holder,
        ' set it as the new leader in percentage decrease.
        
        If Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (rowIndexCounterVariable, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                    .Value _
            < decreasePercentageDoubleVariable Then
        
            decreaseTickerStringVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                            .Value
                        
            decreasePercentageDoubleVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                            .Value
        
        End If
        
        
        ' If a record has a larger total stock volume than the previous holder,
        ' set it as the new leader in total stock volume.
        
        If Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (rowIndexCounterVariable, _
                         ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                    .Value _
            > volumeVariantVariable Then
        
            volumeTickerStringVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TICKER) _
                            .Value
                        
            volumeVariantVariable _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rowIndexCounterVariable, _
                                 ColumnGlobalEnumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                            .Value
        
        End If
    
    Next rowIndexCounterVariable ' This statement ends the repetition loop.
             
    
    ' These lines of code write the results to the change data.
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_TICKERS) _
            .Value _
                = increaseTickerStringVariable
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_FIRST_DATA, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES) _
            .Value _
                = increasePercentageDoubleVariable
            
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_PERCENT_DECREASE, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_TICKERS) _
            .Value _
                = decreaseTickerStringVariable
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_PERCENT_DECREASE, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES) _
            .Value _
                = decreasePercentageDoubleVariable
                
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_TICKERS) _
            .Value _
                = volumeTickerStringVariable
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (RowGlobalEnumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
                 ColumnGlobalEnumeration.ENUM_K_CHANGE_VALUES) _
            .Value _
                = volumeVariantVariable
    
End Sub ' This statement ends the private subroutine,
' CalculateAndWriteChangeDataPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  ConvertRangeIntoTablePrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine converts a range of data into a table.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  Integer
 '          rowIntegerParameter
 '                          This is the row number of the upper left corner of the range.
 '  Integer
 '          columnIntegerParameter
 '                          This is the column number of the upper left corner of the range.
 '  String
 '          tableNameStringParameter
 '                          This is the name of the new table.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
    ConvertRangeIntoTablePrivateSubRoutine _
        (ByVal _
            rowIntegerParameter _
                As Integer, _
         ByVal _
            columnIntegerParameter _
                As Integer, _
         ByVal _
            tableNameStringParameter _
                As String)
    
    Dim _
        tempListObject _
            As ListObject
    
    Dim _
        tableNameStringVariable _
            As String
    
    
    ' This line of code selects the range of data.
    
    Worksheets _
        (ActiveSheet.Name) _
            .Cells _
                (rowIntegerParameter, _
                 columnIntegerParameter) _
            .Select
    
    
    ' This line of code assigns the selected range of data to a ListObject.
    
    On Error Resume Next
        
        Set tempListObject _
            = Worksheets _
                    (ActiveSheet.Name) _
                        .Cells _
                            (rowIntegerParameter, _
                             columnIntegerParameter) _
                        .ListObject
    
    On Error GoTo 0
    
    
    ' If there is no ListObject, the script converts the range to a table.
    
    If tempListObject Is Nothing Then
    
        tableNameStringVariable _
            = tableNameStringParameter & ActiveSheet.Name & "Table"
            
        ActiveSheet.ListObjects _
            .Add _
                (xlSrcRange, _
                Selection.CurrentRegion, _
                , _
                xlYes) _
            .Name _
                = tableNameStringVariable
    
    End If
                
End Sub ' This statement ends the private subroutine,
' ConvertRangeIntoTablePrivateSubRoutine.

'*******************************************************************************************
 '
 '  Subroutine Name:  FormatYearlyChangeCellPrivateSubRoutine
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This function formats the newly assigned yearly change cell in the summary data
 '      based on the row index.
 '
 '  Subroutine Parameters:
 '
 '  Type     Name                  Description
 '  -----   -------------   ----------------------------------------------
 '  ByVal
 '          rowIndexIntegerParameter
 '                          This parameter holds the row index for the current record
 '                          in the summary table.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Sub _
    FormatYearlyChangeCellPrivateSubRoutine _
        (ByVal _
            rowIndexIntegerParameter _
                As Integer)
    
    ' If the yearly change is zero or positive, the script changes the background color
    ' to green.

    If Worksheets _
            (ActiveSheet.Name) _
                .Cells _
                    (rowIndexIntegerParameter, _
                     ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
                .Value _
        >= 0 Then
        
        Worksheets _
            (ActiveSheet.Name) _
                .Cells _
                    (rowIndexIntegerParameter, _
                     ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
                .Interior _
                .ColorIndex _
                    = 4
        
    Else ' If the yearly change is negative, the script changes the background color to red.
    
        Worksheets _
            (ActiveSheet.Name) _
                .Cells _
                    (rowIndexIntegerParameter, _
                     ColumnGlobalEnumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
                .Interior _
                .ColorIndex _
                    = 3
    
    End If
    
End Sub ' This statement ends the private subroutine,
' FormatYearlyChangeCellPrivateSubRoutine.

'*******************************************************************************************
 '
 '  Function Name:  CalculateYearlyChangePrivateFunction
 '
 '  Function Type: Private
 '
 '  Function Description:
 '      This function calculates the yearly change between the first opening price
 '      and the last closing price of the year for a single ticker.
 '
 '  Function Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  ByVal
 '          openingPriceDoubleParameter
 '                          This parameter holds the first opening price of a ticker.
 '  ByVal
 '          closingPriceDoubleParameter
 '                          This parameter holds the last closing price of a ticker.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Function _
    CalculateYearlyChangePrivateFunction _
        (ByVal _
            openingPriceDoubleParameter _
                As Double, _
        ByVal _
            closingPriceDoubleParameter _
                As Double) _
As Double
    
    CalculateYearlyChangePrivateFunction _
        = closingPriceDoubleParameter _
            - openingPriceDoubleParameter
            
End Function ' This statement ends the private function,
' CalculateYearlyChangePrivateFunction.

'*******************************************************************************************
 '
 '  Function Name:  CalculatePercentChangePrivateFunction
 '
 '  Function Type: Private
 '
 '  Function Description:
 '      This function calculates the percent change between the first opening price
 '      and the last closing price of the year for a ticker.
 '
 '  Function Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  ByVal
 '          openingPriceDoubleParameter
 '                          This parameter holds the first opening price of a ticker.
 '  ByVal
 '          closingPriceDoubleParameter
 '                          This parameter holds the last closing price of a ticker.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      N. James George
 '
 '******************************************************************************************/

Private Function _
    CalculatePercentChangePrivateFunction( _
        ByVal _
            openingPriceDoubleParameter _
                As Double, _
        ByVal _
            closingPriceDoubleParameter _
                As Double) _
As Double

    CalculatePercentChangePrivateFunction _
        = (closingPriceDoubleParameter - openingPriceDoubleParameter) _
            / openingPriceDoubleParameter

End Function ' This statement ends the private function,
' CalculatePercentChangePrivateFunction.
