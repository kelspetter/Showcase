Attribute VB_Name = "PivotCreator"
Option Explicit
Option Compare Text

'NOTE: Requires reference to Microsoft Scripting Runtime (under Tools > References...)

'The name of the template sheets
Private Const PivotTemplateSheet As String = "_PivotTemplates_"

Private PivotDef As Scripting.Dictionary        'Pivot definition dictionary
Private PivotFields As Scripting.Dictionary     'List of fields provided in the pivot table

'The text string that denotes a new Pivot Definition
Public Const PivotDefString As String = "Pivot Definition"

'Logging
Private Enum LogLevel
    LogLevelEmpty = 0
    LogLevelNote = 1
    LogLevelWarn = 2
    LogLevelError = 3
End Enum

Private Const LogEnabled As Boolean = True
Private Const logSheet As String = "_Logs_"

Private wsLog As Worksheet
Private LogLine As Integer
Private logPivot As String

'This routine creates a pivot based upon the currently selected pivot definition
Public Sub CreateFromSelected()
    Dim c As Range
    Dim found As Boolean
    Dim flg As Boolean
    
    SetupLogSheet
    
    Set c = ActiveCell
    
    If c <> PivotDefString Then
        'Find the left-most part of the range
        flg = False
        Do Until flg = True
            If c.Column = 1 Then
                'Is the first column
                flg = True
            ElseIf c.Offset(0, -1) = "" Then
                'There is a blank next to the first column
                flg = True
            Else
                Set c = c.End(xlToLeft)
            End If
        Loop
        
        'Find the top-most part of the range
        flg = False
        Do Until flg = True
            If c.Row = 1 Then
                'Is the first row
                flg = True
            ElseIf c.Offset(-1) = "" Then
                'There is a blank above the first row
                flg = True
            Else
                Set c = c.End(xlUp)
            End If
        Loop
    End If
    
    CreatePivotFromDefinition c
End Sub

'This routine creates the pivots based upon the chart definitions found on the current worksheet
Public Sub CreateFromWorksheet()
    Dim defcell As Range                    'Pivot definition cell
    Dim FirstCell As String                 'First pivot definition cell found
    Dim curWS As Worksheet                  'Current worksheet
    
    SetupLogSheet
    
    Set curWS = ActiveSheet
    'Find cells in the active worksheet(s) that have a value 'Pivot Definition'
    Set defcell = Cells.Find(PivotDefString, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If Not defcell Is Nothing Then
        FirstCell = defcell.Address
        
        Do
            CreatePivotFromDefinition defcell
            
            curWS.Activate
            Set defcell = Cells.FindNext(defcell)
        Loop While defcell.Address <> FirstCell
    End If
    
End Sub

'This routine processes a definition in order to create a pivot
Private Sub CreatePivotFromDefinition(ByVal defcell As Range)
    'Check there is an actual pivot definition
    If defcell.Value <> PivotDefString Then
        Log LogLevelError, "No pivot definition at sheet " & ActiveSheet.Name & ", cell " & defcell.Address
        Exit Sub
    End If
    
    'Set up the PivotDef dictionary
    Set PivotDef = New Scripting.Dictionary
    PivotDef.CompareMode = TextCompare
    
    PivotDef("pivot definition source") = defcell.Address
    
    'Process the key-value pairs under the pivot definition
    Dim cOffset As Integer, cKey As String, cVal As String
    cOffset = 1
    
On Error GoTo ErrCode
    Do Until defcell.Offset(cOffset) = ""
        cKey = KeyCheck(Trim(defcell.Offset(cOffset)))
        cVal = Trim(defcell.Offset(cOffset, 1))
        
        Log LogLevelNote, "Evaluating key " & cKey
        
        'If the key is 'Template', process the template first
        If cKey = "Template" Then
            ProcessPivotTemplate cVal
        'Special case for a number format - get the number format in the cell
        ElseIf Left(cKey, 14) = "number format " And Left(cKey, 20) <> "number format string" Then
            PivotDef("number format string " & Mid(cKey, 15)) = defcell.Offset(cOffset, 1).NumberFormat
        'Special case for a title
        ElseIf cKey = "title" Then
            'Ensure we don't have a blank value, in which case just ignore it
            If cVal <> "" Then
                PivotDef(cKey) = "'" & defcell.Parent.Name & "'!" & defcell.Offset(cOffset, 1).Address
            End If
        'Special case for a color - get the color index of the interior (background colour)
        ElseIf Right(cKey, 7) = " colour" Then
            PivotDef(cKey) = defcell.Offset(cOffset, 1).Interior.Color
        'Special case for a colour rgb - use the value in the cell, but override the colour (not colour rgb)
        ElseIf Right(cKey, 11) = " colour rgb" Then
            PivotDef(Left(cKey, Len(cKey) - 4)) = cVal
        'Special case for a font - get the attributes based upon the actual font in the cell
        ElseIf Right(cKey, 5) = " font" Then
            ProcessFont cKey, defcell.Offset(cOffset, 1)
        Else
            PivotDef(cKey) = cVal
        End If
        
        cOffset = cOffset + 1
    Loop
On Error GoTo 0
    
    CreatePivot
    
ExitCode:
    Log LogLevelNote, "Finished creating pivot from definition on worksheet " & defcell.Parent.Name & ", cell " & defcell.Address
    Set PivotDef = Nothing
    
    Exit Sub
    
ErrCode:
    Log LogLevelError, "Unexpected error while processing pivot definition: " & Err.Description
    Resume ExitCode
End Sub

'This routine processes a pivot template, adding its values to the PivotDef dictionary
Private Sub ProcessPivotTemplate(ByVal TemplateName As String)
    Dim wsTemplate As Worksheet                 'The _PivotTemplates_ worksheet, which must be present
    Dim defcell As Range                        'Template cell
    Dim cOffset As Integer, cKey As String, cVal As String
    Dim found As Boolean
    
    Log LogLevelNote, "Looking for Pivot Template: " & TemplateName
    
    Set wsTemplate = Worksheets(PivotTemplateSheet)
    found = False
    
    For Each defcell In wsTemplate.Range("1:1")
    
        If defcell = TemplateName Then
            Log LogLevelNote, "Found " & TemplateName & " at cell " & defcell.Address
            found = True
        
            'Process the key-value pairs under the template
            cOffset = 1
            
            Do Until defcell.Offset(cOffset) = ""
                cKey = KeyCheck(Trim(defcell.Offset(cOffset)))
                cVal = Trim(defcell.Offset(cOffset, 1))
        
                Log LogLevelNote, "Evaluating key " & cKey
                
                'If the key is 'Template', process the template first
                If cKey = "Template" Then
                    ProcessPivotTemplate cVal
                'Special case for a number format - get the number format in the cell
                ElseIf Left(cKey, 14) = "number format " And Left(cKey, 20) <> "number format string" Then
                    PivotDef("number format string " & Mid(cKey, 15)) = defcell.Offset(cOffset, 1).NumberFormat
                'Special case for a color - get the colour index of the interior (background colour)
                ElseIf Right(cKey, 7) = " colour" Then
                    PivotDef(cKey) = defcell.Offset(cOffset, 1).Interior.Color
                'Special case for a colour rgb - use the value in the cell, but override the colour (not colour rgb)
                ElseIf Right(cKey, 11) = " colour rgb" Then
                    PivotDef(Left(cKey, Len(cKey) - 4)) = cVal
                'Special case for a font - get the attributes based upon the actual font in the cell
                ElseIf Right(cKey, 5) = " font" Then
                    ProcessFont cKey, defcell.Offset(cOffset, 1)
                Else
                    PivotDef(cKey) = cVal
                End If
                
                cOffset = cOffset + 1
            Loop
            
            Exit For
        End If
        
    Next
    
    If Not found Then Log LogLevelWarn, "Unable to find template " & TemplateName
End Sub

'This routine creates the pivot based upon the provided pivot definition dictionary
Private Sub CreatePivot()
    Dim pvtConn As WorkbookConnection   'Workbook connection for pivot cache
    Dim pvtCache As PivotCache          'Pivot cache object
    Dim pvtTable As PivotTable          'Pivot table object
    Dim pvtTableIt As PivotTable        'Iterator of pivot table objects
    
    Dim pvtId As String                 'Pivot ID
    Dim pvtCacheSource As PivotCache    'Pivot cache source object
    
    Dim defaultSeparator As String      'Default separator for lists of values
    Dim filterFieldVals As Scripting.Dictionary   'List of values that are to be filtered
    
    Dim fld As PivotField               'Pivot Field object
    Dim fldFound As Boolean             'Whether or not the requested field has been found
    Dim fldType As XlPivotFieldOrientation  'Field type / orientation
    Dim calcFldName As String           'Calculated field name
    Dim calcFldFormula As String        'Calculated field formula
    
    Dim srcWorksheet As Worksheet       'Source worksheet
    Dim srcCells As Range               'Source cell range
    Dim srcString As String             'String representation of source
    
    Dim dstWorksheet As Worksheet       'Destination worksheet
    Dim dstCells As Range               'Destination cell range
    Dim dstString As String             'String representation of destination
    Dim dstWorkbook As Workbook         'Destination workbook (used to create pivot cache)
    Dim filterOffset As Long            'Offset the destination by this number of rows to cater for filters
    
    Dim placeUnderRows As Integer       'Number of rows to place pivot table under another
    Dim placeUnderId As String          'ID of pivot table to place this one under
    Dim placeUnderFound As Boolean      'Flag to show we have found the required pivot table
    Dim placeUnderPvtTable As PivotTable    'The pivot table we are placing this one under
    Dim placeAtRow As Long              'Row to place pivot table at
    Dim placeAtColumn As Long           'Column to place pivot table at
    Dim ws As Worksheet                 'Temp worksheet var
    Dim c As Range                      'Temp range var
    
    Dim titleRows As Integer            'Number of rows to place between title and pivot table
    
    Dim nextRow As Integer, nextCol As Integer 'Position of next row/col/filter
    Dim pos As Integer                  'Position the field will be put in
    Dim i As Integer
    
    Dim numVal As Double, boolVal As Boolean, strVal As String 'Used to return values from GetString/Num/Bool

    'Variable declarations for filtering
    Dim hasFilter As Boolean
    Dim filterType As XlPivotFilterType, filterTypeGroup As Integer
    Dim lblValue1 As String, lblValue2 As String
    Dim filterValue1 As Double, filterValue2 As Double
    Dim dateValue1 As String, dateValue2 As String
    Dim topN As Double
    Dim wholeDays As Boolean
    Dim filterFieldName As String, filterField As PivotField
    Dim hasValue1 As Boolean, hasValue2 As Boolean
    Dim hasValueField As Boolean, hasTopN As Boolean
    Dim hasDate1 As Boolean, hasDate2 As Boolean, hasWholeDays As Boolean
    
    'Set the default separator as a comma
    defaultSeparator = ","
    
On Error GoTo MissingRequirement
    If (PivotDef.Exists("source") = False And PivotDef.Exists("source cache") = False) Or _
        (PivotDef.Exists("destination") = False And PivotDef.Exists("place under") = False) Then
        Err.Raise vbError, "CreatePivot requirement check", "Missing Requirement"
    End If
On Error GoTo 0

    'Get a full evaluation of the source and destination
    If PivotDef.Exists("source") Then ParseSource PivotDef("source"), srcWorksheet, srcCells, srcString
    
    'Set the default place under rows
    placeUnderRows = 3
    
    placeUnderFound = False
    
    'If we are placing under another pivot table...
    If GetString("place under", placeUnderId) Then
        'Check if there is a different place under rows value
        If GetNumber("place under rows", numVal) Then
            If numVal >= 0 Then placeUnderRows = numVal
        End If
        
        'Find the pivot that we need to place under
        For Each ws In ThisWorkbook.Worksheets
            For Each pvtTableIt In ws.PivotTables
                If pvtTableIt.Name = placeUnderId Then
                    placeUnderFound = True
                    
                    Set dstWorksheet = pvtTableIt.TableRange2.Parent
                    Set c = pvtTableIt.TableRange2.Range("A1")
                    placeAtColumn = c.Column
                    placeAtRow = c.Row + pvtTableIt.TableRange2.Rows.Count + placeUnderRows
                    
                    Set dstCells = dstWorksheet.Cells(placeAtRow, placeAtColumn)
                    dstString = "'" & dstWorksheet.Name & "'!" & dstCells.Address
                    
                    Exit For
                End If
            Next pvtTableIt
            
            If placeUnderFound Then Exit For
        Next ws
        
        'We were unable to find the referenced pivot table, so raise an error
        If placeUnderFound = False Then
            Log LogLevelError, "Place under specified, but pivot table with ID " & placeUnderId & " not found"
            Exit Sub
        End If
    'Otherwise use the destination
    Else
        ParseDestination PivotDef("destination"), dstWorksheet, dstCells
    End If

On Error Resume Next
    'Get the ID of the pivot, so we can delete an existing pivot if any
    GetString "id", pvtId
    
    If pvtId <> "" Then
        Log LogLevelNote, "Deleting any existing pivots with ID " & pvtId & " on worksheet " & dstWorksheet.Name
        
        For Each pvtTableIt In dstWorksheet.PivotTables
            If pvtTableIt.Name = pvtId Then
                pvtTableIt.TableRange2.Clear
            End If
        Next
    
    End If
On Error GoTo 0
    
    'From this point forward, trap any errors but continue to the next line.
On Error GoTo ErrCode

    'If there is a clear method to be called, do it now
    If GetString("clear method", strVal) Then
        If strVal = "all" Then
            dstCells.Parent.Cells.Clear
        ElseIf strVal = "below" Then
            Range(dstCells, dstCells.SpecialCells(xlCellTypeLastCell)).EntireRow.Clear
        ElseIf strVal = "block" Then
            Range(dstCells, dstCells.SpecialCells(xlCellTypeLastCell)).Clear
        End If
    End If
    
    'If there is a title, then copy the value and format, and move the pivot down
    If GetString("title", strVal) Then
        'First, put the title where dstCells is currently
        Range(strVal).Copy
        dstCells.PasteSpecial xlPasteFormats
        dstCells.PasteSpecial xlPasteValues
        Application.CutCopyMode = xlCopy
        
        'Move dstCells down
        titleRows = 1 'Default number of rows between
        If GetNumber("title rows", numVal) Then titleRows = CInt(numVal)  'Override if defined
        Set dstCells = dstCells.Offset(titleRows + 1)
    End If
    
    'Before we create the pivot table, we need to cater for the additional rows needed for each filter.
    'We want to ensure that the pivot table is created at the location specified, not above it, which
    'is what happens when we start adding filters. So we will 'roughly' get the number of filters, and
    'add that many rows to the destination cell.
    filterOffset = 0
    i = 1
    
    Do
        fldFound = False
        
        If GetString("add field " & i, strVal) Then
            fldFound = True
            
            'By default a field is a filter
            If GetString("type field " & i, strVal) = False Then
                filterOffset = filterOffset + 1
            'Get the declared type if present, and if it is a filter, add to the offset
            ElseIf GetString("type field " & i, strVal) And strVal = "filter" Then
                filterOffset = filterOffset + 1
            End If
        End If
        
        i = i + 1
    Loop Until fldFound = False
    
    'Change the destination by the number of filters + 1 rows
    If filterOffset > 0 Then Set dstCells = dstCells.Offset(filterOffset + 1)

    'Create a master pivot table
    If GetString("type", strVal) And strVal = "master" Then
        Set dstWorkbook = dstWorksheet.Parent
        Set pvtCache = dstWorkbook.PivotCaches.Create(xlDatabase, srcCells, 8)
        
        If pvtId <> "" Then
            Set pvtTable = pvtCache.CreatePivotTable(TableDestination:=dstCells, TableName:=pvtId, DefaultVersion:=8)
        Else
            Set pvtTable = pvtCache.CreatePivotTable(TableDestination:=dstCells, DefaultVersion:=8)
        End If
    'If the type is OLAP, create a new OLAP pivot cache
'    ElseIf strVal = "olap" Then
'
'        'Create an OLAP connection so that we have Distinct Count available
'        Set pvtConn = dstWorkbook.Connections.Add2(pvtId, "Workbook connection for Pivot Cache/Table " & pvtId, _
'            "WORKSHEET;" & srcWorksheet.Parent.Name, srcWorksheet.Name & "!" & srcCells.Address, _
'            xlCmdExcel, True, False _
'        )
'
'        'Create the pivot cache using the OLAP connection
'        Set pvtCache = dstWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:=pvtConn, Version:=8)
        
    Else
        'It is not a master pivot table, so we need to locate the pivot table with the specified name
        i = dstWorksheet.PivotTables.Count
        If GetString("source cache", strVal) And strVal <> "" Then
            Do Until i = 0
            
                If dstWorksheet.PivotTables(i).Name = strVal Then
                    Set pvtCacheSource = dstWorksheet.PivotTables(i).PivotCache
                    
                    If pvtId <> "" Then
                        Set pvtTable = pvtCacheSource.CreatePivotTable(TableDestination:=dstCells, TableName:=pvtId, DefaultVersion:=8)
                    Else
                        Set pvtTable = pvtCacheSource.CreatePivotTable(TableDestination:=dstCells, DefaultVersion:=8)
                    End If
                    
                    Exit Do
                End If
                
                i = i - 1
            Loop
            
            If i = 0 Then
                Log LogLevelError, "Unable to locate pivot cache " & strVal
                GoTo ExitCode
            End If
        Else
            Log LogLevelError, "Normal pivot table must have a Source Cache attribute"
            GoTo ExitCode
        End If
    End If
    
    'Pivot table style
    If GetString("style", strVal) Then pvtTable.TableStyle2 = strVal
    
    'Pivot table layout options
    If GetBool("grand total columns", boolVal) Then pvtTable.ColumnGrand = boolVal
    If GetBool("grand total rows", boolVal) Then pvtTable.RowGrand = boolVal
    If GetString("layout", strVal) Then
        Select Case strVal
            Case "Compact"
                pvtTable.RowAxisLayout xlCompactRow
            Case "Outline"
                pvtTable.RowAxisLayout xlOutlineRow
            Case "Tabular"
                pvtTable.RowAxisLayout xlTabularRow
            Case Else
                Log LogLevelWarn, "Unexpected layout value '" & strVal & "', ignoring"
        End Select
    End If
    
    'If a new default separator is specified, get its value
    If GetString("default separator", strVal) Then
        If strVal <> "" And Left(strVal, 1) <> " " Then defaultSeparator = Left(strVal, 1)
    End If
    
    'Setup for adding calculated fields
    'Calculated fields cannot be created when the pivot table is empty. So to workaround this,
    'add the first field as a filter, add the calculated fields, then hide the first field
    
    'Temporarily add the first field
    pvtTable.PivotFields(1).Orientation = xlPageField
    
    'Add calculated fields
    i = 1
    
    Do
        fldFound = False
        
        'If we are adding a calculated field...
        If GetString("add calculated field " & i, calcFldName) Then
            'Get the formula
            If GetString("formula calculated field " & i, calcFldFormula) Then
                fldFound = True
                pvtTable.CalculatedFields.Add calcFldName, calcFldFormula
            End If
        End If
        
        i = i + 1
        
    Loop Until fldFound = False
    
    'Remove the first field
    pvtTable.PivotFields(1).Orientation = xlHidden
    
    'Identify fields in the pivot table plus the calculated fields
    Set PivotFields = New Scripting.Dictionary
    
    For Each fld In pvtTable.PivotFields
        PivotFields.Add fld.Name, fld
    Next fld
    
    'Add data/value fields to the pivot table
    i = 1
    
    Do
        fldFound = False
        
        'Handle data fields
        Dim func As XlConsolidationFunction
        Dim sCaption As String, hasCaption As Boolean
        Dim sNumFmt As String, hasNumFmt As Boolean
        Dim sCalc As XlPivotFieldCalculation, hasCalc As Boolean
        Dim sBaseField As String, hasBaseField As Boolean
        Dim sBaseItem As String, hasBaseItem As Boolean
        
        'If we are adding a value...
        If GetString("add value " & i, strVal) Then
            'And we have found the field...
            If GetPivotFieldByName(pvtTable, strVal, fld) Then
                fldFound = True
                
                'Get the aggregation
                func = xlCount
                If GetString("agg value " & i, strVal) Then
                    Select Case strVal
                        Case "count"
                            func = xlCount
                        Case "sum"
                            func = xlSum
                        Case "average", "avg"
                            func = xlAverage
                        Case "maximum", "max"
                            func = xlMax
                        Case "minimum", "min"
                            func = xlMin
                        Case "product", "prod"
                            func = xlProduct
                        Case "count numbers", "count n"
                            func = xlCountNums
                        Case "standard deviation", "stddev"
                            func = xlStDev
                        Case "standard deviation population", "stddevp"
                            func = xlStDevP
                        Case "variance", "var"
                            func = xlVar
                        Case "variance population", "varp"
                            func = xlVarP
                        Case Else
                            Log LogLevelWarn, "Invalid aggregation '" & strVal & "', defaulting to Count"
                    End Select
                End If
                
                'Get the caption
                hasCaption = GetString("caption value " & i, sCaption)
                
                'Get the number format
                hasNumFmt = GetString("number format string value " & i, sNumFmt)
                
                'Get the calculation
                hasCalc = GetString("calculation value " & i, strVal)
                If hasCalc Then
                    sCalc = xlNoAdditionalCalculation
                    
                    Select Case strVal
                        Case "No Calculation"
                            sCalc = xlNoAdditionalCalculation
                        Case "% of grand total"
                            sCalc = xlPercentOfTotal
                        Case "% of Column Total"
                            sCalc = xlPercentOfColumn
                        Case "% of Row Total"
                            sCalc = xlPercentOfRow
                        Case "% of"
                            sCalc = xlPercentOf
                        Case "% of Parent Row Total"
                            sCalc = xlPercentOfParentRow
                        Case "% of Parent Column Total"
                            sCalc = xlPercentOfParentColumn
                        Case "% of Parent Total"
                            sCalc = xlPercentOfParent
                        Case "Difference From"
                            sCalc = xlDifferenceFrom
                        Case "% Difference From"
                            sCalc = xlPercentDifferenceFrom
                        Case "Running Total In"
                            sCalc = xlRunningTotal
                        Case "% Running Total In"
                            sCalc = xlPercentRunningTotal
                        Case "Rank Smallest to Largest"
                            sCalc = xlRankAscending
                        Case "Rank Largest to Smallest"
                            sCalc = xlRankDecending
                        Case "Index"
                            sCalc = xlIndex
                        Case Else
                            Log LogLevelWarn, "Invalid additional calculation '" & strVal & "', defaulting to no additional calculation"
                    End Select
                End If
                
                'Get the base field and items for the calculation
                hasBaseField = GetString("base field value " & i, sBaseField)
                hasBaseItem = GetString("base item value " & i, sBaseItem)
                
                'Finally add the field
                If hasCaption Then
                    Set fld = pvtTable.AddDataField(fld, sCaption, func)
                Else
                    Set fld = pvtTable.AddDataField(fld, , func)
                End If
                
                'Set the calculation
                If hasCalc Then
                    fld.Calculation = sCalc
                    If hasBaseField Then fld.BaseField = sBaseField
                    If hasBaseItem Then fld.BaseItem = sBaseItem
                End If
                
                'Set the number format
                If hasNumFmt Then
                    fld.NumberFormat = sNumFmt
                End If
            End If
        End If
        
        i = i + 1
    Loop Until fldFound = False
    'End adding data/values to the pivot table
    
    'Add fields to the pivot table
    i = 1
    nextRow = 1
    nextCol = 1
    
    Do
        fldFound = False
        
        'If we are adding a field...
        If GetString("add field " & i, strVal) Then
            'And we have found the field...
            If GetPivotFieldByName(pvtTable, strVal, fld) Then
                fldFound = True
                
                'Get the field type (filter, row, column)
                fldType = xlPageField
                If GetString("type field " & i, strVal) Then
                    Select Case strVal
                        Case "filter"
                            fldType = xlPageField
                            pos = 1
                        Case "row"
                            fldType = xlRowField
                            pos = nextRow
                            nextRow = nextRow + 1
                        Case "column"
                            fldType = xlColumnField
                            pos = nextCol
                            nextCol = nextCol + 1
                        Case "hidden"
                            fldType = xlHidden
                            pos = 1
                        Case Else
                            Log LogLevelWarn, "Invalid field type '" & strVal & "', defaulting to Filter"
                    End Select
                End If
                
                fld.Orientation = fldType
                fld.Position = pos
                
                'Get the caption
                If GetString("caption field " & i, strVal) Then fld.Caption = strVal
                
                'Get the number format
                If GetString("number format string field " & i, strVal) Then fld.NumberFormat = strVal
                
                'Check for show/hide values
                Dim itemlist As Variant, ic As Integer, pvtItem As PivotItem
                
                'Basic filters by labels
                If GetString("show labels field " & i, strVal) Then
                    'Break down the string into separated list of values
                    itemlist = Split(strVal, defaultSeparator, -1, vbTextCompare)
                    Set filterFieldVals = New Scripting.Dictionary
                    
                    'Add the values to a dictionary for easier searching
                    If UBound(itemlist) = -1 Then
                        'In the case of a blank itemlist, we mean literally blank values
                        filterFieldVals("") = ""
                    Else
                        For ic = LBound(itemlist) To UBound(itemlist)
                            filterFieldVals(itemlist(ic)) = itemlist(ic)
                        Next ic
                    End If
                    
                    'Iterate through the values and hide if the value is not in the list for rows and columns
                    For Each pvtItem In fld.PivotItems
                        If filterFieldVals.Exists(pvtItem.Name) = False Then
                            pvtItem.Visible = False
                        Else
                            pvtItem.Visible = True
                        End If
                    Next
                ElseIf GetString("hide labels field " & i, strVal) Then
                    'Break down the string into separated list of values
                    itemlist = Split(strVal, defaultSeparator, -1, vbTextCompare)
                    Set filterFieldVals = New Scripting.Dictionary
                    
                    'Add the values to a dictionary for easier searching
                    If UBound(itemlist) = -1 Then
                        'In the case of a blank itemlist, we mean literally blank values
                        filterFieldVals("") = ""
                    Else
                        For ic = LBound(itemlist) To UBound(itemlist)
                            filterFieldVals(itemlist(ic)) = itemlist(ic)
                        Next ic
                    End If
                    
                    'Iterate through the values and show if the value is not in the list for rows and columns
                    For Each pvtItem In fld.PivotItems
                        If filterFieldVals.Exists(pvtItem.Name) = True Then
                            pvtItem.Visible = False
                        Else
                            pvtItem.Visible = True
                        End If
                    Next
                ElseIf GetString("show single label field " & i, strVal) Then
                    'Check it is a filter
                    If fld.Orientation = xlPageField Then
                        'A single string is to be shown, so hide everything else
                        fld.ClearAllFilters
                        fld.CurrentPage = strVal
                    Else
                        Log LogLevelError, "Show Single Value Field cannot be called on non-filter field"
                    End If
                'Advanced filters by labels
                ElseIf GetString("label filter type field " & i, strVal) Then
                    hasFilter = True
                    filterTypeGroup = 1 '1 = Filter requires one parameter
                    
                    Select Case strVal
                        
                        Case "Equals", "Equal", "="
                            filterType = xlCaptionEquals
                        Case "Does Not Equal", "Not Equal", "!=", "<>"
                            filterType = xlCaptionDoesNotEqual
                        Case "Contains"
                            filterType = xlCaptionContains
                        Case "Does Not Contain"
                            filterType = xlCaptionDoesNotContain
                        Case "Greater", ">"
                            filterType = xlCaptionIsGreaterThan
                        Case "Greater or Equal", ">="
                            filterType = xlCaptionIsGreaterThanOrEqualTo
                        Case "Less Than", "<"
                            filterType = xlCaptionIsLessThan
                        Case "Less Than or Equal", "<="
                            filterType = xlCaptionIsLessThanOrEqualTo
                        Case "Begins With"
                            filterType = xlCaptionBeginsWith
                        Case "Does Not Begin With"
                            filterType = xlCaptionDoesNotBeginWith
                        Case "Ends With"
                            filterType = xlCaptionEndsWith
                        Case "Does Not End With"
                            filterType = xlCaptionDoesNotEndWith
                        Case "Between"
                            filterType = xlCaptionIsBetween
                            filterTypeGroup = 2 '2 = Filter requires two parameters
                        Case "Not Between"
                            filterType = xlCaptionIsNotBetween
                            filterTypeGroup = 2
                        Case "None"
                            hasFilter = False
                        Case Else
                            hasFilter = False
                            Log LogLevelWarn, "Unknown label filter type '" & strVal & "', ignoring"
                    End Select
                    
                    If hasFilter Then
                        'Get the values for the filters, etc
                        hasValue1 = GetString("label filter value 1 field " & i, lblValue1)
                        hasValue2 = GetString("label filter value 2 field " & i, lblValue2)
                        
                        If filterTypeGroup = 1 Then
                            If hasValue1 Then
                                fld.ClearAllFilters
                                fld.PivotFilters.Add2 filterType, , lblValue1
                            Else
                                Log LogLevelWarn, "Filter requires 1 value, but Label Filter Value 1 Field " & i & " not specified"
                            End If
                        ElseIf filterTypeGroup = 2 Then
                            If hasValue1 And hasValue2 Then
                                fld.ClearAllFilters
                                fld.PivotFilters.Add2 filterType, , lblValue1, lblValue2
                            Else
                                Log LogLevelWarn, "Filter requires 2 values, but not all value filter values specified"
                            End If
                        End If
                    Else
                        fld.ClearAllFilters
                    End If
                'Filters by values
                ElseIf GetString("value filter type field " & i, strVal) Then
                    
                    hasFilter = True
                    filterTypeGroup = 1 '1 = Filter requires one parameter
                    
                    Select Case strVal
                        Case "Equals", "Equal", "="
                            filterType = xlValueEquals
                        Case "Does Not Equal", "Not Equal", "!=", "<>"
                            filterType = xlValueDoesNotEqual
                        Case "Greater", ">"
                            filterType = xlValueIsGreaterThan
                        Case "Greater or Equal", ">="
                            filterType = xlValueIsGreaterThanOrEqualTo
                        Case "Less Than", "<"
                            filterType = xlValueIsLessThan
                        Case "Less Than or Equal", "<="
                            filterType = xlValueIsLessThanOrEqualTo
                        Case "Between"
                            filterType = xlValueIsBetween
                            filterTypeGroup = 2 '2 = Filter requires two parameters
                        Case "Not Between"
                            filterType = xlValueIsNotBetween
                            filterTypeGroup = 2
                        Case "Top Items"
                            filterType = xlTopCount
                            filterTypeGroup = 3
                        Case "Top Sum"
                            filterType = xlTopSum
                            filterTypeGroup = 3
                        Case "Top Percent"
                            filterType = xlTopPercent
                            filterTypeGroup = 3
                        Case "Bottom Items"
                            filterType = xlBottomCount
                            filterTypeGroup = 3
                        Case "Bottom Sum"
                            filterType = xlBottomSum
                            filterTypeGroup = 3
                        Case "Bottom Percent"
                            filterType = xlBottomPercent
                            filterTypeGroup = 3
                        Case "None"
                            hasFilter = False
                        Case Else
                            hasFilter = False
                            Log LogLevelWarn, "Unknown value filter type '" & strVal & "', ignoring"
                    End Select
                    
                    If hasFilter Then
                        'Get the values for the filters, etc, after setting some defaults
                        topN = 10
                        
                        hasValueField = GetString("value filter by field " & i, filterFieldName)
                        hasValue1 = GetNumber("value filter value 1 field " & i, filterValue1)
                        hasValue2 = GetNumber("value filter value 2 field " & i, filterValue2)
                        hasTopN = GetNumber("value filter top n field " & i, topN)
                        
                        If hasValueField Then
                            'Get the actual value field
                            Set filterField = pvtTable.PivotFields(filterFieldName)
                        
                            If filterTypeGroup = 1 Then
                                If hasValue1 Then
                                    fld.ClearAllFilters
                                    fld.PivotFilters.Add2 filterType, filterField, filterValue1
                                Else
                                    Log LogLevelWarn, "Filter requires 1 value, but value filter value 1 field " & i & " not specified"
                                End If
                            ElseIf filterTypeGroup = 2 Then
                                If hasValue1 And hasValue2 Then
                                    fld.ClearAllFilters
                                    fld.PivotFilters.Add2 filterType, filterField, filterValue1, filterValue2
                                Else
                                    Log LogLevelWarn, "Filter requires 2 values, but not all value filter values specified"
                                End If
                            ElseIf filterTypeGroup = 3 Then
                                If hasTopN = False Then Log LogLevelNote, "Top/Bottom filter does not have value filter top n field " & i & " specified, using default " & topN
                                fld.ClearAllFilters
                                fld.PivotFilters.Add2 filterType, filterField, topN
                            End If
                            
                        Else
                            Log LogLevelWarn, "Filter must have 'value filter by field " & i & "' attribute, but one was not found"
                        End If
                    Else
                        fld.ClearAllFilters
                    End If
                
                'Filters by dates
                ElseIf GetString("date filter type field " & i, strVal) Then
                    
                    hasFilter = True
                    filterTypeGroup = 0 '1 = Filter requires no parameters
                    
                    Select Case strVal
                        Case "Equals", "Equal", "="
                            filterType = xlSpecificDate
                            filterTypeGroup = 1 '1 = requires one parameter
                        Case "Does Not Equal", "Not Equal", "!=", "<>"
                            filterType = xlNotSpecificDate
                            filterTypeGroup = 1
                        Case "After", "Greater", ">"
                            filterType = xlAfter
                            filterTypeGroup = 1
                        Case "After or Equal", "Greater or Equal", ">="
                            filterType = xlAfterOrEqualTo
                            filterTypeGroup = 1
                        Case "Before", "Less Than", "<"
                            filterType = xlBefore
                            filterTypeGroup = 1
                        Case "Before or Equal", "Less Than or Equal", "<="
                            filterType = xlBeforeOrEqualTo
                            filterTypeGroup = 1
                        Case "Between"
                            filterType = xlDateBetween
                            filterTypeGroup = 2 '2 = Filter requires two parameters
                        Case "Not Between"
                            filterType = xlDateNotBetween
                            filterTypeGroup = 2
                        Case "Tomorrow"
                            filterType = xlDateTomorrow
                        Case "Today"
                            filterType = xlDateToday
                        Case "Yesterday"
                            filterType = xlDateYesterday
                        Case "Next Week"
                            filterType = xlDateNextWeek
                        Case "This Week"
                            filterType = xlDateThisWeek
                        Case "Last Week"
                            filterType = xlDateLastWeek
                        Case "Next Month"
                            filterType = xlDateNextMonth
                        Case "This Month"
                            filterType = xlDateThisMonth
                        Case "Last Month"
                            filterType = xlDateLastMonth
                        Case "Next Quarter"
                            filterType = xlDateNextQuarter
                        Case "This Quarter"
                            filterType = xlDateThisQuarter
                        Case "Last Quarter"
                            filterType = xlDateLastQuarter
                        Case "Next Year"
                            filterType = xlDateNextYear
                        Case "This Year"
                            filterType = xlDateThisYear
                        Case "Last Year"
                            filterType = xlDateLastYear
                        Case "Year to Date"
                            filterType = xlYearToDate
                        Case "Quarter 1"
                            filterType = xlAllDatesInPeriodQuarter1
                        Case "Quarter 2"
                            filterType = xlAllDatesInPeriodQuarter2
                        Case "Quarter 3"
                            filterType = xlAllDatesInPeriodQuarter3
                        Case "Quarter 4"
                            filterType = xlAllDatesInPeriodQuarter4
                        Case "January"
                            filterType = xlAllDatesInPeriodJanuary
                        Case "February"
                            filterType = xlAllDatesInPeriodFebruary
                        Case "March"
                            filterType = xlAllDatesInPeriodMarch
                        Case "April"
                            filterType = xlAllDatesInPeriodApril
                        Case "May"
                            filterType = xlAllDatesInPeriodMay
                        Case "June"
                            filterType = xlAllDatesInPeriodJune
                        Case "July"
                            filterType = xlAllDatesInPeriodJuly
                        Case "August"
                            filterType = xlAllDatesInPeriodAugust
                        Case "September"
                            filterType = xlAllDatesInPeriodSeptember
                        Case "October"
                            filterType = xlAllDatesInPeriodOctober
                        Case "November"
                            filterType = xlAllDatesInPeriodNovember
                        Case "December"
                            filterType = xlAllDatesInPeriodDecember
                        Case "None"
                            hasFilter = False
                        Case Else
                            hasFilter = False
                            Log LogLevelWarn, "Unknown date filter type '" & strVal & "', ignoring"
                    End Select
                    
                    If hasFilter Then
                        'Get the values for the filters, etc, after setting some defaults
                        dateValue1 = ""
                        dateValue2 = ""
                        wholeDays = False
                        
                        hasDate1 = GetString("date filter value 1 field " & i, dateValue1)
                        hasDate2 = GetString("date filter value 2 field " & i, dateValue2)
                        hasWholeDays = GetBool("date filter whole days field " & i, wholeDays)
                        
                        If filterTypeGroup = 0 Then
                            fld.ClearAllFilters
                            fld.PivotFilters.Add2 filterType, WholeDayFilter:=wholeDays
                        ElseIf filterTypeGroup = 1 Then
                            If hasValue1 Then
                                fld.ClearAllFilters
                                fld.PivotFilters.Add2 filterType, Value1:=dateValue1
                                'For some reason, setting WholeDayFilter in the Add2 command doesn't work - do it separately
                                If wholeDays Then fld.PivotFilters(1).WholeDayFilter = True
                            Else
                                Log LogLevelWarn, "Filter requires 1 date, but date filter value 1 field " & i & " not specified"
                            End If
                        ElseIf filterTypeGroup = 2 Then
                            If hasValue1 And hasValue2 Then
                                fld.ClearAllFilters
                                fld.PivotFilters.Add2 filterType, Value1:=dateValue1, Value2:=dateValue2
                                If wholeDays Then fld.PivotFilters(1).WholeDayFilter = True
                            Else
                                Log LogLevelWarn, "Filter requires 2 dates, but not all date filter values specified"
                            End If
                        End If
                    Else
                        fld.ClearAllFilters
                    End If
                End If
            End If
            
            'Sorting
            If GetString("sort order field " & i, strVal) Then
                Dim sortDir As Long
                
                Select Case strVal
                    Case "Asc", "Ascending"
                        sortDir = xlAscending
                    Case "Desc", "Descending"
                        sortDir = xlDescending
                    Case "Manual"
                        sortDir = 0
                    Case Else
                        sortDir = 0
                        Log LogLevelWarn, "Unknown sort field direction '" & strVal & "', not sorting"
                End Select
                
                'By default we will sort by the same field if it is not specified
                Dim sortField As String
                sortField = fld.Caption
                
                GetString "sort by field " & i, sortField
                
                If sortDir <> 0 Then fld.AutoSort sortDir, sortField
            End If
            
        'End adding a field
        End If
        i = i + 1
    Loop Until fldFound = False
    

ExitCode:
    Set PivotFields = Nothing
    
    Log LogLevelNote, "Finished creating pivot"
    
    Exit Sub
    
ErrCode:
    Log LogLevelError, "Unexpected error whilst creating pivot: " & Err.Description & " Continuing..."
    Resume Next
    
MissingRequirement:
    Log LogLevelError, "A required attribute is missing or invalid, aborting..."
    Resume ExitCode
End Sub

'Checks if the requested field is in the list of fields, and returns it if it is
Private Function GetPivotFieldByName(ByVal pvtTable As PivotTable, ByVal fldName As String, ByRef fld As PivotField) As Boolean
    If PivotFields.Exists(fldName) Then
        Set fld = PivotFields(fldName)
        GetPivotFieldByName = True
    Else
        GetPivotFieldByName = False
    End If
End Function

'Replaces words in a key with alternatives, e.g. color = colour, to simplify code
Private Function KeyCheck(ByVal key As String) As String
    KeyCheck = Replace(key, "color", "colour", compare:=vbTextCompare)
End Function

'Check if a numeric value for the key exists, and if so, return to the out variable
Private Function GetNumber(ByVal key As String, ByRef out As Double) As Boolean
    Log LogLevelNote, "Checking for key " & key
    
    key = KeyCheck(key)
    If PivotDef.Exists(key) = False Then
        GetNumber = False
    ElseIf IsNumeric(PivotDef(key)) Then
        out = Val(PivotDef(key))
        GetNumber = True
    Else
        Log LogLevelWarn, "Expected numeric value but got " & PivotDef(key)
        GetNumber = False
    End If
    
    If GetNumber Then Log LogLevelNote, "Found numeric value " & out
End Function

'Check if a string value for the key exists, and if so, return to the out variable
Private Function GetString(ByVal key As String, ByRef out As String) As Boolean
    Log LogLevelNote, "Checking for key " & key
    
    key = KeyCheck(key)
    If PivotDef.Exists(key) = False Then
        GetString = False
    Else
        GetString = True
        out = PivotDef(key)
    End If
    
    If GetString Then Log LogLevelNote, "Found string value " & out
End Function

'Check if a boolean value for the key exists, and if so, return to the out variable
Private Function GetBool(ByVal key As String, ByRef out As Boolean) As Boolean
    Log LogLevelNote, "Checking for key " & key
    
    key = KeyCheck(key)
    If PivotDef.Exists(key) = False Then
        GetBool = False
    ElseIf PivotDef(key) = "True" Or PivotDef(key) = "Yes" Then
        out = True
        GetBool = True
    ElseIf PivotDef(key) = "False" Or PivotDef(key) = "No" Then
        out = False
        GetBool = True
    ElseIf IsNumeric(PivotDef(key)) Then
        If Val(PivotDef(key)) = 0 Then
            out = False
        Else
            out = True
        End If
        GetBool = True
    Else
        Log LogLevelWarn, "Expected boolean value but got " & PivotDef(key)
        GetBool = False
    End If
    
    If GetBool Then Log LogLevelNote, "Found boolean value " & out
End Function

Private Sub ProcessFont(ByVal KeyPrefix As String, ByVal refCell As Range)
    Log LogLevelNote, "Processing font for key prefix " & KeyPrefix & ", cell " & refCell.Address
    
    Dim s As String
    
    PivotDef(KeyPrefix & " Name") = refCell.Font.Name
    PivotDef(KeyPrefix & " Size") = refCell.Font.Size
    PivotDef(KeyPrefix & " Bold") = refCell.Font.Bold
    PivotDef(KeyPrefix & " Italic") = refCell.Font.Italic
    
    Select Case refCell.Font.Underline
        Case msoNoUnderline, -4142:
            s = "None"
        Case msoUnderlineSingleLine, msoUnderlineHeavyLine:
            s = "Single"
        Case msoUnderlineDoubleLine, msoUnderlineDottedLine, -4119:
            s = "Double"
        Case Else:
            Log LogLevelWarn, "Unsupported font underline style encountered"
            s = "None"
    End Select
    PivotDef(KeyPrefix & " Underline") = s
    
    PivotDef(KeyPrefix & " Colour") = refCell.Font.Color
End Sub

Private Sub SetFontAttribs(ByVal KeyPrefix As String, ByVal fObj As Font)
    Log LogLevelNote, "Setting font attributes for key prefix " & KeyPrefix
    
    Dim strOut As String, numOut As Double, boolOut As Boolean
    
    If GetString(KeyPrefix & " font name", strOut) Then fObj.Name = strOut
    If GetNumber(KeyPrefix & " font size", numOut) Then fObj.Size = numOut
    If GetBool(KeyPrefix & " font bold", boolOut) Then fObj.Bold = boolOut
    If GetBool(KeyPrefix & " font italic", boolOut) Then fObj.Italic = boolOut
    If GetString(KeyPrefix & " font underline", strOut) Then
        Select Case strOut
            Case "true", "single", "double":
                fObj.Underline = True
            Case "false", "none":
                fObj.Underline = False
            Case Else:
                Log LogLevelWarn, "Unexpected font underline value " & strOut
        End Select
    End If
    If GetNumber(KeyPrefix & " font colour", numOut) Then fObj.Color = numOut
End Sub

Private Sub SetFont2Attribs(ByVal KeyPrefix As String, ByVal fObj As Font2)
    Log LogLevelNote, "Setting font2 attributes for key prefix " & KeyPrefix
    
    Dim strOut As String, numOut As Double, boolOut As Boolean
    
    If GetString(KeyPrefix & " font name", strOut) Then fObj.Name = strOut
    If GetNumber(KeyPrefix & " font size", numOut) Then fObj.Size = numOut
    If GetBool(KeyPrefix & " font bold", boolOut) Then fObj.Bold = boolOut
    If GetBool(KeyPrefix & " font italic", boolOut) Then fObj.Italic = boolOut
    If GetString(KeyPrefix & " font underline", strOut) Then
        Select Case strOut
            Case "true", "single":
                fObj.UnderlineStyle = msoUnderlineSingleLine
            Case "double":
                fObj.UnderlineStyle = msoUnderlineDoubleLine
            Case "false", "none":
                fObj.UnderlineStyle = msoNoUnderline
            Case Else:
                Log LogLevelWarn, "Unexpected font underline value " & strOut
        End Select
    End If
    If GetNumber(KeyPrefix & " font colour", numOut) Then fObj.Fill.ForeColor.RGB = numOut
End Sub

'Extracts the worksheet and source cell ranges from the source.
'It will include the adjacent cells, and work out where the top-left cell is.
Private Sub ParseSource(ByVal src As String, ByRef outWorksheet As Worksheet, ByRef outCells As Range, ByRef outString As String)
    Log LogLevelNote, "Parsing the source string " & src
    
    Dim c As Range
    Dim topleft As Range, xoff As Integer, yoff As Integer
    Dim bottomright As Range
    Dim flg As Boolean
    Dim sWorksheet As String
    Dim sRange As String
    
On Error GoTo ErrCode
    'Get the worksheet name
    Set outWorksheet = Range(src).Parent
    
    'Determine the top-left of the range
    Set c = Range(src)
    
    'Find the left-most part of the range
    flg = False
    Do Until flg = True
        If c.Column = 1 Then
            'Is the first column
            flg = True
        ElseIf c.Offset(0, -1) = "" Then
            'There is a blank next to the first column
            flg = True
        Else
            Set c = c.End(xlToLeft)
        End If
    Loop
    
    'Find the top-most part of the range
    flg = False
    Do Until flg = True
        If c.Row = 1 Then
            'Is the first row
            flg = True
        ElseIf c.Offset(-1) = "" Then
            'There is a blank above the first row
            flg = True
        Else
            Set c = c.End(xlUp)
        End If
    Loop
    
    Set topleft = c
    
    'Get the x offset
    If topleft.Offset(0, 1) = "" Then
        xoff = 0
    Else
        xoff = topleft.End(xlToRight).Column - topleft.Column
    End If
    
    'Get the y offset
    If topleft.Offset(1) = "" Then
        yoff = 0
    Else
        yoff = topleft.End(xlDown).Row - topleft.Row
    End If
    
    'Get the bottom right
    Set bottomright = topleft.Offset(yoff, xoff)
    
    'Combine the ranges
    Set outCells = Range(topleft, bottomright)
    
    'Final source string
    outString = "'" & outWorksheet.Name & "'!" & outCells.Address
    
    Log LogLevelNote, "Final source string " & outString
    
On Error GoTo 0
    
ExitCode:
    
    Exit Sub
    
ErrCode:
    Log LogLevelError, "Unexpected error parsing source string: " & Err.Description
    Resume ExitCode
End Sub

'Returns the destination as a worksheet and cell reference, or sets a default if not specified
Private Sub ParseDestination(ByVal dest As String, ByRef outWorksheet As Worksheet, ByRef outCell As Range)
    Dim dstAddress As String
    
    Log LogLevelNote, "Parsing destination string " & dest
    
    If dest = "" Then
        Set outWorksheet = ActiveSheet
        Set outCell = Range("L10")
    Else
        'Get the worksheet
        Set outWorksheet = Range(dest).Parent
        
        'Get the first cell
        Set outCell = Range(dest).Range("A1")
    End If
    
    'Final destination string
    dstAddress = "'" & outWorksheet.Name & "'!" & outCell.Address
    
    Log LogLevelNote, "Final destination cell " & dstAddress
End Sub

'Setup the log sheet
Private Sub SetupLogSheet()
    If LogEnabled = False Then Exit Sub
    
    Set wsLog = Nothing
    On Error Resume Next
    Set wsLog = Sheets(logSheet)
    If wsLog Is Nothing Then
        Set wsLog = Sheets.Add
        wsLog.Name = logSheet
    End If
    On Error GoTo 0
    
    wsLog.Cells.Delete
    wsLog.Range("A1") = "Pivot ID / Source"
    wsLog.Range("B1") = "Type"
    wsLog.Range("C1") = "Message"
    
    LogLine = 1
End Sub

'Add a log entry
Private Sub Log(ByVal level As LogLevel, ByVal message As String)
    If LogEnabled = False Then Exit Sub
    
    If Not PivotDef Is Nothing Then
        If PivotDef.Exists("ID") Then
            logPivot = PivotDef("ID")
        ElseIf PivotDef.Exists("Source") Then
            logPivot = "Source: " & PivotDef("Source")
        ElseIf PivotDef.Exists("Pivot Definition Source") Then
            logPivot = "Pivot Definition Source: " & PivotDef("Pivot Definition Source")
        Else
            logPivot = "(Unknown pivot)"
        End If
    Else
        logPivot = ""
    End If
    
    wsLog.Range("A1").Offset(LogLine) = logPivot
    wsLog.Range("B1").Offset(LogLine) = IIf(level = LogLevelNote, "NOTE", IIf(level = LogLevelWarn, "WARN", IIf(level = LogLevelError, "ERROR", "")))
    wsLog.Range("C1").Offset(LogLine) = message
    
    With wsLog.Range(wsLog.Cells(1, 2).Offset(LogLine), wsLog.Cells(1, 3).Offset(LogLine)).Font
        If level = LogLevelError Then
            .ColorIndex = 3 'Red
            .Bold = True
        ElseIf level = LogLevelWarn Then
            .ColorIndex = 45 'Orange
            .Bold = True
        End If
    End With
    
    LogLine = LogLine + 1
End Sub

