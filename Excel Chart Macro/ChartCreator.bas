Attribute VB_Name = "ChartCreator"
Option Explicit
Option Compare Text

'NOTE: Requires reference to Microsoft Scripting Runtime (under Tools > References...)

'The name of the template sheets
Private Const ChartTemplateSheet As String = "_ChartTemplates_"
Private Const SeriesTemplateSheet As String = "_SeriesTemplates_"
Private Const PointTemplateSheet As String = "_PointTemplates_"

Private ChartDef As Scripting.Dictionary        'Chart definition dictionary
Private SeriesDef As Scripting.Dictionary       'Series definition dictionary
Private PointDef As Scripting.Dictionary        'Point definition dictionary

'The text string that denotes a new Chart Definition
Public Const ChartDefString As String = "Chart Definition"

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
Private logChart As String

'This routine creates a chart based upon the currently selected chart definition
Public Sub CreateFromSelected()
    Dim c As Range
    Dim found As Boolean
    Dim flg As Boolean
    
    SetupLogSheet
    
    Set c = ActiveCell
    
    If c <> ChartDefString Then
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
    
    CreateChartFromDefinition c
End Sub

'This routine creates the charts based upon the chart definitions found on the current worksheet
Public Sub CreateFromWorksheet()
    Dim defcell As Range                    'Chart definition cell
    Dim FirstCell As String                 'First chart definition cell found
    Dim curWS As Worksheet                  'Current worksheet
    
    SetupLogSheet
    
    Set curWS = ActiveSheet
    'Find cells in the active worksheet(s) that have a value 'Chart Definition'
    Set defcell = Cells.Find(ChartDefString, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If Not defcell Is Nothing Then
        FirstCell = defcell.Address
        
        Do
            CreateChartFromDefinition defcell
            
            curWS.Activate
            Set defcell = Cells.FindNext(defcell)
        Loop While defcell.Address <> FirstCell
    End If
    
End Sub

'This routine processes a definition in order to create a chart
Private Sub CreateChartFromDefinition(ByVal defcell As Range)
    Log LogLevelNote, "Creating chart from definition on worksheet " & defcell.Parent.Name & ", cell " & defcell.Address
    
    'Check there is an actual chart definition
    If defcell.Value <> ChartDefString Then
        Log LogLevelError, "No chart definition at sheet " & ActiveSheet.Name & ", cell " & defcell.Address
        Exit Sub
    End If
    
    'Set up the ChartDef dictionary
    Set ChartDef = New Scripting.Dictionary
    ChartDef.CompareMode = TextCompare
    
    ChartDef("chart definition source") = defcell.Address
    
    'Process the key-value pairs under the chart definition
    Dim cOffset As Integer, cKey As String, cVal As String
    cOffset = 1
    
On Error GoTo ErrCode
    Do Until defcell.Offset(cOffset) = ""
        cKey = KeyCheck(Trim(defcell.Offset(cOffset)))
        cVal = Trim(defcell.Offset(cOffset, 1))
        
        Log LogLevelNote, "Evaluating key " & cKey
        
        'If the key is 'Template', process the template first
        If cKey = "Template" Then
            ProcessChartTemplate cVal
        'If the key is for a Point template, process it separately
        ElseIf Left(cKey, 7) = "Series " And Right(cKey, 15) = " Point Template" Then
            ProcessPointTemplate cKey, cVal
        'If the key is for a Series template, process it separately
        ElseIf Right(cKey, 9) = " Template" And Left(cKey, 7) = "Series " Then
            ProcessSeriesTemplate cKey, cVal
        'Special case for a color - get the color index of the interior (background colour)
        ElseIf Right(cKey, 7) = " colour" Then
            ChartDef(cKey) = defcell.Offset(cOffset, 1).Interior.Color
        'Special case for a colour rgb - use the value in the cell, but override the colour (not colour rgb)
        ElseIf Right(cKey, 11) = " colour rgb" Then
            ChartDef(Left(cKey, Len(cKey) - 4)) = cVal
        'Special case for a font - get the attributes based upon the actual font in the cell
        ElseIf Right(cKey, 5) = " font" Then
            ProcessFont cKey, defcell.Offset(cOffset, 1)
        Else
            ChartDef(cKey) = cVal
        End If
        
        cOffset = cOffset + 1
    Loop
On Error GoTo 0
    
    CreateChart
    
ExitCode:
    Log LogLevelNote, "Finished creating chart from definition on worksheet " & defcell.Parent.Name & ", cell " & defcell.Address
    Set ChartDef = Nothing
    
    Exit Sub
    
ErrCode:
    Log LogLevelError, "Unexpected error while processing chart definition: " & Err.Description
    Resume ExitCode
End Sub

'This routine processes a chart template, adding its values to the ChartDef dictionary
Private Sub ProcessChartTemplate(ByVal TemplateName As String)
    Dim wsTemplate As Worksheet                 'The _ChartTemplates_ worksheet, which must be present
    Dim defcell As Range                        'Template cell
    Dim cOffset As Integer, cKey As String, cVal As String
    Dim found As Boolean
    
    Log LogLevelNote, "Looking for Chart Template: " & TemplateName
    
    Set wsTemplate = Worksheets(ChartTemplateSheet)
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
                    ProcessChartTemplate cVal
                'If the key is for a Point template, process it separately
                ElseIf Left(cKey, 7) = "Series " And Right(cKey, 15) = " Point Template" Then
                    ProcessPointTemplate cKey, cVal
                'If the key is for a Series template, process it separately
                ElseIf Right(cKey, 9) = " Template" And Left(cKey, 7) = "Series " Then
                    ProcessSeriesTemplate cKey, cVal
                'Special case for a color - get the colour index of the interior (background colour)
                ElseIf Right(cKey, 7) = " colour" Then
                    ChartDef(cKey) = defcell.Offset(cOffset, 1).Interior.Color
                'Special case for a colour rgb - use the value in the cell, but override the colour (not colour rgb)
                ElseIf Right(cKey, 11) = " colour rgb" Then
                    ChartDef(Left(cKey, Len(cKey) - 4)) = cVal
                'Special case for a font - get the attributes based upon the actual font in the cell
                ElseIf Right(cKey, 5) = " font" Then
                    ProcessFont cKey, defcell.Offset(cOffset, 1)
                Else
                    ChartDef(cKey) = cVal
                End If
                
                cOffset = cOffset + 1
            Loop
            
            Exit For
        End If
        
    Next
    
    If Not found Then Log LogLevelWarn, "Unable to find template " & TemplateName
End Sub

'This routine processes a series template, adding its values to the ChartDef dictionary
Private Sub ProcessSeriesTemplate(ByVal sKey As String, ByVal TemplateName As String)
    Dim wsTemplate As Worksheet                   'The _SeriesTemplates_ worksheet, which must be present
    Dim defcell As Range                          'Template cell
    Dim cOffset As Integer, cKey As String, cVal As String
    Dim found As Boolean
    Dim sID As Integer, sTemp As String
    
    Log LogLevelNote, "Looking for Series Template: " & TemplateName
    
    'Get the series ID by stripping away the Series and Template strings
    sTemp = Mid(sKey, 8, Len(sKey) - 9 - 7)
    If Not IsNumeric(sTemp) Then
        Log LogLevelWarn, "Invalid series ID value " & sTemp & " from key " & sKey
        Exit Sub
    End If
    sID = Val(sTemp)
    
    Set wsTemplate = Worksheets(SeriesTemplateSheet)
    found = False
    
    For Each defcell In wsTemplate.Range("1:1")
    
        If defcell = TemplateName Then
            Log LogLevelNote, "Found series template " & TemplateName & " at cell " & defcell.Address
            found = True
        
            'Process the key-value pairs under the template
            cOffset = 1
            
            Do Until defcell.Offset(cOffset) = ""
                cKey = KeyCheck("Series " & sID & " " & Trim(defcell.Offset(cOffset)))
                cVal = Trim(defcell.Offset(cOffset, 1))
        
                Log LogLevelNote, "Evaluating key " & cKey
                
                'If the key is 'Template', process the template first
                If cKey = "Template" Then
                    ProcessSeriesTemplate sKey, cVal
                'Special case for a color - get the colour index of the interior (background colour)
                'If the key is for a Point template, process it separately
                ElseIf Right(cKey, 15) = " Point Template" Then
                    ProcessPointTemplate cKey, cVal
                ElseIf Right(cKey, 7) = " colour" Then
                    ChartDef(cKey) = defcell.Offset(cOffset, 1).Interior.Color
                'Special case for a colour rgb - use the value in the cell, but override the colour (not colour rgb)
                ElseIf Right(cKey, 11) = " colour rgb" Then
                    ChartDef(Left(cKey, Len(cKey) - 4)) = cVal
                Else
                    ChartDef(cKey) = cVal
                End If
                
                cOffset = cOffset + 1
            Loop
            
            Exit For
        End If
        
    Next
    
    If Not found Then Log LogLevelWarn, "Unable to find template " & TemplateName
End Sub

'This routine processes a point template, adding its values to the ChartDef dictionary
Private Sub ProcessPointTemplate(ByVal sKey As String, ByVal TemplateName As String)
    Dim wsTemplate As Worksheet           'The _PointTemplates_ worksheet, which must be present
    Dim defcell As Range                  'Template cell
    Dim cOffset As Integer, rOffset As Integer, cKey As String, cVal As String
    Dim found As Boolean
    Dim pID As Integer, seriesPrefix As String
    
    Log LogLevelNote, "Processing point template for key: " & sKey
    Log LogLevelNote, "Looking for Point Template: " & TemplateName
    
    'sKey will include the series - get the series from it
    seriesPrefix = Left(sKey, InStr(8, sKey, " ", vbTextCompare))
    
    Set wsTemplate = Worksheets(PointTemplateSheet)
    found = False
    
    For Each defcell In wsTemplate.Range("1:1")
    
        If defcell = TemplateName Then
            Log LogLevelNote, "Found point template " & TemplateName & " at cell " & defcell.Address
            found = True
            
            'The row to process - start at 2 cells under the template name
            rOffset = 2
            
            Do Until defcell.Offset(rOffset) = ""
                'Process the key-value pairs under the template
                cOffset = 1
                
                'Get the point ID
                pID = defcell.Offset(rOffset)
                
                'The first key is assumed to be the point number
                Do Until defcell.Offset(1, cOffset) = ""
                    cKey = KeyCheck(seriesPrefix & "Point " & pID & " " & Trim(defcell.Offset(1, cOffset)))
                    cVal = Trim(defcell.Offset(rOffset, cOffset))
            
                    Log LogLevelNote, "Evaluating key " & cKey
                    
                    'Special case for a color - get the colour index of the interior (background colour)
                    If Right(cKey, 7) = " colour" Then
                        ChartDef(cKey) = defcell.Offset(rOffset, cOffset).Interior.Color
                    'Special case for a colour rgb - use the value in the cell, but override the colour (not colour rgb)
                    ElseIf Right(cKey, 11) = " colour rgb" Then
                        ChartDef(Left(cKey, Len(cKey) - 4)) = cVal
                    Else
                        ChartDef(cKey) = cVal
                    End If
                    
                    cOffset = cOffset + 1
                Loop 'End key iteration loop
                
                rOffset = rOffset + 1
                
            Loop 'End row iteration loop
            
            Exit For
        End If
        
    Next
End Sub

'This routine creates the chart based upon the provided chart definition dictionary
Private Sub CreateChart()
    Dim chtShape As Shape               'Shape object
    Dim chtObject As Chart              'Chart object
    Dim chtChartObject As ChartObject   'ChartObject object
    Dim cType As XlChartType            'Chart type
    
    Dim srcWorksheet As Worksheet       'Source worksheet
    Dim srcCells As Range               'Source cell range
    Dim srcString As String             'String representation of source
    
    Dim dstWorksheet As Worksheet       'Destination worksheet
    Dim dstCells As Range               'Destination cell range
    
    Dim seriesId As Integer             'Series ID counter
    Dim seriesPrefix As String          'Series string prefix
    Dim seriesObj As Series             'Series object
    
    Dim axisId As Integer               'Axis ID counter
    Dim axisPrefix As String            'Axis string prefix
    Dim axisObj As Axis                 'Axis object
    
    Dim pointId As Integer, pointId2 As Integer 'Point ID counters
    Dim pointCount As Integer           'Number of points in the series
    Dim pointPrefix As String           'Point string prefix
    Dim pointObj As Point                 'Point object
    
    Dim numVal As Double, boolVal As Boolean, strVal As String
    Dim i As Integer
    
On Error GoTo MissingRequirement
    'Get a full evaluation of the source and destination
    ParseSource ChartDef("source"), srcWorksheet, srcCells, srcString
    ParseDestination ChartDef("destination"), dstWorksheet, dstCells
On Error GoTo 0

On Error Resume Next
    'Get the ID of the chart, so we can delete an existing chart if any
    If ChartDef.Exists("id") Then
        If ChartDef("id") <> "" Then
            Log LogLevelNote, "Deleting any existing charts with ID " & ChartDef("id") & " on worksheet " & dstWorksheet.Name
            
            i = dstWorksheet.Shapes.Count
            Do Until i = 0
                If dstWorksheet.Shapes(i).Name = ChartDef("id") Then dstWorksheet.Shapes(i).Delete
                i = i - 1
            Loop
        End If
    End If
On Error GoTo 0

    'Establish the chart style enum
    cType = xlColumnClustered 'Default if not specified
    If GetString("Chart Type", strVal) Then cType = GetChartType(strVal)
        
    'Insert a chart based upon the source into the destination sheet
    Log LogLevelNote, "Creating new chart"
    
On Error GoTo AddChart2Failed
    Set chtShape = dstWorksheet.Shapes.AddChart2(-1, cType, dstCells.Left, dstCells.Top)
On Error GoTo 0
    
    'From this point forward, trap any errors but continue to the next line.
On Error GoTo ErrCode

    If ChartDef("id") <> "" Then chtShape.Name = ChartDef("id")
    Set chtObject = chtShape.Chart
    Set chtChartObject = chtObject.Parent
    chtObject.SetSourceData Range(srcString)
    
    'Set the attributes as they are set
    
    'Plot by rows or columns
    If GetString("plot by", strVal) Then
        Select Case strVal
            Case "rows":
                chtObject.PlotBy = xlRows
            Case "columns":
                chtObject.PlotBy = xlColumns
            Case Else:
                Log LogLevelWarn, "Invalid plot by value " & strVal
        End Select
    End If
    
    'Chart dimensions
    If GetNumber("width", numVal) Then chtShape.Width = numVal
    If GetNumber("height", numVal) Then chtShape.Height = numVal
    
    'Background Colour
    If GetNumber("background colour", numVal) Then chtShape.Fill.ForeColor.RGB = numVal
    
    'Background Transparency
    If GetNumber("background transparency", numVal) Then chtShape.Fill.Transparency = numVal
    
    'Background visible
    If GetBool("background visible", boolVal) Then chtShape.Fill.Visible = boolVal
    
    'Border colour
    If GetNumber("border colour", numVal) Then chtShape.Line.ForeColor.RGB = numVal
    
    'Border width
    If GetNumber("border width", numVal) Then chtShape.Line.Weight = numVal
    
    'Border transparency
    If GetNumber("border transparency", numVal) Then chtShape.Line.Transparency = numVal
    
    'Border visible
    If GetBool("border visible", boolVal) Then chtShape.Line.Visible = boolVal
    
    'Chart title
    If GetString("title text", strVal) Then
        If chtObject.HasTitle = False Then chtObject.SetElement msoElementChartTitleAboveChart
        
        chtObject.ChartTitle.Text = strVal
        
        'Other title attributes
        SetFont2Attribs "title text", chtObject.ChartTitle.Format.TextFrame2.TextRange.Font
    Else
        chtObject.SetElement msoElementChartTitleNone
    End If
    
    'Chart legend
    If GetString("legend", strVal) Then
        Select Case strVal
            Case "bottom":
                chtObject.SetElement msoElementLegendBottom
            Case "top":
                chtObject.SetElement msoElementLegendTop
            Case "left":
                chtObject.SetElement msoElementLegendLeft
            Case "right":
                chtObject.SetElement msoElementLegendRight
            Case "left overlay":
                chtObject.SetElement msoElementLegendLeftOverlay
            Case "right overlay":
                chtObject.SetElement msoElementLegendRightOverlay
            Case "none":
                chtObject.SetElement msoElementLegendNone
            Case Else:
                Log LogLevelWarn, "Unexpected legend position " & strVal
        End Select
        
        'Legend text attributes
        SetFont2Attribs "legend", chtObject.Legend.Format.TextFrame2.TextRange.Font
    Else
        chtObject.SetElement msoElementLegendNone
    End If
    
    'Series Gap Width
    If GetNumber("Series Gap Width", numVal) Then chtObject.ChartGroups(1).GapWidth = numVal
    
    'Series Overlap
    If GetNumber("Series Overlap", numVal) Then chtObject.ChartGroups(1).Overlap = numVal
    
    'Format the series
    For seriesId = 1 To chtObject.SeriesCollection.Count
        'Get the series object
        Set seriesObj = chtObject.SeriesCollection(seriesId)
        
        'Get the string prefix for the attributes
        seriesPrefix = "series " & seriesId & " "
        
        'Process the attributes, one by one
        
        'Chart type
        If GetString(seriesPrefix & "Chart Type", strVal) Then
            cType = GetChartType(strVal)
            seriesObj.ChartType = cType
        End If
        
        'Axis
        If GetString(seriesPrefix & "axis", strVal) Then
            Select Case strVal
                Case "1", "Primary", "Main":
                    seriesObj.AxisGroup = xlPrimary
                Case "2", "Secondary":
                    seriesObj.AxisGroup = xlSecondary
                Case Else:
                    Log LogLevelWarn, "Unexpected axis value " & strVal
            End Select
        End If
        
        'Fill Style
        If GetString(seriesPrefix & "Style", strVal) Then
            seriesObj.Format.Fill.Visible = msoTrue
            Select Case strVal
                Case "None":
                    seriesObj.Format.Fill.Visible = msoFalse
                Case "Solid":
                    seriesObj.Format.Fill.Solid
            End Select
        End If
        
        'Fill Colour
        If GetNumber(seriesPrefix & "Fill Colour", numVal) Then seriesObj.Format.Fill.ForeColor.RGB = numVal
        
        'Line Colour
        If GetNumber(seriesPrefix & "Line Colour", numVal) Then seriesObj.Format.Line.ForeColor.RGB = numVal
        
        'Fill Secondary Colour
        If GetNumber(seriesPrefix & "Fill Secondary Colour", numVal) Then seriesObj.Format.Fill.BackColor.RGB = numVal
        
        'Line Secondary Colour
        If GetNumber(seriesPrefix & "Line Secondary Colour", numVal) Then seriesObj.Format.Line.BackColor.RGB = numVal
        
        'Fill Transparency
        If GetNumber(seriesPrefix & "Fill Transparency", numVal) Then seriesObj.Format.Fill.Transparency = numVal
        
        'Line Transparency
        If GetNumber(seriesPrefix & "Line Transparency", numVal) Then seriesObj.Format.Line.Transparency = numVal
        
        'Line Width
        If GetNumber(seriesPrefix & "Line Width", numVal) Then seriesObj.Format.Line.Weight = numVal
        
        'Fill Visible
        If GetBool(seriesPrefix & "Fill Visible", boolVal) Then seriesObj.Format.Fill.Visible = boolVal
        
        'Line Visible
        If GetBool(seriesPrefix & "Line Visible", boolVal) Then seriesObj.Format.Line.Visible = boolVal
        
        'Data Labels
        Dim dlPosn As XlDataLabelPosition
        Dim dlCallout As Boolean
        Dim dlCalloutPct As Boolean
        Dim dlDataLabels As DataLabels
        
        dlCallout = False
        dlCalloutPct = False
        
        If GetString(seriesPrefix & "Data Labels", strVal) Then
            Select Case strVal
                Case "Callout"
                    dlPosn = xlLabelPositionOutsideEnd
                    dlCallout = True
                Case "callout percent", "callout %"
                    dlPosn = xlLabelPositionOutsideEnd
                    dlCallout = True
                    dlCalloutPct = True
                Case "Center"
                    dlPosn = xlLabelPositionCenter
                Case "Inside Base"
                    dlPosn = xlLabelPositionInsideBase
                Case "Inside End"
                    dlPosn = xlLabelPositionInsideEnd
                Case "Left"
                    dlPosn = xlLabelPositionLeft
                Case "Right"
                    dlPosn = xlLabelPositionRight
                Case "Outside End"
                    dlPosn = xlLabelPositionOutsideEnd
                Case Else
                    dlPosn = xlLabelPositionOutsideEnd
                    Log LogLevelWarn, "Unexpected Data Labels value " & strVal & ", using default of xlLabelPositionOutsideEnd"
            End Select
            
            seriesObj.HasDataLabels = True
            Set dlDataLabels = seriesObj.DataLabels
            
            'Data callout labels need to be done in several steps because it is harder to get the Chart.SetElement method to work on individual series
            If dlCallout Then
                dlDataLabels.Position = xlLabelPositionOutsideEnd
                dlDataLabels.Format.AutoShapeType = msoShapeRectangularCallout
                dlDataLabels.Format.Fill.ForeColor.RGB = 16777215
                dlDataLabels.Format.Line.Visible = msoTrue
                dlDataLabels.Format.Line.ForeColor.RGB = 12566463
                dlDataLabels.ShowCategoryName = True
                
                If dlCalloutPct Then
                    dlDataLabels.ShowPercentage = True
                    dlDataLabels.ShowValue = False
                End If
            Else
                dlDataLabels.Position = dlPosn
            End If
            
        End If
        
        'Delete the series
        If GetBool(seriesPrefix & "Delete", boolVal) Then
            If boolVal = True Then seriesObj.Delete
        End If
        
        'Format the points - positive point numbers
        pointCount = seriesObj.Points.Count
        
        'Do positive and then negative point numbers
        For i = 1 To 2
        
            For pointId = 1 To pointCount
                
                If i = 1 Then
                    pointId2 = pointId
                    Set pointObj = seriesObj.Points(pointId2)
                Else
                    pointId2 = -pointId
                    Set pointObj = seriesObj.Points(pointCount - pointId + 1)
                End If
                
                pointPrefix = seriesPrefix & "Point " & pointId2 & " "
                
                'Fill visible
                If GetBool(pointPrefix & "Fill visible", boolVal) Then pointObj.Format.Fill.Visible = boolVal
                
                'Line visible
                If GetBool(pointPrefix & "Line visible", boolVal) Then pointObj.Format.Line.Visible = boolVal
                
                'Fill colour
                If GetNumber(pointPrefix & "Fill Colour", numVal) Then pointObj.Format.Fill.ForeColor.RGB = numVal
                
                'Line colour
                If GetNumber(pointPrefix & "Line Colour", numVal) Then pointObj.Format.Line.ForeColor.RGB = numVal
                
                'Fill transparency
                If GetNumber(pointPrefix & "Fill transparency", numVal) Then pointObj.Format.Fill.Transparency = numVal
                
                'Line transparency
                If GetNumber(pointPrefix & "Line transparency", numVal) Then pointObj.Format.Line.Transparency = numVal
                
                'Line width
                If GetNumber(pointPrefix & "Line width", numVal) Then pointObj.Format.Line.Weight = numVal
                
            Next pointId
            
        Next i
        
    Next
    
    
    'Axis attributes
    Dim cAxis
    For axisId = 0 To 2
        If axisId = 0 Then
            axisPrefix = "X "
            Set axisObj = chtObject.Axes(xlCategory)
        ElseIf axisId = 1 Then
            axisPrefix = "Y1 "
            Set axisObj = chtObject.Axes(xlValue)
        ElseIf axisId = 2 Then
            axisPrefix = "Y2 "
            Set axisObj = chtObject.Axes(xlValue, xlSecondary)
        End If
        
        'Tick label font
        SetFontAttribs axisPrefix & "label", axisObj.TickLabels.Font
        
        'Minimum and maximum values
        If GetNumber(axisPrefix & "min", numVal) Then axisObj.MinimumScale = numVal
        If GetNumber(axisPrefix & "max", numVal) Then axisObj.MaximumScale = numVal
        
        'Minor and major unit
        If GetNumber(axisPrefix & "minor unit", numVal) Then axisObj.MinorUnit = numVal
        If GetNumber(axisPrefix & "major unit", numVal) Then axisObj.MajorUnit = numVal
        
        'X-axis crosses (at), only for Y1/Y2
        If axisId <> 0 Then
            If GetString(axisPrefix & "x crosses", strVal) Then
                Select Case strVal
                    Case "Automatic":
                        axisObj.Crosses = xlAxisCrossesAutomatic
                    Case "Custom":
                        axisObj.Crosses = xlAxisCrossesCustom
                    Case "Minimum":
                        axisObj.Crosses = xlAxisCrossesMinimum
                    Case "Maximum":
                        axisObj.Crosses = xlAxisCrossesMaximum
                    Case Else:
                        Log LogLevelWarn, "Unexpected axis crosses value " & strVal
                End Select
            End If
            
            If GetNumber(axisPrefix & "x crosses at", numVal) Then axisObj.CrossesAt = numVal
        End If
    Next axisId
    
    'Pie explosion
    If GetNumber("Pie explosion", numVal) Then chtObject.FullSeriesCollection(1).Explosion = numVal
    
    'Pie first slice angle
    If GetNumber("Pie first slice angle", numVal) Then chtObject.ChartGroups(1).FirstSliceAngle = numVal
        
ExitCode:
    Log LogLevelNote, "Finished creating chart"
    
    Exit Sub
    
AddChart2Failed:
    Log LogLevelError, "Unable to add new chart due to error: " & Err.Description & " Aborting..."
    Resume ExitCode
    
ErrCode:
    Log LogLevelError, "Unexpected error whilst creating chart: " & Err.Description & " Continuing..."
    Resume Next
    
MissingRequirement:
    Log LogLevelError, "A required attribute is missing or invalid, aborting..."
    Resume ExitCode
End Sub

'Replaces words in a key with alternatives, e.g. color = colour, to simplify code
Private Function KeyCheck(ByVal key As String) As String
    KeyCheck = Replace(key, "color", "colour", compare:=vbTextCompare)
End Function

'Check if a numeric value for the key exists, and if so, return to the out variable
Private Function GetNumber(ByVal key As String, ByRef out As Double) As Boolean
    Log LogLevelNote, "Checking for key " & key
    
    key = KeyCheck(key)
    If ChartDef.Exists(key) = False Then
        GetNumber = False
    ElseIf IsNumeric(ChartDef(key)) Then
        out = Val(ChartDef(key))
        GetNumber = True
    Else
        Log LogLevelWarn, "Expected numeric value but got " & ChartDef(key)
        GetNumber = False
    End If
    
    If GetNumber Then Log LogLevelNote, "Found numeric value " & out
End Function

'Check if a string value for the key exists, and if so, return to the out variable
Private Function GetString(ByVal key As String, ByRef out As String) As Boolean
    Log LogLevelNote, "Checking for key " & key
    
    key = KeyCheck(key)
    If ChartDef.Exists(key) = False Then
        GetString = False
    Else
        GetString = True
        out = ChartDef(key)
    End If
    
    If GetString Then Log LogLevelNote, "Found string value " & out
End Function

'Check if a boolean value for the key exists, and if so, return to the out variable
Private Function GetBool(ByVal key As String, ByRef out As Boolean) As Boolean
    Log LogLevelNote, "Checking for key " & key
    
    key = KeyCheck(key)
    If ChartDef.Exists(key) = False Then
        GetBool = False
    ElseIf ChartDef(key) = "True" Or ChartDef(key) = "Yes" Then
        out = True
        GetBool = True
    ElseIf ChartDef(key) = "False" Or ChartDef(key) = "No" Then
        out = False
        GetBool = True
    ElseIf IsNumeric(ChartDef(key)) Then
        If Val(ChartDef(key)) = 0 Then
            out = False
        Else
            out = True
        End If
        GetBool = True
    Else
        Log LogLevelWarn, "Expected boolean value but got " & ChartDef(key)
        GetBool = False
    End If
    
    If GetBool Then Log LogLevelNote, "Found boolean value " & out
End Function

Private Sub ProcessFont(ByVal KeyPrefix As String, ByVal refCell As Range)
    Log LogLevelNote, "Processing font for key prefix " & KeyPrefix & ", cell " & refCell.Address
    
    Dim s As String
    
    ChartDef(KeyPrefix & " Name") = refCell.Font.Name
    ChartDef(KeyPrefix & " Size") = refCell.Font.Size
    ChartDef(KeyPrefix & " Bold") = refCell.Font.Bold
    ChartDef(KeyPrefix & " Italic") = refCell.Font.Italic
    
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
    ChartDef(KeyPrefix & " Underline") = s
    
    ChartDef(KeyPrefix & " Colour") = refCell.Font.Color
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

'Get Chart Type enum from the string provided
Private Function GetChartType(ByVal ChartString As String) As XlChartType
    Log LogLevelNote, "Getting chart type for " & ChartString
    
    Dim cType As XlChartType
    
    Select Case ChartString
        Case "Bar":
            cType = xlBarClustered
        Case "Bar stacked":
            cType = xlBarStacked
        Case "Bar stacked 100":
            cType = xlBarStacked100
        Case "3d Bar":
            cType = xl3DBarClustered
        Case "3d Bar stacked":
            cType = xl3DBarStacked
        Case "3d Bar stacked 100":
            cType = xl3DBarStacked100
        Case "Column":
            cType = xlColumnClustered
        Case "Column stacked":
            cType = xlColumnStacked
        Case "Column stacked 100":
            cType = xlColumnStacked100
        Case "3d Column":
            cType = xl3DColumnClustered
        Case "3d Column stacked":
            cType = xl3DColumnStacked
        Case "3d Column stacked 100":
            cType = xl3DColumnStacked100
        Case "Line":
            cType = xlLine
        Case "Line stacked":
            cType = xlLineStacked
        Case "Line stacked 100":
            cType = xlLineStacked100
        Case "3d Line":
            cType = xl3DLine
        Case "Pie"
            cType = xlPie
        Case "3d pie"
            cType = xl3DPie
        Case "Doughnut", "Donut"
            cType = xlDoughnut
        Case Else:
            cType = xlColumnClustered
            Log LogLevelWarn, "Unsupported chart type " & ChartString & ", defaulting to Column"
    End Select
    
    GetChartType = cType
End Function

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
    
    Log LogLevelNote, "Final destination cell " & outCell.Address
End Sub

'Helper function to get the RGB value of the cell
Public Function GetRGB(ByVal ref As String) As LongLong
    Dim r As Range
    Set r = Range(ref).Range("A1")
    GetRGB = r.Interior.Color
End Function

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
    wsLog.Range("A1") = "Chart ID / Source"
    wsLog.Range("B1") = "Type"
    wsLog.Range("C1") = "Message"
    
    LogLine = 1
End Sub

'Add a log entry
Private Sub Log(ByVal level As LogLevel, ByVal message As String)
    If LogEnabled = False Then Exit Sub
    
    If Not ChartDef Is Nothing Then
        If ChartDef.Exists("ID") Then
            logChart = ChartDef("ID")
        ElseIf ChartDef.Exists("Source") Then
            logChart = "Source: " & ChartDef("Source")
        ElseIf ChartDef.Exists("Chart Definition Source") Then
            logChart = "Chart Definition Source: " & ChartDef("Chart Definition Source")
        Else
            logChart = "(Unknown chart)"
        End If
    Else
        logChart = ""
    End If
    
    wsLog.Range("A1").Offset(LogLine) = logChart
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
