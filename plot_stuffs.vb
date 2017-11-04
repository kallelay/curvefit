Imports OxyPlot
Imports OxyPlot.Series
Imports OxyPlot.Axes

Module plot_stuffs
    Public CR! = 2
    Public PLOT_WINDOW_X% = 0, PLOT_WINDOW_Y% = 0
    Public selected_idx_x = 0, selected_idx_y = 0, selected_idx_z = 0
    Public chosenMode = -1


    Public a#, b#, c#, d#, eoff# 'slopes
    Public curpts#()
    Public polyCoef#() = {0.0}
    Public G_READY As Bitmap = Nothing

    Dim pnt() As Single

    Public SHOW_MEAS As Boolean = True
    Public SHOW_ESTIM As Boolean = True
    Public AUTO_SCALE As Boolean = True
    Public LEGEND_VISIBLE As Boolean = True

    Public PLOT_FONT As New Font((New OxyPlot.PlotModel).DefaultFont, (New OxyPlot.PlotModel).DefaultFontSize)
    
    'Public PLOT_FONT As New Font("Segoe UI Light", 12)

    Public XAXIS_TITLE = ""
    Public YAXIS_TITLE = ""
    Public MEAS_TITLE = ""
    Public ESTIM_TITLE = ""
    Public TITLE = ""
    Public LegendPosition As LegendPosition = OxyPlot.LegendPosition.TopLeft



    Public plotModel As PlotModel = Nothing
    Sub makeGraphics()
        Try
            'Dim img As Image = New Image()

            If list_x Is Nothing Then Exit Sub
            If list_x.Count = 0 Then Exit Sub

            'Function (imported)
            Dim lineSeries As New LineSeries()
            lineSeries.MarkerFill = OxyColors.DimGray
            lineSeries.StrokeThickness = 0.1

            lineSeries.Color = OxyColors.LightGray
            lineSeries.MarkerSize = 2
            lineSeries.MarkerType = MarkerType.Circle
            lineSeries.Points.AddRange(list_pts)
            ' lineSeries.Font = PLOT_FONT.Name
            'lineSeries.FontSize = PLOT_FONT.Size
            '  lineSeries.FontWeight = PLOT_FONT.Bold
            lineSeries.Title = MEAS_TITLE

            plotModel.Series.Clear()
            ' If plotModel.Series.Count = 0 Then
            If SHOW_MEAS Then plotModel.Series.Add(lineSeries)
            '  Else

            '  End If 

            Dim predSeries As New LineSeries()
            predSeries.Color = OxyColors.CadetBlue
            predSeries.Points.AddRange(list_pred_pts)
            predSeries.Title = ESTIM_TITLE
            predSeries.Font = PLOT_FONT.Name
            predSeries.FontSize = PLOT_FONT.Size
            If SHOW_ESTIM Then plotModel.Series.Add(predSeries)

            'plotModel.Axes(0) = plotModel.DefaultXAxis
            '  plotModel.Axes(1) = plotModel.DefaultYAxis

            Dim Z_TP = 0.75
            Dim Z_LT = 2 - Z_TP


            If AUTO_SCALE Then

                plotModel.Axes(0) = New LinearAxis() With {.Position = AxisPosition.Bottom, .MaximumPadding = 0.1, .MinimumPadding = 0.1} ', .Minimum = list_x.Min, .Maximum = list_x.Max}
                plotModel.Axes(1) = New LinearAxis() With {.Position = AxisPosition.Left, .MaximumPadding = 0.1, .MinimumPadding = 0.1} ', .Minimum = list_y.Min * Z_LT, .Maximum = list_y.Max * Z_TP}

                plotModel.Axes(0).ZoomAt(list_x.Min * Z_LT, list_x.Max * Z_TP)
                plotModel.Axes(1).ZoomAt(list_y.Min * Z_LT, list_y.Max * Z_TP)
            End If


            plotModel.Title = TITLE
            plotModel.Axes(0).Title = XAXIS_TITLE

            plotModel.Axes(1).Title = YAXIS_TITLE

            plotModel.LegendFont = PLOT_FONT.Name
            plotModel.LegendFontSize = PLOT_FONT.Size

            plotModel.LegendPlacement = LegendPlacement.Inside

            plotModel.LegendPosition = LegendPosition.TopLeft


            plotModel.InvalidatePlot(True)
            '      plotModel.ResetAllAxes()



            '-------------------------- old code
            Exit Sub
            Dim bmp As New Bitmap(PLOT_WINDOW_X, PLOT_WINDOW_Y)
            Dim g As Graphics = Graphics.FromImage(bmp)
            g.Clear(Color.White)


            For i = 0 To list_x.Count - 1
                pnt = convertCoordinatesToPlot(list_x(i), list_y(i))
                g.FillEllipse(Brushes.Black, pnt(0) - CR, pnt(1) - CR, CR * 2, CR * 2)
            Next

            ReDim ppts(lines.Count - 1)
            Dim pp(1) As Single
            For i = 0 To lines.Count - 1
                pp = convertCoordinatesToPlot(lines(i).X, lines(i).Y)
                ppts(i) = New PointF(pp(0), pp(1))
            Next

            If lines IsNot Nothing Then
                g.DrawLines(Pens.RoyalBlue, ppts)
            End If

            g.Save()
            'g.Flush()

            G_READY = bmp.Clone
        Catch
        End Try


    End Sub

    Dim sX, sY, oX, oY As Double
    Dim lastpx%, lastpy%
    Public Function convertCoordinatesToPlot(ByVal x#, ByVal y#) As Single()

        ' If lastpx <> PLOT_WINDOW_X Or lastpy <> PLOT_WINDOW_Y Then

        sX = PLOT_WINDOW_X * 0.98 / (v_maxs(selected_idx_x) - v_mins(selected_idx_x))
        sY = PLOT_WINDOW_Y * 0.98 / (v_maxs(selected_idx_y) - v_mins(selected_idx_y))

        oX = v_mins(selected_idx_x) * 1.02
        oY = v_mins(selected_idx_y) * 1.02
        '  End If

        Return New Single() {oX * sX + sX * (x - oX), PLOT_WINDOW_Y - oY * sY + sY * (y - oY)}

    End Function

End Module
