Imports MICRO_USB2_CLR_DLL
'Imports MICRO_USB2_CLR
Imports System.Threading
Imports System.IO
Imports System.Windows.Forms.DataVisualization.Charting
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Linq
Imports System.Runtime.InteropServices


Public Class Form1
    Dim g_lDevID As Long
    Dim obuMICROUSB2 As New MICRO_USB2_CLR()
    Public Const EXPOSURE_TIME As Integer = 20000
    Public Const PIXEL_SIZE As Integer = 288
    Dim captureThread As Thread
    Dim capturedData(PIXEL_SIZE - 1) As UShort
    ' Private capturedData As New List(Of UShort)
    Private graphSeries As New Series("Intensity")
    Dim isCapturing As Boolean = False
    'Dim capturedData() As UShort '

    ' ✅ Define Coefficients for Wavelength Calculation
    'Dim A0 As Double = 304.349493963413
    'Dim B1 As Double = 2.68852555790766
    'Dim B2 As Double = -0.000855332088537141
    'Dim B3 As Double = -0.0000094620758458323
    'Dim B4 As Double = 0.0000000138925108579558
    'Dim B5 As Double = 0.000000000000860288677236112

    Public A0 As Double
    Public B1 As Double
    Public B2 As Double
    Public B3 As Double
    Public B4 As Double
    Public B5 As Double
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize Chart Series
        Chart1.Series.Clear()
        graphSeries.ChartType = SeriesChartType.Line
        Chart1.Series.Add(graphSeries)

        ' Configure Zooming
        ConfigureChartZooming()

        ' Disable Stop & Save buttons initially
        btnStop.Enabled = False
        btnSaveCSV.Enabled = False

    End Sub

    ''' =========================== 🚀 START CAPTURE =========================== '''

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        If isCapturing Then
            MessageBox.Show("Capturing is already in progress!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        isCapturing = True
        btnStart.Enabled = False ' Start button disable during capture
        btnStop.Enabled = True
        btnSaveCSV.Enabled = False ' Save should only work after stopping
        lblStatus.Text = "Status: Capturing..."

        graphSeries.Points.Clear() ' Clear graph data

        captureThread = New Thread(AddressOf CaptureData)
        captureThread.IsBackground = True
        captureThread.Start()
    End Sub



    Private Sub CaptureDataLoop()
        Try
            While isCapturing
                Dim lReturn As Long
                Dim aryusImageData(PIXEL_SIZE - 1) As UShort
                Dim arylTime(0) As ULong

                ' Initialize device
                lReturn = obuMICROUSB2.MICRO_USB2_initialize()
                If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then Exit While

                ' Get device list
                Dim aryDeviceId(8) As Short
                Dim usNum As UShort
                lReturn = obuMICROUSB2.MICRO_USB2_getModuleConnectionList(aryDeviceId, usNum)
                If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then Exit While

                ' Set Device ID
                g_lDevID = CLng(aryDeviceId(0))

                ' Open device
                lReturn = obuMICROUSB2.MICRO_USB2_open(g_lDevID)
                If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then Exit While

                ' Set Exposure Time
                lReturn = obuMICROUSB2.MICRO_USB2_setExposureTime(g_lDevID, EXPOSURE_TIME)
                If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then Exit While

                ' Start Capture
                lReturn = obuMICROUSB2.MICRO_USB2_captureStart(g_lDevID)
                If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then Exit While

                ' Wait for data
                Thread.Sleep(100)

                ' Get Data
                lReturn = obuMICROUSB2.MICRO_USB2_getImageData(g_lDevID, aryusImageData, arylTime)
                If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then Exit While

                ' Stop Capture
                obuMICROUSB2.MICRO_USB2_captureStop(g_lDevID)

                ' Save Data to Global Variable
                capturedData = aryusImageData

                ' Update Graph on UI Thread
                Me.Invoke(New MethodInvoker(Sub() DisplayGraph(capturedData.ToArray())))

                ' Close Device
                obuMICROUSB2.MICRO_USB2_close(g_lDevID)
                obuMICROUSB2.MICRO_USB2_uninitialize()

                ' Delay Before Next Capture
                Thread.Sleep(100) ' 1 second delay

                ' Check Stop Condition
                If Not isCapturing Then
                    Exit While
                End If
            End While
        Catch ex As Exception
            MessageBox.Show("Error in capturing process: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Ensure cleanup and enable buttons after stopping
            Me.Invoke(New MethodInvoker(Sub()
                                            lblStatus.Text = "Status: Stopped"
                                            btnStart.Enabled = True
                                            btnStop.Enabled = False
                                            btnSaveCSV.Enabled = True
                                        End Sub))
        End Try
    End Sub

    ''' =========================== 🛑 STOP CAPTURE =========================== '''
    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        If Not isCapturing Then
            MessageBox.Show("No capturing is in progress!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        isCapturing = False ' Stop capturing

        ' ✅ Thread Stop
        If captureThread IsNot Nothing AndAlso captureThread.IsAlive Then
            captureThread.Join(1000)
        End If

        btnStart.Enabled = True
        btnStop.Enabled = False
        btnSaveCSV.Enabled = True ' ✅ Enable Save now
        lblStatus.Text = "Status: Stopped"

        MessageBox.Show("Capture stopped successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub




    ''' =========================== 📡 CAPTURE FUNCTION =========================== '''
    Private Sub CaptureData()
        Dim aryusImageData(PIXEL_SIZE - 1) As UShort
        Dim arylTime(0) As ULong

        ' Initialize spectrometer
        Dim lReturn As Long = obuMICROUSB2.MICRO_USB2_initialize()
        If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then
            MessageBox.Show("Initialization failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim aryDeviceId(8) As Short
        Dim usNum As UShort
        'Dim dataList As List(Of UShort) = aryusImageData.ToList()

        lReturn = obuMICROUSB2.MICRO_USB2_getModuleConnectionList(aryDeviceId, usNum)
        If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then
            MessageBox.Show("No device found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        g_lDevID = CLng(aryDeviceId(0))
        lReturn = obuMICROUSB2.MICRO_USB2_open(g_lDevID)
        If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then
            MessageBox.Show("Device open failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        '  Get Calibration Coefficients (AFTER DEVICE OPEN)
        Dim coeffs(6) As Double ' Assuming 6 calibration coefficients
        Dim result As Long = obuMICROUSB2.MICRO_USB2_getCalibrationCoefficient(g_lDevID, 0, coeffs)

        If result <> Usb2Struct.Cusb2Err.usb2Success Then
            MessageBox.Show("Calibration coefficient retrieval failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If result = Usb2Struct.Cusb2Err.usb2Success Then
            A0 = coeffs(0)
            B1 = coeffs(1)
            B2 = coeffs(2)
            B3 = coeffs(3)
            B4 = coeffs(4)
            B5 = coeffs(5)
        Else
            MessageBox.Show("Calibration coefficient retrieval failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If


        lReturn = obuMICROUSB2.MICRO_USB2_setExposureTime(g_lDevID, EXPOSURE_TIME)
        If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then
            MessageBox.Show("Exposure time setting failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' Continuous data collection while capturing
        While isCapturing
            lReturn = obuMICROUSB2.MICRO_USB2_captureStart(g_lDevID)
            If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then
                MessageBox.Show("Capture start failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Thread.Sleep(500) ' Delay for capture
            lReturn = obuMICROUSB2.MICRO_USB2_getImageData(g_lDevID, aryusImageData, arylTime)
            If lReturn <> Usb2Struct.Cusb2Err.usb2Success Then
                MessageBox.Show("Data retrieval failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ' Store captured data
            capturedData = aryusImageData

            Me.Invoke(Sub() DisplayGraph(capturedData))

            obuMICROUSB2.MICRO_USB2_captureStop(g_lDevID)
            Thread.Sleep(100)

            ' Update graph in UI thread
            Me.Invoke(Sub()
                          graphSeries.Points.Clear()
                          For Each intensity As UShort In capturedData
                              graphSeries.Points.AddY(intensity)
                          Next
                      End Sub)
        End While

        ' Stop and cleanup
        obuMICROUSB2.MICRO_USB2_captureStop(g_lDevID)
        obuMICROUSB2.MICRO_USB2_close(g_lDevID)
        obuMICROUSB2.MICRO_USB2_uninitialize()
    End Sub


    ''' =========================== 📈 DISPLAY GRAPH WITH LABELS =========================== '''
    Private Sub DisplayGraph(data() As UShort)
        If Chart1.InvokeRequired Then
            Chart1.Invoke(New MethodInvoker(Sub() DisplayGraph(data)))
        Else
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()

            ' Create a new chart area
            Dim chartArea As New ChartArea()
            Chart1.ChartAreas.Add(chartArea)

            ' Customize chart labels
            Chart1.Titles.Clear()
            Chart1.Titles.Add("Spectrometer Data") ' ✅ Add Chart Title
            Chart1.Titles(0).Font = New Font("Arial", 14, FontStyle.Bold)

            ' Set X and Y axis labels
            chartArea.AxisX.Title = "Wavelength (nm)" ' ✅ X-Axis Label
            chartArea.AxisX.TitleFont = New Font("Arial", 12, FontStyle.Bold)
            chartArea.AxisX.LabelStyle.Enabled = True ' ✅ Enable Labels
            chartArea.AxisX.MajorGrid.Enabled = True  ' ✅ Enable Grid Lines
            chartArea.AxisX.IsMarginVisible = False  ' ✅ Remove Extra Margins

            chartArea.AxisY.Title = "Intensity" ' ✅ Y-Axis Label
            chartArea.AxisY.TitleFont = New Font("Arial", 12, FontStyle.Bold)

            ' ✅ Enable Auto Scale Mode (Removes Fixed Min, Max, Interval)
            chartArea.AxisX.Minimum = Double.NaN
            chartArea.AxisX.Maximum = Double.NaN
            chartArea.AxisX.Interval = Double.NaN

            ' ✅ Enable Zooming
            chartArea.AxisX.ScaleView.Zoomable = True
            chartArea.AxisY.ScaleView.Zoomable = True
            chartArea.CursorX.IsUserEnabled = True
            chartArea.CursorX.IsUserSelectionEnabled = True
            chartArea.CursorY.IsUserEnabled = True
            chartArea.CursorY.IsUserSelectionEnabled = True

            ' Create and configure the series
            Dim series As New Series("Spectrometer Data") With {
            .ChartType = SeriesChartType.Line,
            .BorderWidth = 1,
            .Color = Color.Blue
        }

            ' Add data points with labels
            For i As Integer = 0 To data.Length - 1
                Dim wavelength As Integer = CalculateWavelength(i + 1) ' Pixel starts from 1
                Dim point As New DataPoint(wavelength, data(i))
                series.Points.Add(point)
            Next

            ' Add series to the chart
            Chart1.Series.Add(series)
        End If
    End Sub

    '----------------------- Graph functionality -----------------------------------'
    Private Sub GraphForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigureChartZooming()
    End Sub

    Private Sub ConfigureChartZooming()
        Dim chartArea As ChartArea = Chart1.ChartAreas(0)
        With Chart1.ChartAreas(0)
            .AxisX.ScaleView.Zoomable = True  ' X-Axis Zoom Enable
            .AxisY.ScaleView.Zoomable = True  ' Y-Axis Zoom Enable
            .CursorX.IsUserEnabled = True
            .CursorX.IsUserSelectionEnabled = True
            .CursorY.IsUserEnabled = True
            .CursorY.IsUserSelectionEnabled = True
        End With
    End Sub

    Private Sub Chart1_MouseWheel(sender As Object, e As MouseEventArgs) Handles Chart1.MouseWheel
        Dim ca As ChartArea = Chart1.ChartAreas(0)
        Dim zoomFactor As Double = 0.9
        If e.Delta > 0 Then ' Scroll Up - Zoom In
            ca.AxisX.ScaleView.Zoom(ca.AxisX.ScaleView.Position, ca.AxisX.ScaleView.Size * zoomFactor)
            ca.AxisY.ScaleView.Zoom(ca.AxisY.ScaleView.Position, ca.AxisY.ScaleView.Size * zoomFactor)
        ElseIf e.Delta < 0 Then ' Scroll Down - Zoom Out
            ca.AxisX.ScaleView.Zoom(ca.AxisX.ScaleView.Position, ca.AxisX.ScaleView.Size / zoomFactor)
            ca.AxisY.ScaleView.Zoom(ca.AxisY.ScaleView.Position, ca.AxisY.ScaleView.Size / zoomFactor)
        End If

        ' ✅ Dynamically Adjust Labels After Zooming
        AdjustAxisLabels()
    End Sub

    Private Sub Chart1_DoubleClick(sender As Object, e As EventArgs) Handles Chart1.DoubleClick
        With Chart1.ChartAreas(0)
            .AxisX.ScaleView.ZoomReset()
            .AxisY.ScaleView.ZoomReset()
        End With

        ' ✅ Reset Axis Label Scaling
        AdjustAxisLabels()
    End Sub

    Private Sub AdjustAxisLabels()
        Dim ca As ChartArea = Chart1.ChartAreas(0)
        Dim minX As Double = ca.AxisX.ScaleView.ViewMinimum
        Dim maxX As Double = ca.AxisX.ScaleView.ViewMaximum

        ' ✅ Automatically Adjust Interval Based on Zoom Level
        Dim range As Double = maxX - minX
        If range > 200 Then
            ca.AxisX.Interval = 50
        ElseIf range > 100 Then
            ca.AxisX.Interval = 20
        ElseIf range > 50 Then
            ca.AxisX.Interval = 10
        ElseIf range > 20 Then
            ca.AxisX.Interval = 5
        Else
            ca.AxisX.Interval = 1 ' ✅ Maximum Detail at High Zoom
        End If
    End Sub


    ''' =========================== 🔢 CALCULATE WAVELENGTH =========================== '''
    Private Function CalculateWavelength(i As Integer) As Integer
        Return CInt(Math.Round(A0 + (B1 * i) + (B2 * i ^ 2) + (B3 * i ^ 3) + (B4 * i ^ 4) + (B5 * i ^ 5)))
    End Function



    ''' =========================== 💾 SAVE TO CSV =========================== '''

    Private Sub btnSaveCSV_Click(sender As Object, e As EventArgs) Handles btnSaveCSV.Click
        If isCapturing Then
            MessageBox.Show("Stop capturing before saving!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' ✅ Ensure the user has entered a data type description
        Dim dataType As String = txtDataType.Text.Trim()
        If dataType = "" Then
            MessageBox.Show("Please enter the type of data being read before saving!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' ✅ Ensure captured data is available
        If capturedData.Count = 0 Then
            MessageBox.Show("No data available to save!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' ✅ Set File Path
        Dim savePath As String = "D:\Visual Studio\Spectrometre\Spectrometer_Data.csv"
        Dim fileExists As Boolean = File.Exists(savePath)

        Try
            ' ✅ Calculate Wavelengths (Ensure Pixel Starts from 1)
            Dim wavelengths As New List(Of Integer)
            For pix As Integer = 1 To 288 ' 🔹 Pixel 1 to 288
                wavelengths.Add(CalculateWavelength(pix))
            Next

            ' ✅ Prepare Wavelength Headers
            Dim wavelengthHeader As String = "wavelengths," & String.Join(",", wavelengths)

            ' ✅ Ensure Captured Data Matches Expected Length
            If capturedData.Count <> 288 Then
                MessageBox.Show("Captured data count mismatch! Expected 288 values but got " & capturedData.Count, "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ' ✅ Append Data to CSV
            Using writer As New StreamWriter(savePath, True)
                ' ✅ Write Header Only If File is New
                If Not fileExists Then
                    writer.WriteLine(wavelengthHeader)
                End If
                ' ✅ Append Data Row
                writer.WriteLine(dataType & "," & String.Join(",", capturedData))
            End Using

            MessageBox.Show("✅ Data saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As IOException
            MessageBox.Show("Error: The file is in use. Please close it and try again.", "File Access Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            MessageBox.Show("Error saving file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

End Class
