Imports System.Text.RegularExpressions
Imports System.Globalization
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1

    Public path
    Public destination


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ofd_DestinationExcel.FileName = Label2.Text
        If ofd_DestinationExcel.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Label2.Text = ofd_DestinationExcel.FileName
            destination = ofd_DestinationExcel.FileName
        End If
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub Label26_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label27_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label43_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label47_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MsgBox("Help Box is not avaiable yet. Please Contact Dan Gibbons (631) 456-7733 or dgibbo03@nyit.edu for further information or help.")
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim fileReader As String = My.Computer.FileSystem.ReadAllText(path)
        Console.WriteLine("Text File Imported!")

        Dim stringPattern_Date As New Regex("Study Date and Time\s*\=\s*(?<dateResult>\w{3} \d{2}\, \d{4}).*", RegexOptions.IgnoreCase)
        Dim stringPattern_Filename As New Regex("Filename Prefix\s*\=\s*(?<fileResult>.*)", RegexOptions.IgnoreCase)
        Dim stringPattern_Voltage As New Regex("Source Voltage \(kV\)\s*\=\s*(?<voltageResult>\d+)", RegexOptions.IgnoreCase)
        Dim stringPattern_Current As New Regex("Source Current \(uA\)\s*\=\s*(?<currentResult>\d+)", RegexOptions.IgnoreCase)
        Dim stringPattern_Resolution As New Regex("Number of Rows\s*\=\s*(?<resolutionResult>\d+)", RegexOptions.IgnoreCase)
        Dim stringPattern_Exposure As New Regex("Exposure \(ms\)\s*\=\s*(?<exposureResult>\d+)", RegexOptions.IgnoreCase)
        Dim stringPattern_PixelSize As New Regex("Pixel Size\s*\(um\)\s*=(?<pixelsizeResult>\d+\.\d+)", RegexOptions.IgnoreCase)
        Dim stringPattern_ImageFormat As New Regex("Image Format\s*\=\s*(?<imageformatResult>\w+)", RegexOptions.IgnoreCase)
        Dim stringPattern_RotationStep As New Regex("Rotation Step\s*\(deg\)\s*\=\s*(?<rotationstepResult>\d+\.\d+)", RegexOptions.IgnoreCase)
        Dim stringPattern_Frames As New Regex("Frame Averaging\s*\=[ON|OFF]{2,3}\s*\((?<framesResult>\d+)\)", RegexOptions.IgnoreCase)
        Dim stringPattern_RandomMovement As New Regex("Random Movement\s*\=[ON|OFF]{2,3}\s*\((?<randommovementResult>\d+)\)", RegexOptions.IgnoreCase)
        Dim stringPattern_360 As New Regex("Use 360 Rotation\s*\=\s*(?<result360>\w+)", RegexOptions.IgnoreCase)
        Dim stringPattern_ScanDuration As New Regex("Scan duration\s*\=\s*(?<scandurationResult>\d{2}\:\d{2}\:\d{2})", RegexOptions.IgnoreCase)

        Dim output_Date As String
        Dim output_Filename As String
        Dim output_Voltage As String
        Dim output_Current As String
        Dim output_Resolution As String
        Dim output_Exposure As String
        Dim output_PixelSize As String
        Dim output_ImageFormat As String
        Dim output_RotationStep As String
        Dim output_Frames As String
        Dim output_RandomMovement As String
        Dim output_360 As String
        Dim output_ScanDuration As String

        Dim matchesDate As MatchCollection = stringPattern_Date.Matches(fileReader)
        Dim outDate As Date
        Dim provider As CultureInfo = CultureInfo.InvariantCulture
        If matchesDate.Count > 0 Then
            For Each match As Match In matchesDate
                output_Date = match.Groups.Item("dateResult").Value
                'output_Date.ToString()

                Try
                    outDate = Date.ParseExact(output_Date.ToString, "MMM dd, yyyy", provider)
                    TextBox3.Text = outDate
                Catch b As FormatException
                    TextBox3.Text = output_Date
                End Try
            Next
        End If

        Dim matchesFilename As MatchCollection = stringPattern_Filename.Matches(fileReader)
        If matchesFilename.Count > 0 Then
            For Each match As Match In matchesFilename
                output_Filename = match.Groups.Item("fileResult").Value
                TextBox6.Text = output_Filename
            Next
        End If

        Dim matchesVoltage As MatchCollection = stringPattern_Voltage.Matches(fileReader)
        If matchesVoltage.Count > 0 Then
            For Each match As Match In matchesVoltage
                output_Voltage = match.Groups.Item("voltageResult").Value
                TextBox7.Text = output_Voltage
            Next
        End If

        Dim matchesCurrent As MatchCollection = stringPattern_Current.Matches(fileReader)
        If matchesCurrent.Count > 0 Then
            For Each match As Match In matchesCurrent
                output_Current = match.Groups.Item("currentResult").Value
                TextBox8.Text = output_Current
            Next
        End If

        Dim matchesResolution As MatchCollection = stringPattern_Resolution.Matches(fileReader)
        If matchesResolution.Count > 0 Then
            For Each match As Match In matchesResolution
                output_Resolution = match.Groups.Item("resolutionResult").Value
                Dim resolut As Int32 = Convert.ToInt32(output_Resolution)
                Dim true_resolut As Int32 = resolut / 500

                If true_resolut = 1 Then
                    ComboBox1.SelectedItem = ".5K"
                ElseIf true_resolut = 2 Then
                    ComboBox1.SelectedItem = "1K"
                ElseIf true_resolut = 4 Then
                    ComboBox1.SelectedItem = "2K"
                Else
                End If

            Next
        End If

        Dim matchesExposure As MatchCollection = stringPattern_Exposure.Matches(fileReader)
        If matchesExposure.Count > 0 Then
            For Each match As Match In matchesExposure
                output_Exposure = match.Groups.Item("exposureResult").Value
                TextBox10.Text = output_Exposure
            Next
        End If

        Dim matchesPixelSize As MatchCollection = stringPattern_PixelSize.Matches(fileReader)
        If matchesPixelSize.Count > 0 Then
            For Each match As Match In matchesPixelSize
                output_PixelSize = match.Groups.Item("pixelsizeResult").Value
                TextBox11.Text = output_PixelSize
            Next
        End If

        Dim matchesImageFormat As MatchCollection = stringPattern_ImageFormat.Matches(fileReader)
        If matchesImageFormat.Count > 0 Then
            For Each match As Match In matchesImageFormat
                output_ImageFormat = match.Groups.Item("imageformatResult").Value
                TextBox12.Text = output_ImageFormat
            Next
        End If

        Dim matchesRotationStep As MatchCollection = stringPattern_RotationStep.Matches(fileReader)
        If matchesRotationStep.Count > 0 Then
            For Each match As Match In matchesRotationStep
                output_RotationStep = match.Groups.Item("rotationstepResult").Value
                TextBox14.Text = output_RotationStep
            Next
        End If

        Dim matchesFrames As MatchCollection = stringPattern_Frames.Matches(fileReader)
        If matchesFrames.Count > 0 Then
            For Each match As Match In matchesFrames
                output_Frames = match.Groups.Item("framesResult").Value
                TextBox15.Text = output_Frames
            Next
        End If

        Dim matchesRandomMovement As MatchCollection = stringPattern_RandomMovement.Matches(fileReader)
        If matchesRandomMovement.Count > 0 Then
            For Each match As Match In matchesRandomMovement
                output_RandomMovement = match.Groups.Item("randommovementResult").Value
                TextBox16.Text = output_RandomMovement
            Next
        End If

        Dim matches360 As MatchCollection = stringPattern_360.Matches(fileReader)
        If matches360.Count > 0 Then
            For Each match As Match In matches360
                output_360 = match.Groups.Item("result360").Value
                If String.Compare(output_360.ToString, "YES", True) Then
                    RadioButton1.Checked = True
                Else
                    RadioButton8.Checked = True
                End If
            Next
        End If

        Dim matchesScanDuration As MatchCollection = stringPattern_ScanDuration.Matches(fileReader)
        If matchesScanDuration.Count > 0 Then
            For Each match As Match In matchesScanDuration
                output_ScanDuration = match.Groups.Item("scandurationResult").Value
                TextBox17.Text = output_ScanDuration
            Next
        End If



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim appXL As Excel.Application
        Dim wbXL As Excel.Workbook
        Dim shXL As Excel.Worksheet
        'Dim raXL As Excel.Range


        Dim filename As String = destination
        Dim sheetname As String = "Sheet1"


        appXL = New Excel.Application
        wbXL = appXL.Workbooks.Open(destination)
        shXL = wbXL.Worksheets(sheetname)

        Dim ws As Excel.Worksheet = appXL.ActiveSheet
        Dim lRow = ws.Range("A" & ws.Rows.Count).End(Excel.XlDirection.xlUp).Row

        'For i = 1 To 5
        'ws.Range("A" & lRow).Offset(1, 0).Value = "this is a test! :)"
        'Next



        Dim selectedRB1 As String = Nothing
        Dim selectedRB2 As String = Nothing
        Dim selectedRB3 As String = Nothing

        If RadioButton1.Checked = True Then
            selectedRB1 = "yes"
        ElseIf RadioButton8.Checked = True Then
            selectedRB1 = "no"
        Else
            MsgBox("A Radio Button for Use 360 Rotation needs to be selected!")
            wbXL.Close()
            appXL.Quit()

            releaseObject(appXL)
            releaseObject(wbXL)
            releaseObject(shXL)
            Exit Sub
        End If

        If RadioButton2.Checked = True Then
            selectedRB2 = "yes"
        ElseIf RadioButton3.Checked = True Then
            selectedRB2 = "no"
        Else
            MsgBox("A Radio Button for Log File Archived? needs to be selected!")
            wbXL.Close()
            appXL.Quit()

            releaseObject(appXL)
            releaseObject(wbXL)
            releaseObject(shXL)
            Exit Sub
        End If

        If RadioButton5.Checked = True Then
            selectedRB3 = "yes"
        ElseIf RadioButton4.Checked = True Then
            selectedRB3 = "no"
        Else
            MsgBox("A Radio Button for On Drobo? needs to be selected!")
            wbXL.Close()
            appXL.Quit()

            releaseObject(appXL)
            releaseObject(wbXL)
            releaseObject(shXL)
            Exit Sub
        End If


        Dim DataArray As Object = {TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text,
            ComboBox1.Text, TextBox9.Text, TextBox10.Text, TextBox11.Text, TextBox12.Text, TextBox13.Text, TextBox14.Text, TextBox15.Text, TextBox16.Text,
            selectedRB1, TextBox17.Text, TextBox18.Text, selectedRB2, selectedRB3}

        ws.Range("A" & lRow & ":V" & lRow).Offset(1, 0).Value = DataArray

        wbXL.Close()
        appXL.Quit()

        releaseObject(appXL)
        releaseObject(wbXL)
        releaseObject(shXL)
    End Sub


    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ofd_DestinationExcel.FileOk

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ofd_LogFile.FileName = Label3.Text
        If ofd_LogFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Label3.Text = ofd_LogFile.FileName
            path = ofd_LogFile.FileName
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' RESET ALL FIELDS
        Dim known As String = "Known"
        Dim enter As String = "Enter Here"

        TextBox1.Text = enter
        TextBox2.Text = enter
        TextBox3.Text = known
        TextBox4.Text = enter
        TextBox5.Text = enter
        TextBox6.Text = known
        TextBox7.Text = known
        TextBox8.Text = known
        ComboBox1.Text = "Click to Select"
        TextBox9.Text = enter
        TextBox10.Text = known
        TextBox11.Text = known
        TextBox12.Text = known
        TextBox13.Text = enter
        TextBox14.Text = known
        TextBox15.Text = known
        TextBox16.Text = known
        RadioButton1.Checked = False
        RadioButton8.Checked = False
        TextBox17.Text = known
        TextBox18.Text = enter
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton5.Checked = False
        RadioButton4.Checked = False
    End Sub
End Class
