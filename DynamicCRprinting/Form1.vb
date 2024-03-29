Imports System.IO
Imports System.Text
Imports System.Timers
Imports System.Data.SqlClient
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.CrystalReports
Imports CrystalDecisions.Shared
Imports QRCoder
Imports System.Globalization
Imports System.Configuration
Imports System.Drawing
Imports System.Drawing.Printing
Public Class Form1
    Dim timer As Timer
    Dim filen As String = ""
    Dim header_picture As String = ""
    Dim InvoiceDataTable As New DataSet1.InvoiceDataTable
    Public Delegate Sub MethodInvoker()
    Public Delegate Sub PrintDelegate()
 

    Private Shared Sub OnTimedEvent(ByVal source As Object, ByVal e As System.Timers.ElapsedEventArgs)
        Console.WriteLine("The Elapsed event was raised at {0}", e.SignalTime)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        timer = New Timer(500)
        AddHandler timer.Elapsed, New ElapsedEventHandler(AddressOf TimerElapsed)
        timer.AutoReset = True
        timer.Enabled = True
        timer.Start()
    End Sub

    Sub TimerElapsed(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        ' this function check to print for new file


        Dim time As DateTime = e.SignalTime



        If (Not New DirectoryInfo(My.Settings.to_print).EnumerateFiles("*.txt").Any()) Then

            ' write on the log screen

            txtlogs.Invoke(Sub() Display("no file found: " + time))
        Else


            ' write on the log screen

            txtlogs.Invoke(Sub() Display(" file found: " + time))
            timer.Stop()

            ' parse the folder to print

            Dim folderBrowseDialog As FolderBrowserDialog = New FolderBrowserDialog()
            folderBrowseDialog.SelectedPath = My.Settings.to_print
            Dim driNFO As New DirectoryInfo(folderBrowseDialog.SelectedPath)
            Dim txtFiles As FileInfo() = driNFO.GetFiles("*.txt")

            ' parse all txt file on the folder

            For Each txt_file As FileInfo In txtFiles
                filen = txt_file.Name.Replace(".txt", "")

                txtlogs.Invoke(Sub() Display(" file found name : " + txt_file.Name))
                'txtlogs.AppendText(txt_file.Name)

                Try
                    'read file to print and create pdf file and send to printer

                    readprintfile(My.Settings.to_print + txt_file.Name)


                    'move treated to done folder

                    File.Delete(My.Settings.done + txt_file.Name)
                    File.Move(My.Settings.to_print + txt_file.Name, My.Settings.done + txt_file.Name)
                    filen = ""
                    InvoiceDataTable.Rows.Clear()
                Catch ex As Exception
                    ' if file not treated  move the file to error folder

                    txtlogs.Invoke(Sub() Display("  error in this file : " + txt_file.Name))
                    txtlogs.Invoke(Sub() Display(ex.Message.ToString()))
                    File.Delete(My.Settings.error_pdf + txt_file.Name)
                    File.Move(My.Settings.to_print + txt_file.Name, My.Settings.error_pdf + txt_file.Name)
                End Try
            Next
            timer.Start()

        End If
    End Sub


    Public Sub readprintfile(filename As String)


        'variable of first  Line
        Dim Printer_name As String = ""
        Dim number_of_copies As String = ""

        'variable of Second Line
        Dim Company_VAT_number As String = ""
        Dim company_name As String = ""
        Dim company_address As String = ""
        Dim company_email As String = ""
        Dim company As String = ""


        'variable of third Line
        Dim Invoice_date = "", Invoice_Time = "", Reference_No1 = "", Reference_No2 = "", Invoice_number = "", Invoice_Description = "", customer_name = "", customer_Vat_number = "", customer_address = "", customer_number = "", Warehouse_number = "", Warehouse_Name = "", total_weight = "", Invoice_Type = "", Customer_Other_Information = "", Salesman_No = "", Salesman_Name = "", Extra_Field_Caption1 = "", Extra_Field_Date1 = "", Extra_Field_Caption2 = "", Extra_Field_Date2 = "", Extra_Field_Caption3 = "", Extra_Field_Date3 = ""

        'variable of fourth Line 

        Dim Total_excluding_VAT = "", Discount = "", VAT_amount = "", total_amount = "", Amount_in_words = "", User_ID = "", User_Name = "", Additional_Charge_Amount = "", Additional_Charge_Description

        'variable of  fourth Line

        Dim Item_Number = "", Item_Arabic_Name = "", Item_English_Name = "", Unit = "", Expiry_Date = "", Batch_No = "", Quantity = "", Unit_Price = "", Total_Excluding_vat1 = "", VAT = "", Vat_amount1 = "", Total = ""

        ' Dim vText As String
        Dim vstring(-1) As String
        Dim p1 As String() = {vbTab} 'Note I am using a string array and the vbTab Constant
        Dim vData As String = ""
        Dim rownumber As Integer = 0

        'Read the invoice txt file 
        'Using rvsr As New StreamReader(filename, Encoding.GetEncoding("windows-1256"))
        Dim rows = IO.File.ReadAllLines(filename, Encoding.GetEncoding("windows-1256"))
        'While rvsr.Peek <> -1
        For Each row In rows

            'vText = rvsr.ReadLine()
            Dim vText = row

            rownumber = rownumber + 1

            'read the first line  and save fields into variable

            If rownumber = 1 Then
                vstring = vText.Split(p1, StringSplitOptions.RemoveEmptyEntries) 'I am also using the option to remove empty entries a
                Printer_name = vstring(0)
                number_of_copies = vstring(1)
                'vData = vData + vbCrLf + "Printer name: " + Printer_name + "  number of copies : " + number_of_copies
            End If

            'read the second line  and save fields into variable

            If rownumber = 2 Then
                vstring = vText.Split(p1, StringSplitOptions.None) 'I am also using the option to remove empty entries a

                Company_VAT_number = vstring(0)
                company_name = Mid(vstring(1), 1, 1) + Replace(Mid(vstring(1), 2, 1), " ", "") + Mid(vstring(1), 3)
                company_address = vstring(2)
                company_email = vstring(3)
                company = vstring(4)
                'vData = vData + vbCrLf + Company_VAT_number + " --- " + company_name + " --- " + company_address + "---" + company_email + "---" + company

                ' verify if the key in the config file equal to vat number of company else the software will close
                If DecodeBase64(My.Settings.key) <> Company_VAT_number Then
                    MsgBox("you are using incorrect key, please contact the software owner!!!!")
                    Me.CloseMe()
                End If
            End If

            'read the third line  and save fields into variable
            If rownumber = 3 Then
                vstring = vText.Split(p1, StringSplitOptions.None) 'I am also using the option to remove empty entries a

                Invoice_date = vstring(0) : Invoice_Time = vstring(1) : Reference_No1 = vstring(2) : Reference_No2 = vstring(3) : Invoice_number = vstring(4) : Invoice_Description = vstring(5) : customer_name = vstring(6) : customer_Vat_number = vstring(7) : customer_address = vstring(8) : customer_number = vstring(9) : Warehouse_number = vstring(10) : Warehouse_Name = vstring(11) : total_weight = vstring(12) : Invoice_Type = vstring(13) : Customer_Other_Information = vstring(14) : Salesman_No = vstring(15) : Salesman_Name = vstring(16) : Extra_Field_Caption1 = vstring(17) : Extra_Field_Date1 = vstring(18) : Extra_Field_Caption2 = vstring(19) : Extra_Field_Date2 = vstring(20) : Extra_Field_Caption3 = vstring(21) : Extra_Field_Date3 = vstring(22)
            End If

            'read the fourth line  and save fields into variable

            If rownumber = 4 Then
                vstring = vText.Split(p1, StringSplitOptions.None) 'I am also using the option to remove empty entries a
                Total_excluding_VAT = vstring(0) : Discount = vstring(1) : VAT_amount = vstring(2) : total_amount = vstring(3) : Amount_in_words = vstring(4) : User_ID = vstring(5) : User_Name = vstring(0) : Additional_Charge_Amount = vstring(6) : Additional_Charge_Description = vstring(7)
            End If

            'read the détails line  and save fields into variable and insert the lines into datatable
            If rownumber >= 5 Then
                vstring = vText.Split(p1, StringSplitOptions.None) 'I am also using the option to remove empty entries a
                Item_Number = vstring(0) : Item_Arabic_Name = vstring(1) : Item_English_Name = vstring(2) : Unit = vstring(3) : Expiry_Date = vstring(4) : Batch_No = vstring(5) : Quantity = vstring(6) : Unit_Price = vstring(7) : Total_Excluding_vat1 = vstring(8) : VAT = vstring(9) : Vat_amount1 = vstring(10) : Total = vstring(11)

                ' generate the VAT code number and generate the Qr csode
                'Dim Date1 As String = Mid(Invoice_date, 7, 4) + Mid(Invoice_date, 3, 4) + Mid(Invoice_date, 1, 2) + " " + Invoice_Time

                Dim originalDateTime As DateTime = DateTime.ParseExact(Invoice_date & " " & Invoice_Time, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture)
                Dim Date1 As String = originalDateTime.ToString("yyyy-MM-ddTHH:mm:sszzz")
                Dim Qrcode As Byte() = GenerateQrCode(generateTLV(company_name.Trim(), Company_VAT_number.Trim(), Date1, total_amount.Trim.ToString.Replace(",", ""), VAT_amount.ToString.Replace(",", "").Trim()))

                ' insert the line into datatable
                InvoiceDataTable.Rows.Add(Company_VAT_number, company_name, company_address, company_email, company, Invoice_date, Invoice_Time, Reference_No1, Reference_No2, Invoice_number, Invoice_Description, customer_name, customer_Vat_number, customer_address, customer_number, Warehouse_number, Warehouse_Name, total_weight, Invoice_Type, Customer_Other_Information, Salesman_No, Salesman_Name, Extra_Field_Caption1, Extra_Field_Date1, Extra_Field_Caption2, Extra_Field_Date2, Extra_Field_Caption3, Extra_Field_Date3, Total_excluding_VAT, Discount, VAT_amount, total_amount, Amount_in_words, User_ID, User_Name, Additional_Charge_Amount, Additional_Charge_Description, Item_Number, Item_Arabic_Name, Item_English_Name, Unit, Expiry_Date, Batch_No, Quantity, Unit_Price, Total_Excluding_vat1, VAT, Vat_amount1, Total, header_picture, Qrcode, My.Settings.footerenglish, My.Settings.footerarabic)

                ' To log prnting duplicate items
                My.Computer.FileSystem.WriteAllText(
                My.Settings.error_pdf + "\LogFile.Txt", filename + " " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " " + Item_Number + vbNewLine, True)
            End If

            'End While
        Next

        ' read the header picture
        header_picture = My.Settings.header + "header.jpg"
        'generate the pdf file and  Send file to printer


        print(filename, Printer_name, number_of_copies)


            'Generate VAT number like Fatora

            'MsgBox(generateTLV(company_name, Company_VAT_number, Invoice_date + "t" + Invoice_Time, total_amount.Trim, VAT_amount))
            'Dim oDate As DateTime = Convert.ToDateTime(Invoice_date)
            'Invoice_date = oDate.ToString("yyyy-MM-dd")
            'MsgBox(generateTLV(company_name, Company_VAT_number, Invoice_date + "T" + Invoice_Time + "Z", Total_excluding_VAT.Trim, total_amount.Trim))



        'End Using

    End Sub

    ' display text on text logs screen
    Private Sub Display(text As String)
        txtlogs.AppendText(Environment.NewLine + text)
        txtlogs.ScrollToCaret()
    End Sub

    ' close the screen
    Private Sub CloseMe()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf CloseMe))
            Exit Sub
        End If
        Me.Close()
    End Sub

    ' if we want show report on crystal report viewer

    Private Sub Displayreport(report As CrystalDecisions.CrystalReports.Engine.ReportDocument)
        'CrystalReportViewer1.ReportSource = report

    End Sub

    ' this export report to pdf and  send file to printer
    Public Sub print(filename As String, printername As String, copynumber As String)
        Dim dt As DataTable = InvoiceDataTable
        Dim Report As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        'Try

        Report.Load("rpt_report_41.rpt")
        Report.Database.Tables("Invoice").SetDataSource(dt)

        'CrystalReportViewer1.Invoke(Sub() Displayreport(Report))


        Dim CrExportOptions As ExportOptions
        Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
        Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
        CrDiskFileDestinationOptions.DiskFileName = My.Settings.invoice + filen + ".pdf"
        CrExportOptions = Report.ExportOptions
        With CrExportOptions
            .ExportDestinationType = ExportDestinationType.DiskFile
            .ExportFormatType = ExportFormatType.PortableDocFormat
            .DestinationOptions = CrDiskFileDestinationOptions
            .FormatOptions = CrFormatTypeOptions
        End With
        Report.Export(CrExportOptions)

        ' send report to printer using their name
        Try
            Report.PrintOptions.PrinterName = printername
            Report.PrintToPrinter(Val(copynumber.Trim), False, 0, 0)
        Catch ex As Exception

        End Try


        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    Report.Dispose()
        'End Try
    End Sub

    ' send the form to system tray icon

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            NotifyIcon1.Visible = True
            Me.Hide()
            NotifyIcon1.BalloonTipText = "Hi from right system tray"
            NotifyIcon1.ShowBalloonTip(500)
        End If
    End Sub

Private Sub NotifyIcon1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        NotifyIcon1.Visible = False
    End Sub


    Private Function GenerateQrCode(qrmsg As String) As Byte()
        Dim code As String = qrmsg
        Dim qrGenerator As New QRCodeGenerator()
        Dim qrCode As QRCodeGenerator.QRCode = qrGenerator.CreateQrCode(code, QRCodeGenerator.ECCLevel.H)
        Dim imgBarCode As New System.Web.UI.WebControls.Image()
        imgBarCode.Height = 150
        imgBarCode.Width = 150
        Using bitMap As Bitmap = qrCode.GetGraphic(20)
            Using ms As New MemoryStream()
                bitMap.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                Dim byteImage As Byte() = ms.ToArray()
                Return byteImage
            End Using
        End Using
    End Function



    'get TLV of value
    Private Function getTLV(tag As String, value As String) As String
        Return Chr(tag) & Chr(value.Length) & value
    End Function


    'generate TLV  tag vakue lentgh encoded to base 64 
    Private Function generateTLV(value1 As String, value2 As String, value3 As String, value4 As String, value5 As String) As String
        Try
 
            value1 = getTLV("01", value1)
            value2 = getTLV("02", value2)
            value3 = getTLV("03", value3)
            value4 = getTLV("04", value4)
            value5 = getTLV("05", value5)
            Dim b As Byte() = System.Text.Encoding.UTF8.GetBytes(value1 & value2 & value3 & value4 & value5)
            Dim t As String = Convert.ToBase64String(b)

            Return t
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    ' Decode base 64 
    Public Function DecodeBase64(input As String) As String
        Dim base64Decoded As String
        Dim data() As Byte
        data = System.Convert.FromBase64String(input)
        base64Decoded = System.Text.ASCIIEncoding.ASCII.GetString(data)
        'System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(input)) 
        Return base64Decoded
    End Function

  
End Class
