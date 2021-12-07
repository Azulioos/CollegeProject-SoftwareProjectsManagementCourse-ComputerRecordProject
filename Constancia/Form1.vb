Imports System.Management
Imports System.Text
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class Form1

    Dim _software As String = ""
    Dim _version As String = ""
    Dim _desarrollador As String = ""
    Dim _fechadeuso As String = ""



    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim _nameSpace$ = "root\CIMV2"

        Dim wql = "SELECT * FROM WIN32_PROCESSOR"

        Dim _strbuilder As New StringBuilder

        Using _moSearcher As New ManagementObjectSearcher(_nameSpace, wql)

            For Each _mobject As ManagementObject In _moSearcher.Get
                Label7.Text = $"{_mobject("Name")}"
            Next

        End Using


        Dim wql2 = "SELECT * FROM WIN32_NetworkAdapter Where AdapterType='Ethernet 802.3'"

        Dim _strbuilder2 As New StringBuilder

        Using _moSearcher2 As New ManagementObjectSearcher(_nameSpace, wql2)

            For Each _mobject2 As ManagementObject In _moSearcher2.Get
                Label11.Text = $"{_mobject2("MACAddress")}"
            Next

        End Using


        Dim wql4 = "SELECT * FROM WIN32_BIOS"

        Dim _strbuilder4 As New StringBuilder

        Using _moSearcher4 As New ManagementObjectSearcher(_nameSpace, wql4)

            For Each _mobject4 As ManagementObject In _moSearcher4.Get
                Label18.Text = $"{_mobject4("SerialNumber")}"
            Next

        End Using



        Dim memo = My.Computer.Info.TotalPhysicalMemory
        Dim memo2 = memo * (9.31 * (10 ^ -10))

        Label4.Text = My.Computer.Name
        Label5.Text = My.User.Name
        Label6.Text = My.Computer.Info.OSFullName + " Version: " + My.Computer.Info.OSVersion
        Label9.Text = memo2

        Label25.Text = DateTime.Now.ToString("HH:mm:ss")
        Label26.Text = DateTime.Now.ToString("dddd,dd,MMMM,yyy")




    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)


        Dim _nameSpace$ = "root\CIMV2"

        Dim wql3 = "SELECT * FROM WIN32_Product"

        Dim _strbuilder3 As New StringBuilder


        Using _moSearcher3 As New ManagementObjectSearcher(_nameSpace, wql3)


            For Each _mobject3 As ManagementObject In _moSearcher3.Get
                _software = _software + $"{_mobject3("Name")}" & vbCrLf
                _version = _version + $"{_mobject3("version")}" & vbCrLf
                _desarrollador = _desarrollador + $"{_mobject3("vendor")}" & vbCrLf
                _fechadeuso = _fechadeuso + $"{_mobject3("installdate")}" & vbCrLf



            Next

        End Using




    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click

    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click

    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click

    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click

    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim SaveFileDialog As New SaveFileDialog
        Dim ruta As String
        With SaveFileDialog
            .Title = "Guardar"
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Filter = "Archivo pdf (*.pdf)|*.pdf"
            .FileName = "Proyecto ordinario - Administracion de software"
            .OverwritePrompt = True
            .CheckPathExists = True
        End With

        If SaveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            ruta = SaveFileDialog.FileName
        Else
            ruta = String.Empty
            Exit Sub
        End If

        Try
            Dim document As New iTextSharp.text.Document(PageSize.LETTER)
            document.PageSize.Rotate()

            document.AddAuthor(Label1.ToString)
            document.AddTitle("Crear pdf")


            Dim writer As PdfWriter = PdfWriter.GetInstance(document, New System.IO.FileStream _
            (ruta, System.IO.FileMode.Create))
            writer.ViewerPreferences = PdfWriter.PageLayoutSinglePage

            document.Open()
            Dim cb As PdfContentByte = writer.DirectContent
            Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)

            cb.SetFontAndSize(bf, 10)

            Dim pdfTable As New PdfPTable(2)

            Dim PdfTitulo As New PdfPCell(New Phrase("Constancia de uso de un ordenador - Proyecto ordinario", New Font(bf, 20, Font.Bold)))
            PdfTitulo.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER
            PdfTitulo.Colspan = 4
            PdfTitulo.Border = 0
            PdfTitulo.FixedHeight = 60
            pdfTable.AddCell(PdfTitulo)



            Dim Table2 As PdfPTable = New PdfPTable(4)
            Dim Table1 As PdfPTable = New PdfPTable(4)
            Dim Table3 As PdfPTable = New PdfPTable(4)
            Dim Table4 As PdfPTable = New PdfPTable(2)

            Dim width1 As Single() = New Single() {1.0F, 1.0F, 1.0F, 1.0F}
            Dim width2 As Single() = New Single() {2.5F, 0.5F, 1.0F, 0.5F}
            Dim width3 As Single() = New Single() {2.5F, 0.5F, 1.0F, 0.5F}
            Dim width4 As Single() = New Single() {1.0F, 1.0F}

            Table1.WidthPercentage = 90
            Table2.WidthPercentage = 90
            Table3.WidthPercentage = 90
            Table4.WidthPercentage = 90

            Dim CVacio As PdfPCell = New PdfPCell(New Phrase(""))

            Dim Columna_1 As PdfPCell
            Dim Columna_2 As PdfPCell
            Dim Columna_3 As PdfPCell
            Dim Columna_4 As PdfPCell

            Dim Columna_5 As PdfPCell
            Dim Columna_6 As PdfPCell
            Dim Columna_7 As PdfPCell
            Dim Columna_8 As PdfPCell

            Dim Columna_9 As PdfPCell
            Dim Columna_10 As PdfPCell
            Dim Columna_11 As PdfPCell
            Dim Columna_12 As PdfPCell

            Dim Columna_13 As PdfPCell
            Dim Columna_14 As PdfPCell

            CVacio.Border = 0
            Table1.SetWidths(width1)
            Table2.SetWidths(width2)
            Table3.SetWidths(width3)
            Table4.SetWidths(width4)

            Table1.AddCell(CVacio)
            Table1.AddCell(CVacio)
            Table1.AddCell(CVacio)
            Table1.AddCell(CVacio)

            Table2.AddCell(CVacio)
            Table2.AddCell(CVacio)
            Table2.AddCell(CVacio)
            Table2.AddCell(CVacio)


            Table3.AddCell(CVacio)
            Table3.AddCell(CVacio)
            Table3.AddCell(CVacio)
            Table3.AddCell(CVacio)

            Table4.AddCell(CVacio)
            Table4.AddCell(CVacio)

            Columna_5 = New PdfPCell(New Phrase("ID del equipo: " & Me.Label18.Text, New Font(bf, 6)))
            Columna_5.Border = 0
            Table1.AddCell(Columna_5)

            Columna_6 = New PdfPCell(New Phrase("Nombre del equipo: " & Me.Label4.Text, New Font(bf, 6)))
            Columna_6.Border = 0
            Table1.AddCell(Columna_6)

            Columna_7 = New PdfPCell(New Phrase("Procesador del equipo: " & Me.Label7.Text, New Font(bf, 6)))
            Columna_7.Border = 0
            Table1.AddCell(Columna_7)

            Columna_8 = New PdfPCell(New Phrase("Memoria Instalada: " & Me.Label9.Text + "GB", New Font(bf, 6)))
            Columna_8.Border = 0
            Table1.AddCell(Columna_8)



            Columna_5 = New PdfPCell(New Phrase("Direccion MAC: " & Me.Label11.Text, New Font(bf, 6)))
            Columna_5.Border = 0
            Table1.AddCell(Columna_5)

            Columna_6 = New PdfPCell(New Phrase("Nombre del Usuario: " & Me.Label5.Text, New Font(bf, 6)))
            Columna_6.Border = 0
            Table1.AddCell(Columna_6)

            Columna_7 = New PdfPCell(New Phrase("Sistema Operativo: " & Me.Label6.Text, New Font(bf, 6)))
            Columna_7.Border = 0
            Table1.AddCell(Columna_7)

            Columna_8 = New PdfPCell(New Phrase("-----------------", New Font(bf, 6)))
            Columna_8.Border = 0
            Table1.AddCell(Columna_8)

            Table1.AddCell(CVacio)
            Table1.AddCell(CVacio)
            Table1.AddCell(CVacio)
            Table1.AddCell(CVacio)

            Columna_9 = New PdfPCell(New Phrase("Software", New Font(bf, 5)))
            Columna_9.Border = 0
            Table3.AddCell(Columna_9)

            Columna_10 = New PdfPCell(New Phrase("Version", New Font(bf, 5)))
            Columna_10.Border = 0
            Table3.AddCell(Columna_10)

            Columna_11 = New PdfPCell(New Phrase("Desarrolladores", New Font(bf, 5)))
            Columna_11.Border = 0
            Table3.AddCell(Columna_11)

            Columna_12 = New PdfPCell(New Phrase("Fecha de uso", New Font(bf, 5)))
            Columna_12.Border = 0
            Table3.AddCell(Columna_12)



            Dim _nameSpace$ = "root\CIMV2"

            Dim wql3 = "SELECT * FROM WIN32_Product"

            Dim _strbuilder3 As New StringBuilder

            Using _moSearcher3 As New ManagementObjectSearcher(_nameSpace, wql3)

                For Each _mobject3 As ManagementObject In _moSearcher3.Get


                    Columna_1 = New PdfPCell(New Phrase($"{_mobject3("Name")}", New Font(bf, 5)))
                    Columna_1.Border = 0
                    Table2.AddCell(Columna_1)

                    Columna_2 = New PdfPCell(New Phrase($"{_mobject3("version")}", New Font(bf, 5)))
                    Columna_2.Border = 0
                    Table2.AddCell(Columna_2)

                    Columna_3 = New PdfPCell(New Phrase($"{_mobject3("vendor")}", New Font(bf, 5)))
                    Columna_3.Border = 0
                    Table2.AddCell(Columna_3)

                    Columna_4 = New PdfPCell(New Phrase($"{_mobject3("installdate")}", New Font(bf, 5)))
                    Columna_4.Border = 0
                    Table2.AddCell(Columna_4)



                Next

            End Using

            Table2.AddCell(CVacio)
            Table2.AddCell(CVacio)
            Table2.AddCell(CVacio)
            Table2.AddCell(CVacio)

            Columna_13 = New PdfPCell(New Phrase("Nombre del firmante: " & Me.TextBox2.Text, New Font(bf, 10)))
            Columna_13.Border = 0
            Table4.AddCell(Columna_13)

            Columna_14 = New PdfPCell(New Phrase("CURP: " & Me.TextBox1.Text, New Font(bf, 10)))
            Columna_14.Border = 0
            Table4.AddCell(Columna_14)

            Columna_13 = New PdfPCell(New Phrase("Fecha y hora de firmado: " & Me.Label26.Text, New Font(bf, 10)))
            Columna_13.Border = 0
            Table4.AddCell(Columna_13)

            Columna_14 = New PdfPCell(New Phrase(Me.Label25.Text, New Font(bf, 10)))
            Columna_14.Border = 0
            Table4.AddCell(Columna_14)

            document.Add(pdfTable)
            document.Add(Table1)
            document.Add(Table3)
            document.Add(Table2)
            document.Add(Table4)

            document.Close()

        Catch ex As Exception
            MessageBox.Show("Error de la generacion", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs)

    End Sub
End Class
