Imports PA.Office.ExcelObjects
Imports System.IO

Public Class Form1

	Public Sub Test()

		Dim cellCollection As ExcelCellObjectCollection = New ExcelCellObjectCollection()
		cellCollection(0).Value = "aaa"

		Dim row As New ExcelRowObject
		row.Cells(0).Value = "aaa"

	End Sub

	Private Sub Button1_Click( _
		ByVal sender As System.Object, _
		ByVal e As System.EventArgs) _
		Handles Button1.Click

		Dim excel As New ExcelBookControl

		Dim sheet1 As New ExcelSheetObject
		sheet1.Name = "�e�X�g�P"
		SetSheet1DataAt(sheet1)
		excel.Sheets.Add(sheet1)

		Dim sheet2 As New ExcelSheetObject
		sheet2.Name = "�e�X�g�Q"
		SetSheet2DataAt(sheet2)
		excel.Sheets.Add(sheet2)

		Dim sheet3 As New ExcelSheetObject
		sheet3.Name = "�O���t"

		Dim chart As New ExcelChartObject
		chart.ChartName = "���i�ϓ��O���t"
		chart.ChartType = ExcelChartType.LineMarkers
        chart.PositionX = 100
        chart.PositionY = 100
        chart.SetDataSource(1, 0, 0, 3, 0)

		sheet3.Charts.Add(chart)
		excel.Sheets.Add(sheet3)

		Dim sheet4 As New ExcelSheetObject
		sheet4.Name = "�B���V�[�g"
		sheet4.Visible = False
		excel.Sheets.Add(sheet4)

		excel.Show()

	End Sub

	Private Sub SetSheet1DataAt(ByVal sheet As ExcelSheetObject)

		'---- row1
		Dim row1 As New ExcelRowObject

		Dim cell1_1 As New ExcelCellObject

		cell1_1.Value = 10

		row1.Cells.Add(cell1_1)

		sheet.Rows.Add(row1)


		'---- row2
		Dim row2 As New ExcelRowObject

		Dim cell2_1 As New ExcelCellObject

		cell2_1.Value = 20

		row2.Cells.Add(cell2_1)

		sheet.Rows.Add(row2)


		'---- row3
		Dim row3 As New ExcelRowObject

		Dim cell3_1 As New ExcelCellObject

		cell3_1.Value = 40

		row3.Cells.Add(cell3_1)

		sheet.Rows.Add(row3)

	End Sub

	Private Sub SetSheet2DataAt(ByVal sheet As ExcelSheetObject)

		'---- row1
		Dim row1 As New ExcelRowObject

		Dim cell1_1 As New ExcelCellObject
		Dim cell1_2 As New ExcelCellObject
		Dim cell1_3 As New ExcelCellObject

		cell1_1.Value = "����������"
		cell1_2.Value = "����������"
		cell1_3.Value = "����������"

		row1.Cells.Add(cell1_1)
		row1.Cells.Add(cell1_2)
		row1.Cells.Add(cell1_3)

		sheet.Rows.Add(row1)


		'---- row2
		Dim row2 As New ExcelRowObject

		Dim cell2_1 As New ExcelCellObject
		Dim cell2_2 As New ExcelCellObject
		Dim cell2_3 As New ExcelCellObject

		cell2_1.Value = "�����Ă�"
		cell2_2.Value = "�Ȃɂʂ˂�"
		cell2_3.Value = "�܂݂ނ߂�"

		row2.Cells.Add(cell2_1)
		row2.Cells.Add(cell2_2)
		row2.Cells.Add(cell2_3)

		sheet.Rows.Add(row2)


		'---- row3
		Dim row3 As New ExcelRowObject

		Dim cell3_1 As New ExcelCellObject
		Dim cell3_2 As New ExcelCellObject
		Dim cell3_3 As New ExcelCellObject

		cell3_1.Value = "�₢�䂦��"
		cell3_2.Value = "������"
		cell3_3.Value = "�����"

		row3.Cells.Add(cell3_1)
		row3.Cells.Add(cell3_2)
		row3.Cells.Add(cell3_3)

		sheet.Rows.Add(row3)

	End Sub

	Private Sub Button2_Click( _
		ByVal sender As System.Object, _
		ByVal e As System.EventArgs) _
		Handles Button2.Click

		Dim excel As New ExcelBookControl()
        excel.AddLoadingAreaSetting(100, 100)
        'excel.AddLoadingAreaSetting(3, 5)

        excel.LoadFrom("D:\Excel\Template\�����L�^_template.xlsx")

        Dim obj As Object = excel.Sheets(0).Rows(2 - 1).Cells(2 - 1).Value

        'MsgBox("(2,2)�̃f�[�^��[" + excel.Sheets(1).Rows(2 - 1).Cells(2 - 1).Value.ToString() + "]�ł�")

        'Dim chart As New ExcelChartObject
        'chart.SetDataSource(1, 0, 1, 3, 1)

        excel.Sheets(0).Rows(2 - 1).Cells(2 - 1).Value = 100
        'excel.Sheets(0).Charts.Add(chart)
        'excel.SaveAs("C:\VisualStudioSolution\pa_common\SRC\PA.Office\PA.OfficeAllVersion.TestForm\bin\Debug\templete.xls")
        'excel.SaveAs("C:\temp\templete.xls")
        excel.SaveAs("D:\Excel\���[�o��\�����L�^.xlsx")

        MsgBox("(2,2)�̃f�[�^�����������܂����B�m�F���ĉ������B")

	End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        '[���O��t���ĕۑ�](�_�C�A���O)
        Me.saveExcelFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Me.saveExcelFileDialog.FileName = "TestOutput.xls"
        Me.saveExcelFileDialog.Filter = "XLS�`���i*.xls�j|*.xls"

        Dim dialogResult As DialogResult = Me.saveExcelFileDialog.ShowDialog()

        If dialogResult = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        Dim outputFileName As String = Me.saveExcelFileDialog.FileName
        Dim templeteFileName As String = "C:\Documents and Settings\wcheng\�f�X�N�g�b�v\�R�s�[ �` InspectStatusListTemplate.xls"

        File.Copy(templeteFileName, outputFileName, True)
        File.SetAttributes(outputFileName, FileAttributes.Normal)

        'Dim cells As New List(Of ExcelCellObject)
        'Dim cell1 As New ExcelCellObject
        'cells.Add(cell1)
        'cell1.ColIndex = 22
        'cell1.RowIndex = 6
        ''cell1.Range = "V6"
        'cell1.Value = "����P�Q�R"
        'cell1.ColorIndex = 3


        'Dim cell2 As New ExcelCellObject
        'cells.Add(cell2)
        'cell2.ColIndex = 8
        'cell2.RowIndex = 24
        ''cell2.Range = "H24"
        'cell2.Value = "�e�X�g�`�F�b�N���X�g���ڏڍ׈ꗗ�`�F�b�N���X�g���ڏڍ׈ꗗ"
        'cell2.ColorIndex = 5
        'cell2.FontColorIndex = 3

        'Dim images As New List(Of ExcelImageObject)
        'Dim image1 As New ExcelImageObject
        'image1.RowIndex = 4
        'image1.ColIndex = 4
        'image1.ImageData = Image.FromFile("C:\Documents and Settings\wcheng\�f�X�N�g�b�v\Winter.jpg")
        'image1.ImageData = image1.ImageData.GetThumbnailImage(75, 75, Nothing, IntPtr.Zero)
        'images.Add(image1)

        '�ڕW�s�̉�����R�s�[����
        Dim excelFileWriter As ExcelFileSingleton = ExcelFileSingleton.GetInstance()
        Try
            excelFileWriter.OpenExcel(outputFileName)
            'excelFileWriter.WriterCellsToSheet("CL�ڍ�", cells)

            'excelFileWriter.InsertRowOfSheet("CL�ڍ�", 38)

            'excelFileWriter.SheetRangeCopy("CL�ڍ�", "D19", "H27")
            'excelFileWriter.AddWorksheetAfter("�Y�t1", "�Y�t")
            'excelFileWriter.WriteImagesToSheet("�Y�t", images, AddressOf SetClipboardDataObject)
            excelFileWriter.SetRangeLineStyle("Sheet1", "A25:E30", _
            xlLineStyle.xlNone, _
            xlLineStyle.xlNone, _
            xlLineStyle.xlContinuous, _
            xlLineStyle.xlContinuous, _
            xlLineStyle.xlContinuous, _
            xlLineStyle.xlContinuous)


        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        '[���O��t���ĕۑ�](�_�C�A���O)
        Me.saveExcelFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Me.saveExcelFileDialog.FileName = "TestOutput.xls"
        Me.saveExcelFileDialog.Filter = "XLS�`���i*.xls�j|*.xls"

        Dim dialogResult As DialogResult = Me.saveExcelFileDialog.ShowDialog()

        If dialogResult = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        Dim outputFileName As String = Me.saveExcelFileDialog.FileName
        Dim templeteFileName As String = "C:\Documents and Settings\wcheng\�f�X�N�g�b�v\IssueListTemplate.xls"

        File.Copy(templeteFileName, outputFileName, True)
        File.SetAttributes(outputFileName, FileAttributes.Normal)

        '�V�[�g��
        Dim sheetName As String = "���_L�i�o�j"

        Dim rows As List(Of ExcelRowObject) = GetRowsData()

        '�ڕW�s�̉�����R�s�[����
        Dim excelSingleton As ExcelFileSingleton = ExcelFileSingleton.GetInstance()
        Try
            excelSingleton.OpenExcel(outputFileName)

            excelSingleton.WriteRowsToSheetByArray(sheetName, rows, 8, 2)

        Finally
            excelSingleton.CloseExcel()
        End Try

    End Sub

    Private Function GetRowsData() As List(Of ExcelRowObject)

        Dim rows As New List(Of ExcelRowObject)

        Dim row1 As New ExcelRowObject
        Dim row2 As New ExcelRowObject
        Dim row3 As New ExcelRowObject

        Dim cell82 As New ExcelCellObject
        cell82.RowIndex = 8
        cell82.ColIndex = 2
        cell82.Value = "82"
        Dim cell83 As New ExcelCellObject
        cell83.RowIndex = 8
        cell83.ColIndex = 3
        cell83.Value = "83"
        Dim cell84 As New ExcelCellObject
        cell84.RowIndex = 8
        cell84.ColIndex = 4
        cell84.Value = "84"
        row1.Cells.Add(cell82)
        row1.Cells.Add(cell83)
        row1.Cells.Add(cell84)

        Dim cell95 As New ExcelCellObject
        cell95.RowIndex = 9
        cell95.ColIndex = 5
        cell95.Value = "95"
        Dim cell96 As New ExcelCellObject
        cell96.RowIndex = 9
        cell96.ColIndex = 6
        cell96.Value = "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "12345678901" + _
        "1234567890"
        Dim cell97 As New ExcelCellObject
        cell97.RowIndex = 9
        cell97.ColIndex = 7
        cell97.Value = "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" + _
        "12345678901" 
        row1.Cells.Add(cell95)
        row1.Cells.Add(cell96)
        row1.Cells.Add(cell97)

        rows.Add(row1)
        rows.Add(row2)
        ' rows.Add(row3)

        Return rows

    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        '[���O��t���ĕۑ�](�_�C�A���O)
        Me.saveExcelFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Me.saveExcelFileDialog.FileName = "TestOutput.xlsx"
        Me.saveExcelFileDialog.Filter = "XLS�`���i*.xlsx�j|*.xlsx"

        Dim dialogResult As DialogResult = Me.saveExcelFileDialog.ShowDialog()

        If dialogResult = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        Dim outputFileName As String = Me.saveExcelFileDialog.FileName
        Dim templeteFileName As String = "D:\Excel\Template\�����L�^_template_1.xlsx"

        File.Copy(templeteFileName, outputFileName, True)
        File.SetAttributes(outputFileName, FileAttributes.Normal)

        '�V�[�g��
        Dim sheetName As String = "�����L�^_��.�X�ܑS��"

        'Dim rows As List(Of ExcelRowObject) = GetRowsData()

        '�ڕW�s�̉�����R�s�[����
        Dim excelSingleton As ExcelFileSingleton = ExcelFileSingleton.GetInstance()
        Try
            excelSingleton.OpenExcel(outputFileName)

            'excelSingleton.WriteRowsToSheetByArray(sheetName, rows, 8, 2)

            excelSingleton.InsertColOfSheet(sheetName, "D", 3)

        Finally
            excelSingleton.CloseExcel()
        End Try

        MsgBox("���[�o�͂��܂���")

    End Sub
End Class
