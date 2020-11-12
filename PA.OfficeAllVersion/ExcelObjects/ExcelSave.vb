Imports System
Imports System.Diagnostics
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.IO

Namespace ExcelObjects

	''' <summary>
	''' EXCELファイルをPDFファイルに変換する
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelSave

		Public Sub SaveAsPdf(excelFilePathName As String, saveAsPathName As String)
			Dim application As Application = Nothing
			Dim books As Workbooks = Nothing
			Dim book As Workbook = Nothing

			Try
				''Applicationクラスのインスタンス作成
				application = New Application()
				books = application.Workbooks
				book = books.Open(excelFilePathName)

				If File.Exists(saveAsPathName) Then
					File.Delete(saveAsPathName)

				End If

				book.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF,
										 saveAsPathName,
										 XlFixedFormatQuality.xlQualityStandard)


			Catch ex As Exception

				Dim excels As Process() = Process.GetProcessesByName("EXCEL")
				For Each x As Process In excels
					x.Kill()
				Next
				Throw ex

			Finally
				If book IsNot Nothing Then
					book.Close(True)
				End If

				If application IsNot Nothing Then
					application.Quit()
				End If

				FinalReleaseComObjects(book, books, application)

			End Try
		End Sub

		Private Sub FinalReleaseComObjects(ByVal ParamArray objects As Object())
			For Each obj As Object In objects
				Try
					If obj Is Nothing Then
						Continue For
					End If

					If Marshal.IsComObject(obj) = False Then
						Continue For
					End If

					Marshal.FinalReleaseComObject(obj)
				Catch ex As Exception

				End Try
			Next
		End Sub

	End Class


End Namespace
