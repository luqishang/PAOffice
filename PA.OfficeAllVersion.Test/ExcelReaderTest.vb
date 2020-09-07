Imports NUnit.Framework

Imports PA.Office

<TestFixture()> _
Public Class ExcelReaderTest

	Private filepath As String = Nothing

	''' <summary>
	''' コンストラクタ
	''' </summary>
	''' <remarks>
	''' <para>テスト用Excelシートファイルのパスを設定します。</para>	
	''' </remarks>
	Public Sub New()

		filepath = System.AppDomain.CurrentDomain.BaseDirectory & "\test.xls"

	End Sub

	''' <summary>
	''' 
	''' </summary>
	''' <remarks></remarks>
	<Test()> _
	Public Sub Excelファイル読み込み()

		Dim target As ExcelReader = New ExcelReader(filepath)
		target.Load(1, 4, 3)

		Console.WriteLine("内容：" + vbCrLf + target.ToString())
		Console.WriteLine("(1,1)=" + target.SheetData(1, 1))

		Assert.AreEqual(target.SheetData(1, 1), "aaa", "セルの内容が等価か")
		Assert.AreEqual(target.SheetData(3, 3), Nothing, "セルの内容が等価か（空白セル）")
		Assert.AreEqual(target.SheetName, "TestSheet", "シート名が等価か")

		target = Nothing

	End Sub

	<Test()> _
	Public Sub Excelファイル読み込み２()

		Dim target As ExcelReader = New ExcelReader(filepath)
		target.Load(1, 4)

		Console.WriteLine("内容：" + vbCrLf + target.ToString())
		Console.WriteLine("(1,1)=" + target.SheetData(1, 1))

		Assert.AreEqual(target.SheetData(1, 1), "aaa", "セルの内容が等価か")
		Assert.AreEqual(target.SheetName, "TestSheet", "シート名が等価か")

		target = Nothing

	End Sub

	''' <summary>
	''' 
	''' </summary>
	''' <remarks></remarks>
	<Test()> _
	Public Sub Excelシート数取得()

		Dim target As ExcelReader = New ExcelReader(Me.filepath)

		Dim count As Integer = target.GetSheetCount()
		Assert.AreEqual(count, 3, "シート数が正しく取得できるか？")

		target = Nothing

	End Sub

	<Test()> _
	Public Sub Excel行数取得()

		Dim target1 As New ExcelReader(Me.filepath)
		target1.Load(1, 3, 3)

		Assert.AreEqual(target1.EndRowIndex, 3, "行数（固定）")
		Assert.AreEqual(target1.EndColumnIndex, 3, "列数（固定）")


		Dim target2 As New ExcelReader(Me.filepath)
		target2.Load(1, 3)

		Assert.AreEqual(target2.EndRowIndex, 1, "行数（可変）")
		Assert.AreEqual(target2.EndColumnIndex, 3, "列数（可変）")

	End Sub

End Class
