Imports NUnit.Framework

Imports PA.Office

<TestFixture()> _
Public Class ExcelHandleTest

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

		Dim target As ExcelHandle = New ExcelHandle(filepath)
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

		Dim target As ExcelHandle = New ExcelHandle(filepath)
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

		Dim target As ExcelHandle = New ExcelHandle(Me.filepath)

		Dim count As Integer = target.GetSheetCount()
		Assert.AreEqual(count, 3, "シート数が正しく取得できるか？")

		target = Nothing

	End Sub

End Class
