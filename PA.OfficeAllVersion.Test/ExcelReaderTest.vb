Imports NUnit.Framework

Imports PA.Office

<TestFixture()> _
Public Class ExcelReaderTest

	Private filepath As String = Nothing

	''' <summary>
	''' �R���X�g���N�^
	''' </summary>
	''' <remarks>
	''' <para>�e�X�g�pExcel�V�[�g�t�@�C���̃p�X��ݒ肵�܂��B</para>	
	''' </remarks>
	Public Sub New()

		filepath = System.AppDomain.CurrentDomain.BaseDirectory & "\test.xls"

	End Sub

	''' <summary>
	''' 
	''' </summary>
	''' <remarks></remarks>
	<Test()> _
	Public Sub Excel�t�@�C���ǂݍ���()

		Dim target As ExcelReader = New ExcelReader(filepath)
		target.Load(1, 4, 3)

		Console.WriteLine("���e�F" + vbCrLf + target.ToString())
		Console.WriteLine("(1,1)=" + target.SheetData(1, 1))

		Assert.AreEqual(target.SheetData(1, 1), "aaa", "�Z���̓��e��������")
		Assert.AreEqual(target.SheetData(3, 3), Nothing, "�Z���̓��e���������i�󔒃Z���j")
		Assert.AreEqual(target.SheetName, "TestSheet", "�V�[�g����������")

		target = Nothing

	End Sub

	<Test()> _
	Public Sub Excel�t�@�C���ǂݍ��݂Q()

		Dim target As ExcelReader = New ExcelReader(filepath)
		target.Load(1, 4)

		Console.WriteLine("���e�F" + vbCrLf + target.ToString())
		Console.WriteLine("(1,1)=" + target.SheetData(1, 1))

		Assert.AreEqual(target.SheetData(1, 1), "aaa", "�Z���̓��e��������")
		Assert.AreEqual(target.SheetName, "TestSheet", "�V�[�g����������")

		target = Nothing

	End Sub

	''' <summary>
	''' 
	''' </summary>
	''' <remarks></remarks>
	<Test()> _
	Public Sub Excel�V�[�g���擾()

		Dim target As ExcelReader = New ExcelReader(Me.filepath)

		Dim count As Integer = target.GetSheetCount()
		Assert.AreEqual(count, 3, "�V�[�g�����������擾�ł��邩�H")

		target = Nothing

	End Sub

	<Test()> _
	Public Sub Excel�s���擾()

		Dim target1 As New ExcelReader(Me.filepath)
		target1.Load(1, 3, 3)

		Assert.AreEqual(target1.EndRowIndex, 3, "�s���i�Œ�j")
		Assert.AreEqual(target1.EndColumnIndex, 3, "�񐔁i�Œ�j")


		Dim target2 As New ExcelReader(Me.filepath)
		target2.Load(1, 3)

		Assert.AreEqual(target2.EndRowIndex, 1, "�s���i�ρj")
		Assert.AreEqual(target2.EndColumnIndex, 3, "�񐔁i�ρj")

	End Sub

End Class
