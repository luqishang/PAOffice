Imports System.Text
Imports System.Reflection
Imports System.Runtime.InteropServices
Namespace ExcelObjects

	''' <summary>
	''' �G�N�Z���u�b�N�𑀍삷�邽�߂̃N���X
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelBookControl

#Region "Public Const Fields"

#End Region

#Region "Private Const Fields"

		Private Const ApplicationVisible As Boolean = True

#End Region

#Region "Private Fields"

		Private _filePath As String

		Private _sheets As New ExcelSheetObjectCollection

		Private _loadingAreaRow As IList(Of Integer) = New List(Of Integer)

		Private _loadingAreaColumn As IList(Of Integer) = New List(Of Integer)

		Private _isDefineColumns As Boolean = False

#End Region

#Region "Public Constructor"

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <remarks></remarks>
		Public Sub New()

		End Sub

		''' <summary>
		''' �R���X�g���N�^�B�ǂݏ������s�Ȃ��Ώۂ̃t�@�C�������w�肵�܂��B
		''' </summary>
		''' <param name="filePath">�t�@�C���p�X</param>
		''' <remarks></remarks>
		Public Sub New(ByVal filePath As String)

			Me._filePath = filePath

		End Sub

#End Region

#Region "Public Property"

		''' <summary>
		''' �t�@�C���p�X
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>EXCEL�u�b�N�t�@�C���̃t���p�X��ݒ�A�擾���܂��B</remarks>
		Public Property FilePath() As String
			Get
				Return Me._filePath
			End Get
			Set(ByVal value As String)
				Me._filePath = value
			End Set
		End Property

		''' <summary>
		''' EXCEL�V�[�g�̓��e
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Sheets() As ExcelSheetObjectCollection
			Get
				Return Me._sheets
			End Get
		End Property

#End Region

#Region "Public Methods"

		''' <summary>
		''' EXCEL�u�b�N��ǂݍ��݂܂��i�������j
		''' </summary>
		''' <remarks></remarks>
		Public Sub Load()

		End Sub

		''' <summary>
		''' �w��̃t�@�C���p�X�i�t���p�X�j��EXCEL�u�b�N��ǂݍ��݂܂��B
		''' </summary>
		''' <param name="filePath"></param>
		''' <remarks>
		''' �ǂݍ��ݔ͈͂��w�肷�邽�߁A�ǂݍ��݃G���A�����O�ɐݒ肷�� AddLoadingAreaSetting() ���\�b�h�����s���܂��B
		''' �ǂݍ��ݔ͈͖��ݒ莞�A���̃I�u�W�F�N�g�͉����ǂ܂��ɏ������I�����܂��B
		''' </remarks>
		Public Sub LoadFrom(ByVal filePath As String)

			Dim application As Object = Nothing
			Dim books As Object = Nothing
			Dim book As Object = Nothing
			Dim sheets As Object = Nothing
			Dim sheetList As IList(Of Object) = New List(Of Object)

			If Not System.IO.File.Exists(filePath) Then
				Throw New System.IO.FileNotFoundException("�u�b�N�t�@�C����������܂���B", filePath)
			End If

			Try

				application _
					= Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"))

				application.GetType().InvokeMember( _
					"Visible", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {False})
				application.GetType().InvokeMember( _
					"DisplayAlerts", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {False})



				'
				' �u�b�N�̓ǂݍ���
				'
				books _
					= application.GetType().InvokeMember( _
						"Workbooks", _
						BindingFlags.GetProperty, _
						Nothing, _
						application, _
						Nothing)
				book _
					= books.GetType().InvokeMember( _
						"Open", _
						BindingFlags.InvokeMethod, _
						Nothing, _
						books, _
						New Object() {filePath})
				sheets _
					= book.GetType().InvokeMember( _
						"Worksheets", _
						BindingFlags.GetProperty, _
						Nothing, _
						book, _
						Nothing)
				Dim sheetCountMax As Integer _
					= DirectCast( _
						sheets.GetType().InvokeMember( _
							"Count", _
							BindingFlags.GetProperty, _
							Nothing, _
							sheets, _
							Nothing), _
						Integer)


				'
				' �V�[�g�̓ǂݍ���
				' 
				For countSheet As Integer = 1 To Me._loadingAreaRow.Count Step 1

					' �V�[�g���擾����
					sheetList.Add( _
						sheets.GetType().InvokeMember( _
							"Item", _
							BindingFlags.GetProperty, _
							Nothing, _
							sheets, _
							new Object(){countSheet}))


					Dim sheetData As ExcelSheetObject

					If countSheet <= sheetCountMax Then
						' �ΏۃV�[�g��ǂݍ���

						sheetData = Me.GetSheetDataFrom( _
								sheetList(sheetList.Count - 1), _
								Me._loadingAreaRow(countSheet - 1), _
								Me._loadingAreaColumn(countSheet - 1))

					Else

						sheetData = New ExcelSheetObject()

					End If

					Me._sheets.Add(sheetData)

				Next


				application.GetType().InvokeMember( _
					"Visible", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {False})

			Catch ex As Exception

				Console.WriteLine(ex.ToString())

				' ��O�͂����ŏ��������A�ăX���[����
				Throw ex

			Finally

				For Each sheet As Object In sheetList
					If sheet IsNot Nothing Then
						Marshal.ReleaseComObject(sheet)
						sheet = Nothing
					End If
				Next
				If sheets IsNot Nothing Then
					Marshal.ReleaseComObject(sheets)
					sheets = Nothing
				End If
				If book IsNot Nothing Then
					book.GetType().InvokeMember( _
						"Close", _
						BindingFlags.InvokeMethod, _
						Nothing, _
						book, _
						Nothing)
					Marshal.ReleaseComObject(book)
					book = Nothing
				End If
				If books IsNot Nothing Then
					books.GetType().InvokeMember( _
						"Close" _
						, BindingFlags.InvokeMethod _
						, Nothing _
						, books _
						, Nothing)
					Marshal.ReleaseComObject(books)
					books = Nothing
				End If
				If application IsNot Nothing Then
					application.GetType().InvokeMember( _
						"Quit", _
						BindingFlags.InvokeMethod, _
						Nothing, _
						application, _
						Nothing)
					Marshal.ReleaseComObject(application)
					application = Nothing
				End If

			End Try

		End Sub

		''' <summary>
		''' �V�[�g�̓ǂݍ��ݔ͈͂��w�肵�܂��B
		''' </summary>
		''' <param name="row"></param>
		''' <param name="column"></param>
		''' <remarks></remarks>
		Public Sub AddLoadingAreaSetting( _
			ByVal row As Integer, _
			ByVal column As Integer)

			If row < 0 Then
				Throw New ArgumentOutOfRangeException("row", "row �ɕ������ݒ肳��Ă��܂��B")
			End If
			If column < 0 Then
				Throw New ArgumentOutOfRangeException("column", "column �ɕ������ݒ肳��Ă��܂��B")
			End If

			Me._loadingAreaRow.Add(row)
			Me._loadingAreaColumn.Add(column)

		End Sub

		''' <summary>
		''' �V�[�g�̍s�ǂݍ��ݔ͈͂��擾���܂��B
		''' </summary>
		''' <param name="sheetIndex"></param>
		''' <remarks></remarks>
		Public Function GetLoadingAreaSettingRow(ByVal sheetIndex As Integer) As Integer

			If Me._loadingAreaRow.Count <= sheetIndex Then
				Return 0
			End If
			Return Me._loadingAreaRow(sheetIndex)

		End Function

		''' <summary>
		''' �V�[�g�̗�ǂݍ��ݔ͈͂��擾���܂��B
		''' </summary>
		''' <param name="sheetIndex"></param>
		''' <remarks></remarks>
		Public Function GetLoadingAreaSettingColumn(ByVal sheetIndex As Integer) As Integer

			If Me._loadingAreaColumn.Count <= sheetIndex Then
				Return 0
			End If
			Return Me._loadingAreaColumn(sheetIndex)

		End Function

		''' <summary>
		''' �V�[�g�̓ǂݍ��ݔ͈͂��擾���܂��B
		''' </summary>
		''' <param name="sheetIndex"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function GetLoadingAreaSettingSignature(ByVal sheetIndex As Integer) As String

			If Me._loadingAreaRow.Count <= sheetIndex Then
				Return ""
			End If
			If Me._loadingAreaColumn.Count <= sheetIndex Then
				Return ""
			End If

			Return ExcelBookControl.GetCellSignature( _
				Me._loadingAreaColumn(sheetIndex), _
				Me._loadingAreaRow(sheetIndex))

		End Function

		''' <summary>
		''' EXCEL�u�b�N��ۑ����܂��i�������j
		''' </summary>
		''' <remarks></remarks>
		Public Sub Save()

		End Sub

		''' <summary>
		''' �w��̃t�@�C���p�X��EXCEL�u�b�N��ۑ����܂��B�����t�@�C��������ꍇ�́A�㏑���ҏW���܂��B
		''' </summary>
		''' <param name="filePath"></param>
		''' <remarks></remarks>
		Public Sub SaveAs(ByVal filePath As String)

			Dim application As Object = Nothing
			Dim books As Object = Nothing
			Dim book As Object = Nothing
			Dim sheets As Object = Nothing
			Dim sheetList As IList(Of Object) = New List(Of Object)

			Dim backupSheetsInNewWorkBook As Double = 3

			Try

				application _
					= Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"))

				application.GetType().InvokeMember( _
					"Visible", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {False})
				application.GetType().InvokeMember( _
					"DisplayAlerts", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {False})
				application.GetType().InvokeMember( _
					"AlertBeforeOverwriting", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {False})


				' �����V�[�g���� 1 �ŌŒ�
				backupSheetsInNewWorkBook _
					= DirectCast( _
						application.GetType().InvokeMember( _
							"SheetsInNewWorkbook", _
							BindingFlags.GetProperty, _
							Nothing, _
							application, _
							Nothing), _
						Double)
				application.GetType().InvokeMember( _
					"SheetsInNewWorkbook", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {1})


				'
				' �u�b�N�̐V�K�쐬�E�J��
				'
				books _
					= application.GetType().InvokeMember( _
						"Workbooks", _
						BindingFlags.GetProperty, _
						Nothing, _
						application, _
						Nothing)

				If System.IO.File.Exists(filePath) Then
					' �Y���t�@�C������̂Ƃ��͊J��

					book = books.GetType().InvokeMember( _
							"Open", _
							BindingFlags.InvokeMethod, _
							Nothing, _
							books, _
							New Object() {filePath})
					sheets = book.GetType().InvokeMember( _
							"Worksheets", _
							BindingFlags.GetProperty, _
							Nothing, _
							book, _
							Nothing)
					Dim sheetsCountMax As Integer _
						= DirectCast( _
							sheets.GetType().InvokeMember( _
								"Count", _
								BindingFlags.GetProperty, _
								Nothing, _
								sheets, _
								Nothing), _
							Integer)


					For countSheetForExists As Integer = 1 To Me._sheets.Count Step 1

						If countSheetForExists <= sheetsCountMax Then
							' �����t�@�C���ɃV�[�g������ꍇ
							sheetList.Add( _
								sheets.GetType().InvokeMember( _
									"Item", _
									BindingFlags.GetProperty, _
									Nothing, _
									sheets, _
									New Object() {countSheetForExists}))
						Else
							' �����t�@�C���ɃV�[�g���Ȃ��ꍇ
							sheetList.Add( _
								sheets.GetType().InvokeMember( _
									"Add", _
									BindingFlags.InvokeMethod, _
									Nothing, _
									sheets, _
									Nothing))
						End If

					Next

				Else
					' �Y���t�@�C���Ȃ��̂Ƃ��͐V�K�쐬����

					book = books.GetType().InvokeMember( _
							"Add", _
							BindingFlags.InvokeMethod, _
							Nothing, _
							books, _
							Nothing)

					sheets = book.GetType().InvokeMember( _
							"Worksheets", _
							BindingFlags.GetProperty, _
							Nothing, _
							book, _
							Nothing)

					' ��ڂ̃V�[�g�����X�g�ɒǉ�
                    sheetList.Add( _
                     sheets.GetType().InvokeMember( _
                      "Item", _
                      BindingFlags.GetProperty, _
                      Nothing, _
                      sheets, _
                      New Object() {1}))

                    ' �V�[�g�̖��O��ݒ�
                    Me.SetNameSheetAt(sheetList(sheetList.Count - 1), Me._sheets(0).Name)

					' ��ڈȍ~�̃V�[�g��ǉ�
					For countAdditionSheet As Integer = 2 To Me._sheets.Count Step 1

                        sheetList.Add( _
                         sheets.GetType().InvokeMember( _
                          "Add", _
                          BindingFlags.InvokeMethod, _
                          Nothing, _
                          sheets, _
                          Nothing))

                        ' �V�[�g�̖��O��ݒ�
                        Me.SetNameSheetAt(sheetList(sheetList.Count - 1), Me._sheets(countAdditionSheet - 1).Name)

					Next

				End If

				' �����V�[�g�������ɖ߂�
				application.GetType().InvokeMember( _
					"SheetsInNewWorkbook", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {backupSheetsInNewWorkBook})


				'
				' �V�[�g�̏�������
				' 
				For countSheet As Integer = 1 To Me._sheets.Count Step 1

					' �ΏۃV�[�g����������
					Me.SetDataSheetAt( _
							sheetList(countSheet - 1), _
							Me._sheets(countSheet - 1))

					' �ΏۃV�[�g�̃O���t��ҏW����
					Me.SetChartSheetAt( _
							sheets, _
							sheetList(countSheet - 1), _
							Me._sheets(countSheet - 1))

				Next


				'
				' �V�[�g�̔�\��
				'
				For countHiddenSheet As Integer = Me._sheets.Count To 1 Step -1

					' �ΏۃV�[�g�̔�\��
					Me.SetHiddenSheetAt( _
						sheetList(countHiddenSheet - 1), _
						Me._sheets(countHiddenSheet - 1))

				Next


				book.GetType().InvokeMember( _
					"SaveAs", _
					BindingFlags.InvokeMethod, _
					Nothing, _
					book, _
					New Object() {filePath})
				'application.GetType().InvokeMember( _
				'	"Visible", _
				'	BindingFlags.SetProperty, _
				'	Nothing, _
				'	application, _
				'	New Object() {False})

			Catch ex As Exception

				Console.WriteLine(ex.ToString())

				' ��O�͂����ŏ��������A�ăX���[����
				Throw ex

			Finally

				For Each sheet As Object In sheetList
					If sheet IsNot Nothing Then
						Marshal.ReleaseComObject(sheet)
						sheet = Nothing
					End If
				Next
				If sheets IsNot Nothing Then
					Marshal.ReleaseComObject(sheets)
					sheets = Nothing
				End If
				If book IsNot Nothing Then
					book.GetType().InvokeMember( _
						"Close", _
						BindingFlags.InvokeMethod, _
						Nothing, _
						book, _
						Nothing)
					Marshal.ReleaseComObject(book)
					book = Nothing
				End If
				If books IsNot Nothing Then
					books.GetType().InvokeMember( _
						"Close" _
						, BindingFlags.InvokeMethod _
						, Nothing _
						, books _
						, Nothing)
					Marshal.ReleaseComObject(books)
					books = Nothing
				End If
				If application IsNot Nothing Then
					application.GetType().InvokeMember( _
						"Quit", _
						BindingFlags.InvokeMethod, _
						Nothing, _
						application, _
						Nothing)
					Marshal.ReleaseComObject(application)
					application = Nothing
				End If


			End Try

		End Sub

		''' <summary>
		''' EXCEL�A�v���P�[�V�������N�����AEXCEL�u�b�N�̓��e��\�����܂��B
		''' </summary>
		''' <remarks></remarks>
		Public Sub Show()

			Me.Show(1)

		End Sub

		''' <summary>
		''' EXCEL�A�v���P�[�V�������N�����A�����Ŏw�肵���V�[�g�̓��e��\�����܂��B�i�������j
		''' </summary>
		''' <param name="sheetName"></param>
		''' <remarks></remarks>
		Public Sub Show(ByVal sheetName As String)

		End Sub

		''' <summary>
		''' EXCEL�A�v���P�[�V�������N�����A�����Ŏw�肵���V�[�g�̓��e��\�����܂��B
		''' </summary>
		''' <param name="sheetIndex"></param>
		''' <remarks></remarks>
		Public Sub Show(ByVal sheetIndex As Integer)

			Dim application As Object = Nothing
			Dim books As Object = Nothing
			Dim book As Object = Nothing
			Dim sheets As Object = Nothing
			Dim sheetList As IList(Of Object) = New List(Of Object)
			Dim chartList As IList(Of Object) = New List(Of Object)

			Dim backupSheetsInNewWorkBook As Double = 3

			Try

				application _
					= Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"))

				application.GetType().InvokeMember( _
					"Visible", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {False})
				application.GetType().InvokeMember( _
					"DisplayAlerts", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {False})


				'
				' �����ݒ�
				'

				' �V�K�u�b�N�쐬���̃V�[�g���� 1 �ŌŒ�
				backupSheetsInNewWorkBook _
					= DirectCast( _
						application.GetType().InvokeMember( _
							"SheetsInNewWorkbook", _
							BindingFlags.GetProperty, _
							Nothing, _
							application, _
							Nothing), _
						Double)
				application.GetType().InvokeMember( _
					"SheetsInNewWorkbook", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {1})


				'
				' �V�����u�b�N�̍쐬
				'
				books _
					= application.GetType().InvokeMember( _
						"Workbooks", _
						BindingFlags.GetProperty, _
						Nothing, _
						application, _
						Nothing)
				book = books.GetType().InvokeMember( _
						"Add", _
						BindingFlags.InvokeMethod, _
						Nothing, _
						books, _
						Nothing)
				sheets = book.GetType().InvokeMember( _
						"Worksheets", _
						BindingFlags.GetProperty, _
						Nothing, _
						book, _
						Nothing)


				'
				' �V�[�g�̒ǉ�
				' 
				For countAdditionSheet As Integer = 2 To Me._sheets.Count Step 1

					sheets.GetType().InvokeMember( _
						"Add", _
						BindingFlags.InvokeMethod, _
						Nothing, _
						sheets, _
						Nothing)

				Next


				'
				' �V�[�g���̏���
				'
				For countSheet As Integer = 1 To Me._sheets.Count Step 1

					sheetList.Add( _
						sheets.GetType().InvokeMember( _
							"Item", _
							BindingFlags.GetProperty, _
							Nothing, _
							sheets, _
							New Object() {countSheet}))

					' �V�[�g�̖��O��ݒ�
					Me.SetNameSheetAt(sheetList(sheetList.Count - 1), Me._sheets(countSheet - 1).Name)

					' �f�[�^�̗�������
					Me.SetDataSheetAt(sheetList(sheetList.Count - 1), Me._sheets(countSheet - 1))

					' �ΏۃV�[�g�̃O���t�ҏW
					Me.SetChartSheetAt( _
							sheets, _
							sheetList(sheetList.Count - 1), _
							Me._sheets(countSheet - 1))

				Next


				'
				' �V�[�g�̔�\��
				'
				For countHiddenSheet As Integer = Me._sheets.Count To 1 Step -1

					' �ΏۃV�[�g�̔�\��
					Me.SetHiddenSheetAt( _
						sheetList(countHiddenSheet - 1), _
						Me._sheets(countHiddenSheet - 1))

				Next

				' �����V�[�g�������ɖ߂�
				application.GetType().InvokeMember( _
					"SheetsInNewWorkbook", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {backupSheetsInNewWorkBook})
				application.GetType().InvokeMember( _
					"Visible", _
					BindingFlags.SetProperty, _
					Nothing, _
					application, _
					New Object() {True})

			Catch ex As Exception

				For Each chart As Object In chartList
					If chart IsNot Nothing Then
						Marshal.ReleaseComObject(chart)
						chart = Nothing
					End If
				Next
				If sheets IsNot Nothing Then
					Marshal.ReleaseComObject(sheets)
					sheets = Nothing
				End If
				For Each sheet As Object In sheetList
					If sheet IsNot Nothing Then
						Marshal.ReleaseComObject(sheet)
						sheet = Nothing
					End If
				Next
				If book IsNot Nothing Then
					book.GetType().InvokeMember( _
						"Close", _
						BindingFlags.InvokeMethod, _
						Nothing, _
						book, _
						Nothing)
					Marshal.ReleaseComObject(book)
					book = Nothing
				End If
				If books IsNot Nothing Then
					books.GetType().InvokeMember( _
						"Close" _
						, BindingFlags.InvokeMethod _
						, Nothing _
						, books _
						, Nothing)
					Marshal.ReleaseComObject(books)
					books = Nothing
				End If
				If application IsNot Nothing Then
					application.GetType().InvokeMember( _
						"Quit", _
						BindingFlags.InvokeMethod, _
						Nothing, _
						application, _
						Nothing)
					Marshal.ReleaseComObject(application)
					application = Nothing
				End If

				' ��O�͂����ŏ��������A�ăX���[����
				Throw ex

			End Try

		End Sub

		''' <summary>
		''' EXCEL�u�b�N��ǂݍ��݁A�A�v���P�[�V�������N�����ĕ\�����܂��B�i�������j
		''' </summary>
		''' <param name="filePath"></param>
		''' <remarks></remarks>
		Public Sub LoadAndShow(ByVal filePath As String)

		End Sub

		''' <summary>
		''' EXCEL�u�b�N��ǂݍ��݁A�A�v���P�[�V�������N�����Ďw��̃V�[�g��\�����܂��B�i�������j
		''' </summary>
		''' <param name="filePath"></param>
		''' <param name="sheetIndex"></param>
		''' <remarks></remarks>
		Public Sub LoadAndShow(ByVal filePath As String, ByVal sheetIndex As Integer)

		End Sub

		''' <summary>
		''' EXCEL�u�b�N��ǂݍ��݁A�A�v���P�[�V�������N�����Ďw��̃V�[�g��\�����܂��B�i�������j
		''' </summary>
		''' <param name="filePath"></param>
		''' <param name="sheetName"></param>
		''' <remarks></remarks>
		Public Sub LoadAndShow(ByVal filePath As String, ByVal sheetName As String)

		End Sub

#End Region

#Region "Friend Methods"

		''' <summary>
		''' �s�Ɨ�̔ԍ�����AEXCEL�Z�����̂��擾���܂��B
		''' </summary>
		''' <param name="columnCount"></param>
		''' <param name="rowCount"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Shared Function GetCellSignature( _
			ByVal columnCount As Integer, _
			ByVal rowCount As Integer) _
			As String

			Dim sb As StringBuilder = New StringBuilder()
			Return sb.Append(GetColumnSignature(columnCount)).Append(rowCount).ToString()

		End Function


		''' <summary>
		''' ��̔ԍ�����AEXCEL�񖼏̂��擾���܂��B
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Shared Function GetColumnSignature( _
			ByVal columnCount As Integer) _
			As String

			Dim sb As StringBuilder = New StringBuilder()

			Dim first As Char = CType("", Char)
			Dim second As Char = CType("", Char)

			If 256 < columnCount Then
				Throw New ArgumentOutOfRangeException("columnCount �ɓ��͂���Ă���񐔂��AExcel���e�񐔂𒴉߂��Ă��܂��B")
			End If

			If 26 < columnCount Then
				' 26��ȏ�̏ꍇ�A�񖼂�2����
				first = Chr(CInt(Math.Truncate((columnCount - 1) / 26) + Asc("A") - 1))
				sb.Append(first)

				columnCount = columnCount Mod 26
				If columnCount = 0 Then
					columnCount = 26
				End If
			End If

			second = Chr(columnCount + Asc("A") - 1)
			sb.Append(second)

			Return sb.ToString()

		End Function

#End Region

#Region "Private Methods"

		''' <summary>
		''' EXCEL�V�[�g�ɑ΂����O���Z�b�g����
		''' </summary>
		''' <param name="sheet"></param>
		''' <param name="name"></param>
		''' <remarks></remarks>
		Private Sub SetNameSheetAt( _
			ByVal sheet As Object, _
			ByVal name As String)

			If name IsNot Nothing AndAlso name <> "" Then

				sheet.GetType().InvokeMember( _
					"Name", _
					BindingFlags.SetProperty, _
					Nothing, _
					sheet, _
					New Object() {name})

			End If

		End Sub

		''' <summary>
		''' EXCEL�V�[�g���\���ɂ���
		''' </summary>
		''' <param name="sheet"></param>
		''' <param name="dataSource"></param>
		''' <remarks></remarks>
		Private Sub SetHiddenSheetAt( _
			ByVal sheet As Object, _
			ByVal dataSource As ExcelSheetObject)

			If dataSource.Visible = False Then

				'Dim visibilityObject As Object _
				'	= Activator.CreateInstance(Type.GetTypeFromProgID("Excel.XlSheetVisibility"))
				''Dim visibilityObject As Object _
				''	= Activator.CreateInstance( _
				''		"Interop.Excel", _
				''		"XlSheetVisibility")

				'Dim sheetHiddenValue As Object _
				'	= visibilityObject.GetType().InvokeMember( _
				'		"xlSheetHidden", _
				'		BindingFlags.GetField, _
				'		Nothing, _
				'		visibilityObject, _
				'		Nothing)

				'sheet.GetType().InvokeMember( _
				'	"Visible", _
				'	BindingFlags.SetProperty, _
				'	Nothing, _
				'	sheet, _
				'	New Object() {sheetHiddenValue})
				sheet.GetType().InvokeMember( _
					"Visible", _
					BindingFlags.SetProperty, _
					Nothing, _
					sheet, _
					New Object() {0})	' ������ �����̒萔���ǂ��ɂ���肽�� ������

				'sheet.Visible = XlSheetVisibility.xlSheetHidden

			End If

		End Sub

		''' <summary>
		''' EXCEL�V�[�g�ɑ΂��f�[�^���Z�b�g����
		''' </summary>
		''' <param name="sheet">�f�[�^���Z�b�g����EXCEL�V�[�g�I�u�W�F�N�g</param>		
		''' <param name="dataSource">�Z�b�g����f�[�^���i�[���Ă���I�u�W�F�N�g</param>		
		''' <remarks></remarks>
		Private Sub SetDataSheetAt( _
			ByVal sheet As Object, _
			ByVal dataSource As ExcelSheetObject)

			Dim rangeList As IList(Of Object) = New List(Of Object)

			Try

				For countRow As Integer = 1 To dataSource.Rows.Count Step 1

					Dim row As ExcelRowObject = dataSource.Rows(countRow - 1)

					For countColumn As Integer = 1 To row.Cells.Count Step 1

						Dim signature As String _
							= ExcelBookControl.GetCellSignature(countColumn, countRow)

						rangeList.Add( _
							sheet.GetType().InvokeMember( _
								"Range", _
								BindingFlags.GetProperty, _
								Nothing, _
								sheet, _
								New Object() {signature}))

						If row.Cells(countColumn - 1).Changed Then

							rangeList(rangeList.Count - 1).GetType().InvokeMember( _
								"Value", _
								BindingFlags.SetProperty, _
								Nothing, _
								rangeList(rangeList.Count - 1), _
								New Object() {row.Cells(countColumn - 1).Value})

						End If

					Next

					row = Nothing

				Next

			Finally
				' COM�I�u�W�F�N�g���

				For Each range As Object In rangeList
					If range IsNot Nothing Then
						Marshal.ReleaseComObject(range)
						range = Nothing
					End If
				Next

			End Try

		End Sub

		''' <summary>
		''' EXCEL�V�[�g�ɑ΂��O���t���Z�b�g����
		''' </summary>
		''' <param name="sheet"></param>
		''' <remarks></remarks>
		Private Sub SetChartSheetAt( _
			ByVal sheets As Object, _
			ByVal sheet As Object, _
			ByVal dataSource As ExcelSheetObject)

			Dim comCharts As Object _
				= sheet.GetType().InvokeMember( _
					"ChartObjects", _
					BindingFlags.InvokeMethod, _
					Nothing, _
					sheet, _
					Nothing)
			Dim comChartList As IList(Of Object) = New List(Of Object)
			Dim comChildChartList As IList(Of Object) = New List(Of Object)
			Dim comChartsCount As Integer _
				= DirectCast( _
					comCharts.GetType().InvokeMember( _
						"Count", _
						BindingFlags.GetProperty, _
						Nothing, _
						comCharts, _
						Nothing), _
					Integer)

			Dim comRangeList As IList(Of Object) = New List(Of Object)

			Dim comDataSourceSheetList As IList(Of Object) = New List(Of Object)




			Try

				For count As Integer = 1 To dataSource.Charts.Count Step 1

					Dim chart As ExcelChartObject = dataSource.Charts(count - 1)

					If comChartsCount < count Then

						' �o�^����Ă���O���t�����ۂ�菭�Ȃ��Ƃ��A
						' �O���t��ǉ�����
						comChartList.Add( _
							comCharts.GetType().InvokeMember( _
								"Add", _
								BindingFlags.InvokeMethod, _
								Nothing, _
								comCharts, _
								New Object() {0, 0, 400, 300}))

					Else

						' �o�^����Ă���O���t���̂Ƃ��́A�����O���t��ҏW����
						comChartList.Add( _
							comCharts.GetType().InvokeMember( _
								"Item", _
								BindingFlags.GetProperty, _
								Nothing, _
								comCharts, _
								New Object() {count}))

					End If

					' �O���t�I�u�W�F�N�g�̎擾
					comChildChartList.Add( _
						comChartList(comChartList.Count - 1).GetType().InvokeMember( _
							"Chart", _
							BindingFlags.GetProperty, _
							Nothing, _
							comChartList(comChartList.Count - 1), _
							Nothing))

					'�O���t�̎�ސݒ�()
					'�i�����I�ɂ͕����̎�ނɑΉ�����j
					'comChildChartList(comChildChartList.Count - 1).GetType().InvokeMember( _
					'	"ChartType", _
					'	BindingFlags.SetProperty, _
					'	Nothing, _
					'	comChildChartList(comChildChartList.Count - 1), _
					'	New Object() {/*chartType*/})


					' ���f�[�^�͈͂̐ݒ�
					comDataSourceSheetList.Add( _
						sheets.GetType().InvokeMember( _
							"Item", _
							BindingFlags.GetProperty, _
							Nothing, _
							sheets, _
							New Object() {chart.DataSourceSheetIndex + 1}))

					Dim rangeString As String _
						= New StringBuilder( _
							).Append( _
								ExcelBookControl.GetCellSignature( _
									chart.DataSourceStartColumnIndex + 1, _
									chart.DataSourceStartRowIndex + 1) _
							).Append( _
								":" _
							).Append( _
								ExcelBookControl.GetCellSignature( _
									chart.DataSourceEndColumnIndex + 1, _
									chart.DataSourceEndRowIndex + 1) _
							).ToString()

					comRangeList.Add( _
						comDataSourceSheetList(comDataSourceSheetList.Count - 1).GetType().InvokeMember( _
							"Range", _
							BindingFlags.GetProperty, _
							Nothing, _
							comDataSourceSheetList(comDataSourceSheetList.Count - 1), _
							New Object() {rangeString}))


					comChildChartList(comChildChartList.Count - 1).GetType().InvokeMember( _
						"SetSourceData", _
						BindingFlags.InvokeMethod, _
						Nothing, _
						comChildChartList(comChildChartList.Count - 1), _
						New Object() {comRangeList(comRangeList.Count - 1)})

				Next

			Finally

				' COM�I�u�W�F�N�g���
				For Each comDataSourceSheet As Object In comDataSourceSheetList
					If comDataSourceSheet IsNot Nothing Then
						Marshal.ReleaseComObject(comDataSourceSheet)
						comDataSourceSheet = Nothing
					End If
				Next

				For Each comRange As Object In comRangeList
					If comRange IsNot Nothing Then
						Marshal.ReleaseComObject(comRange)
						comRange = Nothing
					End If
				Next

				For Each comChildChart As Object In comChildChartList
					If comChildChart IsNot Nothing Then
						Marshal.ReleaseComObject(comChildChart)
						comChildChart = Nothing
					End If
				Next

				For Each comChart As Object In comChartList
					If comChart IsNot Nothing Then
						Marshal.ReleaseComObject(comChart)
						comChart = Nothing
					End If
				Next

				If comCharts IsNot Nothing Then

					Marshal.ReleaseComObject(comCharts)
					comCharts = Nothing

				End If

			End Try

		End Sub

		''' <summary>
		''' EXCEL�V�[�g����f�[�^���擾����
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function GetSheetDataFrom( _
			ByVal sheet As Object, _
			ByVal rowUpper As Integer, _
			ByVal columnUpper As Integer) As ExcelSheetObject

			Dim rangeList As IList(Of Object) = New List(Of Object)

			Dim returnData As New ExcelSheetObject

			Try

				For countRow As Integer = 1 To rowUpper Step 1

					Dim rowData As New ExcelRowObject

					For countColumn As Integer = 1 To columnUpper Step 1

						Dim cellData As New ExcelCellObject

						Dim sig As String _
							= ExcelBookControl.GetCellSignature(countColumn, countRow)
						rangeList.Add( _
							sheet.GetType().InvokeMember( _
								"Range", _
								BindingFlags.GetProperty, _
								Nothing, _
								sheet, _
								New Object() {sig}))
						cellData.Value _
							= rangeList(rangeList.Count - 1).GetType().InvokeMember( _
								"Value", _
								BindingFlags.GetProperty, _
								Nothing, _
								rangeList(rangeList.Count - 1), _
								Nothing)
						cellData.SetOldValue = cellData.Value

						rowData.Cells.Add(cellData)

					Next

					returnData.Rows.Add(rowData)

				Next

			Catch ex As Exception

				Console.WriteLine(ex.ToString())

				Throw ex

			Finally

				For Each range As Object In rangeList
					If range IsNot Nothing Then
						Marshal.ReleaseComObject(range)
						range = Nothing
					End If
				Next

			End Try

			Return returnData

        End Function

#End Region

    End Class

End Namespace
