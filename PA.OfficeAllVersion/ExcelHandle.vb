Imports System.Text
Imports System.Reflection
Imports System.Runtime.InteropServices

''' <summary>
''' EXCEL�t�@�C���𑀍삷�邽�߂̋@�\��񋟂���N���X�B
''' </summary>
''' <remarks>
''' <para>���̃N���X�ł́AEXCEL�𑀍삷�邽�߂̃v���p�e�B����у��\�b�h��񋟂��Ă��܂��B</para>
''' <para>
''' <paramref name="Load" />���\�b�h���g�p����ƁAEXCEL�t�@�C����1�u�b�N1�V�[�g��ǂݍ��݁A���̃N���X��DataTable�Ɋi�[���܂��B
''' DataTable�ҏW������<paramref name="Save" />���\�b�h���g�p����ƁAEXCEL�t�@�C���ɕҏW���e���X�V���܂��B
''' </para>
''' <para><font color="red">���̃N���X�͋����ł��B�����N���X�Ƃ̌݊����̂��߂ɑ��݂��Ă��܂��B�V������ <seealso>ExcelReader</seealso>���g�p���ĉ������B</font></para>
''' </remarks>
Public Class ExcelHandle

#Region "public static field"

	''' <summary>
	''' Excel�̍ő�s��
	''' </summary>
	''' <remarks></remarks>
	Public Const MaxRowCount As Integer = 65536

#End Region

#Region "private field"
	Private _dirty As Boolean = False
	Private _filepath As String = ""
	Private _sheetName As String = Nothing

	Private _sheetData(,) As Object
	Private _isChangedData(,) As Boolean
#End Region

#Region "constructor"

    ''' <summary>
    ''' �R���X�g���N�^�B����Excel�I�u�W�F�N�g�ő��삷��EXCEL�̃t�@�C�����w�肵�܂��B
    ''' </summary>
    ''' <param name="filepath">EXCEL�t�@�C���̃t���p�X</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal filepath As String)

        Me._filepath = filepath

    End Sub

#End Region

#Region "public property"

    ''' <summary>
    ''' Excel�t�@�C���̃t���p�X���擾�A�ݒ肵�܂��B
    ''' </summary>
    ''' <value>Excel�t�@�C���̃t���p�X</value>
    ''' <returns>Excel�t�@�C���̃t���p�X</returns>
    ''' <remarks>
    ''' <para>���̃I�u�W�F�N�g�ő��삷��Excel�t�@�C���̃t���p�X���擾�A�ݒ肵�܂��B</para>    
    ''' <para>���̒l��ύX���邱�Ƃɂ��A�ǂݍ��݁A�������ݑΏۂ̃t�@�C����ύX���܂��B</para>    
    ''' </remarks>
    Public ReadOnly Property FilePath() As String
        Get
            Return Me._filepath
        End Get
    End Property

	''' <summary>
	''' Excel�V�[�g�̖��O��ݒ�A�擾���܂��B
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks>
	''' <para>Excel�V�[�g�̖��O��ݒ�A�܂��͎擾���܂��BLoad���\�b�h���s�O�́ANull�iVisual Basic�̏ꍇ��Nothing�j���ݒ肳��Ă��܂��B</para>	
	''' </remarks>
	Public Property SheetName() As String
		Get
			Return Me._sheetName
		End Get
		Set(ByVal value As String)
			Me._sheetName = value
		End Set
	End Property

    ''' <summary>
    ''' Excel�̃Z���̓��e��ݒ�A�擾���܂��B
    ''' </summary>
    ''' <param name="row">�s�ԍ� 1�`</param>
    ''' <param name="col">��ԍ� 1�`</param>
    ''' <value>�Z���ɃZ�b�g����l</value>
    ''' <returns>�Z������擾�����l</returns>
    ''' <remarks></remarks>
    Public Property SheetData(ByVal row As Integer, ByVal col As Integer) As Object
        Get
            Return Me._sheetData(row - 1, col - 1)
        End Get
        Set(ByVal value As Object)
            _sheetData(row - 1, col - 1) = value
        End Set
    End Property

#End Region

#Region "public method"

    ''' <summary>
    ''' Excel�t�@�C���̐V�K�V�[�g���쐬���܂��B
    ''' </summary>
    ''' <param name="sheetIndex"></param>
    ''' <param name="columnCount"></param>
    ''' <param name="rowCount"></param>
    ''' <remarks>
    ''' <para>Excel�t�@�C���̐V�K�V�[�g���쐬���܂��B</para> 
    ''' <para>�ҏW�������eSave()���\�b�h�ɂĕۑ����邱�Ƃ��ł��܂��B</para>       
    ''' </remarks>
    Public Sub InitializeNewSheet( _
          ByVal sheetIndex As Integer _
        , ByVal columnCount As Integer _
        , ByVal rowCount As Integer)
        ' ���� ������ ����
    End Sub

	''' <summary>
	''' �u�b�N�ɓo�^����Ă���V�[�g�̐����擾���܂��B
	''' </summary>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetSheetCount() As Integer

		Dim appl As Object = Nothing
		Dim books As Object = Nothing
		Dim book As Object = Nothing
		Dim sheets As Object = Nothing

		Dim sheetCount As Integer

		Try

			'�x���o�C���f�B���O

			appl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"))
			appl.GetType().InvokeMember("Visible" _
				, BindingFlags.SetProperty _
				, Nothing _
				, appl _
				, New Object() {False})
			books = appl.GetType().InvokeMember("Workbooks" _
				, BindingFlags.GetProperty _
				, Nothing _
				, appl _
				, Nothing)
			book = books.GetType().InvokeMember("Open" _
				, BindingFlags.GetProperty _
				, Nothing _
				, books _
				, New Object() {Me._filepath})
			sheets = book.GetType().InvokeMember("Sheets" _
				, BindingFlags.GetProperty _
				, Nothing _
				, book _
				, Nothing)

			sheetCount = CInt( _
				CType(sheets, Object).GetType().InvokeMember("Count" _
					, BindingFlags.GetProperty _
					, Nothing _
					, sheets _
					, Nothing))

		Catch ex As Exception

			Me._sheetData = Nothing

			'��O�͍ăX���[����
			Throw ex

		Finally

			'�I�������BCOM�I�u�W�F�N�g��S�ĉ������B

			If Not sheets Is Nothing Then
				Marshal.ReleaseComObject(sheets)
			End If

			If Not book Is Nothing Then
				book.GetType().InvokeMember("Close" _
					, BindingFlags.InvokeMethod _
					, Nothing _
					, book _
					, Nothing)
				Marshal.ReleaseComObject(book)
			End If

			If Not books Is Nothing Then
				books.GetType().InvokeMember("Close" _
					, BindingFlags.InvokeMethod _
					, Nothing _
					, books _
					, Nothing)
				Marshal.ReleaseComObject(books)
			End If

			If Not appl Is Nothing Then
				' �A�v���P�[�V�����̏I��
				appl.GetType().InvokeMember("Quit" _
					, BindingFlags.InvokeMethod _
					, Nothing _
					, appl _
					, Nothing)
				Marshal.ReleaseComObject(appl)
			End If

		End Try

		Return sheetCount

	End Function

	''' <summary>
	''' Excel�t�@�C����ǂݍ��݁A�t�@�C���̓��e�����̃I�u�W�F�N�g�ɓW�J���܂��B
	''' </summary>
	''' <param name="sheetIndex">�V�[�g�ԍ�</param>
	''' <param name="columnCount">�ǂݍ��ޗ�</param>
	''' <remarks>
	''' <para>Excel�̃t�@�C����ǂݍ��݂܂��B�l�A�����Ȃǂ����̃I�u�W�F�N�g�̔z��v���p�e�B�ɂĐݒ肵�܂��B</para>
	''' <para>�s���̑S�ẴZ�����󔒂ɂȂ������_�œǂݍ��݂��I�����܂��B�󔒍s�ȍ~�̓��e�͈�ؓǂݍ��܂�Ȃ����Ƃɒ��ӂ��ĉ������B</para>	
	''' </remarks>
	Public Sub Load( _
		ByVal sheetIndex As Integer _
		, ByVal columnCount As Integer)

		Dim appl As Object = Nothing
		Dim books As Object = Nothing
		Dim book As Object = Nothing
		Dim sheets As Object = Nothing
		Dim sheet As Object = Nothing
		Dim listRows As ArrayList = New ArrayList()
		Dim listCells(columnCount - 1) As Object ' listRows�̓��e�ɂȂ�܂�

		Dim listRanges As ArrayList = New ArrayList()

		Try

			'���O�o�C���f�B���O�iDLL��͂̓r���h���j
			appl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"))
			appl.GetType().InvokeMember("Visible" _
				, BindingFlags.SetProperty _
				, Nothing _
				, appl _
				, New Object() {False})
			books = appl.GetType().InvokeMember("Workbooks" _
				, BindingFlags.GetProperty _
				, Nothing _
				, appl _
				, Nothing)
			book = books.GetType().InvokeMember("Open" _
				, BindingFlags.GetProperty _
				, Nothing _
				, books _
				, New Object() {Me._filepath})
			sheets = book.GetType().InvokeMember("Sheets" _
				, BindingFlags.GetProperty _
				, Nothing _
				, book _
				, Nothing)
			sheet = sheets.GetType().InvokeMember("Item" _
				, BindingFlags.GetProperty _
				, Nothing _
				, sheets _
				, New Object() {sheetIndex})

			Dim i As Integer
			Dim j As Integer
			Dim containData As Boolean
			For i = 1 To MaxRowCount Step 1

				' �s���f�[�^���ݒl�̏�����
				containData = False

				' �z��̏�����
				listCells = New Object(columnCount - 1) {}

				' �s�f�[�^�̎擾
				For j = 1 To columnCount Step 1
					Dim cellRange As Object _
						= sheet.GetType().InvokeMember("Range" _
							, BindingFlags.GetProperty _
							, Nothing _
							, sheet _
							, New Object() {getCellSignature(j, i)})
					listCells(j - 1) _
						= cellRange.GetType().InvokeMember("Value" _
							, BindingFlags.GetProperty _
							, Nothing _
							, cellRange _
							, Nothing)
					listRanges.Add(cellRange)	' ReleaseComObject���邽�ߔz��Ɋi�[����
					If Not (listCells(j - 1) Is Nothing) Then
						If Not (CStr(listCells(j - 1)).Trim().Equals(String.Empty)) Then
							' �f�[�^���s���Ɋ܂܂��
							containData = True
						End If
					End If
				Next

				' �s�Ƀf�[�^���܂܂�Ȃ��Ƃ��A���[�v���I������
				If (Not containData) Then
					Exit For
				End If

				' �s�f�[�^�����X�g�ɒǉ�
				listRows.Add(listCells)

			Next

			Me._sheetData = New Object(listRows.Count - 1, columnCount - 1) {}
			Me._isChangedData = New Boolean(listRows.Count - 1, columnCount - 1) {}

			' �擾�f�[�^��񎟌��z��Ɋi�[������
			For x As Integer = 0 To Me._sheetData.GetLength(0) - 1 Step 1
				Dim rowObject() As Object = CType(listRows(x), Object())
				For y As Integer = 0 To Me._sheetData.GetLength(1) - 1 Step 1
					Me._sheetData(x, y) = rowObject(y)
				Next
			Next


			'�z���������
			For xChange As Integer = 0 To Me._isChangedData.GetLength(0) - 1 Step 1
				For yChange As Integer = 0 To Me._isChangedData.GetLength(1) - 1 Step 1
					Me._isChangedData(xChange, yChange) = False
				Next
			Next

			' �V�[�g��
			Me._sheetName = CStr( _
				sheet.GetType().InvokeMember("Name" _
					, BindingFlags.GetProperty _
					, Nothing _
					, sheet _
					, Nothing))

		Catch ex As Exception
			'��O�͍ăX���[����
			Throw ex
		Finally

			For count As Integer = 0 To listRanges.Count - 1
				If Not (listRanges(count) Is Nothing) Then
					Marshal.ReleaseComObject(listRanges(count))
				End If
			Next

			If Not sheet Is Nothing Then
				Marshal.ReleaseComObject(sheet)
			End If

			If Not sheets Is Nothing Then
				Marshal.ReleaseComObject(sheets)
			End If

			If Not book Is Nothing Then
				book.GetType().InvokeMember("Close" _
					, BindingFlags.InvokeMethod _
					, Nothing _
					, book _
					, Nothing)
				Marshal.ReleaseComObject(book)
			End If

			If Not books Is Nothing Then
				books.GetType().InvokeMember("Close" _
					, BindingFlags.InvokeMethod _
					, Nothing _
					, books _
					, Nothing)
				Marshal.ReleaseComObject(books)
			End If

			If Not appl Is Nothing Then
				' �A�v���P�[�V�����̏I��
				appl.GetType().InvokeMember("Quit" _
					, BindingFlags.InvokeMethod _
					, Nothing _
					, appl _
					, Nothing)
				Marshal.ReleaseComObject(appl)
			End If

		End Try

	End Sub

	''' <summary>
	''' Excel�t�@�C����ǂݍ��݁A�t�@�C���̓��e�����̃N���X�ɓW�J���܂��B
	''' </summary>
	''' <param name="sheetIndex">�V�[�g�ԍ�</param>    
	''' <param name="columnCount">�ǂݍ��ޗ�</param>    
	''' <param name="rowCount">�������ޗ�</param>    
	''' <remarks>
	''' <para>Excel�̃t�@�C����ǂݍ��݂܂��B�l�A�����Ȃǂ����̃I�u�W�F�N�g�̔z��v���p�e�B�ɂĐݒ肵�܂��B</para>    
	''' </remarks>
    Public Sub Load( _
        ByVal sheetIndex As Integer _
      , ByVal columnCount As Integer _
      , ByVal rowCount As Integer)

        Dim appl As Object = Nothing
        Dim books As Object = Nothing
        Dim book As Object = Nothing
        Dim sheets As Object = Nothing
        Dim sheet As Object = Nothing
        Dim ranges(rowCount - 1, columnCount - 1) As Object

        Me._sheetData = New Object(rowCount - 1, columnCount - 1) {}
		Me._isChangedData = New Boolean(rowCount - 1, columnCount - 1) {}

		'�z���������
		For x As Integer = 0 To Me._isChangedData.GetLength(0) - 1 Step 1
			For y As Integer = 0 To Me._isChangedData.GetLength(1) - 1 Step 1
				Me._isChangedData(x, y) = False
			Next
		Next

		Try

			'�x���o�C���f�B���O

			appl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"))
			appl.GetType().InvokeMember("Visible" _
				, BindingFlags.SetProperty _
				, Nothing _
				, appl _
				, New Object() {False})
			books = appl.GetType().InvokeMember("Workbooks" _
				, BindingFlags.GetProperty _
				, Nothing _
				, appl _
				, Nothing)
			book = books.GetType().InvokeMember("Open" _
				, BindingFlags.GetProperty _
				, Nothing _
				, books _
				, New Object() {Me._filepath})
			sheets = book.GetType().InvokeMember("Sheets" _
				, BindingFlags.GetProperty _
				, Nothing _
				, book _
				, Nothing)
			sheet = sheets.GetType().InvokeMember("Item" _
				, BindingFlags.GetProperty _
				, Nothing _
				, sheets _
				, New Object() {sheetIndex})

			Dim i As Integer
			Dim j As Integer
			For i = 1 To rowCount
				For j = 1 To columnCount
					ranges(i - 1, j - 1) = sheet.GetType().InvokeMember("Range" _
						, BindingFlags.GetProperty _
						, Nothing _
						, sheet _
						, New Object() {getCellSignature(j, i)})
					_sheetData(i - 1, j - 1) = ranges(i - 1, j - 1).GetType().InvokeMember("Value" _
						, BindingFlags.GetProperty _
						, Nothing _
						, ranges(i - 1, j - 1) _
						, Nothing)
				Next
			Next

			Me._sheetName = CStr( _
				sheet.GetType().InvokeMember("Name" _
					, BindingFlags.GetProperty _
					, Nothing _
					, sheet _
					, Nothing))

		Catch ex As Exception

			Me._sheetData = Nothing

			'��O�͍ăX���[����
			Throw ex

		Finally

			'�I�������BCOM�I�u�W�F�N�g��S�ĉ������B

			Dim i As Integer
			Dim j As Integer

			For i = 1 To rowCount
				For j = 1 To columnCount
					If Not ranges(i - 1, j - 1) Is Nothing Then
						Marshal.ReleaseComObject(ranges(i - 1, j - 1))
					End If
				Next
			Next

			If Not sheet Is Nothing Then
				Marshal.ReleaseComObject(sheet)
			End If

			If Not sheets Is Nothing Then
				Marshal.ReleaseComObject(sheets)
			End If

			If Not book Is Nothing Then
				book.GetType().InvokeMember("Close" _
					, BindingFlags.InvokeMethod _
					, Nothing _
					, book _
					, Nothing)
				Marshal.ReleaseComObject(book)
			End If

			If Not books Is Nothing Then
				books.GetType().InvokeMember("Close" _
					, BindingFlags.InvokeMethod _
					, Nothing _
					, books _
					, Nothing)
				Marshal.ReleaseComObject(books)
			End If

			If Not appl Is Nothing Then
				' �A�v���P�[�V�����̏I��
				appl.GetType().InvokeMember("Quit" _
					, BindingFlags.InvokeMethod _
					, Nothing _
					, appl _
					, Nothing)
				Marshal.ReleaseComObject(appl)
			End If

		End Try

    End Sub

    ''' <summary>
    ''' ���̃I�u�W�F�N�g�̓��e��\����������擾���܂��B
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function ToString() As String

        Dim sb As StringBuilder = New StringBuilder()

        If Me._sheetData Is Nothing Then
            Return ""
        End If

        Dim i As Integer = 0
        Dim j As Integer = 0

        For i = 0 To Me._sheetData.GetLength(0) - 1
            For j = 0 To Me._sheetData.GetLength(1) - 1
                sb.Append("(").Append(i + 1).Append(",").Append(j + 1).Append(")=")
                If Me._sheetData(i, j) Is Nothing Then
                    sb.Append("Undefined ")
                Else
                    sb.Append(Me._sheetData(i, j).ToString()).Append(" ")
                End If
            Next
            sb.Append(vbCrLf)
        Next

        Return sb.ToString()

    End Function

#End Region

#Region "private method"

    ''' <summary>
    ''' �s�Ɨ�̔ԍ�����AEXCEL�Z�����̂��擾���܂��B
    ''' </summary>
    ''' <param name="columnCount"></param>
    ''' <param name="rowCount"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Private Function getCellSignature(ByVal columnCount As Integer, ByVal rowCount As Integer) As String

		Dim sb As StringBuilder = New StringBuilder()
		Return sb.Append(GetColumnSignature(columnCount)).Append(rowCount).ToString()

	End Function

    ''' <summary>
    ''' ��̔ԍ�����AEXCEL�񖼏̂��擾���܂��B
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetColumnSignature(ByVal columnCount As Integer) As String

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

End Class
