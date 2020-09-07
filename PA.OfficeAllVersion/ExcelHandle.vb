Imports System.Text
Imports System.Reflection
Imports System.Runtime.InteropServices

''' <summary>
''' EXCELファイルを操作するための機能を提供するクラス。
''' </summary>
''' <remarks>
''' <para>このクラスでは、EXCELを操作するためのプロパティおよびメソッドを提供しています。</para>
''' <para>
''' <paramref name="Load" />メソッドを使用すると、EXCELファイルの1ブック1シートを読み込み、このクラスのDataTableに格納します。
''' DataTable編集をして<paramref name="Save" />メソッドを使用すると、EXCELファイルに編集内容を更新します。
''' </para>
''' <para><font color="red">このクラスは旧式です。既存クラスとの互換性のために存在しています。新しくは <seealso>ExcelReader</seealso>を使用して下さい。</font></para>
''' </remarks>
Public Class ExcelHandle

#Region "public static field"

	''' <summary>
	''' Excelの最大行数
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
    ''' コンストラクタ。このExcelオブジェクトで操作するEXCELのファイルを指定します。
    ''' </summary>
    ''' <param name="filepath">EXCELファイルのフルパス</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal filepath As String)

        Me._filepath = filepath

    End Sub

#End Region

#Region "public property"

    ''' <summary>
    ''' Excelファイルのフルパスを取得、設定します。
    ''' </summary>
    ''' <value>Excelファイルのフルパス</value>
    ''' <returns>Excelファイルのフルパス</returns>
    ''' <remarks>
    ''' <para>このオブジェクトで操作するExcelファイルのフルパスを取得、設定します。</para>    
    ''' <para>この値を変更することにより、読み込み、書き込み対象のファイルを変更します。</para>    
    ''' </remarks>
    Public ReadOnly Property FilePath() As String
        Get
            Return Me._filepath
        End Get
    End Property

	''' <summary>
	''' Excelシートの名前を設定、取得します。
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks>
	''' <para>Excelシートの名前を設定、または取得します。Loadメソッド実行前は、Null（Visual Basicの場合はNothing）が設定されています。</para>	
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
    ''' Excelのセルの内容を設定、取得します。
    ''' </summary>
    ''' <param name="row">行番号 1～</param>
    ''' <param name="col">列番号 1～</param>
    ''' <value>セルにセットする値</value>
    ''' <returns>セルから取得した値</returns>
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
    ''' Excelファイルの新規シートを作成します。
    ''' </summary>
    ''' <param name="sheetIndex"></param>
    ''' <param name="columnCount"></param>
    ''' <param name="rowCount"></param>
    ''' <remarks>
    ''' <para>Excelファイルの新規シートを作成します。</para> 
    ''' <para>編集した内容Save()メソッドにて保存することができます。</para>       
    ''' </remarks>
    Public Sub InitializeNewSheet( _
          ByVal sheetIndex As Integer _
        , ByVal columnCount As Integer _
        , ByVal rowCount As Integer)
        ' ☆☆ 未実装 ☆☆
    End Sub

	''' <summary>
	''' ブックに登録されているシートの数を取得します。
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

			'遅延バインディング

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

			'例外は再スローする
			Throw ex

		Finally

			'終了処理。COMオブジェクトを全て解放する。

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
				' アプリケーションの終了
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
	''' Excelファイルを読み込み、ファイルの内容をこのオブジェクトに展開します。
	''' </summary>
	''' <param name="sheetIndex">シート番号</param>
	''' <param name="columnCount">読み込む列数</param>
	''' <remarks>
	''' <para>Excelのファイルを読み込みます。値、書式などをこのオブジェクトの配列プロパティにて設定します。</para>
	''' <para>行内の全てのセルが空白になった時点で読み込みを終了します。空白行以降の内容は一切読み込まれないことに注意して下さい。</para>	
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
		Dim listCells(columnCount - 1) As Object ' listRowsの内容になります

		Dim listRanges As ArrayList = New ArrayList()

		Try

			'事前バインディング（DLL解析はビルド時）
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

				' 行内データ存在値の初期化
				containData = False

				' 配列の初期化
				listCells = New Object(columnCount - 1) {}

				' 行データの取得
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
					listRanges.Add(cellRange)	' ReleaseComObjectするため配列に格納する
					If Not (listCells(j - 1) Is Nothing) Then
						If Not (CStr(listCells(j - 1)).Trim().Equals(String.Empty)) Then
							' データが行内に含まれる
							containData = True
						End If
					End If
				Next

				' 行にデータが含まれないとき、ループを終了する
				If (Not containData) Then
					Exit For
				End If

				' 行データをリストに追加
				listRows.Add(listCells)

			Next

			Me._sheetData = New Object(listRows.Count - 1, columnCount - 1) {}
			Me._isChangedData = New Boolean(listRows.Count - 1, columnCount - 1) {}

			' 取得データを二次元配列に格納し直す
			For x As Integer = 0 To Me._sheetData.GetLength(0) - 1 Step 1
				Dim rowObject() As Object = CType(listRows(x), Object())
				For y As Integer = 0 To Me._sheetData.GetLength(1) - 1 Step 1
					Me._sheetData(x, y) = rowObject(y)
				Next
			Next


			'配列を初期化
			For xChange As Integer = 0 To Me._isChangedData.GetLength(0) - 1 Step 1
				For yChange As Integer = 0 To Me._isChangedData.GetLength(1) - 1 Step 1
					Me._isChangedData(xChange, yChange) = False
				Next
			Next

			' シート名
			Me._sheetName = CStr( _
				sheet.GetType().InvokeMember("Name" _
					, BindingFlags.GetProperty _
					, Nothing _
					, sheet _
					, Nothing))

		Catch ex As Exception
			'例外は再スローする
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
				' アプリケーションの終了
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
	''' Excelファイルを読み込み、ファイルの内容をこのクラスに展開します。
	''' </summary>
	''' <param name="sheetIndex">シート番号</param>    
	''' <param name="columnCount">読み込む列数</param>    
	''' <param name="rowCount">書き込む列数</param>    
	''' <remarks>
	''' <para>Excelのファイルを読み込みます。値、書式などをこのオブジェクトの配列プロパティにて設定します。</para>    
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

		'配列を初期化
		For x As Integer = 0 To Me._isChangedData.GetLength(0) - 1 Step 1
			For y As Integer = 0 To Me._isChangedData.GetLength(1) - 1 Step 1
				Me._isChangedData(x, y) = False
			Next
		Next

		Try

			'遅延バインディング

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

			'例外は再スローする
			Throw ex

		Finally

			'終了処理。COMオブジェクトを全て解放する。

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
				' アプリケーションの終了
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
    ''' このオブジェクトの内容を表す文字列を取得します。
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
    ''' 行と列の番号から、EXCELセル名称を取得します。
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
    ''' 列の番号から、EXCEL列名称を取得します。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetColumnSignature(ByVal columnCount As Integer) As String

        Dim sb As StringBuilder = New StringBuilder()

		Dim first As Char = CType("", Char)
		Dim second As Char = CType("", Char)

        If 256 < columnCount Then
            Throw New ArgumentOutOfRangeException("columnCount に入力されている列数が、Excel許容列数を超過しています。")
        End If

        If 26 < columnCount Then
            ' 26列以上の場合、列名は2文字
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
