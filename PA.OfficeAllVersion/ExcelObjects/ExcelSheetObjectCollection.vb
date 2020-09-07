Namespace ExcelObjects

	''' <summary>
	''' Excelシートの内容を格納するコレクションクラス
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelSheetObjectCollection
		Inherits System.Data.InternalDataCollectionBase

#Region "Private Fields"

		Private _sheets As IList(Of ExcelSheetObject) = New List(Of ExcelSheetObject)

#End Region

#Region "Public Default Properties"

		''' <summary>
		''' EXCELのシートの情報を取得します。
		''' </summary>
		''' <param name="index"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Sheets(ByVal index As Integer) As ExcelSheetObject
			Get
				Return Me._sheets(index)
			End Get
		End Property

#End Region

#Region "Public Properties"

		''' <summary>
		''' このブックに登録されているEXCELシートの数を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides ReadOnly Property Count() As Integer
			Get
				Return Me._sheets.Count
			End Get
		End Property

#End Region

#Region "Public Methods"

		''' <summary>
		''' EXCELのシートを追加します。
		''' </summary>
		''' <param name="sheet"></param>
		''' <remarks></remarks>
		Public Sub Add(ByVal sheet As ExcelSheetObject)
			Me._sheets.Add(sheet)
		End Sub

#End Region

#Region "Public Overrides Methods"

		''' <summary>
		''' このコレクションに登録されているシートオブジェクトの列挙を取得します。
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides Function GetEnumerator() As System.Collections.IEnumerator
			Return Me._sheets.GetEnumerator()
		End Function

#End Region

	End Class

End Namespace
