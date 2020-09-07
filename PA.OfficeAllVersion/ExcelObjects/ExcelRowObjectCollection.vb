Namespace ExcelObjects

	''' <summary>
	''' EXCELシートの行コレクション機能を提供します。
	''' </summary>
	''' <remarks></remarks>
	Public NotInheritable Class ExcelRowObjectCollection
		Inherits System.Data.InternalDataCollectionBase

#Region "Private Fields"

		Private _rows As IList(Of ExcelRowObject) = New List(Of ExcelRowObject)

#End Region

#Region "Public Default Properties"

		''' <summary>
		''' EXCELシートの行の情報を取得します。
		''' </summary>
		''' <param name="index"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Rows(ByVal index As Integer) As ExcelRowObject
			Get
				Return Me._rows(index)
			End Get
		End Property

#End Region

#Region "Public Properties"

		''' <summary>
		''' 登録されているEXCEL行情報のレコード数を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides ReadOnly Property Count() As Integer
			Get
				Return Me._rows.Count
			End Get
		End Property

#End Region

#Region "Public Methods"

		''' <summary>
		''' Excelの行を追加します。
		''' </summary>
		''' <param name="row"></param>
		''' <remarks></remarks>
		Public Sub Add(ByVal row As ExcelRowObject)
			Me._rows.Add(row)
		End Sub

#End Region

#Region "Public Overrides Methods"

		''' <summary>
		''' このコレクションに登録されている行オブジェクトの列挙を取得します。
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides Function GetEnumerator() As System.Collections.IEnumerator
			Return Me._rows.GetEnumerator()
		End Function

#End Region

	End Class

End Namespace
