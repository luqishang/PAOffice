Namespace ExcelObjects

	''' <summary>
	''' EXCELシートのセルのコレクション機能を提供します。
	''' </summary>
	''' <remarks></remarks>
	Public NotInheritable Class ExcelCellObjectCollection
		Inherits System.Data.InternalDataCollectionBase

#Region "Private Fields"

		Private _cells As IList(Of ExcelCellObject) = New List(Of ExcelCellObject)

#End Region

#Region "Public Default Properties"

		''' <summary>
		''' EXCELのセルの内容を設定、取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Cells(ByVal index As Integer) As ExcelCellObject
			Get
				Return Me._cells(index)
			End Get
		End Property

#End Region

#Region "Public Properties"

		''' <summary>
		''' EXCELシート行の内容が登録されているセルの個数を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides ReadOnly Property Count() As Integer
			Get
				Return Me._cells.Count
			End Get
		End Property

#End Region

#Region "Public Methods"

		''' <summary>
		''' EXCELのセルの内容を追加します。
		''' </summary>
		''' <param name="cell"></param>
		''' <remarks></remarks>
		Public Sub Add(ByVal cell As ExcelCellObject)
			Me._cells.Add(cell)
		End Sub

#End Region

#Region "Public Overrides Methods"

		''' <summary>
		''' このコレクションに登録されているセルオブジェクトの列挙を取得します。
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides Function GetEnumerator() As System.Collections.IEnumerator
			Return Me._cells.GetEnumerator()
		End Function

#End Region

	End Class

End Namespace
