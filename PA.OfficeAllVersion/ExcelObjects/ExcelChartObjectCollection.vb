Namespace ExcelObjects

	''' <summary>
	''' Excelグラフオブジェクトコレクションです。
	''' </summary>
	''' <remarks></remarks>
	Public NotInheritable Class ExcelChartObjectCollection
		Inherits System.Data.InternalDataCollectionBase

#Region "Private Fields"

		Private _charts As IList(Of ExcelChartObject) = New List(Of ExcelChartObject)

#End Region

#Region "Public Default Properties"

		''' <summary>
		''' 登録されているEXCELグラフオブジェクトを取得します。
		''' </summary>
		''' <param name="index"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Item(ByVal index As Integer) As ExcelChartObject
			Get
				Return Me._charts(index)
			End Get
		End Property

#End Region

#Region "Public Overrides Properties"

		''' <summary>
		''' 登録されているグラフオブジェクトの個数を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides ReadOnly Property Count() As Integer
			Get
				Return Me._charts.Count
			End Get
		End Property

#End Region

#Region "Public Methods"

		''' <summary>
		''' シートに新しくグラフオブジェクトを追加します。
		''' </summary>
		''' <param name="chart"></param>
		''' <remarks></remarks>
		Public Sub Add(ByVal chart As ExcelChartObject)
			Me._charts.Add(chart)
		End Sub

#End Region

#Region "Public Overrides Methods"

		''' <summary>
		''' このコレクションが所有しているグラフオブジェクトの列挙を取得します。
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides Function GetEnumerator() As System.Collections.IEnumerator
			Return Me._charts.GetEnumerator()
		End Function

#End Region

	End Class

End Namespace
