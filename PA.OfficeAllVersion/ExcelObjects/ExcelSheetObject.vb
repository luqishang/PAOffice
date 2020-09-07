Namespace ExcelObjects

	''' <summary>
	''' EXCELのシートを定義するクラスです。
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelSheetObject

#Region "Private Fields"

		Private _rows As New ExcelRowObjectCollection

		Private _charts As New ExcelChartObjectCollection

		Private _name As String

		Private _oldName As String

		Private _visible As Boolean = True

		Private _displayGridLine As Boolean = True

#End Region

#Region "Public Property"

		''' <summary>
		''' 行のコレクションを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Rows() As ExcelRowObjectCollection
			Get
				Return Me._rows
			End Get
		End Property

		''' <summary>
		''' グラフのコレクションを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Charts() As ExcelChartObjectCollection
			Get
				Return Me._charts
			End Get
		End Property

		''' <summary>
		''' シートの名前を取得、設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Name() As String
			Get
				Return Me._name
			End Get
			Set(ByVal value As String)
				Me._name = value
			End Set
		End Property

		''' <summary>
		''' ファイル読み込み時のシートの名前を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property OldName() As String
			Get
				Return Me._oldName
			End Get
		End Property

		''' <summary>
		''' このシートの情報を表示設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>Falseのとき、このシートは非表示になります。</remarks>
		Public Property Visible() As Boolean
			Get
				Return Me._visible
			End Get
			Set(ByVal value As Boolean)
				Me._visible = value
			End Set
		End Property

		''' <summary>
		''' このシートの枠線表示設定をします。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property DisplayGridLine() As Boolean
			Get
				Return Me._displayGridLine
			End Get
			Set(ByVal value As Boolean)
				Me._displayGridLine = value
			End Set
		End Property

#End Region

#Region "Friend Properties"

		''' <summary>
		''' ファイル読み込み時のシートの名前を設定します。
		''' </summary>
		''' <value></value>
		''' <remarks></remarks>
		Public WriteOnly Property SetOldName() As String
			Set(ByVal value As String)
				Me._oldName = value
			End Set
		End Property

#End Region

	End Class

End Namespace
