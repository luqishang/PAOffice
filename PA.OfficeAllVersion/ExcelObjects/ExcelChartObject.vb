Imports System.Text
Namespace ExcelObjects

	''' <summary>
	''' EXCELのグラフオブジェクトを定義するクラスです。なお、現在は列方向データのみ対応となってます。
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelChartObject

#Region "Private Fields"

		Private _chartType As ExcelChartType

		Private _dataSourceSheetIndex As Integer = 0

		Private _dataSourceStartRowIndex As Integer = 0

		Private _dataSourceStartColumnIndex As Integer = 0

		Private _dataSourceEndRowIndex As Integer = 0

		Private _dataSourceEndColumnIndex As Integer = 0

		Private _positionX As Integer = 0

		Private _positionY As Integer = 0

		Private _chartName As String

#End Region

#Region "Public Properties"

		''' <summary>
		''' グラフの種類を設定、取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ChartType() As ExcelChartType
			Get
				Return Me._chartType
			End Get
			Set(ByVal value As ExcelChartType)
				Me._chartType = value
			End Set
		End Property

		''' <summary>
		''' 表示位置（X座標）を設定、取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PositionX() As Integer
			Get
				Return Me._positionX
			End Get
			Set(ByVal value As Integer)
				Me._positionX = value
			End Set
		End Property

		''' <summary>
		''' 表示位置（Y座標）を設定、取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PositionY() As Integer
			Get
				Return Me._positionY
			End Get
			Set(ByVal value As Integer)
				Me._positionY = value
			End Set
		End Property

		''' <summary>
		''' グラフのタイトルを設定、取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ChartName() As String
			Get
				Return Me._chartName
			End Get
			Set(ByVal value As String)
				Me._chartName = value
			End Set
		End Property

		''' <summary>
		''' 
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property DataSourceSheetIndex() As Integer
			Get
				Return Me._dataSourceSheetIndex
			End Get
		End Property

		''' <summary>
		''' 
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property DataSourceStartRowIndex() As Integer
			Get
				Return Me._dataSourceStartRowIndex
			End Get
		End Property

		''' <summary>
		''' 
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property DataSourceStartColumnIndex() As Integer
			Get
				Return Me._dataSourceStartColumnIndex
			End Get
		End Property

		''' <summary>
		''' 
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property DataSourceEndRowIndex() As Integer
			Get
				Return Me._dataSourceEndRowIndex
			End Get
		End Property

		''' <summary>
		''' 
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property DataSourceEndColumnIndex() As Integer
			Get
				Return Me._dataSourceEndColumnIndex
			End Get
		End Property

#End Region

#Region "Friends Property"

#End Region

#Region "Public Methods"

		'''' <summary>
		'''' データ範囲を行、列の番号で設定します。（未実装）
		'''' </summary>
		'''' <param name="fromRowIndex"></param>
		'''' <param name="fromColumnIndex"></param>
		'''' <param name="toRowIndex"></param>
		'''' <param name="toColumnIndex"></param>
		'''' <remarks></remarks>
		'Public Sub SetDataSource( _
		'	ByVal fromRowIndex As Integer, _
		'	ByVal fromColumnIndex As Integer, _
		'	ByVal toRowIndex As Integer, _
		'	ByVal toColumnIndex As Integer)

		'	Dim sb As New StringBuilder()
		'	sb.Append(ExcelBookControl.GetCellSignature(fromColumnIndex, fromRowIndex))
		'	sb.Append(":")
		'	sb.Append(ExcelBookControl.GetCellSignature(toColumnIndex, toRowIndex))
		'	Me._dataSource = sb.ToString()

		'End Sub

		'''' <summary>
		'''' データ範囲をセル名（ZZ99形式）で設定します。
		'''' </summary>
		'''' <param name="fromCellName"></param>
		'''' <param name="toCellName"></param>
		'''' <remarks></remarks>
		'Public Sub SetDataSource( _
		'	ByVal fromCellName As String, _
		'	ByVal toCellName As String)

		'	Me._dataSource _
		'		= New StringBuilder().Append(fromCellName).Append(toCellName).ToString()

		'End Sub

		''' <summary>
		''' データ範囲を指定します。
		''' </summary>
		''' <param name="sheetIndex"></param>		
		''' <param name="startRowIndex"></param>		
		''' <param name="startColumnIndex"></param>		
		''' <param name="endRowIndex"></param>		
		''' <param name="endColumnIndex"></param>		
		''' <remarks></remarks>
		Public Sub SetDataSource( _
			ByVal sheetIndex As Integer, _
			ByVal startRowIndex As Integer, _
			ByVal startColumnIndex As Integer, _
			ByVal endRowIndex As Integer, _
			ByVal endColumnIndex As Integer)

			Me._dataSourceSheetIndex = sheetIndex
			Me._dataSourceStartRowIndex = startRowIndex
			Me._dataSourceStartColumnIndex = startColumnIndex
			Me._dataSourceEndRowIndex = endRowIndex
			Me._dataSourceEndColumnIndex = endColumnIndex

		End Sub

#End Region

	End Class

End Namespace
