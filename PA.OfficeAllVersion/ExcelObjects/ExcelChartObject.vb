Imports System.Text
Namespace ExcelObjects

	''' <summary>
	''' EXCEL�̃O���t�I�u�W�F�N�g���`����N���X�ł��B�Ȃ��A���݂͗�����f�[�^�̂ݑΉ��ƂȂ��Ă܂��B
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
		''' �O���t�̎�ނ�ݒ�A�擾���܂��B
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
		''' �\���ʒu�iX���W�j��ݒ�A�擾���܂��B
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
		''' �\���ʒu�iY���W�j��ݒ�A�擾���܂��B
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
		''' �O���t�̃^�C�g����ݒ�A�擾���܂��B
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
		'''' �f�[�^�͈͂��s�A��̔ԍ��Őݒ肵�܂��B�i�������j
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
		'''' �f�[�^�͈͂��Z�����iZZ99�`���j�Őݒ肵�܂��B
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
		''' �f�[�^�͈͂��w�肵�܂��B
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
