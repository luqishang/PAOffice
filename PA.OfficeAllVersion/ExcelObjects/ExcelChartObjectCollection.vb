Namespace ExcelObjects

	''' <summary>
	''' Excel�O���t�I�u�W�F�N�g�R���N�V�����ł��B
	''' </summary>
	''' <remarks></remarks>
	Public NotInheritable Class ExcelChartObjectCollection
		Inherits System.Data.InternalDataCollectionBase

#Region "Private Fields"

		Private _charts As IList(Of ExcelChartObject) = New List(Of ExcelChartObject)

#End Region

#Region "Public Default Properties"

		''' <summary>
		''' �o�^����Ă���EXCEL�O���t�I�u�W�F�N�g���擾���܂��B
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
		''' �o�^����Ă���O���t�I�u�W�F�N�g�̌����擾���܂��B
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
		''' �V�[�g�ɐV�����O���t�I�u�W�F�N�g��ǉ����܂��B
		''' </summary>
		''' <param name="chart"></param>
		''' <remarks></remarks>
		Public Sub Add(ByVal chart As ExcelChartObject)
			Me._charts.Add(chart)
		End Sub

#End Region

#Region "Public Overrides Methods"

		''' <summary>
		''' ���̃R���N�V���������L���Ă���O���t�I�u�W�F�N�g�̗񋓂��擾���܂��B
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides Function GetEnumerator() As System.Collections.IEnumerator
			Return Me._charts.GetEnumerator()
		End Function

#End Region

	End Class

End Namespace
