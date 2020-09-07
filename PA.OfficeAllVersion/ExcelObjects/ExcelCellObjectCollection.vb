Namespace ExcelObjects

	''' <summary>
	''' EXCEL�V�[�g�̃Z���̃R���N�V�����@�\��񋟂��܂��B
	''' </summary>
	''' <remarks></remarks>
	Public NotInheritable Class ExcelCellObjectCollection
		Inherits System.Data.InternalDataCollectionBase

#Region "Private Fields"

		Private _cells As IList(Of ExcelCellObject) = New List(Of ExcelCellObject)

#End Region

#Region "Public Default Properties"

		''' <summary>
		''' EXCEL�̃Z���̓��e��ݒ�A�擾���܂��B
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
		''' EXCEL�V�[�g�s�̓��e���o�^����Ă���Z���̌����擾���܂��B
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
		''' EXCEL�̃Z���̓��e��ǉ����܂��B
		''' </summary>
		''' <param name="cell"></param>
		''' <remarks></remarks>
		Public Sub Add(ByVal cell As ExcelCellObject)
			Me._cells.Add(cell)
		End Sub

#End Region

#Region "Public Overrides Methods"

		''' <summary>
		''' ���̃R���N�V�����ɓo�^����Ă���Z���I�u�W�F�N�g�̗񋓂��擾���܂��B
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides Function GetEnumerator() As System.Collections.IEnumerator
			Return Me._cells.GetEnumerator()
		End Function

#End Region

	End Class

End Namespace
