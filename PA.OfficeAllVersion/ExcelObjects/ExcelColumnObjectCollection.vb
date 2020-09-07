Namespace ExcelObjects

	''' <summary>
	''' 
	''' </summary>
	''' <remarks></remarks>
	Public NotInheritable Class ExcelColumnObjectCollection
		Inherits System.Data.InternalDataCollectionBase

#Region "Private Fields"

		Private _columns As IList(Of ExcelColumnObject) = New List(Of ExcelColumnObject)

#End Region

#Region "Public Default Properties"

		''' <summary>
		''' EXCEL�̗�̓��e���擾���܂��B
		''' </summary>
		''' <param name="index"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Columns(ByVal index As Integer) As ExcelColumnObject
			Get
				Return Me._columns(index)
			End Get
		End Property

#End Region

#Region "Public Properties"

		''' <summary>
		''' ��`����Ă����̌����擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides ReadOnly Property Count() As Integer
			Get
				Return Me._columns.Count
			End Get
		End Property

#End Region

#Region "Public Methods"

		''' <summary>
		''' ���`��ǉ����܂��B
		''' </summary>
		''' <param name="column"></param>
		''' <remarks></remarks>
		Public Sub Add(ByVal column As ExcelColumnObject)

			Me._columns.Add(column)

		End Sub

#End Region

#Region "Public Overrides Methods"

		''' <summary>
		''' ���̃R���N�V�����ɓo�^����Ă����I�u�W�F�N�g�̗񋓂��擾���܂��B
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides Function GetEnumerator() As System.Collections.IEnumerator
			Return Me._columns.GetEnumerator()
		End Function

#End Region

	End Class

End Namespace
