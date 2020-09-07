Namespace ExcelObjects

	''' <summary>
	''' EXCEL�V�[�g�̍s�R���N�V�����@�\��񋟂��܂��B
	''' </summary>
	''' <remarks></remarks>
	Public NotInheritable Class ExcelRowObjectCollection
		Inherits System.Data.InternalDataCollectionBase

#Region "Private Fields"

		Private _rows As IList(Of ExcelRowObject) = New List(Of ExcelRowObject)

#End Region

#Region "Public Default Properties"

		''' <summary>
		''' EXCEL�V�[�g�̍s�̏����擾���܂��B
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
		''' �o�^����Ă���EXCEL�s���̃��R�[�h�����擾���܂��B
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
		''' Excel�̍s��ǉ����܂��B
		''' </summary>
		''' <param name="row"></param>
		''' <remarks></remarks>
		Public Sub Add(ByVal row As ExcelRowObject)
			Me._rows.Add(row)
		End Sub

#End Region

#Region "Public Overrides Methods"

		''' <summary>
		''' ���̃R���N�V�����ɓo�^����Ă���s�I�u�W�F�N�g�̗񋓂��擾���܂��B
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides Function GetEnumerator() As System.Collections.IEnumerator
			Return Me._rows.GetEnumerator()
		End Function

#End Region

	End Class

End Namespace
