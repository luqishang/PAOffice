Namespace ExcelObjects

	''' <summary>
	''' Excel�V�[�g�̓��e���i�[����R���N�V�����N���X
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelSheetObjectCollection
		Inherits System.Data.InternalDataCollectionBase

#Region "Private Fields"

		Private _sheets As IList(Of ExcelSheetObject) = New List(Of ExcelSheetObject)

#End Region

#Region "Public Default Properties"

		''' <summary>
		''' EXCEL�̃V�[�g�̏����擾���܂��B
		''' </summary>
		''' <param name="index"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Sheets(ByVal index As Integer) As ExcelSheetObject
			Get
				Return Me._sheets(index)
			End Get
		End Property

#End Region

#Region "Public Properties"

		''' <summary>
		''' ���̃u�b�N�ɓo�^����Ă���EXCEL�V�[�g�̐����擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides ReadOnly Property Count() As Integer
			Get
				Return Me._sheets.Count
			End Get
		End Property

#End Region

#Region "Public Methods"

		''' <summary>
		''' EXCEL�̃V�[�g��ǉ����܂��B
		''' </summary>
		''' <param name="sheet"></param>
		''' <remarks></remarks>
		Public Sub Add(ByVal sheet As ExcelSheetObject)
			Me._sheets.Add(sheet)
		End Sub

#End Region

#Region "Public Overrides Methods"

		''' <summary>
		''' ���̃R���N�V�����ɓo�^����Ă���V�[�g�I�u�W�F�N�g�̗񋓂��擾���܂��B
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overrides Function GetEnumerator() As System.Collections.IEnumerator
			Return Me._sheets.GetEnumerator()
		End Function

#End Region

	End Class

End Namespace
