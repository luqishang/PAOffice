Namespace ExcelObjects

	''' <summary>
	''' EXCEL�u�b�N�̍s�̏���\���N���X�ł��B
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelRowObject

#Region "Public Constructor"

        Public Sub New()

        End Sub

        ''' <summary>
        ''' �R���X�g���N�^
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New(ByVal Columns As Integer)
            For i As Integer = 0 To Columns - 1 Step 1
                Dim excelCell As New ExcelCellObject
                Me.Cells.Add(excelCell)
            Next
        End Sub

#End Region

		Private _cellCollection As New ExcelCellObjectCollection

		''' <summary>
		''' �Z���̃R���N�V�������擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Cells() As ExcelCellObjectCollection
			Get
				Return Me._cellCollection
			End Get
		End Property

	End Class

End Namespace
