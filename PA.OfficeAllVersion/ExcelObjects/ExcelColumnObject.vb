Namespace ExcelObjects

	''' <summary>
	''' EXCEL�̗��`������N���X�ł��B
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelColumnObject

#Region "Private Fields"
		Private _name As String
#End Region

#Region "Public Constructor"

		''' <summary>
		''' ���̃N���X�̃C���X�^���X�𐶐����܂��B
		''' </summary>
		''' <remarks></remarks>
		Public Sub New()

		End Sub

		''' <summary>
		''' EXCEL��̒�`����ݒ肵�A�C���X�^���X�𐶐����܂��B
		''' </summary>
		''' <param name="name"></param>
		''' <remarks></remarks>
		Public Sub New(ByVal name As String)

			Me._name = name

		End Sub

#End Region

#Region "Public Properties"

		''' <summary>
		''' 
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

#End Region

	End Class

End Namespace
