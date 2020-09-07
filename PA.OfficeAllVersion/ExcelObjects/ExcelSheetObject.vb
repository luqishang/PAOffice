Namespace ExcelObjects

	''' <summary>
	''' EXCEL�̃V�[�g���`����N���X�ł��B
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
		''' �s�̃R���N�V�������擾���܂��B
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
		''' �O���t�̃R���N�V�������擾���܂��B
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
		''' �V�[�g�̖��O���擾�A�ݒ肵�܂��B
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
		''' �t�@�C���ǂݍ��ݎ��̃V�[�g�̖��O���擾���܂��B
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
		''' ���̃V�[�g�̏���\���ݒ肵�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>False�̂Ƃ��A���̃V�[�g�͔�\���ɂȂ�܂��B</remarks>
		Public Property Visible() As Boolean
			Get
				Return Me._visible
			End Get
			Set(ByVal value As Boolean)
				Me._visible = value
			End Set
		End Property

		''' <summary>
		''' ���̃V�[�g�̘g���\���ݒ�����܂��B
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
		''' �t�@�C���ǂݍ��ݎ��̃V�[�g�̖��O��ݒ肵�܂��B
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
