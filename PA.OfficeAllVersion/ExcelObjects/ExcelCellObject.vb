Namespace ExcelObjects

	''' <summary>
	''' EXCEL�̃Z�������i�[���܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelCellObject

#Region "Public Properties"

		Private _value As Object

		''' <summary>
		''' 
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Value() As Object
			Get
				Return Me._value
			End Get
			Set(ByVal value As Object)
				Me._value = value
			End Set
		End Property

		Private _oldValue As Object

		''' <summary>
		''' ExcelBookControl��Load()�Ŏ擾�����Z���̓��e�B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property OldValue() As Object
			Get
				Return Me._oldValue
			End Get
		End Property

		''' <summary>
		''' �Z���̓��e�ɕύX������Ƃ��́ATrue ���擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Changed() As Boolean
			Get
				If Me._oldValue Is Nothing And Me._value Is Nothing Then
					Return False
				End If
				If Me._oldValue Is Nothing And Me._value IsNot Nothing Then
					Return True
				End If
				If Me._oldValue IsNot Nothing And Me._value Is Nothing Then
					Return True
				End If
				If Not Me._oldValue.GetType().Equals(Me._value.GetType()) Then
					Return True
				End If
				If TypeOf Me._oldValue Is Integer Then
					Return DirectCast(Me._oldValue, Integer) <> DirectCast(Me._value, Integer)
				End If
				If TypeOf Me._oldValue Is Long Then
					Return DirectCast(Me._oldValue, Long) <> DirectCast(Me._value, Long)
				End If
				If TypeOf Me._oldValue Is Short Then
					Return DirectCast(Me._oldValue, Short) <> DirectCast(Me._value, Short)
				End If
				If TypeOf Me._oldValue Is Double Then
					Return DirectCast(Me._oldValue, Double) <> DirectCast(Me._value, Double)
				End If
				If TypeOf Me._oldValue Is Decimal Then
					Return DirectCast(Me._oldValue, Decimal) <> DirectCast(Me._value, Decimal)
				End If
				If TypeOf Me._oldValue Is DateTime Then
					Return DirectCast(Me._oldValue, DateTime) <> DirectCast(Me._value, DateTime)
				End If
				Return False
			End Get
        End Property

        Private _color As Object
        ''' <summary>
        ''' �Z���̐F
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Color() As Object
            Get
                Return Me._color
            End Get
            Set(ByVal value As Object)
                Me._color = value
            End Set
        End Property

        Private _colorIndex As Nullable(Of Long)
        ''' <summary>
        ''' �Z���̐F�̔ԍ�
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ColorIndex() As Nullable(Of Long)
            Get
                Return Me._colorIndex
            End Get
            Set(ByVal value As Nullable(Of Long))
                Me._colorIndex = value
            End Set
        End Property

        Private _fontColor As Object
        ''' <summary>
        ''' �\�� Font �̐F
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FontColor() As Object
            Get
                Return Me._fontColor
            End Get
            Set(ByVal value As Object)
                Me._fontColor = value
            End Set
        End Property

        Private _fontColorIndex As Nullable(Of Long)
        ''' <summary>
        ''' �\�� Font �̐F
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FontColorIndex() As Nullable(Of Long)
            Get
                Return Me._fontColorIndex
            End Get
            Set(ByVal value As Nullable(Of Long))
                Me._fontColorIndex = value
            End Set
        End Property

        Private _rowIndex As Integer = 0
        ''' <summary>
        ''' �Z���̐e�s�̃C���f�b�N�X
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RowIndex() As Integer
            Get
                Return Me._rowIndex
            End Get
            Set(ByVal value As Integer)
                Me._rowIndex = value
            End Set
        End Property

        Private _colIndex As Integer = 0
        ''' <summary>
        ''' �Z���̗�̃C���f�b�N�X
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ColIndex() As Integer
            Get
                Return Me._colIndex
            End Get
            Set(ByVal value As Integer)
                Me._colIndex = value
            End Set
        End Property

        Private _range As String = Nothing
        ''' <summary>
        ''' �Z���̗�̃C���f�b�N�X
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Range() As String
            Get
                Return Me._range
            End Get
            Set(ByVal value As String)
                Me._range = value
            End Set
        End Property

#End Region

#Region "Friend Properties"

		''' <summary>
		''' OldValue�v���p�e�B�ɒl��ݒ肵�܂��B���̃v���p�e�B�͊O���A�Z���u������̎Q�Ƃ͂ł��܂���B
		''' </summary>
		''' <value></value>
		''' <remarks></remarks>
		Friend WriteOnly Property SetOldValue() As Object
			Set(ByVal value As Object)
				Me._oldValue = value
			End Set
		End Property

#End Region

	End Class

End Namespace
