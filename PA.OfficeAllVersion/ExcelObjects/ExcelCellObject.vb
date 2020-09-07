Namespace ExcelObjects

	''' <summary>
	''' EXCELのセル情報を格納します。
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
		''' ExcelBookControlのLoad()で取得したセルの内容。
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
		''' セルの内容に変更があるときは、True を取得します。
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
        ''' セルの色
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
        ''' セルの色の番号
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
        ''' 表す Font の色
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
        ''' 表す Font の色
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
        ''' セルの親行のインデックス
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
        ''' セルの列のインデックス
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
        ''' セルの列のインデックス
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
		''' OldValueプロパティに値を設定します。このプロパティは外部アセンブリからの参照はできません。
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
