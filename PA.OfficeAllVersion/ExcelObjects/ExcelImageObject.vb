Imports System.Drawing

Namespace ExcelObjects

    ''' <summary>
    ''' EXCELの画像情報を格納します。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ExcelImageObject

#Region "Public Properties"

        Private _imageData As Image

        ''' <summary>
        ''' Image型の画像データ
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ImageData() As Image
            Get
                Return Me._imageData
            End Get
            Set(ByVal value As Image)
                Me._imageData = value
            End Set
        End Property

        Private _rowIndex As Integer
        ''' <summary>
        ''' セルの行のインデックス（画像表示用）
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

        Private _colIndex As Integer
        ''' <summary>
        ''' セルの列のインデックス（画像表示用）
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

#End Region

    End Class

End Namespace
