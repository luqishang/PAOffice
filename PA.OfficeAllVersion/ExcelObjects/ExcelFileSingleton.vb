' =============================================================================
'
'  名前空間        ：ExcelObjects
'  プログラム名称  ：Excel出入力についての共通部品
'  機能概要        ：注意点、COMを読み込んでいる時、スレッドを使う場合、
'                    STAモードで使用してください。またExcelFileSingletonのオブジェクトをLockしてください。
'  更新履歴        ：新規作成  2008/10/24  wcheng
'
'  Copyright (C)  2008  Pactera.Ltd.
' =============================================================================
Imports System.Text
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Diagnostics
Imports System.Windows.Forms

Namespace ExcelObjects


    Public Class ExcelFileSingleton

#Region "Singleton Design Pattern"

        Private Shared ReadOnly myInstance As New ExcelFileSingleton()

        Private Sub New()

        End Sub

        Public Shared Function GetInstance() As ExcelFileSingleton

            Return myInstance

        End Function

#End Region

#Region "private 変数"
        Private Shared m_application As Object = Nothing
        Public Shared m_books As Object = Nothing
        Public Shared m_book As Object = Nothing
        Private m_sheets As Object = Nothing
        Private m_sheetList As IList(Of Object) = New List(Of Object)
        Private m_filePath As String = String.Empty
#End Region

#Region "Private メソッド"

        ''' <summary>
        ''' シート名によって、シートインスタンスを取得する。
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetSheetByName(ByVal sheetName As String) As Object

            If m_sheets Is Nothing Then
                Return Nothing
            End If

            Dim sheet As Object = m_sheets.GetType.InvokeMember( _
                        "Item", _
                        BindingFlags.GetProperty, _
                        Nothing, _
                        m_sheets, _
                        New Object() {sheetName})

            Return sheet

        End Function

        ''' <summary>
        ''' シートの名前を取得する
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetNameOfSheet(ByVal sheet As Object) As String

            Return CType(sheet.GetType.InvokeMember( _
                        "Name", _
                        BindingFlags.GetProperty, _
                        Nothing, _
                        sheet, _
                        Nothing), String)

        End Function

        ''' <summary>
        ''' パラメータCell1Cell2によってワークシートのRangeオブジェクトを取得する
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="Cell1Cell2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetExcelRange(ByVal sheet As Object, ByVal Cell1Cell2 As Object) As Object

            Return sheet.GetType().InvokeMember( _
                       "Range", _
                       BindingFlags.GetProperty, _
                       Nothing, _
                       sheet, _
                       New Object() {Cell1Cell2})
        End Function

       
        ''' <summary>
        ''' rangeのBordersを取得する。
        ''' </summary>
        ''' <param name="range"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetRangeBorders(ByVal range As Object) As Object

            Return range.GetType().InvokeMember( _
                       "Borders", _
                       BindingFlags.GetProperty, _
                       Nothing, _
                       range, _
                       Nothing)
        End Function

        ''' <summary>
        ''' 罫線の線のスタイルを設定します。
        ''' </summary>
        ''' <param name="borders"></param>
        ''' <param name="XlLineStyleDown"></param>
        ''' <param name="XlLineStyleUp"></param>
        ''' <param name="XlLineStyleLeft"></param>
        ''' <param name="XlLineStyleTop"></param>
        ''' <param name="XlLineStyleBottom"></param>
        ''' <param name="XlLineStyleRight"></param>
        ''' <remarks></remarks>
        Private Sub SetBordersLineStyle(ByVal borders As Object, _
            ByVal XlLineStyleDown As Integer, _
            ByVal XlLineStyleUp As Integer, _
            ByVal XlLineStyleLeft As Integer, _
            ByVal XlLineStyleTop As Integer, _
            ByVal XlLineStyleRight As Integer, _
            ByVal XlLineStyleBottom As Integer)

            'セル範囲の各セルの左上隅から右下隅への罫線
            Dim borderDown As Object = Nothing
            'セル範囲の各セルの左下隅から右上隅への罫線
            Dim borderUp As Object = Nothing
            'セル範囲の左側の罫線
            Dim borderLeft As Object = Nothing
            'セル範囲の上側の罫線
            Dim borderTop As Object = Nothing
            'セル範囲の右側の罫線
            Dim borderRight As Object = Nothing
            'セル範囲の下側の罫線
            Dim borderBottom As Object = Nothing
            
            Try
                'セル範囲の各セルの左上隅から右下隅への罫線
                borderDown = GetRangeBorder(borders, XlBordersIndex.xlDiagonalDown)
                SetBorderLineStyle(borderDown, XlLineStyleDown)
                'セル範囲の各セルの左下隅から右上隅への罫線
                borderUp = GetRangeBorder(borders, XlBordersIndex.xlDiagonalUp)
                SetBorderLineStyle(borderUp, XlLineStyleUp)
                'セル範囲の左側の罫線
                borderLeft = GetRangeBorder(borders, XlBordersIndex.xlEdgeLeft)
                SetBorderLineStyle(borderLeft, XlLineStyleLeft)
                'セル範囲の上側の罫線
                borderTop = GetRangeBorder(borders, XlBordersIndex.xlEdgeTop)
                SetBorderLineStyle(borderTop, XlLineStyleTop)
                'セル範囲の右側の罫線
                borderRight = GetRangeBorder(borders, XlBordersIndex.xlEdgeRight)
                SetBorderLineStyle(borderRight, XlLineStyleRight)
                'セル範囲の下側の罫線
                borderBottom = GetRangeBorder(borders, XlBordersIndex.xlEdgeBottom)
                SetBorderLineStyle(borderBottom, XlLineStyleBottom)
                
            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(borderDown)
                ReleaseComObject(borderUp)
                ReleaseComObject(borderLeft)
                ReleaseComObject(borderTop)
                ReleaseComObject(borderBottom)
                ReleaseComObject(borderRight)
            End Try

        End Sub

        ''' <summary>
        ''' 罫線の列挙型
        ''' </summary>
        ''' <remarks></remarks>
        Private Enum XlBordersIndex
            'セル範囲の各セルの左上隅から右下隅への罫線
            xlDiagonalDown = 5
            'セル範囲の各セルの左下隅から右上隅への罫線
            xlDiagonalUp = 6
            'セル範囲の下側の罫線
            xlEdgeBottom = 9
            'セル範囲の左側の罫線
            xlEdgeLeft = 7
            'セル範囲の右側の罫線
            xlEdgeRight = 10
            'セル範囲の上側の罫線
            xlEdgeTop = 8
        End Enum

        ''' <summary>
        ''' 罫線の線のスタイルを設定する。
        ''' </summary>
        ''' <param name="border"></param>
        ''' <param name="xlLineStyle"></param>
        ''' <remarks></remarks>
        Private Sub SetBorderLineStyle(ByVal border As Object, ByVal xlLineStyle As Integer)

            border.GetType().InvokeMember( _
                   "LineStyle", _
                   BindingFlags.SetProperty, _
                   Nothing, _
                   border, _
                   New Object() {xlLineStyle})

        End Sub

        ''' <summary>
        ''' 罫線を取得する。
        ''' </summary>
        ''' <param name="borders"></param>
        ''' <param name="index"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetRangeBorder(ByVal borders As Object, ByVal index As Integer) As Object

            Return borders.GetType().InvokeMember( _
                   "Item", _
                   BindingFlags.GetProperty, _
                   Nothing, _
                   borders, _
                   New Object() {index})

        End Function

        ''' <summary>
        ''' Interior 型オブジェクトを取得する
        ''' </summary>
        ''' <param name="range"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetRangeInterior(ByVal range As Object) As Object
            Return range.GetType().InvokeMember( _
                       "Interior", _
                       BindingFlags.GetProperty, _
                       Nothing, _
                       range, _
                       Nothing)
        End Function

        ''' <summary>
        ''' Rangeのフォント属性を取得する
        ''' </summary>
        ''' <param name="range"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetRangeFont(ByVal range As Object) As Object

            Return range.GetType().InvokeMember( _
                       "Font", _
                       BindingFlags.GetProperty, _
                       Nothing, _
                       range, _
                       Nothing)
        End Function

        ''' <summary>
        ''' RangeのValue属性を取得する
        ''' </summary>
        ''' <param name="range"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetRangeValue(ByVal range As Object) As Object

            Return range.GetType().InvokeMember( _
                       "Value", _
                       BindingFlags.GetProperty, _
                       Nothing, _
                       range, _
                       Nothing)

        End Function

        ''' <summary>
        ''' 指定したinteriorオブジェクトの色を設定する
        ''' </summary>
        ''' <param name="interior"></param>
        ''' <param name="color"></param>
        ''' <remarks></remarks>
        Private Sub SetRangeInteriorColor(ByVal interior As Object, ByVal color As Object)

            interior.GetType().InvokeMember( _
                       "Color", _
                       BindingFlags.SetProperty, _
                       Nothing, _
                       interior, _
                       New Object() {color})
        End Sub

        ''' <summary>
        ''' 指定したinteriorオブジェクトの色コードを設定する
        ''' </summary>
        ''' <param name="interior"></param>
        ''' <param name="colorIndex"></param>
        ''' <remarks></remarks>
        Private Sub SetRangeInteriorColorIndex(ByVal interior As Object, ByVal colorIndex As Object)

            interior.GetType().InvokeMember( _
                       "ColorIndex", _
                       BindingFlags.SetProperty, _
                       Nothing, _
                       interior, _
                       New Object() {colorIndex})
        End Sub

        ''' <summary>
        ''' 指定したフォントオブジェクトの色を設定する
        ''' </summary>
        ''' <param name="font"></param>
        ''' <param name="color"></param>
        ''' <remarks></remarks>
        Private Sub SetRangeFontColor(ByVal font As Object, ByVal color As Object)

            font.GetType().InvokeMember( _
                       "Color", _
                       BindingFlags.SetProperty, _
                       Nothing, _
                       font, _
                       New Object() {color})
        End Sub

        ''' <summary>
        ''' 指定したフォントオブジェクトの色コードを設定する。
        ''' </summary>
        ''' <param name="font"></param>
        ''' <param name="colorIndex"></param>
        ''' <remarks></remarks>
        Private Sub SetRangeFontColorIndex(ByVal font As Object, ByVal colorIndex As Object)

            font.GetType().InvokeMember( _
                       "ColorIndex", _
                       BindingFlags.SetProperty, _
                       Nothing, _
                       font, _
                       New Object() {colorIndex})
        End Sub

        ''' <summary>
        ''' 指定したRangeのValueを設定する
        ''' </summary>
        ''' <param name="range"></param>
        ''' <param name="value"></param>
        ''' <remarks></remarks>
        Private Sub SetRangeValue(ByVal range As Object, ByVal value As Object)

            range.GetType().InvokeMember( _
                    "Value", _
                    BindingFlags.SetProperty, _
                    Nothing, _
                    range, _
                    New Object() {value})

        End Sub

        ''' <summary>
        ''' 指定したrangeSourceをrangeDestにコピーする
        ''' </summary>
        ''' <param name="rangeSource"></param>
        ''' <param name="rangeDest"></param>
        ''' <remarks></remarks>
        Private Sub RangeCopy(ByVal rangeSource As Object, ByVal rangeDest As Object)

            rangeSource.GetType().InvokeMember( _
                       "Copy", _
                       BindingFlags.InvokeMethod, _
                       Nothing, _
                       rangeSource, _
                       New Object() {rangeDest})

        End Sub

        ''' <summary>
        ''' 指定したRangeのInsertメソッドを呼び出す。
        ''' </summary>
        ''' <param name="range"></param>
        ''' <remarks></remarks>
        Private Sub RangeInsert(ByVal range As Object)

            range.GetType().InvokeMember(
                       "Insert",
                       BindingFlags.InvokeMethod,
                       Nothing,
                       range,
                       Nothing)

        End Sub

        ''' <summary>
        ''' 指定したRangeのDeleteメソッドを呼び出す。
        ''' </summary>
        ''' <param name="range"></param>
        ''' <remarks></remarks>
        Private Sub RangeDelete(ByVal range As Object)

            range.GetType().InvokeMember(
                       "Delete",
                       BindingFlags.InvokeMethod,
                       Nothing,
                       range,
                       Nothing)

        End Sub

        ''' <summary>
        ''' 指定したワークブックを閉じる
        ''' </summary>
        ''' <param name="book"></param>
        ''' <remarks></remarks>
        Private Sub CloseBook(ByVal book As Object)

            If book IsNot Nothing Then

                Try

                    book.GetType.InvokeMember( _
                                         "Close", _
                                         BindingFlags.InvokeMethod, _
                                         Nothing, _
                                         book, _
                                         Nothing)
                Catch ex As Exception

                End Try


            End If
        End Sub

        ''' <summary>
        ''' 主Windowを持っていないExcelのProcessをキールする。
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub KillExcelProcess()

            For Each proc As Process In Process.GetProcessesByName("EXCEL")
                If (proc.MainWindowHandle = IntPtr.Zero) Then
                    proc.Kill()
                End If
            Next

        End Sub

        ''' <summary>
        ''' Microsoft Excel を終了します。 
        ''' </summary>
        ''' <param name="app"></param>
        ''' <remarks></remarks>
        Private Sub QuitApp(ByVal app As Object)
            If app IsNot Nothing Then

                Try
                    app.GetType().InvokeMember( _
                                     "Quit", _
                                     BindingFlags.InvokeMethod, _
                                     Nothing, _
                                     app, _
                                     Nothing)
                Catch ex As Exception

                    KillExcelProcess()

                End Try


            End If
        End Sub

        ''' <summary>
        ''' COMオブジェクトを解放する
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <remarks></remarks>
        Private Sub ReleaseComObject(ByRef obj As Object)
            If obj IsNot Nothing Then
                If Marshal.IsComObject(obj) Then

                    Try
                        Marshal.ReleaseComObject(obj)
                        obj = Nothing
                    Catch ex As Exception

                    End Try

                End If

            End If
        End Sub

        ''' <summary>
        ''' セル単位で背景色を設定する。
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="cell"></param>
        ''' <remarks></remarks>
        Private Sub SetCellBgColor(ByVal sheet As Object, ByVal cell As ExcelCellObject)

            Dim range As Object = Nothing
            Dim interior As Object = Nothing

            Dim signature As String = Nothing
            If cell.RowIndex <> 0 And cell.ColIndex <> 0 Then
                signature = ExcelBookControl.GetCellSignature(cell.ColIndex, cell.RowIndex)
            End If
            If cell.Range IsNot Nothing Then
                signature = cell.Range
            End If

            Try

                range = Me.GetExcelRange(sheet, signature)
                If (range IsNot Nothing) Then

                    interior = Me.GetRangeInterior(range)
                    '背景色
                    If (cell.ColorIndex.HasValue) Then
                        Me.SetRangeInteriorColorIndex(interior, cell.ColorIndex.Value)
                    End If
                    If (cell.Color IsNot Nothing) Then
                        Me.SetRangeInteriorColor(interior, cell.Color)
                    End If
                End If

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(interior)
                ReleaseComObject(range)
            End Try

        End Sub

        ''' <summary>
        ''' 指定したワークシートに指定したセルの値とか、背景色とか、フォント色とか設定する。
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="cell"></param>
        ''' <remarks></remarks>
        Private Sub WriteCellToSheet(ByVal sheet As Object, ByVal cell As ExcelCellObject)

            Dim range As Object = Nothing
            Dim interior As Object = Nothing
            Dim font As Object = Nothing

            Dim signature As String = Nothing
            If cell.RowIndex <> 0 And cell.ColIndex <> 0 Then
                signature = ExcelBookControl.GetCellSignature(cell.ColIndex, cell.RowIndex)
            End If
            If cell.Range IsNot Nothing Then
                signature = cell.Range
            End If

            Try

                range = Me.GetExcelRange(sheet, signature)

                If (range IsNot Nothing) Then

                    interior = Me.GetRangeInterior(range)
                    font = Me.GetRangeFont(range)

                    '背景色
                    If (cell.ColorIndex.HasValue) Then
                        Me.SetRangeInteriorColorIndex(interior, cell.ColorIndex.Value)
                    End If
                    If (cell.Color IsNot Nothing) Then
                        Me.SetRangeInteriorColor(interior, cell.Color)
                    End If
                    'Font色
                    If (cell.FontColorIndex.HasValue) Then
                        Me.SetRangeFontColorIndex(font, cell.FontColorIndex.Value)
                    End If
                    If (cell.FontColor IsNot Nothing) Then
                        Me.SetRangeFontColor(font, cell.FontColor)
                    End If

                    Me.SetRangeValue(range, cell.Value)

                End If

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(interior)
                ReleaseComObject(font)
                ReleaseComObject(range)
            End Try

        End Sub

        ''' <summary>
        ''' 指定したワークシートに指定したセルに、画像情報を出力する。
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="image"></param>
        ''' <remarks></remarks>
        Private Sub WriteImageToSheet(ByVal sheet As Object, ByVal image As ExcelImageObject)

            Dim range As Object = Nothing
            Dim signature As String = ExcelBookControl.GetCellSignature(image.ColIndex, image.RowIndex)

            Try
                range = Me.GetExcelRange(sheet, signature)

                'データをコピーする
                Clipboard.SetDataObject(image.ImageData, True)

                'データを貼り付け
                sheet.GetType().InvokeMember( _
                     "Paste", _
                     BindingFlags.InvokeMethod, _
                     Nothing, _
                     sheet, _
                     New Object() {range, Type.Missing})

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(range)
            End Try

        End Sub

        ''' <summary>
        ''' 指定したワークシート、行インデックス、列インデックスでセルの情報を取得する。
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="colIndex"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetCellData(ByVal sheet As Object, ByVal rowIndex As Integer, ByVal colIndex As Integer) As ExcelCellObject

            Dim cell As New ExcelCellObject

            Dim signature As String = ExcelBookControl.GetCellSignature(colIndex, rowIndex)
            Dim range As Object = Nothing

            Try
                range = Me.GetExcelRange(sheet, signature)

                'valueオブジェクトはCOMオブジェクトじゃない。
                Dim value As Object = GetRangeValue(range)
                cell.Value = value
                cell.RowIndex = rowIndex
                cell.ColIndex = colIndex

            Catch ex As Exception

            Finally
                ReleaseComObject(range)
            End Try

            Return cell

        End Function

        ''' <summary>
        ''' シートの名前を設定する。
        ''' </summary>
        ''' <param name="sheet">シート</param>
        ''' <param name="sheetName">シート名</param>
        ''' <remarks></remarks>
        Private Sub SetSheetName(ByVal sheet As Object, ByVal sheetName As String)

            If sheet Is Nothing Then
                Return
            End If

            sheet.GetType().InvokeMember( _
                          "Name", _
                          BindingFlags.SetProperty, _
                          Nothing, _
                          sheet, _
                          New Object() {sheetName})

        End Sub

#End Region

#Region "public メソッド"

        ''' <summary>
        ''' 指定したファイルパスで、Excelファイルを読み込みます。
        ''' </summary>
        ''' <param name="filePath">読み込みパス</param>
        ''' <remarks></remarks>
        Public Sub OpenExcel(ByVal filePath As String)

            Try
                m_filePath = filePath

                If m_application Is Nothing Then
                    m_application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"))

                    m_application.GetType().InvokeMember( _
                    "Visible", _
                    BindingFlags.SetProperty, _
                    Nothing, _
                    m_application, _
                    New Object() {False})

                    m_application.GetType().InvokeMember( _
                        "ScreenUpdating", _
                        BindingFlags.SetProperty, _
                        Nothing, _
                        m_application, _
                    New Object() {False})

                    m_application.GetType().InvokeMember( _
                         "DisplayAlerts", _
                         BindingFlags.SetProperty, _
                         Nothing, _
                         m_application, _
                         New Object() {False})
                    m_application.GetType().InvokeMember( _
                         "AlertBeforeOverwriting", _
                         BindingFlags.SetProperty, _
                         Nothing, _
                         m_application, _
                         New Object() {False})
                End If

                ' ブックの新規作成・開く
                If m_books Is Nothing Then
                    m_books = m_application.GetType().InvokeMember( _
                                         "Workbooks", _
                                         BindingFlags.GetProperty, _
                                         Nothing, _
                                         m_application, _
                                         Nothing)
                End If

                If (File.Exists(filePath)) Then
                    m_book = m_books.GetType().InvokeMember( _
                       "Open", _
                       BindingFlags.InvokeMethod, _
                       Nothing, _
                       m_books, _
                       New Object() {filePath})

                    m_sheets = m_book.GetType().InvokeMember( _
                       "Worksheets", _
                       BindingFlags.GetProperty, _
                       Nothing, _
                       m_book, _
                       Nothing)

                    Dim sheetsCountMax As Integer _
                         = DirectCast( _
                          m_sheets.GetType().InvokeMember( _
                           "Count", _
                           BindingFlags.GetProperty, _
                           Nothing, _
                           m_sheets, _
                           Nothing), _
                          Integer)

                    For i As Integer = 1 To sheetsCountMax

                        ' 既存ファイルにシートがある場合
                        m_sheetList.Add( _
                         m_sheets.GetType().InvokeMember( _
                          "Item", _
                          BindingFlags.GetProperty, _
                          Nothing, _
                          m_sheets, _
                          New Object() {i}))

                    Next

                End If


            Catch ex As Exception

                'Excelについてのクローズ処理
                CloseExcel()

            End Try

        End Sub

        ''' <summary>
        ''' 該当ExcelファイルについてのCOMオブジェクトを解放する。
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub CloseExcel()

            Try
                m_book.GetType().InvokeMember( _
                    "SaveAs", _
                    BindingFlags.InvokeMethod, _
                    Nothing, _
                    m_book, _
                    New Object() {m_filePath})

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する。
                For Each sheet As Object In m_sheetList
                    ReleaseComObject(sheet)
                Next
                m_sheetList.Clear()
                ReleaseComObject(m_sheets)

                CloseBook(m_book)
                ReleaseComObject(m_book)

                CloseBook(m_books)
                ReleaseComObject(m_books)

                QuitApp(m_application)
                ReleaseComObject(m_application)
            End Try

        End Sub

        ''' <summary>
        ''' 指定のファイルパスの指定シートには 指定範囲sourceを指定範囲destにコピーする
        ''' </summary>
        ''' <param name="sheetName">指定シート名前</param>
        ''' <param name="rangeSource">指定範囲source</param>
        ''' <param name="rangeDest">指定範囲dest</param>
        ''' <remarks></remarks>
        Public Sub SheetRangeCopy(ByVal sheetName As String, _
                                  ByVal rangeSource As String, _
                                  ByVal rangeDest As String)

            Dim sheet As Object = Nothing
            Dim rangeFrom As Object = Nothing
            Dim rangeTo As Object = Nothing

            Try
                'シート名によって、シートを取得する
                sheet = Me.GetSheetByName(sheetName)

                If sheet IsNot Nothing Then

                    rangeFrom = Me.GetExcelRange(sheet, rangeSource)
                    rangeTo = Me.GetExcelRange(sheet, rangeDest)

                    'rangeFromからrangeToにデータをコピーする
                    RangeCopy(rangeFrom, rangeTo)

                End If

            Catch ex As Exception

            Finally

                'COMオブジェクトを解放する
                ReleaseComObject(rangeTo)
                ReleaseComObject(rangeFrom)
                ReleaseComObject(sheet)
            End Try


        End Sub

        ''' <summary>
        ''' 指定したワークシートにデータを出力する。
        ''' </summary>
        ''' <param name="sheetName">シート名</param>
        ''' <param name="cells">出力データ「List(Of ExcelCellObject)型」</param>
        ''' <remarks></remarks>
        Public Sub WriterCellsToSheet(ByVal sheetName As String, ByVal cells As List(Of ExcelCellObject))

            Dim sheet As Object = Nothing

            Try

                sheet = Me.GetSheetByName(sheetName)

                If (sheet IsNot Nothing) Then
                    For Each cell As ExcelCellObject In cells
                        WriteCellToSheet(sheet, cell)
                    Next
                End If

            Catch ex As Exception

            Finally

                'COMオブジェクトを解放する
                ReleaseComObject(sheet)

            End Try

        End Sub

        ''' <summary>
        ''' 指定したワークシートに、データを出力する。
        ''' 出力データが多い場合、性能が悪くなったので、WriteRowsToSheetByArrayを利用してください。
        ''' </summary>
        ''' <param name="sheetName">シート名</param>
        ''' <param name="rows">出力データ「List(Of ExcelRowObject)型」</param>
        ''' <remarks></remarks>
        Public Sub WriteRowsToSheet(ByVal sheetName As String, ByVal rows As List(Of ExcelRowObject))

            Dim sheet As Object = Nothing

            Try
                sheet = Me.GetSheetByName(sheetName)

                For Each row As ExcelRowObject In rows

                    For i As Integer = 0 To row.Cells.Count - 1

                        Dim cell As ExcelCellObject = row.Cells(i)
                        WriteCellToSheet(sheet, cell)

                    Next

                Next

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' 指定したワークシートの指定した行、指定した列から指定した列までの情報を取得する。
        ''' </summary>
        ''' <param name="sheetName">シート名</param>
        ''' <param name="rowIndex">開始行</param>
        ''' <param name="colIndexFrom">開始列</param>
        ''' <param name="colIndexTo">終了列</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReadRowData(ByVal sheetName As String, ByVal rowIndex As Integer, ByVal colIndexFrom As Integer, ByVal colIndexTo As Integer) As List(Of ExcelCellObject)

            Dim cells As New List(Of ExcelCellObject)

            Dim sheet As Object = Nothing

            Try
                sheet = Me.GetSheetByName(sheetName)

                For colIndex As Integer = colIndexFrom To colIndexTo

                    Dim cell As ExcelCellObject = GetCellData(sheet, rowIndex, colIndex)
                    cells.Add(cell)

                Next

            Catch ex As Exception

            Finally

                'COMオブジェクトを解放する
                ReleaseComObject(sheet)

            End Try

            Return cells

        End Function

        ''' <summary>
        ''' 指定したワークシートの指定した範囲から情報を取得する。
        ''' </summary>
        ''' <param name="sheetName">シート名</param>
        ''' <param name="keyCol">キー列</param>
        ''' <param name="startRowIndex">開始行</param>
        ''' <param name="colFrom">開始列</param>
        ''' <param name="colTo">終了列</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReadRowsData(ByVal sheetName As String, ByVal keyCol As Integer, ByVal startRowIndex As Integer, ByVal colFrom As Integer, ByVal colTo As Integer) As List(Of ExcelRowObject)

            Dim rows As New List(Of ExcelRowObject)


            Dim sheet As Object = Nothing

            Try

                sheet = Me.GetSheetByName(sheetName)

                'キー列のセルを取得する
                Dim keyCell As ExcelCellObject = GetCellData(sheet, startRowIndex, keyCol)

                While keyCell.Value IsNot Nothing

                    Dim row As New ExcelRowObject
                    rows.Add(row)

                    'キー列のセルの値は空白ではない場合、データを取得しつつ
                    For colIndex As Integer = colFrom To colTo
                        Dim cell As ExcelCellObject = GetCellData(sheet, startRowIndex, colIndex)
                        row.Cells.Add(cell)
                    Next

                    '次の行のデータを取得する
                    startRowIndex += 1
                    keyCell = GetCellData(sheet, startRowIndex, keyCol)

                End While

            Catch ex As Exception

            Finally

                'COMオブジェクトを解放する
                ReleaseComObject(sheet)

            End Try

            Return rows

        End Function

        ''' <summary>
        ''' 指定したシートの指定したcells情報を読み込みます。
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="cells"></param>
        ''' <remarks></remarks>
        Public Sub ReadCellsData(ByVal sheetName As String, ByVal cells As List(Of ExcelCellObject))

            Dim sheet As Object = Nothing

            Try
                sheet = Me.GetSheetByName(sheetName)

                For Each cell As ExcelCellObject In cells

                    Dim cellData As ExcelCellObject = GetCellData(sheet, cell.RowIndex, cell.ColIndex)
                    cell.Value = cellData.Value

                Next

            Catch ex As Exception

            Finally

                'COMオブジェクトを解放する
                ReleaseComObject(sheet)

            End Try


        End Sub

        ''' <summary>
        ''' 該当ワークブックの全部のシート名を取得する。
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSheetNames() As List(Of String)

            Dim sheetNames As New List(Of String)

            For Each sheet As Object In m_sheetList
                sheetNames.Add(GetNameOfSheet(sheet))
            Next

            Return sheetNames

        End Function

        ''' <summary>
        ''' 指定したワークシートの指定した行に新しい行を挿入する。
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="rowIndex"></param>
        ''' <remarks></remarks>
        Public Sub InsertRowOfSheet(ByVal sheetName As String, ByVal rowIndex As Integer, ByVal count As Integer)

            If count <= 0 Then
                Exit Sub
            End If

            Dim sheet As Object = Nothing
            Dim range As Object = Nothing

            Try
                sheet = Me.GetSheetByName(sheetName)

                If (sheet IsNot Nothing) Then

                    Dim cell1cell2 As String = rowIndex.ToString() + ":" + rowIndex.ToString()

                    range = Me.GetExcelRange(sheet, cell1cell2)

                    For i As Integer = 1 To count Step 1
                        rangeInsert(range)
                    Next

                End If

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(range)
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' 指定したワークシートの指定した列に新しい列を挿入する。
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="colIndex"></param>
        ''' <param name="count"></param>
        ''' <remarks></remarks>
        Public Sub InsertColOfSheet(ByVal sheetName As String, ByVal colIndex As Integer, ByVal count As Integer)

            If count <= 0 Then
                Exit Sub
            End If

            Dim sheet As Object = Nothing
            Dim range As Object = Nothing

            Try
                sheet = Me.GetSheetByName(sheetName)

                If (sheet IsNot Nothing) Then

                    Dim col1 As String = ExcelBookControl.GetColumnSignature(colIndex)
                    Dim cell1cell2 As String = col1 + ":" + col1

                    range = Me.GetExcelRange(sheet, cell1cell2)

                    For i As Integer = 1 To count Step 1
                        RangeInsert(range)
                    Next

                End If

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(range)
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' 指定したワークシートの指定した列以降、指定した列数の列を削除する。
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="colIndex"></param>
        ''' <param name="count"></param>
        ''' <remarks></remarks>
        Public Sub DeleteColOfSheet(ByVal sheetName As String, ByVal colIndex As Integer, ByVal count As Integer)

            If count <= 0 Then
                Exit Sub
            End If

            Dim sheet As Object = Nothing
            Dim range As Object = Nothing

            Try
                sheet = Me.GetSheetByName(sheetName)

                If (sheet IsNot Nothing) Then

                    Dim col1 As String = ExcelBookControl.GetColumnSignature(colIndex + 1)
                    Dim col2 As String = ExcelBookControl.GetColumnSignature(colIndex + count)
                    Dim cell1cell2 As String = col1 + ":" + col2

                    range = Me.GetExcelRange(sheet, cell1cell2)
                    RangeDelete(range)

                End If

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(range)
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' 指定したシート名「afterSheetName」の後に、「sheetName」というシートを挿入する。
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="afterSheetName"></param>
        ''' <remarks></remarks>
        Public Sub AddWorksheetAfter(ByVal sheetName As String, ByVal afterSheetName As String)

            Dim afterSheet As Object = Nothing

            Try
                afterSheet = Me.GetSheetByName(afterSheetName)

                If afterSheet IsNot Nothing Then

                    Dim sheet As Object = m_sheets.GetType().InvokeMember( _
                             "Add", _
                             BindingFlags.InvokeMethod, _
                             Nothing, _
                             m_sheets, _
                             New Object() {Type.Missing, afterSheet, Type.Missing, Type.Missing})

                    m_sheetList.Add(sheet)

                    If (Not String.IsNullOrEmpty(sheetName)) Then
                        'シートの名前を設定する。
                        SetSheetName(sheet, sheetName)
                    End If

                End If

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(afterSheet)
            End Try

        End Sub

        ''' <summary>
        ''' 指定したワークシートに、データを出力する。
        ''' </summary>
        ''' <param name="sheetName">シート名</param>
        ''' <param name="images">出力データ「List(Of ExcelRowObject)型」</param>
        ''' <remarks></remarks>
        Public Sub WriteImagesToSheet(ByVal sheetName As String, ByVal images As List(Of ExcelImageObject))

            Dim sheet As Object = Nothing

            Try
                sheet = Me.GetSheetByName(sheetName)

                For Each image As ExcelImageObject In images

                    WriteImageToSheet(sheet, image)

                Next

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' 指定したシートを選択する。
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <remarks></remarks>
        Public Sub WorksheetSelect(ByVal sheetName As String)

            Dim sheet As Object = Nothing

            Try
                sheet = Me.GetSheetByName(sheetName)

                If (sheet IsNot Nothing) Then
                    sheet.GetType.InvokeMember( _
                        "Select", _
                        BindingFlags.InvokeMethod, _
                        Nothing, _
                        sheet, _
                        Nothing)
                End If

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' 罫線の線のスタイルの設定する。
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <param name="signature"></param>
        ''' <param name="down"></param>
        ''' <param name="up"></param>
        ''' <param name="left"></param>
        ''' <param name="top"></param>
        ''' <param name="bottom"></param>
        ''' <param name="right"></param>
        ''' <remarks></remarks>
        Public Sub SetRangeLineStyle(ByVal sheetName As String, ByVal signature As String, _
            ByVal down As Integer, _
            ByVal up As Integer, _
            ByVal left As Integer, _
            ByVal top As Integer, _
            ByVal right As Integer, _
            ByVal bottom As Integer)

            Dim sheet As Object = Nothing
            Dim range As Object = Nothing
            Dim borders As Object = Nothing

            Try
                sheet = Me.GetSheetByName(sheetName)
                range = Me.GetExcelRange(sheet, signature)
                borders = Me.GetRangeBorders(range)

                SetBordersLineStyle(borders, down, up, left, top, right, bottom)

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(sheet)
                ReleaseComObject(range)
                ReleaseComObject(borders)
            End Try


        End Sub

        ''' <summary>
        ''' 出力データが数多い場合、性能が良くなる為に、
        ''' 配列でデータをExcelファイルに出力するように
        ''' </summary>
        ''' <param name="sheetName">シート名</param>
        ''' <param name="rows">出力詳細データ</param>
        ''' <param name="startRowIndex">開始行番号</param>
        ''' <param name="startColIndex">開始列番号</param>
        ''' <remarks></remarks>
        Public Sub WriteRowsToSheetByArray(ByVal sheetName As String, _
            ByVal rows As List(Of ExcelRowObject), _
            ByVal startRowIndex As Integer, _
            ByVal startColIndex As Integer)

            Dim cells As New List(Of ExcelCellObject)
            Dim arr As Array = Array.CreateInstance(GetType(String), rows.Count, 256)

            Dim bgCells As New List(Of ExcelCellObject)

            For Each row As ExcelRowObject In rows
                For Each cell As ExcelCellObject In row.Cells

                    '背景色があれば、
                    If (cell.Color IsNot Nothing Or cell.ColorIndex.HasValue) Then
                        bgCells.Add(cell)
                    End If

                    'セルの値が無ければ、次のセルの処理へ
                    If cell.Value Is Nothing Then
                        Continue For
                    End If

                    'セルの文字が912以上場合、セル単位で値を出力する。
                    'Excel 2003で、長い文字列配列を代入すると、実行時エラー1004を回避する為
                    If CStr(cell.Value).Length > 911 Then
                        cells.Add(cell)
                    Else
                        arr.SetValue(cell.Value, cell.RowIndex - startRowIndex, cell.ColIndex - startColIndex)
                    End If
                Next
            Next

            Dim sheet As Object = Nothing
            Dim range As Object = Nothing

            Try
                sheet = Me.GetSheetByName(sheetName)

                Dim sig1 As String = ExcelBookControl.GetCellSignature(startColIndex, startRowIndex)
                Dim sig2 As String = ExcelBookControl.GetCellSignature(256, startRowIndex + rows.Count - 1)
                range = GetExcelRange(sheet, sig1 + ":" + sig2)

                '一括でセルの値をセットする。
                range.GetType().InvokeMember( _
                       "Value2", _
                       BindingFlags.SetProperty, _
                       Nothing, _
                       range, _
                       New Object() {arr})

                'セルの値が912文字以上の場合、セルずつ値をセットする。
                WriterCellsToSheet(sheetName, cells)

                '背景色があれば、背景色をセットする。
                For Each cell As ExcelCellObject In bgCells
                    SetCellBgColor(sheet, cell)
                Next

            Catch ex As Exception

            Finally
                'COMオブジェクトを解放する
                ReleaseComObject(range)
                ReleaseComObject(sheet)
            End Try

        End Sub

#End Region

    End Class

    ''' <summary>
    ''' 罫線の線の種類の列挙型
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum xlLineStyle

        ''' <summary>
        ''' 実線
        ''' </summary>
        ''' <remarks></remarks>
        xlContinuous = 1
        ''' <summary>
        ''' 何もない
        ''' </summary>
        ''' <remarks></remarks>
        xlNone = -4142

    End Enum

End Namespace

