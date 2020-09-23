' =============================================================================
'
'  ���O���        �FExcelObjects
'  �v���O��������  �FExcel�o���͂ɂ��Ă̋��ʕ��i
'  �@�\�T�v        �F���ӓ_�ACOM��ǂݍ���ł��鎞�A�X���b�h���g���ꍇ�A
'                    STA���[�h�Ŏg�p���Ă��������B�܂�ExcelFileSingleton�̃I�u�W�F�N�g��Lock���Ă��������B
'  �X�V����        �F�V�K�쐬  2008/10/24  wcheng
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

#Region "private �ϐ�"
        Private Shared m_application As Object = Nothing
        Public Shared m_books As Object = Nothing
        Public Shared m_book As Object = Nothing
        Private m_sheets As Object = Nothing
        Private m_sheetList As IList(Of Object) = New List(Of Object)
        Private m_filePath As String = String.Empty
#End Region

#Region "Private ���\�b�h"

        ''' <summary>
        ''' �V�[�g���ɂ���āA�V�[�g�C���X�^���X���擾����B
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
        ''' �V�[�g�̖��O���擾����
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
        ''' �p�����[�^Cell1Cell2�ɂ���ă��[�N�V�[�g��Range�I�u�W�F�N�g���擾����
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
        ''' range��Borders���擾����B
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
        ''' �r���̐��̃X�^�C����ݒ肵�܂��B
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

            '�Z���͈͂̊e�Z���̍��������E�����ւ̌r��
            Dim borderDown As Object = Nothing
            '�Z���͈͂̊e�Z���̍���������E����ւ̌r��
            Dim borderUp As Object = Nothing
            '�Z���͈͂̍����̌r��
            Dim borderLeft As Object = Nothing
            '�Z���͈͂̏㑤�̌r��
            Dim borderTop As Object = Nothing
            '�Z���͈͂̉E���̌r��
            Dim borderRight As Object = Nothing
            '�Z���͈͂̉����̌r��
            Dim borderBottom As Object = Nothing
            
            Try
                '�Z���͈͂̊e�Z���̍��������E�����ւ̌r��
                borderDown = GetRangeBorder(borders, XlBordersIndex.xlDiagonalDown)
                SetBorderLineStyle(borderDown, XlLineStyleDown)
                '�Z���͈͂̊e�Z���̍���������E����ւ̌r��
                borderUp = GetRangeBorder(borders, XlBordersIndex.xlDiagonalUp)
                SetBorderLineStyle(borderUp, XlLineStyleUp)
                '�Z���͈͂̍����̌r��
                borderLeft = GetRangeBorder(borders, XlBordersIndex.xlEdgeLeft)
                SetBorderLineStyle(borderLeft, XlLineStyleLeft)
                '�Z���͈͂̏㑤�̌r��
                borderTop = GetRangeBorder(borders, XlBordersIndex.xlEdgeTop)
                SetBorderLineStyle(borderTop, XlLineStyleTop)
                '�Z���͈͂̉E���̌r��
                borderRight = GetRangeBorder(borders, XlBordersIndex.xlEdgeRight)
                SetBorderLineStyle(borderRight, XlLineStyleRight)
                '�Z���͈͂̉����̌r��
                borderBottom = GetRangeBorder(borders, XlBordersIndex.xlEdgeBottom)
                SetBorderLineStyle(borderBottom, XlLineStyleBottom)
                
            Catch ex As Exception

            Finally
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(borderDown)
                ReleaseComObject(borderUp)
                ReleaseComObject(borderLeft)
                ReleaseComObject(borderTop)
                ReleaseComObject(borderBottom)
                ReleaseComObject(borderRight)
            End Try

        End Sub

        ''' <summary>
        ''' �r���̗񋓌^
        ''' </summary>
        ''' <remarks></remarks>
        Private Enum XlBordersIndex
            '�Z���͈͂̊e�Z���̍��������E�����ւ̌r��
            xlDiagonalDown = 5
            '�Z���͈͂̊e�Z���̍���������E����ւ̌r��
            xlDiagonalUp = 6
            '�Z���͈͂̉����̌r��
            xlEdgeBottom = 9
            '�Z���͈͂̍����̌r��
            xlEdgeLeft = 7
            '�Z���͈͂̉E���̌r��
            xlEdgeRight = 10
            '�Z���͈͂̏㑤�̌r��
            xlEdgeTop = 8
        End Enum

        ''' <summary>
        ''' �r���̐��̃X�^�C����ݒ肷��B
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
        ''' �r�����擾����B
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
        ''' Interior �^�I�u�W�F�N�g���擾����
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
        ''' Range�̃t�H���g�������擾����
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
        ''' Range��Value�������擾����
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
        ''' �w�肵��interior�I�u�W�F�N�g�̐F��ݒ肷��
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
        ''' �w�肵��interior�I�u�W�F�N�g�̐F�R�[�h��ݒ肷��
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
        ''' �w�肵���t�H���g�I�u�W�F�N�g�̐F��ݒ肷��
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
        ''' �w�肵���t�H���g�I�u�W�F�N�g�̐F�R�[�h��ݒ肷��B
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
        ''' �w�肵��Range��Value��ݒ肷��
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
        ''' �w�肵��rangeSource��rangeDest�ɃR�s�[����
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
        ''' �w�肵��Range��Insert���\�b�h���Ăяo���B
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
        ''' �w�肵��Range��Delete���\�b�h���Ăяo���B
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
        ''' �w�肵�����[�N�u�b�N�����
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
        ''' ��Window�������Ă��Ȃ�Excel��Process���L�[������B
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
        ''' Microsoft Excel ���I�����܂��B 
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
        ''' COM�I�u�W�F�N�g���������
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
        ''' �Z���P�ʂŔw�i�F��ݒ肷��B
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
                    '�w�i�F
                    If (cell.ColorIndex.HasValue) Then
                        Me.SetRangeInteriorColorIndex(interior, cell.ColorIndex.Value)
                    End If
                    If (cell.Color IsNot Nothing) Then
                        Me.SetRangeInteriorColor(interior, cell.Color)
                    End If
                End If

            Finally
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(interior)
                ReleaseComObject(range)
            End Try

        End Sub

        ''' <summary>
        ''' �w�肵�����[�N�V�[�g�Ɏw�肵���Z���̒l�Ƃ��A�w�i�F�Ƃ��A�t�H���g�F�Ƃ��ݒ肷��B
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

                    '�w�i�F
                    If (cell.ColorIndex.HasValue) Then
                        Me.SetRangeInteriorColorIndex(interior, cell.ColorIndex.Value)
                    End If
                    If (cell.Color IsNot Nothing) Then
                        Me.SetRangeInteriorColor(interior, cell.Color)
                    End If
                    'Font�F
                    If (cell.FontColorIndex.HasValue) Then
                        Me.SetRangeFontColorIndex(font, cell.FontColorIndex.Value)
                    End If
                    If (cell.FontColor IsNot Nothing) Then
                        Me.SetRangeFontColor(font, cell.FontColor)
                    End If

                    Me.SetRangeValue(range, cell.Value)

                End If

            Finally
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(interior)
                ReleaseComObject(font)
                ReleaseComObject(range)
            End Try

        End Sub

        ''' <summary>
        ''' �w�肵�����[�N�V�[�g�Ɏw�肵���Z���ɁA�摜�����o�͂���B
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="image"></param>
        ''' <remarks></remarks>
        Private Sub WriteImageToSheet(ByVal sheet As Object, ByVal image As ExcelImageObject)

            Dim range As Object = Nothing
            Dim signature As String = ExcelBookControl.GetCellSignature(image.ColIndex, image.RowIndex)

            Try
                range = Me.GetExcelRange(sheet, signature)

                '�f�[�^���R�s�[����
                Clipboard.SetDataObject(image.ImageData, True)

                '�f�[�^��\��t��
                sheet.GetType().InvokeMember( _
                     "Paste", _
                     BindingFlags.InvokeMethod, _
                     Nothing, _
                     sheet, _
                     New Object() {range, Type.Missing})

            Catch ex As Exception

            Finally
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(range)
            End Try

        End Sub

        ''' <summary>
        ''' �w�肵�����[�N�V�[�g�A�s�C���f�b�N�X�A��C���f�b�N�X�ŃZ���̏����擾����B
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

                'value�I�u�W�F�N�g��COM�I�u�W�F�N�g����Ȃ��B
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
        ''' �V�[�g�̖��O��ݒ肷��B
        ''' </summary>
        ''' <param name="sheet">�V�[�g</param>
        ''' <param name="sheetName">�V�[�g��</param>
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

#Region "public ���\�b�h"

        ''' <summary>
        ''' �w�肵���t�@�C���p�X�ŁAExcel�t�@�C����ǂݍ��݂܂��B
        ''' </summary>
        ''' <param name="filePath">�ǂݍ��݃p�X</param>
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

                ' �u�b�N�̐V�K�쐬�E�J��
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

                        ' �����t�@�C���ɃV�[�g������ꍇ
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

                'Excel�ɂ��ẴN���[�Y����
                CloseExcel()

            End Try

        End Sub

        ''' <summary>
        ''' �Y��Excel�t�@�C���ɂ��Ă�COM�I�u�W�F�N�g���������B
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
                'COM�I�u�W�F�N�g���������B
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
        ''' �w��̃t�@�C���p�X�̎w��V�[�g�ɂ� �w��͈�source���w��͈�dest�ɃR�s�[����
        ''' </summary>
        ''' <param name="sheetName">�w��V�[�g���O</param>
        ''' <param name="rangeSource">�w��͈�source</param>
        ''' <param name="rangeDest">�w��͈�dest</param>
        ''' <remarks></remarks>
        Public Sub SheetRangeCopy(ByVal sheetName As String, _
                                  ByVal rangeSource As String, _
                                  ByVal rangeDest As String)

            Dim sheet As Object = Nothing
            Dim rangeFrom As Object = Nothing
            Dim rangeTo As Object = Nothing

            Try
                '�V�[�g���ɂ���āA�V�[�g���擾����
                sheet = Me.GetSheetByName(sheetName)

                If sheet IsNot Nothing Then

                    rangeFrom = Me.GetExcelRange(sheet, rangeSource)
                    rangeTo = Me.GetExcelRange(sheet, rangeDest)

                    'rangeFrom����rangeTo�Ƀf�[�^���R�s�[����
                    RangeCopy(rangeFrom, rangeTo)

                End If

            Catch ex As Exception

            Finally

                'COM�I�u�W�F�N�g���������
                ReleaseComObject(rangeTo)
                ReleaseComObject(rangeFrom)
                ReleaseComObject(sheet)
            End Try


        End Sub

        ''' <summary>
        ''' �w�肵�����[�N�V�[�g�Ƀf�[�^���o�͂���B
        ''' </summary>
        ''' <param name="sheetName">�V�[�g��</param>
        ''' <param name="cells">�o�̓f�[�^�uList(Of ExcelCellObject)�^�v</param>
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

                'COM�I�u�W�F�N�g���������
                ReleaseComObject(sheet)

            End Try

        End Sub

        ''' <summary>
        ''' �w�肵�����[�N�V�[�g�ɁA�f�[�^���o�͂���B
        ''' �o�̓f�[�^�������ꍇ�A���\�������Ȃ����̂ŁAWriteRowsToSheetByArray�𗘗p���Ă��������B
        ''' </summary>
        ''' <param name="sheetName">�V�[�g��</param>
        ''' <param name="rows">�o�̓f�[�^�uList(Of ExcelRowObject)�^�v</param>
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
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' �w�肵�����[�N�V�[�g�̎w�肵���s�A�w�肵���񂩂�w�肵����܂ł̏����擾����B
        ''' </summary>
        ''' <param name="sheetName">�V�[�g��</param>
        ''' <param name="rowIndex">�J�n�s</param>
        ''' <param name="colIndexFrom">�J�n��</param>
        ''' <param name="colIndexTo">�I����</param>
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

                'COM�I�u�W�F�N�g���������
                ReleaseComObject(sheet)

            End Try

            Return cells

        End Function

        ''' <summary>
        ''' �w�肵�����[�N�V�[�g�̎w�肵���͈͂�������擾����B
        ''' </summary>
        ''' <param name="sheetName">�V�[�g��</param>
        ''' <param name="keyCol">�L�[��</param>
        ''' <param name="startRowIndex">�J�n�s</param>
        ''' <param name="colFrom">�J�n��</param>
        ''' <param name="colTo">�I����</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReadRowsData(ByVal sheetName As String, ByVal keyCol As Integer, ByVal startRowIndex As Integer, ByVal colFrom As Integer, ByVal colTo As Integer) As List(Of ExcelRowObject)

            Dim rows As New List(Of ExcelRowObject)


            Dim sheet As Object = Nothing

            Try

                sheet = Me.GetSheetByName(sheetName)

                '�L�[��̃Z�����擾����
                Dim keyCell As ExcelCellObject = GetCellData(sheet, startRowIndex, keyCol)

                While keyCell.Value IsNot Nothing

                    Dim row As New ExcelRowObject
                    rows.Add(row)

                    '�L�[��̃Z���̒l�͋󔒂ł͂Ȃ��ꍇ�A�f�[�^���擾����
                    For colIndex As Integer = colFrom To colTo
                        Dim cell As ExcelCellObject = GetCellData(sheet, startRowIndex, colIndex)
                        row.Cells.Add(cell)
                    Next

                    '���̍s�̃f�[�^���擾����
                    startRowIndex += 1
                    keyCell = GetCellData(sheet, startRowIndex, keyCol)

                End While

            Catch ex As Exception

            Finally

                'COM�I�u�W�F�N�g���������
                ReleaseComObject(sheet)

            End Try

            Return rows

        End Function

        ''' <summary>
        ''' �w�肵���V�[�g�̎w�肵��cells����ǂݍ��݂܂��B
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

                'COM�I�u�W�F�N�g���������
                ReleaseComObject(sheet)

            End Try


        End Sub

        ''' <summary>
        ''' �Y�����[�N�u�b�N�̑S���̃V�[�g�����擾����B
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
        ''' �w�肵�����[�N�V�[�g�̎w�肵���s�ɐV�����s��}������B
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
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(range)
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' �w�肵�����[�N�V�[�g�̎w�肵����ɐV�������}������B
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
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(range)
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' �w�肵�����[�N�V�[�g�̎w�肵����ȍ~�A�w�肵���񐔂̗���폜����B
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
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(range)
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' �w�肵���V�[�g���uafterSheetName�v�̌�ɁA�usheetName�v�Ƃ����V�[�g��}������B
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
                        '�V�[�g�̖��O��ݒ肷��B
                        SetSheetName(sheet, sheetName)
                    End If

                End If

            Catch ex As Exception

            Finally
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(afterSheet)
            End Try

        End Sub

        ''' <summary>
        ''' �w�肵�����[�N�V�[�g�ɁA�f�[�^���o�͂���B
        ''' </summary>
        ''' <param name="sheetName">�V�[�g��</param>
        ''' <param name="images">�o�̓f�[�^�uList(Of ExcelRowObject)�^�v</param>
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
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' �w�肵���V�[�g��I������B
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
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(sheet)
            End Try

        End Sub

        ''' <summary>
        ''' �r���̐��̃X�^�C���̐ݒ肷��B
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
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(sheet)
                ReleaseComObject(range)
                ReleaseComObject(borders)
            End Try


        End Sub

        ''' <summary>
        ''' �o�̓f�[�^���������ꍇ�A���\���ǂ��Ȃ�ׂɁA
        ''' �z��Ńf�[�^��Excel�t�@�C���ɏo�͂���悤��
        ''' </summary>
        ''' <param name="sheetName">�V�[�g��</param>
        ''' <param name="rows">�o�͏ڍ׃f�[�^</param>
        ''' <param name="startRowIndex">�J�n�s�ԍ�</param>
        ''' <param name="startColIndex">�J�n��ԍ�</param>
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

                    '�w�i�F������΁A
                    If (cell.Color IsNot Nothing Or cell.ColorIndex.HasValue) Then
                        bgCells.Add(cell)
                    End If

                    '�Z���̒l��������΁A���̃Z���̏�����
                    If cell.Value Is Nothing Then
                        Continue For
                    End If

                    '�Z���̕�����912�ȏ�ꍇ�A�Z���P�ʂŒl���o�͂���B
                    'Excel 2003�ŁA����������z���������ƁA���s���G���[1004����������
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

                '�ꊇ�ŃZ���̒l���Z�b�g����B
                range.GetType().InvokeMember( _
                       "Value2", _
                       BindingFlags.SetProperty, _
                       Nothing, _
                       range, _
                       New Object() {arr})

                '�Z���̒l��912�����ȏ�̏ꍇ�A�Z�����l���Z�b�g����B
                WriterCellsToSheet(sheetName, cells)

                '�w�i�F������΁A�w�i�F���Z�b�g����B
                For Each cell As ExcelCellObject In bgCells
                    SetCellBgColor(sheet, cell)
                Next

            Catch ex As Exception

            Finally
                'COM�I�u�W�F�N�g���������
                ReleaseComObject(range)
                ReleaseComObject(sheet)
            End Try

        End Sub

#End Region

    End Class

    ''' <summary>
    ''' �r���̐��̎�ނ̗񋓌^
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum xlLineStyle

        ''' <summary>
        ''' ����
        ''' </summary>
        ''' <remarks></remarks>
        xlContinuous = 1
        ''' <summary>
        ''' �����Ȃ�
        ''' </summary>
        ''' <remarks></remarks>
        xlNone = -4142

    End Enum

End Namespace

