
Imports System.Runtime.InteropServices

Namespace Excel

    ''' <summary>
    ''' ブックにあるすべての Name オブジェクトのコレクション
    ''' </summary>
    Public Class NamesWrapper
        Inherits AbstractExcelWrapper

        Private _book As BookWrapper

        Private _names As Object

#Region " Logging For Log4net "
        ''' <summary>Logging For Log4net</summary>
        Private Shared ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(String.Empty)
#End Region

#Region " コンストラクタ "

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="book">親のシート</param>
        ''' <param name="names">Excel.Range</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal book As BookWrapper, ByVal names As Object)
            MyBase.New(book.ApplicationWrapper)
            _book = book
            _names = names
        End Sub

#End Region

#Region " Overrides "

        ''' <summary>
        ''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub MyDispose()
            ReleaseExcelObject(_names)
        End Sub

        ''' <summary>
        ''' 取得した Excel インスタンス
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Overrides ReadOnly Property OrigianlInstance As Object
            Get
                Return _names
            End Get
        End Property

#End Region
#Region " プロパティ "

        ''' <summary>
        ''' Excel.Application のラッパー
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property App() As ExcelWrapper
            Get
                Return DirectCast(MyBase.xlsWrapper, ExcelWrapper)
            End Get
        End Property

        ''' <summary>
        ''' 親のブック
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Book() As BookWrapper
            Get
                Return _book
            End Get
        End Property

        ''' <summary>
        ''' コレクションに含まれるオブジェクトの数
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Count() As Integer
            Get
                Return DirectCast(InvokeGetProperty(_names, "Count", Nothing), Integer)
            End Get
        End Property

        ''' <summary>
        ''' 指定されたオブジェクトの親オブジェクトを返します
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Parent() As BookWrapper
            Get
                Dim obj As Object
                Dim wrap As BookWrapper

                obj = InvokeGetProperty(_names, "Parent", Nothing)

                wrap = New BookWrapper(App, obj)
                addXlsObject(wrap)

                Return wrap
            End Get
        End Property

#End Region
#Region " Method "

        ''' <summary>
        ''' セル範囲の新しい名前を定義します
        ''' </summary>
        ''' <returns></returns>
        Public Function Add(
                           Optional ByVal Name As String = Nothing,
                           Optional ByVal RefersTo As Object = Nothing,
                           Optional ByVal Visible As Boolean = True,
                           Optional ByVal MacroType As XlXLMMacroType = XlXLMMacroType.xlNotXLM,
                           Optional ByVal ShortcutKey As String = Nothing,
                           Optional ByVal Category As Object = Nothing,
                           Optional ByVal NameLocal As String = Nothing,
                           Optional ByVal RefersToLocal As String = Nothing,
                           Optional ByVal CategoryLocal As String = Nothing,
                           Optional ByVal RefersToR1C1 As String = Nothing,
                           Optional ByVal RefersToR1C1Local As String = Nothing
                           ) As NameWrapper
            Dim argsV As New List(Of Object)
            Dim argsN As New List(Of String)
            Dim xl As Object
            Dim obj As NameWrapper
            obj = Nothing

            If Name IsNot Nothing Then
                argsV.Add(Name)
                argsN.Add("Name")
            End If
            If RefersTo IsNot Nothing Then
                argsV.Add(RefersTo)
                argsN.Add("RefersTo")
            End If
            argsV.Add(Visible)
            argsN.Add("Visible")
            argsV.Add(MacroType)
            argsN.Add("MacroType")
            If ShortcutKey IsNot Nothing Then
                argsV.Add(ShortcutKey)
                argsN.Add("ShortcutKey")
            End If
            If Category IsNot Nothing Then
                argsV.Add(Category)
                argsN.Add("Category")
            End If
            If NameLocal IsNot Nothing Then
                argsV.Add(NameLocal)
                argsN.Add("NameLocal")
            End If
            If RefersToLocal IsNot Nothing Then
                argsV.Add(RefersToLocal)
                argsN.Add("RefersToLocal")
            End If
            If CategoryLocal IsNot Nothing Then
                argsV.Add(CategoryLocal)
                argsN.Add("CategoryLocal")
            End If
            If RefersToR1C1 IsNot Nothing Then
                argsV.Add(RefersToR1C1)
                argsN.Add("RefersToR1C1")
            End If
            If RefersToR1C1Local IsNot Nothing Then
                argsV.Add(RefersToR1C1Local)
                argsN.Add("RefersToR1C1Local")
            End If

            xl = InvokeMethod(_names, "Add", argsV.ToArray, argsN.ToArray)
            If xl IsNot Nothing Then
                obj = New NameWrapper(Me, xl)
                addXlsObject(obj)
            End If
            Return obj
        End Function

        ''' <summary>
        ''' Names コレクションから、単一の Name オブジェクトを返します
        ''' </summary>
        ''' <returns></returns>
        Public Function Item(
                            Optional ByVal Index As Object = Nothing,
                            Optional ByVal IndexLocal As Object = Nothing,
                            Optional ByVal RefersTo As Object = Nothing
                            ) As NameWrapper
            Dim argsV As New List(Of Object)
            Dim argsN As New List(Of String)
            Dim xl As Object
            Dim obj As NameWrapper
            obj = Nothing

            If Index IsNot Nothing Then
                argsV.Add(Index)
                argsN.Add("Index")
            End If
            If IndexLocal IsNot Nothing Then
                argsV.Add(IndexLocal)
                argsN.Add("IndexLocal")
            End If
            If RefersTo IsNot Nothing Then
                argsV.Add(RefersTo)
                argsN.Add("RefersTo")
            End If

            xl = InvokeMethod(_names, "Item", argsV.ToArray, argsN.ToArray)
            If xl IsNot Nothing Then
                obj = New NameWrapper(Me, xl)
                addXlsObject(obj)
            End If
            Return obj
        End Function

#End Region

    End Class

End Namespace
