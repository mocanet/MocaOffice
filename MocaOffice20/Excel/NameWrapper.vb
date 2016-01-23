
Imports System.Runtime.InteropServices

Namespace Excel

    Public Class NameWrapper
        Inherits AbstractExcelWrapper

        Private _names As NamesWrapper

        Private _name As Object

#Region " Logging For Log4net "
        ''' <summary>Logging For Log4net</summary>
        Private Shared ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(String.Empty)
#End Region

#Region " コンストラクタ "

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="names">親のNames</param>
        ''' <param name="name">Excel.Name</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal names As NamesWrapper, ByVal name As Object)
            MyBase.New(names.ApplicationWrapper)
            _names = names
            _name = name
        End Sub

#End Region

#Region " Overrides "

        ''' <summary>
        ''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub MyDispose()
            ReleaseExcelObject(_name)
        End Sub

        ''' <summary>
        ''' 取得した Excel インスタンス
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Overrides ReadOnly Property OrigianlInstance As Object
            Get
                Return _name
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
        ''' 親のNames
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Names() As NamesWrapper
            Get
                Return _names
            End Get
        End Property

        ''' <summary>
        ''' 指定した名前の分類名を、コード記述時の言語で設定します
        ''' </summary>
        ''' <returns></returns>
        Public Property Category() As String
            Get
                Return DirectCast(InvokeGetProperty(_name, "Category", Nothing), String)
            End Get
            Set(ByVal value As String)
                InvokeSetProperty(_name, "Category", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' 指定した名前がユーザー定義の関数やマクロ (マクロ シートの関数マクロやコマンド マクロ) を参照するように定義されている場合、その分類名をコード記述時の言語の文字列として設定します
        ''' </summary>
        ''' <returns></returns>
        Public Property CategoryLocal() As String
            Get
                Return DirectCast(InvokeGetProperty(_name, "CategoryLocal", Nothing), String)
            End Get
            Set(ByVal value As String)
                InvokeSetProperty(_name, "CategoryLocal", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' 同種のオブジェクトのコレクション内のオブジェクトを特定するインデックス番号を取得します
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Index() As Integer
            Get
                Return DirectCast(InvokeGetProperty(_name, "Index", Nothing), Integer)
            End Get
        End Property

        ''' <summary>
        ''' 指定した名前の参照先を設定します
        ''' </summary>
        ''' <returns></returns>
        Public Property MacroType() As XlXLMMacroType
            Get
                Return DirectCast(InvokeGetProperty(_name, "MacroType", Nothing), XlXLMMacroType)
            End Get
            Set(ByVal value As XlXLMMacroType)
                InvokeSetProperty(_name, "MacroType", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' オブジェクトの名前を設定します
        ''' </summary>
        ''' <returns></returns>
        Public Property Name() As String
            Get
                Return DirectCast(InvokeGetProperty(_name, "Name", Nothing), String)
            End Get
            Set(ByVal value As String)
                InvokeSetProperty(_name, "Name", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' コード実行時の言語でオブジェクトの名前を設定します
        ''' </summary>
        ''' <returns></returns>
        Public Property NameLocal() As String
            Get
                Return DirectCast(InvokeGetProperty(_name, "NameLocal", Nothing), String)
            End Get
            Set(ByVal value As String)
                InvokeSetProperty(_name, "NameLocal", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' 指定されたオブジェクトの親オブジェクトを返します
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Parent() As SheetWrapper
            Get
                Dim obj As Object
                Dim wrap As SheetWrapper

                obj = InvokeGetProperty(_name, "Parent", Nothing)

                wrap = New SheetWrapper(_names.Parent, obj)
                addXlsObject(wrap)

                Return wrap
            End Get
        End Property

        ''' <summary>
        ''' 指定した名前が参照する数式を設定します。数式は等号 (=) で始まり、コード記述時の言語を使用して A1 形式で記述します
        ''' </summary>
        ''' <returns></returns>
        Public Property RefersTo() As String
            Get
                Return DirectCast(InvokeGetProperty(_name, "RefersTo", Nothing), String)
            End Get
            Set(ByVal value As String)
                InvokeSetProperty(_name, "RefersTo", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' 指定した名前が参照する数式 (セル範囲) を設定します
        ''' </summary>
        ''' <returns></returns>
        Public Property RefersToLocal() As String
            Get
                Return DirectCast(InvokeGetProperty(_name, "RefersToLocal", Nothing), String)
            End Get
            Set(ByVal value As String)
                InvokeSetProperty(_name, "RefersToLocal", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' 指定した名前が参照する数式 (セル範囲) を設定します
        ''' </summary>
        ''' <returns></returns>
        Public Property RefersToR1C1() As String
            Get
                Return DirectCast(InvokeGetProperty(_name, "RefersToR1C1", Nothing), String)
            End Get
            Set(ByVal value As String)
                InvokeSetProperty(_name, "RefersToR1C1", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' 指定した名前が参照する数式 (セル範囲) を設定します
        ''' </summary>
        ''' <returns></returns>
        Public Property RefersToR1C1Local() As String
            Get
                Return DirectCast(InvokeGetProperty(_name, "RefersToR1C1Local", Nothing), String)
            End Get
            Set(ByVal value As String)
                InvokeSetProperty(_name, "RefersToR1C1Local", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' Name オブジェクトの参照先の Range オブジェクトを取得します
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property RefersToRange() As RangeWrapper
            Get
                Dim obj As Object
                Dim wrap As RangeWrapper

                obj = InvokeGetProperty(_name, "RefersToRange", Nothing)

                wrap = New RangeWrapper(Parent, obj)
                addXlsObject(wrap)

                Return wrap
            End Get
        End Property

        ''' <summary>
        ''' ユーザー定義マクロ コマンドとして定義されている名前にショートカット キーを設定します
        ''' </summary>
        ''' <returns></returns>
        Public Property ShortcutKey() As String
            Get
                Return DirectCast(InvokeGetProperty(_name, "ShortcutKey", Nothing), String)
            End Get
            Set(ByVal value As String)
                InvokeSetProperty(_name, "ShortcutKey", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' 名前の参照先の数式を含む文字列を取得します
        ''' </summary>
        ''' <returns></returns>
        Public Property Value() As String
            Get
                Return DirectCast(InvokeGetProperty(_name, "Value", Nothing), String)
            End Get
            Set(ByVal value As String)
                InvokeSetProperty(_name, "Value", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' オブジェクトを表示するか、非表示にするかを決定します
        ''' </summary>
        ''' <returns></returns>
        Public Property Visible() As Boolean
            Get
                Return DirectCast(InvokeGetProperty(_name, "Visible", Nothing), Boolean)
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty(_name, "Visible", New Object() {value})
            End Set
        End Property

#End Region
#Region " Method "

        ''' <summary>
        ''' オブジェクトを削除します
        ''' </summary>
        Public Sub Delete()
            InvokeMethod(_name, "Delete", Nothing)
        End Sub

#End Region

    End Class

End Namespace
