Namespace Excel

	''' <summary>
	''' シート内容を構成する為のインタフェース
	''' </summary>
	''' <remarks>
	''' 複雑な帳票設計など、セルに対して詳細に操作が必要なときに使用します。<br/>
	''' </remarks>
	Public Interface ISheetContents

		''' <summary>
		''' ブックに存在するベースとなるシート名
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' 出力対象となるブック上に存在するシート名を返すようにする。
		''' 新規に追加するシートのときは空文字を返すようにする。
		''' </remarks>
		Property BaseSheetName() As String

		''' <summary>
		''' 保存時に使用するシート名
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' 保存時に使用するシート名を返すようにする。
		''' <see cref="BaseSheetName"/> で指定したシート名と同一の場合は空文字を返す。
		''' </remarks>
		ReadOnly Property SaveSheetName() As String

		''' <summary>
		''' 出力内容をセルへ設定する
		''' </summary>
		''' <param name="sheet">該当するシート操作インスタンス</param>
		''' <remarks>
		''' </remarks>
		Sub WriteContents(ByVal sheet As SheetWrapper)

	End Interface

End Namespace
