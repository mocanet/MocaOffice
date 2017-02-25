
''' <summary>
''' 変換するファイルフォーマット
''' </summary>
''' <remarks>
''' https://msdn.microsoft.com/ja-jp/library/office/ff838270.aspx
''' </remarks>
Public Enum WdExportFormat As Integer
    ''' <summary>
    ''' PDFファイル
    ''' </summary>
    ''' <remarks></remarks>
    wdExportFormatPDF = 17
    ''' <summary>
    ''' XPSファイル
    ''' </summary>
    ''' <remarks></remarks>
    wdExportFormatXPS
End Enum


''' <summary>
''' 変換品質
''' </summary>
''' <remarks>
''' https://msdn.microsoft.com/ja-jp/library/office/ff845865.aspx
''' </remarks>
Public Enum WdExportOptimizeFor As Integer
    ''' <summary>
    ''' 画質が細かくファイル サイズが大きい、印刷用
    ''' </summary>
    ''' <remarks></remarks>
    wdExportOptimizeForPrint
    ''' <summary>
    ''' 画質が粗くファイル サイズが小さい、画面用
    ''' </summary>
    ''' <remarks></remarks>
    wdExportOptimizeForOnScreen
End Enum

''' <summary>
''' エクスポートする文書の範囲
''' </summary>
''' <remarks>
''' https://msdn.microsoft.com/ja-jp/library/office/ff194747.aspx
''' </remarks>
Public Enum WdExportRange As Integer
    ''' <summary>
    ''' 文書全体
    ''' </summary>
    wdExportAllDocument
    ''' <summary>
    ''' 現在の選択範囲
    ''' </summary>
    wdExportSelection
    ''' <summary>
    ''' 現在のページ
    ''' </summary>
    wdExportCurrentPage
    ''' <summary>
    ''' 開始位置と終了位置を使用して、指定範囲
    ''' </summary>
    wdExportFromTo
End Enum

''' <summary>
''' 更履歴とコメントを含めて文書をエクスポートするかどうか
''' </summary>
''' <remarks>
''' https://msdn.microsoft.com/ja-jp/library/office/ff821431.aspx
''' </remarks>
Public Enum WdExportItem As Integer
    ''' <summary>
    ''' 変更履歴とコメントを含めずに
    ''' </summary>
    wdExportDocumentContent
    ''' <summary>
    ''' 変更履歴とコメントを含めて
    ''' </summary>
    wdExportDocumentWithMarkup = 7
End Enum

''' <summary>
''' 文書をエクスポートするときに含めるブックマーク
''' </summary>
Public Enum WdExportCreateBookmarks As Integer
    wdExportCreateNoBookmarks
    wdExportCreateHeadingBookmarks
    wdExportCreateWordBookmarks
End Enum

''' <summary>
''' クロの実行中に特定の警告とメッセージを処理する方法
''' </summary>
Public Enum WdAlertLevel As Integer
    wdAlertsMessageBox = -2
    wdAlertsAll = -1
    wdAlertsNone = 0
End Enum

''' <summary>
''' 保留中の変更をどのように処理するか
''' </summary>
Public Enum WdSaveOptions As Integer
    wdPromptToSaveChanges = -2
    wdSaveChanges = -1
    wdDoNotSaveChanges = 0
End Enum

''' <summary>
''' 文書形式を指定します。この列挙は通常、文書の保存時に使用
''' </summary>
Public Enum WdOriginalFormat As Integer
    wdWordDocument
    wdOriginalDocumentFormat
    wdPromptUser
End Enum
