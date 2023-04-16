Option Strict On
Option Explicit On

''' <summary>ファイルの項目情報。</summary>
Public Interface IFileItem

    ''' <summary>エスケープを解除した文字列を返す。</summary>
    ''' <returns><エスケープを解除した文字列。/returns>
    ReadOnly Property UnEscape As String

    ''' <summary>項目の文字列を取得する。</summary>
    ''' <returns>項目の文字列。</returns>
    ReadOnly Property Text As String

    ''' <summary>生の文字配列を取得します。</summary>
    ''' <returns>生の文字配列。</returns>
    ReadOnly Property Raw As ReadOnlyMemory

End Interface
