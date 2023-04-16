Option Strict On
Option Explicit On

''' <summary>ライブラリ共通例外。</summary>
Public Class LegacyFilesException
    Inherits Exception

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="message">例外メッセージ。</param>
    Public Sub New(message As String)
        MyBase.New(message)
    End Sub

End Class