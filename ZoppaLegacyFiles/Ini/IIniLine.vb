Option Strict On
Option Explicit On

Namespace Ini

    ''' <summary>行情報を表現するインターフェイスです。</summary>
    Friend Interface IIniLine

        ''' <summary>行位置を取得します（途中挿入用に小数を指定可）</summary>
        ''' <returns>行位置。</returns>
        Property LineNo As Double

        ''' <summary>書き込み用文字列を取得します。</summary>
        ''' <returns>書き込み文字列。</returns>
        ReadOnly Property WriteStr As String

    End Interface

End Namespace
