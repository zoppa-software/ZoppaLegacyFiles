Option Strict On
Option Explicit On

Namespace Ini

    ''' <summary>値、セクションではない行を示します（空の行、コメントのみの行など）</summary>
    Friend NotInheritable Class EmptyLine
        Implements IIniLine

        ''' <summary>行位置を取得します（途中挿入用に小数を指定可）</summary>
        ''' <returns>行位置。</returns>
        Public Property LineNo As Double Implements IIniLine.LineNo

        ''' <summary>書き込み用文字列を取得します。</summary>
        ''' <returns>書き込み文字列。</returns>
        Public ReadOnly Property WriteStr As String Implements IIniLine.WriteStr

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="row">行位置。</param>
        ''' <param name="str">書き込む文字列。</param>
        Public Sub New(row As Integer, str As String)
            Me.LineNo = row
            Me.WriteStr = str
        End Sub
    End Class

End Namespace