Option Strict On
Option Explicit On

Namespace Ini

    ''' <summary>INIファイル取得値を表現します。</summary>
    Public Structure ValueResult
        Implements IFileItem

        ''' <summary>指定したセクション、キーの値があれば真を返します。</summary>
        Public ReadOnly Property IsSome As Boolean

        ''' <summary>エスケープ解除後の文字列。</summary>
        Public ReadOnly Property UnEscape As String Implements IFileItem.UnEscape

        ''' <summary>エスケープ解除前の文字列。</summary>
        Public ReadOnly Property Text As String Implements IFileItem.Text

        ''' <summary>生の文字配列を取得します。</summary>
        ''' <returns>生の文字配列。</returns>
        Public ReadOnly Property Raw As ReadOnlyMemory Implements IFileItem.Raw
            Get
                Throw New NotSupportedException("生の文字列を取得出来ません")
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="sm">有効な値ならば真。</param>
        ''' <param name="unEsc">エスケープ後の文字列。</param>
        ''' <param name="text">エスケープ前の文字列。</param>
        Public Sub New(sm As Boolean, unEsc As String, text As String)
            Me.IsSome = sm
            Me.UnEscape = unEsc
            Me.Text = text
        End Sub

        ''' <summary>指定した型で値を取得します。</summary>
        ''' <typeparam name="T">指定した型。</typeparam>
        ''' <returns>値。</returns>
        Public Function Convert(Of T)() As T
            Dim conv = DataTypeConverter.Convert(GetType(T))
            Return CType(conv.ToValue(Me.UnEscape), T)
        End Function

    End Structure

End Namespace
