Option Strict On
Option Explicit On

''' <summary>読み込み専用メモリ。</summary>
Public Structure ReadOnlyMemory

    ''' <summary>参照する文字配列。</summary>
    Private ReadOnly mItems As Char()

    ''' <summary>開始位置。</summary>
    Private ReadOnly mStart As Integer

    ''' <summary>文字数。</summary>
    Private ReadOnly mLength As Integer

    ''' <summary>文字数を取得します。</summary>
    ''' <returns>文字数。</returns>
    Public ReadOnly Property Length() As Integer
        Get
            Return If(Me.mItems IsNot Nothing, Me.mLength, 0)
        End Get
    End Property

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="item">参照文字配列。</param>
    Public Sub New(item As Char())
        Me.mItems = item
        Me.mStart = 0
        Me.mLength = item.Length
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="source">元にした読み込み専用メモリ。</param>
    ''' <param name="start">新しい開始位置。</param>
    ''' <param name="length">新しい文字数。</param>
    Public Sub New(source As ReadOnlyMemory, start As Integer, length As Integer)
        Me.mItems = source.mItems
        Me.mStart = If(source.mStart <= start, start, source.mStart)
        Dim maxlen = source.mStart + source.mLength
        Me.mLength = If(maxlen >= Me.mStart + length, length, maxlen - Me.mStart)
    End Sub

    ''' <summary>読み込み専用メモリから部分を取得します。</summary>
    ''' <param name="start">部分開始位置。</param>
    ''' <param name="length">部分文字数。</param>
    ''' <returns>部分の読み込み専用メモリ。</returns>
    Public Function Slice(start As Integer, length As Integer) As ReadOnlyMemory
        Dim ln = If(Me.mLength - start >= length, length, Me.mLength - start)
        Return New ReadOnlyMemory(Me, Me.mStart + start, ln)
    End Function

    ''' <summary>文字列表現を取得します。</summary>
    ''' <returns>文字列表現。</returns>
    Public Overrides Function ToString() As String
        Return If(Me.mItems IsNot Nothing, New String(Me.mItems, Me.mStart, Me.mLength), "")
    End Function

End Structure