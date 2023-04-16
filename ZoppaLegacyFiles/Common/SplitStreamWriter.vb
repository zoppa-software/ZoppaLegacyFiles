Option Strict On
Option Explicit On

''' <summary>特定の区切り文字で分割した値を出力するためのストリームです。</summary>
Public MustInherit Class SplitStreamWriter
    Implements IDisposable

    ' 出力ストリーム
    Private mWriter As IO.StreamWriter

    ' カーソル位置は左端ならば真
    Private mCursorHome As Boolean = True

    ''' <summary>自動フラッシュを設定、取得します。</summary>
    Public Property AutoFlush As Boolean
        Get
            Return Me.mWriter.AutoFlush
        End Get
        Set(value As Boolean)
            Me.mWriter.AutoFlush = value
        End Set
    End Property

    ''' <summary>基底のストリームを取得します。</summary>
    Public ReadOnly Property BaseStream As IO.Stream
        Get
            Return Me.mWriter.BaseStream
        End Get
    End Property

    ''' <summary>ストリームのテキストエンコードを取得します。</summary>
    Public ReadOnly Property Encoding As Text.Encoding
        Get
            Return Me.mWriter.Encoding
        End Get
    End Property

    ''' <summary>改行文字を設定、取得します。</summary>
    Public Property NewLine As String
        Get
            Return Me.mWriter.NewLine
        End Get
        Set(value As String)
            Me.mWriter.NewLine = value
        End Set
    End Property

    ''' <summary>区切り文字を取得します。</summary>
    ''' <returns>区切り文字。</returns>
    Protected MustOverride ReadOnly Property SplitStr As String

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">対象ストリーム。</param>
    Public Sub New(stream As IO.Stream)
        Me.mWriter = New IO.StreamWriter(stream)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">対象パス。</param>
    Public Sub New(path As String)
        Me.mWriter = New IO.StreamWriter(path)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">対象ストリーム。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding)
        Me.mWriter = New IO.StreamWriter(stream, encoding)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">対象パス。</param>
    ''' <param name="append">追記モードならば真。</param>
    Public Sub New(path As String, append As Boolean)
        Me.mWriter = New IO.StreamWriter(path, append)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">対象ストリーム。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="bufferSize">バッファサイズ。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding, bufferSize As Integer)
        Me.mWriter = New IO.StreamWriter(stream, encoding, bufferSize)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">対象パス。</param>
    ''' <param name="append">追記モードならば真。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    Public Sub New(path As String, append As Boolean, encoding As Text.Encoding)
        Me.mWriter = New IO.StreamWriter(path, append, encoding)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">対象パス。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="bufferSize">バッファサイズ。</param>
    ''' <param name="leaveOpen">オープンフラグ。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding, bufferSize As Integer, leaveOpen As Boolean)
        Me.mWriter = New IO.StreamWriter(stream, encoding, bufferSize, leaveOpen)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">対象パス。</param>
    ''' <param name="append">追記モードならば真。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="bufferSize">バッファサイズ。</param>
    Public Sub New(path As String, append As Boolean, encoding As Text.Encoding, bufferSize As Integer)
        Me.mWriter = New IO.StreamWriter(path, append, encoding, bufferSize)
    End Sub

    ''' <summary>ストリームをクローズします。</summary>
    Public Sub Close()
        Me.mWriter.Close()
    End Sub

    ''' <summary>ストリームをフラッシュします。</summary>
    Public Sub Flush()
        Me.mWriter.Flush()
    End Sub

    ''' <summary>ストリームを非同期でフラッシュします。</summary>
    ''' <returns>タスクオブジェクト。</returns>
    Public Function FlushAsync() As Task
        Return Me.mWriter.FlushAsync()
    End Function

    ''' <summary>カーソルが先頭位置以外ならば、区切り文字とエスケープした文字列を取得します。</summary>
    ''' <param name="str">出力する文字列。</param>
    Private Sub JoinItem(str As String)
        If Not Me.mCursorHome Then
            mWriter.Write(Me.SplitStr)
        End If
        Me.mWriter.Write(Me.Escape(str))
        Me.mCursorHome = False
    End Sub

    ''' <summary>文字を出力します。</summary>
    ''' <param name="c">文字。</param>
    Public Sub Write(c As Char)
        Me.JoinItem(c)
    End Sub

    ''' <summary>数値を出力します。</summary>
    ''' <param name="dec">数値。</param>
    Public Sub Write(dec As Decimal)
        Me.JoinItem(dec.ToString())
    End Sub

    ''' <summary>数値を出力します。</summary>
    ''' <param name="dbl">数値。</param>
    Public Sub Write(dbl As Double)
        Me.JoinItem(dbl.ToString())
    End Sub

    ''' <summary>数値を出力します。</summary>
    ''' <param name="i32">数値。</param>
    Public Sub Write(i32 As Integer)
        Me.JoinItem(i32.ToString())
    End Sub

    ''' <summary>数値を出力します。</summary>
    ''' <param name="i64">数値。</param>
    Public Sub Write(i64 As Long)
        Me.JoinItem(i64.ToString())
    End Sub

    ''' <summary>オブジェクトを出力します。</summary>
    ''' <param name="obj">オブジェクト。</param>
    Public Sub Write(obj As Object)
        Me.JoinItem(obj.ToString())
    End Sub

    ''' <summary>数値を出力します。</summary>
    ''' <param name="sng">数値。</param>
    Public Sub Write(sng As Single)
        Me.JoinItem(sng.ToString())
    End Sub

    ''' <summary>文字列を出力します。</summary>
    ''' <param name="str">文字列。</param>
    Public Sub Write(str As String)
        Me.JoinItem(str)
    End Sub

    ''' <summary>数値を出力します。</summary>
    ''' <param name="ui32">数値。</param>
    Public Sub Write(ui32 As UInteger)
        Me.JoinItem(ui32.ToString())
    End Sub

    ''' <summary>数値を出力します。</summary>
    ''' <param name="ui64">数値。</param>
    Public Sub Write(ui64 As ULong)
        Me.JoinItem(ui64.ToString())
    End Sub

    ''' <summary>連続した項目を出力します。</summary>
    ''' <param name="items">項目リスト。</param>
    Public Sub Write(ParamArray items As Object())
        If items.Length > 0 Then
            If Not Me.mCursorHome Then
                Me.mWriter.Write(Me.SplitStr)
            End If
            Me.mWriter.Write(Me.Escape(items(0).ToString()))

            For i As Integer = 1 To items.Length - 1
                Me.mWriter.Write(Me.SplitStr)
                Me.mWriter.Write(Me.Escape(items(i).ToString()))
            Next
        End If
        Me.mCursorHome = False
    End Sub

    ''' <summary>カーソルが先頭位置以外ならば、区切り文字とエスケープした文字列を取得します。</summary>
    ''' <param name="str">出力する文字列。</param>
    Private Sub NewLineItem(str As String)
        If Not Me.mCursorHome Then
            mWriter.Write(Me.SplitStr)
        End If
        Me.mWriter.WriteLine(Me.Escape(str))
        Me.mCursorHome = True
    End Sub

    Public Sub WriteLine()
        Me.mWriter.WriteLine()
        Me.mCursorHome = True
    End Sub

    ''' <summary>文字を出力して改行します。</summary>
    ''' <param name="c">文字。</param>
    Public Sub WriteLine(c As Char)
        Me.NewLineItem(c)
    End Sub

    ''' <summary>数値を出力して改行します。</summary>
    ''' <param name="dec">数値。</param>
    Public Sub WriteLine(dec As Decimal)
        Me.NewLineItem(dec.ToString())
    End Sub

    ''' <summary>数値を出力して改行します。</summary>
    ''' <param name="dbl">数値。</param>
    Public Sub WriteLine(dbl As Double)
        Me.NewLineItem(dbl.ToString())
    End Sub

    ''' <summary>数値を出力して改行します。</summary>
    ''' <param name="i32">数値。</param>
    Public Sub WriteLine(i32 As Integer)
        Me.NewLineItem(i32.ToString())
    End Sub

    ''' <summary>数値を出力して改行します。</summary>
    ''' <param name="i64">数値。</param>
    Public Sub WriteLine(i64 As Long)
        Me.NewLineItem(i64.ToString())
    End Sub

    ''' <summary>オブジェクトを出力して改行します。</summary>
    ''' <param name="obj">オブジェクト。</param>
    Public Sub WriteLine(obj As Object)
        Me.NewLineItem(obj.ToString())
    End Sub

    ''' <summary>数値を出力して改行します。</summary>
    ''' <param name="sng">数値。</param>
    Public Sub WriteLine(sng As Single)
        Me.NewLineItem(sng.ToString())
    End Sub

    ''' <summary>文字列を出力して改行します。</summary>
    ''' <param name="str">文字列。</param>
    Public Sub WriteLine(str As String)
        Me.NewLineItem(str)
    End Sub

    ''' <summary>数値を出力して改行します。</summary>
    ''' <param name="ui32">数値。</param>
    Public Sub WriteLine(ui32 As UInteger)
        Me.NewLineItem(ui32.ToString())
    End Sub

    ''' <summary>数値を出力して改行します。</summary>
    ''' <param name="ui64">数値。</param>
    Public Sub WriteLine(ui64 As ULong)
        Me.NewLineItem(ui64.ToString())
    End Sub

    ''' <summary>連続した項目を出力して改行します。</summary>
    ''' <param name="items">項目リスト。</param>
    Public Sub WriteLine(ParamArray items As Object())
        If items.Length > 0 Then
            If Not Me.mCursorHome Then
                Me.mWriter.Write(Me.SplitStr)
            End If
            Me.mWriter.Write(Me.Escape(items(0).ToString()))

            For i As Integer = 1 To items.Length - 1
                Me.mWriter.Write(Me.SplitStr)
                Me.mWriter.Write(Me.Escape(items(i).ToString()))
            Next
            Me.mWriter.WriteLine()
        End If
        Me.mCursorHome = True
    End Sub

    ''' <summary>指定の文字列をエスケープした文字列へ変換して返します。</summary>
    ''' <param name="str">指定の文字列。</param>
    ''' <returns>エスケープした文字列。</returns>
    Protected MustOverride Function Escape(str As String) As String

    ''' <summary>リソースを解放します。</summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        Me.mWriter.Dispose()
    End Sub

    ''' <summary>インスタンスの文字列表現を取得します。</summary>
    ''' <returns>インスタンスの文字列表現。</returns>
    Public Overrides Function ToString() As String
        Return If(Me.mWriter.ToString(), "")
    End Function

End Class
