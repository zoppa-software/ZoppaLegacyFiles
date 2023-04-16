Option Strict On
Option Explicit On

Namespace Csv

    ''' <summary>CSV形式でストリーム出力するためのクラスです。</summary>
    Public NotInheritable Class CsvStreamWriter
        Inherits SplitStreamWriter

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">対象ストリーム。</param>
        Public Sub New(stream As IO.Stream)
            MyBase.New(stream)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">対象パス。</param>
        Public Sub New(path As String)
            MyBase.New(path)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">対象ストリーム。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding)
            MyBase.New(stream, encoding)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">対象パス。</param>
        ''' <param name="append">追記モードならば真。</param>
        Public Sub New(path As String, append As Boolean)
            MyBase.New(path, append)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">対象ストリーム。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="bufferSize">バッファサイズ。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding, bufferSize As Integer)
            MyBase.New(stream, encoding, bufferSize)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">対象パス。</param>
        ''' <param name="append">追記モードならば真。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        Public Sub New(path As String, append As Boolean, encoding As Text.Encoding)
            MyBase.New(path, append, encoding)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">対象パス。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="bufferSize">バッファサイズ。</param>
        ''' <param name="leaveOpen">オープンフラグ。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding, bufferSize As Integer, leaveOpen As Boolean)
            MyBase.New(stream, encoding, bufferSize, leaveOpen)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">対象パス。</param>
        ''' <param name="append">追記モードならば真。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="bufferSize">バッファサイズ。</param>
        Public Sub New(path As String, append As Boolean, encoding As Text.Encoding, bufferSize As Integer)
            MyBase.New(path, append, encoding, bufferSize)
        End Sub

        ''' <summary>区切り文字を取得します。</summary>
        ''' <returns>区切り文字。</returns>
        Protected Overrides ReadOnly Property SplitStr As String
            Get
                Return ","
            End Get
        End Property

        ''' <summary>指定の文字列をエスケープした文字列へ変換して返します。</summary>
        ''' <param name="str">指定の文字列。</param>
        ''' <returns>エスケープした文字列。</returns>
        Protected Overrides Function Escape(str As String) As String
            ' エスケープが必要か判定
            Dim needEsc As Boolean = False
            If str.Length > 0 AndAlso Char.IsWhiteSpace(str(0)) Then
                needEsc = True
            ElseIf str.Length > 0 AndAlso Char.IsWhiteSpace(str(str.Length - 1)) Then
                needEsc = True
            Else
                For Each c In str
                    If c = ","c OrElse c = """"c OrElse c = ChrW(10) Then
                        needEsc = True
                        Exit For
                    End If
                Next
            End If

            ' エスケープが必要ならばエスケープして返す
            If needEsc Then
                str = str.Replace("""", """""")
                Return $"""{str}"""
            End If
            Return str
        End Function

    End Class

End Namespace
