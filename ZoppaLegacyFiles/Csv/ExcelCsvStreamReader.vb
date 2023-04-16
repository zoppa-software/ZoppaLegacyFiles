Option Strict On
Option Explicit On

Namespace Csv

    ''' <summary>カンマ区切りファイル読み込みストリームです（EXCEL）</summary>
    Public NotInheritable Class ExcelCsvStreamReader
        Inherits SplitStreamReader(Of ExcelCsvSpliter, ExcelCsvItem)

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        Public Sub New(stream As IO.Stream)
            MyBase.New(stream)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">入力ファイルパス。</param>
        Public Sub New(path As String)
            MyBase.New(path)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        Public Sub New(stream As IO.Stream, detectEncodingFromByteOrderMarks As Boolean)
            MyBase.New(stream, detectEncodingFromByteOrderMarks)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding)
            MyBase.New(stream, encoding)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">入力ファイルパス。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        Public Sub New(path As String, detectEncodingFromByteOrderMarks As Boolean)
            MyBase.New(path, detectEncodingFromByteOrderMarks)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">入力ファイルパス。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        Public Sub New(path As String, encoding As Text.Encoding)
            MyBase.New(path, encoding)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">入力ファイルパス。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        Public Sub New(path As String, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean)
            MyBase.New(path, encoding, detectEncodingFromByteOrderMarks)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        ''' <param name="bufferSize">バッファサイズ。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">入力ファイルパス。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        ''' <param name="bufferSize">バッファサイズ。</param>
        Public Sub New(path As String, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer)
            MyBase.New(path, encoding, detectEncodingFromByteOrderMarks, bufferSize)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        ''' <param name="bufferSize">バッファサイズ。</param>
        ''' <param name="leaveOpen">ストリームを開いたままにするならば真。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, leaveOpen As Boolean)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize, leaveOpen)
        End Sub

    End Class

End Namespace
