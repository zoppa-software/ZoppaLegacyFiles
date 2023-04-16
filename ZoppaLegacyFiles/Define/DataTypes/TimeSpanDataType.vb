Option Strict On
Option Explicit On

Imports System.Globalization

''' <summary>TimeSpanのデータ型を表現するクラスです。</summary>
Public NotInheritable Class TimeSpanDataType
    Inherits DataType

    ' カスタムフォーマット
    Private mCustomFormat As String = ""

    ''' <summary>デフォルトフォーマットを取得または設定します。</summary>
    ''' <returns>フォーマット。</returns>
    Public Shared Property DefaultFormat As String = ""

    ''' <summary>カスタムフォーマットを取得または設定します。</summary>
    ''' <returns>フォーマット。</returns>
    Public ReadOnly Property CustomFormat As String
        Get
            If Me.mCustomFormat?.Trim() <> "" Then
                Return Me.mCustomFormat
            Else
                Return DefaultFormat
            End If
        End Get
    End Property

    ''' <summary>変換する型を取得します。</summary>
    ''' <returns>変換する型。</returns>
    Public Overrides ReadOnly Property DotNetType As Type
        Get
            Return GetType(TimeSpan?)
        End Get
    End Property

    ''' <summary>新しいインスタンスを初期化します。</summary>
    Friend Sub New()

    End Sub

    ''' <summary>新しいインスタンスを初期化します。</summary>
    ''' <param name="format">カスタムフォーマット。</param>
    Public Sub New(format As String)
        Me.mCustomFormat = format
    End Sub

    ''' <summary>文字列を変換できるかどうかを判定します。</summary>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換できる場合は<c>True</c>、そうでない場合は<c>False</c>。</returns>
    Public Overrides Function CanTo(input As String) As Boolean
        If input Is Nothing OrElse input = "" Then
            Return True
        Else
            Dim tm As TimeSpan
            If Me.CustomFormat?.Trim() <> "" Then
                Return TimeSpan.TryParseExact(input, Me.CustomFormat, CultureInfo.CurrentCulture, TimeSpanStyles.None, tm)
            Else
                Return TimeSpan.TryParse(input, CultureInfo.CurrentCulture, tm)
            End If
        End If
    End Function

    ''' <summary>文字列を変換します。</summary>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換された値。</returns>
    Public Overrides Function ToValue(input As String) As Object
        If input Is Nothing OrElse input = "" Then
            Return Nothing
        Else
            Dim tm As TimeSpan
            If Me.CustomFormat?.Trim() <> "" Then
                If TimeSpan.TryParseExact(input, Me.CustomFormat, CultureInfo.CurrentCulture, TimeSpanStyles.None, tm) Then
                    Return tm
                End If
            ElseIf TimeSpan.TryParse(input, CultureInfo.CurrentCulture, tm) Then
                Return tm
            End If
            Throw New FormatException($"unable to convert ""{input}"" to TimeSpan")
        End If
    End Function

End Class