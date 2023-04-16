Option Strict On
Option Explicit On

Imports System.Globalization

''' <summary>Dateのデータ型を表現するクラスです。</summary>
Public NotInheritable Class DateTimeDataType
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
            Return GetType(Date?)
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
            Dim dt As Date
            If Me.CustomFormat?.Trim() <> "" Then
                Return DateTime.TryParseExact(input, Me.CustomFormat, CultureInfo.CurrentCulture, DateTimeStyles.None, dt)
            Else
                Return DateTime.TryParse(input, dt)
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
            Dim dt As DateTime
            If Me.CustomFormat?.Trim() <> "" Then
                If DateTime.TryParseExact(input, Me.CustomFormat, CultureInfo.CurrentCulture, DateTimeStyles.None, dt) Then
                    Return dt
                End If
            ElseIf DateTime.TryParse(input, dt) Then
                Return dt
            End If
            Throw New FormatException($"unable to convert ""{input}"" to DateTime")
        End If
    End Function

End Class