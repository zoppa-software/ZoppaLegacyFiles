﻿Option Strict On
Option Explicit On

''' <summary>バイトのデータ型を表現するクラスです。</summary>
Public NotInheritable Class ByteDataType
    Inherits DataType

    ''' <summary>変換する型を取得します。</summary>
    ''' <returns>変換する型。</returns>
    Public Overrides ReadOnly Property DotNetType As Type
        Get
            Return GetType(Byte?)
        End Get
    End Property

    ''' <summary>コンストラクタ。</summary>
    Friend Sub New()

    End Sub

    ''' <summary>文字列を変換できるかどうかを判定します。</summary>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換できる場合は<c>True</c>、そうでない場合は<c>False</c>。</returns>
    Public Overrides Function CanTo(input As String) As Boolean
        If input Is Nothing OrElse input = "" Then
            Return True
        Else
            Dim tmp As Byte
            Return Byte.TryParse(input, tmp)
        End If
    End Function

    ''' <summary>文字列を変換します。</summary>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換された値。</returns>
    Public Overrides Function ToValue(input As String) As Object
        If input Is Nothing OrElse input = "" Then
            Return Nothing
        Else
            Dim tmp As Byte
            If Byte.TryParse(input, tmp) Then
                Return tmp
            Else
                Throw New FormatException($"unable to convert ""{input}"" to byte")
            End If
        End If
    End Function

End Class
