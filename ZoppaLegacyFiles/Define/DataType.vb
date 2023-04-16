Option Strict On
Option Explicit On

''' <summary>データの型を表現するクラスです。</summary>
Public MustInherit Class DataType
    Implements IDataType

    ''' <summary>変換する型を取得します。</summary>
    ''' <returns>変換する型。</returns>
    Public MustOverride ReadOnly Property DotNetType As Type Implements IDataType.DotNetType

    ''' <summary>分割項目を変換できるかどうかを判定します。</summary>
    ''' <param name="input">変換する分割項目。</param>
    ''' <returns>変換できる場合は<c>True</c>、そうでない場合は<c>False</c>。</returns>
    Public Function CanTo(input As IFileItem) As Object Implements IDataType.CanTo
        Return Me.CanTo(input.UnEscape)
    End Function

    ''' <summary>文字列を変換できるかどうかを判定します。</summary>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換できる場合は<c>True</c>、そうでない場合は<c>False</c>。</returns>
    Public MustOverride Function CanTo(input As String) As Boolean Implements IDataType.CanTo

    ''' <summary>分割項目を変換します。</summary>
    ''' <param name="input">変換する分割項目。</param>
    ''' <returns>変換された値。</returns>
    Public Function ToValue(input As IFileItem) As Object Implements IDataType.ToValue
        Return Me.ToValue(input.UnEscape)
    End Function

    ''' <summary>文字列を変換します。</summary>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換された値。</returns>
    Public MustOverride Function ToValue(input As String) As Object Implements IDataType.ToValue

End Class
