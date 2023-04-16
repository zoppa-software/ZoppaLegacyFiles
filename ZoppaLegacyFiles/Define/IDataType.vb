Option Strict On
Option Explicit On

''' <summary>データの型を表現するインターフェイスです。</summary>
Public Interface IDataType

    ''' <summary>変換する型を取得します。</summary>
    ''' <returns>変換する型。</returns>
    ReadOnly Property DotNetType As Type

    ''' <summary>文字列を変換できるかどうかを判定します。</summary>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換できる場合は<c>True</c>、そうでない場合は<c>False</c>。</returns>
    Function CanTo(input As String) As Boolean

    ''' <summary>分割項目を変換できるかどうかを判定します。</summary>
    ''' <param name="input">変換する分割項目。</param>
    ''' <returns>変換できる場合は<c>True</c>、そうでない場合は<c>False</c>。</returns>
    Function CanTo(input As IFileItem) As Object


    ''' <summary>文字列を変換します。</summary>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換された値。</returns>
    Function ToValue(input As String) As Object

    ''' <summary>分割項目を変換します。</summary>
    ''' <param name="input">変換する分割項目。</param>
    ''' <returns>変換された値。</returns>
    Function ToValue(input As IFileItem) As Object

End Interface
