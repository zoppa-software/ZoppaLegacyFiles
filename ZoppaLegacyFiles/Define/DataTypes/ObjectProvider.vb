Option Strict On
Option Explicit On

''' <summary>オブジェクトデータ型のチェック、変換メソッドを登録するプロバイダです。</summary>
Public Module ObjectProvider

    ' チェックメソッド
    Private ReadOnly mCheckers As New Dictionary(Of Type, Func(Of String, Boolean))(8)

    ' 変換メソッド
    Private ReadOnly mCreaters As New Dictionary(Of Type, Func(Of String, Object))(8)

    ''' <summary>指定の型のチェックメソッドを設定します。</summary>
    ''' <typeparam name="T">指定の型。</typeparam>
    ''' <param name="method">チェックメソッド。</param>
    Public Sub SetChecker(Of T)(method As Func(Of String, Boolean))
        mCheckers.Add(GetType(T), method)
    End Sub

    ''' <summary>指定の型の変換メソッドを設定します。</summary>
    ''' <typeparam name="T">指定の型。</typeparam>
    ''' <param name="method">変換メソッド。</param>
    Public Sub SetCreater(Of T)(method As Func(Of String, T))
        mCreaters.Add(GetType(T), Function(input As String) method(input))
    End Sub

    ''' <summary>指定の型に変換できるかどうかを判定します。</summary>
    ''' <typeparam name="T">指定の型。</typeparam>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換可能ならば真。</returns>
    Friend Function CanTo(Of T)(input As String) As Boolean
        Try
            Dim method As Func(Of String, Boolean) = Nothing
            Return mCheckers.TryGetValue(GetType(T), method) AndAlso method(input)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>文字列を指定の型へ変換します。</summary>
    ''' <typeparam name="T">指定の型。</typeparam>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換後の型。</returns>
    Friend Function ToValue(Of T)(input As String) As T
        Dim method As Func(Of String, Object) = Nothing
        Return If(mCreaters.TryGetValue(GetType(T), method), CType(method(input), T), Nothing)
    End Function

    ''' <summary>設定した各メソッドを消去します。</summary>
    Public Sub Clear()
        mCheckers.Clear()
        mCreaters.Clear()
    End Sub

End Module
