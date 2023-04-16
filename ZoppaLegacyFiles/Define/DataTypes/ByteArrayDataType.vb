Option Strict On
Option Explicit On

''' <summary>バイト配列のデータ型を表現するクラスです。</summary>
Public NotInheritable Class ByteArrayDataType
    Inherits DataType

    ' チェックメソッドを設定するメソッド
    Private Shared ReadOnly mSync As New Key()

    ' チェックメソッド
    Private Shared mCheckMethod As Func(Of String, Boolean) = Nothing

    ' 変換メソッド
    Private Shared mConvertMethod As Func(Of String, Byte()) = Nothing

    ''' <summary>変換する型を取得します。</summary>
    ''' <returns>変換する型。</returns>
    Public Overrides ReadOnly Property DotNetType As Type
        Get
            Return GetType(Byte())
        End Get
    End Property

    ''' <summary>コンストラクタ。</summary>
    Friend Sub New()

    End Sub

    ''' <summary>
    ''' チェックメソッドと変換メソッドを設定します。
    ''' 両方とも<c>Nothing</c>を設定すると、デフォルトのチェックメソッドに戻ります。
    ''' </summary>
    ''' <param name="checkMethod">チェックメソッド。</param>
    ''' <param name="convertMethod">変換メソッド。</param>
    Public Shared Sub SetConvertMethod(checkMethod As Func(Of String, Boolean), convertMethod As Func(Of String, Byte()))
        If (checkMethod IsNot Nothing AndAlso convertMethod IsNot Nothing) OrElse
           (checkMethod Is Nothing AndAlso convertMethod Is Nothing) Then
            mCheckMethod = checkMethod
            mConvertMethod = convertMethod
        End If
    End Sub

    ''' <summary>文字列を変換できるかどうかを判定します。</summary>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換できる場合は<c>True</c>、そうでない場合は<c>False</c>。</returns>
    Public Overrides Function CanTo(input As String) As Boolean
        If input Is Nothing OrElse input = "" Then
            Return True
        ElseIf mCheckMethod IsNot Nothing Then
            SyncLock mSync
                Return mCheckMethod(input)
            End SyncLock
        Else
            For Each c In input.ToCharArray()
                Select Case c
                    Case "0"c To "9"c

                    Case "a"c To "f"c

                    Case "A"c To "F"c

                    Case Else
                        Return False
                End Select
            Next
            Return True
        End If
    End Function

    ''' <summary>文字列を変換します。</summary>
    ''' <param name="input">変換する文字列。</param>
    ''' <returns>変換された値。</returns>
    Public Overrides Function ToValue(input As String) As Object
        If input Is Nothing OrElse input = "" Then
            Return Nothing
        ElseIf mConvertMethod IsNot Nothing Then
            SyncLock mSync
                Return mConvertMethod(input)
            End SyncLock
        Else
            Dim ans As New List(Of Byte)()

            Dim buf = New Char(1) {}
            Using sr As New IO.StringReader(If(input.Length Mod 2 = 0, input, "0" & input))
                Do While sr.Peek() <> -1
                    sr.Read(buf, 0, 2)

                    Dim b = 0
                    For i As Integer = 0 To 1
                        Select Case buf(i)
                            Case "0"c To "9"c
                                b = b * 16 + (AscW(buf(i)) - AscW("0"c))
                            Case "a"c To "f"c
                                b = b * 16 + 10 + (AscW(buf(i)) - AscW("a"c))
                            Case "A"c To "F"c
                                b = b * 16 + 10 + (AscW(buf(i)) - AscW("A"c))
                        End Select
                    Next
                    ans.Add(CByte(b))
                Loop
            End Using

            Return ans.ToArray()
        End If
    End Function

    ''' <summary>同期ロックインスタンス。</summary>
    Private NotInheritable Class Key

    End Class

End Class