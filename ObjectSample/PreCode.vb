Option Strict On
Option Explicit On
''' <summary>オブジェクトデータ型を表現するクラスです。</summary>

Public Class PreCode

    Public ReadOnly Value As String

    Public Sub New(value As String)
        Me.Value = value
    End Sub

End Class
