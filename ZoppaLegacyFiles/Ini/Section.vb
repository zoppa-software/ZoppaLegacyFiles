Option Strict On
Option Explicit On

Namespace Ini

    ''' <summary>セクションを表現します。</summary>
    Friend NotInheritable Class Section
        Implements IIniLine

        ''' <summary>セクションが無名セクションならば真を返します。</summary>
        Public ReadOnly Property DefaultSection As Boolean

        ''' <summary>セクションの名前を取得します。</summary>
        Public ReadOnly Property Name As String

        ''' <summary>行位置を取得します（途中挿入用に小数を指定可）</summary>
        ''' <returns>行位置。</returns>
        Public Property LineNo As Double Implements IIniLine.LineNo

        ''' <summary>書き込み用文字列を取得します。</summary>
        ''' <returns>書き込み文字列。</returns>
        Public ReadOnly Property WriteStr As String Implements IIniLine.WriteStr

        ''' <summary>コンストラクタ。</summary>
        Public Sub New()
            Me.DefaultSection = True
            Me.Name = ""
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="secname">セクション名。</param>
        Public Sub New(secname As String)
            Me.DefaultSection = False
            Me.Name = secname
            Me.LineNo = -1
            Me.WriteStr = $"[{secname}]"
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="secname">セクション名。</param>
        ''' <param name="row">行位置。</param>
        ''' <param name="str">書き込む文字列。</param>
        Public Sub New(secname As String, row As Integer, str As String)
            Me.DefaultSection = False
            Me.Name = secname
            Me.LineNo = row
            Me.WriteStr = str
        End Sub

        ''' <summary>オブジェクトが一致するならば真を返します。</summary>
        ''' <param name="obj">比較するオブジェクト。</param>
        ''' <returns>一致するならば真。</returns>
        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is Section Then
                Dim other = CType(obj, Section)
                Return (Me.DefaultSection = other.DefaultSection AndAlso Me.Name = other.Name)
            Else
                Return False
            End If
        End Function

        ''' <summary>ハッシュコード値を取得します。</summary>
        ''' <returns>ハッシュコード値。</returns>
        Public Overrides Function GetHashCode() As Integer
            Return Me.DefaultSection.GetHashCode() Xor Me.Name.GetHashCode()
        End Function

    End Class

End Namespace