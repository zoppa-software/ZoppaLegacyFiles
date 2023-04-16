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

        Public Property LineNo As Double Implements IIniLine.LineNo

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
        ''' <param name="row"></param>
        ''' <param name="str"></param>
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