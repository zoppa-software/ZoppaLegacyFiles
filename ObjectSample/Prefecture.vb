Option Strict On
Option Explicit On

Public Class Prefecture

    Public ReadOnly area As AreaEnum

    Public ReadOnly no As Integer

    Public ReadOnly code As PreCode

    Public ReadOnly name As String

    Public ReadOnly yomi As String

    Public ReadOnly center As String

    Public ReadOnly centerYomi As String

    ''' <summary>各 CSVの列に対応したパラメータを定義したコンストラクタです。</summary>
    ''' <param name="area">地方区分</param>
    ''' <param name="no">都道府県番号</param>
    ''' <param name="code">団体コード</param>
    ''' <param name="name">都道府県名</param>
    ''' <param name="yomi">都道府県名_よみ</param>
    ''' <param name="center">県庁所在地</param>
    ''' <param name="centerYomi">県庁所在地_よみ</param>
    Public Sub New(area As AreaEnum,
                   no As Integer,
                   code As PreCode,
                   name As String,
                   yomi As String,
                   center As String,
                   centerYomi As String)
        Me.area = area
        Me.no = no
        Me.code = code
        Me.name = name
        Me.yomi = yomi
        Me.center = center
        Me.centerYomi = centerYomi
    End Sub

End Class
