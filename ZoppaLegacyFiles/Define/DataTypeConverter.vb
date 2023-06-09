﻿Option Strict On
Option Explicit On

''' <summary>データ型を取得するクラスです。</summary>
Public Module DataTypeConverter

    ' BooleanDataType
    Private ReadOnly mBooleanData As New Lazy(Of BooleanDataType)(Function() New BooleanDataType())

    ' ByteArrayDataType
    Private ReadOnly mByteArrayData As New Lazy(Of ByteArrayDataType)(Function() New ByteArrayDataType())

    ' ByteDataType
    Private ReadOnly mByteData As New Lazy(Of ByteDataType)(Function() New ByteDataType())

    ' DateTimeDataType
    Private ReadOnly mDateTimeData As New Lazy(Of DateTimeDataType)(Function() New DateTimeDataType())

    ' DecimalDataType
    Private ReadOnly mDecimalData As New Lazy(Of DecimalDataType)(Function() New DecimalDataType())

    ' DoubleDataType
    Private ReadOnly mDoubleData As New Lazy(Of DoubleDataType)(Function() New DoubleDataType())

    ' IntegerDataType
    Private ReadOnly mIntegerData As New Lazy(Of IntegerDataType)(Function() New IntegerDataType())

    ' LongDataType
    Private ReadOnly mLongData As New Lazy(Of LongDataType)(Function() New LongDataType())

    ' ShortDataType
    Private ReadOnly mShortData As New Lazy(Of ShortDataType)(Function() New ShortDataType())

    ' SingleDataType
    Private ReadOnly mSingleData As New Lazy(Of SingleDataType)(Function() New SingleDataType())

    ' StringDataType
    Private ReadOnly mStringData As New Lazy(Of StringDataType)(Function() New StringDataType())

    ' TimeSpanDataType
    Private ReadOnly mTimeSpanData As New Lazy(Of TimeSpanDataType)(Function() New TimeSpanDataType())

    ''' <summary>BooleanDataType を取得します。</summary>
    ''' <returns>BooleanDataType。</returns>
    Public ReadOnly Property BooleanData As BooleanDataType
        Get
            Return mBooleanData.Value
        End Get
    End Property

    ''' <summary>ByteArrayDataType を取得します。</summary>
    ''' <returns>ByteArrayDataType。</returns>
    Public ReadOnly Property ByteArrayData As ByteArrayDataType
        Get
            Return mByteArrayData.Value
        End Get
    End Property

    ''' <summary>ByteDataType を取得します。</summary>
    ''' <returns>ByteDataType。</returns>
    Public ReadOnly Property ByteData As ByteDataType
        Get
            Return mByteData.Value
        End Get
    End Property

    ''' <summary>DateTimeDataType を取得します。</summary>
    ''' <returns>DateTimeDataType。</returns>
    Public ReadOnly Property DateTimeData As DateTimeDataType
        Get
            Return mDateTimeData.Value
        End Get
    End Property

    ''' <summary>DecimalDataType を取得します。</summary>
    ''' <returns>DecimalDataType。</returns>
    Public ReadOnly Property DecimalData As DecimalDataType
        Get
            Return mDecimalData.Value
        End Get
    End Property

    ''' <summary>DoubleDataType を取得します。</summary>
    ''' <returns>DoubleDataType。</returns>
    Public ReadOnly Property DoubleData As DoubleDataType
        Get
            Return mDoubleData.Value
        End Get
    End Property

    ''' <summary>IntegerDataType を取得します。</summary>
    ''' <returns>IntegerDataType。</returns>
    Public ReadOnly Property IntegerData As IntegerDataType
        Get
            Return mIntegerData.Value
        End Get
    End Property

    ''' <summary>LongDataType を取得します。</summary>
    ''' <returns>LongDataType。</returns>
    Public ReadOnly Property LongData As LongDataType
        Get
            Return mLongData.Value
        End Get
    End Property

    ''' <summary>ShortDataType を取得します。</summary>
    ''' <returns>ShortDataType。</returns>
    Public ReadOnly Property ShortData As ShortDataType
        Get
            Return mShortData.Value
        End Get
    End Property

    ''' <summary>SingleDataType を取得します。</summary>
    ''' <returns>SingleDataType。</returns>
    Public ReadOnly Property SingleData As SingleDataType
        Get
            Return mSingleData.Value
        End Get
    End Property

    ''' <summary>StringDataType を取得します。</summary>
    ''' <returns>StringDataType。</returns>
    Public ReadOnly Property StringData As StringDataType
        Get
            Return mStringData.Value
        End Get
    End Property

    ''' <summary>TimeSpanDataType を取得します。</summary>
    ''' <returns>TimeSpanDataType。</returns>
    Public ReadOnly Property TimeSpanData As TimeSpanDataType
        Get
            Return mTimeSpanData.Value
        End Get
    End Property

    ''' <summary>指定した型の DataType を取得します。</summary>
    ''' <param name="itemType">.netの型。</param>
    ''' <returns>DataType。</returns>
    Public Function Convert(itemType As Type) As IDataType
        Select Case itemType
            Case GetType(Boolean), GetType(Boolean?)
                Return BooleanData
            Case GetType(Byte())
                Return ByteArrayData
            Case GetType(Byte), GetType(Byte?)
                Return ByteData
            Case GetType(DateTime), GetType(DateTime?)
                Return DateTimeData
            Case GetType(Decimal), GetType(Decimal?)
                Return DecimalData
            Case GetType(Double), GetType(Double?)
                Return DoubleData
            Case GetType(Integer), GetType(Integer?)
                Return IntegerData
            Case GetType(Long), GetType(Long?)
                Return LongData
            Case GetType(Short), GetType(Short?)
                Return ShortData
            Case GetType(Single), GetType(Single?)
                Return SingleData
            Case GetType(String)
                Return StringData
            Case GetType(TimeSpan), GetType(TimeSpan?)
                Return TimeSpanData
            Case Else
                If itemType.IsEnum Then
                    Dim tp = Type.GetType($"{GetType(DataType).Namespace}.EnumDataType`1[[{itemType.AssemblyQualifiedName}]]")
                    Return TryCast(Activator.CreateInstance(tp), IDataType)
                ElseIf itemType.IsClass OrElse itemType.IsValueType Then
                    Dim tp = Type.GetType($"{GetType(DataType).Namespace}.ObjectDataType`1[[{itemType.AssemblyQualifiedName}]]")
                    Return TryCast(Activator.CreateInstance(tp), IDataType)
                End If
        End Select

        Throw New InvalidCastException($"{itemType.Name} は DataType として定義していません")
    End Function

End Module
