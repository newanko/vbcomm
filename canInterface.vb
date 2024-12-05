

Imports System.Threading.Tasks
Imports System.Threading



''' <summary>
''' CANフレームクラス
''' </summary>
''' <remarks></remarks>
Public Class CANDFRAME

    Public Const CANID_MAX As UInteger = &H7FF
    Public Const CANDATALEN_MAX As UInteger = 8

    Public id As UInteger       'CAN-IDの数値(標準なら11bit,拡張なら29bit)
    Public data() As Byte       'データフレーム部(max 8byte)
    Public timeStamp As Long    'タイムスタンプ値
    Public dlc As Integer       'データ長(adta.lengthでいいとおもうが)
    Public rtr As Boolean
    Public ide As Boolean
    Public slot As Integer


    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="inid">CANID</param>
    ''' <param name="indata">データ本体(8バイト)</param>
    ''' <param name="ref_timestamp">タイムスタンプ値</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal inid As UInteger, ByVal indata() As Byte, ByVal ref_timestamp As Long)

        Me.id = inid
        If indata Is Nothing Then
            Me.data = New Byte(7) {0, 0, 0, 0, 0, 0, 0, 0}
        Else
            Me.data = IIf(indata.Length <= 8, indata, New Byte(7) {0, 0, 0, 0, 0, 0, 0, 0})
        End If
        Me.timeStamp = ref_timestamp

    End Sub


    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="inid">CANID</param>
    ''' <param name="indata">データ本体(8バイト)</param>
    ''' <param name="isRemote"></param>
    ''' <param name="isExt"></param>
    ''' <param name="times">タイムスタンプ値</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal inid As UInteger, ByVal indata() As Byte, ByVal isRemote As Boolean, ByVal isExt As Boolean, ByVal times As Long)

        Me.id = inid
        If indata Is Nothing Then
            Me.dlc = 8
        ElseIf indata.Length < 8 Then
            Me.dlc = indata.Length
        Else
            Me.dlc = 8
        End If
        Me.data = New Byte(Me.dlc - 1) {}
        For i = 0 To Me.data.Length - 1
            Me.data(i) = indata(i)
        Next
        Me.rtr = isRemote
        Me.ide = isExt
        Me.timeStamp = times

    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        Me.id = 0
        Me.data = New Byte(7) {0, 0, 0, 0, 0, 0, 0, 0}
        Me.timeStamp = 0

    End Sub

    ''' <summary>
    ''' CANフレームのデータのみを転写する
    ''' </summary>
    ''' <param name="inid"></param>
    ''' <param name="indata"></param>
    ''' <param name="times"></param>
    ''' <remarks></remarks>
    Public Sub Copy(ByVal inid As UInteger, ByVal indata() As Byte, ByVal ide_r As Boolean, ByVal rtr_r As Boolean, ByVal times As Long)

        Me.id = inid
        Me.timeStamp = times
        Dim i As Integer
        ReDim data(indata.Length - 1)
        For i = 0 To indata.Length - 1
            Me.data(i) = indata(i)
        Next

        Me.ide = ide_r
        Me.rtr = rtr_r

    End Sub

    ''' <summary>
    ''' CANフレームを文字列化する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function toString() As String

        Dim datastr As New System.Text.StringBuilder
        For i = 0 To Me.data.Length - 1
            datastr.Append(Me.data(i).ToString("X2") & ",") '00,11,22,33,44,55,66,77のようなカンマ区切り形式になる
        Next
        Return datastr.ToString.TrimEnd(",")

    End Function

End Class


''' <summary>
''' イベントで通知するCANデータクラス
''' </summary>
''' <remarks></remarks>
Public Class myEventArgs
    Inherits System.EventArgs

    Public data As CANDFRAME
    Public ret As Boolean

    ''' <summary>
    ''' 引数なしコンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Sub New()
        data = New CANDFRAME
        ret = True
    End Sub

    ''' <summary>
    ''' 引数付きコンストラクタ
    ''' </summary>
    ''' <param name="cand"></param>
    ''' <param name="cant"></param>
    ''' <remarks></remarks>
    Sub New(ByVal cand As CANDFRAME, ByVal cant As Boolean)
        data = cand
        ret = cant
    End Sub

End Class


''' <summary>
''' CAN通信クラスのインターフェース
''' </summary>
''' <remarks></remarks>
Public Interface canInterface

    ''' <summary>
    ''' 通信速度kHz
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property baudRate As Integer


    ''' <summary>
    ''' 特別な受信IDの定義
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property specialID As List(Of Integer)


    ''' <summary>
    ''' 使用可能なノード文字列配列
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getNodeNames() As String()


    ''' <summary>
    ''' ノードをオープンする
    ''' </summary>
    ''' <param name="port"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function open(ByVal port As String) As Boolean


    ''' <summary>
    ''' ノードをクローズする
    ''' </summary>
    ''' <remarks></remarks>
    Sub close()


    ''' <summary>
    ''' ポートがあいているか
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function isOpen() As Boolean


    ''' <summary>
    ''' 受信フィルタを設定する
    ''' </summary>
    ''' <param name="index"></param>
    ''' <param name="id"></param>
    ''' <param name="mask"></param>
    ''' <param name="ide"></param>
    ''' <param name="rtr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function setFilter(ByVal index As UShort, ByVal id As Integer, ByVal mask As Integer, ByVal ide As Boolean, ByVal rtr As Boolean) As Boolean


    ''' <summary>
    ''' データフレームを送信する
    ''' </summary>
    ''' <param name="cand"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Tx(ByVal cand As CANDFRAME) As Boolean


    ''' <summary>
    ''' データフレームの送受信
    ''' </summary>
    ''' <param name="txd"></param>
    ''' <param name="timeout"></param>
    ''' <param name="rxd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function TxRx(ByVal txd As CANDFRAME, ByVal timeout As UShort, ByRef rxd As CANDFRAME) As Boolean


    ''' <summary>
    ''' 受信イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Event OnComm(ByVal sender As Object, ByVal e As myEventArgs)


End Interface


''' <summary>
''' クバサーでの通信クラス
''' </summary>
''' <remarks></remarks>
Public Class KVSR_CAN
    Inherits CAN_Kvaser
    Implements canInterface



    Private sss As WindowsFormsSynchronizationContext

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ss"></param>
    ''' <remarks></remarks>
    Sub New(ByVal ss As WindowsFormsSynchronizationContext)
        sss = ss
    End Sub


    Public Property baudRate As Integer Implements canInterface.baudRate


    Public Shadows Property specialID As List(Of Integer) Implements canInterface.specialID
        Get
            Return MyBase.specialID
        End Get
        Set(value As List(Of Integer))
            MyBase.specialID = value
        End Set
    End Property



    Public Overloads Sub close() Implements canInterface.close
        MyBase.close()
    End Sub


    Public Function isopen() As Boolean Implements canInterface.isOpen
        Return True
    End Function



    ''' <summary>
    ''' ノード一覧
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getNodeNames() As String() Implements canInterface.getNodeNames
        Return MyBase.DeviceSearch
    End Function


    ''' <summary>
    ''' 受信イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Shadows Event OnComm(sender As Object, e As myEventArgs) Implements canInterface.OnComm


    ''' <summary>
    ''' 受信イベントバイパス
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="arg"></param>
    ''' <remarks></remarks>
    Private Sub CanDataRcv(ByVal sender As Object, ByVal arg As myEventArgs) Handles MyBase.OnComm

        RaiseEvent OnComm(Me, New myEventArgs(arg.data, True))

    End Sub


    ''' <summary>
    ''' オープン
    ''' </summary>
    ''' <param name="port"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function open(port As String) As Boolean Implements canInterface.open
        Return MyBase.open(port)
    End Function


    ''' <summary>
    ''' アクセプタンスフィルタ
    ''' </summary>
    ''' <param name="index"></param>
    ''' <param name="id"></param>
    ''' <param name="mask"></param>
    ''' <param name="ide"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function setFilter(index As UShort, id As Integer, mask As Integer, ide As Boolean, ByVal rtr As Boolean) As Boolean Implements canInterface.setFilter

        'Return MyBase.kvaserFilterSetting(id, mask, ide)
        Return True

    End Function


    ''' <summary>
    ''' 送信
    ''' </summary>
    ''' <param name="cand"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Tx(cand As CANDFRAME) As Boolean Implements canInterface.Tx
        Return MyBase.Tx(cand)
    End Function


    ''' <summary>
    ''' 送受信
    ''' </summary>
    ''' <param name="txd"></param>
    ''' <param name="timeout"></param>
    ''' <param name="rxd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function TxRx(txd As CANDFRAME, timeout As UShort, ByRef rxd As CANDFRAME) As Boolean Implements canInterface.TxRx
        Return MyBase.TxRx(txd, timeout, rxd)
    End Function


End Class


''' <summary>
''' 内製ジグでの通信クラス
''' </summary>
''' <remarks></remarks>
Public Class UART_CAN
    Implements canInterface


    ''' <summary>
    ''' １ビットのクオンタム設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Const bitQuantams As Integer = 16

    ''' <summary>
    ''' タイムセグメント2
    ''' </summary>
    ''' <remarks></remarks>
    Private Const tseg2 As Integer = 3

    ''' <summary>
    ''' タイムセグメント1
    ''' </summary>
    ''' <remarks></remarks>
    Private Const tseg1 As Integer = (bitQuantams - tseg2 - 1)

    ''' <summary>
    ''' シンクロナイゼーションジャンプ
    ''' </summary>
    ''' <remarks></remarks>
    Private Const sjw As Integer = 3


    ''' <summary>
    ''' 内製通信ジグインスタンス
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents comm As CAN_CAT.CAN_ON_JIG


    ''' <summary>
    ''' 送信データ
    ''' </summary>
    ''' <remarks></remarks>
    Private senddata As New CAN_CAT.CANDFRAME


    ''' <summary>
    ''' 最新の受信データ
    ''' </summary>
    ''' <remarks></remarks>
    Private lastdata As CANDFRAME


    ''' <summary>
    ''' 送受信成功フラグ
    ''' </summary>
    ''' <remarks></remarks>
    Private dataRecieved As Boolean


    ''' <summary>
    ''' 非同期受信イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="arg"></param>
    ''' <remarks></remarks>
    Private Sub CanDataRcv(ByVal sender As Object, ByVal arg As CAN_CAT.CAN_ON_JIG.myEventArgs) Handles comm.OnComm

        If Not (arg Is Nothing) Then
            If Not (arg.data Is Nothing) Then
                If (arg.data.id) <> senddata.id Then
                    Dim candata As CANDFRAME = New CANDFRAME(arg.data.id, arg.data.data, arg.data.rtr, arg.data.ide, arg.data.timeStamp)

                    If Me.specialID.Contains(candata.id) Then
                        lastdata = candata
                        dataRecieved = True
                    Else
                        RaiseEvent OnComm(Me, New myEventArgs(candata, True))
                    End If

                Else
                    dataRecieved = False
                End If
            End If
        End If


    End Sub


    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <param name="ss"></param>
    ''' <remarks></remarks>
    Sub New(ByVal ss As WindowsFormsSynchronizationContext)

        comm = New CAN_CAT.CAN_ON_JIG(ss)

    End Sub


    ''' <summary>
    ''' 通信速度
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property baudRate As Integer Implements canInterface.baudRate


    ''' <summary>
    ''' デバッグ用のID
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property specialID As New List(Of Integer) Implements canInterface.specialID


    ''' <summary>
    ''' ポート閉じる
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub close() Implements canInterface.close
        comm.close()
    End Sub


    ''' <summary>
    ''' ポート名取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getNodeNames() As String() Implements canInterface.getNodeNames

        Return System.IO.Ports.SerialPort.GetPortNames()

    End Function


    ''' <summary>
    ''' 受信割り込み
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event OnComm(sender As Object, e As myEventArgs) Implements canInterface.OnComm


    ''' <summary>
    ''' ポートopen
    ''' </summary>
    ''' <param name="port"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function open(port As String) As Boolean Implements canInterface.open
        Try
            If comm.open(port) Then
                Return True
            End If
        Catch ex As Exception

        End Try
        Return False
    End Function


    ''' <summary>
    ''' オープン状態確認
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isopen() As Boolean Implements canInterface.isOpen
        Return comm.isopen
    End Function


    ''' <summary>
    ''' アクセプタンスフィルタの設定
    ''' </summary>
    ''' <param name="index"></param>
    ''' <param name="id"></param>
    ''' <param name="mask"></param>
    ''' <param name="ide"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function setFilter(index As UShort, id As Integer, mask As Integer, ide As Boolean, ByVal rtr As Boolean) As Boolean Implements canInterface.setFilter
        Return comm.setFilter(index, id, mask, rtr, ide)
    End Function


    ''' <summary>
    ''' データフレームの送信
    ''' </summary>
    ''' <param name="cand"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Tx(cand As CANDFRAME) As Boolean Implements canInterface.Tx

        If Me.isopen Then
            Dim txdata As New CAN_CAT.CANDFRAME(cand.id, cand.data, cand.rtr, cand.ide, cand.timeStamp, 0)
            Return comm.tx(txdata)
        End If
        Return False

    End Function

    Dim st As New Stopwatch


    ''' <summary>
    ''' 特定の受信データを期待する送信アクション
    ''' </summary>
    ''' <param name="txd"></param>
    ''' <param name="timeout"></param>
    ''' <param name="rxd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TxRx(txd As CANDFRAME, timeout As UShort, ByRef rxd As CANDFRAME) As Boolean Implements canInterface.TxRx

        If Me.isopen Then
            dataRecieved = False
            senddata = New CAN_CAT.CANDFRAME(txd.id, txd.data, txd.rtr, txd.ide, txd.timeStamp, 0)

            st.Reset()
            st.Start()
            If comm.tx(senddata) Then

                Dim ktimer As New Timers.Timer
                ktimer.AutoReset = False
                ktimer.Interval = timeout
                ktimer.Enabled = True
                Do
                    Application.DoEvents()
                Loop While (dataRecieved = False) And (ktimer.Enabled = True)
                st.Stop()
                Debug.WriteLine("Elapsed Time=" & st.ElapsedMilliseconds)
                ktimer.Enabled = False
                ktimer.Dispose()

                If dataRecieved Then
                    rxd = lastdata
                    Return True
                End If

            End If
        End If

        Return False

    End Function


End Class






