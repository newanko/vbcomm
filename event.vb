

Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Text
Imports System.Net
Imports System.Console
Imports System.Timers



''' <summary>
''' </summary>
''' <remarks></remarks>
Public Class Form1


    ''' <summary>
    ''' 通信インターフェース
    ''' </summary>
    ''' <remarks></remarks>
    Public WithEvents CAN As canInterface

    Private kcan As KVSR_CAN


    ''' <summary>
    ''' CAN受信
    ''' </summary>
    ''' <remarks></remarks>
    Private Async Sub CanDataRcv(ByVal sender As Object, ByVal arg As myEventArgs) Handles CAN.OnComm

        If Not (arg Is Nothing) Then
            If Not (arg.data Is Nothing) Then

                Dim candf As CANDFRAME = arg.data

                If CheckBox_viewEnable.Checked = True Then
                    dispcan(candf)
                End If

                If CheckBox_charger_enable.Checked Then
                    pargseCharger(candf)
                End If

                If dbgdata.acceptanceID.Contains(candf.id) Then
                    canParseDataDebug(candf)
                Else
                    canParseData(candf)
                    Await canParseDataAsync(candf)
                End If



            End If
        End If

    End Sub


    ''' <summary>
    ''' フォームコンストラクタ
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        kcan = New KVSR_CAN(SynchronizationContext.Current)
        canInstance_init(kcan)


    End Sub



End Class
