Imports System

Namespace test

    Public Structure Pixel
        Dim index As Integer
        Dim size As Integer
    End Structure

    Public Class Form1
        Inherits System.Windows.Forms.Form

        Public Enum InterfaceColors
            Mis tyRose = &HE1E4FF&
            SlateGray = &H908070&
            DodgerBlue = &HFF901E&
            DeepSkyBlue = &HFFBF00&
        End Enum

        <Flags()> Public Enum FilePermissions As Integer
            None = 0
            Create = 1
            Read = 2
            Update = 4
            Delete = 8
        End Enum

        Function Update(ByVal thisSale As Decimal) As Decimal
            Static totalSales As Decimal = 0
            totalSales += thisSale
            Return totalSales
        End Function

        Public DeviceCount As Integer

        Private _count As Integer
        Public Property Number() As Double
            Get
                Return _count
            End Get
            Set(ByVal value As Integer)
                _count = value
            End Set
        End Property

        Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (
    ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
        Sub GetUser()
            Dim buffer As String = New String(CChar(" "), 25)
            Dim retVal As Integer = GetUserName(buffer, 25)
            Dim userName As String = Strings.Left(buffer, InStr(buffer, Chr(0)) - 1)
            MsgBox(userName)
        End Sub
        Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

            Dim DeviceCount As Integer
            Dim DeviceIndex As Integer
            Dim TempDevString As String
            Dim BytesWritten As Integer
            Dim TempStringData As String
            Dim BytesRead As Integer

            ' Get the number of device attached
            FT_Status = FT_GetNumberOfDevices(DeviceCount, vbNullChar, FT_LIST_NUMBER_ONLY)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If
            ' Display device count on form
            TextBox1.Text = DeviceCount.ToString

            ' Get serial number of device with index 0
            ' Allocate space for string variable
            TempDevString = Space(16)
            FT_Status = FT_GetDeviceString(DeviceIndex, TempDevString, FT_LIST_BY_INDEX Or FT_OPEN_BY_SERIAL_NUMBER)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If
            FT_Serial_Number = Microsoft.VisualBasic.Left(TempDevString, InStr(1, TempDevString, vbNullChar) - 1)
            ' Display serial number on form
            TextBox2.Text = FT_Serial_Number

            ' Get description of device with index 0
            ' Allocate space for string variable
            TempDevString = Space(64)
            FT_Status = FT_GetDeviceString(DeviceIndex, TempDevString, FT_LIST_BY_INDEX Or FT_OPEN_BY_DESCRIPTION)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If
            FT_Description = Microsoft.VisualBasic.Left(TempDevString, InStr(1, TempDevString, vbNullChar) - 1)
            ' Display serial number on form
            TextBox3.Text = FT_Description

            'Open device by serial number
            FT_Status = FT_OpenByDescription(FT_Description, 2, FT_Handle)
            If FT_Status <> FT_OK Then
                MsgBox("Failed to open device.", , )
                Exit Sub
            End If

            ' Reset device
            FT_Status = FT_ResetDevice(FT_Handle)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If

            ' Purge buffers
            FT_Status = FT_Purge(FT_Handle, FT_PURGE_RX Or FT_PURGE_TX)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If

            ' Set Baud Rate
            FT_Status = FT_SetBaudRate(FT_Handle, 115200)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If

            ' Set parameters
            FT_Status = FT_SetDataCharacteristics(FT_Handle, FT_DATA_BITS_8, FT_STOP_BITS_1, FT_PARITY_NONE)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If

            ' Set Flow Control
            FT_Status = FT_SetFlowControl(FT_Handle, FT_FLOW_RTS_CTS, 0, 0)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If

            ' Set RTS
            FT_Status = FT_SetRts(FT_Handle)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If

            ' Set DTR
            FT_Status = FT_SetDtr(FT_Handle)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If

            ' Write string data to device
            FT_Status = FT_Write_String(FT_Handle, TextBox4.Text, Len(TextBox4.Text), BytesWritten)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If

            ' Wait
            Sleep(100)

            ' Get number of bytes waiting to be read
            FT_Status = FT_GetQueueStatus(FT_Handle, FT_RxQ_Bytes)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If

            ' Read number of bytes waiting
            ' Allocate string to recieve data
            TempStringData = Space(FT_RxQ_Bytes + 1)
            FT_Status = FT_Read_String(FT_Handle, TempStringData, FT_RxQ_Bytes, BytesRead)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If
            ' Display string on form
            TextBox5.Text = Trim(TempStringData)

            ' Close device
            FT_Status = FT_Close(FT_Handle)
            If FT_Status <> FT_OK Then
                Exit Sub
            End If
        End Sub

        Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
            Close()
        End Sub

        Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            MsgBox("Connect an FTDI device with D2XX drivers." & Chr(13) & Chr(10) & "Connect a loop-back connector to write and read data through the device.", , )
        End Sub
    End Class

End Namespace
