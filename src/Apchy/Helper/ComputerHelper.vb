Imports System.Drawing
Imports System.Management
Imports System.Net.NetworkInformation

Namespace Helper
    Public Class ComputerHelper
        Public Shared Function GetCPUID() As String
            Dim Mc As ManagementClass = New ManagementClass("Win32_Processor")
            Dim Moc As ManagementObjectCollection
            Moc = Mc.GetInstances()
            Dim strCpuID As String = ""
            For Each mo As ManagementObject In Moc
                strCpuID = mo.Properties("ProcessorId").Value.ToString()
                Exit For
            Next

            Return strCpuID
        End Function

        Public Shared Function GetMainHardDiskId() As String  '获取硬盘序列号
            Dim cmicWmi As New System.Management.ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")

            Dim Uint32 As UInt32

            For Each cmicWmiObj As ManagementObject In cmicWmi.Get
                Uint32 = cmicWmiObj("signature")
            Next

            Return Uint32.ToString
        End Function

        Public Shared Function Name() As String
            Return My.Computer.Name
        End Function

        Public Shared Function GetComputerInfo() As ComputerInfo
            Dim myInfo As New ComputerInfo

            With myInfo
                .Name = My.Computer.Name
                .OSFullName = My.Computer.Info.OSFullName
                .OSPlatform = My.Computer.Info.OSPlatform
                .OSVersion = My.Computer.Info.OSVersion

                .TotalPhysicalMemory = My.Computer.Info.TotalPhysicalMemory
                .AvailablePhysicalMemory = My.Computer.Info.AvailablePhysicalMemory

                .Resolution = My.Computer.Screen.Bounds
            End With

            Return myInfo
        End Function

        Public Shared Function GetResolution() As Rectangle
            Return My.Computer.Screen.Bounds
        End Function

        Public Shared Function GetNetwork() As List(Of NetWorkInfo)
            Dim myInfo As New List(Of NetWorkInfo)

            Dim myNetworkInfo As NetWorkInfo

            For Each myNetworkInterface As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces
                Dim myAddress As PhysicalAddress = myNetworkInterface.GetPhysicalAddress()

                myNetworkInfo = New NetWorkInfo
                With myNetworkInfo
                    .Name = myNetworkInterface.Name
                    .MacAddess = myNetworkInterface.GetPhysicalAddress.ToString
                    .IPAddress = myNetworkInterface.GetIPProperties.ToString
                End With

                myInfo.Add(myNetworkInfo)
            Next

            Return myInfo
        End Function

        Public Shared Function GetDiskModels() As List(Of String)
            Dim mosDisks As ManagementObjectSearcher = New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")

            Dim myString As New List(Of String)

            For Each moDisk In mosDisks.Get
                myString.Add(moDisk("Model"))
            Next

            Return myString
        End Function

        Public Shared Function GetDiskInfo(ByVal ModelName As String) As DiskInfo
            Dim mosDisks As ManagementObjectSearcher = New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive WHERE Model = '" & ModelName & "'")
            Dim myDiskInfo As New DiskInfo

            On Error Resume Next

            For Each moDisk In mosDisks.Get
                With myDiskInfo
                    .MediaType = moDisk("MediaType")
                    .Model = moDisk("Model")
                    .SerialNumber = moDisk("SerialNumber")
                    .InterfaceType = moDisk("InterfaceType")
                    .Capacity = moDisk("Size")
                    .Partitions = moDisk("Partitions")
                    .Signature = moDisk("Signature")
                    .FirmwareRevision = moDisk("FirmwareRevision")
                    .TotalCylinders = moDisk("TotalCylinders")
                    .TotalHeads = moDisk("TotalHeads")
                    .TotalSectors = moDisk("TotalSectors")
                    .TotalTracks = moDisk("TotalTracks")
                    .BytesPerSector = moDisk("BytesPerSector")
                    .SectorsPerTrack = moDisk("SectorsPerTrack")
                    .TracksPerCylinder = moDisk("TracksPerCylinder")
                    .PNPDeviceID = moDisk("PNPDeviceID")
                End With
            Next

            Return myDiskInfo
        End Function
    End Class

    Public Class DiskInfo
        Public MediaType As String
        Public Model As String
        Public SerialNumber As String
        Public InterfaceType As String
        Public Capacity As String
        Public Partitions As Integer
        Public Signature As String

        Public PNPDeviceID As String
        Public FirmwareRevision As String
        Public TotalCylinders As Integer
        Public TotalSectors As Integer
        Public TotalHeads As Integer
        Public TotalTracks As Integer
        Public BytesPerSector As Integer
        Public SectorsPerTrack As Integer
        Public TracksPerCylinder As Integer
    End Class

    Public Class NetWorkInfo
        Public Name As String
        Public MacAddess As String
        Public IPAddress As String
    End Class

    Public Class ComputerInfo
        Public Name As String
        Public OSPlatform As String
        Public OSFullName As String
        Public OSVersion As String

        Public TotalPhysicalMemory As ULong
        Public AvailablePhysicalMemory As ULong

        Public Resolution As Rectangle
    End Class

End Namespace