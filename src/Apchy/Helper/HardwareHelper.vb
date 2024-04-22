Imports System.Management

Namespace Helper
    Public Class HardwareHelper
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

End Namespace