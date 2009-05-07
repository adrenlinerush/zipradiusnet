Option Explicit On
Imports Microsoft.Win32
Imports System.Math
Public Class clsRadius
    Private ZipsInRadius As ArrayList
    Private dsZipsInBox As DataTable
    Private objConnection As New SqlClient.SqlConnection
    Private objDA As New SqlClient.SqlDataAdapter
    Private Const dblEarthRadiusMiles As Double = 3956.0
    Private dblLat As Double
    Private dblLng As Double
    Public Function GetZipsInRadius(ByVal strZip As String, ByVal dblDistance As Double) As Object
        Dim strRet As String
        strRet = OpenDB()
        If strRet = "" Then
            strRet = GetZipsInBox(strZip, dblDistance)
            If strRet = "" Then
                strRet = RemoveOutRadiusZips(strZip, dblDistance)
                If strRet = "" Then
                    Call CloseDB()
                    Return ZipsInRadius
                Else
                    Call CloseDB()
                    Return strRet
                End If
            Else
                Call CloseDB()
                Return strRet
            End If
        Else
            Call CloseDB()
            Return strRet
        End If
    End Function

    Private Function GetZipsInBox(ByVal strZip As String, ByVal dblDistance As Double) As String
        Try
            Dim strSQL As String = "Select * from ZipCoordinates Where Zip = " & strZip
            Dim dsStartZip As New DataTable
            objDA = New SqlClient.SqlDataAdapter(strSQL, objConnection)
            objDA.Fill(dsStartZip)

            If dsStartZip.Rows.Count > 0 Then
                dblLat = dsStartZip.Rows(0).Item("Latitude").ToString
                dblLng = dsStartZip.Rows(0).Item("Longitude").ToString
                Dim dblLatRad As Double = dblLat * (PI / 180)
                Dim dblLngRad As Double = dblLng * (PI / 180)
                Dim dblRRad As Double = dblDistance / dblEarthRadiusMiles

                Dim dblSLat As Double = (dblLatRad - dblRRad) * (180 / PI)
                Dim dblNLat As Double = (dblLatRad + dblRRad) * (180 / PI)

                Dim lat As Double = Asin(Sin(dblLatRad) * Cos(dblRRad))
                Dim lon As Double = Atan2(Sin(PI / 2) * Sin(dblRRad) * Cos(dblLatRad), Cos(dblRRad) - Sin(dblLatRad) * Sin(lat))
                Dim dblELng As Double = (((dblLngRad + lon + PI) Mod (2.0 * PI)) - PI) * (180 / PI)

                lon = Atan2(Sin(3.0 * PI / 2.0) * Sin(dblRRad) * Cos(dblLatRad), Cos(dblRRad) - Sin(dblLatRad) * Sin(lat))
                Dim dblWLng As Double = (((dblLngRad + lon + PI) Mod (2.0 * PI)) - PI) * (180 / PI)

                dsZipsInBox = New DataTable

                strSQL = "Select * from ZipCoordinates Where Longitude < " & dblELng & _
                    " and Longitude > " & dblWLng & " and Latitude > " & dblSLat & _
                    " and Latitude < " & dblNLat

                objDA = New SqlClient.SqlDataAdapter(strSQL, objConnection)
                objDA.Fill(dsZipsInBox)
            Else
                Return "Zip Not Found"
            End If
            Return ""
        Catch
            Return "GetZipsInBox: " & Err.Description
        End Try
    End Function
    Private Function RemoveOutRadiusZips(ByVal strZip As String, ByVal dblDistance As Double) As String
        Try
            ZipsInRadius = New ArrayList
            Dim iZip As Integer
            Dim curDistance As Double
            Dim curLat As Double
            Dim curLng As Double

            For iZip = 0 To dsZipsInBox.Rows.Count - 1
                curLat = dsZipsInBox.Rows(iZip).Item("Latitude").ToString
                curLng = dsZipsInBox.Rows(iZip).Item("Longitude").ToString

                curDistance = Sin((curLat - dblLat) / 2) ^ 2 + Cos(dblLat) * Cos(curLat) * Sin((curLng - dblLng) / 2) ^ 2

                If curDistance <= dblDistance Then
                    ZipsInRadius.Add(dsZipsInBox.Rows(iZip).Item("Zip").ToString)
                End If
            Next

            Return ""
        Catch
            Return "RemoveOutRadiusZips: " & Err.Description
        End Try
    End Function
    Private Function OpenDB() As String
        Dim regShipsCommon As RegistryKey
        Dim systemSettings As New DataSet
        Dim connectionString As String

        Try
            regShipsCommon = Registry.LocalMachine.OpenSubKey("Software\Ships\Common")
            connectionString = regShipsCommon.GetValue("connectionStringNET")

            objConnection.ConnectionString = connectionString
            objConnection.Open()
            OpenDB = True
            Return ""
        Catch
            Return Err.Description
        End Try

    End Function
    Private Sub CloseDB()
        Try
            objConnection.Close()
        Catch
        End Try
    End Sub
End Class
