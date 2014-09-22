Partial Class demos_zip_code_distance_Default
    Inherits System.Web.UI.Page

	Sub GetInitialCoordinates() Handles btnZip.Click
		'This subroutine requires a Label control named lblStatus
		
		'Prepare to connect to db and execute stored procedure
		Dim objConn As New SqlConnection(ConfigurationManager.ConnectionStrings("YOUR CONNECTION STRING").ConnectionString)
		Dim objCmd As New SqlCommand("get_by_zip_code", objConn)
		objCmd.CommandType = CommandType.StoredProcedure

		'we need to supply the ZIP code as an input parameter to our stored procedure
		objCmd.Parameters.Add(New SqlParameter("zip_code", SqlDbType.Char, 5))
		objCmd.Parameters("zip_code").Value = tbZip.Text

		'sglMinLat = south, sglMaxLat = north, sglMinLon = west, sglMaxLon = east
		Dim sglMinLat As Single
		Dim sglMaxLat As Single
		Dim sglMinLon As Single
		Dim sglMaxLon As Single

		Try
			'open connection
			objConn.Open()
			'put results into datareader
			Dim objReader As SqlDataReader
			objReader = objCmd.ExecuteReader()
			If objReader.HasRows Then
				'if starting point found, calculate box points
				objReader.Read()
				sglMinLat = Rad2Deg(CalculateLatitudeCoordinate(Deg2Rad(objReader("latitude")), 3959, Deg2Rad(180), ddlDistance.SelectedValue))
				sglMaxLat = Rad2Deg(CalculateLatitudeCoordinate(Deg2Rad(objReader("latitude")), 3959, Deg2Rad(0), ddlDistance.SelectedValue))
				sglMinLon = Rad2Deg(CalculateLongitudeCoordinate(Deg2Rad(objReader("longitude")), Deg2Rad(objReader("latitude")), Deg2Rad(sglMinLat), 3959, Deg2Rad(270), ddlDistance.SelectedValue))
				sglMaxLon = Rad2Deg(CalculateLongitudeCoordinate(Deg2Rad(objReader("longitude")), Deg2Rad(objReader("latitude")), Deg2Rad(sglMinLat), 3959, Deg2Rad(90), ddlDistance.SelectedValue))
				
				'report starting point details to lblStatus
				Dim strOut As String
				strOut = "ZIP Code " & tbZip.Text & " is assigned to " & objReader("city_name") & ", " & objReader("state_code") & ".<br />"
				strOut &= "It is located at latitude " & objReader("latitude") & ", longitude " & objReader("longitude") & ".<br /><br />"
				strOut &= "At a distance of " & ddlDistance.SelectedValue & " miles, the search box coordinates are:<br />"
				strOut &= "Maximum latitude (North): " & sglMaxLat & "<br />"
				strOut &= "Miniumum latitude (South): " & sglMinLat & "<br />"
				strOut &= "Maximum longitude (East): " & sglMaxLon & "<br />"
				strOut &= "Minimum longitude (West): " & sglMinLon & "<br />"
				lblStatus.Text = strOut

				'populate gridview
				PopulateGridView(sglMinLat, sglMaxLat, sglMinLon, sglMaxLon, objReader("latitude"), objReader("longitude"))
			Else
				'starting point not found
				lblStatus.Text = "Error retrieving initial ZIP Code coordinates: No record found for " & tbZip.Text & "."
			End If
			objConn.Close()
			objCmd.Dispose()
			objConn.Dispose()
		Catch ex As Exception
			'technical problem running the query
			lblStatus.Text = "Error executing database query for initial coordinates: " & ex.Message
		End Try
	End Sub

	Sub PopulateGridView(ByVal sglMinLat As Single, ByVal sglMaxLat As Single, ByVal sglMinLon As Single, ByVal sglMaxLon As Single, ByVal sglStartLat As Single, ByVal sglStartLon As Single)
		sqlZip.SelectParameters("minlat").DefaultValue = sglMinLat
		sqlZip.SelectParameters("maxlat").DefaultValue = sglMaxLat
		sqlZip.SelectParameters("minlon").DefaultValue = sglMinLon
		sqlZip.SelectParameters("maxlon").DefaultValue = sglMaxLon
		sqlZip.SelectParameters("startlat").DefaultValue = sglStartLat
		sqlZip.SelectParameters("startlon").DefaultValue = sglStartLon

		gvZIP.DataBind()
	End Sub

	Function CalculateLatitudeCoordinate(ByVal sglLat1 As Single, ByVal intRadius As Integer, ByVal intBearing As Integer, ByVal intDistance As Integer) As Single
		Return Math.Asin(Math.Sin(sglLat1) * Math.Cos(intDistance / intRadius) + Math.Cos(sglLat1) * Math.Sin(intDistance / intRadius) * Math.Cos(intBearing))
	End Function

	Function CalculateLongitudeCoordinate(ByVal sglLon1 As Single, ByVal sglLat1 As Single, ByVal sglLat2 As Single, intRadius As Integer, ByVal intBearing As Integer, ByVal intDistance As Integer) As Single
		Return sglLon1 + Math.Atan2(Math.Sin(intBearing) * Math.Sin(intDistance / intRadius) * Math.Cos(sglLat1), Math.Cos(intDistance / intRadius) - Math.Sin(sglLat1) * Math.Sin(sglLat2))
	End Function

	Function Deg2Rad(ByVal sglDegrees As Single) As Single
		Return sglDegrees * (Math.PI / 180.0)
	End Function

	Function Rad2Deg(ByVal sglRadians As Single) As Single
		Return sglRadians * (180.0 / Math.PI)
	End Function

End Class
