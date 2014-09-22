<%@ Page Title="Getting All ZIP Codes Within A Known Distance To A Given Point / ZIP Code Via ASP.NET / SQL Server" Language="VB" MasterPageFile="~/Default.master" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="demos_zip_code_distance_Default" %>

<asp:Content ID="cntSidebar" ContentPlaceHolderID="cphSidebar" Runat="Server">
</asp:Content>
<asp:Content ID="cntMain" ContentPlaceHolderID="cphMain" Runat="Server">
	<!--
	
        Getting All ZIP Codes Within A Known Distance To A Given Point / ZIP Code Via ASP.NET / SQL Server
        Copyright (C) 2012 Doug Vanderweide

        This program is free software: you can redistribute it and/or modify
        it under the terms of the GNU General Public License as published by
        the Free Software Foundation, either version 3 of the License, or
        (at your option) any later version.

        This program is distributed in the hope that it will be useful,
        but WITHOUT ANY WARRANTY; without even the implied warranty of
        MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
        GNU General Public License for more details.

        You should have received a copy of the GNU General Public License
        along with this program.  If not, see <http://www.gnu.org/licenses/>.

    -->

    <h2>Getting All ZIP Codes Within A Known Distance To A Given Point / ZIP Code Via ASP.NET / SQL Server</h2>
	
	<p class="note">Accepts from a user a starting ZIP Code and a distance; returns in a GridView all ZIP Codes that fall within a square of coordinates that distance North, South, East and West of the starting point.</p>

	<p>
		Select all ZIP Codes within 

		<asp:DropDownList runat="server" ID="ddlDistance">
			<asp:ListItem Selected="True">5</asp:ListItem>
			<asp:ListItem>10</asp:ListItem>
			<asp:ListItem>25</asp:ListItem>
			<asp:ListItem>50</asp:ListItem>
			<asp:ListItem>100</asp:ListItem>
		</asp:DropDownList>

		miles of ZIP Code

		<asp:TextBox runat="server" ID="tbZip" Columns="5" />
		<asp:RequiredFieldValidator 
			runat="server" 
			ID="rfvZip" 
			ControlToValidate="tbZip" 
			ErrorMessage="Please provide a ZIP Code" 
			CssClass="warning" 
			Display="Dynamic" 
		/>
		<asp:RegularExpressionValidator 
			runat="server" 
			ID="revZip" 
			ControlToValidate="tbZip" 
			ValidationExpression="^[0-9]{5}$" 
			ErrorMessage="Please enter a valid five-digit ZIP Code" 
			CssClass="warning" 
			Display="Dynamic" 
		/>

		<asp:Button runat="server" ID="btnZip" Text="Get ZIP Codes" />
	</p>

	<p><asp:Label runat="server" ID="lblStatus" Text="Status messages will appear here" /></p>

	<asp:GridView 
		runat="server" 
		ID="gvZIP"
		DataSourceID = "sqlZip" 
		AutoGenerateColumns="false" 
		AllowSorting="true"
		AllowPaging = "true"
		PageSize = "20"
		HeaderStyle-BackColor="Yellow" 
		HeaderStyle-Font-Bold="true" 
		HeaderStyle-HorizontalAlign="Center"
		AlternatingRowStyle-BackColor="WhiteSmoke" 
		CellPadding="5"
	>
		<Columns>
			<asp:BoundField HeaderText="City" DataField="city_name" SortExpression="city_name" />
			<asp:BoundField HeaderText="State" DataField="state_code" SortExpression="state_code" />
			<asp:BoundField HeaderText="ZIP Code" DataField="zip_code" SortExpression="zip_code" />
			<asp:BoundField HeaderText="Latitude" DataField="latitude" SortExpression="latitude" />
			<asp:BoundField HeaderText="Longitude" DataField="longitude" SortExpression="longitude" />
			<asp:BoundField HeaderText="Distance" DataField="distance" SortExpression="distance" />
		</Columns>
	</asp:GridView>

	<asp:SqlDataSource 
		runat="server" 
		ID="sqlZip" 
		SelectCommand="sp_us_zip_codes_get_by_radius" 
		SelectCommandType="StoredProcedure" 
		ConnectionString="<%$ ConnectionStrings:connMain %>"
	>
		<SelectParameters>
			<asp:Parameter Name="maxlat" DbType="Decimal" DefaultValue="0.0" />
			<asp:Parameter Name="minlat" DbType="Decimal" DefaultValue="0.0" />
			<asp:Parameter Name="maxlon" DbType="Decimal" DefaultValue="0.0" />
			<asp:Parameter Name="minlon" DbType="Decimal" DefaultValue="0.0" />
			<asp:Parameter Name="startlat" DbType="Decimal" DefaultValue="0.0" />
			<asp:Parameter Name="startlon" DbType="Decimal" DefaultValue="0.0" />
			<asp:Parameter Name="radius" DbType="Int16" DefaultValue="3959" />
		</SelectParameters>
	</asp:SqlDataSource>

    <h4>Page Code</h4>

    <pre class="brush: xml">
        &lt;p&gt;
		    Select all ZIP Codes within 

		    &lt;asp:DropDownList runat="server" ID="ddlDistance"&gt;
			    &lt;asp:ListItem Selected="True"&gt;5&lt;/asp:ListItem&gt;
			    &lt;asp:ListItem&gt;10&lt;/asp:ListItem&gt;
			    &lt;asp:ListItem&gt;25&lt;/asp:ListItem&gt;
			    &lt;asp:ListItem&gt;50&lt;/asp:ListItem&gt;
			    &lt;asp:ListItem&gt;100&lt;/asp:ListItem&gt;
		    &lt;/asp:DropDownList&gt;

		    miles of ZIP Code

		    &lt;asp:TextBox runat="server" ID="tbZip" Columns="5" /&gt;
		    &lt;asp:RequiredFieldValidator 
			    runat="server" 
			    ID="rfvZip" 
			    ControlToValidate="tbZip" 
			    ErrorMessage="Please provide a ZIP Code" 
			    CssClass="warning" 
			    Display="Dynamic" 
		    /&gt;
		    &lt;asp:RegularExpressionValidator 
			    runat="server" 
			    ID="revZip" 
			    ControlToValidate="tbZip" 
			    ValidationExpression="^[0-9]{5}$" 
			    ErrorMessage="Please enter a valid five-digit ZIP Code" 
			    CssClass="warning" 
			    Display="Dynamic" 
		    /&gt;

		    &lt;asp:Button runat="server" ID="btnZip" Text="Get ZIP Codes" /&gt;
	    &lt;/p&gt;

	    &lt;p&gt;&lt;asp:Label runat="server" ID="lblStatus" Text="Status messages will appear here" /&gt;&lt;/p&gt;

	    &lt;asp:GridView 
		    runat="server" 
		    ID="gvZIP"
		    DataSourceID = "sqlZip" 
		    AutoGenerateColumns="false" 
		    AllowSorting="true"
		    AllowPaging = "true"
		    PageSize = "20"
		    HeaderStyle-BackColor="Yellow" 
		    HeaderStyle-Font-Bold="true" 
		    HeaderStyle-HorizontalAlign="Center"
		    AlternatingRowStyle-BackColor="WhiteSmoke" 
		    CellPadding="5"
	    &gt;
		    &lt;Columns&gt;
			    &lt;asp:BoundField HeaderText="City" DataField="cityname" SortExpression="cityname" /&gt;
			    &lt;asp:BoundField HeaderText="State" DataField="statecode" SortExpression="statecode" /&gt;
			    &lt;asp:BoundField HeaderText="ZIP Code" DataField="zipcode" SortExpression="zipcode" /&gt;
			    &lt;asp:BoundField HeaderText="Latitude" DataField="latitude" SortExpression="latitude" /&gt;
			    &lt;asp:BoundField HeaderText="Longitude" DataField="longitude" SortExpression="longitude" /&gt;
			    &lt;asp:BoundField HeaderText="Distance" DataField="distance" SortExpression="distance" /&gt;
		    &lt;/Columns&gt;
	    &lt;/asp:GridView&gt;

	    &lt;asp:SqlDataSource 
		    runat="server" 
		    ID="sqlZip" 
		    SelectCommand="sp_get_zips_in_radius" 
		    SelectCommandType="StoredProcedure" 
		    ConnectionString="YOUR CONNECTION STRING"
	    &gt;
		    &lt;SelectParameters&gt;
			    &lt;asp:Parameter Name="maxlat" DbType="Decimal" DefaultValue="0.0" /&gt;
			    &lt;asp:Parameter Name="minlat" DbType="Decimal" DefaultValue="0.0" /&gt;
			    &lt;asp:Parameter Name="maxlon" DbType="Decimal" DefaultValue="0.0" /&gt;
			    &lt;asp:Parameter Name="minlon" DbType="Decimal" DefaultValue="0.0" /&gt;
			    &lt;asp:Parameter Name="startlat" DbType="Decimal" DefaultValue="0.0" /&gt;
			    &lt;asp:Parameter Name="startlon" DbType="Decimal" DefaultValue="0.0" /&gt;
			    &lt;asp:Parameter Name="radius" DbType="Int16" DefaultValue="3959" /&gt;
		    &lt;/SelectParameters&gt;
	    &lt;/asp:SqlDataSource&gt;
    </pre>

    <h4>Code Behind</h4>

    <pre class="brush: vb">
    	Sub GetInitialCoordinates() Handles btnZip.Click
	        'This subroutine requires a Label control named lblStatus
	
	        'Prepare to connect to db and execute stored procedure
	        Dim objConn As New SqlConnection("YOUR CONNECTION STRING")
	        Dim objCmd As New SqlCommand("sp_get_zip_code", objConn)
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
			        strOut = "ZIP Code " &amp; tbZip.Text &amp; " is assigned to " &amp; objReader("cityname") &amp; ", " &amp; objReader("statecode") &amp; ".&lt;br /&gt;"
			        strOut &amp;= "It is located at latitude " &amp; objReader("latitude") &amp; ", longitude " &amp; objReader("longitude") &amp; ".&lt;br /&gt;&lt;br /&gt;"
			        strOut &amp;= "At a distance of " &amp; ddlDistance.SelectedValue &amp; " miles, the search box coordinates are:&lt;br /&gt;"
			        strOut &amp;= "Maximum latitude (North): " &amp; sglMaxLat &amp; "&lt;br /&gt;"
			        strOut &amp;= "Miniumum latitude (South): " &amp; sglMinLat &amp; "&lt;br /&gt;"
			        strOut &amp;= "Maximum longitude (East): " &amp; sglMaxLon &amp; "&lt;br /&gt;"
			        strOut &amp;= "Minimum longitude (West): " &amp; sglMinLon &amp; "&lt;br /&gt;"
			        lblStatus.Text = strOut

			        'populate gridview
			        PopulateGridView(sglMinLat, sglMaxLat, sglMinLon, sglMaxLon, objReader("latitude"), objReader("longitude"))
		        Else
			        'starting point not found
			        lblStatus.Text = "Error retrieving initial ZIP Code coordinates: No record found for " &amp; tbZip.Text &amp; "."
		        End If
		        objConn.Close()
		        objCmd.Dispose()
		        objConn.Dispose()
	        Catch ex As Exception
		        'technical problem running the query
		        lblStatus.Text = "Error executing database query for initial coordinates: " &amp; ex.Message
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
    </pre>

</asp:Content>
