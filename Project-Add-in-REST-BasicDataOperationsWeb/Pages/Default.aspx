<%-- Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file. --%>
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Project_Add_in_REST_BasicDataOp.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
 <div>
            <center><h2>Projects</h2></center>
       <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePartialRendering="true" />
      <asp:UpdatePanel ID="ListManagementPanel" runat="server" UpdateMode="Always" ChildrenAsTriggers="true" EnableViewState ="true">
        <ContentTemplate>   
            <center>
                <asp:Button runat="server" ID="btnCreateProject" Text="Create Project Name" OnClick="btnCreateProject_Click"  />
                <asp:TextBox ID="txtNewProjectName" runat="server" Width="500px"/>
                <br />
                <asp:HiddenField ID="hidProjectId" runat="server" Value="" />
                <asp:Button runat="server" ID="btnUpdateProject" Text="Update Project Name" OnClick="btnUpdateProject_Click" />
                <asp:TextBox ID="txtUpdateProjectName" runat="server" Width="500" /></center>
            
            </center>
      

            <asp:GridView ID="grdProjects" runat="server"  AutoGenerateColumns="False" OnRowCommand="grdProjects_RowCommand" HorizontalAlign="Center" CellPadding="4" ForeColor="#333333" GridLines="None"

    >

                <AlternatingRowStyle BackColor="White" />

    <Columns>

        <asp:BoundField DataField="id" HeaderText="Id" />
        <asp:BoundField DataField="name" HeaderText="Name" />
        <asp:BoundField DataField="startDate" HeaderText="Start Date"  />
        <asp:BoundField DataField="endDate" HeaderText="End Date"  />
        <asp:CheckBoxField DataField="isCheckedOut" HeaderText="Checked Out" 
                           ReadOnly="True" />
        <%--<asp:BoundField DataField="isCheckedOut" HeaderText="isCheckedOut" DataFormatString="b" />--%>
        <asp:ButtonField buttontype="Button" 
            commandname="Edit"
            headertext="Edit" 
            text="Edit" />
          <asp:ButtonField buttontype="Button" 
            commandname="Delete"
            headertext="Delete" 
            text="Delete" />

    </Columns>

                <EditRowStyle BackColor="#2461BF" />
                <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                <RowStyle BackColor="#EFF3FB" />
                <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                <SortedAscendingCellStyle BackColor="#F5F7FB" />
                <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                <SortedDescendingCellStyle BackColor="#E9EBEF" />
                <SortedDescendingHeaderStyle BackColor="#4870BE" />

</asp:GridView>
            
        </ContentTemplate>
        </asp:UpdatePanel>
    </div>
    </form>
</body>
</html>

<%--
SharePoint Add-in REST/OData Basic Data Operations, https://github.com/OfficeDev/SharePoint-Add-in-REST-OData-BasicDataOperations
 
Copyright (c) Microsoft Corporation
All rights reserved. 
 
MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
 
The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.    
  
--%>