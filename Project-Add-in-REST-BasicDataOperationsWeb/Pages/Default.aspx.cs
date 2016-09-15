// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;


namespace Project_Add_in_REST_BasicDataOp
{
    public partial class Default : System.Web.UI.Page
    {
     
        SharePointContextToken contextToken;
        string accessToken;
        Uri sharepointUrl;
        String digestValue;

        //QueueMsgType enumeration values used from ProjectServer.Client
        const int qMsgType_ProjectPublish = 24;
        const int qMsgType_ProjectUpdate  = 28;
        const int qMsgType_ProjectCheckIn = 21;

        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        // The Page_load method gets the context token and the access token. The access token is used 
        // by all moethods interacting with the service.
        protected void Page_Load(object sender, EventArgs e)
        {
            string contextTokenString = TokenHelper.GetContextTokenFromRequest(Request);

            if (contextTokenString != null)
            {
                contextToken =
                    TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Url.Authority);

                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
                accessToken =
                    TokenHelper.GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken;

                // In a production add-in, you should cache the access token somewhere, such as in a database
                // or ASP.NET Session Cache. (Do not put it in a cookie.) Your code should also check to see 
                // if it is expired before using it (and use the refresh token to get a new one when needed). 
                // For more information, see the MSDN topic at https://msdn.microsoft.com/library/office/dn762763.aspx
                // For simplicity, this sample does not follow these practices. 

                ViewState["accessToken"] = accessToken;

                //Requests to SharePoint that can make changes must have a Form Digest
                //More info:https://msdn.microsoft.com/en-us/library/office/ms472879(v=office.14).aspx
                SetDigestValue(sharepointUrl);
                DisplayProjects();

            }
            else if (!IsPostBack)
            {
             //   Response.Write("Could not find a context token.");
            }
            else 
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            accessToken = (string)ViewState["accessToken"];
        }

        //Will make a request to the Projects endpoint and retrieve project info
        //as JSON
        private void DisplayProjects()
        {
            
            String jsonPayload = GetReSTRequest(sharepointUrl, "ProjectServer/Projects");

            var serializer = new JavaScriptSerializer();
            var jsonObject = serializer.DeserializeObject(jsonPayload) as Dictionary<string, object>;
       
            object[] projectsList = (object[])jsonObject["value"];
            List<Project> projs = new List<Project>();
            Array.ForEach(projectsList, proj => 
            {
                Project p = ConvertJsontoProject(proj as Dictionary<string, object>);
                projs.Add(p);
            }); 

            grdProjects.DataSource = projs;
            grdProjects.DataBind();
        }

        private static Project ConvertJsontoProject(Dictionary<string, object> projDetails)
        {
            string id = (string)projDetails["Id"];
            string name = (string)projDetails["Name"];
            DateTime startDate = DateTime.Parse((string)projDetails["StartDate"]);
            DateTime endDate = DateTime.Parse((string)projDetails["FinishDate"]);
            bool isCheckedOut = (bool)projDetails["IsCheckedOut"];

            Project p = new Project();
            p.id = id;
            p.name = name;
            p.startDate = startDate;
            p.endDate = endDate;
            p.isCheckedOut = isCheckedOut;
            
            return p;
        }

        //Waiting for the project queue helper class
        private void WaitForQueue(string siteUrl, string projectId, int jobtype)
        {
            string projId = projectId;
            string sJobType = jobtype.ToString();
            int jobType = jobtype;

            object[] jobsListX = null;

            var serializer = new JavaScriptSerializer();
           
            while(true)
            {
                String endPoint = "ProjectServer/Projects('" + projId + "')/queuejobs?filter=MessageType eq " + jobType + "'";
                String jsonPayload = GetReSTRequest(siteUrl.ToString(), endPoint);

                var jsonObject = serializer.DeserializeObject(jsonPayload) as Dictionary<string, object>;
                jobsListX = (object[])jsonObject["value"];

                if (jobsListX.Length != 0)
                {
                    System.Threading.Thread.Sleep(500);
                } else
                {
                    break;
                }
            } 
        }
        //Waiting for the project queue helper class
        private void WaitForQueue(Uri siteUrl, string projectId, int jobtype)
        {
            WaitForQueue(siteUrl.ToString(), projectId, jobtype);
        }


        //Generic build REST request method
        //Will build up the request as needed adding form digest if provided
        //All interaction is assumed to be in json format
        private HttpWebRequest BuildReSTRequest(String siteUrl, String endpoint, String method, String body = null)
        {
            String url = siteUrl.TrimEnd('/') + "/_api/" + endpoint;
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(url);

            req.Method = method;
            Boolean isReadOnly = Array.Exists(new[] { "GET", "HEAD" }, c => c.Equals(req.Method));
            Boolean isDigestRequest = endpoint.Contains("contextinfo");
            if (string.IsNullOrEmpty(body))
            {
                req.ContentLength = 0;
            }
            else
            {
                req.ContentLength = body.Length;
                req.ContentType = "application/json";
            }

            //add the digest if the request can make changes to the service
            if (!isDigestRequest && !isReadOnly)
            {
                
                req.Headers.Add("X-RequestDigest", digestValue);
            }

            req.Headers.Add("Authorization", "Bearer " + accessToken);
            req.Accept = "application/json";


            //write out the request body if set
            if (!string.IsNullOrEmpty(body))
            {
                StreamWriter writer = new StreamWriter(req.GetRequestStream());
                writer.Write(body);
                writer.Close();
                writer.Dispose();
            }
            return req;
        }

        private string PostReSTRequest(string siteUrl, string endpoint, string body = null)
        {
            String returnVal = null;

            HttpWebRequest request = BuildReSTRequest(siteUrl, endpoint, "POST", body);
            WebResponse resp = request.GetResponse();
            if (resp != null)
            {
                StreamReader reader = new StreamReader(resp.GetResponseStream());
                returnVal = reader.ReadToEnd();

                reader.Dispose();
            }

            return returnVal;
        }

        private string PostReSTRequest(Uri siteUri, string endpoint, string body = null)
        {
            return PostReSTRequest(siteUri.ToString(), endpoint, body);
        }

        private string PatchReSTRequest(string siteUrl, string endpoint, string body)
        {
            string returnVal = null;

            HttpWebRequest request = BuildReSTRequest(siteUrl, endpoint, "PATCH", body);
            WebResponse resp = request.GetResponse();
            if (resp != null)
            {
                StreamReader reader = new StreamReader(resp.GetResponseStream());
                reader.ReadToEnd();
                reader.Dispose();
            }

            return returnVal;
        }
        //Patch Request helper method
        private string PatchReSTRequest(Uri siteUri, string endpoint, string body = null)
        {
            return PatchReSTRequest(siteUri.ToString(), endpoint, body);
        }

        //Get Request helper method
        private String GetReSTRequest(string siteUrl, string endpoint, string body = null)
        {
            String returnVal = null;

            HttpWebRequest request = BuildReSTRequest(siteUrl, endpoint, "GET", body);
            try
            {
                WebResponse resp = request.GetResponse();
                if (resp != null)
                {
                    StreamReader reader = new StreamReader(resp.GetResponseStream());
                    returnVal = reader.ReadToEnd();
                    reader.Dispose();
                }
            }
            catch { }

            return returnVal;
        }
        //Get Request helper method
        private string GetReSTRequest(Uri siteUri, string endpoint, string body = null)
        {
            return GetReSTRequest(siteUri.ToString(), endpoint, body);
        }

        //Delete Request helper method
        private string DeleteReSTRequest(string siteUrl, string endpoint, string body = null)
        {
            String returnVal = null;

            HttpWebRequest request = BuildReSTRequest(siteUrl, endpoint, "DELETE", body);
            WebResponse resp = request.GetResponse();

            if (resp != null)
            {
                StreamReader reader = new StreamReader(resp.GetResponseStream());
                returnVal = reader.ReadToEnd();
                reader.Dispose();
            }

            return returnVal;
        }

        //Delete Request helper method
        private string DeleteReSTRequest(Uri siteUri, string endpoint, string body = null)
        {
            return DeleteReSTRequest(siteUri.ToString(), endpoint, body);
        }


        //gets the form digest for the Patch/Post/Delete requests
        private void SetDigestValue(Uri siteUrl)
        {
            SetDigestValue(siteUrl.ToString());
        }
        //see https://msdn.microsoft.com/en-us/library/office/ms472879(v=office.14).aspx to understand what form digest is used for
        private void SetDigestValue(String siteUrl)
        {
            var serializer = new JavaScriptSerializer();
            string jsonPayload = PostReSTRequest(siteUrl, "contextinfo", null);
            var jsonObject = serializer.DeserializeObject(jsonPayload) as Dictionary<string, object>;
            if (jsonObject.ContainsKey("FormDigestValue"))
            {
                digestValue = (string)jsonObject["FormDigestValue"];
            }
           
        }
        
        
        //Handles updating the project name
        protected void btnUpdateProject_Click(object sender, EventArgs e)
        {
            //Update the project - in this case, it's the project name.

            //Check out the project - prep and call
            string projectId = hidProjectId.Value.ToString();
            string projEndpoint = "ProjectServer/Projects('" + projectId + "')/checkOut";
            
            string returnVal = PostReSTRequest(sharepointUrl, projEndpoint, null);


            //Update the project - prep and call
            //siteUrl and Id are unchanged.
            string updateEndpoint = "ProjectServer/Projects('" + projectId + "')/Draft";
            string projBody = "{'Name':'" + txtUpdateProjectName.Text + "'}";

            PatchReSTRequest(sharepointUrl, updateEndpoint, projBody);

            //Let command complete before check in 
            WaitForQueue(sharepointUrl, projectId, qMsgType_ProjectUpdate);

            //Publish (check in) project - qMsgType_ProjectCheckIn
            string checkinEndpoint = "ProjectServer/Projects('" + projectId + "')/Draft/publish(true)";
            PostReSTRequest(sharepointUrl, checkinEndpoint, null);

            //Let command complete before updating the display
            WaitForQueue(sharepointUrl, projectId, qMsgType_ProjectPublish);

            DisplayProjects();
        }

        //Handles the new project flow
        //Will post ot the Add Project endpoint with the project name as a parameter
        //Then we wait for the queue to finish processing the publish event
        //and reload the project list
        protected void btnCreateProject_Click(object sender, EventArgs e)
        {
            string projEndpoint = "ProjectServer/Projects/Add";

            //Info for the new project
            string newName;
            string newId = Guid.NewGuid().ToString();
            string newDesc = "Auto generated. An update to this field is appropriate.";

            //Get the name of the new project (or fabricate one)
            if (txtNewProjectName.Text.Length > 0)
            {
                newName = txtNewProjectName.Text;
            }
            else
            {
                newName = "Project Rial created";
            }


            //Build JSON parameter string and make the call.
            string projBody =
                "{'parameters': {'Id':'" + newId + "', 'Name':'" + newName + "', 'Description':'" + newDesc + "'} }";

            string returnVal = PostReSTRequest(sharepointUrl, projEndpoint, projBody);

            //Project additions/changes go through the queue on saving
            WaitForQueue(sharepointUrl, newId, qMsgType_ProjectPublish);

            DisplayProjects();

        }


        //Handles the row buttons
        protected void grdProjects_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            switch (e.CommandName) {
                case "Edit":
                    {
                        // Convert the row index stored in the CommandArgument
                        // property to an Integer.
                        int index = Convert.ToInt32(e.CommandArgument);
                        GridViewRow row = grdProjects.Rows[index];
                        hidProjectId.Value = row.Cells[0].Text;
                        txtUpdateProjectName.Text = row.Cells[1].Text;
                        e.Handled = true;
                    }
                    break;
                case "Delete":
                    {
                        int index = Convert.ToInt32(e.CommandArgument);
                        GridViewRow row = grdProjects.Rows[index];
                        String projectId = row.Cells[0].Text;

                        // Build and send delete request to server
                        string projEndpoint = "ProjectServer/Projects('" + projectId + "')/";
                        string returnVal = DeleteReSTRequest(sharepointUrl, projEndpoint);

                        // Add call to get project - when the delete completes, a "not found" error occurs. 
                        // Can move on from there.

                        while (GetReSTRequest(sharepointUrl, projEndpoint) != null)
                        {
                            System.Threading.Thread.Sleep(500);
                        }

                        DisplayProjects();
                        e.Handled = true;
                    }
                    break;
            }
        }
    }
}

/*
Project Add-in REST/OData Basic Data Operations, https://github.com/OfficeDev/Project-Add-in-REST-OData-BasicDataOperations
 
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
  
*/
