This Collection has REST API calls for Site Admin User Management in ALM instace.
Steps to follow to get started. Import the json collection into Postman Workspace. Select the "Variables tab in this collection" and update the following fields:
- ALMServer - in the format: https://almserver/qcbin
- If authenticating using APIKey and APISecret use the relevant fields and execute "Login_APIKey" to authenticate.
- If authenticating using Username and Password use the relevant fields and execute "Login_Username1" and "Login_Username2" to authenticate.
"Check - is authenticated" is optional in case you wish to verify the authenticated user.
Use "Logout" once you're ready to logout.
"Get customer ID" Gets Customer ID, which is required for many subsiquent calls . hence running this is mandatory.
"Get GetAllSiteUser-LimitPageSize" - this call gets the list of all Site Users.You can define start Index and Page Size in Query Parm in Params tab.
"GetAllSiteUsers" - this call gets all the users in the ALM instance. If there are many users then would recommend using limit page size.
"Post_CreateUser" - Creates SiteUser on ALM instance. Update the body with User details required to create user.
"PutDeactivateUser" -Deactivates existing user in ALM instance. Update Path Param "username" in Params tab.
"PutActivateUser" - Activate exisitng user in ALM instance. Update Path Param "username" in Params tab.
"GetUserProject" - Get the list of projects the User is assigned to.Update Path Param "username" in Params tab.
"GetUserProperty"-Get details about user.Update Path Param "username" in Params tab.



