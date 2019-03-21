#r "Newtonsoft.Json"
using System;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

public static string ReplaceLastOccurrence(string Source, string Find, string Replace)
{
    int Place = Source.LastIndexOf(Find);
    string result = Source.Remove(Place, Find.Length).Insert(Place, Replace);
    return result;
}

/*
public static int GetIndentLevel(string Source)
{
    int lastIndent = Source.LastIndexOf("\n ");
    string indents = Source.Substring(lastIndent);
    string justIndents = indents.Substring(0, indents.IndexOf("\""));
    int count = 0;
    foreach (char c in justIndents) {
        if (c == ' ') count++;
    }
    return count;
}
*/

public static string Base64Encode(string plainText) {
    var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
    return System.Convert.ToBase64String(plainTextBytes);
}

public static string AddIfNotContained(string listString, string newEntry) {
    if (!listString.Contains(newEntry)) {
        listString += "<br />" + newEntry;
    }
    return listString;
}
 
public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    string PAT = Environment.GetEnvironmentVariable("PAT", EnvironmentVariableTarget.Process);

    string ORG = "domoreexp";
    string TEAM = "Teamspace";
    string PROJECT = "Teamspace-Web";
    string CONFIG_PATH = "/ecs-configs/config-ecs-audience-prod.json";

    string VSTS_GIT_API_BASE = $"https://dev.azure.com/{ORG}/{TEAM}/_apis/git/repositories/{PROJECT}/";
    string API_VERSION = "api-version=4.1";
    string GET_DEFAULT_BRANCH_ENDPOINT = $"{VSTS_GIT_API_BASE}refs/heads/develop?{API_VERSION}";
    string GET_CONFIG_ENDPOINT =         $"{VSTS_GIT_API_BASE}items?path={CONFIG_PATH}&{API_VERSION}";
    string CREATE_BRANCH_ENDPOINT =      $"{VSTS_GIT_API_BASE}refs?{API_VERSION}";
    string GET_COMMIT_BASE =             $"{VSTS_GIT_API_BASE}";
    string PUSH_CHANGE_ENDPOINT =        $"{VSTS_GIT_API_BASE}pushes?{API_VERSION}";
    string CREATE_TASK_ENDPOINT =        $"https://dev.azure.com/{ORG}/MSTeams/_apis/wit/workitems/$task?{API_VERSION}";
    string CREATE_PR_ENDPOINT =          $"{VSTS_GIT_API_BASE}pullrequests?{API_VERSION}";
    
    string AUTH_STRING = "Basic " + Base64Encode(":" + PAT);
    string JUSTIFICATION_STRING = "Use the latest template that most closely matches your PR from http://aka.ms/CentralTriage<br /><br />1. This is an ECS Tenant only change<br /><br />2. I understand a build isn’t necessary for this change and that CT may or may not complete the PR to speed up deployment time or reduce the number of builds.<br /><br />3. I understand it will take up to 15 minutes after the PR shows as complete until it is effective in the environment it was targeting.";
    string MASTER_FEATURE_URL = "https://domoreexp.visualstudio.com/MSTeams/_workitems/edit/157597";

    string REVIEWER_ID = "b8815472-82db-422c-9fae-5a6466b85624";


    log.Info("Webhook was triggered!");
    HttpContent requestContent = req.Content;
    string jsonContent = requestContent.ReadAsStringAsync().Result;

    log.Info(jsonContent);

    JObject jo = JObject.Parse(jsonContent);
    log.Info(jo.ToString());
    log.Info(jo["value"].ToString());
    JArray requests = (JArray)jo["value"];

    DateTime aFortnightPast = DateTime.Now.Date.AddDays(-14);
    log.Info(aFortnightPast.ToString());

    // Get the current config string
    string config = "";
    using (var client = new HttpClient())
    {
        client.DefaultRequestHeaders.Add("Authorization", AUTH_STRING);

        log.Info("Getting config response");

        var response = await client.GetAsync(GET_CONFIG_ENDPOINT);
        config = response.Content.ReadAsStringAsync().Result;
    }
    log.Info("Got config response");

    Dictionary<string, List<String>> notifyDict = new Dictionary<string, List<String>>();
    // Keys are submitter emails
    // Value is initially a table and thead with header values
    // When adding a new entry to their value, add a new trow and td's for each field we're interested in
    // At the end, iterate through the dict and put each table into the workitem.

    int ring1_5Start = config.IndexOf("\"ring1_5\":");
    int ring1_5UserIdsStart = config.Substring(ring1_5Start).IndexOf("\"userIds\":") + ring1_5Start;
    int ring1_5UserIdsEnd = config.Substring(ring1_5UserIdsStart).IndexOf("]") + ring1_5UserIdsStart;
    string originalRing1_5UserIds = config.Substring(ring1_5UserIdsStart, ring1_5UserIdsEnd - ring1_5UserIdsStart);
    string ring1_5UserIds = originalRing1_5UserIds;
    // Add a terminnal comma, since we're adding more
    ring1_5UserIds = ring1_5UserIds.Replace("\" //", "\", //");
    ring1_5UserIds = ring1_5UserIds.Replace("\"  //", "\",  //");
    ring1_5UserIds = ReplaceLastOccurrence(ring1_5UserIds, "\n", "");

    int ring1_5TenantIdsStart = config.Substring(ring1_5Start).IndexOf("\"tenantIds\":") + ring1_5Start;
    int ring1_5TenantIdsEnd = config.Substring(ring1_5TenantIdsStart).IndexOf("]") + ring1_5TenantIdsStart;
    string originalRing1_5TenantIds = config.Substring(ring1_5TenantIdsStart, ring1_5TenantIdsEnd - ring1_5TenantIdsStart);
    string ring1_5TenantIds = originalRing1_5TenantIds;
    ring1_5TenantIds = ring1_5TenantIds.Replace("\" //", "\", //");
    ring1_5TenantIds = ring1_5TenantIds.Replace("\"  //", "\",  //");  // Additional check for when there's two spaces
    ring1_5TenantIds = ReplaceLastOccurrence(ring1_5TenantIds, "\n", "");

    int ring3Start = config.IndexOf("\"ring3\":");
    int ring3TenantIdsStart = config.Substring(ring3Start).IndexOf("\"tenantIds\":") + ring3Start;
    int ring3TenantIdsEnd = config.Substring(ring3TenantIdsStart).IndexOf("]") + ring3TenantIdsStart;
    string originalRing3TenantIds = config.Substring(ring3TenantIdsStart, ring3TenantIdsEnd - ring3TenantIdsStart);
    string ring3TenantIds = originalRing3TenantIds;
    ring3TenantIds = ring3TenantIds.Replace("\" //", "\", //");
    ring3TenantIds = ring3TenantIds.Replace("\"  //", "\",  //");  // Additional check for when there's two spaces
    ring3TenantIds = ReplaceLastOccurrence(ring3TenantIds, "\n", "");

    int ring3UserIdsStart = config.Substring(ring3Start).IndexOf("\"userIds\":") + ring3Start;
    int ring3UserIdsEnd = config.Substring(ring3UserIdsStart).IndexOf("]") + ring3UserIdsStart;
    string originalRing3UserIds = config.Substring(ring3UserIdsStart, ring3UserIdsEnd - ring3UserIdsStart);
    string ring3UserIds = originalRing3UserIds;
    ring3UserIds = ring3UserIds.Replace("\" //", "\", //");
    ring3UserIds = ring3UserIds.Replace("\"  //", "\",  //");
    ring3UserIds = ReplaceLastOccurrence(ring3UserIds, "\n", "");

    log.Info("Got them");

    //log.Info(originalRing1_5UserIds);
    //log.Info(originalRing3TenantIds);
    log.Info(originalRing3UserIds);

    int ring1_5UserIndentLevel = 14;  // Hard-code this for now, since it's only one ring
    int ring1_5TenantIndentLevel = 14;
    int ring3TenantIndentLevel = 14;
    int ring3UserIndentLevel = 8;

    log.Info("Got indent level");

    string ring1_5UserAdditions = "";
    string ring1_5UserRemovals = "";

    string ring1_5TenantAdditions = "";
    string ring1_5TenantRemovals = "";

    string ring3UserAdditions = "";
    string ring3UserRemovals = "";

    string ring3TenantAdditions = "";
    string ring3TenantRemovals = "";

    string problems = "";

    string successEmails = "";
    string rejectEmails = "";

    string NOTIFICATION_TABLE_HEADER = "<table style='width=100%;border=1px solid black;'><thead><tr><td>Type</td><td>Email/Domain</td><td>ObjectID</td><td>Ring</td><td>Add/Remove</td><td>Status</td><td>Explanation</td></thead><tbody>";
    
    foreach (JObject request in requests)
    {
        log.Info(request.ToString());
        string userOrTenant = request["UserOrTenant"].ToString();
        string ring = request["Ring"].ToString();
        string addRemove = request["AddRemove"].ToString();

        string name = request["Name"].ToString();
        name = name.Replace("\n", "");  // Definitely don't want newlines in this
        string email = request["Email"].ToString();
        string id = request["ObjectID"].ToString();

        string submitterEmail = request["SubmitterEmail"].ToString();

        if (!notifyDict.ContainsKey(submitterEmail)) {
            notifyDict[submitterEmail] = new List<string>();
        }

        string reject = request["Reject"].ToString();
        if (reject.Contains("Yes")) {
            problems += "<br />" + id + " (" + email + ") was rejected manually";
            rejectEmails = AddIfNotContained(rejectEmails, submitterEmail);
            notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Failure</td><td>Rejected manually by admin</td></tr>");
            continue;
        }

        string domain = "";
        try {
            domain = email.Substring(email.LastIndexOf('@') + 1);
        } catch (Exception e) {
            domain = request["Email"].ToString();
        }

        string category = request["UserCategory"].ToString();

        DateTime datetime;
        try {
            datetime = (DateTime)request["Datetime"];
        } catch (Exception e)
        {
            double oadate = (double)request["Datetime"];
            datetime = DateTime.FromOADate(oadate);
        }

        if (userOrTenant == "User") {
            if (addRemove == "Add") {
                if (ring == "1.5") {
                    if (ring1_5UserIds.Contains(id)) {
                        log.Info("A user with id " + id + " was already provisioned. Skipping this one");
                        problems += "<br />" + id + " already present in Ring 1.5";
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Failure</td><td>User already provisioned</td></tr>");
                    } else {                    
                        string newLine = "\n" + new String(' ', ring1_5UserIndentLevel) + "\"" + id + "\", // " + domain + " | " + category;
                        ring1_5UserIds += newLine;
                        ring1_5UserAdditions += "<br />" + newLine;
                        successEmails = AddIfNotContained(successEmails, submitterEmail);
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Success</td><td></td></tr>");

                    }
                } else if (ring == "3") {
                    if (ring3UserIds.Contains(id)) {
                        log.Info("A user with id " + id + " was already provisioned. Skipping this one");
                        problems += "<br />" + id + " already present in Ring 3";
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Failure</td><td>User already provisioned</td></tr>");
                    } else {                    
                        string newLine = "\n" + new String(' ', ring3UserIndentLevel) + "\"" + id + "\", // " + domain + " | " + category;
                        ring3UserIds += newLine;
                        ring3UserAdditions += "<br />" + newLine;
                        successEmails = AddIfNotContained(successEmails, submitterEmail);
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Success</td><td></td></tr>");
                    }
                }

            } else if (addRemove == "Remove") {
                if (ring == "1.5") {
                    log.Info("Looking for: " + id);
                    int ind = -1;

                    if (!ring1_5UserIds.Contains(id)) {
                        log.Info("Userid " + id + " not found, skipping its removal");
                        problems += "<br />" + id + " not found in Ring 1.5";
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Failure</td><td>User not found</td></tr>");
                    }
                    // New: Remove all instances of that ID.
                    while (ring1_5UserIds.Contains(id)) {
                        ind = ring1_5UserIds.IndexOf(id);
                        while (ring1_5UserIds[ind] != '\n') {
                            ind -= 1;
                        }
                        // Start including the previous newline, and end including the next newline
                        string subst = ring1_5UserIds.Substring(ind + 1);
                        string thisLine = "\n" + subst.Substring(0, subst.IndexOf("\n"));
                        ring1_5UserIds = ring1_5UserIds.Replace(thisLine, "");
                        ring1_5UserRemovals += "<br /> " + thisLine;
                        successEmails = AddIfNotContained(successEmails, submitterEmail);
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Success</td><td></td></tr>");
                        log.Info("Removing: " + thisLine);
                    }
                }
                else if (ring == "3") {
                    log.Info("Looking for: " + id);
                    int ind = -1;
                    if (!ring3UserIds.Contains(id)) {
                        log.Info("Userid " + id + " not found, skipping its removal");
                        problems += "<br />" + id + " not found in Ring 3";
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Failure</td><td>User not found</td></tr>");
                    }
                    while (ring3UserIds.Contains(id)) {
                        ind = ring3UserIds.IndexOf(id);
                        while (ring3UserIds[ind] != '\n') {
                            ind -= 1;
                        }
                        // Start including the previous newline, and end including the next newline
                        string subst = ring3UserIds.Substring(ind + 1);
                        string thisLine = "\n" + subst.Substring(0, subst.IndexOf("\n"));
                        ring3UserIds = ring3UserIds.Replace(thisLine, "");
                        ring3UserRemovals += "<br /> " + thisLine;
                        successEmails = AddIfNotContained(successEmails, submitterEmail);
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Success</td><td></td></tr>");
                        log.Info("Removing: " + thisLine);
                    }
                }

            }
        } else if (userOrTenant == "Tenant") {
            if (addRemove == "Add") {
                if (ring == "1.5") {
                    if (ring1_5TenantIds.Contains(id)) {
                        log.Info("Tenant with id " + id + " was already provisioned. Skipping this one");
                        problems += "<br />" + id + " already present in Ring 1_5";
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Failure</td><td>Tenant already provisioned</td></tr>");
                    } else {
                        string newLine = "\n" + new String(' ', ring1_5TenantIndentLevel) + "\"" + id + "\", // " + domain + " | " + name + " " + category;
                        ring1_5TenantIds += newLine;
                        ring1_5TenantAdditions += "<br />" + newLine;
                        successEmails = AddIfNotContained(successEmails, submitterEmail);
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Success</td><td></td></tr>");
                    }
                } else {
                    if (ring3TenantIds.Contains(id)) {
                        log.Info("Tenant with id " + id + " was already provisioned. Skipping this one");
                        problems += "<br />" + id + " already present in Ring 3";
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Failure</td><td>Tenant already provisioned</td></tr>");
                    } else {
                        string newLine = "\n" + new String(' ', ring3TenantIndentLevel) + "\"" + id + "\", // " + domain + " | " + name + " " + category;
                        ring3TenantIds += newLine;
                        ring3TenantAdditions += "<br />" + newLine;
                        successEmails = AddIfNotContained(successEmails, submitterEmail);
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Success</td><td></td></tr>");
                    }
                }
            } else if (addRemove == "Remove") {
                if (ring == "1.5") {
                    int ind = -1;
                    if (!ring1_5TenantIds.Contains(id)) {
                        log.Info("Tenant " + id + " not found, skipping its removal");
                        problems += "<br />" + id + " not found in Ring 1_5";
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Failure</td><td>Tenant not found</td></tr>");
                    }
                    while (ring1_5TenantIds.Contains(id)) {
                        ind = ring1_5TenantIds.IndexOf(id);
                        log.Info(ring1_5TenantIds[ind].ToString());
                        while (ring1_5TenantIds[ind] != '\n') {
                            ind -= 1;
                            log.Info(ring1_5TenantIds[ind].ToString());
                        }
                        // Start after the previous newline, and end including the next newline
                        string subst = ring1_5TenantIds.Substring(ind + 1);
                        string thisLine = "\n" + subst.Substring(0, subst.IndexOf("\n"));
                        ring1_5TenantIds = ring1_5TenantIds.Replace(thisLine, "");
                        ring1_5TenantRemovals += "<br /> " + thisLine;
                        successEmails = AddIfNotContained(successEmails, submitterEmail);
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Success</td><td></td></tr>");
                    }
                } else {
                    int ind = -1;
                    if (!ring3TenantIds.Contains(id)) {
                        log.Info("Tenant " + id + " not found, skipping its removal");
                        problems += "<br />" + id + " not found in Ring 3";
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Failure</td><td>Tenant not found</td></tr>");
                    }
                    while (ring3TenantIds.Contains(id)) {
                        ind = ring3TenantIds.IndexOf(id);
                        log.Info(ring3TenantIds[ind].ToString());
                        while (ring3TenantIds[ind] != '\n') {
                            ind -= 1;
                            log.Info(ring3TenantIds[ind].ToString());
                        }
                        // Start after the previous newline, and end including the next newline
                        string subst = ring3TenantIds.Substring(ind + 1);
                        string thisLine = "\n" + subst.Substring(0, subst.IndexOf("\n"));
                        ring3TenantIds = ring3TenantIds.Replace(thisLine, "");
                        ring3TenantRemovals += "<br /> " + thisLine;
                        successEmails = AddIfNotContained(successEmails, submitterEmail);
                        notifyDict[submitterEmail].Add($"<tr><td>{userOrTenant}</td><td>{email}</td><td>{id}</td><td>{ring}</td><td>{addRemove}</td><td>Success</td><td></td></tr>");
                    }
                }
            }
        }
        
    }

    ring1_5UserIds = ReplaceLastOccurrence(ring1_5UserIds, "\", //", "\" //");
    ring1_5UserIds += "\n";

    ring1_5TenantIds = ReplaceLastOccurrence(ring1_5TenantIds, "\", //", "\" //");
    ring1_5TenantIds += "\n";

    ring3UserIds = ReplaceLastOccurrence(ring3UserIds, "\", //", "\" //");
    ring3UserIds += "\n";

    ring3TenantIds = ReplaceLastOccurrence(ring3TenantIds, "\", //", "\" //");
    ring3TenantIds += "\n";

    string editedConfig = config.Replace(originalRing1_5UserIds, ring1_5UserIds);
    editedConfig = editedConfig.Replace(originalRing1_5TenantIds, ring1_5TenantIds);
    editedConfig = editedConfig.Replace(originalRing3UserIds, ring3UserIds);
    editedConfig = editedConfig.Replace(originalRing3TenantIds, ring3TenantIds);
    //log.Info(editedConfig);

    log.Info(ring1_5UserIds);

    log.Info(problems);


    // Get the default branch ID
    string defaultBranchId = "";
    using (var client = new HttpClient())
    {
        client.DefaultRequestHeaders.Add("Authorization", AUTH_STRING);
        var response = await client.GetAsync(GET_DEFAULT_BRANCH_ENDPOINT);
        JObject branchCollection = JObject.Parse(response.Content.ReadAsStringAsync().Result);
        defaultBranchId = branchCollection["value"][0]["objectId"].ToString();
        log.Info(defaultBranchId);
    }

    // Create a new branch
    string branchName = "refs/heads/provisioning-" + DateTime.Now.Date.ToString("yyyy-MM-dd");
    log.Info(branchName);
    using (var client = new HttpClient())
    {
        client.DefaultRequestHeaders.Add("Authorization", AUTH_STRING);
        dynamic d = new {
                name = branchName,
                oldObjectId = "0000000000000000000000000000000000000000",
                newObjectId = defaultBranchId
        };

        //log.Info(d.ToString());
        string vstsJson = JObject.FromObject(d).ToString();
        vstsJson = "[ " + vstsJson + "]";   // Needs to be wrapped in an array
        //log.Info(vstsJson);
        var response = await client.PostAsync(
            CREATE_BRANCH_ENDPOINT,
            new StringContent(vstsJson, System.Text.Encoding.UTF8, "application/json")
        );
        string result = response.Content.ReadAsStringAsync().Result;
        log.Info(result);
    }
    
    // Get the ID for the commit we want to add to
    // (Yep, this is still necessary. We only get the branch ID from previous step, not commit ID)
    string commitId = "";
    string GET_COMMIT_ENDPOINT = GET_COMMIT_BASE + branchName;
    using (var client = new HttpClient())
    {
        client.DefaultRequestHeaders.Add("Authorization", AUTH_STRING);
        var response = await client.GetAsync(GET_COMMIT_ENDPOINT);
        JObject commit = JObject.Parse(response.Content.ReadAsStringAsync().Result);
        commitId = commit["value"][0]["objectId"].ToString();
        //log.Info(commitId);
    }

    // Now create a commit and push it to the branch
    using (var client = new HttpClient())
    {
        // Create a POST for VSTS
        dynamic d = new {
            refUpdates = new [] {
                new {
                    name = branchName,
                    oldObjectId = commitId
                }
            },
            commits = new [] {
                new {
                    comment = "[Automated] Provision and de-provision requested users for TAP",
                    changes = new [] {
                            new {
                            changeType = "edit",
                            item = new {
                                path = CONFIG_PATH
                            },
                            newContent = new {
                                content = editedConfig,
                                contentType = 0
                            }
                        }
                    }
                }
            }
        };

        string vstsJson = JObject.FromObject(d).ToString();

        // Create the request
        client.DefaultRequestHeaders.Add("Authorization", AUTH_STRING);
        //log.Info(vstsJson);
        var response = await client.PostAsync(
            PUSH_CHANGE_ENDPOINT,
            new StringContent(vstsJson, System.Text.Encoding.UTF8, "application/json")
        );

        string result = response.Content.ReadAsStringAsync().Result;
        //log.Info(result);
    }

    string descriptionString = "";
    descriptionString += "Add to Ring 1.5:<br />" + ring1_5UserAdditions;
    descriptionString += "<br /><br />Remove from Ring 1.5:<br />" + ring1_5UserRemovals;
    descriptionString += "<br /><br />Add Users to Ring 3:<br />" + ring3UserAdditions;
    descriptionString += "<br /><br />Remove Users from Ring 3:<br />" + ring3UserRemovals; 
    descriptionString += "<br />Add Tenants to Ring 3:<br />" + ring3TenantAdditions;
    descriptionString += "<br /><br />Remove Tenants from Ring 3:<br />" + ring3TenantRemovals;
    descriptionString += "<br /><br />Problems:<br />" + problems;
    descriptionString += "<br /><br />Notify Successes:<br />" + successEmails;
    descriptionString += "<br /><br />Notify Rejections:<br />" + rejectEmails;

    foreach(KeyValuePair<string, List<string>> entry in notifyDict)
    {
        string table = $"<br /><div id={entry.Key}><p>Notification table for {entry.Key}:</p><br />" + NOTIFICATION_TABLE_HEADER;
        foreach(string row in entry.Value) {
            table += row;
        }
        table += "</tbody></table></div>";
        descriptionString += table;
        descriptionString = descriptionString.Replace("<td>", "<td style='border=1px solid black;'>");

    }

    int taskId = -1;
    string taskUrl = "";
    using (var client = new HttpClient())
    {
        // Create a POST for VSTS
        JObject vstsPost = new JObject();
        JArray body = new JArray();

        JObject body1 = JObject.FromObject(new { 
            op = "add",
            path = "/fields/System.Title",
            value = $"[Automated] TAP: Provision/de-provision users/tenants for TAP ({DateTime.Now.Date.ToString("yyyy-MM-dd")})"
        });

        JObject body2 = JObject.FromObject(new {
            op = "add",
            path = "/fields/System.AreaPath",
            value = "MSTeams\\Customer Feedback"
        });

        JObject body3 = JObject.FromObject(new {
            op = "add",
            path = "/fields/System.IterationPath",
            value = "MSTeams\\Backlog"
        });

        JObject body4 = JObject.FromObject(new {
            op = "add",
            path = "/fields/MicrosoftTeamsCMMI.CentralTriageJustification",
            value = JUSTIFICATION_STRING
        });

        JObject body5 = JObject.FromObject(new {
            op = "add",
            path = "/fields/System.Description",
            value = descriptionString
        });

        JObject body6 = JObject.FromObject(new {
            op = "add",
            path = "/fields/System.AssignedTo",
            value = "Kasi Viswanathan Thirunavukkarasu <v-kathir@microsoft.com>"
        });

        JObject body7 = JObject.FromObject(new {
            op = "add",
            path = "/relations/-",
            value = new {
                rel = "System.LinkTypes.Hierarchy-Reverse",
                url = MASTER_FEATURE_URL
            }
        });

        JObject body8 = JObject.FromObject(new {
            op = "add",
            path = "/fields/System.Tags",
            value = "TAP100Engine; Provisioning"
        });

        body.Add(body1);
        body.Add(body2);
        body.Add(body3);
        body.Add(body4);
        body.Add(body5);
        body.Add(body6);
        body.Add(body7);
        body.Add(body8);

        string vstsJson = body.ToString();
        log.Info(vstsJson);

        // Create the request
        client.DefaultRequestHeaders.Add("Authorization", AUTH_STRING);
        var response = await client.PostAsync(
            CREATE_TASK_ENDPOINT,
            new StringContent(vstsJson, System.Text.Encoding.UTF8, "application/json-patch+json")
        );

        log.Info("Creating a new task: " + response.Content.ReadAsStringAsync().Result);
        taskId = Convert.ToInt32(JObject.Parse(response.Content.ReadAsStringAsync().Result)["id"]);
        taskUrl = JObject.Parse(response.Content.ReadAsStringAsync().Result)["_links"]["self"]["href"].ToString();
    }
    
    // Still need to set the task to be Resolved. First we must make it Active, then Resolved.
    string EDIT_TASK_ENDPOINT = $"https://dev.azure.com/{ORG}/MSTeams/_apis/wit/workitems/{taskId}?{API_VERSION}";
    using (var client = new HttpClient()) {
        JObject vstsPost = new JObject();
        JArray body = new JArray();
        JObject body1 = JObject.FromObject(new {
            op = "add",
            path = "/fields/System.State",
            value = "Active"
        });

        body.Add(body1);
        string vstsJson = body.ToString();

        client.DefaultRequestHeaders.Add("Authorization", AUTH_STRING);

        var request = new HttpRequestMessage(new HttpMethod("PATCH"), EDIT_TASK_ENDPOINT);
        request.Content = new StringContent(vstsJson, System.Text.Encoding.UTF8, "application/json-patch+json");
        var response = await client.SendAsync(request);

        log.Info("Setting task to active: " + response.Content.ReadAsStringAsync().Result);
    }

    using (var client = new HttpClient()) {
        JObject vstsPost = new JObject();
        JArray body = new JArray();
        JObject body1 = JObject.FromObject(new {
            op = "add",
            path = "/fields/System.State",
            value = "Resolved"
        });

        body.Add(body1);
        string vstsJson = body.ToString();

        client.DefaultRequestHeaders.Add("Authorization", AUTH_STRING);
        
        var request = new HttpRequestMessage(new HttpMethod("PATCH"), EDIT_TASK_ENDPOINT);
        request.Content = new StringContent(vstsJson, System.Text.Encoding.UTF8, "application/json-patch+json");
        var response = await client.SendAsync(request);

        log.Info("Setting task to resolved: " + response.Content.ReadAsStringAsync().Result);
    }

    // Finally, open a pull request for our changes, referencing that work item.
    string prId = "";
    using (var client = new HttpClient())
    {
        JObject vstsPost = JObject.FromObject(new {
            sourceRefName = branchName,
            targetRefName = "refs/heads/develop",
            title = "[Automated] TAP: Provision and de-provision users for TAP",
            description = "Automatically generated pull request, adding/removing the requested users to/from Ring 1.5.",
            workitemRefs = new [] {
                new {
                    id = taskId,
                    url = taskUrl
                }
            }
        });

        string vstsJson = vstsPost.ToString();

        log.Info(vstsJson);

        client.DefaultRequestHeaders.Add("Authorization", AUTH_STRING);
        var response = await client.PostAsync(
            CREATE_PR_ENDPOINT,
            new StringContent(vstsJson, System.Text.Encoding.UTF8, "application/json")
        );

        //prId = Convert.ToInt32(JObject.Parse(response.Content.ReadAsStringAsync().Result)["id"]);

        log.Info("Creating a new PR: " + response.Content.ReadAsStringAsync().Result);

        JObject prCreation = JObject.Parse(response.Content.ReadAsStringAsync().Result);
        prId = prCreation["pullRequestId"].ToString();
    }

    // Add Kasi as an approver
    using (var client = new HttpClient()) {
        // Send a PUT request to this url:
        string put_url = $"https://dev.azure.com/domoreexp/Teamspace/_apis/git/repositories/Teamspace-Web/pullRequests/{prId}/reviewers/{REVIEWER_ID}?api-version=5.0";

        JObject vstsPut = JObject.FromObject(new {
            vote = 0
        });
        string putJson = vstsPut.ToString();

        client.DefaultRequestHeaders.Add("Authorization", AUTH_STRING);
        var response = await client.PutAsync(
            put_url,
            new StringContent(putJson, System.Text.Encoding.UTF8, "application/json")
        );

        log.Info("Adding a PR reviewer: " + response.Content.ReadAsStringAsync().Result);
    }

    return req.CreateResponse(HttpStatusCode.OK);
}
