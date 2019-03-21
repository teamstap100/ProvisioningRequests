#r "Newtonsoft.Json"
using System;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("Webhook was triggered!");
    HttpContent requestContent = req.Content;
    string jsonContent = requestContent.ReadAsStringAsync().Result;

    log.Info(jsonContent);

    JObject jo = JObject.Parse(jsonContent);
    //log.Info(jo["resource"]["revision"]["fields"]["System.Description"].ToString());

    string status = jo["resource"]["revision"]["fields"]["System.State"].ToString();
    
    // Only do this when the PR closes!
    if (status != "Closed") {
        log.Info("Not Closed, so not doing anything");
        return req.CreateResponse(HttpStatusCode.OK);
    }


    string desc = jo["resource"]["revision"]["fields"]["System.Description"].ToString();
    string tablePart = desc.Substring(desc.IndexOf("<div id="));
    string[] tables = tablePart.Split(new string[] { "<div id=\"" }, StringSplitOptions.None);
    log.Info("About to look at tables");
    foreach(string table in tables) {
        if (table == "") {
            continue;
        }
        //log.Info("Here's a table: " + table);
        string emailRecipient = table.Split('"')[0];

        log.Info(table);
        string finalTable = table.Substring(table.IndexOf("<table"));
        finalTable = finalTable.Remove(finalTable.IndexOf("</div>"), 6);
        //finalTable = finalTable.Replace("<td", "<td style='border:black solid 1px;'");
        log.Info("Recipient: " + emailRecipient);
        log.Info(finalTable);

        log.Info(" ");

        JObject post = JObject.FromObject(new {
            recipient = emailRecipient,
            table = finalTable
        });

        string postJson = post.ToString();

        using (var client = new HttpClient())
        {
            var response = await client.PostAsync(
                "https://prod-29.westcentralus.logic.azure.com:443/workflows/8838916ff2634fbb9a651b13b199a3ba/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ChdP4-n9HbGa1npqxTey7SKvwLSdrQk2kleICo8pCto", 
                new StringContent(postJson, Encoding.UTF8, "application/json"));
        }
    }

    return req.CreateResponse(HttpStatusCode.OK);
}
