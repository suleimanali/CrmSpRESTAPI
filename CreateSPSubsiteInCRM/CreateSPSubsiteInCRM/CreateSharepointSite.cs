using System;
using System.Collections.Generic;
using System.Linq;
using System;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System.Net;
using System.IO;

namespace CreateSPSubsiteInCRM
{
    internal class CreateSharepointSite
    {
        internal static bool CreateSubsite(IOrganizationService service, string siteName, string urlString)
        {
            var siteUrl = urlString; //@"sharepointUrl";
            var _username = @"crmuserName";
            var _password = "crmPassword";
            var spSite = new Uri(siteUrl);
            var restURL = siteUrl + "/_api/web/webs/add";

            var stringData = string.Concat("{'parameters': { '__metadata': { 'type': 'SP.WebCreationInformation' },",
                "'Title': '" + siteName + "', 'Url': '" + siteName + "', 'WebTemplate': 'siteTemplate', 'UseSamePermissionsAsParentSite': true } }");

            var endpointRequest = (HttpWebRequest)HttpWebRequest.Create(restURL);
            var spo = SpoAuthUtility.Create(spSite, _username, WebUtility.HtmlEncode(_password), false);
            var cookies = spo.GetCookieContainer();
            var formDigest = spo.GetRequestDigest();
            endpointRequest.CookieContainer = cookies;
            endpointRequest.Method = "POST";
            endpointRequest.ContentLength = stringData.Length;
            endpointRequest.Accept = "application/json;odata=verbose";
            endpointRequest.ContentType = "application/json;odata=verbose";
            endpointRequest.Headers.Add("X-RequestDigest", formDigest);

            var writer = new StreamWriter(endpointRequest.GetRequestStream());
            writer.Write(stringData);
            writer.Flush();

            var endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
            try
            {
                var webResponse = endpointRequest.GetResponse();
                var webStream = webResponse.GetResponseStream();
                var responseReader = new StreamReader(webStream);
                var response = responseReader.ReadToEnd();
                var jobj = JObject.Parse(response);
                responseReader.Close();
                return true;

            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException("Inside create SHarpoint Plugin " + e.Message);
            }

        }
    }

}
