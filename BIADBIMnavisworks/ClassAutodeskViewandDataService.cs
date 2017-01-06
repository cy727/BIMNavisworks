using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Web;
using System.IO;

using RestSharp;

namespace BIADBIMnavisworks
{
    class ClassAutodeskViewandDataService
    {
        const String strClient = "https://developer.api.autodesk.com";
        RestClient _client = new RestClient(strClient);

        public String strConsumerKey = "IfK1GjUH3X9adQtLSA2YCqvmnlxmncEU";
        public String strConsumerSecret = "ATokCTujbNyhTPzq";

        public String _token = "";

        public bool authentication()
        {
            RestRequest authReq = new RestRequest();
            authReq.Resource = "authentication/v1/authenticate";
            authReq.Method = Method.POST;
            authReq.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            authReq.AddParameter("client_id", strConsumerKey);
            authReq.AddParameter("client_secret", strConsumerSecret);
            authReq.AddParameter("grant_type", "client_credentials");

            IRestResponse result = _client.Execute(authReq);
            if (result.StatusCode == System.Net.HttpStatusCode.OK)
            {
                String responseString = result.Content;
                int len = responseString.Length;
                int index = responseString.IndexOf("\"access_token\":\"") + "\"access_token\":\"".Length;
                responseString = responseString.Substring(index, len - index - 1);
                int index2 = responseString.IndexOf("\"");
                _token = responseString.Substring(0, index2);

                //now set the token.
                RestRequest setTokenReq = new RestRequest();
                setTokenReq.Resource = "utility/v1/settoken";
                setTokenReq.Method = Method.POST;
                setTokenReq.AddHeader("Content-Type", "application/x-www-form-urlencoded");
                setTokenReq.AddParameter("access-token", _token);

                IRestResponse resp = _client.Execute(setTokenReq);
                if (resp.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    return true;
                }
            }
            return false;

        }
 

    }
}
