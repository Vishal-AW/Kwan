using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net.Http;
using System.Security.Policy;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text;
using System.Web.Http;
using System.Web.Http.Results;
using System.Configuration;
using static KwanWebAPI.Controllers.GSTController;
using System.Net;

namespace KwanWebAPI.Controllers
{
    public class IRNController : ApiController
    {
        // GET: IRN



        [System.Web.Http.Route("api/GSTDetails/GenerateIRN")]
        public string GenerateIRN(JObject Jsondata)
        {


            try
            {
                using (var client = new HttpClient())
                {

                    // var endpoint = ConfigurationManager.AppSettings["UAT_IRNURL"];// "https://uat.logitax.in/TransactionAPI/GetGSTINDetails";


                    var endpoint = Jsondata["URL"].ToString();

                    FinalData finalData = new FinalData()
                    {   
                        USERCODE = ConfigurationManager.AppSettings["UAT_USERCODE"],//"Collective_DEMO",
                        CLIENTCODE = ConfigurationManager.AppSettings["UAT_CLIENTCODE"],
                        PASSWORD = ConfigurationManager.AppSettings["UAT_PASSWORD"],
                        LID= Jsondata["LID"].ToString(),
                        TransactionType =Jsondata["TransactionType"].ToString(),
                        ListName = Jsondata["ListName"].ToString(),
                        json_data = Jsondata["json_data"]
                    };


                    var newpostjson = JsonConvert.SerializeObject(finalData);
                    var payload = new StringContent(newpostjson, Encoding.UTF8, "application/json");
                    var result = client.PostAsync(endpoint, payload).Result.Content.ReadAsStringAsync().Result;


                    return result;
                }

            }
            catch (Exception ex)
            {
                return "erroe";
            }
        }
     }


    public class FinalData
    {
        public string USERCODE { get; set; }
        public string CLIENTCODE { get; set; }
        public string PASSWORD { get; set; }
        public string LID { get; set; }
        public string TransactionType { get; set; }
        public string ListName { get; set; }
        public object json_data { get; set; }



    }
}