using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net.Http;
using System.Security.Policy;
using Newtonsoft.Json;
using System.Text;
using System.Web.Http;
using System.Web.Http.Results;
using System.Configuration;

namespace KwanWebAPI.Controllers
{
    public class GSTController : ApiController
    {
        

       // [HttpGet]
        [System.Web.Http.Route("api/GSTDetails/getGSTDetails/{GSTNo}")]
        public string getGSTDetails(string GSTNo)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    
                    var endpoint = ConfigurationManager.AppSettings["UAT_GSTURL"];// "https://uat.logitax.in/TransactionAPI/GetGSTINDetails";

                    GST usergst = new GST();
                    usergst.GSTIN = GSTNo; // "29AAACW4202F1ZM";
                    GST[] Us = new GST[1];
                    Us[0] = usergst;

                    UserGST userGST = new UserGST()
                    {
                        USERCODE = ConfigurationManager.AppSettings["UAT_USERCODE"],//"Collective_DEMO",
                        CLIENTCODE = ConfigurationManager.AppSettings["UAT_CLIENTCODE"],
                        //"DOAgw",
                        PASSWORD = ConfigurationManager.AppSettings["UAT_PASSWORD"],
                        //"Collective@123",
                        RequestorGSTIN = ConfigurationManager.AppSettings["UAT_RequestorGSTIN"],
                        //"29AAACW4202F1ZM",
                        gstinlist = Us
                    };
                    var newpostjson = JsonConvert.SerializeObject(userGST);
                    var payload = new StringContent(newpostjson, Encoding.UTF8, "application/json");
                    var result = client.PostAsync(endpoint, payload).Result.Content.ReadAsStringAsync().Result;


                    return result;
                }
                
            }
            catch(Exception ex)
            {
                return "erroe";
            }

            
        }



        [System.Web.Http.Route("api/PANDetails/getPANDetails/{PANNo}")]
        public string getPANDetails(string PANNo)
        {
            try
            {
                using (var client = new HttpClient())
                {

                    var endpoint = ConfigurationManager.AppSettings["UAT_PAN"];// "https://uat.logitax.in/TransactionAPI/GetGSTINDetails";

                    PanCard panCard = new PanCard();
                    panCard.id_number = PANNo;

                    var newpostjson = JsonConvert.SerializeObject(panCard);
                    var payload = new StringContent(newpostjson, Encoding.UTF8, "application/json");
                   // client.DefaultRequestHeaders.Add("Content-Type", "application/json");
                   // client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
                    client.DefaultRequestHeaders.Add("api-key", ConfigurationManager.AppSettings["UAT_PANToken"]);

                    var result = client.PostAsync(endpoint, payload).Result.Content.ReadAsStringAsync().Result;


                    return result;
                }

            }
            catch (Exception ex)
            {
                return "erroe";
            }


        }


        public class UserGST
        {
            public string USERCODE { get; set; }
            public string CLIENTCODE { get; set; }
            public string PASSWORD { get; set; }
            public string RequestorGSTIN { get; set; }
            public GST[] gstinlist { get; set; }



        }
        public class PanCard
        {
            public string id_number { get; set;}
        }



        public class GST
        {
            public string GSTIN { get; set; }
        }
    }
}