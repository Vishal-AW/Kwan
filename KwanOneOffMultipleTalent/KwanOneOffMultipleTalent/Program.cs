using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KwanOneOffMultipleTalent.Application;
namespace KwanOneOffMultipleTalent
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<ErrorApplication> errorList = new List<ErrorApplication>();

            try
            {
                var siteUrl = ConfigurationManager.AppSettings["SP_Address_Live"].ToString();
                var ListName = ConfigurationManager.AppSettings["ListName"].ToString();
                Task<List<ErrorApplication>> errorModel = sharepointOperation.GetActiveErrorListAsync(siteUrl, ListName);

                if (errorModel.Result.Count > 0) // Access the result using Result property
                {
                    //sharepointOperation.getSharepointLibraryExcel(siteUrl, ListName);
                }
            }
            catch (Exception e)
            { 

            }
        }
    }
}