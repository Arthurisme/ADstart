using ADstart.Models;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using Microsoft.VisualBasic.FileIO;

namespace ADstart.Controllers
{
    public class HomeController : Controller
    {


        string StrCurrentTestOU = "OU1";
        string feedbackString;

        [HttpPost]
        public IActionResult addGroup(DirectoryModel directory)
        {
            int userId = directory.UserId;
            string name = directory.Name;
            string operatorPerson = directory.OperatorPerson;
            string CN = directory.CN;
            string samAccountName = directory.SamAccountName;
            string gender = directory.Gender;
            string city = directory.City;


            try
            {

                // Bind to the domain that this user is currently connected to.
                //DirectoryEntry dom = new DirectoryEntry("LDAP://adtest", "Administrator", "Testtest01");
                DirectoryEntry dom = new DirectoryEntry("LDAP://192.168.86.130", "Administrator", "Testtest01");

                // Find the container (in this case, the Consulting organizational unit) that you 
                // wish to add the new group to.
                DirectoryEntry ou = dom.Children.Find("OU=" + StrCurrentTestOU);

                dom.RefreshCache();

                // Add the new group Practice Managers.
                DirectoryEntry group = ou.Children.Add("CN=" + CN, "group");

                // Set the samAccountName for the new group.
                group.Properties["samAccountName"].Value = samAccountName;

                // Commit the new group to the directory.
                group.CommitChanges();

            }
            catch (System.Runtime.InteropServices.COMException COMEx)
            {
                // If a COMException is thrown, then the following code example can catch the text of the error.
                // For more information about handling COM exceptions, see Handling Errors.
                Console.WriteLine(COMEx.ErrorCode);
            }

            ////////////////////////////////////////////////////////////////////////////


           

              feedbackString  = "Active Directory is commited successfully. \r" +
                   "You added " + CN + " as Group in the " + StrCurrentTestOU + ".";

            return RedirectToAction("FeedbackMethod", "Home",new { operatorperson= operatorPerson, feedbackstring = feedbackString });
        }

        [HttpPost]
        public IActionResult addOU(DirectoryModel directory)
        {
            int userId = directory.UserId;
            string name = directory.Name;
            string operatorPerson = directory.OperatorPerson;
            string CN = directory.CN;
            string samAccountName = directory.SamAccountName;
            string gender = directory.Gender;
            string city = directory.City;

            string strOUForAdd = directory.OU;
            string strOUDescription = directory.OUDescription;



            try
            {
                // Bind to the domain that this user is currently connected to.
                //DirectoryEntry dom = new DirectoryEntry("LDAP://adtest", "Administrator", "Testtest01");
                DirectoryEntry dom = new DirectoryEntry("LDAP://192.168.86.130", "Administrator", "Testtest01");

                // Find the container (in this case, the Consulting organizational unit) that you 
                // wish to add the new group to.
                DirectoryEntry ou = dom.Children.Find("OU=" + StrCurrentTestOU);


                dom.RefreshCache();


                // Add the new group Practice Managers.
                DirectoryEntry ouForAdd = ou.Children.Add("OU=" + strOUForAdd, "OrganizationalUnit");

                // Set the samAccountName for the new group.
                ouForAdd.Properties["description"].Add(strOUDescription);

                // Commit the new group to the directory.
                ouForAdd.CommitChanges();

            }

            catch (System.Runtime.InteropServices.COMException COMEx)
            {
                // If a COMException is thrown, then the following code example can catch the text of the error.
                // For more information about handling COM exceptions, see Handling Errors.
                Console.WriteLine(COMEx.ErrorCode);
            }

            ////////////////////////////////////////////////////////////////////////////


            feedbackString = "Active Directory is commited successfully. \r "
                + "You added " + strOUForAdd + " as OU in the " + StrCurrentTestOU + ".";

            return RedirectToAction("FeedbackMethod", "Home", new { operatorperson = operatorPerson, feedbackstring = feedbackString });
        }

        [HttpPost]
        public IActionResult createUser(DirectoryModel directory)
        {
            int userId = directory.UserId;
            string operatorPerson = directory.OperatorPerson;
            string name = directory.Name;
            string CN = directory.CN;
            string samAccountName = directory.SamAccountName;
            string gender = directory.Gender;
            string city = directory.City;

            string strOUForAdd = directory.OU;
            string strOUDescription = directory.OUDescription;

            string strUserForAdd = directory.UserName;



            try
            {
                // Bind to the domain that this user is currently connected to.
                //DirectoryEntry dom = new DirectoryEntry("LDAP://adtest", "Administrator", "Testtest01");
                DirectoryEntry dom = new DirectoryEntry("LDAP://192.168.86.130", "Administrator", "Testtest01");

                // Find the container (in this case, the Consulting organizational unit) that you 
                // wish to add the new group to.
                DirectoryEntry ou = dom.Children.Find("OU=" + StrCurrentTestOU);


                dom.RefreshCache();


                // Use the Add method to add a user to an organizational unit.
                DirectoryEntry usr = ou.Children.Add("CN=" + strUserForAdd, "user");
                // Set the samAccountName, then commit changes to the directory.
                usr.Properties["samAccountName"].Value = samAccountName;
                usr.CommitChanges();

                feedbackString = "Active Directory is commited successfully. \r "
               + "You added " + strUserForAdd + " as User in the " + StrCurrentTestOU + ".";

            }

            catch (System.Runtime.InteropServices.COMException COMEx)
            {
                // If a COMException is thrown, then the following code example can catch the text of the error.
                // For more information about handling COM exceptions, see Handling Errors.
                Console.WriteLine(COMEx.ErrorCode);
                feedbackString = "There are some Error on the server or code, please contact admin.";
            }

            ////////////////////////////////////////////////////////////////////////////






            return RedirectToAction("FeedbackMethod", "Home", new { operatorperson = operatorPerson, feedbackstring = feedbackString });
        }

        [HttpPost]
        public IActionResult CreatBulkADUsersFromCSVFile(DirectoryModel directory)
        {

            string operatorPerson = directory.OperatorPerson;




            string csv_File_Path = @"C:\NewUsers.csv";
            TextFieldParser csvReader = new TextFieldParser(csv_File_Path);
            csvReader.SetDelimiters(new string[] { "," });
            csvReader.HasFieldsEnclosedInQuotes = true;

            // reading column fields 
            string[] colFields = csvReader.ReadFields();
            int index_Name = colFields.ToList().IndexOf("Name");
            int index_samaccountName = colFields.ToList().IndexOf("samAccountName");
            int index_ParentOU = colFields.ToList().IndexOf("ParentOU");
            while (!csvReader.EndOfData)
            {



                // reading user fields 
                string[] csvData = csvReader.ReadFields();
                //DirectoryEntry ouEntry = new DirectoryEntry("LDAP://" + csvData[index_ParentOU]);

                // Bind to the domain that this user is currently connected to.
                //DirectoryEntry dom = new DirectoryEntry("LDAP://adtest", "Administrator", "Testtest01");
                DirectoryEntry dom = new DirectoryEntry("LDAP://192.168.86.130", "Administrator", "Testtest01");

                DirectoryEntry ou = dom.Children.Find("OU=" + StrCurrentTestOU);


                try
                {

                    // Use the Add method to add a user to an organizational unit.
                    DirectoryEntry usr = ou.Children.Add("CN=" + csvData[index_Name], "user");
                    // Set the samAccountName, then commit changes to the directory.
                    usr.Properties["samAccountName"].Value = csvData[index_samaccountName];
                    usr.CommitChanges();

                    feedbackString = "Active Directory is commited successfully. \r "
+ "You added bulk users in the " + StrCurrentTestOU + ".";


                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    feedbackString = "There are some Error on the server or code, please contact admin.";
                }
            }
            csvReader.Close();

            TempData["operatorPerson"] = operatorPerson;

            return RedirectToAction("FeedbackMethod", "Home", new { operatorperson = operatorPerson, feedbackstring = feedbackString });
        }


        public IActionResult Directory()
        {
            ViewData["Message"] = "Your Directory page.";

            return View();
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Error()
        {
            return View();
        }

        //public string FeedbackMethod()
        //{
        //    string operatorPerson = TempData["operatorPerson"] as String;
        //    string feedbackString = TempData["feedbackString"] as String;


        //    return operatorPerson + ": You have clicked submit. <br>" +
        //          feedbackString;
        //}

        public string FeedbackMethod(string operatorperson , string feedbackstring)
        {
            string operatorPerson = operatorperson;
            string feedbackString = feedbackstring;


            return operatorPerson + ": You have clicked submit. \n" +
                  feedbackString;
        }


    }
}
