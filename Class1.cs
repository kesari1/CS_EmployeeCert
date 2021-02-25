using Deltek.Vision.WorkflowAPI.Server;
using Deltek.Vision.Ancestors.Server;
using Deltek.Framework.API.Server;
using Deltek.Framework.Workflow.Server;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System;
using System.Linq;

using System.Net;
using System.Security;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;

namespace CS_EmployeeCerts
{
    public class Class1 : WorkflowBaseClass
    {
        public void CreateEmployeeLicense(string empID, string empLicense, string empEarnerdDate, string empState, string empLicNumber, string empExpiryDate, string empCertificatesSearch, string empCourseListSearch)

        {
            try
            {






                string login = "sp-designer"; //give your username here  
                string password = "Consor@2019!"; //give your password  
                                                  // var securePassword = new SecureString();
                                                  // foreach (char c in password)
                                                  // {
                                                  //      securePassword.AppendChar(c);
                                                  // }


                // AddInformation("NETWORK CRED");
                //  AddError("NETWORK CRED");
                NetworkCredential myCred = new NetworkCredential(login, password, "consor");
                CredentialCache myCache = new CredentialCache();

                myCache.Add(new Uri("https://portal.consoreng.com"), "Basic", myCred);

                string siteUrl = "https://portal.consoreng.com";
                //AddInformation("After site url");
                DateTime dateExp;

                if (string.IsNullOrEmpty(empExpiryDate)) { empExpiryDate = "1/1/1999"; }
                else
                {
                    DateTime.TryParse(empExpiryDate, out dateExp);
                    empExpiryDate = dateExp.Date.ToString();
                }
                if (string.IsNullOrEmpty(empEarnerdDate)) { empEarnerdDate = "1/1/1999"; }
                else
                {
                    DateTime.TryParse(empEarnerdDate, out dateExp);
                    empEarnerdDate = dateExp.Date.ToString();
                }

                // if (string.IsNullOrEmpty(empLastRenewal)) { empLastRenewal = "1/1/1999"; }
                // else
                //{
                //      DateTime.TryParse(empLastRenewal, out dateExp);
                //      empLastRenewal = dateExp.Date.ToString();
                //  }
                string listName = "CS_EmployeeCert";
                ClientContext clientContext = new ClientContext(siteUrl);
                SP.List oList = clientContext.Web.Lists.GetByTitle(listName);
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);
                oListItem["EmployeeID"] = empID;
                oListItem["Experience"] = empLicense;
                oListItem["Earned"] = empEarnerdDate;
                oListItem["State"] = empState;
                oListItem["LicenseNumber"] = empLicNumber;
                oListItem["Expiration"] = empExpiryDate;
                oListItem["CertificatesSearch"] = empCertificatesSearch;
                oListItem["CourseListSearch"] = empCourseListSearch;

                //oListItem["LastRenewal"] = empLastRenewal;
                oListItem.Update();
                clientContext.Credentials = myCred;
                clientContext.ExecuteQuery();
            }

            catch (Exception e)
            {
                AddError(e.Message + e.Source);
            }
        }

        public void UpdateByLicenseNumber(string empID, string empLicense, string empEarnerdDate, string empState, string empLicNumber, string empExpiryDate, string empCertificatesSearch, string empCourseListSearch)
        {
            try
            {
                string login = "sp-designer"; //give your username here  
                string password = "Consor@2019!"; //give your password  
                                                  // var securePassword = new SecureString();
                                                  // foreach (char c in password)
                                                  // {
                                                  //      securePassword.AppendChar(c);
                                                  // }

                DateTime ntpdate2 = DateTime.Now;
                DateTime endDate2 = DateTime.Now;

                NetworkCredential myCred = new NetworkCredential(login, password, "consor");
                CredentialCache myCache = new CredentialCache();

                myCache.Add(new Uri("https://portal.consoreng.com"), "Basic", myCred);

                string siteUrl = "https://portal.consoreng.com";

                DateTime dateExp;
                string listName = "CS_EmployeeCert";

                if (string.IsNullOrEmpty(empExpiryDate)) { empExpiryDate = "1/1/1999"; }
                else
                {
                    DateTime.TryParse(empExpiryDate, out dateExp);
                    empExpiryDate = dateExp.Date.ToString();
                }
                if (string.IsNullOrEmpty(empEarnerdDate)) { empEarnerdDate = "1/1/1999"; }
                else
                {
                    DateTime.TryParse(empEarnerdDate, out dateExp);
                    empEarnerdDate = dateExp.Date.ToString();
                }

                //if (string.IsNullOrEmpty(empLastRenewal)) { empLastRenewal = "1/1/1999"; }
                //else
                //{
                //    DateTime.TryParse(empLastRenewal, out dateExp);
                //    empLastRenewal = dateExp.Date.ToString();
                //}


                ClientContext clientContext = new ClientContext(siteUrl);
                SP.List oList = clientContext.Web.Lists.GetByTitle(listName);

                CamlQuery camlQuery = new CamlQuery();
                string EmployeeID = empID;

                string QueryString = "<View><Query><Where>" +
                         "<Eq>" +
                           "<FieldRef Name=\"LicenseNumber\"/>" +
                            "<Value Type=\"Text\">" + empLicNumber + "</Value>" +
                         "</Eq>" +
                        "</Where></Query></View>";
                camlQuery.ViewXml = QueryString;
                ListItemCollection collListItem = oList.GetItems(camlQuery);

                clientContext.Load(collListItem);
                clientContext.Credentials = myCred;

                // Retrieve SP Timezone 
                var spTimeZone = clientContext.Web.RegionalSettings.TimeZone;
                clientContext.Load(spTimeZone);
                clientContext.ExecuteQuery();

                // SP Time Zone 
                // Resolve SP time Zone Code 

                var fixedTimeZoneName = spTimeZone.Description.Replace("and", "&");
                var timeZoneInfo = TimeZoneInfo.GetSystemTimeZones().FirstOrDefault(tz => tz.DisplayName == fixedTimeZoneName);

                // End SP TimeZone 


                clientContext.ExecuteQuery();
                // AddInformation("Query Executed");
                foreach (ListItem oListItem in collListItem)
                {

                    oListItem["EmployeeID"] = empID;
                    oListItem["Experience"] = empLicense;
                    oListItem["Earned"] = empEarnerdDate;
                    oListItem["State"] = empState;
                    oListItem["LicenseNumber"] = empLicNumber;
                    oListItem["Expiration"] = empExpiryDate;
                    oListItem["CertificatesSearch"] = empCertificatesSearch;
                    oListItem["CourseListSearch"] = empCourseListSearch;
                    //  oListItem["LastRenewal"] = empLastRenewal;
                    oListItem.Update();
                    clientContext.Credentials = myCred;
                    clientContext.ExecuteQuery();
                    return;

                }
                //  if (collListItem.Count == 0)
                //  {
                CreateEmployeeLicense(empID, empLicense, empEarnerdDate, empState, empLicNumber, empExpiryDate, empCertificatesSearch, empCourseListSearch);
                //  }

            }
            catch (Exception e)
            {
                // LogTosp("Error", e.Message.ToString());
                AddError(e.Message + e.Source);

            }

        }

        /// <summary>
        /// Send Parameters Separated by comma
        /// </summary>
        /// <param name="empID"></param>
        /// <param name="empLicense"></param>
        /// <param name="empEarnerdDate"></param>
        /// <param name="empState"></param>
        /// <param name="empLicNumber"></param>
        /// <param name="empExpiryDate"></param>
        /// <param name="empLastRenewal"></param>
        public void updateListOfLicense(string empID, string empLicense, string empEarnerdDate, string empState, string empLicNumber, string empExpiryDate, string empCertificatesSearch, string empCourseListSearch)
        {
            if (!string.IsNullOrEmpty(empID))
            {

                string[] empIDs = empID.Split(',');
                string[] empEarnerdDates = empEarnerdDate.Split(',');
                string[] empLicenses = empLicense.Split(',');
                string[] empStates = empState.Split(',');
                string[] empLicNumbers = empLicNumber.Split(',');
                string[] empExpiryDates = empExpiryDate.Split(',');


                int limit = empIDs.Length;



                for (int i = 0; i < limit; i++)
                {


                    License license = new License();

                    license.EmployeeId = empIDs[i];
                    license.empLicense = empLicenses[i];
                    license.empEarnerdDate = empEarnerdDates[i];
                    license.empState = empStates[i];
                    license.empLicNumber = empLicNumbers[i];
                    license.empExpiryDate = empExpiryDates[i];
                    license.empCertificatesSearch = empCertificatesSearch;
                    license.empCourseListSearch = empCourseListSearch;


                    UpdateByLicenseNumber(license.EmployeeId, license.empLicense, license.empEarnerdDate, license.empState, license.empLicNumber, license.empExpiryDate, empCertificatesSearch, empCourseListSearch);


                }
            }
        }
        public void updateListOfCertification(string empCertID, string empID, string empCertAgency, string empCertNumber, string empCertTitle, string empExpiryDate, string empNoExpiry, string empCustCode, string empCertificatesSearch)
        {

            if (!string.IsNullOrEmpty(empID))
            {
                string[] empIDs = empID.Split(',');
                string[] empExpiryDates = empExpiryDate.Split(',');
                string[] empCertIDs = empCertID.Split(',');
                string[] empCertAgencys = empCertAgency.Split(',');
                string[] empCertNumbers = empCertNumber.Split(',');
                string[] empCertTitles = empCertTitle.Split(',');
                string[] empNoExpirys = empNoExpiry.Split(',');
                string[] empCustCodes = empCustCode.Split(',');
                int limit = empIDs.Length;
                for (int i = 0; i < limit; i++)
                {

                    {
                        Certificate certificate = new Certificate();

                        certificate.empID = empIDs[i];
                        certificate.empExpiryDate = empExpiryDates[i];
                        certificate.empCertID = empCertIDs[i];
                        certificate.empCertAgency = empCertAgencys[i];
                        certificate.empCertNumber = empCertNumbers[i];
                        certificate.empCertTitle = empCertTitles[i];
                        certificate.empNoExpiry = empNoExpirys[i];
                        certificate.empCustCode = empCustCodes[i];
                        certificate.empCertificatesSearch = empCertificatesSearch;

                        UpdateByCertID(certificate.empCertID, certificate.empID, certificate.empCertAgency, certificate.empCertNumber, certificate.empCertTitle, certificate.empExpiryDate, certificate.empNoExpiry, certificate.empCustCode, certificate.empCertificatesSearch);
                    }


                }
            }

            // Create Object for Certifications 
            // New Method

            //******************************************************************************************************************************************************************************
        }
        public void CreateEmployeeCertifications(string empCertID, string empID, string empCertAgency, string empCertNumber, string empCertTitle, string empExpiryDate, string empNoExpiry, string empCustCode, string empCertificatesSearch)
        {
            try
            {






                string login = "sp-designer"; //give your username here  
                string password = "Consor@2019!"; //give your password  
                                                  // var securePassword = new SecureString();
                                                  // foreach (char c in password)
                                                  // {
                                                  //      securePassword.AppendChar(c);
                                                  // }


                // AddInformation("NETWORK CRED");
                //  AddError("NETWORK CRED");
                NetworkCredential myCred = new NetworkCredential(login, password, "consor");
                CredentialCache myCache = new CredentialCache();

                myCache.Add(new Uri("https://portal.consoreng.com"), "Basic", myCred);

                string siteUrl = "https://portal.consoreng.com";
                //AddInformation("After site url");
                DateTime dateExp;

                if (string.IsNullOrEmpty(empExpiryDate)) { empExpiryDate = "1/1/1999"; }
                else
                {
                    DateTime.TryParse(empExpiryDate, out dateExp);
                    empExpiryDate = dateExp.Date.ToString();
                }

                string listName = "CS_EmployeeCert";
                ClientContext clientContext = new ClientContext(siteUrl);
                SP.List oList = clientContext.Web.Lists.GetByTitle(listName);
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);
                oListItem["EmployeeID"] = empID;
                oListItem["CertID"] = empCertID;
                oListItem["CertAgency"] = empCertAgency;
                oListItem["CertNumber"] = empCertNumber;
                oListItem["CertTitle"] = empCertTitle;
                oListItem["Expiration"] = empExpiryDate;
                oListItem["NoExpiry"] = empNoExpiry;
                oListItem["CertCode"] = empCustCode;
                oListItem["CertificatesSearch"] = empCertificatesSearch;
                oListItem.Update();
                clientContext.Credentials = myCred;
                clientContext.ExecuteQuery();
            }

            catch (Exception e)
            {
                AddError(e.Message + e.Source);
            }
        }



        public void UpdateByCertID(string empCertID, string empID, string empCertAgency, string empCertNumber, string empCertTitle, string empExpiryDate, string empNoExpiry, string empCustCode, string empCertificatesSearch)
        {
            try
            {
                string login = "sp-designer"; //give your username here  
                string password = "Consor@2019!"; //give your password  
                                                  // var securePassword = new SecureString();
                                                  // foreach (char c in password)
                                                  // {
                                                  //      securePassword.AppendChar(c);
                                                  // }

                DateTime ntpdate2 = DateTime.Now;
                DateTime endDate2 = DateTime.Now;

                NetworkCredential myCred = new NetworkCredential(login, password, "consor");
                CredentialCache myCache = new CredentialCache();

                myCache.Add(new Uri("https://portal.consoreng.com"), "Basic", myCred);

                string siteUrl = "https://portal.consoreng.com";

                DateTime dateExp;
                string listName = "CS_EmployeeCert";

                if (string.IsNullOrEmpty(empExpiryDate)) { empExpiryDate = "1/1/1999"; }
                else
                {
                    DateTime.TryParse(empExpiryDate, out dateExp);
                    empExpiryDate = dateExp.Date.ToString();
                }

                ClientContext clientContext = new ClientContext(siteUrl);
                SP.List oList = clientContext.Web.Lists.GetByTitle(listName);

                CamlQuery camlQuery = new CamlQuery();
                string EmployeeID = empID;

                string QueryString = "<View><Query><Where>" +
                         "<Eq>" +
                           "<FieldRef Name=\"CertID\"/>" +
                            "<Value Type=\"Text\">" + empCertID + "</Value>" +
                         "</Eq>" +
                        "</Where></Query></View>";
                camlQuery.ViewXml = QueryString;
                ListItemCollection collListItem = oList.GetItems(camlQuery);

                clientContext.Load(collListItem);
                clientContext.Credentials = myCred;

                // Retrieve SP Timezone 
                var spTimeZone = clientContext.Web.RegionalSettings.TimeZone;
                clientContext.Load(spTimeZone);
                clientContext.ExecuteQuery();

                // SP Time Zone 
                // Resolve SP time Zone Code 

                var fixedTimeZoneName = spTimeZone.Description.Replace("and", "&");
                var timeZoneInfo = TimeZoneInfo.GetSystemTimeZones().FirstOrDefault(tz => tz.DisplayName == fixedTimeZoneName);

                // End SP TimeZone 


                clientContext.ExecuteQuery();
                // AddInformation("Query Executed");
                foreach (ListItem oListItem in collListItem)
                {

                    oListItem["EmployeeID"] = empID;
                    oListItem["CertID"] = empCertID;
                    oListItem["CertAgency"] = empCertAgency;
                    oListItem["CertNumber"] = empCertNumber;
                    oListItem["CertTitle"] = empCertTitle;
                    oListItem["Expiration"] = empExpiryDate;
                    oListItem["NoExpiry"] = empNoExpiry;
                    oListItem["CertCode"] = empCustCode;
                    oListItem["CertificatesSearch"] = empCertificatesSearch;
                    oListItem.Update();
                    clientContext.Credentials = myCred;
                    clientContext.ExecuteQuery();
                    return;

                }
                //  if (collListItem.Count == 0)
                //  {
                CreateEmployeeCertifications(empCertID, empID, empCertAgency, empCertNumber, empCertTitle, empExpiryDate, empNoExpiry, empCustCode, empCertificatesSearch);
                // }

            }
            catch (Exception e)
            {
                // LogTosp("Error", e.Message.ToString());
                AddError(e.Message + e.Source);

            }
        }
        // *********************************************************************************************************************************
        //public void CreateEmployeeCourse(string empID, string agency, string courseno, string coursename, string date, string classcost, string miscscost, string courseseq)
        //{
        //    try
        //    {






        //        string login = "sp-designer"; //give your username here  
        //        string password = "Consor@2019!"; //give your password  
        //                                          // var securePassword = new SecureString();
        //                                          // foreach (char c in password)
        //                                          // {
        //                                          //      securePassword.AppendChar(c);
        //                                          // }


        //        // AddInformation("NETWORK CRED");
        //        //  AddError("NETWORK CRED");
        //        NetworkCredential myCred = new NetworkCredential(login, password, "consor");
        //        CredentialCache myCache = new CredentialCache();

        //        myCache.Add(new Uri("https://portal.consoreng.com"), "Basic", myCred);

        //        string siteUrl = "https://portal.consoreng.com";
        //        //AddInformation("After site url");
        //        DateTime dateExp;

        //        if (string.IsNullOrEmpty(date)) { date = "1/1/1999"; }
        //        else
        //        {
        //            DateTime.TryParse(date, out dateExp);
        //            date = dateExp.Date.ToString();
        //        }
        //        string listName = "EmployeeCourses";
        //        ClientContext clientContext = new ClientContext(siteUrl);
        //        SP.List oList = clientContext.Web.Lists.GetByTitle(listName);
        //        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
        //        ListItem oListItem = oList.AddItem(itemCreateInfo);
        //        oListItem["EmployeeID"] = empID;
        //        oListItem["Agency"] = agency;
        //        oListItem["date"] = date;
        //        oListItem["CourseNo"] = courseno;
        //        oListItem["CourseName"] = coursename;
        //        oListItem["ClassCost"] = classcost;
        //        oListItem["MiscCost"] = miscscost;
        //        oListItem["CourseSeq"] = courseseq;

        //        oListItem.Update();
        //        clientContext.Credentials = myCred;
        //        clientContext.ExecuteQuery();
        //    }

        //    catch (Exception e)
        //    {
        //        AddError(e.Message + e.Source);
        //    }
        //}
        //*************************************Update By Course ID ************************************************************

        //public void UpdateByCourseID(string empID, string agency, string courseno, string coursename, string date, string classcost, string miscscost, string courseseq)
        //{
        //    try
        //    {
        //        string login = "sp-designer"; //give your username here  
        //        string password = "Consor@2019!"; //give your password  
        //                                          // var securePassword = new SecureString();
        //                                          // foreach (char c in password)
        //                                          // {
        //                                          //      securePassword.AppendChar(c);
        //                                          // }

        //        DateTime ntpdate2 = DateTime.Now;
        //        DateTime endDate2 = DateTime.Now;

        //        NetworkCredential myCred = new NetworkCredential(login, password, "consor");
        //        CredentialCache myCache = new CredentialCache();

        //        myCache.Add(new Uri("https://portal.consoreng.com"), "Basic", myCred);

        //        string siteUrl = "https://portal.consoreng.com";

        //        DateTime dateExp;
        //        string listName = "EmployeeCourses";

        //        if (string.IsNullOrEmpty(date)) { date = "1/1/1999"; }
        //        else
        //        {
        //            DateTime.TryParse(date, out dateExp);
        //            date = dateExp.Date.ToString();
        //        }

        //        ClientContext clientContext = new ClientContext(siteUrl);
        //        SP.List oList = clientContext.Web.Lists.GetByTitle(listName);

        //        CamlQuery camlQuery = new CamlQuery();
        //        string EmployeeID = empID;

        //        string QueryString = "<View><Query><Where>" +
        //                 "<Eq>" +
        //                   "<FieldRef Name=\"CourseSeq\"/>" +
        //                    "<Value Type=\"Text\">" + courseseq + "</Value>" +
        //                 "</Eq>" +
        //                "</Where></Query></View>";
        //        camlQuery.ViewXml = QueryString;
        //        ListItemCollection collListItem = oList.GetItems(camlQuery);

        //        clientContext.Load(collListItem);
        //        clientContext.Credentials = myCred;

        //        // Retrieve SP Timezone 
        //        var spTimeZone = clientContext.Web.RegionalSettings.TimeZone;
        //        clientContext.Load(spTimeZone);
        //        clientContext.ExecuteQuery();

        //        // SP Time Zone 
        //        // Resolve SP time Zone Code 

        //        var fixedTimeZoneName = spTimeZone.Description.Replace("and", "&");
        //        var timeZoneInfo = TimeZoneInfo.GetSystemTimeZones().FirstOrDefault(tz => tz.DisplayName == fixedTimeZoneName);

        //        // End SP TimeZone 


        //        clientContext.ExecuteQuery();
        //        // AddInformation("Query Executed");
        //        foreach (ListItem oListItem in collListItem)
        //        {

        //            oListItem["EmployeeID"] = empID;
        //            oListItem["Agency"] = agency;
        //            oListItem["date"] = date;
        //            oListItem["CourseNo"] = courseno;
        //            oListItem["CourseName"] = coursename;
        //            oListItem["ClassCost"] = classcost;
        //            oListItem["MiscCost"] = miscscost;
        //            oListItem["CourseSeq"] = courseseq;
        //            oListItem.Update();
        //            clientContext.Credentials = myCred;
        //            clientContext.ExecuteQuery();
        //            return;

        //        }
        //        //  if (collListItem.Count == 0)
        //        //  {
        //        CreateEmployeeCourse(empID, agency, courseno, coursename, date, classcost, miscscost, courseseq);
        //        // }

        //    }
        //    catch (Exception e)
        //    {
        //        // LogTosp("Error", e.Message.ToString());
        //        AddError(e.Message + e.Source);

        //    }
        //}
        //********************************************************End Update Method ****************************************************************************************

        //*********************************************Update by Course ********************************************************
        //public void updateListOfCourse(string empID, string agency, string courseno, string coursename, string date, string classcost, string miscscost, string courseseq)
        //{

        //    if (!string.IsNullOrEmpty(empID))
        //    {
        //        string[] empIDs = empID.Split(',');
        //        string[] agencys = agency.Split(',');
        //        string[] coursenos = courseno.Split(',');
        //        string[] coursenames = coursename.Split(',');
        //        string[] dates = date.Split(',');
        //        string[] classcosts = classcost.Split(',');
        //        string[] miscscosts = miscscost.Split(',');
        //        string[] courseseqs = courseseq.Split(',');
        //        int limit = empIDs.Length;
        //        for (int i = 0; i < limit; i++)
        //        {

        //            {
        //                Course course = new Course();

        //                course.empID = empIDs[i];
        //                course.agency = agencys[i];
        //                course.courseno = coursenos[i];
        //                course.coursename = coursenames[i];
        //                course.date = dates[i];
        //                course.classcost = classcosts[i];
        //                course.misccost = miscscosts[i];
        //                course.courseseq = courseseqs[i];

        //                UpdateByCourseID(course.empID, course.agency, course.courseno, course.coursename, course.date, course.classcost, course.misccost, course.courseseq);
        //            }


        //        }
        //    }
        //}

        /// <summary>
        /// License Object
        /// </summary>
        public class License
        {

            public string EmployeeId { get; set; }
            public string empLicense { get; set; }
            public string empEarnerdDate { get; set; }
            public string empState { get; set; }
            public string empLicNumber { get; set; }
            public string empExpiryDate { get; set; }
            public string empCertificatesSearch { get; set; }
            public string empCourseListSearch { get; set; }





        }


        /// <summary>
        /// Certification Object
        /// </summary>
        public class Certificate
        {

            public string empID { get; set; }
            public string empCertID { get; set; }
            public string empCertAgency { get; set; }
            public string empCertNumber { get; set; }
            public string empCertTitle { get; set; }
            public string empExpiryDate { get; set; }
            public string empNoExpiry { get; set; }
            public string empCustCode { get; set; }
            public string empCertificatesSearch { get; set; }
        }

        public class Course
        {

            public string empID { get; set; }
            public string agency { get; set; }
            public string courseno { get; set; }
            public string coursename { get; set; }
            public string date { get; set; }
            public string classcost { get; set; }
            public string misccost { get; set; }
            public string courseseq { get; set; }

        }



    }

}

