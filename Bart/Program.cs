/*
 * BART - Active Directory to SQL users exporter.
 * This script exports users records from Active Directory in Windows and inserts them in to SQL Server Express database.
 * It will allow to lookup users when new tickets are created by technicians.
 */

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;

namespace Bart
{
    class Program
    {
        //public static string GlobalFilePath = "@\"C:\\userInfo.txt\"";

        static void Main(string[] args)
        {
            //Step 1. clear file before running export
            System.IO.File.WriteAllText("C:\\userInfo.txt", string.Empty);

            //Step 2. Execute export from AD to file
            getGroupMembers(); 

            //Step 3. SQL part. Insert values from files to SQL
            insertUsersToSQL();

            //Step 4.clear file before running export for WebappOperators Group
            System.IO.File.WriteAllText("C:\\WebAppOperatorsUsers.txt", string.Empty);

            //Step 5. Read contents of WebAppOperators group
            //         Read each member properties
            //          Save them to text file
            getWebAppOperatorsMembers();

            //Step 6. Read WebAppOperatorsUsers.txt contents and insert infromation to SQL if not already there
            insertWAOPUsersToSQL();
            
            Console.WriteLine("*********************** ALL DONE");
            Console.ReadLine();
        }


        /********************************************************************
         * Function to access and read users details from Active Directory.
         * It reads Group members and passed this information to 
         * getUserProperties
         ********************************************************************/
        public static void getGroupMembers()
        {
            Console.WriteLine("Hello, this script will now start exporting. Please press a key to continue");
            Console.ReadKey();

            // ################## PrincipalContext method ########################
            //http://stackoverflow.com/questions/4901749/get-user-names-in-an-active-directory-group-via-net
            //http://blogs.technet.com/b/brad_rutkowski/archive/2008/04/15/c-getting-members-of-a-group-the-easy-way-with-net-3-5-discussion-groups-nested-recursive-security-groups-etc.aspx



            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "project.local");
            GroupPrincipal grp = GroupPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, "WebAppUsers");

            if (grp != null)
            {
                foreach (Principal p in grp.GetMembers())
                {
                    string uname = String.Format("{0}", p.SamAccountName);
                    Console.WriteLine("The user name found is: " + uname);
                    getUserProperties(uname);
                }

                grp.Dispose();
                ctx.Dispose();
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("\nWe did not find that group in that domain, perhaps the group resides in a different domain?");
                Console.ReadLine();
            }
        }


               

        /********************************************************************
         * getUserProperties function reads users information passed by 
         * getGroupMembers function and saves their details to a file 
         * on a disk.
         ********************************************************************/
        public static void getUserProperties(string unamep)
        {
           
            PrincipalContext cty = new PrincipalContext(ContextType.Domain, "project.local");
            UserPrincipal usr = UserPrincipal.FindByIdentity(cty, unamep);
            
            string cn = usr.Name;
            string givenName = usr.GivenName ;
            string sn = usr.Surname;
            string sAMAccountName = unamep;
            string mail = usr.EmailAddress;
            string telephoneNumber = usr.VoiceTelephoneNumber;
            //string physicalDeliveryOfficeName = ""; // requires DirectoryEntry object

            string fullUserInfo = cn + "," + givenName + "," + sn + "," + sAMAccountName + "," + mail + "," + telephoneNumber;
            Console.WriteLine (fullUserInfo);
            saveADUsersToFile(fullUserInfo, cn);  // call function to save record to text file  

        }



        /********************************************************************
         * saveADUsersToFile receives full user information from getUserProperties
         * and saves this infotmation as string, coma delimited, to a text file
         ********************************************************************/
        public static void saveADUsersToFile(string userInfo, string uname)
        {

            using (System.IO.StreamWriter file = new System.IO.StreamWriter("C:\\userInfo.txt", true))
            {
                file.WriteLine(userInfo);
            }
            
            //System.IO.File.WriteAllText(@"C:\userInfo.txt", userInfo);
            Console.WriteLine("User " + uname + " has been saved to file");

        }

        /*
         * Open file
         * Open SQL connection
         * read line to a string.
         * Delimit each element by coma and insert it to an array[5]. Element [2] is unique.
         * array[5] user = {"","","","","",""}
         * Run SQL Query to:
         *  -check if element[2] exist. IF exists read next line
         *  -if does not exisit, do insert
         * 
         * 
         * 
         * Close SQL connection
         */
        public static void insertUsersToSQL()
        {

            //https://msdn.microsoft.com/en-us/library/aa287535(v=vs.71).aspx
            //Open file from C:\

            string line;
            string[] lineArr = new string[6];


            //SQL Connection Create            
            System.Data.SqlClient.SqlConnection sqlconn1 = new System.Data.SqlClient.SqlConnection("Server=lk-dit-ad.project.local;database=helpdesk;Trusted_Connection=true");
            System.Data.SqlClient.SqlConnection sqlconn2 = new System.Data.SqlClient.SqlConnection("Server=lk-dit-ad.project.local;database=helpdesk;Trusted_Connection=true");

            string sqlQueryforCheck = "";
            string sqlQueryForInsert = "";

            // Read the file and display it line by line.
            System.IO.StreamReader file = new System.IO.StreamReader("C:\\userInfo.txt");

            try
            {
                sqlconn1.Open();
                sqlconn2.Open();
                Console.WriteLine("Connected to SQL");
            }

            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                //Console.ReadLine();
            }

            //read every single line of the file
            while ((line = file.ReadLine()) != null)
            {

                line = line.ToString();
                Console.WriteLine("**** " + line);
                lineArr = line.Split(','); //Split each line by coma

                // if [4] or [5] are blank throw ERROR. do not add to SQL
                //************************************************************
                if (lineArr[4] == "" || lineArr[5] == "")
                {
                    Console.WriteLine("ERROR: Missing user data. Skipping. User:" + lineArr[0].ToString());
                    return;
                }
                else
                {

                    //pass username from file into SQL query to be checked for existance in helpdesk DB
                    sqlQueryforCheck = "SELECT ID from users WHERE username =" + "'" + lineArr[3].ToString() + "'";
                    System.Data.SqlClient.SqlCommand scmd = new System.Data.SqlClient.SqlCommand(sqlQueryforCheck, sqlconn1);

                    //SQL that inserts non existent user informaction from file to SQL DB (Parametrised)
                    sqlQueryForInsert = "INSERT INTO users (FNAME,LNAME,USERNAME,EMAIL,PHONE,USER_TYPE) values(@fname,@lname,@username,@email,@phone,@user_type)";
                    System.Data.SqlClient.SqlCommand sqlInsert = new System.Data.SqlClient.SqlCommand(sqlQueryForInsert, sqlconn2);

                    // Create parameters
                    sqlInsert.Parameters.Add("@fname", System.Data.SqlDbType.Text);
                    sqlInsert.Parameters.Add("@lname", System.Data.SqlDbType.Text);
                    sqlInsert.Parameters.Add("@username", System.Data.SqlDbType.Text);
                    sqlInsert.Parameters.Add("@email", System.Data.SqlDbType.Text);
                    sqlInsert.Parameters.Add("@phone", System.Data.SqlDbType.Text);
                    sqlInsert.Parameters.Add("@user_type", System.Data.SqlDbType.TinyInt);
                    //sqlInsert.Parameters.Add("@department", System.Data.SqlDbType.Text);

                    //Assign parameters from from array elements
                    sqlInsert.Parameters["@fname"].Value = lineArr[1].ToString();
                    sqlInsert.Parameters["@lname"].Value = lineArr[2].ToString();
                    sqlInsert.Parameters["@username"].Value = lineArr[3].ToString();
                    sqlInsert.Parameters["@email"].Value = lineArr[4].ToString();
                    sqlInsert.Parameters["@phone"].Value = lineArr[5].ToString();
                    sqlInsert.Parameters["@user_type"].Value = 0;


                    System.Data.SqlClient.SqlDataReader bhread = scmd.ExecuteReader();

                    /* 
                    if bhread has rows, that means username is already in the database so skip it and move on the the next line in the file
                    for another username to be checked against database
                    if found, insert it to database
                    */
                    if (bhread.HasRows)
                    {
                        Console.WriteLine("User alrady DB: " + lineArr[3].ToString());
                    }
                    else
                    {
                        sqlInsert.ExecuteNonQuery();
                        Console.WriteLine("Insert Success");
                    }

                    bhread.Close();

                }
            }
                file.Close();
                sqlconn1.Close(); //connection stays open outside while processing
                sqlconn2.Close();
                file.Close();
                Console.WriteLine("Connection Closed");
                Console.ReadLine();

            
        }
        /*************************************** MODIFIED CODE FROM ABOVE FOR WEB APP OPERATORS *****************************************/



        /*getWebAppOperatorsMembers finction: Finds members of the webappoperators group and send their username to getWAOPUserProperties function to get more details */
        public static void getWebAppOperatorsMembers()
        {
            Console.WriteLine("Reading WebAppOperators group Members");

            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "project.local"); //point to the domain
            GroupPrincipal grp = GroupPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, "WebAppOperators");

            if (grp != null)
            {
                //loop through members of the group
                foreach (Principal p in grp.GetMembers())
                {
                    string uname = String.Format("{0}", p.SamAccountName);
                    Console.WriteLine(uname);
                    getWAOPUserProperties(uname); //pass each member username function that gets more information
                }

                grp.Dispose();
                ctx.Dispose();

            }
            else
            {
                Console.WriteLine("WebAppOperators Windows Group Not Found");
                Console.ReadLine();
            }
        }


        /* This function receives username from getWebAppOperatorsMembers and reads more details
         * all details are concatenated into single line string and pased to savetofile function
         * 
         */
        public static void getWAOPUserProperties(string unamep)
        {

            PrincipalContext cty = new PrincipalContext(ContextType.Domain, "project.local");
            UserPrincipal usr = UserPrincipal.FindByIdentity(cty, unamep);

            string cn = usr.Name;
            string givenName = usr.GivenName;
            string sn = usr.Surname;
            string sAMAccountName = unamep;
            string mail = usr.EmailAddress;
            string telephoneNumber = usr.VoiceTelephoneNumber;
            //string physicalDeliveryOfficeName = ""; // requires DirectoryEntry object

            string fullUserInfo = cn + "," + givenName + "," + sn + "," + sAMAccountName + "," + mail + "," + telephoneNumber;
            Console.WriteLine(fullUserInfo);
            saveWAOPUsersToFile(fullUserInfo, cn);

        }

        /*This function receives full user info and username from getWAOPUserProperties
         * and saves all information in the text file
         */
        public static void saveWAOPUsersToFile(string userInfo, string uname) //uname only used for display purposes
        {

            using (System.IO.StreamWriter file = new System.IO.StreamWriter("C:\\WebAppOperatorsUsers.txt", true))
            {
                file.WriteLine(userInfo);
            }

            Console.WriteLine("User " + uname + " has been saved to file");

        }

        /* SQL reads file and inserts values to SQL users table. Check if they exisit and if not commits insert */
        public static void insertWAOPUsersToSQL()
        {
            string line;
            string[] lineArr = new string[6];

            //SQL Connection Create            
            System.Data.SqlClient.SqlConnection sqlconn1 = new System.Data.SqlClient.SqlConnection("Server=lk-dit-ad.project.local;database=helpdesk;Trusted_Connection=true");
            System.Data.SqlClient.SqlConnection sqlconn2 = new System.Data.SqlClient.SqlConnection("Server=lk-dit-ad.project.local;database=helpdesk;Trusted_Connection=true");

            string sqlQueryforCheck = "";
            string sqlQueryForInsert = "";

            // Read the file and display it line by line.
            System.IO.StreamReader file = new System.IO.StreamReader("C:\\WebAppOperatorsUsers.txt");

            try
            {
                sqlconn1.Open();
                sqlconn2.Open();
                Console.WriteLine("Connected to SQL");
            }

            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            //read every single line of the file
            while ((line = file.ReadLine()) != null)
            {

                line = line.ToString();
                Console.WriteLine("**** " + line);
                lineArr = line.Split(','); //Split each line by coma

                //pass username from file into SQL query to be checked for existance in helpdesk DB
                sqlQueryforCheck = "SELECT ID from users WHERE username =" + "'" + lineArr[3].ToString() + "'";
                System.Data.SqlClient.SqlCommand scmd = new System.Data.SqlClient.SqlCommand(sqlQueryforCheck, sqlconn1);

                //SQL that inserts non existent user informaction from file to SQL DB (Parametrised)
                sqlQueryForInsert = "INSERT INTO users (FNAME,LNAME,USERNAME,EMAIL,PHONE,USER_TYPE) values(@fname,@lname,@username,@email,@phone,@user_type)";
                System.Data.SqlClient.SqlCommand sqlInsert = new System.Data.SqlClient.SqlCommand(sqlQueryForInsert, sqlconn2);

                // Create parameters
                sqlInsert.Parameters.Add("@fname", System.Data.SqlDbType.Text);
                sqlInsert.Parameters.Add("@lname", System.Data.SqlDbType.Text);
                sqlInsert.Parameters.Add("@username", System.Data.SqlDbType.Text);
                sqlInsert.Parameters.Add("@email", System.Data.SqlDbType.Text);
                sqlInsert.Parameters.Add("@phone", System.Data.SqlDbType.Text);
                sqlInsert.Parameters.Add("@user_type", System.Data.SqlDbType.TinyInt);
                //sqlInsert.Parameters.Add("@department", System.Data.SqlDbType.Text);

                //Assign parameters from from array elements
                sqlInsert.Parameters["@fname"].Value = lineArr[1].ToString();
                sqlInsert.Parameters["@lname"].Value = lineArr[2].ToString();
                sqlInsert.Parameters["@username"].Value = lineArr[3].ToString();
                sqlInsert.Parameters["@email"].Value = lineArr[4].ToString();
                sqlInsert.Parameters["@phone"].Value = lineArr[5].ToString();
                sqlInsert.Parameters["@user_type"].Value = 1;


                System.Data.SqlClient.SqlDataReader bhread = scmd.ExecuteReader();

                /* 
                if bhrad has rows, that means username is already in the database so skip it and move on the the next line in the file
                for another username to be checked against database
                if found, insert it to database
                */
                if (bhread.HasRows)
                {
                    Console.WriteLine("User alrady DB: " + lineArr[3].ToString());
                }
                else
                {
                    sqlInsert.ExecuteNonQuery();
                    Console.WriteLine("Insert Success");
                }

                bhread.Close();

            }
            file.Close();
            sqlconn1.Close(); //connection stays open outside while processing
            sqlconn2.Close();
            file.Close();
            Console.WriteLine("Connection Closed");
            Console.ReadLine();

        }


    }
}

