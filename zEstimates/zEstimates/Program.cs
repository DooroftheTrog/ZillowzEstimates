using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Xml;

namespace ConsoleApplication2
{
    class zEstimatesAccess
    {
        static void Main(string[] args)
        {
            //create a session for the execution of the program
            /*
            SDKSession session = new SDKSession();
            SDKApplication app = null;

            //login into session
            Console.WriteLine("Please enter your username:");
            string user = Console.ReadLine();
            Console.WriteLine("Please enter your password:");
            string pass = Console.ReadLine();
            session.Login(user, pass, "Server");
            session.Authorize("Atlantic Bay Mortgage", "36040", "773329768226018262");
            */
            //Access Excel Data
            Excel.Application x = new Excel.Application();
            if (x == null)
            {
                Console.WriteLine("This bad boy Empty YO!");
                return;
            }
            else
            {
                x.Visible = true;
                Console.WriteLine("Enter The file path of the zEstimate File :");
                string path =  Console.ReadLine();
                Excel.Workbook wb = x.Workbooks.Open(path);
                Excel.Worksheet ws = wb.ActiveSheet;
                Excel.Range r = ws.UsedRange;
                int rows = r.Rows.Count;
                int cols = r.Columns.Count;
                String[,] searchInfo = new String[rows, cols];

                //list of zillow ids used to make calls
                String[] ids = { "X1-ZWz1ey2snjmlfv_4aijl", "X1-ZWz1ey2wllupe3_4dboj", "X1-ZWz1a2fkcts83v_4eq90", "X1-ZWz1ey30jo2tcb_4g4th", "X1-ZWz1a2fgerk45n_4hjdy", "X1-ZWz1ey38fsj18r_4lr3d", "X1-ZWz1a2f8in3w97_4n5nu", "X1-ZWz1ey3cdur56z_4ok8b", "X1-ZWz1a2f4kkvsaz_4pyss", "X1-ZWz1ey3gbwz957_4rdd9", "X1-ZWz1eycznaksuj_1k6qj" };

                if (ws == null)
                {
                    Console.WriteLine("Work Sheet could not be created YO!");
                    return;
                }
                else
                {
                    //transfer the data from the excel file into the 2d array searchInfo                
                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= cols; j++)
                        {
                            string str = r.Cells[i, j].Text;
                            searchInfo[i - 1, j - 1] = str;

                            Console.Write(str + " ");
                        }
                        Console.WriteLine();
                    }
                }
                //Create Web Connection and Make web Calls
                //selects a new key when 1000 calls have been made
                for (int numID = 0; numID <= ids.Length; numID++)//change numID back to 0 tomorrow && < !!!!!!!!!!
                {
                    //iterate over the columns of search info to get the address, city, state, and zip
                    int numIterations = 1000 * numID; //use this as an additive to i on access of data so that when items are over 1000 it still iterates the correct amount  
                    if (rows == 0) { break; } //check to see if the amount of rows is 0 break out of the loop                
                    int numofCalls;
                    numofCalls = (rows > 1000) ? 1000 + numIterations : rows + numIterations;
                    for (int i = numIterations; i < numofCalls; i++)
                    {
                        string address = formatAddress(searchInfo[i, 1]);
                        string citystatezip = "";
                        citystatezip = formatCSZ(searchInfo[i, 2]);

                        //get the XML File for the ZPID based on address and city
                        string url = "http://www.zillow.com/webservice/GetSearchResults.htm?zws-id=" + ids[numID] + "&address=" + address + "&citystatezip=" + citystatezip;

                        //gets the ZPID from the XML File
                        XmlNodeList elements = getDoc(url).GetElementsByTagName("zpid");

                        for (int k = 0; k < elements.Count; k++)
                        {

                            //gets the XML File from zillow for the ZEstimate
                            url = "http://www.zillow.com/webservice/GetZestimate.htm?zws-id=" + ids[numID] + "&zpid=" + elements[k].InnerText.ToString();
                            //gets the zEstimate from the XML File
                            XmlNodeList zEstimate = getDoc(url).GetElementsByTagName("amount");
                            for (int j = 0; j < zEstimate.Count; j++)
                            {
                                Console.WriteLine(zEstimate[j].InnerText.ToString());
                                //store z estimate in byte using sdk
                                //input zEstimate[j].InnerText.ToString() into the loan where the loan id equals(searchInfo[i+1,0])
                                // SDKFile writeTo = null;
                                //app = session.GetApplication();

                                ws.Cells[i + 2, "D"] = zEstimate[j].InnerText.ToString();
                                //writeTo = app.OpenFile(searchInfo[i, 0], false);
                                Console.WriteLine(i);
                                //writeTo.SetField("ExtendedFields.zEstimate", zEstimate[j].InnerText.ToString());
                                //writeTo.Save();
                                //app.CloseFile(writeTo);
                            }
                        }
                    }
                    rows -= 1000; //subtract 1000 from rows so that you end with the remainder of rows, example the last 456 out of 1456
                }
                wb.Save();
                Console.ReadLine();
            }
        }
        public static XmlDocument getDoc(string url)
        {
            HttpWebRequest xmlRequest = WebRequest.Create(url) as HttpWebRequest;
            HttpWebResponse xmlresponse = xmlRequest.GetResponse() as HttpWebResponse;
            XmlDocument response = new XmlDocument();
            response.Load(xmlresponse.GetResponseStream());
            return response;
        }
        //formats the city, state, and zip so that they can be used in the api call
        public static string formatCSZ(String citystatezip)
        {

            if (citystatezip != "" && citystatezip != null)
            {
                citystatezip.Replace(" ", "+");
                citystatezip.Replace(",", "%2C");
            }
            return citystatezip;
        }
        //formats address's so that they can be used in an api call
        public static string formatAddress(string Address)
        {
            Address = Address.Replace(" ", "+");
            Address = Address.Replace(",", "%2C");
            return Address;
        }
    }
}