# Auto Generate Word Document with Excel Chart and Email Report
This tutorial will show you how to use C# to generate word document or PDF file with some Excel charts in it, and email this report to who is concerned. The motivation of this tutorial is because I don't want to use some third party library which cost or maybe free but with watermark on your report. 

# 1. What is the Goal <br />
Let's take a look on what kind of document (.docx or .pdf) we are going to generate using C# Console Application.<br />
<p align="center">
<img src="/image/report_img.JPG" height="40%" width="40%"> 
</p>  
Apparently, there is title, some contents, some pictures, and some charts in the report. That is what we are going to achieve using program.

# 2. The Idea in behind <br />
We will use C# program to use Microsoft Office 2010 Word and Excel to generate report for us, so you will need Office 2010 installed in your PC. 
<p align="center">
<img src="/image/idea_process.JPG" height="90%" width="90%"> 
</p>  
<br />
You will need the following:
* Visual Studio 2017
* Office 2010 Word, Excel
* SQL database (optional)
* Exchange Server Account (optional)

# 3. Prepare Word Template
Reference 1: https://vivekcek.wordpress.com/tag/interop-word-template/   (Create Word Template)  <br />
Reference 2: https://stackoverflow.com/questions/32105775/how-to-add-an-image-at-a-specific-location-in-msword-using-c  (Place inlineshape at specific location)  <br />

This is very crucial step, because this will save your life. You are not required to prepare one template, but just using pure code to create a Word document. However, you will end up finding few resources online teaching you how to use program to place your content at the location you wish or may be achieve the effect you wanted. Using template will be more reasonable.
Let's take a look what is a template looks like
<p align="center">
<img src="/image/template_img.JPG" height="40%" width="40%"> 
</p>  
It looks almost the same as the report which we are going to generate, right? Only some part with some strange symbol such as ```<<    >>```.

### 3.1 Steps to Create Template
1. Open a new Word document in Office 2010.
2. Put some content like you usually do in Word, remember that, at this step, all the content will remain the same without any change.
3. This step is crucial, we will put some special mark in this document which our program will recognize so that we can put some dynamic content at the place we marked. At ```Insert``` tab, select ```Quick Parts```, select ```Field```, select ```MergeField``` in the Field names. You can then name it whatever you like in the Field name.
4. You can then add some permenant picture in the content.
5. Reserve a location for our dynamic Excel chart. Choose a location, click on ```Insert``` tab, select ```Bookmark```, put a name you like, click ```Add```.
<p align="center">
<img src="/image/template_step.gif" height="90%" width="90%"> 
</p> 

# 4. Create C# Project
Create a C# Console Application Solution in Visual Studio 2017, we name our solution ```motor_alarm```. The IDE will generate some files for you including ```Program.cs``` which we are going to put most of our main code in this file. We then add ```New Item``` C# class with name ```SQL_str.cs``` where we are going to put some SQL connecting and query string. This will make our solution more organize and clean. 
<p align="center">
<img src="/image/solution_explorer_img.JPG" height="20%" width="20%"> 
</p>  

Our code will do the following: 
1. Check if Remote SQL Database is alive on the LAN (Local Area Network). If good, Connect to SQL database, do some quering, return dataset.
2. Create Word Document using template which we've created. Draw Excel chart according to the dataset from SQL
3. Convert .docx file to .pdf file
4. Attach pdf file in email and email it to Exchange Server.

### 4.1 SQL
This section will talk about everthing you need to know when connect to SQL database. The following code will all be in ```SQL_str.cs``` under ```class SQL_str```
#### 4.1.1 Check if SQL on the LAN
```c#
using System.Net;
using System.Net.NetworkInformation;
using System.Diagnostics;

//...... The function will be in the class
public bool PingHost(string nameOrAddress)
{
            bool pingable = false;
            Ping pinger = null;
            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send(nameOrAddress);
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException)
            {
                // Discard PingExceptions and return false;
            }
            finally
            {
                if (pinger != null)
                {
                    pinger.Dispose();
                }
            }
            return pingable;
}
//......
```
#### 4.1.2 Connect to SQL and Do Query
The following is just a function.
```c#
using System.Data.SqlClient;
using System.Data;
//......The function will be in the class
public DataSet Query_database(string conn_str, string query_str)
{
            DataSet ds = new DataSet();
            if (ds.Tables.Count > 0)
            {
                ds.Tables[0].Columns.Clear();
                ds.Tables[0].Clear();
                ds.Reset();  //I forgot to put this and cause the "Object reference not set to an instance of an object"
                //ds.Tables[0].Rows.Clear();
            }
            //ds.Tables[0].Rows.Clear();
            using (SqlConnection Myconn2 = new SqlConnection(conn_str))
            using (SqlCommand Mycomm2 = new SqlCommand(query_str, Myconn2))
            {
                try
                {
                    Myconn2.Open();
                    SqlDataAdapter MyAd2 = new SqlDataAdapter();
                    MyAd2.SelectCommand = Mycomm2;
                    //DataTable dTable2 = new DataTable();
                    //MyAd2.Fill(dTable2);
                    MyAd2.Fill(ds);
                    Myconn2.Close();
                }
                catch (Exception ex)
                {

                }
            }
            return ds;
}        
```
The following code will show you the connecting string required to be put into the function above
```c# 
internal string connection_database = "Data Source = 172.16.246.78; Initial Catalog = SQL_Database; Persist Security Info=True;User ID = sa; Password = some_pwd; Connection Timeout=1";

internal string select_std = @"SELECT STDEV(A.TagValue) as 'std', AVG(A.TagValue) as 'avg' FROM(
                                        SELECT TOP 288 [TagName]
                                              ,[TagValue]
                                              ,[TagTime]
                                          FROM [DCS].[dbo].[DataSL910]
                                          where 
                                          1=1
                                          AND TagName like 'AMA021-2%'
                                          AND TagValue >= '10'
                                          AND TagTime < replace(convert(varchar, DATEADD(MINUTE,-5,GETDATE()),111),'/','') + replace(convert(varchar, DATEADD(MINUTE,-5,GETDATE()),108),':','')
                                          order by TagTime desc
                                     ) A";
                                     
internal string select_records = @"SELECT TOP 40 [TagName]
                                            ,[TagValue]
                                            ,[TagTime]	  
                                            , SUBSTRING([TagTime], 1, 4) + '-' + SUBSTRING([TagTime], 5, 2) + '-' + SUBSTRING([TagTime], 7, 2) + ' ' 
                                            + SUBSTRING([TagTime], 9, 2) + ':' + SUBSTRING([TagTime], 11, 2) + ':'+ SUBSTRING([TagTime], 13, 2) as 'TagDate'
                                            , SUBSTRING([TagTime], 9, 2) + ':' + SUBSTRING([TagTime], 11, 2) + ':'+ SUBSTRING([TagTime], 13, 2) as 'TagTimeHour'
                                            FROM [DCS].[dbo].[DataSL910]
                                            where 1=1
                                            AND TagName like 'AMA021-2%'
                                            ORDER BY TagTime Desc";                                     
```
The following code will be inside ```Program.cs``` under ```class Program```. Which will call SQL connect and query function every 5 min
```c++
static SQL_str sql_str = new SQL_str();
static double current_std = 0.0;
static double current_avg = 0.0;
static double current_val = 0.0;
        
static void Main(string[] args)
{
            Timer _timer = new Timer(TimerCallback, null, 0, 300 * 1000); //300 means 300 seconds --> 5 min
            Console.ReadKey();          
}

private static void TimerCallback(Object o)
{
            Query_database_STD(sql_str.connection_database, sql_str.select_std);
            CreateDocument_template(Query_database_Record(sql_str.connection_sl910, sql_str.select_records));  // We will talk about this later
            Console.WriteLine("VAL: " + current_val + "  STD: " + current_std.ToString() + "  AVG: " + current_avg.ToString());
            ConvertDocument_PDF(docx_filename, pdf_filename); // We will talk about this later
            if (current_val - current_avg > 3.0 * current_std)
            {               
                Console.WriteLine("Abnormal");               
                Send_Mail(rec, pdf_filename);  // We will talk about this later
            }
            else
            {
                Console.WriteLine("Normal");
            }
            Console.WriteLine("In TimerCallback: " + DateTime.Now);
}

```



