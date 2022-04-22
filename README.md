# Auto Generate Word Document with Excel Chart and Email Report
This tutorial will show you how to use C# to generate word document or PDF file with some Excel charts in it, and email this report to who is concerned. The motivation of this tutorial is because I don't want to use some third party library which cost or maybe free but with watermark on your report. 

I spend so much time on researching this topic online, but only to find some fragment of information which is not easy for beginner to understand. There is no any detail tutorial on related topic which teach you to do such complete document. I believe a basic document would include some image, dynamic charts, some text content. Fortunately, my tutorial includes all of them. I hope this tutorial can save much of your time on the tedious work.  

1. [What is the Goal](https://github.com/Dungyichao/Auto_Generate_WordDoc#1-what-is-the-goal-)
2. [The Idea in behind](https://github.com/Dungyichao/Auto_Generate_WordDoc#2-the-idea-in-behind-)
3. [Prepare Word Template](https://github.com/Dungyichao/Auto_Generate_WordDoc#3-prepare-word-template)
    * 3.1 [Steps to Create Template](https://github.com/Dungyichao/Auto_Generate_WordDoc#31-steps-to-create-template)
4. [Create C# Project](https://github.com/Dungyichao/Auto_Generate_WordDoc#4-create-c-project)
    * 4.1 [SQL](https://github.com/Dungyichao/Auto_Generate_WordDoc#41-sql)
    * 4.2 [Generate Word Document and Excel Chart](https://github.com/Dungyichao/Auto_Generate_WordDoc#42-generate-word-document-and-excel-chart)
    * 4.3 [Email Out the Report](https://github.com/Dungyichao/Auto_Generate_WordDoc#43-email-out-the-report)
5. Interact with Web Page Data

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
You will need the following: <br />

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

You also need to add reference ```Microsoft.Office.Interop.Excel``` and ```Microsoft.Office.Interop.Word``` to this project. We use Version 12.0.0.0.
<p align="center">
<img src="/image/vs2017_ref.JPG" height="50%" width="50%"> 
</p> 

Or

<p align="center">
<img src="/image/vs2017_ref_COM.JPG" height="50%" width="50%"> 
</p> 

Our code will do the following: 
1. Check if Remote SQL Database is alive on the LAN (Local Area Network). If good, Connect to SQL database, do some quering, return dataset.
2. Create Word Document using template which we've created. Draw Excel chart according to the dataset from SQL. Optional step is converting .docx file to .pdf file
3. Attach pdf file in email and email it to Exchange Server.

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
                                          FROM [SQL_Database].[dbo].[DataSL910]
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
                                            FROM [SQL_Database].[dbo].[DataSL910]
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
Until this stage, we will have SQL records in hand and we will start generate report and Excel chart in next section

### 4.2 Generate Word Document and Excel Chart
In this section, we will talk about how to use program to call Office 2010 library. 
#### 4.2.1 Create Word
```C#
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

//The following code are inside ```class Program```
static string pdf_filename = @"c:\motor_alarm\template\temp1.PDF";
static string docx_filename = @"c:\motor_alarm\template\temp1.docx";
static object oTemplatePath = @"C:\motor_alarm\template\Word1.dotx";
static object filename_pdf = @"c:\motor_alarm\template\temp1.PDF";
static object filename_doc = @"c:\motor_alarm\template\temp1.docx";

private static void CreateDocument_template(DataSet ds)
{
            try
            {
                String QR_save_fileName = Path.GetTempFileName();
                var result = new Bitmap();
                result.Save(QR_save_fileName);
                //Create an instance for word app  
                Word.Application winword = new Word.Application();

                //Set status for word application is to be visible or not.  
                winword.Visible = false;

                //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;
                //object oTemplatePath = @"C:\Users\admin\Desktop\motor_alarm\template\Word1.dotx";
                //object cTemplatePath = @"C:\Users\admin\Desktop\motor_alarm\template\Chart1.crtx";

                //Create a new document  
                Word.Document document = winword.Documents.Add(ref oTemplatePath, ref missing, ref missing, ref missing);

                foreach (Word.Field myMergeField in document.Fields)
                {
                    Word.Range rngFieldCode = myMergeField.Code;

                    String fieldText = rngFieldCode.Text;

                    // ONLY GETTING THE MAILMERGE FIELDS

                    if (fieldText.StartsWith(" MERGEFIELD"))
                    {

                        // THE TEXT COMES IN THE FORMAT OF

                        // MERGEFIELD  MyFieldName  \\* MERGEFORMAT

                        // THIS HAS TO BE EDITED TO GET ONLY THE FIELDNAME "MyFieldName"

                        Int32 endMerge = fieldText.IndexOf("\\");

                        Int32 fieldNameLength = fieldText.Length - endMerge;

                        String fieldName = fieldText.Substring(11, endMerge - 11);

                        // GIVES THE FIELDNAMES AS THE USER HAD ENTERED IN .dot FILE

                        fieldName = fieldName.Trim();

                        // **** FIELD REPLACEMENT IMPLEMENTATION GOES HERE ****//

                        // THE PROGRAMMER CAN HAVE HIS OWN IMPLEMENTATIONS HERE

                        if (fieldName == "datetime")
                        {
                            myMergeField.Select();
                            winword.Selection.TypeText(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"));
                        }
                        else if (fieldName == "current_val")
                        {
                            myMergeField.Select();
                            winword.Selection.TypeText(current_val.ToString());
                        }
                        else if (fieldName == "current_avg")
                        {
                            myMergeField.Select();
                            winword.Selection.TypeText(Math.Round(current_avg, 4).ToString());
                        }
                        else if (fieldName == "current_std")
                        {
                            myMergeField.Select();
                            winword.Selection.TypeText(Math.Round(current_std, 4).ToString());
                        }
                        else if (fieldName == "3sigma")
                        {
                            myMergeField.Select();
                            winword.Selection.TypeText((3.0 * Math.Round(current_std, 4)).ToString());
                        }
                        else if (fieldName == "QR_CODE")
                        {
                           myMergeField.Select();
                           winword.Selection.InlineShapes.AddPicture(QR_save_fileName, false, true);
                        }

                    }

                }

                object oClassType = "Excel.Chart.8";
                object oEndOfDoc = "\\endofdoc";

                //Word.InlineShape wrdInlineShape = document.InlineShapes.AddOLEObject(oClassType);

                //Place inlineshape at specific location:https://stackoverflow.com/questions/32105775/how-to-add-an-image-at-a-specific-location-in-msword-using-c
                //open your word document > add a bookmark named: "PicHere"
                //https://stackoverflow.com/questions/8483471/how-to-change-the-size-of-a-picture-after-inserting-it-into-a-word-document
                 
                Word.InlineShape wrdInlineShape = document.Bookmarks["PicHere"].Range.InlineShapes.AddOLEObject(oClassType);


                Excel.Workbook obook = (Excel.Workbook)wrdInlineShape.OLEFormat.Object;
                Excel.Worksheet sheet = (Excel.Worksheet)obook.Worksheets["Sheet1"];
                if (wrdInlineShape.OLEFormat.ProgID == "Excel.Chart.8")
                {
                    // Word doesn't keep all of its embedded objects in the running state all the time.
                    // In order to access the interface you first have to ensure the object is in the running state,
                    // ie: OLEFormat.Activate() (or something)
                    object verb = Word.WdOLEVerb.wdOLEVerbHide;
                    wrdInlineShape.OLEFormat.DoVerb(ref verb);
                    //Random rn = new Random();

                    Excel.Range temp_range = sheet.get_Range("A1", "C40");

                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        ((Excel.Range)temp_range.Cells[ds.Tables[0].Rows.Count - i, 1]).Value = ds.Tables[0].Rows[i]["TagTimeHour"];
                        ((Excel.Range)temp_range.Cells[ds.Tables[0].Rows.Count - i, 2]).Value = ds.Tables[0].Rows[i]["TagValue"];
                        ((Excel.Range)temp_range.Cells[ds.Tables[0].Rows.Count - i, 3]).Value = current_avg;
                    }
                    wrdInlineShape.Width = 500;
                    wrdInlineShape.Height = 250;

                    obook.ActiveChart.ChartType = Excel.XlChartType.xlLineMarkers;
                    obook.ActiveChart.HasTitle = true;
                    obook.ActiveChart.ChartTitle.Text = "AMA021-2 Records in past 3hr";
                    obook.ActiveChart.SetSourceData(temp_range.get_Range("B1", "C40"), Excel.XlRowCol.xlColumns);


                    Excel.Series series1 = (Excel.Series)obook.ActiveChart.SeriesCollection(1);
                    series1.Name = "AMA021-2 Amp";
                    Excel.Series series2 = (Excel.Series)obook.ActiveChart.SeriesCollection(2);
                    series2.Name = "24Hr Average";
                    series2.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;

                    //series1.HasDataLabels = true;
                    series1.XValues = sheet.get_Range("A1", "A40").Value;


                    //Axis title: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.tools.excel.chart.axes?view=vsto-2017
                    Excel.Axis axis = (Excel.Axis)obook.ActiveChart.Axes(
                                        Excel.XlAxisType.xlValue,
                                        Excel.XlAxisGroup.xlPrimary);
                    axis.HasTitle = true;
                    axis.AxisTitle.Text = "Amp (A)";

                    //Rotate axis label: https://stackoverflow.com/questions/16275979/rotate-x-axis-in-excel-chart-c-sharp
                    //Axis interface: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.axis?view=excel-pia
                    //TickLabels Interface: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.ticklabels?view=excel-pia
                    //https://social.msdn.microsoft.com/Forums/en-US/e39c28b3-6e9e-43d8-ab45-c1789a574f31/excel-2010-chart-line-how-to-format-aix-x-to-daymonth-hourminute?forum=exceldev
                    obook.ActiveChart.Axes(Excel.XlAxisType.xlCategory).TickLabels.NumberFormat = "HH:mm:ss";


                    ////The following marked code can generate table view of data
                    //Word.Range wrdRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    //object oRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    //wrdRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    //sheet.UsedRange.Copy();
                    //document.SetDefaultTableStyle("Light List - Accent 4", false);
                    //for (int i = 0; i < 19; i++)
                    //{
                    //    wrdRng.InsertBreak(Word.WdBreakType.wdLineBreak);
                    //}
                    //wrdRng.PasteExcelTable(true, true, false);
                    wrdInlineShape.ConvertToShape();
                }


                //Save the document  
                //object filename = @"c:\Users\admin\Desktop\temp1.pdf";
                document.SaveAs(ref filename_doc);
                //https://stackoverflow.com/questions/17777545/closing-excel-application-process-in-c-sharp-after-data-access
                //If forgot to obook.Close, Excel application will still be alive in the process
                obook.Close(0); 
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                Console.WriteLine("Document created successfully !");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
}

```

If you try to put the Word template file into the Solution Resources, you can reference it from the Resources. It will be like the following
```C#
using System.IO;
using Word = Microsoft.Office.Interop.Word;

String fileName = Path.GetTempFileName();
File.WriteAllBytes(fileName, Properties.Resources.Customer_Pick_Up_Carrier_Check_In_Sheet_Template);
object temp_file_name = (object)fileName;
Word.Application winword = new Word.Application();
winword.Visible = false;
object missing = System.Reflection.Missing.Value;
Word.Document document = winword.Documents.Add(ref temp_file_name, ref missing, ref missing, ref missing);
```
The process to put your template file into the Solution Resources is in the following:
```TEXT
https://stackoverflow.com/questions/15925801/visual-studio-c-sharp-how-to-add-a-doc-file-as-a-resource
https://stackoverflow.com/questions/33164270/how-to-open-embedded-resource-word-document

Right-click your project and select the "Properties" option.
Then click the "Resources" tab and it will show the dialog for you to add resources in the design time.
The default page is for add String resources, you can select the combobox in the top-right to select the "file" item.
Then click the "Add Resource" button to select the doc file and click OK.
At last, the doc file will show in the blank area. It means that you have added it successfully.
```

Reference 1: https://vivekcek.wordpress.com/tag/interop-word-template/ </br>
Reference 2: https://stackoverflow.com/questions/3684103/how-to-add-office-graph-in-word </br>
Reference 3: https://stackoverflow.com/questions/32105775/how-to-add-an-image-at-a-specific-location-in-msword-using-c (Place inlineshape at specific location) </br>
Reference 4: https://stackoverflow.com/questions/8483471/how-to-change-the-size-of-a-picture-after-inserting-it-into-a-word-document </br>
Reference 5: https://stackoverflow.com/questions/16275979/rotate-x-axis-in-excel-chart-c-sharp  (Rotate axis label) </br>
Reference 6: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.axis?view=excel-pia  (Axis interface) </br>
Reference 7: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.ticklabels?view=excel-pia  (TickLabels Interface) </br>
Reference 8: https://social.msdn.microsoft.com/Forums/en-US/e39c28b3-6e9e-43d8-ab45-c1789a574f31/excel-2010-chart-line-how-to-format-aix-x-to-daymonth-hourminute?forum=exceldev

#### 4.2.2 Convert .docx to .pdf
This step is optional. If you use ```document.SaveAs()``` with .pdf file name, the PDF file will not be able to read normally by regular PDF reader. So we need to first save the document into .docx file, and then convert to .pdf optionally.
```C#
private static void ConvertDocument_PDF(string file_name_docx, string file_name_pdf)
{
            try
            {
                var wordApp = new Word.Application();
                wordApp.Visible = false;
                object readOnly = true;
                var wordDocument = wordApp.Documents.Open(file_name_docx, ref readOnly);

                wordDocument.ExportAsFixedFormat(file_name_pdf, Word.WdExportFormat.wdExportFormatPDF);

                wordDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges,
                                   Word.WdOriginalFormat.wdOriginalDocumentFormat,
                                   false); //Close document

                wordApp.Quit(); //Important: When you forget this Word keeps running in the background
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
}
```
Reference 1: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.documentclass.exportasfixedformat?view=word-pia </br>
Reference 2: https://social.msdn.microsoft.com/Forums/en-US/877d981c-3dba-4724-881d-749225821757/save-word-document-as-pdf?forum=oxmlsdk


#### 4.2.3 Printer Print out PDF file 
```C#
//https://social.msdn.microsoft.com/Forums/en-US/6634c718-67e9-403b-a301-704f9be545c8/print-a-word-document-from-c-using-printdialog?forum=csharplanguage
//https://stackoverflow.com/questions/3197830/is-there-anyway-to-specify-a-printto-printer-when-spawning-a-process
using System;

string file_name_pdf = string.Empty; //file path of the .pdf file
using (PrintDialog pd = new PrintDialog())
{
    if (pd.ShowDialog() == DialogResult.OK)
    {
         //pd.ShowDialog();
         ProcessStartInfo info = new ProcessStartInfo(file_name_pdf);
         info.Verb = "PrintTo";
         info.Arguments = "\"" + pd.PrinterSettings.PrinterName + "\"";
         info.CreateNoWindow = true;
         info.WindowStyle = ProcessWindowStyle.Hidden;
         Process process = Process.Start(info);
         if (process.HasExited == false)
         {
             process.WaitForExit(10000);
          }
          process.Close();
       }             
}
```

### 4.3 Email Out the Report
If you have any access to Microsoft Exchange Server, you can use the following code to email your report to those who you concern. This step is totally optional.
```C#
using System.Net.Mail;
using System.Net;

// attachment is the string of the path of the report. This function is called in section # 4.1.2
public static void Send_Mail(string[] recipients, string attachment)
{
            //Generate SMTP client object with credentials
            var client = new SmtpClient();

            //SMTP client configuration
            client.Host = "ngltexs94.domain.home.usa";//Mail Server
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential("MyAccountName", "123#myPassWord", "OUR_DOMAIN_NAME");//username, passowrd, domain
            client.Timeout = 300000;

            //MailMessage message = new MailMessage();
            //Need to use using so that message will not hold the resource of attachment file
            using (var message = new MailMessage())
            {
                //Generate Mail Message

                message.From = new MailAddress("NoReply@nglt.nptkm.com");//Mail Sender
                message.Subject = "Motor Alarm - " + DateTime.Now.Date.ToString().Split(' ')[0];//Mail Subject
                //message.IsBodyHtml = true;   // May cause sending to mobile carrier not complete
                message.IsBodyHtml = false;
                message.Body = "Current Amp of AMA021 Motor exceed 3 times of standard deviation from last 24 hour";//Mail body
                //mail auto delete after some days
                string expTime = DateTime.Now.AddDays(7).AddHours(4).ToString("dd MMM yyyy HH:mm");
                message.Headers.Add("expiry-date", expTime);


                //Add recipients
                foreach (string recipient in recipients)
                {
                    message.To.Add(recipient);
                }

                //Add attachment (PDF file)
                message.Attachments.Add(new Attachment(attachment));

                //Send mail
                try
                {
                    client.Send(message);
                    Console.WriteLine("Mail Sent");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
            
}

```
If you want to send to SMS by different carrier, you can use the following to concate phone number with the following email domain
```c#
public Dictionary<string, string> ISP = new Dictionary<string, string>()
        {
            {"AT&T", "@txt.att.net"},
            {"Boost", "@myboostmobile.com" },
            {"Nextel", "@messaging.nextel.com" },
            {"Sprint", "@messaging.sprintpcs.com" },
            {"T-Mobile", "@tmomail.net" },
            {"Verizon", "@vtext.com" },
            {"Virgin", "@vmobl.com" }
        };
```

# 5. Interact with Web Page Data
If you want to retrive some data from web browser, but the vendor didn't give you any library to interact with the data, you can follow the following steps to retrieve data. You just need to make sure you can find out how your vendor's web page dealing with URL to retrieve data.

https://user-images.githubusercontent.com/25232370/164585397-ca74b5bc-132b-4d63-a952-858b16f9289b.mp4

The action in above video are performed by program itself (except the moving mouse part). What the program will do Is the follwing
1. Open chrome with URL: https://172.16.218.199/login
2. input password and press enter itself (the username was input before)
3. You will now see the overview page
4. User can now take the synthetic URL into the “Open URL” (in the above video, we put a fixed URL in the program, so the program will automatically open the equipment URL), the program will “Ctrl+A” which select all the data, the program will “ctrl+C” which will copy the data. The program will then read the clipboard data and put into the textbox which our program can then process the data.

Code is in the following
```C#
using System;
using System.Diagnostics;//Process
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading; //Sleep

namespace BayView
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start("chrome.exe", @"https://172.16.218.199/login");
            timer_pwd.Enabled = true;
        }

        private void timer_pwd_Tick(object sender, EventArgs e)
        {
            SendKeys.Send("somePassWordStringandDigit");
            timer_login.Enabled = true;
            timer_pwd.Enabled = false;
        }

        private void timer_login_Tick(object sender, EventArgs e)
        {
            SendKeys.Send("{ENTER}");
            timer_auto_open_url.Enabled = true;
            timer_login.Enabled = false;            
        }

        private void button_URL_Click(object sender, EventArgs e)
        {
            string bay_view_url = textBox1.Text.Trim().ToString();
            open_chrome_URL(bay_view_url);
            timer_select_all.Enabled = true;
        }

        public void open_chrome_URL(string url_str)
        {
            Process.Start("chrome.exe", url_str);
        }

        private void timer_select_all_Tick(object sender, EventArgs e)
        {
            SendKeys.Send("^(a)");
            Thread.Sleep(1000);
            SendKeys.Send("^(c)");
            //if (Clipboard.ContainsText())
            //{
            //    Console.WriteLine(Clipboard.GetText());
            //}
            timer_paste_clipboard.Enabled = true;
            timer_select_all.Enabled = false;
        }

        private void timer_paste_clipboard_Tick(object sender, EventArgs e)
        {
            textBox_result.Text = "";
            textBox_result.Text = Clipboard.GetText().ToString();
            timer_closePage.Enabled = true;
            timer_paste_clipboard.Enabled = false;
        }

        private void timer_closePage_Tick(object sender, EventArgs e)
        {
            SendKeys.Send("^(w)");
            timer_closePage.Enabled = false;
        }

        private void timer_auto_open_url_Tick(object sender, EventArgs e)
        {
            open_chrome_URL(@"https://172.16.218.199/webChart/query/data/history$3a$2f$2fEquipment$2420ID$2fN1_TESLAGOOGLEC0001$2f$2fNANYA_SC0001$2fN1_TESLAGOOGLEC0001_BO_POS$3fstart$3d2022$2d04$2d21T13$3a19$3a30$2e015$2d04$3a00");
            timer_select_all.Enabled = true;
            timer_auto_open_url.Enabled = false;
        }
    }
}

```


