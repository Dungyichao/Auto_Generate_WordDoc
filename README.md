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

# 4. Create C# Project
Create a C# Console Application Solution in Visual Studio 2017, we name our solution ```motor_alarm```. The IDE will generate some files for you including ```Program.cs``` which we are going to put most of our main code in this file. We then add ```New Item``` C# class with name ```SQL_str.cs``` where we are going to put some SQL connecting and query string. This will make our solution more organize and clean. 
