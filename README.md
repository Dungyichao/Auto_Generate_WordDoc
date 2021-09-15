# Auto Generate Word Document with Excel Chart and Email Report
This tutorial will show you how to use C# to generate word document or PDF file with some Excel charts in it, and email this report to who is concerned. The motivation of this tutorial is because I don't want to use some third party library which cost or maybe free but with watermark on your report. 

# 1. What is the Goal <br />
Let's take a look on what kind of document (.docx or .pdf) we are going to generate using C# Console Application.<br />
<p align="center">
<img src="/image/report_img.JPG" height="40%" width="40%"> 
</p>  
Apperntly, there is title, some contents, some pictures, and some charts in the report. That is what we are going to achieve using program.

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

