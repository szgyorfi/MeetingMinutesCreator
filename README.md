# Meeting Minutes Creator
MS Outlook script to create meeting minutes .docx file based on Appointment data.

1. The code assumes that you have Miscrosoft Office installed (or at least Outlook and Word). <br>
Tests performed only in MS Office 2010.
2. You'll have to add the **Microsoft Word 14.0 Object Library** reference to your Outlook Project. <br>
You can do this by going **Tools** > **References** and search the list for this object library. Simply tick the mark and click **OK**.
3. The file **CreateMeetingMinutes.bas** must be added to your Outlooks VBA project. <br>
You can find instruction on how to import the file into outlook [here](http://www.outlookcode.com/article.aspx?id=28) at the file Import/Export section.
4. After import maybe you want to add a Ribbon Button to the Appointment / Meeting menu. <br>
Instructions can be found [here](http://www.howto-outlook.com/howto/macrobutton.htm#ribbon)
5. If you setted all this up, open an Appointment / Meeting and Run the macro <br>
(if you did not have created a Ribbon Button you can simply access available macros by pressing the **Alt+F8** key combination - choose *createMinutes* macro and press **Run**.)

Good luck!
