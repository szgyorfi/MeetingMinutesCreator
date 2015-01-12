# Meeting Minutes Creator
MS Outlook script to create meeting minutes .docx file based on Appointment data. Basically it will create a word document where you can write down results / assumptions / tasks issued on the meeting and distribute this .docx file to the meeting attendees.
<h3>What it does</h3>
Outputs a file containing the following
- - -
<div>

  <h4>Meeting Minutes</h4><br>

  <p>

    <b>Subject</b>: <i>[Your subject field content]</i><br>
    <b>Importance</b>: <i>[Represented by a number from 0 - high to 2 low]</i><br>
    <b>Location</b>: <i>[Location]</i><br>
    <b>Start</b>: <i>[Start In StartTimeZone]</i><br>
    <b>Organizer</b>: <i>[Organizer]</i><br>
    <b>Required</b>: <i>[Required Attendees]</i><br>
    <b>Optional</b>: <i>[Optional Attendees]</i><br><br>
      <b><i>Present:</b></i>: <i>[Your actual list has to be filled in]</i><br>
    <br>
    <h4>Results</h4>
    ...
  </p>

</div>
- - - 

<h3>How to use</h3>
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
