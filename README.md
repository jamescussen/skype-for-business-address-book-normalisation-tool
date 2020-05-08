Skype for Business Address Book Normalisation Tool
==================================================

            


![Image](https://github.com/jamescussen/skype-for-business-address-book-normalisation-tool/raw/master/addressbooknormscreenshot1.00.png)


 


**Tool Features:**




  *  Import Existing “Company_Phone_Number_Normalization_Rules.txt” files into the system.

  *  Add/Edit address book rules to the system. If the rule you are setting has a name that matches an existing rule, then the existing rule will be edited. If the rule’s name does not match an existing rule then it will be added as a new rule to the list.

  *  Delete rules from the system. 
  *  Create new Site based Address Book Normalisation Rules policies. 
  *  Change the priority of rules. 
  *  Custom written rule testing code for testing pattern and translation matches as well as the resultant number.

  *  Export rules back into a “Company_Phone_Number_Normalization_Rules.txt” file format.

  *  Test the rules! Skype for Business currently (at the time of writing this) doesn’t have Address Book Normalisation testing capabilities. So I wrote a custom testing engine into the tool providing this feature. By entering a number into the Test textbox
 and pressing the Test Number button, the tool will highlight all of the rules that match in the currently selected Global/Site level Policy patterns in **blue**. The rule that has the highest priority and
 matches the tested number will be highlighted in **red**. The pattern and translation of the highest priority match (the one highlighted in red) will be used to do the translation on the Test Number and
 the resultant translated number will be displayed by the Test Result label. 

 


**Version 1.01 Update (13/10/2015):**



  *  Added warning message on the Remove policy button to save you from yourself :)

  *  Removed second .txt from the export name. 


**Version 1.02 Update (20/1/2015):**



  *  Script now doesn't strip ';' char before applying regex. (Thanks Daniel Appleby for reporting)

  *  Updated Code Signing Signature (25/5/2016) 

**Version 1.0.3 Update (28/11/2019):**


  *  Updated the Test Number capability to be much more accurate. 
  *  Moved around some UI elements to make things a bit clearer. 
  *  Added some more error checking. 
  *  Updated Icon. 

**All information on this tool can be found here: **[http://www.myteamslab.com/2015/07/skype-for-business-address-book.html](http://www.myteamslab.com/2015/07/skype-for-business-address-book.html)


 






        
    
