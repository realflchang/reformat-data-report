<h1>Reformat data report.xlsm - Prerequisites and Instructions</h1>
This macro is to convert an unsorted, unformatted Excel data report into a sorted, formatted Excel data report. Useful when the software used to generate the data report is older or incapable to provide formatting needs.<br />

<!---
<h2>Video Demonstration</h2>

- ### [YouTube: How To Install osTicket with Prerequisites](https://www.youtube.com) -->

<h2>Environments and Technologies Used</h2>

- Microsoft Excel
- Visual Basic / Macros

<h2>Operating Systems Used </h2>

- Windows 10</b> (21H2)

<h2>List of Prerequisites</h2>

- Computer running Windows 10. Possible to run in Mac but I have not tested it.
- Excel data report file – contains the unsorted, unformatted data (“sample initial.xls”)
- Macros in Excel are enabled

<h2>Steps</h2>

<p>1.	Create a folder in Windows, My Documents folder to store Excel macros. Ie. My Documents\Macros</p>

<p>2.	Download the Excel macro file (“reformat data report.xlsm”) into that folder</p>

<img src="https://github.com/user-attachments/assets/fddda079-ecf5-44ce-80e0-9c989ee752fb" alt="Download the reformat data report.xlsm file into My Documents Macros folder" />

<p>3.	Next, we want to allow Excel to be able to run the Macro. There are 2 ways:</p>

  *	One way is to click “Enable Content” when this .xlsm file is opened. Problem with this is this will have to be done every time the macro file is opened.
<img src="https://github.com/user-attachments/assets/be5559b0-f82c-4645-9faf-e12d8fe7043d" alt="Macros have been disabled Security Warning" />

  *	Another more permanent way is to allow Excel to trust the location of the .xlsm file. 
    -	To do that, go to File, then click “Options” at the bottom of the menu
<img src="https://github.com/user-attachments/assets/8c7df4c9-9649-4f5a-8de4-fe0b34f8ce2e" alt="Menu File Options" />

    - Click on Trust Center, Trust Center Settings…
<img src="https://github.com/user-attachments/assets/33e1aa94-9133-4b17-bf6a-d6b60a9fb7d1" alt="Trust Center Settings" />

    - Trusted Locations, Add new location…
<img src="https://github.com/user-attachments/assets/c190314e-e4ef-43fe-a41e-6074e4771f91" alt="Add new location" />

    - Browse to the Macros folder that was just created and click OK:
<img src="https://github.com/user-attachments/assets/79751c53-4235-4f35-b18a-997a204e6755" alt="Browse to Macros folder" />

    - Added new location:
<img src="https://github.com/user-attachments/assets/b0ce28c0-f958-4150-b9c7-ad1c09c3292e" alt="Added new location" />

    - Next time we open the .xlsm file in Macros folder, there will not be a tooltip message from Excel 

<p>4.	Next, open the unsorted, unformatted data file. In my case it is “sample initial.xls”.  This is an export from my company's internal, older software:</p>
<img src="https://github.com/user-attachments/assets/bbbb984d-e53e-4ce9-b3eb-3d11490efe46" alt="Open the sample" />

<p>5.	To run the macro, go to View, Macros, View Macros…:</p>
<img src="https://github.com/user-attachments/assets/b5d73dc2-a629-4f7f-93fc-cc2e3bbcd77f" alt="View Macros" />

<p>6.	Click on the Macro from the .xlsm file that includes the name “datareformat”, and click Run:</p>
<img src="https://github.com/user-attachments/assets/0067cb7b-a91d-4764-859e-ab1a03c9dabe" alt="Run Macro" />

  * The Options… button provides an option to assign a key shortcut, so that in the future, instead of browsing to View, Macros, View Macros, we can just do the key shortcut combination to run the same macro.

<img src="https://github.com/user-attachments/assets/de6ca386-1c0d-47b9-98c7-12aece59c964" alt="Run Macro option key shortcut" />

<p>7.	Result of the macro. Result is sorted by the first column and other formats as programmed in the macro:</p>

<img src="https://github.com/user-attachments/assets/b2355517-858a-462d-ae84-abb1f6e9bb5e" alt="Result" />

<p>8.	This macro is customized to format the data report in a particular way. It can be recustomized to perform other formats as well.  With the below result, our colleagues can continue work on the data report. Useful if report needs to be worked on weekly or daily.</p>

<img src="https://github.com/user-attachments/assets/622d01f8-a253-4225-848f-1ba1a6497751" alt="Result with one group expanded" />



