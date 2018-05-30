# emailmerge
Creating email merges based on Excel Spreadsheet data on a Mac using R

This is an inital sketch of the process of getting an Excel Spreadsheet to act as a source for emails (inclduing attaching a different file to each message). 

The scripts:

* are not smart- there is no effort made to make decisions about what to do if entries are blank
* are not escaped- there is no effort to make sure quote marks are escaped properly

This is intended as a first draft that shows the worklfow involved:

* R, through using readxl, understands the spreadsheet and that each row contains information about a message to be sent
* arranges that data into a form for sending (via Open Scripting Architecture (OSA) Applescript commands) to either Mail.App or Microsoft Outlook (Mac).

Contents:

A Mail.app focused example

A MS Outlook (Mac) focused example

An excel sheet that works with the examples.

