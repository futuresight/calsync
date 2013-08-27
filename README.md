calsync
=======

Calendar Sync for Outlook

.NET 4.5 Office 2013 Add-In to Sync Outlook Calendar with Google account.

Features
--------

* Sync all calendars between Outlook and Google.
* Supports calendar event subject/summary, description, time/duration, and location.
* Events are synced based on last modification time.
* Efficient sync method only syncs data changed since last sync (option to perform full sync included).
* Good integration with the Office Outlook Ribbon.
* Extensive options on sync time and modes, allows exclusion of calendars from sync.
* Bonus feature: if properly set up, adds an "Upcoming birthdays" section to Outlook Today that automatically gets birtdays from the Contact birthdays calendar and lists ones in the next 3 months.

Build preparation
-----------------

This project is best built with Visual Studio 2012. It requires a professional version of Visual Studio with the latest VSTO installed. To run it, you need Microsoft Outlook 2013.

All DLL dependencies are included in the bin/Debug directory, but if you encounter problems with these dependencies you should download the NuGet package https://www.nuget.org/packages/Google.Apis.Calendar.v3

Before building, you must configure the program to use your personal access keys to the Google APIs. To do this, follow the instructions given in the APIConfig.vb class.

Build
-----

Build normally using F5 or the build solution button. To actually install the DLL into Office 2013, the easiest and most practical way to do so is simply Debugging the program. From then on, running the addin from Visual Studio is not required: it will start automatically with Outlook.

Uninstall from Office
---------------------

You can unregister the addin from the Trust Center in Outlook.

Unfixed Bugs
------------

Sometimes, when calendars are deleted from Outlook they are not deleted from the Google server.
