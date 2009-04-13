= WRITER'S TOOLS EXTENSION FOR OPENOFFICE.ORG =

Copyright 2008 Dmitri Popov
Anyone may run the code associated with this document. If you wish to modify or distribute this document and or its code, then you are granted a license to do so under the terms of the GNU General Public License. See:  http://www.gnu.org/licenses/gpl.html

Thanks to JohnV, Mark Bain, Konstantin Lavrov, and many OOoForum users for their help.

Official page:
http://code.google.com/p/writertools/

The latest development release is available in the following SVN repository:
http://writertools.googlecode.com/svn/trunk/

==CHANGELOG==

VERSION 1.3.7
  * Post to Microblog tool added (supports Identi.ca and Twitter).
  * Quotations Book reference added to the Lookup Tool.

VERSION 1.3.5 2009-04-07
  * Dutch added to the Google Translate tool.
  * Published status added to the Writer's Submission Tracker.
  * Email Backup error message fixed.
  * Tasks tool now marks overdue tasks.

VERSION 1.3.3 2009-03-21
  * All dialogs and forms now use the Droid font family.
  * Convert To tool removed.
  * Writer's Submission Tracker can store documents.
  * Add to Submission Tracker tool added (not in the menu).
  * Amazon S3 Backup tool is now based on aws: http://timkay.com/aws/

VERSION 1.1.9 2009-01-25
  * Encyclopedia Britannica added to the Lookup Tool.
  * WriterTemplates.oxt removed

VERSION 1.1.7 2008-11-15
  * Writer's Submission tracker added to WriterDB.
  * Amazon S3 Backup tool added.
  * Wikify Word tool removed from the main menu.
  * Minor fixes.

VERSION 1.1.5 2008-09-02
  * Minor fixes.
  * Quick Converter tool improved.

VERSION 1.1.3 2008-08-27
  * Email Backup tool fixes and improvements.

VERSION 1.1.1 2008-08-24
  * Code and comments cleaned up.
  * Minor fixes and improvements in the Email Backup tool.

VERSION 1.1.0 2008-08-19
  * Python-based email backup feature added.

VERSION 1.0.0 2008-08-03
  * New icons.
  * Basket tool renamed to Notebook.
  * Path to the default browser is not saved in the .ini file (yay!).

VERSION 0.9.47 2008-07-27
  * Visual Word Count tool added.
  * Message boxes minor tweaks.
  * Games removed.
  * Confusing Words dictionary added to the Lookup tool.
  

VERSION 0.9.45 2008-07-11
  * Minor fixes in the Quick converter tool.
  * New Bookmarks tool added.
  * Export tasks to Calc macro has been reworked. It now creates a pie chart of the finished and unfinished tasks.
  * Default browser fall back fix. If default browser is not found, the user can enter the path to the desired browser manually.

VERSION 0.9.43 2008-06-27
  * New languages added to the Google Translate tool: Bulgarian, Croatian, Czech, Danish, Finnish, Hindi, Norwegian, Polish, Romanian, and Swedish.
  * Minor tweaks to the Post to Twitter tool.
  * Quick Converter tool added.
  * Games removed from the Writer's Tools menu.

VERSION 0.9.41 2008-05-09
  * Added a cancel action to the Convert To tool.
  * Functions moved to a separate module.
  * Verb Conjugator tool added (not in the main menu).
  * Writer's Tasks report added (Required the Sun Report Builder extension).
  * Export to Calc feature added to the Task Viewer.
  * Detect language option added to the Google Translate tool.

VERSION 0.9.39 2008-04-07
  * Error handling added to the OpenDatabase routine.

VERSION 0.9.37 2008-04-07
  * The OpenForm function has been replaced with the OpenDatabase routine, which opens the WriterDB database instead of individual forms.
  * Minor tweaks in the WriterDB database.

VERSION 0.9.35 2008-04-07
  * Completely rewritten the OpenForm function.
  * Minor tweaks and improvements.

VERSION 0.9.33 2008-04-07
  * Important fix in the OpenForm function.
  * Minor tweaks and improvements.

VERSION 0.9.31 2008-04-07
  * "cmd /c start" is used to launch the default browser on Windows.
  * "/etc/alternatives/x-www-browser" is used to launch the default browser on Linux.
  * Minor tweaks and improvements.

VERSION 0.9.29 2008-04-07
  * Reworked and simplified Tasks tool.
  * Minor tweaks and improvements.

VERSION 0.9.27 2008-03-29
  * Minor bug fixes and tweaks.
  * All messages and error warnings moved to the Tools module for easier localization.

VERSION 0.9.25 2008-03-19
  * Minor improvements.
  * Error handling improved.

VERSION 0.9.23 2008-03-19
  * The Letter Hunt game improved. It now uses words without any separators.

VERSION 0.9.21 2008-03-19
  * Improvements based on cenacchi's code.
  * BrowseURL function is now used to detect and open the default browser.
  * Lookup Tool can handle multiple words searches (e.g. you can lookup "World War II").
  * Lookup Took can handle special characters (e.g. you can lookup "Procter&Gamble", quoted
text, etc.).

VERSION 0.9.19 2008-03-19
  * Minor improvement: OpenForm function added.
  * Notes field added to the tasks table.
  * Writer's Tasks form added. It can be opened from the Tasks Manager window.

VERSION 0.9.17 2008-03-18
  * Writer's Notes database table added.
  * Notes tool added.

VERSION 0.9.15 2008-03-17
  * The Tasks tool has been rewritten. The Task Manager allows the user to create, view, and delete tasks from a single interface. Tasks can be deleted individually or in bulk.
  * Code cleaned up. Database connection is now established using a function.

VERSION 0.9.13 2008-03-15
  * List boxes in the Lookup and Convert To tools are populated dynamically via arrays.
  * App Launcher tool added.

VERSION 0.9.11 2008-03-11
  * Minor regrouping of macros.
  * Post to Twitter tool added (hidden).
  * Multi-format backup now uses the built-in zip feature.

VERSION 0.9.10 2007-03-09
  * Rewritten Convert To tool replaced the Convert to HTML and Covert to DOC tools.

VERSION 0.9.9 2007-03-07
  * Convert to HTML and Covert to DOC tools added.
  * Q and A game improvements:
    * The system marks correct answers, and they are not ued in the game again.
	* User can exclude records from the game by ticking the Status check box.
	* The Counter field keeps track of the number of correct/incorrect answers.
  * Letter Hunt Game added.
  * All games moved into the WriterGames module.

VERSION 0.9.7 2008-02-27
  * Proper error message when the default browser is not found.
  * Basket form improvements.
  * Tasks to Calc export tool added.
  * Support for the Extension Manager Update feature.

VERSION 0.9.6 2008-01-30
  * Tasks in the WriterDB database can now be marked as done, which are not displayed when running the Show Tasks command.
  * Basket form and Browse Basket command added.

VERSION 0.9.5 2007-12-11
  * Minor improvements to the Translit tool.
  * Q and A game added.

VERSION 0.9.3 2007-10-13
  * Removed StatusText from the StartTiming macro due to its erratic behavior.
  * Minor fixes for the Writer Templates extension.
  * Save Session dialog box added to the Timer tool. It's now possible to add notes to the saved session.
  * Translit tool added. Converts latin to cyrillic (experimental and is not in the Writer's Tools menu).
  * Post to Gelato tool removed.

VERSION 0.9.2 2007-10-11
  * Added a link to the Writer for Writers book instead of Donate.

VERSION 0.9.1 2007-10-10
  * Status messages added to the Timer tool.
  * wnotes form added to the WriterDB database.
  * Meaningful message when the task list in the Tasks tool is empty.
  * Added a confirmation dialog when purging tasks.
  * Show on the Map tool switched to Google Maps.
  * Multi-format Backup tool now uses timestamps for folder names.
  * Multi-format Backup can use 7-Zip (must be installed separately) to compress the resulting backup folder.

VERSION 0.9.0 2007-10-08
  * Minor fixes in the Word Hunt game.
  * Minor cosmetic tweak to dialog windows.
  * Added package meta info (provides extension info in Extension Manager).
  * Post to Gelato tool added (experimental and is not in the Writer's Tools menu).
  * Tasks tool added.

VERSION 0.8.9.2 2007-10-04
  * Word Source reference added to the Lookup tool.
  * Fixed a few serious bugs in Google Translate.
  * Improved default browser detection on Windows and Linux.

VERSION 0.8.8 2007-09-29
  * Wikify Word tool added.
  * Writer templates (generic article and book) added.
  * User-friendly error messages for failed database connection.
  * Minor fixes.

VERSION 0.8.7.1 2007-09-22
  * Minor bug fixes.
  * Donate menu item added.

VERSION 0.8.7 2007-09-21
  * Completely reworked the Add to Basket tool.
  * Convert to DokuWiki tool removed (OpenOffice.org 2.3 has a built-in MediaWiki export filter).

VERSION 0.8.6.1 2007-08-04
  * The Mode option added to the Google Translate Tool.
  * Word Count statistics added to the Timer tool.
  * Minor fixes in the Timer tool.

VERSION 0.8.5.1 2007-06-30
  * Minor improvements and bug fixes.

VERSION 0.8.5 2007-06-20
  * The Open FTP Document tool has been removed in favor of the new Remote Backup tool.
  * Minor improvements to the Word Hunt game.

VERSION 0.8.3 2007-06-19
  * Cosmetic changes of the messages boxes.
  * Minor improvements to the Word Hunt game.
  * Show on the Map tool added.

VERSION 0.8.1 2007-06-17
  * The Word Hunt game improved.
  * Added BEOLINGUS reference to the Lookup tool.
  * Word of the Day tool added.
  * Add to Basket tool added.

VERSION 0.7.2 2007-06-11
  * Ask Oxford reference added to the Lookup tool.
  * Added Russian to the Google Translate tool (thanks to Konstantin Lavrov).
  * Fixed the HTTP link for the Google Translate tool on Linux.

VERSION 0.7.1 2007-06-08
  * The Timer tool (Start/Stop Timer) can now save the time data in the accompanying WriterDB database.
  * Added a simple Word Hunt game, where you have to guess the word randomly picked by Writer.
