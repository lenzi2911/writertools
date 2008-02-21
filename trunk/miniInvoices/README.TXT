= MINIINVOICES =

Copyright 2008 Dmitri Popov
Anyone may run the code associated with this document. If you wish to modify or distribute this document and or its code, then you are granted a license to do so under the terms of the GNU General Public License. See:  http://www.gnu.org/licenses/gpl.html

Official page:
http://code.google.com/p/writertools/

The latest development release is available in the following SVN repository:
http://writertools.googlecode.com/svn/trunk/

= NOTES =

miniInvoices is a simple invoicing solution designed specifically for writing professionals. miniInvoices requires the Sun Report Builder extension (http://extensions.services.openoffice.org/project/reportdesign).

To print a single invoice, switch to the Reports section, and run the Single_Invoice report by double clicking on it. In the Parameter Input dialog window, enter the desired invoice number into the Value field and press OK.

Preparing miniInvoices:

  * In the Tables section, use the Publications table to add publications your write for.
  * In the Reports sections, modify the default header of the Single_Invoice report.

==CHANGELOG==

VERSION 0.2.1
  * Conditional Print Expression added to the Country field in the Single_Invoice report.
  * Tweaked logo.
  * Invoice_no field is no Text [VARCHAR].
  * Notes field in the Invoices table added.
  * Conditional print expression fixed.
  * Earnings_by_Year and Earnings_by_Month reports and related queries added.
  * Invoices form cosmetical tweaks.

VERSION 0.1.0 2008-01-25
  * Initial release