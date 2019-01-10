# Converting QuickBooks contacts to EST 2.0 format
Program converts QuickBooks contacts exported to an Excel file into an EST 2.0 (Canada Post) importable text file.

* See [here](https://smallbusiness.chron.com/export-quickbooks-customer-list-60392.html) for exporting your QuickBooks contacts to an Excel file.

* See [here](https://www.canadapost.ca/cpc/en/business/shipping/find-rates-ship/est-2.page) for downloading EST 2.0 and [here](https://www.canadapost.ca/cpo/mc/assets/pdf/business/import_2016_en.pdf) for information about the file format required for importing contacts.

If updating addresses, I recommend deleting your entire address book from EST 2.0 before importing the new data to avoid double entries.

Additional comments:

* Place your QuickBooks export Excel file into the same directory as this script, run script, and follow prompts.
* You may need to install the following modules:
  * openpyxl: pip install openpyxl
  * pycountry: pip install pycountry
* Tested and created on QuickBooks Pro 2015.
* See [here](https://medium.com/dreamcatcher-its-blog/making-an-stand-alone-executable-from-a-python-script-using-pyinstaller-d1df9170e263) to convert script into a self-contained executable (Windows) file.
