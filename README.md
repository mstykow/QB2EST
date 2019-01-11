# Converting QuickBooks contacts to EST 2.0 format
Program converts QuickBooks contacts exported to an Excel file into an EST 2.0 (Canada Post) importable text file.

After an order has been invoiced and a customer profile has been created in QuickBooks, it would be nice if there was a quick way to get this customer information into Canada Post's large-volume shipping label tool called EST 2.0. That's what this program does.

* See [here](https://smallbusiness.chron.com/export-quickbooks-customer-list-60392.html) for exporting your QuickBooks contacts to an Excel file.

* See [here](https://www.canadapost.ca/cpc/en/business/shipping/find-rates-ship/est-2.page) for downloading EST 2.0 and [here](https://www.canadapost.ca/cpo/mc/assets/pdf/business/import_2016_en.pdf) for information about the file format required for importing contacts.

If updating addresses, I recommend deleting your entire address book from EST 2.0 before importing the new data to avoid double entries.

Additional comments:

* Place your QuickBooks Excel export file into the same directory as these two scripts, run QB2EST.py, and follow prompts.
* You may need to install the following modules:
  * openpyxl: pip install openpyxl
  * pycountry: pip install pycountry
* Tested and created on QuickBooks Pro 2015.
