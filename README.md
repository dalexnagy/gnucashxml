# gnucashxml
Python library to read information from a GnuCash XML file

This is a slightly modified version of the original Python library created by Jorgen Schaefer in 2012.  All credit for copyright goes to the author Jorgen Schaefer <forcer@forcix.cx>.

The library seems to be abandoned and I wanted to add support for the transaction 'num' field.  This version 1.1 does have support to return the transaction num.

If you install 'gnucashxml' through PIP, you will get version 1.0 without support described above.  I found copies of this library in three folders on my Ubuntu 20.04 & Python 3.8 system:
  ~/lib/python3.8/site-packages
  /usr/local/lib/python3.8/site-packages
  ~/venv/lib/python3.8/site-packages

In this repository are examples of Python code written using this updated version gnucashxml to:
  1. Dump the GnuCash structure elements (helped my development process)
  2. Create a spreadsheet of current account balances
  3. Create a report and/or spreadsheet of transactions in an account for a specific period (PyQT GUI)
  4. Check if any entries exist in the 'Imbalance' account and notify me via email if any are found

