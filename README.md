# gnucashxml
Python library to read information from a GnuCash XML file

This is an updated and slightly modified version of the original Python library created by Jorgen Schaefer in 2012. All credit for copyright goes to the author Jorgen Schaefer (jorgen.schaefer@gmail.com).

This version (2.0) support Python V3 (3.6 & 3.8 support is verified - I have not verified other levels) and has support to return the transaction num.

-------------

class Book(object):
	A book is the main container for GNU Cash data.

    It doesn't really do anything at all by itself, except to have
    a reference to the accounts, transactions, and commodities.
	
	Implemented:
	 - book:id
	 - book:slots
	 - gnc:commodity
	 - gnc:account
	 - gnc:transaction
	
	Not implemented:
	 - gnc:schedxaction
	 - gnc:template-transactions
	 - gnc:count-data

	
class Commodity(object):
	A commodity is something that's stored in GNU Cash accounts.

    It consists of a name (or id) and a space (namespace).
	
	Implemented:
	 - cmdty:id
	 - cmdty:space

	Not implemented:
	 - cmdty:get_quotes => unknown, empty, optional
	 - cmdty:quote_tz => unknown, empty, optional
	 - cmdty:source => text, optional, e.g. "currency"
	 - cmdty:name => optional, e.g. "template"
	 - cmdty:xcode => optional, e.g. "template"
	 - cmdty:fraction => optional, e.g. "1"


class Account(object):
    An account is part of a tree structure of accounts and contains splits.
		
	Implemented:
	 - act:name
	 - act:id
	 - act:type
	 - act:description
	 - act:commodity
	 - act:commodity-scu
	 - act:parent
	 - act:slots


class Transaction(object):

    A transaction is a balanced group of splits.
	
	Implemented:
	 - trn:id
	 - trn:currency
	 - trn:num
	 - trn:date-posted
	 - trn:date-entered
	 - trn:description
	 - trn:splits / trn:split
	 - trn:slots	

class Split(object):

    A split is one entry in a transaction.
	
	Implemented:
	 - split:id
	 - split:memo
	 - split:reconciled-state
	 - split:reconcile-date
	 - split:value
	 - split:quantity
	 - split:account
	 - split:slots
		 Implemented:
		 - slot
		 - slot:key
		 - slot:value
		 - ts:date
		 - gdate
