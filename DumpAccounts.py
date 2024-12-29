import gnucashxml
from datetime import date

# Dumps information for selected accounts (in 'test_list')

# Name of your GnuCash book file
GnuCash_Book = "<your book name>.gnucash"

# Get current date and set other variables
today = date.today()

test_list = ["Charity", "Donation", "Health Insurance", "Long Term Care Insurance", "Mortgage Interest",
"Investing Expenses", "Medical Expenses", "Devices, etc.", "Doctor", "Eyecare", "Medicine", "Taxes",
"Federal US", "Estimated Tax Payment", "Prior Year Tax Payment", "Other Tax", "Property Tax", "State/Local"]

def formatNone(val):
    if val is None:
        return "*None*"
    else:
        return val

def process_account(par, acc):
    struct = par+":"+acc.name
    #print("{:112} TYPE:{:16} #CHILDREN:{:4} #SPLITS:{:4}"
    #      .format(struct, formatNone(acc.actype), len(acc.children), len(acc.splits)))
    #print("--{} Name:{:32} Parent:{:32} Type: {:16} #CHILDREN:{:4} #Splits:{:4} #Slots:{:4}"
    #      .format(acc_type, acc.name, acc.parent.name, acc.actype, len(acc.children), len(acc.splits), len(acc.slots)))

    if len(acc.splits) > 0:
        print("{:112} TYPE:{:16} #CH:{:4} "
              "#SP:{:4}"
              .format(struct, formatNone(acc.actype), len(acc.children), len(acc.splits)))
    #    for split in acc.splits:
    #        process_split(split)

    if len(acc.children) > 0:
        for child in acc.children:
            #print("Child:", child.name)
            process_account(struct, child)

def process_split(spl):
    #print("--SPLIT: Acct:{:32} Value:{:9.2F} #Slots:{:4}"
    #      .format(split.account.name, split.value, len(split.slots)))
    process_transaction(spl.account.name, spl.transaction)

def process_transaction(acc_name, trans):
    #print("----TRX: Date:", trans.date.strftime("%m/%d/%Y"), "Num:", trans.num, "Desc:", trans.description, "#Splits:", len(trans.splits))
    #print("----Trans: Date:{:10} Num:{} Desc:{:32} #Splits:{:4}".format(trans.date.strftime("%m/%d/%Y"), trans.num, trans.description, len(trans.splits)))
    print("ACC_NAME:", acc_name)
    spl_cnt = len(trans.splits)
    spl_ctr = 0
    for spl in trans.splits:
        print("--TRX-SPLIT", spl_ctr, "/", spl_cnt, ": DATE:", trans.date.strftime("%m/%d/%Y"), " NUM:", trans.num, " DESC:", trans.description,
              " SPL-ACCT:", spl.account.name, "  AMT:", spl.value)
        spl_ctr += 1

# --------------------------------------------------------------------------------

book = gnucashxml.from_filename(GnuCash_Book)

for account, children, splits in book.walk():
    if len(splits) > 0:
        print("ROOT.{:48} PARENT: {:32} TYPE: {:16} #CHILDREN:{:3} #SPLITS:{:4}".format(account.name, formatNone(account.parent.name), formatNone(account.actype), len(children), len(splits)))
        if account.name in test_list:
            for split in account.splits:
