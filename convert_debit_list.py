import debit_list

file_name = '/Users/frodin/work/projects/python/TSffDebitList.xlsx'
dl = debit_list.DebitList(file_name)

# write email addresses
dl.write_email_list('EmailList.xls')
dl.write_short_estate_list('EstateList.xls')
