import debit_list
from shutil import copyfile

copyfile('/Users/frodin/Downloads/Alla Aktiva (generell) 201810140719.xls',
         r'/Users/frodin/work/projects/debit_list_converter/AllaAktiva.xls')

file_name = '/Users/frodin/work/projects/debit_list_converter/AllaAktiva.xls'
dl = debit_list.DebitList(file_name)

# write email addresses
dl.write_email_list('EmailList.xls')
dl.write_short_estate_list('EstateList.xls')


