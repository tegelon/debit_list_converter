import debit_list, os
from shutil import copyfile

def match_dump_files(filename):
    import re
    match = 'Alla Aktiva\s\(generell\)\s\d{12}\.xls'
    m = re.search(match,filename)
    return m.group(0) if m is not None else None


if __name__ == "__main__":
    
    dump_dir = '/Users/frodin/Downloads/'
    out_dir = '/Users/frodin/work/projects/debit_list_converter/'
    file_list = os.listdir(dump_dir)
    dump_files = [match_dump_files(f) for f in file_list if match_dump_files(f) is not None]
    if dump_files:
        latest_dump_file = dump_files[len(dump_files)-1]
        print('Reading file: '+ latest_dump_file)
        dl = debit_list.DebitList(dump_dir+latest_dump_file,out_dir)

        # Write lists
        dl.write_email_list('EmailList.xls')
        dl.write_shared_estate_list('SharedEstateList.xls')
        dl.write_estate_list('EstateList.xls')
    else:
        print('No database dump file available, try download a file from foreningshuset.')




