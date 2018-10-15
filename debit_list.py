#!/usr/bin/python
          # -*- coding: utf-8 -*-
import os, sys, xlrd, xlwt, shutil, errno, re, locale
from enum import Enum
locale.setlocale(locale.LC_ALL, 'sv_SE.utf-8'); 

class HeaderFields(Enum):
    estateId = 1
    firstName = 2
    lastName = 3
    email = 4
    address = 5
    zipcode = 6
    city = 7
    parkinglot = 8
    GA2mooring = 9
    GA3mooring = 10

    
class DebitList:
    def __init__(self,filepath):
        workDir, fileName = os.path.split(filepath)
        self.filepath = filepath;
        self.__workDir__ = workDir;
        self.__book__ = None;

        
    # Public methods #
    def write_email_list(self, fileName):

        file_path = os.path.join(self.__workDir__, fileName);
        book = xlwt.Workbook()
        emailSheet = book.add_sheet("EmailList")

        # get all contacts
        self.__parse_list__();

        # write header
        email_list = [(self.translate_write(HeaderFields.firstName),
                       self.translate_write(HeaderFields.lastName),
                       self.translate_write(HeaderFields.email))];

        for n, est in enumerate(self.__estates__):
            contacts = est.get_contacts();        
            for contact in contacts:
                email_list.append([contact.get_firstname(),
                                   contact.get_lastname(),
                                   contact.get_email()
                ]);

        email_list = self.__remove_list_duplicates__(email_list);
                     
        for rowidx, rowdata in enumerate(email_list):
            row = emailSheet.row(rowidx);
            for colidx, str in enumerate(rowdata):
                row.write(colidx, str)
        
        # save worksheet
        book.save(file_path);
        print 'Wrote file ' + file_path

    def write_short_estate_list(self, fileName):

        file_path = os.path.join(self.__workDir__, fileName);
        book = xlwt.Workbook()
        emailSheet = book.add_sheet("EmailList")

        # get all contacts
        self.__parse_list__();

        # write header
        short_estate = [[self.translate_write(HeaderFields.firstName),
                         self.translate_write(HeaderFields.lastName),
                         self.translate_write(HeaderFields.email),
                         self.translate_write(HeaderFields.estateId),
                         self.translate_write(HeaderFields.parkinglot),
                         self.translate_write(HeaderFields.GA2mooring),
                         self.translate_write(HeaderFields.GA3mooring)]];
        
        for n, est in enumerate(self.__estates__):
            contacts = est.get_contacts()
            if contacts:
                firstname = contacts[0].get_firstname()
                lastname = contacts[0].get_lastname()
                email = contacts[0].get_email()

                for k, contact in enumerate(contacts):
                    if k > 0:
                        firstname = firstname + '\n' + contact.get_firstname()
                        lastname = lastname + '\n' + contact.get_lastname()
                        email = email + '\n' + contact.get_email()

                mooring_sommarbo = est.get_GA2_mooring();
                m_sommarbo = '';
                if mooring_sommarbo:
                    m_sommarbo = mooring_sommarbo[0];
                    for l, m in enumerate(mooring_sommarbo):
                        if l > 0:
                            m_sommarbo = m_sommarbo + '\n' + m

                
                mooring_tegelon = est.get_GA3_mooring();
                m_tegelon = '';
                if mooring_tegelon:
                    m_tegelon = mooring_tegelon[0];
                    for l, m in enumerate(mooring_tegelon):
                        if l > 0:
                            m_tegelon = m_tegelon + '\n' + m

                
                short_estate.append([firstname,
                                     lastname,
                                     email,
                                     est.get_estate(),
                                     est.get_parkinglot(),
                                     m_sommarbo,
                                     m_tegelon])                     

        for rowidx, rowdata in enumerate(short_estate):
            row = emailSheet.row(rowidx);
            for colidx, str in enumerate(rowdata):
                if str:
                    row.write(colidx, str);
                
        # save worksheet
        book.save(file_path);
        print 'Wrote file ' + file_path

        
    # Private methods #
    def __parse_list__(self):
        
        sheetIdx = 0;
        self.__read_xls_sheet__(sheetIdx);
        header = DebitListHeader(self.__get_sheet__());
        contents = self.__get_contents__();

        self.__estates__ = [Estate(header, row) for row in contents];
        for est in self.__estates__:
            est.parse_estate()

    def __read_xls_sheet__(self, sheetIdx):
        if self.filepath:
            if not self.__book__:
                self.__book__ = xlrd.open_workbook(self.filepath);
            self.__sheet__ = self.__book__.sheet_by_index(sheetIdx);

    def __remove_list_duplicates__(self, debit_list):
        
        unique_dict = dict();
        unique_list = list();
        for rowdata in debit_list:
            key = ''.join(rowdata);
            if key not in unique_dict:
                unique_dict[key] = None;
                unique_list.append(rowdata);
        
        return unique_list;
        

    def __copyfile__(self, dst):
        try:
            shutil.copytree(self.filepath, dst)
        except OSError as exc: # python >2.5
            if exc.errno == errno.ENOTDIR:
                shutil.copy(self.filepath, dst)
            else: raise
            
    def __add_sheet__(self, name):
        self.__sheet__ = self.__book__.add_sheet(name)

    def __get_sheet__(self):
        return(self.__sheet__);

    def __get_row__(self, rowIdx):
        return(self.__sheet__.row_values(rowIdx));

    def __get_contents__(self):
        contents = [self.__get_row__(row)
                    for row in range(0, self.__sheet__.nrows)];
        return contents;

    @staticmethod
    def translate_write(headerFieldEnum):
        return DebitListHeader.translate_write(headerFieldEnum);

    
class DebitListHeader:
    
    __FIRSTCOLUMNTOKEN__ = 'ID';
    
    __translation_dict_read__ = {HeaderFields.estateId:'^Fastighet$',
                                 HeaderFields.firstName:'^F\xf6rnamn$',
                                 HeaderFields.lastName:'^Efternamn$',
                                 HeaderFields.email:'^Epost\s\(bostad\)$',
                                 HeaderFields.address:'^Gatuadress$',
                                 HeaderFields.zipcode:'^Postnr$',
                                 HeaderFields.city:'^Ort$',
                                 HeaderFields.parkinglot:'^Parkeringsplats$',
                                 HeaderFields.GA2mooring:'^GA:2\sB\xe5tplats$',
                                 HeaderFields.GA3mooring:'^GA:3\sB\xe5tplats$'};
    
    
    __translation_dict_write__ = {HeaderFields.estateId:u'Fastighet',
                                  HeaderFields.firstName:u'F\xf6rnamn',
                                  HeaderFields.lastName:u'Efternamn',
                                  HeaderFields.email:u'Epost',
                                  HeaderFields.address:u'Gatuadress',
                                  HeaderFields.zipcode:u'Postnr',
                                  HeaderFields.city:u'Ort',
                                  HeaderFields.parkinglot:u'Parkeringsplats',
                                  HeaderFields.GA2mooring:u'B\xe5tplats Sommarbo',
                                  HeaderFields.GA3mooring:u'B\xe5tplats Tegelön'};
    
    def __init__(self, sheet):

        #ÅÄÖ åäöé
        char_str = '[-\w\xc5\xc4\xd6\xe5\xe4\xf6\xe9]+';

        header_row = self.__read_header_row__(sheet);
        self.__header_row__ = header_row[0]; #self.__readHeaderRow__(sheet);
        self.__row_idx__ = header_row[1];
        self.__key__ =  [(HeaderFields.estateId,'Tegel\xf6n\s\d:\d{1,3}$'),
                         (HeaderFields.firstName,char_str),
                         (HeaderFields.lastName,char_str),
                         (HeaderFields.email,'[\w\.-]+@[\w\.-]+'),
                         (HeaderFields.address,None),
                         (HeaderFields.zipcode,'\d{3}\s?\d{2}'),
                         (HeaderFields.city,char_str),
                         (HeaderFields.parkinglot,'[0-9]+'),
                         (HeaderFields.GA2mooring,'(S[p|k]\s\d{1,3},?)+'),
                         (HeaderFields.GA3mooring,'(Tp\s\d{1,2},?)+')];
        
    def __read_header_row__(self, sheet):
        first_column = sheet.col_values(0);
        for idx, cell in enumerate(first_column):
            if cell == DebitListHeader.__FIRSTCOLUMNTOKEN__:
                return((sheet.row_values(idx), idx));

    def get_header(self):
        return(self.__header_row__);

    def get_row_idx(self):
        return(self.__row_idx__);

    def get_key(self):
        return(self.__key__);

    def get_dictionary(self):
        header = self.get_header();

        compiled_key = [(k,re.compile(self.translate_read(k))) for k,pat in self.__key__];

        # For each cell in header row: match with all keys and create
        # - a dictionary translating to a column number if match
        # - or insert None
        key_dictionary_raw = [{k:idx} if cp.match(cell) != None else None
                              for (k,cp) in compiled_key
                         for (idx, cell) in enumerate(header)];

        # Remove invalid entries
        key_dictionary_list = [k for k in key_dictionary_raw if k is not None];
        
        key_dictionary = {};
        for item in key_dictionary_list:
            key_dictionary.update(item);
        return key_dictionary;

    @staticmethod
    def translate_read(headerFieldEnum):
        return DebitListHeader.__translation_dict_read__.get(headerFieldEnum);
    @staticmethod
    def translate_write(headerFieldEnum):
        return DebitListHeader.__translation_dict_write__.get(headerFieldEnum);

class Contact:

    def __init__(self,firstname,lastname,email,address,zip,city):

        #self.__N__ = len(email);
        #assert(len(first_name) == self.__N__ &
        #       len(last_name) == self.__N__ &
        #       len(address) == self.__N__ &
        #       len(zip) == self.__N__ &
        #       len(city) == self.__N__);
        
        self.__firstname__ = firstname;
        self.__lastname__ = lastname;
        self.__email__ = email;
        self.__address__ = address;
        self.__zip__ = zip;
        self.__city__ = city;

    def get_firstname(self):
        return(self.__firstname__);

    def get_lastname(self):
        return(self.__lastname__);

    def get_email(self):
        return(self.__email__);

    def get_address(self):
        return(self.__address__);

    def get_zip(self):
        return(self.__zip__);

    def get_city(self):
        return(self.__city__);

    def set_first_name(self,firstName):
        self.__firstName__ = firstName;
        
    def set_last_name(self,lastName):
        self.__lastName__ = lastName;

    def set_email(self,email):
        self.__email__ = email;

    def set_address(self,address):
        self.__address__ = address;

    def set_zip(self,zipcode):
        self.__zip__ = zipcode;

    def set_city(self,city):
        self.__city__ = city;

    def __len__(self):
        return len(self.__email__);

    def print_contact(self):
        print(self.__firstName__ + ' ' +
              self.__lastName__ + ' ' +
              self.__email__ + ' ' +
              self.__address__ + ' ' +
              self.__zip__ + ' ' +
              self.__city__);


    
class Estate:

    def __init__(self,header,row):
        self.__header__ = header;
        self.__raw__ = row;
        self.__dictionary__ = None;
        self.__contacts__ = list();
        self.__info__ = dict();

    def get_estate(self):
        return self.__estate__

    def get_contacts(self):
        return self.__contacts__

    def get_parkinglot(self):
        return self.__parkinglot__

    def get_GA2_mooring(self):
        return self.__GA2_mooring__

    def get_GA3_mooring(self):
        return self.__GA3_mooring__

    def __init_contents_dictionary__(self):
        self.__contents_dictionary__ = dict((key,self.__raw__[idx])
                                   for key,idx in
                                   self.__header__.get_dictionary().iteritems());

    def __init_contents_regex_pattern__(self):
        self.__contents_regex_pattern__ = dict((k,re.compile(pat) if pat != None else None)
                                             for k,pat in
                                             self.__header__.get_key());
    def parse_estate(self):

        self.__init_contents_dictionary__();
        self.__init_contents_regex_pattern__();
        
        # Parse contacts
        this_estate = self.__parse_cell__(HeaderFields.estateId);
        self.__estate__ = this_estate;

        if this_estate:
            
            email_addresses = self.__parse_cell__(HeaderFields.email);
            firstnames = self.__parse_cell__(HeaderFields.firstName);
            lastnames = self.__parse_cell__(HeaderFields.lastName);
            address = self.__parse_cell__(HeaderFields.address);
            zipcode = self.__parse_cell__(HeaderFields.zipcode);
            city = self.__parse_cell__(HeaderFields.city);
        
            composite_list = [firstnames, lastnames, email_addresses, address, zipcode, city];
            composite_list_t = [list(x) for x in zip(*composite_list)];
            contacts = [(Contact(fn, ln, email, address, zipcode, city))
                        for fn, ln, email, address, zipcode, city in composite_list_t];
            
            self.__contacts__ = contacts;

            # Parse parkinglot
            parking = self.__parse_cell__(HeaderFields.parkinglot);
            self.__parkinglot__ = parking[0] if parking else None;
            
            # Parse mooring
            ga2_mooring = self.__parse_cell__(HeaderFields.GA2mooring);
            self.__GA2_mooring__ = None if not ga2_mooring else ga2_mooring;
            ga3_mooring = self.__parse_cell__(HeaderFields.GA3mooring);
            self.__GA3_mooring__ = None if not ga3_mooring else ga3_mooring;
        

    def __parse_cell__(self,category):
        data = self.__contents_dictionary__.get(category);
        pattern = self.__contents_regex_pattern__.get(category);
        # print '------------------\n' + data ### DEBUG ###

        return(pattern.findall(data) if pattern != None else [data]);
        

    @staticmethod
    def __extend_list_with_last_item__(this_list,length):
        d = length - len(this_list);
        last_item = this_list[-1] if len(this_list) > 1 else this_list;
        if d > 0:
            for n in range(0,d):
                this_list.extend(last_item);

               
# C-x 8 RET 005B/D RET - inserts [/]
#dl = DebitList('/Users/frodin/work/projects/python/TSffDebitList.xlsx');
#dl.writeEmailList('EmailList.xlsx')
