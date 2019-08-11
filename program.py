# This Python file uses the following encoding: utf-8
import sqlite3
from sqlite3 import Error 
import os
from datetime import datetime
import xlrd
import xlwt
import progressbar
import requests
import json
import socket
import threading
import codecs
import base64
import hashlib
import time

class ForclosureProgram(object):
    v_dict_questions = {}
    v_dict_answers = {}
    v_client_socket = None
    @staticmethod
    def create_connection(db_file):
        """ Create Database Connection
        """    
        try:
            conn = sqlite3.connect(db_file)
            return conn
        except Error as e:
            print(e)
        return None
    @staticmethod
    def create_db_schema(conn):
        """ Create Schema of Database"""
        sql_create_files_table = """ CREATE TABLE IF NOT EXISTS files (
                                            id integer PRIMARY KEY,
                                            name text NOT NULL,
                                            processed_date text
                                        ); """
        sql_create_file_records_table = """ CREATE TABLE IF NOT EXISTS file_records (
                                            id integer PRIMARY KEY,
                                            file_id integer NOT NULL,
                                            record_id integer NOT NULL,
                                            header_id integer NOT NULL,
                                            value text
                                        ); """
        sql_create_file_records_index = """ 
                                            CREATE UNIQUE INDEX IF NOT EXISTS idx_uk_file_records ON file_records(file_id, record_id, header_id);
                                        """
        sql_create_headers_table = """ CREATE TABLE IF NOT EXISTS headers (
                                            id integer PRIMARY KEY,
                                            header text NOT NULL UNIQUE,
                                            is_source integer default 0
                                        ); """
        sql_count_headers = """ select count(*) from headers;
                            """
        sql_insert_headers = """ insert into headers(id, header, is_source) values (1, 'TYPE', 1), (2, 'ATTORNEY', 1), (3, 'PLAINTIFF', 1), 
                                (4, 'Sheriff’s #/Courts Case #', 1), (5, 'DEFENDANT', 1), (6, 'ADDRESS', 1), (7, 'PARCEL', 1), 
                                (8, 'STATUS', 1), (9, 'PRINCIPAL', 1),
                                (10, 'Zillow_Estimate', 0), (11, 'ZIP',0), (12, 'ZLiving Area',0), (13, 'ZBed',0), (14, 'ZBath',0), (15, 'ZStatus',0)
                                , (16, 'ZTime on zillow',0), 
                                (17, 'ZSchool1 Name',0), (18, 'ZSchool1 Grade',0), (19, 'ZSchool1 Distance',0), (20, 'ZSchool1 Rating', 0),
                                (21, 'ZSchool2 Name',0), (22, 'ZSchool2 Grade',0), (23, 'ZSchool2 Distance',0), (24, 'ZSchool2 Rating', 0),
                                (25, 'ZPrice History', 0), (26, 'nccd_response',2), (27, 'NSubdivision',0), (28, 'NMunicipal Info',0), (29, 'NZoning',0),
                                (30, 'NCounty Balance Due', 0), (31, 'NSchool Balance Due', 0), (32, 'NRecent Sales', 0),
                                (33, 'parcel_id', 2), (34, 'zillow_response',2),(35, 'Zillow_ZEstimate', 0)
                             """
        if(conn is not None):
            try:
                c = conn.cursor()
                c.execute(sql_create_files_table)
                c.execute(sql_create_file_records_table)
                c.execute(sql_create_file_records_index)
                c.execute(sql_create_headers_table)
                c.execute(sql_count_headers)
                row = c.fetchone()
                if(row[0] == 0):
                    """Insert default rows if not present in the tables"""
                    c.execute(sql_insert_headers)
                    conn.commit()
            except Error as e:
                print(e)

    @staticmethod
    def patch_db_schema(conn):
        sql_insert_headers = """ insert into headers(id, header, is_source) 
                                SELECT 35, 'Zillow_ZEstimate', 0 
                                WHERE NOT EXISTS(SELECT 1 FROM headers WHERE id = 35)
                             """
        sql2 = """ UPDATE file_records set value = replace(value,'=','') WHERE header_id = 33
                """
        sql3 = """ UPDATE headers set header = 'Zillow_Estimate' WHERE id = 10
                """
                

        if(conn is not None):
            try:
                c = conn.cursor()
                c.execute(sql_insert_headers)
                c.execute(sql2)
                c.execute(sql3)
                conn.commit()
            except Error as e:
                print(e)
 
    @staticmethod
    def create_file(conn, file):
        """create a new file entry and return id"""
        sql = """ insert into files(name, processed_date) 
                    values (?, ?) """
        if(conn is not None):
            try:
                c = conn.cursor()
                c.execute(sql, file)
                conn.commit()
                return c.lastrowid
            except Error as e:
                print(e)
        return None
    @staticmethod
    def create_file_record(conn, file_record):
        """create a new file entry and return id"""
        sql = """ insert into file_records(record_id, file_id, header_id, value) 
                    values (?, ?, ?, ?) """
        if(conn is not None):
            try:
                c = conn.cursor()
                c.execute(sql, file_record)
                conn.commit()
                return c.lastrowid
            except Error as e:
                print(e)
        return None
    @staticmethod
    def select_all_files(conn):
        """
        Query all rows in the tasks table
        :param conn: the Connection object
        :return:
        """
        c = conn.cursor()
        c.execute("SELECT * FROM files")
 
        rows = c.fetchall()
 
        for row in rows:
            print(row)
    @staticmethod
    def get_file_by_id(conn, file_id):
        """
        Query all rows in the tasks table
        :param conn: the Connection object
        :return:
        """
        c = conn.cursor()
        c.execute("SELECT * FROM files where id = ?", (file_id,))
        row = c.fetchone()
        return (row)
    
    @staticmethod
    def get_nccd_id(conn, parcel):
        """
        Get NCCD ID From Databbase
        """
        c = conn.cursor()
        c.execute("""SELECT a.value FROM file_records a 
        JOIN file_records b ON a.record_id = b.record_id and b.header_id = 7 and a.header_id = 33
        where b.value = ?""", (parcel,))
        row = c.fetchone()
        if(row is None):
            return None
        else:
            return (row[0])

    @staticmethod
    def get_header_id_by_name(conn, header):
        """ return header id by header of header """
        sql = """ select a.id from headers a
                  where a.is_source = 1 and a.header = ?; """
        sql_insert = """ insert into headers(header, is_source) values (?, ?); """ 
        c = conn.cursor()
        c.execute(sql,(header,))
        row = c.fetchone()
        if(row is None):
            c.execute(sql_insert,(header,1,))
            conn.commit()
            return c.lastrowid
        else:
            return row[0]
    @staticmethod
    def get_header_dict(conn, is_source=None):
        sql = """ select a.id, a.header from headers a
                  where a.is_source = ? order by a.id ; """
        sql_all = """ select a.id, a.header from headers a
                      order by a.id ; """
        c = conn.cursor()
        if(is_source is None):
            c.execute(sql_all)
        else:
            c.execute(sql, (is_source,))
        dict = {}
        rows = c.fetchall()
        for row in rows:
            dict[row[0]]=row[1]
        return dict
    @staticmethod
    def get_file_records(conn, record_id_set):
        """ return records of a file """
        sql =   """
                    select record_id, header_id, value 
                    from file_records 
                    where file_id = ? and header_id in (%s)
                    order by record_id, header_id
                """ %','.join('?'*(len(record_id_set)-1))
        c = conn.cursor()
        c.execute(sql,record_id_set)
        rows = c.fetchall()
        return rows
    @staticmethod
    def get_next_sheet_row(sheet, row_i, col_i):
        if(row_i < sheet.nrows-1):
            next_row_i = row_i + 1
            while next_row_i < sheet.nrows:
                cell_value2 = sheet.cell(next_row_i, col_i).value
                if(cell_value2 is not None and cell_value2 != ''):
                    break
                next_row_i = next_row_i + 1
        else:
            next_row_i = row_i + 1
        return next_row_i

    @staticmethod
    def get_cell_value_concatenated(sheet, row_i, col_i, next_row_i, header_dict, sheet_header_mapping, actual_col_i):
        return_val = ''
        row_i2 = row_i

        for row_i2 in range(row_i, next_row_i):
            # for Address and Parcel change numeric number to integer
            if(sheet_header_mapping[actual_col_i] in (6,7) and sheet.cell(row_i2, col_i).ctype == 2):
                cell_value = str(int(sheet.cell(row_i2, col_i).value))
            else:
                cell_value = str(sheet.cell(row_i2, col_i).value).replace(header_dict[sheet_header_mapping[actual_col_i]],'').strip().replace('\n',' ')
            if(cell_value != '' and cell_value is not None and header_dict[sheet_header_mapping[actual_col_i]].find(cell_value) == -1 ):
                if(return_val != ''):
                    return_val = return_val + ' ' + cell_value
                else:
                    return_val = cell_value
            # Cleanup Parcel
            if(sheet_header_mapping[actual_col_i] == 7):
                if(return_val.find('and')==-1):
                    return_val = return_val.replace(' ','')
                    return_val.zfill(10)
        return return_val

    @staticmethod
    def get_info_from_zillow(conn, file_row_seq, file_id, address):
        address_transformed = address.replace(' ', '-')+'-_rb/'
        url_zillow1 = 'https://www.zillow.com/homes/'+address_transformed
        r = requests.get(url_zillow1, headers={'Host': 'www.zillow.com', 'User-Agent':'PostmanRuntime/7.15.2'}, )
        content = r.text
        file_record = (file_row_seq, file_id, 34, content)
        ForclosureProgram.create_file_record(conn, file_record)
        #zp
        zpid = content[content.find('_zpid')-8:content.find('_zpid')]
        #ZEstimate
        if(content.find('"ds-value">') != -1):
            new_content = content[content.find('"ds-value">')+11:]
            ZEstimate = new_content[:new_content.find('</span>')]
            file_record = (file_row_seq, file_id, 10, ZEstimate)
            ForclosureProgram.create_file_record(conn, file_record)
        #Living Area
        if(content.find('<span class="ds-bed-bath-living-area"><span>') != -1):
            new_content = content[content.find('<span class="ds-bed-bath-living-area"><span>')+44:]
            living_area = new_content[:new_content.find('</span>')]
            file_record = (file_row_seq, file_id, 12, living_area)
            ForclosureProgram.create_file_record(conn, file_record)
        # Bed
        if(content.find('<span class="ds-vertical-divider ds-bed-bath-living-area"><span>') != -1):
            new_content = content[content.find('<span class="ds-vertical-divider ds-bed-bath-living-area"><span>')+len('<span class="ds-vertical-divider ds-bed-bath-living-area"><span>'):]
            bed = new_content[:new_content.find('</span>')]
            file_record = (file_row_seq, file_id, 13, bed)
            ForclosureProgram.create_file_record(conn, file_record)
        #Bath room
        if(content.find('<span class="ds-vertical-divider ds-bed-bath-living-area"><span>') != -1):
            new_content = content[content.find('<span class="ds-vertical-divider ds-bed-bath-living-area"><span>')+len('<span class="ds-vertical-divider ds-bed-bath-living-area"><span>'):]
            ba = new_content[:new_content.find('</span>')]
            file_record = (file_row_seq, file_id, 14, ba)
            ForclosureProgram.create_file_record(conn, file_record)
        #Status
    
        if(content.find('<span class="zsg-tooltip-launch_keyword">') != -1):
            new_content = content[content.find('<span class="zsg-tooltip-launch_keyword">')+len('<span class="zsg-tooltip-launch_keyword">'):]
            status = new_content[:new_content.find('</span>')]
            file_record = (file_row_seq, file_id, 15, status)
            ForclosureProgram.create_file_record(conn, file_record)
        #Time on Zillow
        if(content.find('<div class="ds-overview-stat-value">') != -1):
            new_content = content[content.find('<div class="ds-overview-stat-value">')+len('<div class="ds-overview-stat-value">'):]
            status = new_content[:new_content.find('</div>')]
            file_record = (file_row_seq, file_id, 16, status)
            ForclosureProgram.create_file_record(conn, file_record)
        #School1
        if(content.find('<div class="ds-nearby-schools-info-section">') != -1):
            new_content1 = content[content.find('<div class="ds-nearby-schools-info-section">')+len('<div class="ds-nearby-schools-info-section">'):]
            new_content = new_content1[new_content1.find('rel="nofollow noopener noreferrer" target="_blank">')+len('rel="nofollow noopener noreferrer" target="_blank">'):]
            school1_name = new_content[:new_content.find('</a>')]
            file_record = (file_row_seq, file_id, 17, school1_name)
            ForclosureProgram.create_file_record(conn, file_record)

            if(new_content1.find('<span class="ds-school-value ds-body-small">') != -1):
                new_content = new_content1[new_content1.find('<span class="ds-school-value ds-body-small">')+len('<span class="ds-school-value ds-body-small">'):]
                school1_grade = new_content[:new_content.find('</span>')]
                file_record = (file_row_seq, file_id, 18, school1_grade)
                ForclosureProgram.create_file_record(conn, file_record)

            if(new_content1.find('Distance:</span><span class="ds-school-value ds-body-small">') != -1):
                new_content = new_content1[new_content1.find('Distance:</span><span class="ds-school-value ds-body-small">')+len('Distance:</span><span class="ds-school-value ds-body-small">'):]
                school1_distance = new_content[:new_content.find('</span>')]
                file_record = (file_row_seq, file_id, 19, school1_distance)
                ForclosureProgram.create_file_record(conn, file_record)

            if(content.find('<span class="ds-hero-headline ds-schools-display-rating">') != -1):
                new_content_rating = content[content.find('<span class="ds-hero-headline ds-schools-display-rating">')+len('<span class="ds-hero-headline ds-schools-display-rating">'):]
                school1_rating = new_content_rating[:new_content_rating.find('</span>')] + '/10'
                file_record = (file_row_seq, file_id, 20, school1_rating)
                ForclosureProgram.create_file_record(conn, file_record)        
        
            #School2
            if(new_content1.find('<div class="ds-nearby-schools-info-section">') != -1):
                new_content2 = new_content1[new_content1.find('<div class="ds-nearby-schools-info-section">')+len('<div class="ds-nearby-schools-info-section">'):]
                new_content = new_content2[new_content2.find('rel="nofollow noopener noreferrer" target="_blank">')+len('rel="nofollow noopener noreferrer" target="_blank">'):]
                school1_name = new_content[:new_content.find('</a>')]
                file_record = (file_row_seq, file_id, 21, school1_name)
                ForclosureProgram.create_file_record(conn, file_record)

                if(new_content2.find('<span class="ds-school-value ds-body-small">') != -1):
                    new_content = new_content2[new_content2.find('<span class="ds-school-value ds-body-small">')+len('<span class="ds-school-value ds-body-small">'):]
                    school1_grade = new_content[:new_content.find('</span>')]
                    file_record = (file_row_seq, file_id, 22, school1_grade)
                    ForclosureProgram.create_file_record(conn, file_record)

                if(new_content2.find('Distance:</span><span class="ds-school-value ds-body-small">') != -1):
                    new_content = new_content2[new_content2.find('Distance:</span><span class="ds-school-value ds-body-small">')+len('Distance:</span><span class="ds-school-value ds-body-small">'):]
                    school1_distance = new_content[:new_content.find('</span>')]
                    file_record = (file_row_seq, file_id, 23, school1_distance)
                    ForclosureProgram.create_file_record(conn, file_record)

                if(new_content_rating.find('<span class="ds-hero-headline ds-schools-display-rating">') != -1):
                    new_content_rating = new_content_rating[new_content_rating.find('<span class="ds-hero-headline ds-schools-display-rating">')+len('<span class="ds-hero-headline ds-schools-display-rating">'):]
                    school1_rating = new_content_rating[:new_content_rating.find('</span>')] + '/10'
                    file_record = (file_row_seq, file_id, 24, school1_rating)
                    ForclosureProgram.create_file_record(conn, file_record)
        #Price Hisotory
        api_endpoint = 'https://www.zillow.com/graphql/?zpid=' + zpid + '&contactFormRenderParameter=&queryId=37deca2c9f17ef9ff5f2e01f0fd10794&operationName=ForSaleDoubleScrollFullRenderQuery'
        data = {
            "operationName": "ForSaleDoubleScrollFullRenderQuery",
            "variables": {
                "zpid": zpid,
                "contactFormRenderParameter": {
                    "zpid": zpid,
                    "platform": "desktop",
                    "isDoubleScroll": 'true'
                }
            },
            "clientVersion": "home-details/5.45.0.0.0.hotfix-2019-07-31.00d1547",
            "queryId": "37deca2c9f17ef9ff5f2e01f0fd10794"
        }
    
        # Get Price History
        r = requests.post(url=api_endpoint, data=json.dumps(data), headers={'Host': 'www.zillow.com', 'User-Agent':'PostmanRuntime/7.15.2', 'Content-Type':'application/json', 'Content-Length':'388','Postman-Token':'e06d0281-62b6-4254-90df-36896dd6cbaa'},)
        price_history_dataset = r.json()["data"]["property"]["priceHistory"]
        price_history_converted_dataset = []
        if( len(price_history_dataset) > 0):
            for price_record in price_history_dataset:
                price={}
                price['date']=datetime.fromtimestamp((price_record['time']/1000)).strftime('%Y-%m-%d')
                price['price']='$'+str(price_record['price'])
                price['priceChangeRate']=str(round((price_record['priceChangeRate']*100),2))+'%'
                price['event']=price_record['event']
                price['source']=price_record['source']
                price['buyerAgent']=price_record['buyerAgent']
                price['sellerAgent']=price_record['sellerAgent']
                price_history_converted_dataset.append(price)
    
            #Insert Price History
            if(len(price_history_converted_dataset)>0):
                file_record = (file_row_seq, file_id, 25, str(price_history_converted_dataset))
                ForclosureProgram.create_file_record(conn, file_record)
        zestimate = r.json()["data"]["property"]["zestimate"]
        file_record = (file_row_seq, file_id, 35, str(zestimate))
        ForclosureProgram.create_file_record(conn, file_record)

    @staticmethod
    def get_info_from_nccd(conn, file_row_seq, file_id, parcel):
        parcel_id = ForclosureProgram.get_nccd_id(conn, parcel)
        if(parcel_id is None or parcel_id == 'NA'):
            ForclosureProgram.v_dict_questions[parcel]=None
            keep_lookup = True
            counter = 0
            while keep_lookup:
                try:
                    answer = ForclosureProgram.v_dict_answers[parcel]
                    parcel_id = answer[1]
                    keep_lookup = False
                    time.sleep(2)
                except:
                    keep_lookup = True
                    time.sleep(2)
                    counter = counter + 1
                    if(counter==(10 if file_row_seq==1 else 45)):
                        print('')
                        print('Open/Refresh "http://www3.nccde.org/parcel/search/default.aspx" in Chrome with CJS extension and JS Code')
        #Return if no answers found else process answers
        if(parcel_id!='NA' and parcel_id != '' and parcel_id is not None):
            #parcel_id
            file_record = (file_row_seq, file_id, 33, parcel_id)
            ForclosureProgram.create_file_record(conn, file_record)
            url_nccd = 'http://www3.nccde.org/parcel/Details/Default.aspx?ParcelKey='+parcel_id
            r = requests.get(url_nccd, headers={'Host': 'www3.nccde.org', 'User-Agent':'PostmanRuntime/7.15.2',
                                                'Postman-Token':'67e7d762-6217-4a2d-8eb7-6d9f9de7a1d4'}, )
            content = r.text
            file_record = (file_row_seq, file_id, 26, content)
            ForclosureProgram.create_file_record(conn, file_record)
            #Subdivision
            if(content.find('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelSubdivision">') != -1):
                new_content = content[content.find('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelSubdivision">')+len('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelSubdivision">'):]
                NSubdivision = new_content[:new_content.find('</span>')]
                file_record = (file_row_seq, file_id, 27, NSubdivision)
                ForclosureProgram.create_file_record(conn, file_record)
            #Municipal Info
            if(content.find('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelParcelStatus">') != -1):
                new_content = content[content.find('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelParcelStatus">')+len('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelParcelStatus">'):]
                NMunicipal = new_content[:new_content.find('</span>')]
                file_record = (file_row_seq, file_id, 28, NMunicipal)
                ForclosureProgram.create_file_record(conn, file_record)
            #Zoning                
            if(content.find('<h4>Zoning</h4>') != -1):
                new_content = content[content.find('<h4>Zoning</h4>')+len('<h4>Zoning</h4>'):]
                new_content = new_content[:new_content.find('</div>')]
                new_content = new_content.replace('\n','').replace('\r','').replace('\t','').replace('  ','').replace('<ul>','')\
                .replace('<li>','').replace('</li>',' ').replace('</ul>','')
                file_record = (file_row_seq, file_id, 29, new_content)
                ForclosureProgram.create_file_record(conn, file_record)
            #County Balance Due
            if(content.find('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelTaxCountyBalanceDue">') != -1):
                new_content = content[content.find('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelTaxCountyBalanceDue">')+len('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelTaxCountyBalanceDue">'):]
                val = new_content[:new_content.find('</span>')]
                file_record = (file_row_seq, file_id, 30, val)
                ForclosureProgram.create_file_record(conn, file_record)
            #School Balance Due
            if(content.find('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelTaxSchoolBalanceDue">') != -1):
                new_content = content[content.find('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelTaxSchoolBalanceDue">')+len('<span id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__LabelTaxSchoolBalanceDue">'):]
                val = new_content[:new_content.find('</span>')]
                file_record = (file_row_seq, file_id, 31, val)
                ForclosureProgram.create_file_record(conn, file_record)
            #Recent Sale
            url_nccd = 'http://www3.nccde.org/parcel/RecentSales/Default.aspx?ParcelKey='+parcel_id
            r = requests.get(url_nccd, headers={'Host': 'www3.nccde.org', 'User-Agent':'PostmanRuntime/7.15.2',
                                                'Postman-Token':'67e7d762-6217-4a2d-8eb7-6d9f9de7a1d4'}, )
            content = r.text
            recent_sale = []
            for i in range(0,4):
                if(content.find('<tr>') == -1):
                    break
                content=content[content.find('<tr>'):]
                if(i == 0):
                    continue
                content=content[content.find('<tr>'):]
                sale={}
                sale['Parcel #']=content[content.find('</a>')-10:content.find('</a>')]
                content=content[content.find('</td>')+len('</td>'):]
                sale['Address']=content[content.find('<td class="left">')+len('<td class="left">'):content.find('</td>')]
                content=content[content.find('</td>')+len('</td>'):]
                sale['Sale Date']=content[content.find('<td>')+len('<td>'):content.find('</td>')]
                content=content[content.find('</td>')+len('</td>'):]
                sale['Sale Amount']=content[content.find('<td class="last">')+len('<td class="last">'):content.find('</td>')]

                recent_sale.append(sale)
            str_recent_sale = json.dumps(recent_sale)
            file_record = (file_row_seq, file_id, 32, str(str_recent_sale))
            ForclosureProgram.create_file_record(conn, file_record)
    @staticmethod
    def create_excel_file(conn, file_id, sheet_header_mapping, header_dict):
        print('\Composing New Excel File:')
        style_text = 'font: bold 1; pattern: back_color orange; border: top thick, right thick, bottom thick, left thick;'
        style = xlwt.easyxf(style_text)
        file = ForclosureProgram.get_file_by_id(conn, file_id)
        header_dict_computed = ForclosureProgram.get_header_dict(conn, 0)
        last_col_i = list(sheet_header_mapping.keys())[-1]
        # Iterating over values
        for header_id, header_name in header_dict_computed.items():
            last_col_i = last_col_i + 1
            sheet_header_mapping[last_col_i] = header_id

        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Foreclosure')
        header_col_mapping = {}
        for col_i, header_id in sheet_header_mapping.items():
            #write headers to sheet
            sheet.write(0, col_i, header_dict[header_id], style)
            #build new mapping of headers to excel columns
            header_col_mapping[header_id] = col_i
        rows = ForclosureProgram.get_file_records(conn,[file_id] + list(header_col_mapping.keys()))
        #Progress Bar
        bar = progressbar.ProgressBar(maxval=len(rows), \
                    widgets=[progressbar.Bar('=', '[', ']'), ' ', progressbar.Percentage()])
        bar.start()
        # write rows to sheet
        processed_count = 0
        for row in rows:
            sheet.write(row[0], header_col_mapping[row[1]], row[2])
            processed_count = processed_count + 1
            # Update Progress
            bar.update(processed_count)
        bar.update(len(rows))
        new_file_name = ('output/'+file[1].split('.')[0]+'-'+file[2]).replace(':','').replace('.','')+'.xls'
        workbook.save(new_file_name)
    @staticmethod
    def process_files(conn):
        source_directory = "to_be_processed/"
        directory = os.fsencode(source_directory)
        for file in os.listdir(directory):
            file_row_seq = 0
            filename = os.fsdecode(file)
            if filename.endswith(".xlsx") or filename.endswith(".xls"): 
                print('Processing file: ' + filename )
                file = (filename, str(datetime.now()))
                file_id = ForclosureProgram.create_file(conn, file)
                wb = xlrd.open_workbook(source_directory+filename, on_demand=False)
                sheet_count = len(wb.sheet_names())
                sheet_header_mapping = {}
            
                # get all shets count
                total_rows = 0
                for sheet_i in range(0, sheet_count):
                    sheet = wb.sheet_by_index(sheet_i)
                    total_rows = total_rows + sheet.nrows
                #Progress Bar
                print('Preparing data set:')
                bar = progressbar.ProgressBar(maxval=total_rows, \
                    widgets=[progressbar.Bar('=', '[', ']'), ' ', progressbar.Percentage()])
                bar.start()
                prev_sheet_row_i = 0
                header_dict = {}
                for sheet_i in range(0, sheet_count):
                    sheet = wb.sheet_by_index(sheet_i)
                    row_i = 0
                    next_row_i = 0
                    skip_cols = []
                    while row_i < sheet.nrows:
                        for col_i in range(0,sheet.ncols):
                            # Skip the empty columns
                            if(col_i in skip_cols):
                                continue
                            # Calculate actual column By skiping empty columns
                            actual_col_i = col_i
                            for skip_col in skip_cols:
                                if(col_i > skip_col):
                                    actual_col_i = actual_col_i - 1
                            if(col_i == 0):
                                next_row_i = ForclosureProgram.get_next_sheet_row(sheet, row_i, col_i)
                            cell_value = sheet.cell_value(row_i, col_i)
                            # Check if there are skip columns in the sheet
                            if(row_i == 0 and (cell_value == '' or cell_value is None)):
                                skip_cols.append(col_i)
                                continue
                            # For first row of first sheet prepare headers
                            if(row_i == 0 and sheet_i == 0):
                                if(cell_value is not None and cell_value != ''):
                                    header_id = ForclosureProgram.get_header_id_by_name(conn, cell_value)
                                    sheet_header_mapping[col_i] = header_id
                            else:
                                if len(list(header_dict.keys())) == 0:
                                    header_dict = ForclosureProgram.get_header_dict(conn, None)
                                cell_value = sheet.cell_value(row_i, col_i)
                                # Get row data spread till the start of next row
                                cell_value_concatenated = ForclosureProgram.get_cell_value_concatenated(sheet, row_i, col_i, next_row_i, header_dict, sheet_header_mapping, actual_col_i)
                                #if address is read then fetch columns from Zillow
                                if(sheet_header_mapping[actual_col_i]==6 and cell_value_concatenated is not None and cell_value_concatenated !=''):
                                    ForclosureProgram.get_info_from_zillow(conn, file_row_seq, file_id, cell_value_concatenated)
                                    #Add ZIP
                                    file_record = (file_row_seq, file_id, 11, cell_value_concatenated[-6:])
                                    ForclosureProgram.create_file_record(conn, file_record)
                                #if address is read then fetch columns from NCCD
                                if(sheet_header_mapping[actual_col_i]==7):
                                    if(cell_value_concatenated!='' and cell_value_concatenated is not None):
                                        if(cell_value_concatenated.find(' and ') != -1):
                                            parcel = cell_value_concatenated[:cell_value_concatenated.find(' and ')].zfill(10)
                                        else:
                                            parcel = cell_value_concatenated
                                        ForclosureProgram.get_info_from_nccd(conn, file_row_seq, file_id, parcel)
                                # skip headers repeaded
                                if(col_i == 0):
                                    # Ignore headers in other sheets
                                    if(cell_value_concatenated == ''):
                                        break
                                    else:
                                        file_row_seq = file_row_seq + 1
                                # Create file record
                                file_record = (file_row_seq, file_id, sheet_header_mapping[actual_col_i], cell_value_concatenated)
                                ForclosureProgram.create_file_record(conn, file_record)
                        # Update Progress
                        bar.update(prev_sheet_row_i + row_i)
                        row_i = next_row_i
                    prev_sheet_row_i = prev_sheet_row_i + next_row_i - 1
                    # Update Progress
                    bar.update(prev_sheet_row_i)
                # Update Progress
                bar.update(total_rows)
                ForclosureProgram.create_excel_file(conn, file_id, sheet_header_mapping, header_dict)
                os.rename(source_directory+filename,'processed/'+filename)
    @staticmethod
    def decode_frame(frame):
        opcode_and_fin = frame[0]
        # assuming it's masked, hence removing the mask bit(MSB) to get len. also assuming len is <125
        payload_len = frame[1] - 128
        mask = frame [2:6]
        encrypted_payload = frame [6: 6+payload_len]
        payload = bytearray([ encrypted_payload[i] ^ mask[i%4] for i in range(payload_len)])
        return payload.decode('utf-8')
    @staticmethod
    def send_frame(payload1):
        payload = bytearray(payload1.encode('utf-8'))
        # setting fin to 1 and opcpde to 0x1
        frame = [129]
        # adding len. no masking hence not doing +128
        frame += [len(payload)]
        # adding payload
        frame_to_send = bytearray(frame) + payload
        ForclosureProgram.v_client_socket.sendall(frame_to_send)
    @staticmethod
    def fn_keep_listening():
        client_sock = ForclosureProgram.v_client_socket
        live = True
        while live:
            try:
                decoded_payload = ForclosureProgram.decode_frame(bytearray(client_sock.recv(9999).strip()))
                if(decoded_payload == 'PARCEL'):
                    parcel = 'NA'
                    try:
                        parcel = list(ForclosureProgram.v_dict_questions.keys())[0]
                    except:
                        parcel = 'NA'
                    ForclosureProgram.send_frame(str(parcel))
                else:
                    parcel_info = decoded_payload.split("|")
                    if(len(parcel_info)>0):
                        ForclosureProgram.v_dict_questions.pop(parcel_info[0],None)
                        ForclosureProgram.v_dict_answers[parcel_info[0]]=parcel_info
                        ForclosureProgram.send_frame('DONE')
            except ConnectionAbortedError as error:
                client_sock.close()
                live = False
            except ConnectionResetError as error:
                client_sock.close()
                live = False
            except Exception as error:
                client_sock.close()
                live = False
    @staticmethod
    def handle_client_connection(client_sock,address,bind_ip,bind_port):
        client_sock.settimeout(60.0)
        # Do Handshake
        received = client_sock.recv(9924).strip()
        MAGIC = b'258EAFA5-E914-47DA-95CA-C5AB0DC85B11'  # Fix key for handshake on server side
        hsKey = hsUpgrade = b''
        for header in received.split(b'\r\n'):
            if header.startswith(b'Sec-WebSocket-Key'): hsKey = header.split(b':')[1].strip()
            if header.startswith(b'Upgrade'): hsUpgrade = header.split(b':')[1].strip()
        if hsUpgrade != b"websocket": return
        digest = base64.b64encode(bytes.fromhex(hashlib.sha1(hsKey + MAGIC).hexdigest())).decode('utf-8')
        response = ('HTTP/1.1 101 Switching Protocols\r\nUpgrade: websocket\r\n'
            'Connection: Upgrade\r\nSec-WebSocket-Accept: {}\r\n\r\n'.format(digest))
        client_sock.send(response.encode('utf8'))
        #Keep Listening
        listen_handler = threading.Thread(
            target=ForclosureProgram.fn_keep_listening,
            args=()
        )
        listen_handler.start();
    @staticmethod
    def start_tcp_server():
        bind_ip = '127.0.0.1'
        bind_port = 5001
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind((bind_ip, bind_port))
        server.listen(1000)  # max backlog of connections

        #print('Listening on {}:{}'.format(bind_ip,bind_port))
    
        while True:
            client_sock, address = server.accept()
            ForclosureProgram.v_client_socket = client_sock
            #print('Accepted connection from {}:{}'.format(address[0], address[1] ))
            client_handler = threading.Thread(
                target=ForclosureProgram.handle_client_connection,
                args=(client_sock,address,bind_ip,bind_port)
            )
            client_handler.start();
    @staticmethod
    def main():
        client_handler = threading.Thread(
                target=ForclosureProgram.start_tcp_server,
                args=()
            )
        client_handler.start();
        database = "D:\\Foreclosure\db\pythonsqlite.db"
        conn = ForclosureProgram.create_connection(database)
        ForclosureProgram.create_db_schema(conn)
        ForclosureProgram.patch_db_schema(conn)
        ForclosureProgram.process_files(conn)
        conn.close()
        print('')
        print('Execution Completed!')
        os._exit(1)
if __name__ == '__main__':
    ForclosureProgram.main()
