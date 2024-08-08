import docx2txt
import os
from dotenv import load_dotenv
import mysql.connector
from functions import * 

load_dotenv()

#DB Credentials
host_name = os.getenv('HOST_NAME')
user_name = os.getenv('USER')
password = os.getenv('PASSWORD')
db_name = os.getenv('DATABASE')

conn = mysql.connector.connect(
    host=host_name,
    user=user_name,
    password=password,
    database=db_name
)

#Path for word doc 
doc_path = f"C:\\Users\\pagrawal\\Downloads\\docparser\\For-Parth.docx"

#Path for txt file for raw output of tables
tx_path = f"C:\\Users\\pagrawal\\Downloads\\docparser\\doc.txt"

#Create raw output txt file
docx_tables_to_txt(doc_path,tx_path)

# Path to the parsed output file
parsed_output_path = f"C:\\Users\\pagrawal\\Downloads\\docparser\\output.txt"

#parse the raw txt file and store the data correctly
main_table_rows,sub_table_rows = parse_txt_file_test(tx_path)

write_output(parsed_output_path,main_table_rows,sub_table_rows)

insert_intro_trust_services_criteria(doc_path, conn) #intro done

headings = extract_specific_headings(doc_path)
insert_headings_into_db(headings, conn) #table 1 done

insert_into_table_2_and_3(main_table_rows, sub_table_rows, conn) #table 2 and table 3 done

