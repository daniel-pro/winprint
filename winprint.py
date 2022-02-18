import tempfile
import win32api
import win32print
import cx_Oracle
import sys
import locale
import time

 
def print_on_default_printer(selling_list):
    locale.setlocale(locale.LC_MONETARY, 'it_IT.utf8')
    receipt_id   = str(selling_list[0][6])
    receipt_date = str(selling_list[0][1].date()) 
    
    filename = tempfile.mktemp (receipt_id+".txt")
    f = open (filename, "w")
    f.write ("           PRIMA HOME\n\n")
    f.write("  Data: {}\n".format(receipt_date))
    f.write("  Vendita numero: {}\n\n\n".format(receipt_id))

    f.write("----------------------------------\n")
    f.write("Art.         Qta   Prezzo     Tot.\n")
    total = 0
    for row in selling_list:
        f.write("{:<10} {:>4} {:>8} {:>8}\n".format(row[0][:10],row[2],locale.currency(row[3], symbol=False, grouping=True),locale.currency(row[7], symbol=False, grouping=True)))
        total = total + row[7]
    f.write("----------------------------------\n\n\n")

    f.write("{:25}€{:>7}\n".format("TOTALE:",locale.currency(total, symbol=False, grouping=True)))
    f.write("\n\n");
    f.write(".")
    f.close()

    win32api.ShellExecute (
      0,
      "print",
      filename,
      #
      # If this is None, the default printer will
      # be used anyway.
      #
      '/d:"%s"' % win32print.GetDefaultPrinter (),
      ".",
      0
    )

def get_receipt_and_print(receipt_id, connection):
    query = """select 
           PQ.ID,
           PQ.TX_DATE,
           PQ.QTY,
           PQ.SELLING_PRICE,
           PQ.COST_PRICE,
           PQ.MOVEMENT,
           PQ.SELLING_GROUP_ID,
           PQ.TOTAL_S_AMOUNT,
           PQ.TOTAL_C_AMOUNT,
           P.DESCRIPTION
      from PRODBA.PRODUCT_QTY PQ, PRODBA.PRODUCT P
     where PQ.MOVEMENT='S' AND PQ.SELLING_GROUP_ID = {} AND PQ.ID != '0xF' AND P.ID=PQ.ID order by PQ.TS""".format(str(receipt_id))

    try:
        cur = connection.cursor()
        cur.execute(query)
        selling_list = cur.fetchall()
        cur.close()
        print_on_default_printer(selling_list)
    except Exception as err:
        print("Whoops! [get_receipt_and_print] ERROR: {}".format(str(err)))
        print(err);


def get_receipt_to_print(connection):
    query = """select 
               SELLING_GROUP_ID
      from PRODBA.PRINT"""
    try:
        cur = connection.cursor()
        cur.execute(query)
        printing_list = cur.fetchall()

        for row in printing_list:
            print("\nPrinting Receipt nbr: "+str(row[0]))
            get_receipt_and_print(row[0], connection)
            delete = "delete from PRODBA.PRINT where SELLING_GROUP_ID="+str(row[0])
            delete_cur = connection.cursor()
            delete_cur.execute(delete)
            connection.commit()
            
        cur.close()

    except Exception as err:
        print("Whoops! [get_receipt_to_print] ERROR: {}".format(str(err)))
        print(err);

def open_db_connection():
    try:
        connection = cx_Oracle.connect(user="printer", password="XXXXXXXXXXXXXXXX",dsn="prodb_tp", encoding="UTF-8")
    except Exception as err:
        print("Whoops! ERROR: {}".format(str(err)))
        print(err);    
    return connection

print("Initializing Oracle Client libraries and config ...",end='')
cx_Oracle.init_oracle_client(lib_dir=r"d:/oracle/ic",config_dir=r"d:/oracle/ic/network/admin")
print(" [OK] ")
while True:
    print("Connecting to the Oracle database ...",end='')
    conn = open_db_connection()
    print(" [OK] ")
    print("Getting receipt from the database ...",end='')
    get_receipt_to_print(conn)
    print(" [OK] ")
    print("Closing Oracle connection ...",end='')
    conn.close()
    print(" [OK] ")
    time.sleep(2.5)
    
