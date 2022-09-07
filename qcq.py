import sys
from MySQLdb import Connect, cursors
import xlrd

mysql_settings = {
    'host': '',
    'username': '',
    'password': '',
    'port': 3306,
    'database': ''
}

update_query_template = "UPDATE ipaddresses SET `InstallOnAP`='{1}' WHERE ip_addr=INET_ATON('{0}');\n"


class MySQL:
    """
    Just a simple class for managing the connection
    """
    con = None
    cur = None
    cur_class = None

    def __init__(self, cfg, cur_class=cursors.DictCursor):
        self.cur_class = cur_class
        self.con = Connect(host=cfg.get('host'),
                           user=cfg.get('username'),
                           password=cfg.get('password'),
                           database=cfg.get('database'),
                           port=cfg.get('port'),
                           cursorclass=cur_class,
                           charset='utf8',
                           use_unicode=True)

    def __enter__(self):
        self.cur = self.con.cursor(self.cur_class)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        del exc_type
        del exc_val
        del exc_tb
        self.con.commit()
        self.cur.close()
        self.con.close()


if __name__ == '__main__':
    print("Quick Crappy Query Generator")
    assert len(sys.argv) > 1, "File path is required."
    filepath = sys.argv[1] if len(sys.argv) > 1 else "crap.xlsx"

    test = '--test' in [str(arg).lower() for arg in sys.argv]

    error = '--error' in [str(arg).lower() for arg in sys.argv]

    export = '--export' in [str(arg).lower() for arg in sys.argv]

    if export:
        idx = sys.argv.index('--export') + 1
        assert len(sys.argv) >= idx + 1, 'Export File path not specified'
        export = sys.argv[idx]

    excel_workbook = xlrd.open_workbook(filepath)
    sheet = excel_workbook.sheet_by_index(0)

    assert sheet.nrows > 1, "Not enough rows to process."
    assert sheet.ncols > 10, "Not enough columns."

    ip_column = 2  # Column C
    capacity_column = 11  # Column L

    update_template = update_query_template

    update_query = ""
    for r in range(1, sheet.nrows):
        if update_query == "":
            update_query = "START TRANSACTION;\n"
        try:
            cap = int(sheet.cell_value(r, capacity_column))
            ip = str(sheet.cell_value(r, ip_column))
            ap_status = "Yes"
            if cap < -1:
                ap_status = "No"
            elif cap < 2:
                ap_status = "If Approved"

            update_query += update_template.format(ip, ap_status)
        except:
            if error:
                print(f'ERROR: Row {r} failed to process.')

    if update_query and not test:
        if mysql_settings.get('host') and mysql_settings.get('username') \
                and mysql_settings.get('password') and mysql_settings.get('port') and mysql_settings.get('database'):
            last_q = ""
            try:
                with MySQL(mysql_settings) as db:
                    print("Connected. Updating...")
                    for r in update_query.split("\n"):
                        last_q = r
                        if r:
                            db.cur.execute(r)
                    print(db.cur.rowcount, "record(s) updated")
            except Exception as ex:
                if error:
                    print(f'ERROR: Query {last_q} aborted the update process.')
                raise ex
        else:
            print(f'WARNING: Query execution aborted because mysql connection settings are missing.')
            if not export:
                export = "default.sql"
    elif test:
        print(update_query)

    if export:
        with open(export, 'w') as file:
            file.write(update_query + 'COMMIT;')
