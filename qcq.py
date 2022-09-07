import sys
from MySQLdb import Connect, cursors
import xlrd
import qcq_custom


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
                           cursorclass=cfg.get('cur_class', cur_class),  # custom cursor classes are allowed
                           charset=cfg.get('charset'),
                           use_unicode=cfg.get('use_unicode'))

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
    filepath = sys.argv[1] if len(sys.argv) > 1 else qcq_custom.default_path

    assert filepath, "File path is required."

    test = '--test' in [str(arg).lower() for arg in sys.argv]

    error = '--error' in [str(arg).lower() for arg in sys.argv]

    export = '--export' in [str(arg).lower() for arg in sys.argv]

    if export:
        idx = sys.argv.index('--export') + 1
        assert len(sys.argv) >= idx + 1, 'Export File path not specified'
        export = sys.argv[idx]

    excel_workbook = xlrd.open_workbook(filepath)
    sheet = excel_workbook.sheet_by_index(0)

    assert sheet.nrows > qcq_custom.required_rows, "Not enough rows to process."
    assert sheet.ncols > qcq_custom.required_cols, "Not enough columns."

    query_template = f"{qcq_custom.query_template}\n"

    generated_query = ""
    for r in range(1, sheet.nrows):
        if generated_query == "" and qcq_custom.upsert_query and qcq_custom.enable_transaction:
            generated_query = "START TRANSACTION;\n"
        try:
            generated_query += query_template.format(*qcq_custom.process_row(r, sheet))
        except:
            if error:
                print(f'ERROR: Row {r} failed to process.')

    if generated_query and not test:
        if qcq_custom.mysql_settings.get('enabled') is True and \
                qcq_custom.mysql_settings.get('host') and qcq_custom.mysql_settings.get('username') \
                and qcq_custom.mysql_settings.get('password') and qcq_custom.mysql_settings.get('port') \
                and qcq_custom.mysql_settings.get('database'):
            last_q = ""
            try:
                with MySQL(qcq_custom.mysql_settings) as db:
                    print("Connected. Running generated query...")
                    for r in generated_query.split("\n"):
                        last_q = r
                        if r:
                            db.cur.execute(r)
                    if qcq_custom.upsert_query:
                        print(db.cur.rowcount, "record(s) updated")
            except Exception as ex:
                if error:
                    print(f'ERROR: Query "{last_q}" aborted.')
                raise ex
        else:
            if qcq_custom.mysql_settings.get('enabled') is True:
                print(f'WARNING: Query execution aborted because mysql connection settings are missing.')
            if not export:  # default to export if there are no mysql connection settings, or if disabled
                export = qcq_custom.default_sql_filename
    elif test:
        print(generated_query)

    if export:
        print(f"Exporting to {export}...")
        with open(export, 'w') as file:
            file.write(f'{generated_query}'
                       f'{"COMMIT;" if qcq_custom.upsert_query and qcq_custom.enable_transaction else ""}')
        print("Done.")
