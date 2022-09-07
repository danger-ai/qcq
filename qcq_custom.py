"""
Example currently configured to generate a query based on a list of IPs and capacity values

All the following variables ARE REQUIRED!
"""

mysql_settings = {
    'host': '',
    'username': '',
    'password': '',
    'port': 3306,
    'database': '',
    'charset': 'utf8',
    'use_unicode': True,
    'enabled': False
}

enable_transaction = True  # only if upsert_query is enabled
upsert_query = True  # if the generated query is treated like an update/insert

required_rows = 1
required_cols = 10

default_path = "crap.xlsx"
default_sql_filename = "default.sql"

query_template = "UPDATE ipaddresses SET `InstallOnAP`='{1}' WHERE ip_addr=INET_ATON('{0}');"


def process_row(row_num, sheet):
    """
    This method can be changed to affect processing of the row
    :param row_num: current row being processed
    :param sheet: the current excel worksheet
    :return: a list or tuple of string values passed back to the query template
    """
    ip_column = 2  # Column C
    capacity_column = 11  # Column L

    cap = int(sheet.cell_value(row_num, capacity_column))
    ip = str(sheet.cell_value(row_num, ip_column))
    ap_status = "Yes"
    if cap < -1:
        ap_status = "No"
    elif cap < 2:
        ap_status = "If Approved"

    return ip, ap_status
