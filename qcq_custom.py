mysql_settings = {
    'host': '',
    'username': '',
    'password': '',
    'port': 3306,
    'database': ''
}

enable_transaction = True

required_rows = 1
required_cols = 10

update_query_template = "UPDATE ipaddresses SET `InstallOnAP`='{1}' WHERE ip_addr=INET_ATON('{0}');\n"


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