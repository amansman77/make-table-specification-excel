import pymysql
import openpyxl
import shutil
import datetime
import config

def showDatabases(cur):
    sql ='''
    SHOW DATABASES
    '''

    cur.execute(sql)
    rows = cur.fetchall()

    print('\tTable count: {:,}'.format(len(rows)))
    return rows

def findTable(cur, database_name):
    sql ='''
    SELECT TABLE_NAME, TABLE_COMMENT
    FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_SCHEMA = %s
    '''

    cur.execute(sql, [database_name])
    rows = cur.fetchall()

    print('\tTable count: {:,}'.format(len(rows)))
    return rows

def findColumn(cur, database_name, table_name):
    sql ='''
    SELECT ORDINAL_POSITION, COLUMN_NAME, DATA_TYPE, IFNULL(NUMERIC_PRECISION, CHARACTER_MAXIMUM_LENGTH) AS DATA_LENGTH, IS_NULLABLE, COLUMN_KEY, COLUMN_COMMENT, EXTRA
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s
    ORDER BY ORDINAL_POSITION ASC
    '''

    cur.execute(sql, [database_name, table_name])
    rows = cur.fetchall()

    print('\tColumn count: {:,}'.format(len(rows)))
    return rows
    
def addSheet(wb, table_row, column_rows):
    ws = wb.copy_worksheet(wb[cell_config['template_sheet_name']])
    ws.title = table_row['TABLE_NAME']
    ws[cell_config['TABLE_ENGLISH_NAME']] = table_row['TABLE_NAME']
    ws[cell_config['TABLE_KOREAN_COMMENT']] = table_row['TABLE_COMMENT']
    ws[cell_config['TABLE_COMMENT']] = table_row['TABLE_COMMENT']

    start_column_idx = cell_config['START_COLUMN_INDEX']
    # insert row의 개념이 이하 row를 move하는 개념인데 move시에 병합된 cell이 깨지는 현상 존재
    # ws.insert_rows(start_column_idx + 1, len(column_rows))
    for i, column_row in enumerate(column_rows):
        row_index = start_column_idx + i
        ws[cell_config['COLUMN_NUMBER'] + str(row_index)] = column_row['ORDINAL_POSITION']
        ws[cell_config['COLUMN_NAME'] + str(row_index)] = column_row['COLUMN_NAME']
        ws[cell_config['DATA_TYPE'] + str(row_index)] = column_row['DATA_TYPE']
        ws[cell_config['DATA_LENGTH'] + str(row_index)] = column_row['DATA_LENGTH']
        ws[cell_config['IS_NULLABLE'] + str(row_index)] = column_row['IS_NULLABLE']
        ws[cell_config['COLUMN_KEY'] + str(row_index)] = column_row['COLUMN_KEY']
        ws[cell_config['EXTRA'] + str(row_index)] = column_row['EXTRA']

if __name__ == '__main__':
    file_config = config.FILE
    database_config = config.DATABASE
    cell_config = config.CELL_INFO

    file_path = file_config['template_full_path']
    file_path_output = file_path[:file_path.rindex('/')] + file_config['file_name']

    today = datetime.datetime.now()
    create_date = today.strftime('%Y%m%d')
    file_path_output = file_path_output.replace('{{date}}', create_date)

    database_name = database_config['database_name']
    file_path_output = file_path_output.replace('{{database_name}}', database_name)

    shutil.copyfile(file_path, file_path_output)

    connection = pymysql.connect(
        host=database_config['host'],
        port=database_config['port'],
        user=database_config['user'],
        passwd=database_config['passwd'],
        database=database_name,
        charset=database_config['charset']
    )
    cursor = connection.cursor(pymysql.cursors.DictCursor)

    table_rows = findTable(cursor, database_name)

    wb = openpyxl.load_workbook(filename = file_path_output)
    for table_row in table_rows:
        column_rows = findColumn(cursor, database_name, table_row['TABLE_NAME'])
        print(table_row)
        print(column_rows)

        addSheet(wb, table_row, column_rows)

    wb.remove(wb[cell_config['template_sheet_name']])
    wb.save(file_path_output)
    