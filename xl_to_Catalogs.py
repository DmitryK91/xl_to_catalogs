import sys
import win32com.client
import pymysql.cursors
import string


Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(u'D:\\Users\\DKondratev\\Desktop\\form.xlsx')
ws = wb.ActiveSheet
mainRng = ws.UsedRange.Cells
RowCount = mainRng.Rows.Count
xRow = RowCount

connection = pymysql.connect(host='192.168.169.94', # dev
                            user='root',
                            password='2wsx2WSX',
                            db='morris_prod1',
                            charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)

#родитель активность - группировка (Аудит)
pId = 2342
pName = 'ft_qq_audit_a'


def SQL():
    Rows = Read()

    a = 0
    for row in Rows:
        if a == 0: #пропускаем заголовок
            a += 1
            continue

        parentID = pId
        parentName = pName
        parentTitle = ''
        i = 0

        for col in row:
            title = col[:64]
            effective_behavior = 'NULL'
            if i == 2:
                effective_behavior =  "'%s'" % col

            par = getParent(parentTitle, parentID)
            if par != {}:
                parentName = par['parentName']
                parentID = par['parentID']

            order = getOrder(parentID)

            #сиснеймы для чаптеров, дальше без них
            if i == 0:
                sysname =  "'%s_c%d'" % (parentName, order)
            else:
                sysname = 'NULL'

            INSERT(title, parentID, order, sysname, effective_behavior)

            parentTitle = title

            i += 1


def INSERT(title, parentID, order, sysname, effective_behavior):
    sql = "INSERT INTO tbl_catalog_activity\
            (created, has_empty_answer, id_parent, `order`, title, system_name, effective_behavior)\
            VALUES(SYSDATE(), 1, %d, %d, '%s', %s, %s)" % (parentID, order, title, sysname, effective_behavior)

    if getParent(title, parentID) == {}:
        UPDATE(sql)


def UPDATE(sql):
    try:
        print(sql)
        with connection.cursor() as cursor:
            cursor.execute(sql)
        connection.commit()
    except Exception as ex:
        connection.rollback()
        print(str(ex))

def setHelp_effective():
    """конкатерация effective_behavior с 3го уровня на 2й в help_effective"""

    sqlHelp = "SELECT a1.id_catalog_activity,\
                GROUP_CONCAT(concat(' - ', a2.effective_behavior) SEPARATOR '\n' ) as help_effective\
	            FROM tbl_catalog_activity a\
	            JOIN tbl_catalog_activity a1 ON a1.id_parent = a.id_catalog_activity\
	            JOIN tbl_catalog_activity a2 ON a2.id_parent = a1.id_catalog_activity\
	            WHERE a.id_parent = %s\
	            GROUP BY a1.id_catalog_activity" % pId

    rows = []
    with connection.cursor() as cursor:
        cursor.execute(sqlHelp)
        rows = cursor.fetchall()

    for row in rows:
        sqlUpdate = "UPDATE tbl_catalog_activity SET help_effective = '%s'\
                WHERE id_catalog_activity = %d" % (row['help_effective'], row['id_catalog_activity'])
        UPDATE(sqlUpdate)


def getParent(c, parentID):
    parentSql = "SELECT id_catalog_activity, system_name FROM tbl_catalog_activity a \
                    WHERE a.title = '%s' AND a.id_parent = %d" % (c, parentID)
    par = SELECT(parentSql)
    parent = {}
    if par:
        parent.update({
            'parentID': par['id_catalog_activity'],
            'parentName': par['system_name']
            })
    return parent


def getOrder(parentID):
    """количество дочерних записей"""

    orderSql = "SELECT Count(*) FROM tbl_catalog_activity a WHERE a.id_parent = %d" % parentID
    order = SELECT(orderSql)['Count(*)'] + 1
    return order


def SELECT(sql):
    with connection.cursor() as cursor:
        cursor.execute(sql)
        return cursor.fetchone()


def Read():
    print('Reading Excel:')
    rows = []
    for row in mainRng.Rows:
        cols = []
        for col in row.Columns:
            if not col.MergeArea[0].Value:
                continue
            cols.append(col.MergeArea[0].Value)

        if len(cols) == 3:
            rows.append(cols)

    print(str(len(rows)) + ' rows Found!')
    return rows


if __name__ == '__main__':
    try:
        SQL()
        setHelp_effective()
    except Exception as ex:
        print(str(ex))
    finally:
        connection.close()
        wb.Close()
        Excel.Quit()
        del Excel
        sys.exit(1)