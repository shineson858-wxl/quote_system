# test_db.py
import pymysql

try:
    conn = pymysql.connect(
        host='127.0.0.1',
        port=3306,
        user='root',
        password='',  # XAMPP 默认密码为空
        database='solar_quotation_db',
        charset='utf8mb4'
    )
    cursor = conn.cursor()
    cursor.execute("SHOW TABLES;")
    tables = cursor.fetchall()
    print(" 数据库连接成功！当前存在的表：")
    for t in tables:
        print(f" - {t[0]}")
    cursor.close()
    conn.close()
except Exception as e:
    print("❌ 连接失败，原因：", e)