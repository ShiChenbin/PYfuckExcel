import pymysql
import tushare as ts

# 调用Tushare的API
ts.set_token('31be4469d37e6f45027a1590e92c64439840d0dc475a0bd84f8b8be1')
pro = ts.pro_api()

# 连接数据库
conn = pymysql.Connect(
    host='localhost',
    user='root',
    password='123456',
    database='hf'
)

# 创建数据库游标
cursor = conn.cursor()

# 查询stock_info表中的所有股票代码
query = "SELECT ts_code FROM stock_info"
cursor.execute(query)
stock_codes = cursor.fetchall()

# 遍历每个股票代码，查询当日收盘价并更新数据库
for stock_code in stock_codes:
    # 查询当日股票收盘价
    df = pro.daily(ts_code=stock_code[0])
    if not df.empty:
        close_price = df.iloc[0]['close']

        # 更新stock_info表中的close项
        update_query = f"UPDATE stock_info SET close = {close_price} WHERE ts_code = '{stock_code[0]}'"
        cursor.execute(update_query)
        conn.commit()

# 关闭数据库连接
cursor.close()
conn.close()

# 运行结束提示
print("end")