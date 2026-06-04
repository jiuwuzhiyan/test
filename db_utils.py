import sqlite3
import pandas as pd
import os
import logging
from datetime import datetime
from contextlib import contextmanager

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class PointsDatabase:
    """
    积分数据SQLite数据库操作类
    提供高效的数据库连接、查询、插入和更新操作
    """
    
    def __init__(self, db_path):
        """
        初始化数据库连接
        :param db_path: SQLite数据库文件路径
        """
        self.db_path = db_path
        self._ensure_directory_exists()
    
    def _ensure_directory_exists(self):
        """确保数据库目录存在"""
        dir_path = os.path.dirname(self.db_path)
        if dir_path and not os.path.exists(dir_path):
            os.makedirs(dir_path, exist_ok=True)
    
    @contextmanager
    def get_connection(self, read_only=False):
        """
        获取数据库连接上下文管理器
        :param read_only: 是否只读模式
        :yield: sqlite3.Connection对象
        """
        conn = None
        try:
            if read_only:
                conn = sqlite3.connect(f'file:{self.db_path}?mode=ro', uri=True)
                # 只读模式下只设置cache_size
                conn.execute('PRAGMA cache_size=-2000;')
            else:
                conn = sqlite3.connect(self.db_path)
                conn.execute('PRAGMA journal_mode=WAL;')
                conn.execute('PRAGMA synchronous=NORMAL;')
                conn.execute('PRAGMA cache_size=-2000;')
            yield conn
        except sqlite3.Error as e:
            logger.error(f"数据库连接错误: {e}")
            raise
        finally:
            if conn:
                conn.close()
    
    def table_exists(self, table_name='积分列表'):
        """
        检查表是否存在
        :param table_name: 表名
        :return: bool
        """
        try:
            with self.get_connection(read_only=True) as conn:
                cursor = conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
                    (table_name,)
                )
                return cursor.fetchone() is not None
        except Exception as e:
            logger.error(f"检查表存在性失败: {e}")
            return False
    
    def create_table(self, table_name='积分列表'):
        """
        创建积分列表表
        :param table_name: 表名
        """
        create_sql = f"""
        CREATE TABLE IF NOT EXISTS {table_name} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            积分时间 TIMESTAMP,
            活动项目 TEXT,
            业态 TEXT,
            租户 TEXT,
            会员卡号 INTEGER,
            手机号 TEXT,
            性别 TEXT,
            等级 TEXT,
            销售单号 TEXT,
            原订单号 TEXT,
            关联券订单号 TEXT,
            消费金额 TEXT,
            消费时间 TEXT,
            积分数 REAL,
            积分方式 TEXT,
            积分类型 TEXT,
            渠道 TEXT,
            创建时间 TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS idx_points_time ON {table_name}(积分时间);
        CREATE INDEX IF NOT EXISTS idx_member_card ON {table_name}(会员卡号);
        CREATE INDEX IF NOT EXISTS idx_points_type ON {table_name}(积分类型);
        """
        try:
            with self.get_connection() as conn:
                conn.executescript(create_sql)
                conn.commit()
            logger.info(f"表 '{table_name}' 创建成功")
        except sqlite3.Error as e:
            logger.error(f"创建表失败: {e}")
            raise
    
    def get_max_date(self, table_name='积分列表'):
        """
        获取最大积分时间
        :param table_name: 表名
        :return: datetime对象或None
        """
        if not self.table_exists(table_name):
            return None
        
        try:
            with self.get_connection(read_only=True) as conn:
                cursor = conn.execute(
                    f"SELECT MAX(积分时间) FROM {table_name}"
                )
                result = cursor.fetchone()
                if result and result[0]:
                    return datetime.strptime(result[0], '%Y-%m-%d %H:%M:%S')
                return None
        except sqlite3.Error as e:
            logger.error(f"获取最大日期失败: {e}")
            return None
    
    def query_data(self, start_date=None, end_date=None, table_name='积分列表', 
                   limit=None, offset=None):
        """
        查询积分数据
        :param start_date: 开始日期 (datetime或字符串)
        :param end_date: 结束日期 (datetime或字符串)
        :param table_name: 表名
        :param limit: 返回行数限制
        :param offset: 偏移量
        :return: pandas.DataFrame
        """
        if not self.table_exists(table_name):
            return pd.DataFrame()
        
        query = f"SELECT * FROM {table_name}"
        params = []
        
        conditions = []
        if start_date:
            if isinstance(start_date, datetime):
                start_date = start_date.strftime('%Y-%m-%d %H:%M:%S')
            conditions.append("积分时间 >= ?")
            params.append(start_date)
        
        if end_date:
            if isinstance(end_date, datetime):
                end_date = end_date.strftime('%Y-%m-%d %H:%M:%S')
            conditions.append("积分时间 <= ?")
            params.append(end_date)
        
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        
        query += " ORDER BY 积分时间 DESC"
        
        if limit:
            query += " LIMIT ?"
            params.append(limit)
        
        if offset:
            query += " OFFSET ?"
            params.append(offset)
        
        try:
            with self.get_connection(read_only=True) as conn:
                df = pd.read_sql_query(query, conn, params=params)
            logger.info(f"查询成功，返回 {len(df)} 条记录")
            return df
        except sqlite3.Error as e:
            logger.error(f"查询数据失败: {e}")
            return pd.DataFrame()
    
    def get_total_count(self, table_name='积分列表'):
        """
        获取总记录数
        :param table_name: 表名
        :return: int
        """
        if not self.table_exists(table_name):
            return 0
        
        try:
            with self.get_connection(read_only=True) as conn:
                cursor = conn.execute(f"SELECT COUNT(*) FROM {table_name}")
                result = cursor.fetchone()
                return result[0] if result else 0
        except sqlite3.Error as e:
            logger.error(f"获取记录数失败: {e}")
            return 0
    
    def insert_data(self, df, table_name='积分列表', batch_size=1000):
        """
        批量插入数据
        :param df: pandas.DataFrame
        :param table_name: 表名
        :param batch_size: 批量插入大小
        :return: 插入的记录数
        """
        if df.empty:
            return 0
        
        # 确保表存在
        if not self.table_exists(table_name):
            self.create_table(table_name)
        
        # 清理DataFrame中的不必要列和空值
        df = df.drop(columns=['id', '创建时间'], errors='ignore')
        
        # 列名映射：处理可能的列名差异
        column_mapping = {
            '消费金额(元)': '消费金额',
            '消费金额': '消费金额'
        }
        
        # 获取数据库中的实际列名
        with self.get_connection(read_only=True) as conn:
            cursor = conn.execute(f"PRAGMA table_info({table_name})")
            db_columns = [col[1] for col in cursor.fetchall()]
        
        # 调整DataFrame的列名以匹配数据库
        new_columns = []
        for col in df.columns:
            if col in column_mapping and column_mapping[col] in db_columns:
                new_columns.append(column_mapping[col])
            else:
                new_columns.append(col)
        df.columns = new_columns
        
        # 只保留数据库中存在的列
        df = df[[col for col in df.columns if col in db_columns]]
        
        # 将datetime转换为字符串
        for col in df.columns:
            if df[col].dtype == 'datetime64[ns]':
                df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        columns = df.columns.tolist()
        placeholders = ','.join(['?' for _ in columns])
        # 列名可能包含特殊字符，需要用引号括起来
        quoted_columns = [f'"{col}"' for col in columns]
        insert_sql = f"""
        INSERT INTO {table_name} ({','.join(quoted_columns)})
        VALUES ({placeholders})
        """
        
        total_inserted = 0
        
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                # 批量插入
                for i in range(0, len(df), batch_size):
                    batch = df.iloc[i:i+batch_size]
                    data = [tuple(row) for row in batch.itertuples(index=False)]
                    cursor.executemany(insert_sql, data)
                    conn.commit()
                    total_inserted += len(data)
                    logger.debug(f"已插入 {total_inserted} 条记录")
                
                logger.info(f"批量插入完成，共插入 {total_inserted} 条记录")
                return total_inserted
        except sqlite3.Error as e:
            logger.error(f"插入数据失败: {e}")
            raise
    
    def delete_data(self, start_date=None, end_date=None, table_name='积分列表'):
        """
        删除指定日期范围内的数据
        :param start_date: 开始日期
        :param end_date: 结束日期
        :param table_name: 表名
        :return: 删除的记录数
        """
        if not self.table_exists(table_name):
            return 0
        
        query = f"DELETE FROM {table_name}"
        params = []
        conditions = []
        
        if start_date:
            if isinstance(start_date, datetime):
                start_date = start_date.strftime('%Y-%m-%d %H:%M:%S')
            conditions.append("积分时间 >= ?")
            params.append(start_date)
        
        if end_date:
            if isinstance(end_date, datetime):
                end_date = end_date.strftime('%Y-%m-%d %H:%M:%S')
            conditions.append("积分时间 <= ?")
            params.append(end_date)
        
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        else:
            logger.warning("未指定删除条件，取消操作")
            return 0
        
        try:
            with self.get_connection() as conn:
                cursor = conn.execute(query, params)
                conn.commit()
                deleted = cursor.rowcount
                logger.info(f"删除完成，共删除 {deleted} 条记录")
                return deleted
        except sqlite3.Error as e:
            logger.error(f"删除数据失败: {e}")
            raise
    
    def get_data_summary(self, table_name='积分列表'):
        """
        获取数据摘要统计
        :param table_name: 表名
        :return: dict
        """
        if not self.table_exists(table_name):
            return {
                'total_records': 0,
                'min_date': None,
                'max_date': None,
                'avg_points': 0,
                'total_points': 0
            }
        
        try:
            with self.get_connection(read_only=True) as conn:
                cursor = conn.execute(f"""
                    SELECT 
                        COUNT(*),
                        MIN(积分时间),
                        MAX(积分时间),
                        AVG(积分数),
                        SUM(积分数)
                    FROM {table_name}
                """)
                result = cursor.fetchone()
                
                return {
                    'total_records': result[0],
                    'min_date': result[1],
                    'max_date': result[2],
                    'avg_points': round(result[3], 2) if result[3] else 0,
                    'total_points': round(result[4], 2) if result[4] else 0
                }
        except sqlite3.Error as e:
            logger.error(f"获取数据摘要失败: {e}")
            return {}

def get_points_db(db_path=None):
    """
    获取积分数据库实例的便捷函数
    :param db_path: 数据库路径，默认为默认路径
    :return: PointsDatabase实例
    """
    if db_path is None:
        _LOCAL_DIR = r"C:\Users\lenovo\Desktop\财务共享资料-测试环境"
        _UNC_DIR = r"\\172.16.103.130\Users\lenovo\Desktop\财务共享资料-测试环境"
        
        if os.path.exists(_LOCAL_DIR):
            BASE_DIR = _LOCAL_DIR
        elif os.path.exists(_UNC_DIR):
            BASE_DIR = _UNC_DIR
        else:
            BASE_DIR = _LOCAL_DIR
        
        db_path = os.path.join(BASE_DIR, "4-报表中心", "积分列表.db")
    
    return PointsDatabase(db_path)

if __name__ == "__main__":
    # 测试数据库操作
    db = get_points_db()
    
    print("=== 数据库摘要 ===")
    summary = db.get_data_summary()
    for key, value in summary.items():
        print(f"{key}: {value}")
    
    print("\n=== 查询最近10条记录 ===")
    df = db.query_data(limit=10)
    print(df[['积分时间', '会员卡号', '积分数', '积分类型']])
