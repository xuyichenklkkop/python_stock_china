# -*- coding = utf-8 -*-
# @Time : 2021/12/18 12:37
# @File : mysql_helper.py
# @Software : PyCharm

import pymysql
from timeit import default_timer
from dbutils.pooled_db import PooledDB


class MySqlConfig:

    def __init__(self, host, db, user, password, port=3306):
        self.host = host
        self.port = port
        self.db = db
        self.user = user
        self.password = password

        self.charset = 'UTF8'  # 不能是 utf-8
        self.minCached = 10
        self.maxCached = 20
        self.maxShared = 10
        self.maxConnection = 100

        self.blocking = True
        self.maxUsage = 100
        self.setSession = None
        self.reset = True


class MysqlPoolConn:
    __pool = None

    def __init__(self, config):
        if not self.__pool:
            self.__class__.__pool = PooledDB(
                creator=pymysql,
                maxconnections=config.maxConnection,
                mincached=config.minCached,
                maxcached=config.maxCached,
                maxshared=config.maxShared,
                blocking=config.blocking,
                maxusage=config.maxUsage,
                setsession=config.setSession,
                charset=config.charset,
                host=config.host,
                port=config.port,
                database=config.db,
                user=config.user,
                password=config.password
            )

    def get_conn(self):
        return self.__pool.connection()


host = 'localhost'
user = 'root'
pwd = 'jtt0304xu'
port = 3306
db = 'normal_stock'
charset = 'utf8'

db_config = MySqlConfig(host, db, user, pwd, port)
g_pool_connection = MysqlPoolConn(db_config)


class MySqlHelper(object):
    def __init__(self, commit=True, logtime=True, log_label="耗时"):
        """
        :param commit:是否在最后提交事务
        :param logtime:是否打印程序运行总时间
        :param log_label:自定义log的文字
        """
        self._log_time = logtime
        self._commit = commit
        self._log_label = log_label

    def __enter__(self):
        # 如果需要记录时间
        if self._log_time is True:
            self._start = default_timer()

        # 从连接池获取数据库连接
        conn = g_pool_connection.get_conn()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        conn.autocommit = False

        # # 在进入的时候自动获取连接和cursor
        # conn = get_connection()
        # cursor = conn.cursor(pymysql.cursors.DictCursor)
        # conn.autocommit = False

        self._conn = conn
        self._cursor = cursor
        return self

    def __exit__(self, *exc_info):
        # 提交事务
        if self._commit:
            self._conn.commit()

        # 在退出的时候自动关闭连接和cursor
        self._cursor.close()
        self._conn.close()

        if self._log_time is True:
            diff = default_timer() - self._start
            print('-- %s: %.6f 秒' % (self._log_label, diff))

    @property
    def cursor(self):
        return self._cursor

    def fetch_all(self, sql, params=None):
        self.cursor.execute(sql, params)
        return self.cursor.fetchall()

    def fetch_by_pk(self, sql, pk):
        self.cursor.execute(sql, (pk,))
        return self.cursor.fetchall()

    def update_by_pk(self, sql, params=None):
        self.cursor.execute(sql, params)