import os

import peewee as pw

database = pw.SqliteDatabase(F'{os.getcwd()}\\Traffic.db', pragmas={'foreign_keys': 4})


class Table(pw.Model):
    class Meta:
        database = database


class Data(Table):
    date = pw.DateField()
    from_ = pw.CharField(max_length=100, default='')
    to_ = pw.CharField(max_length=100, default='')
    distance = pw.IntegerField(default=1)

    class Meta:
        order_by = ('date',)
