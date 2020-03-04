package com.example.readxlsx.sqlite;

import android.content.Context;
import android.database.sqlite.SQLiteDatabase;
import android.database.sqlite.SQLiteOpenHelper;

import androidx.annotation.Nullable;

public class MySQLHelp extends SQLiteOpenHelper {
    private String module = "create table module(" +
            "id Integer PRIMARY KEY AUTOINCREMENT,"+
            "materialID Integer," +//物料ID
            "materialEncoding text," +//物料编码
            "materialName text," +//名称
            "materialModel text," +//编号
            "materialSize text," +//规格
            "unit text," +//单位
            "price Integer," +//单价
            "count Integer," +//数量
            "manufacturers text," +//厂家
            "type text," +//类别
            "receiptor text," +//经手人
            "storagelocation text," +//存放地点
            "materialState text);";//状态

    public MySQLHelp(@Nullable Context context, @Nullable String name, @Nullable SQLiteDatabase.CursorFactory factory, int version) {
        super(context, name, factory, version);
    }

    @Override
    public void onCreate(SQLiteDatabase db) {
        db.execSQL(module);
    }

    @Override
    public void onUpgrade(SQLiteDatabase db, int oldVersion, int newVersion) {

    }
}
