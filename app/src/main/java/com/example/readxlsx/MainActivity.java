package com.example.readxlsx;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

import android.Manifest;
import android.app.Activity;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.Toast;

import com.example.readxlsx.utils.ExcelUtil;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;


public class MainActivity extends AppCompatActivity implements View.OnClickListener {
    Button impo,read_and_save;
    String path = "null";
    List<Map<Integer,Object>> listData = new ArrayList<>();

    int REQD_EXCEL_REQUEST_CODE = 10000;
    int DIR_SELECTOR_CODE = 20000;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        initView();
    }

    private void initView() {
        impo = findViewById(R.id.impo);
        read_and_save = findViewById(R.id.read_and_save);
        setOnClickListener();
    }

    private void setOnClickListener() {
        impo.setOnClickListener(this);
        read_and_save.setOnClickListener(this);
    }

    @Override
    public void onClick(View v) {
        switch (v.getId()) {
            case R.id.impo:
                Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
                intent.setType("application/*");//无类型限制 intent.setType(“video/*;image/*”);同时选择视频和图片
                if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.READ_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
                    ActivityCompat.requestPermissions(MainActivity.this, new String[]{Manifest.permission.READ_EXTERNAL_STORAGE}, 1);
                } else {
                    startActivityForResult(intent, REQD_EXCEL_REQUEST_CODE);
                }
                break;
            case R.id.read_and_save:
                if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
                    ActivityCompat.requestPermissions(MainActivity.this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, 1);
                } else {
                    Intent intent1 = new Intent(Intent.ACTION_CREATE_DOCUMENT);
                    intent1.setType("application/*");
                    intent1.putExtra(Intent.EXTRA_TITLE,System.currentTimeMillis()+".xlsx");
                    startActivityForResult(intent1,DIR_SELECTOR_CODE);
                }
                break;
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == REQD_EXCEL_REQUEST_CODE && resultCode == Activity.RESULT_OK) {
            Uri uri = data.getData();
            path = uri.getPath();
            if (path == null || path.equals("null")) return;
            importExcelDeal(uri);
        }else if (requestCode == DIR_SELECTOR_CODE && resultCode == Activity.RESULT_OK) {
            Uri uri = data.getData();
            if (uri == null) return;
            //您可以修改readExcelList，然后写入excel。
            ExcelUtil excelUtil = new ExcelUtil();
            excelUtil.getDataAndSave(this,uri);
        }
    }

    private void importExcelDeal(final Uri uri) {
        new Thread(new Runnable() {
            @Override
            public void run() {
                if (path.equals(null) || path.equals("null")) return;
                ExcelUtil excelUtil = new ExcelUtil();
                List<Map<Integer, Object>> list = excelUtil.readExcel(MainActivity.this, uri, uri.getPath());
                if (list != null && list.size() > 0){
                    listData.clear();
                    listData.addAll(list);
                    runOnUiThread(new Runnable() {
                        @Override
                        public void run() {
                            Toast.makeText(MainActivity.this, "导入成功", Toast.LENGTH_SHORT).show();
                        }
                    });
                }else
                    runOnUiThread(new Runnable() {
                        @Override
                        public void run() {
                            Toast.makeText(MainActivity.this, "导入失败", Toast.LENGTH_SHORT).show();
                        }
                    });
            }
        }).start();
    }

}