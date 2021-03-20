package com.example.exceleditor;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;

import static com.example.exceleditor.Excel_Commands.getAccName;
import static com.example.exceleditor.Excel_Commands.getAccNumber;
import static com.example.exceleditor.Excel_Commands.getAmmDeposite;
import static com.example.exceleditor.Excel_Commands.getDate;
import static com.example.exceleditor.Excel_Commands.getDefaults;
import static com.example.exceleditor.Excel_Commands.getNumOfInstal;
import static com.example.exceleditor.Excel_Commands.getReferenceNumber;
import static com.example.exceleditor.Excel_Commands.getTotalAmm;
import static com.example.exceleditor.Excel_Commands.insertRows;
import static com.example.exceleditor.Excel_Commands.setAccNum;
import static com.example.exceleditor.Excel_Commands.setAmmEnd;
import static com.example.exceleditor.Excel_Commands.setDate;
import static com.example.exceleditor.Excel_Commands.setDepsiteAmm;
import static com.example.exceleditor.Excel_Commands.setName;
import static com.example.exceleditor.Excel_Commands.setNumOfInstall;
import static com.example.exceleditor.Excel_Commands.setPenalty;
import static com.example.exceleditor.Excel_Commands.setRdDemon;
import static com.example.exceleditor.Excel_Commands.setRefNum;
import static com.example.exceleditor.Excel_Commands.setRefNumEnd;
import static com.example.exceleditor.Excel_Commands.setRefNumMain;
import static com.example.exceleditor.Excel_Commands.takeFile;

public class MainActivity extends AppCompatActivity {

    private Button browse_button , run;
    private TextView result;
    private String  file_path;



    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);


        Log.d("ABC", "onCreate: "+ ContextCompat.checkSelfPermission(MainActivity.this,Manifest.permission.MANAGE_EXTERNAL_STORAGE)+ " " +PackageManager.PERMISSION_DENIED);
        if(ContextCompat.checkSelfPermission(MainActivity.this,Manifest.permission.READ_EXTERNAL_STORAGE)==
                PackageManager.PERMISSION_DENIED){
            ActivityCompat.requestPermissions(MainActivity.this,new String[]{Manifest.permission.READ_EXTERNAL_STORAGE},50);

        }else {
            Toast
                    .makeText(MainActivity.this,
                            "Permission already granted",
                            Toast.LENGTH_SHORT)
                    .show();
        }


        Log.d("ABC", "onCreate: "+ ContextCompat.checkSelfPermission(MainActivity.this,Manifest.permission.MANAGE_EXTERNAL_STORAGE)+ " " +PackageManager.PERMISSION_DENIED);
        if(ContextCompat.checkSelfPermission(MainActivity.this,Manifest.permission.WRITE_EXTERNAL_STORAGE)==
                PackageManager.PERMISSION_DENIED){
            ActivityCompat.requestPermissions(MainActivity.this,new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE},51);

        }else {
            Toast
                    .makeText(MainActivity.this,
                            "Permission already granted",
                            Toast.LENGTH_SHORT)
                    .show();
        }


        browse_button = findViewById(R.id.Browse_File);
        result = findViewById(R.id.result);
        run = findViewById(R.id.run);

        browse_button.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
                intent.setType("application/vnd.ms-excel");
                startActivityForResult(intent, 2);


            }
        });

        run.setOnClickListener(new View.OnClickListener() {

            @Override
            public void onClick(View v) {

                try {
                    runExcelProcedure();
                } catch (IOException e) {
                    Log.d("ABC", "onCreate: "+e.getMessage());
                    e.printStackTrace();
                }
            }
        });


    }


    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {

        super.onActivityResult(requestCode, resultCode, data);
        switch (requestCode) {
            case 2:
                if (resultCode == RESULT_OK) {
                    file_path = data.getData().getPath();
                    Log.d("abcdef",file_path);
                    result.setText(file_path);
                }
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);

        if(grantResults.length > 0 && grantResults[0]==PackageManager.PERMISSION_GRANTED){
            Toast
                    .makeText(MainActivity.this,
                            "Permission granted",
                            Toast.LENGTH_SHORT)
                    .show();
        }
    }

    public void runExcelProcedure() throws IOException {
        // Intializing the no of lots
        int no = 14;

        // Taking datas Out
        Log.d("ABC",file_path);;
        HSSFSheet initalSheet = takeFile("RDInstallmentReport", this.getAssets().open(file_path));
        String  takeReferenceNo = getReferenceNumber(initalSheet);
        String  takeDate = getDate(initalSheet);
        String[]  takeAccNum = getAccNumber(no , initalSheet);
        String[] takeAccName = getAccName(no , initalSheet);
        String[]  takeAmmDeposite = getAmmDeposite( no , initalSheet);
        String[]  takeNumOfInstall = getNumOfInstal(no , initalSheet);
        String[]  takeDefaults = getDefaults( no , initalSheet);
        String  takeTotalAmm = getTotalAmm(initalSheet , no);

        // Input Stream
        InputStream inp = getResources().openRawResource(R.raw.template);
        HSSFWorkbook finalWorkbook = new HSSFWorkbook(new POIFSFileSystem(inp));
        HSSFSheet finalSheet = finalWorkbook.getSheet("Main");

        //Adding style to wraptext
        CellStyle styleWrap = finalWorkbook.createCellStyle();
        styleWrap.setWrapText(true);

        //Adding Styles to end
        CellStyle style = finalWorkbook.createCellStyle();
        Font fontForBottom = finalWorkbook.createFont();
        Short fontSize = 14 ;
        fontForBottom.setFontHeightInPoints(fontSize);
        fontForBottom.setBold(true);
        style.setFont(fontForBottom);

        //Inserting datas
        setDate(finalSheet , takeDate);
        setRefNumMain(finalSheet , takeReferenceNo);
        insertRows(no , finalSheet);
        setRefNum(finalSheet , takeReferenceNo ,no);
        setAccNum(finalSheet , no , takeAccNum );
        setName(finalSheet , no , takeAccName ,styleWrap);
        setRdDemon(finalSheet , no , takeAmmDeposite);
        setDepsiteAmm(finalSheet , no , takeAmmDeposite);
        setNumOfInstall(finalSheet , no , takeNumOfInstall);
        setPenalty(finalSheet , no , takeDefaults);
        setRefNumEnd(finalSheet , takeReferenceNo , no , style);
        setAmmEnd(finalSheet , takeTotalAmm , no , style);
        System.out.println(takeTotalAmm);
        System.out.println(takeReferenceNo);

        // Final Output
        FileOutputStream out = new FileOutputStream(takeReferenceNo+".xls");
        finalWorkbook.write(out);
    }


}