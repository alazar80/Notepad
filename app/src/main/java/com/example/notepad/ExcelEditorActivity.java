package com.example.notepad;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;

public class ExcelEditorActivity extends AppCompatActivity {
    private static final int REQUEST_OPEN = 1;
    private static final int REQUEST_SAVE = 2;

    private EditText editText;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_excel_editor);

        editText = findViewById(R.id.editTextExcel);
        Button btnOpen = findViewById(R.id.btnOpenExcel);
        Button btnSave = findViewById(R.id.btnSaveExcel);

        btnOpen.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            startActivityForResult(intent, REQUEST_OPEN);
        });

        btnSave.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
            intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            intent.putExtra(Intent.EXTRA_TITLE, "sheet.xlsx");
            startActivityForResult(intent, REQUEST_SAVE);
        });
        ((Button)findViewById(R.id.btnClear))
                .setOnClickListener(v -> editText.setText(""));
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (resultCode == RESULT_OK && data != null) {
            Uri uri = data.getData();
            if (requestCode == REQUEST_OPEN) {
                readExcel(uri);
            } else if (requestCode == REQUEST_SAVE) {
                writeExcel(uri);
            }
        }
    }

    private void readExcel(Uri uri) {
        try (InputStream inputStream = getContentResolver().openInputStream(uri)) {
            if (inputStream == null) {
                Toast.makeText(this, "Unable to open file", Toast.LENGTH_SHORT).show();
                return;
            }

            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            StringBuilder builder = new StringBuilder();

            for (Row row : sheet) {
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    builder.append(cell.toString()).append("\t");
                }
                builder.append("\n");
            }

            editText.setText(builder.toString());
            workbook.close();
            Toast.makeText(this, "Excel loaded", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Failed to read Excel: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    private void writeExcel(Uri uri) {
        try (OutputStream outputStream = getContentResolver().openOutputStream(uri)) {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Sheet1");

            String[] lines = editText.getText().toString().split("\n");
            for (int i = 0; i < lines.length; i++) {
                String[] cells = lines[i].split("\t");
                Row row = sheet.createRow(i);
                for (int j = 0; j < cells.length; j++) {
                    row.createCell(j).setCellValue(cells[j]);
                }
            }

            workbook.write(outputStream);
            workbook.close();
            Toast.makeText(this, "Excel saved", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Failed to save Excel", Toast.LENGTH_SHORT).show();
        }
    }
}
