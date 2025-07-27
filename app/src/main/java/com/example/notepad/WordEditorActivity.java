package com.example.notepad;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class WordEditorActivity extends AppCompatActivity {
    private static final int REQUEST_OPEN = 1;
    private static final int REQUEST_SAVE = 2;

    private EditText editText;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_word_editor);

        editText = findViewById(R.id.editTextWord);
        Button btnOpen = findViewById(R.id.btnOpenWord);
        Button btnSave = findViewById(R.id.btnSaveWord);
        Button btnPowerPoint = findViewById(R.id.btnpowerpoint);

        btnPowerPoint.setOnClickListener(v -> {
            Intent intent = new Intent(this, PowerPointEditorActivity.class);
            startActivity(intent);
        });

        btnOpen.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            intent.setType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            startActivityForResult(intent, REQUEST_OPEN);
        });

        btnSave.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
            intent.setType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            intent.putExtra(Intent.EXTRA_TITLE, "document.docx");
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
                readWord(uri);
            } else if (requestCode == REQUEST_SAVE) {
                writeWord(uri);
            }
        }
    }

    private void readWord(Uri uri) {
        try (InputStream inputStream = getContentResolver().openInputStream(uri)) {
            if (inputStream == null) {
                Toast.makeText(this, "Unable to open file", Toast.LENGTH_SHORT).show();
                return;
            }

            XWPFDocument document = new XWPFDocument(inputStream);
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            StringBuilder builder = new StringBuilder();

            for (XWPFParagraph para : paragraphs) {
                builder.append(para.getText()).append("\n");
            }

            editText.setText(builder.toString());
            Toast.makeText(this, "Word document loaded", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Failed to read .docx file: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    private void writeWord(Uri uri) {
        try (OutputStream outputStream = getContentResolver().openOutputStream(uri)) {
            XWPFDocument document = new XWPFDocument();
            String[] lines = editText.getText().toString().split("\n");
            for (String line : lines) {
                document.createParagraph().createRun().setText(line);
            }
            document.write(outputStream);
            document.close();
            Toast.makeText(this, "Word document saved", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Failed to save .docx file", Toast.LENGTH_SHORT).show();
        }
    }
}
