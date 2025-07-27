package com.example.notepad;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;


public class PowerPointEditorActivity extends AppCompatActivity {
    private static final int REQUEST_OPEN = 1;
    private static final int REQUEST_SAVE = 2;

    private EditText editText;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_powerpoint_editor);

        editText = findViewById(R.id.editTextPpt);
        Button btnOpen = findViewById(R.id.btnOpenPpt);
        Button btnSave = findViewById(R.id.btnSavePpt);
        Button btnExcel = findViewById(R.id.btnexcel);

        btnExcel.setOnClickListener(v -> {
            Intent intent = new Intent(this, ExcelEditorActivity.class);
            startActivity(intent);
        });

        btnOpen.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            intent.setType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            startActivityForResult(intent, REQUEST_OPEN);
        });

        btnSave.setOnClickListener(v -> {
            Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
            intent.setType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            intent.putExtra(Intent.EXTRA_TITLE, "presentation.pptx");
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
                readPowerPoint(uri);
            } else if (requestCode == REQUEST_SAVE) {
                writePowerPoint(uri);
            }
        }
    }

    private void readPowerPoint(Uri uri) {
        try (InputStream inputStream = getContentResolver().openInputStream(uri)) {
            if (inputStream == null) {
                Toast.makeText(this, "Unable to open file", Toast.LENGTH_SHORT).show();
                return;
            }

            XMLSlideShow ppt = new XMLSlideShow(inputStream);
            List<XSLFSlide> slides = ppt.getSlides();
            StringBuilder builder = new StringBuilder();

            for (XSLFSlide slide : slides) {
                slide.getShapes().forEach(shape -> {
                    if (shape instanceof XSLFTextShape) {
                        builder.append(((XSLFTextShape) shape).getText()).append("\n");
                    }
                });
                builder.append("---- Slide End ----\n");
            }

            editText.setText(builder.toString());
            Toast.makeText(this, "PowerPoint loaded", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Failed to read PowerPoint: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    private void writePowerPoint(Uri uri) {
        try (OutputStream outputStream = getContentResolver().openOutputStream(uri)) {
            XMLSlideShow ppt = new XMLSlideShow();
            String[] slidesText = editText.getText().toString().split("---- Slide End ----");

            for (String slideContent : slidesText) {
                XSLFSlide slide = ppt.createSlide();
                XSLFTextShape shape = slide.createTextBox();
                shape.setText(slideContent.trim());
            }

            ppt.write(outputStream);
            ppt.close();
            Toast.makeText(this, "PowerPoint saved", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Failed to save PowerPoint", Toast.LENGTH_SHORT).show();
        }
    }
}
