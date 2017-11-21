package com.modeltodocx;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.Document;

public class TestCus
{

    public static void main(String []a) throws FileNotFoundException, IOException, InvalidFormatException
    {

        CustomXWPFDocument document = new CustomXWPFDocument(new FileInputStream(new File("D:/Work/javaweb/new.docx")));
        FileOutputStream fos = new FileOutputStream(new File("D:/Work/javaweb/new.docx"));
        String blipId = document.addPictureData(new FileInputStream(new File("D:/Work/1.jpg")), Document.PICTURE_TYPE_JPEG);
        document.createPicture(blipId,document.getNextPicNameNumber(Document.PICTURE_TYPE_JPEG), 50, 50);
        document.write(fos);
        fos.flush();
        fos.close();

    }

}