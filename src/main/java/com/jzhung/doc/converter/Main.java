package com.jzhung.doc.converter;

import com.jzhung.doc.converter.uitl.DocManager;

/**
 * Created by Jzhung on 2018/1/5.
 */
public class Main {
    public static void main(String[] args) {
        String src = "E:\\Data\\doc\\1.docx";
        //DocManager.docxToDoc(src, "E:\\Data\\doc\\1.doc");
        DocManager.docToHtml(src, "E:\\Data\\doc\\1.html");
    }
}
