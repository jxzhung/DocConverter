package com.jzhung.doc.converter.uitl;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jzhung.doc.converter.entity.WdSaveFormat;

/**
 * Main Interface for document's convert
 *
 * Created by Jzhung on 2018/1/5.
 */
public class DocManager {

    /**
     * convert doc to html
     *
     * @param src    file to convert
     * @param target converted file
     */
    public static void docToHtml(String src, String target) {
        convert(src, target, WdSaveFormat.wdFormatHTML);
    }

    /**
     * convert docx to doc
     *
     * @param src    file to convert
     * @param target converted file
     */
    public static void docxToDoc(String src, String target) {
        convert(src, target, WdSaveFormat.wdFormatDocument);
    }

    /**
     * convert doc to pdf
     *
     * @param src    file to convert
     * @param target converted file
     */
    public static void docToPdf(String src, String target) {
        convert(src, target, WdSaveFormat.wdFormatPDF);
    }


    private static boolean convert(String inputFile, String saveFile, int format) {
        ActiveXComponent app = null;
        try {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", false);
            Dispatch docs = app.getProperty("Documents").toDispatch();
            Dispatch doc = Dispatch.call(docs, "Open", inputFile, false, true)
                    .toDispatch();
            Dispatch.call(doc, "SaveAs2", saveFile, format);
            Dispatch.call(doc, "Close", false);
        } catch (Exception e) {
            return false;
        } finally {
            if (app != null) {
                app.invoke("Quit", 0);
            }
        }
        return true;
    }
}
