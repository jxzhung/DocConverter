package com.jzhung.doc.converter.entity;

/**
 * Doc format table
 * Created by Jzhung on 2018/1/5.
 */
public class WdSaveFormat {
    /**
     * DOC TYPE FROM ： https://msdn.microsoft.com/zh-cn/VBA/Word-VBA/articles/wdsaveformat-enumeration-word
     */

    public static final int wdFormatDocument = 0;//	Microsoft Office Word 97 - 2003 binary file format.
    public static final int wdFormatDOSText = 4;//	Microsoft DOS text format.
    public static final int wdFormatDOSTextLineBreaks = 5;//	Microsoft DOS text with line breaks preserved.
    public static final int wdFormatEncodedText = 7;//	Encoded text format.
    public static final int wdFormatFilteredHTML = 10;//	Filtered HTML format.
    public static final int wdFormatFlatXML = 19;//	Open XML file format saved as a single XML file.
    public static final int wdFormatFlatXMLWithMacros = 20;//	Open XML file format with macros enabled saved as a single XML file.
    public static final int wdFormatFlatXMLTemplate = 21;//	Open XML template format saved as a XML single file.
    public static final int wdFormatFlatXMLTemplateMacroEnabled = 22;//	Open XML template format with macros enabled saved as a single XML file.
    public static final int wdFormatOpenDocumentText = 23;//	OpenDocument Text format.
    public static final int wdFormatHTML = 8;//	Standard HTML format.
    public static final int wdFormatRTF = 6;//	Rich text format (RTF).
    public static final int wdFormatStrictOpenXMLDocument = 24;//	Strict Open XML document format.
    public static final int wdFormatTemplate = 1;//	Word template format.
    public static final int wdFormatText = 2;//	Microsoft Windows text format.
    public static final int wdFormatTextLineBreaks = 3;//	Windows text format with line breaks preserved.
    public static final int wdFormatUnicodeText = 7;//	Unicode text format.
    public static final int wdFormatWebArchive = 9;//	Web archive format.
    public static final int wdFormatXML = 11;//	Extensible Markup Language (XML) format.
    public static final int wdFormatDocument97 = 0;//	Microsoft Word 97 document format.
    public static final int wdFormatDocumentDefault = 16;//	Word default document file format. For Word, this is the DOCX format.
    public static final int wdFormatPDF = 17;//	PDF format.
    public static final int wdFormatTemplate97 = 1;//	Word 97 template format.
    public static final int wdFormatXMLDocument = 12;//	XML document format.
    public static final int wdFormatXMLDocumentMacroEnabled = 13;//	XML document format with macros enabled.
    public static final int wdFormatXMLTemplate = 14;//	XML template format.
    public static final int wdFormatXMLTemplateMacroEnabled = 15;//	XML template format with macros enabled.
    public static final int wdFormatXPS = 18;//	XPS format.
}