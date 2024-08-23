package io.github.orionlibs.orion_docx;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Table;
import io.github.orionlibs.orion_docx.tasks.GetTableToChangeTask;

public class DocxService
{
    public static Document loadFile(String filePathInResourcesFolder)
    {
        Document document = new Document();
        document.loadFromStream(DocxService.class.getResourceAsStream(filePathInResourcesFolder), FileFormat.Auto);
        return document;
    }


    public static void saveAsPDFAndCloseFile(Document document, String tempFileName)
    {
        document.saveToFile(tempFileName, FileFormat.PDF);
        document.close();
        document.dispose();
        document = null;
    }


    public static Table getTableToChange(Document document, int depthOfSearchForTable, String textToLookForInTheTable)
    {
        return GetTableToChangeTask.run(document, depthOfSearchForTable, textToLookForInTheTable);
    }
}