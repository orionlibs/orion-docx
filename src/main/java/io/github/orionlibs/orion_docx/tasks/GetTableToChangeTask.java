package io.github.orionlibs.orion_docx.tasks;

import com.spire.doc.Document;
import com.spire.doc.DocumentObject;
import com.spire.doc.Table;
import com.spire.doc.TableCell;
import com.spire.doc.TableRow;
import com.spire.doc.documents.Paragraph;
import java.util.Iterator;

public class GetTableToChangeTask
{
    public static Table run(Document document, int depthOfSearchForTable, String textToLookForInTheTable)
    {
        Table tableToChange = null;
        Iterator<DocumentObject> iterator = document.getChildObjects().iterator();
        loop1:
        while(iterator.hasNext())
        {
            if(depthOfSearchForTable == 1)
            {
                DocumentObject documentObject = iterator.next();
                tableToChange = findTableInResults(documentObject, textToLookForInTheTable);
                if(tableToChange != null)
                {
                    break loop1;
                }
            }
            else if(depthOfSearchForTable >= 2)
            {
                Iterator<DocumentObject> iterator1 = iterator.next().getChildObjects().iterator();
                while(iterator1.hasNext())
                {
                    if(depthOfSearchForTable == 3)
                    {
                        Iterator<DocumentObject> iterator2 = iterator1.next().getChildObjects().iterator();
                        while(iterator2.hasNext())
                        {
                            DocumentObject documentObject = iterator2.next();
                            tableToChange = findTableInResults(documentObject, textToLookForInTheTable);
                            if(tableToChange != null)
                            {
                                break loop1;
                            }
                        }
                    }
                    else
                    {
                        DocumentObject documentObject = iterator1.next();
                        tableToChange = findTableInResults(documentObject, textToLookForInTheTable);
                        if(tableToChange != null)
                        {
                            break loop1;
                        }
                    }
                }
            }
        }
        return tableToChange;
    }


    private static Table findTableInResults(DocumentObject documentObject, String textToLookForInTheTable)
    {
        Table tableToChange = null;
        if(documentObject instanceof Table)
        {
            Table table = (Table)documentObject;
            Iterator<TableRow> rows = table.getRows().iterator();
            loop2:
            while(rows.hasNext())
            {
                TableRow row = rows.next();
                Iterator<TableCell> cells = row.getCells().iterator();
                while(cells.hasNext())
                {
                    Iterator<Paragraph> paragraphs = cells.next().getParagraphs().iterator();
                    while(paragraphs.hasNext())
                    {
                        String text = paragraphs.next().getText();
                        if(text.indexOf(textToLookForInTheTable) != -1)
                        {
                            tableToChange = table;
                            break loop2;
                        }
                    }
                }
            }
        }
        return tableToChange;
    }
}