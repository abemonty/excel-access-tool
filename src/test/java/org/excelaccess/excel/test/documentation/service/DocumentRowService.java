package org.excelaccess.excel.test.documentation.service;

import org.excelaccess.excel.ExcelAccessor;
import org.excelaccess.excel.RowHandler;
import org.excelaccess.excel.test.documentation.model.Document;
import org.excelaccess.excel.test.documentation.model.excel.DocumentDelivered;

public class DocumentRowService extends RowHandler<Document, DocumentDelivered> {

    public DocumentRowService(ExcelAccessor excelAccessor) {
        super(excelAccessor);
    }

    @Override
    protected boolean identifyResource(DocumentDelivered resource, Document resourceKey) {
        return isMatching(resourceKey.getApplication(), resource.getApplication())
                && isMatching(resourceKey.getType(), resource.getType())
                && isMatching(resourceKey.getTypeDoc(), resource.getTypeDoc())
                && isMatching(resourceKey.getName(), resource.getName())
                && isMatching(resourceKey.getDetails(), resource.getDetails())
                && isMatching(resourceKey.getReference(), resource.getReference())
                && isMatching(resourceKey.getVersion(), resource.getVersion());
    }

}
