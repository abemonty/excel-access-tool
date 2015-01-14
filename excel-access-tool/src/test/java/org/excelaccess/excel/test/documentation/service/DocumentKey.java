package org.excelaccess.excel.test.documentation.service;

import org.excelaccess.excel.test.documentation.model.Document;

/**
 * Ensemble des champs représentant un document. Ici, aucun n'est obligatoire.
 * 
 * @author Loic Abemonty
 * 
 */
public class DocumentKey implements Document {

    public String application;

    public String name;

    public String reference;

    public String type;

    public String typeDoc;

    public Integer version;

    public String dateLivraison;

    public String details;

    @Override
    public String getApplication() {
        return application;
    }

    @Override
    public String getDateLivraison() {
        return this.dateLivraison;
    }

    @Override
    public String getDetails() {
        return this.details;
    }

    @Override
    public String getName() {
        return name;
    }

    @Override
    public String getReference() {
        return reference;
    }

    @Override
    public String getType() {
        return type;
    }

    @Override
    public String getTypeDoc() {
        return typeDoc;
    }

    @Override
    public Integer getVersion() {
        return version;
    }

}
