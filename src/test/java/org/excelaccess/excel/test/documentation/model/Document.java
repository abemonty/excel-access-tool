package org.excelaccess.excel.test.documentation.model;

/**
 * Définition de l'objet Document.
 * 
 * @author Loic Abemonty
 */
public interface Document {

    /**
     * Nom de l'application
     */
    String getApplication();

    /**
     * Date de livraison de la version.
     */
    String getDateLivraison();

    /**
     * Détails du fichier.
     */
    String getDetails();

    /**
     * Nom du fichier.
     */
    String getName();

    /**
     * Référence externe.
     */
    String getReference();

    /**
     * Catégorie de la documentation.
     */
    String getType();

    /**
     * Type du fichier (spec fonctionnelle, spec technique, etc).
     */
    String getTypeDoc();

    /**
     * Numéro de la version livrée.
     */
    Integer getVersion();

}
