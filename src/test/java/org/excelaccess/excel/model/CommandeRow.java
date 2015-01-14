package org.excelaccess.excel.model;

/**
 * Interface d'accès aux éléments d'une commande dans une ligne.
 * 
 * @author Loic Abemonty
 * 
 */
public interface CommandeRow {

    /**
     * Volume prévisionnel de la commande
     * 
     * @param index
     *            n-ième semaine de la commande, 0-based.
     */
    Integer getCommande(int index);

    void setCommande(int index, Integer volumeCommande);
}
