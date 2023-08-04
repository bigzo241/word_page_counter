package org.mamadou;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException, URISyntaxException {
//        URL du dossier contenant les documents Word
        URL dirPath = Main.class.getClassLoader().getResource("doc");

//        Recuperation des documents dans une liste
        List<Path> pathList= Files.list(Paths.get(dirPath.toURI())).toList();

//        Parcours de la liste et determination du nombre de page de chaque document
        for (Path doc : pathList) {
            try (OPCPackage opcPackage = OPCPackage.open(doc.toString());
                 XWPFDocument document = new XWPFDocument(opcPackage)) {

                int pageCount = document.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();
                System.out.println(doc.getFileName().toString() + " : " + pageCount + " page(s)");

            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}