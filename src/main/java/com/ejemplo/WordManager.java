package com.ejemplo;

import org.apache.poi.xwpf.usermodel.*;
import java.io.*;

public class WordManager {

    public void crearWord(String ruta) throws Exception {

        XWPFDocument documento = new XWPFDocument();
        XWPFParagraph p = documento.createParagraph();

        XWPFRun run = p.createRun();

        run.setBold(true);
        run.setFontSize(16);
        run.setText("Proyecto Java Word");

        run.addBreak();

        run.setBold(false);
        run.setText("Linea numero 1");

        run.addBreak();

        run.setText("Linea numero 2");

        run.addBreak();

        run.setText("Linea numero 3");

        FileOutputStream salida =
                new FileOutputStream(ruta);

        documento.write(salida);
        salida.close();
        documento.close();
    }

    public String leerWord(String ruta)
        throws Exception {

        StringBuilder texto =
                new StringBuilder();

        FileInputStream entrada =
                new FileInputStream(ruta);

        XWPFDocument documento =
                new XWPFDocument(entrada);

        for (XWPFParagraph p :
        documento.getParagraphs()) {

            texto.append(p.getText());
            texto.append("\n");
        }

        documento.close();
        entrada.close();

        return texto.toString();
    }
}   
