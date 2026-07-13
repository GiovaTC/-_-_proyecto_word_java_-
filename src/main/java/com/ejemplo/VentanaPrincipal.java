package com.ejemplo;

import javax.swing.*;
import java.awt.*;
import java.io.File;

public class VentanaPrincipal extends JFrame {

    private JTextArea area;

    private WordManager manager;

    public VentanaPrincipal() {

        manager = new WordManager();

        setTitle("Word Java");

        setSize(700,500);

        setLocationRelativeTo(null);

        setDefaultCloseOperation(EXIT_ON_CLOSE);

        JPanel superior = new JPanel();

        JButton crear = new JButton("Generar Word");

        JButton abrir = new JButton("Abrir Word");

        superior.add(crear);

        superior.add(abrir);

        area = new JTextArea();

        JScrollPane scroll =
                new JScrollPane(area);

        add(superior, BorderLayout.NORTH);

        add(scroll, BorderLayout.CENTER);

        crear.addActionListener(e -> {

            JFileChooser chooser =
                    new JFileChooser();

            chooser.setSelectedFile(
                    new File("Documento.docx"));

            int opcion =
                    chooser.showSaveDialog(this);

            if(opcion == JFileChooser.APPROVE_OPTION){

                try{

                    manager.crearWord(
                            chooser.getSelectedFile()
                                    .getAbsolutePath());

                    JOptionPane.showMessageDialog(
                            this,
                            "Archivo creado correctamente");

                }

                catch(Exception ex){

                    JOptionPane.showMessageDialog(
                            this,
                            ex.getMessage());

                }

            }

        });
        abrir.addActionListener(e -> {

            JFileChooser chooser =
                    new JFileChooser();

            int opcion =
                    chooser.showOpenDialog(this);

            if(opcion==JFileChooser.APPROVE_OPTION){

                try{

                    String texto =
                            manager.leerWord(
                                    chooser.getSelectedFile()
                                            .getAbsolutePath());

                    area.setText(texto);

                }

                catch(Exception ex){

                    JOptionPane.showMessageDialog(
                            this,
                            ex.getMessage());
                }
            }

        });

        setVisible(true);
    }   
}