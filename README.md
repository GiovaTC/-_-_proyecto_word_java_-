# -_-_proyecto_word_java_- :.
# 📄 Proyecto Word Java:

<img width="1254" height="1254" alt="image" src="https://github.com/user-attachments/assets/605b5a6f-b8ce-45a8-9766-92d62cf31dc9" />    

<img width="1282" height="1079" alt="image" src="https://github.com/user-attachments/assets/7cea6cbd-0dfc-4237-aa74-85f509dd4a5a" />    

```
Proyecto desarrollado con **Java 21**, **Swing**, **Apache POI** e **IntelliJ IDEA**, que permite generar,
guardar, abrir y visualizar documentos **Microsoft Word (.docx)** mediante una interfaz gráfica sencilla .

---

# 📌 Tecnologías utilizadas

- ☕ Java 21
- 💻 IntelliJ IDEA
- 🖼️ Swing (Interfaz gráfica)
- 📄 Apache POI (poi-ooxml)
- 📦 Maven

---

# 🎯 Funcionalidades

El programa permite:

- ✅ Generar un archivo **Word (.docx)**
- ✅ Escribir varias líneas de texto
- ✅ Guardar el archivo
- ✅ Abrir un documento Word existente
- ✅ Leer todo su contenido
- ✅ Mostrar el contenido dentro de una interfaz gráfica (`JTextArea`)

---

# 📁 Estructura del proyecto

```text
ProyectoWordJava
│
├── pom.xml
│
└── src
    └── main
        ├── java
        │
        └── com
            └── ejemplo
                │
                ├── Main.java
                ├── WordManager.java
                └── VentanaPrincipal.java
```

---

# 📦 Dependencias Maven

```xml
<dependencies>

    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.4.1</version>
    </dependency>

</dependencies>
```

---

# 🖥️ Interfaz gráfica

```text
----------------------------------------------

          ARCHIVOS WORD JAVA

----------------------------------------------

[ Generar Word ]

[ Abrir Word ]

----------------------------------------------

+-----------------------------------------+

Aquí aparecerá el contenido del Word...

+-----------------------------------------+
```

---

# 🚀 Clase Main

```java
package com.ejemplo;

public class Main {

    public static void main(String[] args) {

        new VentanaPrincipal();

    }

}
```

---

# 📄 Clase WordManager

```java
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

        run.setText("Línea número 1");

        run.addBreak();

        run.setText("Línea número 2");

        run.addBreak();

        run.setText("Línea número 3");

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
```

---

# 🖼️ Clase VentanaPrincipal

```java
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
```

---

# ▶️ Resultado esperado

Al ejecutar la aplicación se mostrará una ventana similar a la siguiente:

```text
+------------------------------------------------------+

             ARCHIVOS WORD JAVA

--------------------------------------------------------

[ Generar Word ]   [ Abrir Word ]

--------------------------------------------------------

Proyecto Java Word

Línea número 1

Línea número 2

Línea número 3

--------------------------------------------------------
```

---

# 📚 Explicación del funcionamiento

## 1. Generar Word

Al pulsar el botón **Generar Word**:

- Se abre un `JFileChooser`.
- El usuario selecciona la ubicación.
- Se crea un archivo **.docx**.
- Se escriben varias líneas utilizando **Apache POI**.
- El documento queda almacenado en disco.

---

## 2. Abrir Word

Al pulsar el botón **Abrir Word**:

- Se selecciona un archivo `.docx`.
- Apache POI lee todos los párrafos.
- El contenido se concatena en un `String`.
- Finalmente se muestra dentro del `JTextArea`.

---

# 📌 Librerías utilizadas

| Librería | Función |
|----------|---------|
| Apache POI | Crear documentos Word |
| XWPFDocument | Documento Word |
| XWPFParagraph | Párrafos |
| XWPFRun | Texto con formato |
| Swing | Interfaz gráfica |
| JFileChooser | Selección de archivos |
| JTextArea | Mostrar contenido |
| JScrollPane | Área con desplazamiento |

---

# 💡 Mejoras que se pueden agregar

Este proyecto puede ampliarse para incluir funcionalidades más completas:

- 📄 Crear documentos Word con múltiples párrafos.
- 🖼️ Insertar imágenes en el documento.
- 📊 Agregar tablas con filas y columnas.
- 🎨 Aplicar estilos (títulos, encabezados, colores y fuentes).
- 🔍 Buscar palabras dentro del documento.
- ✏️ Editar y guardar nuevamente el archivo.
- 📋 Copiar el contenido del Word a una tabla (`JTable`).
- 📑 Soporte para varios documentos abiertos mediante `JTabbedPane`.
- 📤 Exportar el contenido a PDF.
- 🗃️ Guardar y recuperar documentos desde una base de datos MySQL.
- 🌙 Interfaz moderna con **FlatLaf** para una apariencia más profesional.

---

# 🏆 Conclusión

Este proyecto constituye una excelente introducción al manejo de documentos **Microsoft Word (.docx)** utilizando **Apache POI** en Java.

A través de una interfaz gráfica desarrollada con **Swing**, el usuario puede generar, guardar y visualizar documentos de forma sencilla y didáctica.

Su arquitectura modular facilita la incorporación de nuevas funcionalidades, convirtiéndolo en una base ideal para desarrollar aplicaciones de:

- 📂 Gestión documental.
- 📝 Editores de texto.
- 📄 Generadores automáticos de informes.
- 📚 Sistemas de administración de documentos.
- 💼 Aplicaciones empresariales con Java.

---

## 👨‍💻 Autor

Proyecto desarrollado con:

- Java 21
- Swing
- Apache POI
- Maven
- IntelliJ IDEA

---
**Proyecto educativo para aprender el manejo de documentos Word en Java.**
:. . / .
