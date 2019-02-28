# IOC_Emotet
Este repositorio es para tener guardado código para obtener IOC del malware de Emotet en documentos Word .doc en formato XML

Está ligado al análisis realizado en el siguiente artículo: https://www.linkedin.com/pulse/c%C3%B3mo-extraer-ioc-de-emotet-doc-xml-parte-3-soc-chile

# Requerimientos
Se necesita tener instalados los siguientes componentes antes:
- __XMLStarlet__ (está en la mayoría de repositorios oficiales http://xmlstar.sourceforge.net/)
- __olevba__ de oletools (está en la mayoría de los repositorios oficiales https://github.com/decalage2/oletools)

# Funcionamiento
Se ejecuta el script con el archivo a analizar como parámetro.

```
$ ./ioc_emotet.sh documento.doc
```

# Detecciones

Este script puede funcionar detectando estos tipos de archivos:
- Archivos .doc con objetos OLE embebidos (Composite Document File V2 Document)
- Archivos .doc con formato Microsoft Office XML (XML 1.0 document)
- Archivos .docx con objetos OLE embebidos (Microsoft Word 2007+)

# TODO
Faltan métodos de ofuscación, principalmente el de detección de código embebido en TextBox dentro del documento.
