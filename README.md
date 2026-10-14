# Aula — Registro de calificaciones

Aplicación de escritorio desarrollada en **Python y Tkinter** para registrar, consultar y organizar las calificaciones de estudiantes. Su interfaz permite gestionar grupos y materias desde la misma aplicación, y exportar los registros a un archivo de Excel.

> Proyecto de entrega de la materia **Programación Basada en Ingeniería en Software**.

## ✨ Funciones principales

- **Registro de calificaciones:** guarda nombre del estudiante, grupo, materia, calificación, fecha y observaciones.
- **Grupos configurables:** incluye los grupos A2, A3, A4 y A5. Desde **Ver materias** se puede cambiar el nombre de cada grupo.
- **Materias por grupo:** cada grupo tiene su propia lista. Al elegir un grupo en el formulario, solo aparecen las materias asignadas a ese grupo.
- **Administración sencilla:** desde **Ver materias** se pueden editar las materias de cada grupo, escribiendo una materia por línea.
- **Edición y eliminación:** permite seleccionar un registro para actualizarlo o eliminarlo.
- **Búsqueda y filtros:** busca por estudiante o materia, filtra por materia y ordena los nombres de A–Z o Z–A.
- **Resumen del grupo:** muestra el total de registros, el promedio general y el porcentaje de aprobación. Se considera aprobatoria una calificación de 60 o más.
- **Exportación a Excel:** genera un archivo `.xlsx` con los datos y formato básico para facilitar su lectura.
- **Guardado local automático:** conserva los datos en una base SQLite junto al programa.
- **Interfaz en español:** diseño visual suave y pensado para facilitar el uso diario.

## 📚 Materias iniciales

Las listas se pueden modificar desde el botón **Ver materias**.

| Grupo | Materias iniciales |
| --- | --- |
| **A2** | Historia, Español, Biología, Álgebra |
| **A3** | Literatura, Geografía, Química, Aritmética |
| **A4** | Inglés, Física, Civismo, Estadística |
| **A5** | Filosofía, Ecología, Trigonometría, Economía |

## 🧰 Requisitos

- Python 3.9 o posterior
- Tkinter (incluido normalmente con Python)
- `openpyxl` para exportar archivos de Excel

## 🚀 Instalación y ejecución

1. Descarga o clona este repositorio.
2. Abre una terminal en la carpeta del proyecto.
3. Instala la dependencia de Excel:

   ```bash
   pip install -r requirements.txt
   ```

4. Inicia la aplicación:

   ```bash
   python registro_calificaciones.py
   ```

Si tu sistema usa el comando `python3`, puedes iniciar el programa con:

```bash
python3 registro_calificaciones.py
```

## 📝 Uso rápido

1. Escribe el nombre del estudiante.
2. Selecciona un grupo; la lista de materias se actualizará automáticamente.
3. Selecciona la materia, indica la calificación entre 0 y 100 y completa los demás datos.
4. Pulsa **Guardar calificación**.
5. Para cambiar un registro, selecciónalo en la tabla y pulsa **Editar seleccionado**. También puedes hacer doble clic en la fila.
6. Para configurar grupos o materias, pulsa **Ver materias** en la parte superior.
7. Usa los filtros y la búsqueda para localizar registros, o pulsa **Exportar Excel** para generar el archivo `.xlsx`.

## 💾 Almacenamiento de datos

La aplicación crea automáticamente el archivo `calificaciones.db` en la misma carpeta que `registro_calificaciones.py`. Este archivo contiene los registros, los nombres de los grupos y las materias configuradas. Para conservar o transferir la información, guarda una copia de `calificaciones.db` junto con el programa.

## 📂 Archivos del proyecto

```text
registro_calificaciones.py   Aplicación de escritorio
requirements.txt             Dependencias de Python
calificaciones.db            Base de datos local, creada al iniciar el programa
```

## 👤 Créditos

**Juan Pablo Cabrera Leo** — GitHub: [@akaInt-Code](https://github.com/akaInt-Code)

Proyecto de entrega de la materia **Programación Basada en Ingeniería en Software**.
