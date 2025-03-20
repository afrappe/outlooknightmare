# El Infierno del Email
Una macro para luchar contra el email, esta macro permite trabajar con un Outlook en Español.
Usa cadenas de caracteres de mensajes que se encuentran en español, como De:, Para:, Entregado:
Puedes ampliarlo para que se ajuste a tus necesidades.

Básicamente lo que hace es identificar el mailbox que tienes configurado en Outlook, y descargar todos los mensajes y pasarlos a un formato tabular.

Para **cargar un archivo `.bas` como macro en Excel**, sigue estos pasos:  

---

### **🔹 1. Abre el Editor de Visual Basic (VBA) 🏗️**  
1️⃣ Abre **Microsoft Excel**.  
2️⃣ Presiona `ALT + F11` para abrir el **Editor de VBA**.  

---

### **🔹 2. Importar el Archivo `.bas` 📂**  
1️⃣ En el Editor de VBA, ve al menú **"Archivo" → "Importar archivo..."**.  
2️⃣ Selecciona tu archivo `.bas` y haz clic en **"Abrir"**.  
3️⃣ El código de la macro se importará en un **módulo nuevo** dentro de "Módulos".  

💡 *Si no ves "Módulos", expande la sección "Módulos" en el Explorador de Proyectos (CTRL + R).*  

---

### **🔹 3. Guardar y Habilitar Macros 💾**  
1️⃣ Guarda el archivo de Excel con macros activadas:  
   - **Archivo → Guardar como → "Libro de Excel habilitado para macros" (`.xlsm`)**.  
2️⃣ Cierra y vuelve a abrir el archivo para asegurarte de que las macros estén habilitadas.  
3️⃣ Si aparece una advertencia de seguridad, haz clic en **"Habilitar contenido"**.  

---

### **🔹 4. Ejecutar la Macro 🚀**  
1️⃣ Vuelve al editor de VBA (`ALT + F11`).  
2️⃣ Abre el módulo donde importaste el archivo `.bas`.  
3️⃣ Presiona `F5` para ejecutar la macro o llámala desde Excel escribiendo su nombre en la ventana de macros (`ALT + F8`).  

---

### **🔹 5. (Opcional) Agregar un Botón en la Hoja de Excel 🖱️**  
1️⃣ Ve a **"Desarrollador" → "Insertar" → "Botón"** (si no ves esta pestaña, actívala en "Opciones de Excel").  
2️⃣ Dibuja el botón en la hoja de cálculo.  
3️⃣ En la ventana emergente, selecciona la macro importada y haz clic en **"Aceptar"**.  
4️⃣ ¡Listo! Ahora puedes ejecutar la macro haciendo clic en el botón.  

---


