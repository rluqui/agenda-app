# 📒 Agenda Tunuyán Superior

Este proyecto está desarrollado con **Google Apps Script** y estructurado para trabajar con **clasp**, facilitando su edición en Visual Studio Code y el control de versiones con GitHub.

---

## 📁 Estructura del proyecto

El proyecto está organizado en varias carpetas para mantener el código ordenado y facilitar su mantenimiento.

```
📁 src/                  → Código fuente principal en Apps Script  
 ┓ 📄 init.js            → Inicializa datos compartidos y funciones comunes  
 ┓ 📄 autoridad.js       → Lógica para gestionar autoridades  
 ┓ 📄 encabezados.js     → Encabezados y estructura de reportes  
 ┓ 📄 modelo.js          → Modelo de datos y utilidades generales  
 ┓ 📄 roles.js           → Definición y validación de roles de usuario  
 ┗ 📄 telefonos.js       → Gestión de números de contacto

📁 .github/workflows/   → Automatización con GitHub Actions  
 ┗ 📄 deploy.yml         → Despliegue automático del proyecto (requiere configuración con OIDC)

📁 backups/             → Copias de seguridad manuales del código  
 ┗ 📄 Codigo_original_backup.js → Versión inicial del código

📄 .clasp.json          → Configuración del proyecto para usar `clasp`  
📄 .gitignore           → Archivos a excluir del control de versiones  
📄 README.md            → Documentación del proyecto
```

---

## ⚙️ Instalación

1. Cloná este repositorio:

   ```bash
   git clone https://github.com/usuario/repositorio.git
   ```

2. Instalá `clasp` si no lo tenés:

   ```bash
   npm install -g @google/clasp
   ```

3. Iniciá sesión en Google:

   ```bash
   clasp login
   ```

4. Sincronizá el proyecto:

   ```bash
   clasp pull
   ```

---

## 🚀 Despliegue automático (GitHub Actions)

> ⚠️ Requiere configuración previa en Google Cloud IAM + Workload Identity Federation.

Cada vez que hacés `git push`, el script se despliega automáticamente a tu proyecto de Apps Script.

Archivo responsable:

```
.github/workflows/deploy.yml
```

---

## 🤩 Uso

Este proyecto está vinculado con **AppSheet** y permite:

* Gestionar autoridades y teléfonos oficiales
* Generar reportes estructurados con encabezados dinámicos
* Controlar el acceso mediante validación de roles

---

## 🧑‍💻 Créditos

* **Desarrollado por**: Ricardo Luqui
* **Asistencia técnica**: ChatGPT
* **Dependencias**:

  * [Google Apps Script](https://developers.google.com/apps-script)
  * [clasp](https://github.com/google/clasp)
  * [GitHub Actions](https://docs.github.com/en/actions)

---

## 🚡 Notas

* ⚠️ Este proyecto puede escalar y recibir mejoras futuras.
* Se recomienda mantener una copia actualizada en la carpeta `backups/`.
