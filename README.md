# 📊 MS-Graph-Mail-Auditor

**High-Performance Email Frequency Reporting & Inbox Peritaje**

Este repositorio contiene una herramienta de auditoría avanzada desarrollada en **PowerShell** que utiliza **Microsoft Graph API v1.0**.  
El script está optimizado para procesar bandejas de entrada masivas (+7,000 registros), permitiendo identificar patrones de tráfico, remitentes frecuentes y dominios para la depuración de datos y peritaje informático.

---

## 🚀 Características Principales

- **Gestión de Paginación Masiva**  
  Implementa el manejo automático de tokens `@odata.nextLink`, permitiendo descargar y procesar miles de correos en bloques de 1,000 registros de forma eficiente.

- **Programación Defensiva (Null-Safety)**  
  Blindado contra errores de tipo *Null Reference* en objetos de Graph (como borradores o correos de sistema sin remitente), asegurando que el proceso no se interrumpa.

- **Arquitectura REST Directa**  
  Utiliza `Invoke-MgRestMethod` para evitar limitaciones de validación de cmdlets estándar, permitiendo una comunicación más limpia con los endpoints de Capa 7.

- **Análisis de Frecuencia en Tiempo Real**  
  Agrupa remitentes y dominios para generar estadísticas de saturación de bandeja de entrada.

---

## 🛠️ Stack Técnico

- **Lenguaje:** PowerShell 7+  
- **Protocolo:** OAuth 2.0 / REST API  
- **Endpoint:** Microsoft Graph API v1.0 (`/me/messages`)  
- **Formato de Salida:** CSV (UTF-8) compatible con Microsoft Excel  

---

## 📦 Uso del Script

### 🔧 Requisitos Previos

Es necesario tener instalado el SDK de Microsoft Graph:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

### ▶️ Ejecución

1. Clona este repositorio  
2. Abre una terminal de PowerShell y ejecuta:

```powershell
Connect-MgGraph -Scopes "Mail.Read"
powershell -ExecutionPolicy Bypass -File .\ms-graph-mail-auditor.ps1
```

---

## 📊 Estructura del Reporte (Excel)

El script genera un archivo llamado:

Auditoria_Final_Wilson.csv

Con las siguientes columnas:

| Columna   | Descripción |
|----------|------------|
| **Remitente** | Dirección de correo completa del emisor |
| **Dominio**   | Dominio raíz extraído (útil para identificar proveedores) |
| **Cantidad**  | Frecuencia total de correos recibidos de esa fuente |
| **Prioridad** | Clasificación automática (CRÍTICO / Newsletter / Normal) |

---

## ⚖️ Aplicaciones en Peritaje Legal

Como herramienta desarrollada con enfoque en **Ingeniería y Derecho**, este script tiene aplicaciones directas en:

- **Auditoría Forense**  
  Identificación de flujos de comunicación persistentes.

- **Preservación de Evidencia**  
  Mapeo de metadatos de remitentes sin alterar el contenido del mensaje.

- **Optimización de Activos Digitales**  
  Depuración proactiva de ruido en comunicaciones corporativas.

---

## 👨‍💻 Autor

**Ing. Alejandro Wilson**

- Ingeniero Industrial & Ingeniero de Software  
- Magíster en Arquitectura de Software (Universidad de los Andes)  
- Estudiante de Derecho (Colombia)  

---

## 📄 Licencia

Este proyecto puede ser adaptado según las necesidades del usuario.
