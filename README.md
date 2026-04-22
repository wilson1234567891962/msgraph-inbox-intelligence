# 馃搳 MS-Graph-Mail-Auditor
### **High-Performance Email Frequency Reporting & Inbox Peritaje**

Este repositorio contiene una herramienta de auditor铆a avanzada desarrollada en **PowerShell** que utiliza **Microsoft Graph API v1.0**. El script est谩 optimizado para procesar bandejas de entrada masivas (+7,000 registros), permitiendo identificar patrones de tr谩fico, remitentes frecuentes y dominios para la depuraci贸n de datos y peritaje inform谩tico.

---

## 馃殌 Caracter铆sticas Principales

* **Gesti贸n de Paginaci贸n Masiva:** Implementa el manejo autom谩tico de tokens `@odata.nextLink`, permitiendo descargar y procesar miles de correos en bloques de 1,000 registros de forma eficiente.
* **Programaci贸n Defensiva (Null-Safety):** Blindado contra errores de "Null Reference" comunes en objetos de Graph (como borradores o correos de sistema sin remitente), asegurando que el proceso no se interrumpa ante datos inconsistentes.
* **Arquitectura REST Directa:** Utiliza `Invoke-MgRestMethod` para evitar las limitaciones de validaci贸n de los cmdlets est谩ndar, permitiendo una comunicaci贸n m谩s limpia con los endpoints de Capa 7.
* **An谩lisis de Frecuencia en Tiempo Real:** Agrupa remitentes y dominios para generar estad铆sticas de saturaci贸n de bandeja de entrada, ideal para identificar spam persistente o newsletters.

---

## 馃洜锔?Stack T茅cnico

* **Lenguaje:** PowerShell 7+
* **Protocolo:** OAuth 2.0 / REST API
* **Endpoint:** Microsoft Graph API v1.0 (`/me/messages`)
* **Formato de Salida:** CSV (Codificaci贸n UTF-8) compatible con Microsoft Excel.

---

## 馃搵 Uso del Script

### **Requisitos Previos**
Es necesario tener instalado el SDK de Microsoft Graph en tu terminal:
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
