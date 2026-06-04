# Generador de Informes Técnicos para Instalación de UPS

## Descripción del proyecto

El Generador de Informes Técnicos para Instalación de UPS es una aplicación web desarrollada para automatizar la generación de informes técnicos a partir de la información recolectada en EpiCollect5 durante procesos de instalación de sistemas UPS.

El sistema permite consultar registros almacenados en EpiCollect, seleccionar uno o varios formularios diligenciados por los técnicos de campo y generar automáticamente informes en formato Excel con evidencia fotográfica y datos técnicos.

## Problema que resuelve

Antes del desarrollo del sistema, la generación de informes técnicos se realizaba manualmente, lo que generaba:

* Retrasos en la entrega de informes.
* Duplicidad de trabajo.
* Errores de digitación.
* Dificultad para consolidar evidencias fotográficas.
* Baja eficiencia en el proceso de documentación técnica.

La solución desarrollada automatiza el proceso completo de generación de informes.

---

## Arquitectura General

EpiCollect5 → Backend Flask (Render) → Frontend React (Netlify) → Usuario Final

### Flujo de información

1. El técnico registra la información en EpiCollect5.
2. El backend consulta los registros mediante la API de EpiCollect.
3. El frontend muestra los registros disponibles.
4. El usuario selecciona uno o varios registros.
5. El backend genera automáticamente los informes.
6. El usuario descarga los archivos generados.

---

## Tecnologías utilizadas

### Frontend

* React
* JavaScript
* CSS

### Backend

* Python
* Flask
* OpenPyXL
* Requests

### Infraestructura

* GitHub
* Netlify
* Render
* EpiCollect5

---

## Estructura del proyecto

```text
frontend/
├── public/
├── src/
│   ├── components/
│   ├── pages/
│   ├── services/
│   └── styles/

backend/
├── data/
├── app.py
└── requirements.txt
```

## Instalación local

### Clonar repositorio

```bash
git clone https://github.com/AnalistaCSA/app_generacion_informes.git
```

### Backend

```bash
cd backend
pip install -r requirements.txt
python app.py
```

### Frontend

```bash
cd frontend
npm install
npm start
```

## Variables de entorno

Ejemplo:

```env
API_URL=https://five.epicollect.net/api/export/entries/...
```

## Producción

### Frontend

https://generador-de-informes-csa.netlify.app/

### Backend

https://app-generacion-informes.onrender.com/datos

## Autor

Proyecto desarrollado como trabajo académico para la automatización de generación de informes técnicos de instalación de UPS.
