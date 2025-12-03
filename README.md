# Cursos 5 CBS 🚒💚

Es una ayuda para poner de manera más fácil las fechas en las que se completa un curso y tener un registro de mejor manera.

## Agradecimientos

- Sergio Cortés Prado
- Clemente Garmendia Pascal

## Features

- ✅ **Búsqueda pública**: Cualquier usuario puede buscar sus cursos ingresando su nombre (sin autenticación)
- ✅ **Editor protegido**: Solo administradores pueden subir y editar archivos Excel
- ✅ **Preservación de formato**: Mantiene el formato original del Excel al agregar registros
- ✅ **Formulario dinámico**: Detecta automáticamente las columnas del Excel y crea campos apropiados
- ✅ **Interfaz moderna**: Diseño limpio y fácil de usar
- ✅ **Descarga de archivos**: Permite descargar el archivo Excel actualizado

## Estructura del Proyecto

```
Cursos-5 V.0/
├── app.py                 # Punto de entrada principal
├── pages/                 # Páginas de la aplicación
│   ├── landing.py        # Página de inicio
│   ├── login.py          # Página de login (admin)
│   └── editor.py         # Página de editor (admin)
├── components/            # Componentes reutilizables
│   └── search.py         # Componente de búsqueda pública
├── utils/                 # Utilidades
│   ├── auth.py           # Funciones de autenticación
│   ├── excel_handler.py  # Manejo de archivos Excel
│   └── styling.py        # Estilos CSS
├── style.css             # Estilos personalizados
└── requirements.txt      # Dependencias
```

## Setup Local

1. Instala las dependencias:
```bash
pip install -r requirements.txt
```

2. Crea un archivo `.env` con las credenciales del administrador:
```
EDITOR_USERNAME=tu_usuario
EDITOR_PASSWORD=tu_contraseña
```

3. Ejecuta la aplicación:
```bash
streamlit run app.py
```

## Despliegue en Streamlit Cloud (Gratuito)

Para desplegar la aplicación en Streamlit Cloud de forma gratuita:

1. **Sube tu código a GitHub**:
   - Crea un repositorio en GitHub
   - Sube todos los archivos del proyecto

2. **Conecta con Streamlit Cloud**:
   - Ve a [share.streamlit.io](https://share.streamlit.io)
   - Inicia sesión con tu cuenta de GitHub
   - Haz clic en "New app"
   - Selecciona tu repositorio y branch
   - Configura el archivo principal como `app.py`

3. **Configura variables de entorno**:
   - En la configuración de la app, ve a "Secrets"
   - Agrega las credenciales:
   ```toml
   EDITOR_USERNAME = "tu_usuario"
   EDITOR_PASSWORD = "tu_contraseña"
   ```

4. **Despliega**:
   - Haz clic en "Deploy"
   - Tu app estará disponible en `https://tu-app.streamlit.app`

## Formato del Archivo Excel

El Excel debe tener dos hojas:
1. **"HONORARIOS"** (o "Honorarios")
2. **"ACTIVOS"** (o "Activos")

El sistema detecta automáticamente las columnas y crea campos de formulario dinámicos.

## Uso

### Para Usuarios Públicos:
1. Haz clic en "🔍 Buscar Mis Cursos"
2. Ingresa tu nombre completo o parte de él
3. Verás todos tus cursos registrados en ambas hojas

### Para Administradores:
1. Haz clic en "✏️ Editor (Admin)"
2. Ingresa tus credenciales
3. Sube el archivo Excel
4. Completa el formulario para agregar nuevos registros
5. Descarga el archivo actualizado

## Notas Importantes

- El archivo Excel debe ser subido por un administrador antes de que los usuarios puedan buscar
- El formato del Excel se preserva al agregar nuevos registros
- Los datos se mantienen en memoria durante la sesión del administrador
- Para que la búsqueda funcione, el administrador debe haber cargado el archivo en la misma sesión

## Tecnologías Utilizadas

- **Streamlit**: Framework web
- **Pandas**: Manipulación de datos
- **OpenPyXL**: Manejo de archivos Excel con formato
- **Python**: Lenguaje de programación 


# .ENV

`````
EDITOR_USERNAME=admin
EDITOR_PASSWORD=Quinta2025
`````