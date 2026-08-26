# Retenciones SRI

Aplicación web para cargar el TXT de comprobantes emitidos del SRI, consultar cada clave de acceso contra el Web Service oficial de autorización y descargar un Excel con el detalle de retenciones.

## Ejecutar localmente

```bash
pip install -r requirements.txt
uvicorn main:app --reload
```

Abrir http://127.0.0.1:8000

## Publicar gratis en Render

1. Crear un repositorio en GitHub y subir estos archivos.
2. En Render: New > Web Service > conectar el repositorio.
3. Runtime: Python 3.
4. Build command: `pip install -r requirements.txt`
5. Start command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
6. Seleccionar plan Free y Deploy.

No requiere usuario ni contraseña del SRI. Usa únicamente las claves de acceso incluidas en el TXT.

> Recomendación: probar primero con 2-5 comprobantes antes de procesar lotes grandes.
