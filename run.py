# Punto de entrada de la aplicación FlowOps.
# Este es el fichero que ejecutas con "python run.py" para arrancar el servidor.

from app import create_app

# Creamos la instancia de la aplicación Flask usando la factory
# definida en app/__init__.py
app = create_app()

if __name__ == "__main__":
    # debug=True activa el recargado automático al guardar cambios
    # y muestra errores detallados en el navegador. Solo para desarrollo.
    app.run(debug=True)