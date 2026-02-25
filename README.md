# Sistema web Fochev (gestión de ventas de productos para el cabello)

Este proyecto es un sistema web profesional para la empresa **Fochev**, que vende productos de belleza para el cabello.

Permite gestionar:
- Clientes (realizan pedidos en línea)
- Distribuidores (ven y gestionan pedidos de su zona)
- Administrador/propietario (control total de productos, distribuidores, zonas y ventas)

Está desarrollado con **Flask**, **SQLite** y **Bootstrap**.

## Requisitos previos

- Python 3.10+ instalado
- (Opcional) Entorno virtual de Python

## Instalación

En una terminal, desde la carpeta del proyecto:

```bash
pip install -r requirements.txt
```

## Ejecutar el sistema
python app.py
Desde la carpeta del proyecto:

```bash
python app.py
```

El sistema arrancará en modo desarrollo en:

```text
http://127.0.0.1:5000/
```

Abre esa URL en tu navegador.

En el primer arranque se creará automáticamente la base de datos `fochev.db`, la carpeta `exports` y datos de ejemplo
(zonas, usuarios y productos).

### Cuentas de prueba

Puedes iniciar sesión con estas cuentas ya creadas:

- **Administrador (dueño)**  
  - Correo: `admin@fochev.com`  
  - Contraseña: `admin123`

- **Distribuidor zona Centro**  
  - Correo: `dist.centro@fochev.com`  
  - Contraseña: `dist123`

- **Distribuidor zona Norte**  
  - Correo: `dist.norte@fochev.com`  
  - Contraseña: `dist123`

- **Cliente demo**  
  - Correo: `cliente.demo@fochev.com`  
  - Contraseña: `cliente123`

### Flujo principal

- **Cliente**
  - Inicia sesión con su cuenta de cliente.
  - Ve el catálogo de productos Fochev, selecciona cantidad y agrega al carrito.
  - Confirma el pedido.  
    - Se registra la venta en la base de datos.
    - Se descuenta stock de los productos.
    - Se genera un archivo Excel con todos los datos del cliente y del pedido en la carpeta `exports/`.

- **Distribuidor**
  - Inicia sesión como distribuidor.
  - Ve los pedidos pendientes de su zona.
  - Entra al detalle de cada pedido, ve datos completos del cliente, productos y total.
  - Actualiza el estatus del pedido (pendiente, asignado, en ruta, entregado, cancelado).

- **Administrador**
  - Inicia sesión como administrador.
  - Panel con resumen de ventas, pedidos recientes y alertas de stock bajo.
  - Gestiona productos (alta, baja, cambio de precio y stock, activar/desactivar).
  - Gestiona distribuidores (alta, baja, asignación de zona, activar/desactivar).
  - Gestiona zonas (alta, baja, actualización de responsable y correo).
  - Consulta todas las ventas/pedidos registrados y el estado de cada uno.

## Archivos Excel de pedidos

Cada vez que un cliente confirma un pedido, se genera automáticamente un archivo Excel en la carpeta:

```text
exports/pedido_<ID>.xlsx
```

Ese archivo contiene:
- Datos del cliente (nombre, teléfono, dirección, ubicación completa)
- Datos de la zona y responsable
- Detalle de productos (cantidad, precio unitario, subtotal)
- Total del pedido

Para enviar ese Excel por correo al responsable de la zona puedes:
- Integrar tu propio servidor SMTP en otro paso, o
- Descargar el archivo desde el servidor y enviarlo manualmente.

