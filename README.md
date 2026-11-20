# Mis gastos · app para el móvil

Contenido del paquete:

- `app/`: la aplicación (se sube tal cual a un hosting).
- `apps-script/Code.gs`: el código que conecta la app con tu Google Sheet.

## 1. Preparar la Google Sheet (5 min)

1. Crea una hoja nueva en la cuenta de Google donde quieras guardar tus gastos.
2. Menú **Extensiones › Apps Script**. Borra lo que haya y pega el contenido de `Code.gs`. Guarda.
3. Arriba, elige la función **setup** y pulsa **Ejecutar**. Acepta los permisos. Google avisará de que la app no está verificada: es tu propio script, así que pulsa *Configuración avanzada › Ir a…*.
   Se crean las pestañas **Gastos**, **Config** y **Resumen**. La clave para la app queda en **Config › B2**.
4. **Implementar › Nueva implementación**. En el engranaje elige **Aplicación web**:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
5. Pulsa **Implementar** y copia la URL que acaba en `/exec`.

> Si más adelante cambias `Code.gs`, usa **Implementar › Gestionar implementaciones › Editar › Nueva versión**, para que la URL no cambie.

## 2. Publicar la app

Ya está publicada con GitHub Pages en https://seabornsmile.github.io/GASTOS-GENERALES/ (este repositorio). Las opciones de abajo solo hacen falta si algún día la mueves a otro sitio.

**Opción fácil: Netlify Drop**
1. Entra en https://app.netlify.com/drop (crea una cuenta gratis si te la pide).
2. Arrastra la carpeta `app` a la página.
3. Te da un enlace `https://algo.netlify.app`. Puedes cambiarle el nombre en *Site settings*.

**Alternativa: GitHub Pages**
1. Crea un repositorio y sube el contenido de la carpeta `app`.
2. *Settings › Pages › Deploy from a branch › main / root*.
3. El enlace será `https://tuusuario.github.io/nombre-repo/`.

## 3. Instalar en el móvil

**Android (Chrome)**: abre el enlace › menú ⋮ › **Instalar aplicación** (o «Añadir a pantalla de inicio»).

**iPhone (Safari)**: abre el enlace en **Safari** › botón Compartir › **Añadir a pantalla de inicio**.

La primera vez, pega la URL `/exec` y la clave. Hay que hacerlo una vez en cada móvil.

## Actualizar a la versión 2 (ingresos, presupuestos por categoría, ahorro…)

1. Abre tu hoja › **Extensiones › Apps Script**, borra el código y pega el nuevo `apps-script/Code.gs`. Guarda.
2. Elige **setup** › **Ejecutar**. Convierte la pestaña *Gastos* en *Movimientos* y crea *Fijos*, *Categorias* y *Objetivos*, sin perder ningún dato.
3. **Implementar › Gestionar implementaciones** › lápiz (Editar) › *Versión: Nueva versión* › **Implementar**. La URL `/exec` no cambia.
4. Abre la app en el móvil con conexión: se actualiza sola. Si ves el aviso «el script de tu hoja es la versión 1», repasa el paso 3.

## Cómo funciona

- Cada gasto que apuntas se guarda como una fila en la pestaña **Gastos**. Puedes editar la hoja a mano y la app lo verá al sincronizar.
- Si no hay cobertura, la app guarda los cambios y los envía cuando vuelve la conexión (el punto de la esquina se pone amarillo).
- **Resumen** muestra el gasto por mes y categoría con una fórmula que se actualiza sola.
- Para pasar los gastos de la versión de Claude: allí pulsa *Exportar todo (CSV)* y aquí *Ajustes › Importar gastos desde CSV*.

## Seguridad

La URL y la clave dan acceso a tus gastos: no las compartas. Si crees que alguien las tiene, borra la propiedad `TOKEN` en Apps Script (*Configuración del proyecto › Propiedades del script*), vuelve a ejecutar **setup** y pon la clave nueva en la app.
