# Canban

"Kanban Board" para LibreOffice Calc.

Este proyecto es una herramienta de organización visual y gestión de tareas en formato Kanban, desarrollada íntegramente dentro de **LibreOffice Calc** mediante programación en **Basic (StarBasic)** de la API de UNO.

El sistema esta diseñado bajo una separación de responsabilidades, que se divide en un tablero visual interactivo (Capa de Presentación) y un registro plano estructurado (Capa de Datos).

## Características
- **Espacio de Trabajo Fluido y Ágil:** El tablero Kanban se actualiza e interactúa de forma inmediata al cambiar de proyecto o de columna de estado, ofreciendo una experiencia de uso ágil y sin retrasos, libre de las ralentizaciones habituales de las hojas de cálculo convencionales.
- **Visualización Dinámica de Detalles:** Con solo hacer clic en cualquier tarea del tablero, el panel lateral muestra al instante su título completo y descripción detallada, permitiendo mantener las tarjetas del Kanban cortas, limpias y legibles.
- **Captura e Interacción Intuitivas:** El formulario único para tareas se adapta visualmente al contexto de uso. Ofrece opciones simplificadas para la creación rápida y transforma sus colores y botones para alertar claramente al usuario en operaciones delicadas como el envío a la papelera o la eliminación definitiva.
- **Seguridad y Trazabilidad de sus Tareas:** Cada tarea está vinculada a un registro único y persistente en la base de datos interna. Esto garantiza que las prioridades se puedan reorganizar, subir o bajar de nivel con total seguridad de que los textos y descripciones nunca se mezclarán ni se perderán.
- **Separación de Operación y Configuración:** El sistema separa la rutina de organización diaria de la estructuración del espacio de trabajo. Mediante un panel de configuración dedicado, es posible añadir o dar de baja proyectos y columnas de forma centralizada sin interferir con la captura rápida de tareas.
- **Enfoque Local y Privacidad Total:** La herramienta es 100% autocontenida y funciona completamente fuera de línea (offline). No requiere registros, cuentas en la nube ni dependencias externas; toda su información permanece resguardada de forma privada en su computadora.

Este proyecto tiene como objetivo ofrecer una alternativa local, ligera, personalizable y orientada a la privacidad para la gestión ágil de proyectos directos, sin depender de servicios web externos ni bases de datos de terceros.

## Cómo empezar (Instalación)
1. Descarga el archivo `canban.ods` de este repositorio.
2. Abre el archivo con LibreOffice Calc.
3. Si aparece un mensaje indicando que las macros están desactivadas, consulta la sección de **Seguridad de Macros** más abajo.

## Seguridad de Macros
Para que las macros funcionen correctamente en tu equipo:
1. En LibreOffice, ve a **Herramientas > Opciones > LibreOffice > Seguridad**.
2. Haz clic en **Seguridad de macros > Orígenes de confianza**.
3. Añade la carpeta donde descargaste este archivo a la lista de **Fuentes de confianza** (recomendado), o bien establece el nivel de seguridad en **Medio**.

## Comandos Fundamentales
1. **Desactive los menus** de libreoffice calc con el comando `Ctrl + Alt + Q`.
2. Para **agregar una nueva tarea**, utilice el comando `Ctrl + Alt + N`.
3. Parar **editar una tarea existente**, utilice el comando `Ctrl + Alt + U`.
4. Para **agregar/eliminar un estado**, utilice el comando `Ctrl + Alt + E`.
5. Para **agregar/eliminar una actividad**, utilice el comando `Ctrl + Alt + A`.
6. Para **enviar a la papelera una tarea**, utilice el comando `Ctrl + Alt + P`.
7. Para **eliminar(borrar permanentemente) una tarea**, utilice el comando `Ctrl + Alt + D`.


## Cómo contribuir o modificar las macros
Si deseas realizar cambios en el código:
1. El código fuente de las macros se encuentra en la carpeta `src/macros.bas`.
2. Puedes hacer modificaciones directamente en el editor de Basic de LibreOffice Calc.
3. Al finalizar, recuerda exportar el módulo modificado y reemplazar el archivo `src/macros.bas` para mantener el repositorio actualizado.

## Licencia
Este proyecto está bajo la Licencia MIT - ver el archivo `LICENSE` para más detalles.
