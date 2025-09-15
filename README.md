# vbscripts

Colección de scripts en VBScript para mantenimiento y actualización de sistemas internos.

## Tabla de contenido

- [Descripción](#descripción)  
- [Archivos incluidos](#archivos-incluidos)  
- [Requisitos](#requisitos)  
- [Uso](#uso)  
- [Configuración](#configuración)  
- [Convenciones de nombres](#convenciones-de-nombres)  
- [Ejemplos](#ejemplos)  
- [Contribuciones](#contribuciones)  
- [Licencia](#licencia)

---

## Descripción

Este repositorio alberga varios scripts VBScript (.vbs) que realizan tareas automáticas como:

- Actualización de APIs / aplicaciones (`updateapi.vbs`, `updateapp.vbs`, etc.)  
- Truncado o limpieza de datos (`truncate_data.vbs`)  
- Restauraciones o restaurar servicios (`uprest.vbs`, `updaterest.vbs`)  
- Otros scripts puntuales relacionados con “scale”, web, terminaciones de día, etc.

---

## Archivos incluidos

| Nombre | Función estimada |
|---|-------------------|
| `updateapi.vbs` / `updscale.vbs` | Actualiza endpoints, interfaces o servicios relacionados con API / escalas. |
| `truncate_data.vbs` | Limpia tablas o datos antiguos. |
| `uprest.vbs` / `updaterest.vbs` | Script para restauración. |
| `actendday.vbs` | Tarea al final del día (“act end day”) posiblemente cierre o backup. |
| `output.log` | Registro de salidas / errores de los scripts. |
| `.vscode/` | Configuración del editor VSCode. |

---

## Requisitos

- Sistema operativo Windows con soporte VBScript.  
- Permisos adecuados para ejecutar scripts en el sistema / acceder a archivos o recursos que los scripts modifiquen.  
- Dependencias externas si los scripts llaman a otros servicios, APIs o bases de datos (no todas documentadas).  
- Programador de tareas (Task Scheduler) u otra herramienta para automatizar ejecución si se requiere.

---

## Uso

1. Colocar los scripts en una carpeta accesible por el sistema.  
2. Verificar que los parámetros internos del script estén correctos (rutas, credenciales, endpoints, nombres de servidores).  
3. Probar ejecución manual de un script para verificar que funciona sin errores.  
4. Si se requiere ejecución automática, configurarlo en el Programador de tareas de Windows.

---

## Configuración

- Editar los scripts para ajustar rutas locales, direcciones de API, nombres de servidor, credenciales.  
- Establecer permisos de lectura/escritura donde se aloje el script.  
- Si se utiliza `output.log`, asegurarse de que la ruta es accesible y que existe o que el script puede crear el archivo.

---

## Convenciones de nombres

- Prefijo `update` o `upd` → tareas de actualización.  
- Prefijo `truncate` → limpieza de datos.  
- Prefijo `act` → acciones rutinarias (ej. al fin de día).  
- `rest` → restauración.  

---

## Ejemplos

```powershell
# Ejecutar manualmente un script
cscript //nologo updateapi.vbs

# Ver log generado
type output.log
