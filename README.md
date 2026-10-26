# Control de Inventarios en Excel VBA

Proyecto de portafolio para registrar productos, controlar entradas y salidas, calcular existencias y generar alertas de reposición.

## Funcionalidades
- Catálogo de productos con SKU único, categoría, unidad, stock mínimo y costo unitario.
- Registro de entradas y salidas con fecha/hora, documento y responsable.
- Validación para impedir salidas mayores al stock disponible.
- Dashboard con productos activos, unidades disponibles, valorización estimada y productos bajo mínimo.
- Reporte de stock actual con estado `OK` o `REPOSICIÓN`.
- Cálculo del stock desde el historial de movimientos para mantener trazabilidad.

## Requisitos
- Microsoft Excel de escritorio para Windows o Mac con soporte para VBA.
- Guardar el archivo de trabajo como **Libro de Excel habilitado para macros (`.xlsm`)**.
- Habilitar macros únicamente si confías en el archivo y entiendes el código.

## Instalación
1. Descarga `Plantilla_Control_Inventarios.xlsx` y ábrela en Excel.
2. Presiona `Alt + F11` para abrir el Editor de Visual Basic.
3. En el editor, selecciona **File > Import File...** y elige `vba/modControlInventarios.bas`.
4. Vuelve a Excel y guarda una copia como `Libro de Excel habilitado para macros (*.xlsm)`.
5. Ejecuta `ConfigurarSistema` desde **Desarrollador > Macros** (o desde el editor VBA con el cursor dentro del procedimiento).
6. Para operar, ejecuta las macros `RegistrarProducto`, `RegistrarEntrada`, `RegistrarSalida`, `ActualizarDashboard` y `GenerarReporteStock`.

## Hojas
- `Productos`: maestro de artículos.
- `Movimientos`: registro histórico de entradas y salidas.
- `Dashboard`: indicadores principales.
- `ReporteStock`: se crea al generar el reporte.

## Flujo recomendado
1. Ejecuta `ConfigurarSistema`.
2. Registra los productos que administras.
3. Registra una **Entrada** inicial para cargar existencias reales. No escribas el stock directamente en la hoja de productos: el stock se calcula a partir de movimientos.
4. Registra cada recepción o despacho como movimiento.
5. Genera el reporte y revisa productos marcados para reposición.

## Modelo de datos
`Productos.SKU` es el identificador único del artículo. `Movimientos` es un kardex simplificado: cada entrada suma y cada salida resta. El costo utilizado para la valorización es el costo unitario registrado en el maestro; no implementa FIFO ni promedio ponderado.

## Limitaciones y siguientes mejoras
- Es una solución local, para un usuario o equipo con un archivo compartido cuidadosamente administrado.
- No incluye control de lotes, vencimientos, series, ubicaciones múltiples, órdenes de compra, perfiles de usuario ni auditoría protegida contra modificaciones manuales.
- El archivo `.xlsx` de plantilla no contiene macros; el código VBA se importa por separado y luego se guarda como `.xlsm`.
- Antes de usarlo en producción, validar el código con datos de prueba, definir copias de seguridad y restringir la edición directa del historial.

## Estructura del repositorio
```text
Control_Inventarios_VBA/
├── Plantilla_Control_Inventarios.xlsx
├── README.md
├── LICENSE
├── .gitignore
├── vba/
│   └── modControlInventarios.bas
└── docs/
    └── pruebas.md
```

## Pruebas manuales sugeridas
Ver [`docs/pruebas.md`](docs/pruebas.md).
