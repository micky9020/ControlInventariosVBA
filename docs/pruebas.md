# Pruebas manuales

1. Ejecutar `ConfigurarSistema` y verificar que se creen las hojas `Productos`, `Movimientos` y `Dashboard`.
2. Registrar el producto `SKU-100`, stock mínimo 5, costo S/ 12.50.
3. Registrar una entrada de 20 unidades para `SKU-100`.
4. Confirmar que el reporte muestre stock actual 20.
5. Registrar una salida de 7 unidades y confirmar stock 13.
6. Intentar registrar una salida de 14 unidades: el sistema debe rechazarla por stock insuficiente.
7. Cambiar el stock mínimo de `SKU-100` a 15 y actualizar el dashboard: debe contar el producto bajo mínimo.
8. Generar el reporte y confirmar el estado `REPOSICIÓN`.
9. Revisar que cada movimiento contenga fecha/hora, tipo, SKU, cantidad, documento y responsable.
