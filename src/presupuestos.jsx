import { useState, useRef, useCallback, useEffect, Component } from "react";
// ─────────────────────────────────────────────────────────────────────
// Componente Presupuestos
// Versión: v2.00.0 (9 Junio 2026)
//
// Convención SemVer:
//   - MAJOR: cambios incompatibles
//   - MINOR: nueva funcionalidad
//   - PATCH: corrección de errores
//
// Histórico reciente:
//   v2.00.0 (9 Junio 2026) - Nuevo: Descuentos → Gestionar Estrategias Descuento (CRUD de estrategias de la tabla descuentos/detalledescuentos + aplicar al presupuesto por grupo). Quitado "Gestionar Descuentos" de Otros
//   v1.99.1 (9 Junio 2026) - Asistente Referencias: opciones ordenadas por posición y, dentro de cada posición, por referencia (posición -1 al final)
//   v1.99.0 (9 Junio 2026) - Asistente Referencias: búsqueda automática en BD al montar la referencia; lista de productos cuya referencia empieza por la montada (como el autocompletado del grid); se elige uno e inserta
//   v1.98.2 (9 Junio 2026) - Asistente Referencias: la opción SOBREESCRIBE desde la posición indicada (no inserta); referencia más grande en la lista de productos
//   v1.98.1 (9 Junio 2026) - Asistente Referencias: posición -1 inserta la opción al final de la referencia (admite varias)
//   v1.98.0 (9 Junio 2026) - Nueva funcionalidad Productos → Asistente Referencias: lista productos+opciones, monta la referencia insertando opciones por posición, consulta productos en BD e inserta en el presupuesto
//   v1.97.0 (8 Junio 2026) - Excel Imprimir: pestaña "Condiciones Comerciales" con las condiciones particulares de suministro (negrita en títulos, hipervínculos clicables)
//   v1.96.0 (8 Junio 2026) - Error Boundary: un fallo de render ya no deja la app en blanco (muestra aviso + Reintentar). commitEdit más defensivo (captura la celda y try/catch) para el blur del editor multilínea
//   v1.95.0 (8 Junio 2026) - Etiqueta de estructura siempre visible: verde "Estructura activa" / gris "Estructura desactivada", clicable para alternar
//   v1.94.4 (8 Junio 2026) - Excel Imprimir: título primera fila "Rev.N" → "Revisión N"
//   v1.94.3 (8 Junio 2026) - Nombre fichero Imprimir: "Rev.2" → "Rev 2" (punto por espacio)
//   v1.94.2 (8 Junio 2026) - Leer Presupuesto: importe total y netos con punto de miles también en 4 cifras (useGrouping always)
//   v1.94.1 (8 Junio 2026) - Diálogo Leer Elemento también redimensionable
//   v1.94.0 (8 Junio 2026) - Diálogos Leer Presupuesto, Leer Producto, Gestionar Clientes y Gestionar Contactos redimensionables (arrastrando la esquina inferior derecha)
//   v1.93.0 (7 Junio 2026) - Todos los recuadros de log (progreso) tienen ahora botón Exportar CSV además de Limpiar log (helper exportarLogCSV)
//   v1.92.0 (7 Junio 2026) - Ayuda actualizada: menús al día (conf/TP, Guardar Producto con actualización, contactos con teléfonos, login, formato numérico, etc.)
//   v1.91.1 (7 Junio 2026) - Mantenimiento BD tabla Contactos: importa también Teléfono 1 y Teléfono 2 desde Excel
//   v1.91.0 (7 Junio 2026) - Ficha contactos: campos Teléfono 1 y Teléfono 2 (formulario y tabla)
//   v1.90.1 (7 Junio 2026) - Ficha clientes: CP también con 5 cifras al editar la celda
//   v1.90.0 (7 Junio 2026) - Ficha clientes: CP siempre con 5 cifras (ceros a la izquierda, 1250 → 01250)
//   v1.89.3 (7 Junio 2026) - Limpieza: quitados los logs de diagnóstico temporal de Guardar Producto
//   v1.89.2 (7 Junio 2026) - Diagnóstico temporal en Guardar Producto (logs en consola) para depurar por qué solo procesa 1 fila
//   v1.89.1 (7 Junio 2026) - Fix Guardar Producto: el rango seleccionado usaba row1/row2 inexistentes (NaN); solo procesaba 1 fila. Ahora usa startRowIdx/endRowIdx
//   v1.89.0 (7 Junio 2026) - Guardar Producto: si el producto existe, diálogo para elegir columnas a actualizar (PVP/desc/nombre/coste/grupo) + confirmación por fila SÍ/NO
//   v1.88.1 (6 Junio 2026) - Excel Imprimir: columna I ancho 18 y texto centrado
//   v1.88.0 (6 Junio 2026) - Excel Imprimir: representación "TP" (Total Posición) → cantidad 1 y neto unitario = total posición (solo en la impresión)
//   v1.87.0 (6 Junio 2026) - Excel Imprimir: representación "conf" aplica estilo CONF a la fila y escribe "A confirmar por el cliente" en columna I (justificado)
//   v1.86.0 (6 Junio 2026) - Configurar Estilos: nuevos estilos PD (producto normal) y CONF (Confirmar por el cliente), configurables
//   v1.85.0 (6 Junio 2026) - PVP en rojo si fechapvp > 1 año (al buscar datos/leer producto); tooltips de fechapvp y fechapreciocoste; se limpia al aplicar estructura
//   v1.84.3 (5 Junio 2026) - Nombre de fichero Excel de Exportar y Formato Simple Quote usa numerocompleto (no número)
//   v1.84.2 (5 Junio 2026) - Leer Producto: la celda Referencia recorta con ... y tooltip si el texto no cabe (no se sale sobre otras columnas)
//   v1.84.1 (5 Junio 2026) - Fix build: celda Descripción en Leer Producto había perdido su <td> de apertura (error JSX)
//   v1.84.0 (5 Junio 2026) - Leer Producto: columnas redimensionables arrastrando los separadores de la cabecera
//   v1.83.1 (5 Junio 2026) - Redimensión de columnas: zona de agarre más ancha y visible (azul al pasar), doble clic restablece ancho
//   v1.83.0 (5 Junio 2026) - Leer Producto: nueva columna "Grupo Dto." con el nombre del grupo descuento (no el id)
//   v1.82.0 (5 Junio 2026) - Grid: tooltip con descripción en celdas de grupodescuento/familia/subfamilia con contenido
//   v1.81.0 (5 Junio 2026) - Comprobar Celda: muestra nº caracteres + color fondo/tinta de la celda; quitado "Ver Familia" del menú
//   v1.80.0 (5 Junio 2026) - Borrar filas con 0: usa la columna seleccionada (una sola); borra filas con 0 en esa columna
//   v1.79.1 (5 Junio 2026) - Tooltip de "Calcular Descuento" más específico
//   v1.79.0 (4 Junio 2026) - Login obligatorio para acceder; bloqueo de 10 min tras 5 intentos fallidos
//   v1.78.0 (4 Junio 2026) - Leer Presupuesto: añade 10 filas en blanco al final del presupuesto cargado
//   v1.77.1 (4 Junio 2026) - Excel Imprimir: nombre del fichero usa numerocompleto + Rev.
//   v1.77.0 (4 Junio 2026) - Apartado: al activar estructura se persiste la numeración (1.1, 1.1.2) en posicion y se vacía en S1-S4/TT/CM; al desactivar se mantiene
//   v1.76.2 (4 Junio 2026) - Excel Imprimir: el nº de presupuesto usa numerocompleto + revisión (antes np)
//   v1.76.1 (4 Junio 2026) - Excel Imprimir: formato moneda € corregido (#,##0.00 €) en Neto Unitario/Posición
//   v1.76.0 (4 Junio 2026) - Excel Imprimir: etiquetas en col D; columnas B/C/E centradas; Neto Unitario/Posición en formato moneda €
//   v1.75.0 (4 Junio 2026) - Barra inferior: muestra nº de celdas seleccionadas (rectángulo azul) y suma de sus valores numéricos (estilo Excel)
//   v1.74.1 (4 Junio 2026) - Tooltips también en los menús de nivel superior (Presupuesto, Celdas, Productos, etc.)
//   v1.74.0 (4 Junio 2026) - Tooltips descriptivos (title) en todos los items de menú/submenú y en el menú de Opciones
//   v1.73.2 (4 Junio 2026) - Fix formato: punto de miles también en números de 4 cifras (useGrouping "always" en fmt/fmtEur)
//   v1.73.1 (4 Junio 2026) - Grid: al editar importes/dto se acepta coma o punto como decimal (se normaliza a número)
//   v1.73.0 (4 Junio 2026) - Grid: todos los numéricos con coma decimal; importes en euros con € y punto de miles (con o sin estructura)
//   v1.72.4 (4 Junio 2026) - Buscar Ref SIEMENS: ignora espacios; detecta refs pegadas sin guiones (bloque base no bloquea la versión larga)
//   v1.72.3 (4 Junio 2026) - Buscar Ref SIEMENS: ignora espacios al detectar (refs partidas por espacios) y mapea posiciones al texto original para el resaltado
//   v1.72.2 (4 Junio 2026) - Fix Buscar Ref SIEMENS: la última referencia aceptada se perdía (estado asíncrono); ahora se aplica correctamente
//   v1.72.1 (4 Junio 2026) - Si la API local no responde al iniciar, muestra diálogo de aviso que el usuario debe aceptar
//   v1.72.0 (4 Junio 2026) - Al iniciar comprueba que la API local está viva (GET /); indicador verde/rojo en barra de estado
//   v1.71.1 (4 Junio 2026) - Fix: escapeEditRef estaba declarado en ClientesDialog en vez de App; el doble clic en celda crasheaba
//   v1.71.0 (4 Junio 2026) - Edición de celda: salir/cambiar de celda confirma el valor (onBlur guarda); solo Escape descarta
//   v1.70.0 (4 Junio 2026) - Nueva fila: inserta encima de la celda seleccionada (rectángulo azul); error si no hay celda seleccionada
//   v1.69.1 (3 Junio 2026) - Estilos por defecto actualizados (verde/rojo/naranja claros, fontSize 11) según fichero del usuario
//   v1.69.0 (3 Junio 2026) - Mantenimiento BD: apartado "Comprobar integridad datos en BD" (detecta registros huérfanos vía /integridad/comprobar)
//   v1.68.3 (3 Junio 2026) - Mantenimiento Clientes/Contactos: recuadro "Mapeo de columnas detectado" con tags verde/rojo (estilo Tarifas)
//   v1.68.2 (3 Junio 2026) - Import XLSX desde "xlsx-js-style" (coincide con package.json y aplica estilos en exportación Excel)
//   v1.68.1 (3 Junio 2026) - Import Clientes: razón social O nombre común (al menos uno); busca por razón social; si solo razón social → nombre común = razón social
//   v1.68.0 (3 Junio 2026) - Clientes: nuevas columnas nif(varchar) y telefono1; ifa ahora integer; cp integer. UI + importación Excel
//   v1.67.2 (3 Junio 2026) - Mantenimiento Clientes/Contactos: log de progreso detallado siempre visible (estilo Tarifas)
//   v1.67.1 (3 Junio 2026) - Mantenimiento Clientes/Contactos: selector "Actualizar existentes" con el mismo estilo toggle que Tarifas
//   v1.67.0 (3 Junio 2026) - Mantenimiento BD: apartados "Mantenimiento tabla Clientes" y "Contactos" (importar Excel + actualizar SI/NO + log)
//   v1.66.0 (2 Junio 2026) - Gestión Clientes: botón Borrar cliente (comprueba uso, confirma, DELETE); nuevos sin guardar se descartan local
//   v1.65.1 (2 Junio 2026) - Fix Gestión Clientes: datalist provincias único (no duplicado por celda) + reset selección tras guardar
//   v1.65.0 (2 Junio 2026) - Gestión Clientes: campo Provincia como desplegable con autocompletado (datalist), muestra nombre y guarda id
//   v1.64.2 (2 Junio 2026) - Crear SimpleQuote: quitada la línea "Indicamos que el cliente es..."
//   v1.64.1 (2 Junio 2026) - Crear SimpleQuote: añade "De: <email contacto>" al final del cuerpo si el presupuesto tiene contacto
//   v1.64.0 (2 Junio 2026) - Fondo bienvenida: ruta /inicio.png (minúscula) para coincidir con el fichero renombrado
//   v1.63.9 (2 Junio 2026) - Aplicar descuentos por Grupo: admite descuentos negativos (-100 a 100)
//   v1.63.8 (2 Junio 2026) - Selector País cliente final: quitadas banderas emoji, solo código + nombre
//   v1.63.7 (2 Junio 2026) - Fondo bienvenida: ruta /Inicio.png (mayúscula) para coincidir con el fichero en Linux/Vercel
//   v1.63.6 (2 Junio 2026) - Eliminado alias 'Image as ImageIcon' sin usar del import lucide-react
//   v1.63.5 (2 Junio 2026) - Limpieza de imports/variables sin usar (iconos lucide, parsePastedTSV) para build Vercel
//   v1.63.4 (2 Junio 2026) - Fix build: setEstilosLocal→setDraft y state resultado faltante en GuardarElementoDialog
//   v1.63.3 (2 Junio 2026) - Quitados comentarios eslint-disable que rompían el build en Vercel (react-scripts)
//   v1.63.2 (2 Junio 2026) - Selector País cliente final: emoji de bandera a partir del código ISO2
//   v1.63.1 (2 Junio 2026) - Fix: importar icono Upload de lucide-react (Configuraciones Varias)
//   v1.63.0 (2 Junio 2026) - Opciones → "Configuraciones Varias" con apartado Crear SimpleQuote (destinatario/CC) + guardar/leer fichero
//   v1.62.2 (1 Junio 2026) - Fix Damex E: valor_pedido tomado de calcTotalPresupuesto().neto (devolvía objeto, no número)
//   v1.62.1 (1 Junio 2026) - Guardar Producto: verifica que la respuesta es realmente del producto antes de marcarlo como "ya existe"
//   v1.62.0 (1 Junio 2026) - codigoUsuario ahora se vincula al usuario logado (codigopresupuestos del POST /usuarios/login)
//   v1.61.2 (1 Junio 2026) - Botón calculadora nº completo: muestra en estado el nuevo valor calculado
//   v1.61.1 (1 Junio 2026) - Damex E: éxito solo si resultado="NN"; otra combinación se notifica como error
//   v1.61.0 (1 Junio 2026) - Hacer Damex E: llama a /damex/ejecutar con título, cliente, país y total del presupuesto
//   v1.60.5 (1 Junio 2026) - API local unificada en 127.0.0.1:8000 (PMD + crear_sq) vía constante API_LOCAL_URL
//   v1.60.4 (1 Junio 2026) - Fix: añadido el componente SelectorClienteDialog que faltaba
//   v1.60.3 (1 Junio 2026) - Botón calculadora junto a Nº Completo: recalcula con codigoUsuario + nº + año
//   v1.60.2 (1 Junio 2026) - Diálogo Sobrescribir: usa numerocompleto (no solo número)
//   v1.60.1 (1 Junio 2026) - Comprobación con BD: diferencias muestran fila del presupuesto actual (no del filtrado)
//   v1.60.0 (1 Junio 2026) - Cabecera: campo Cliente con diálogo selector (similar a "A la atención de")
//   v1.59.4 (1 Junio 2026) - Fix Leer Presupuestos: quitada la cabecera PVP residual de la tabla detalle
//   v1.59.3 (1 Junio 2026) - Guardar Presupuesto: valida que haya cliente seleccionado antes de enviar
//   v1.59.2 (1 Junio 2026) - Fix Guardar: idposicion vacío se envía como null (es INTEGER en BD)
//   v1.59.1 (1 Junio 2026) - Guardar presupuesto envía totalpresupuesto, cliente final, país y netos calculados al backend
//   v1.59.0 (1 Junio 2026) - Cabecera presupuesto: campos "País cliente final" (selector) y "Cliente final" (texto)
//   v1.58.8 (1 Junio 2026) - Leer Presupuestos detalle: Neto Pos. se lee directamente del campo BD precionetoposicion
//   v1.58.7 (31 Mayo 2026) - Leer Presupuestos: lista con columna Total; detalle quita PVP, añade Neto Pos., Referencia más estrecha
//   v1.58.6 (31 Mayo 2026) - Guardar Presupuesto: identifica el presupuesto en BD por numerocompleto + revision
//   v1.58.5 (31 Mayo 2026) - Productos: eliminados también los handlers de las 8 opciones quitadas del menú
//   v1.58.4 (31 Mayo 2026) - Productos: eliminadas 8 opciones no implementadas (Precio Neto, Costes, etc.)
//   v1.58.3 (31 Mayo 2026) - Leer Precios PMD: ya no muestra datos en el diálogo, los aplica directo a la fila y cierra
//   v1.58.2 (31 Mayo 2026) - Crear SimpleQuote: cuerpo en texto plano con columnas alineadas (Outlook-friendly)
//   v1.58.1 (31 Mayo 2026) - Leer Precios de PMD: API local en puerto 8001 (no 8000)
//   v1.58.0 (31 Mayo 2026) - Leer Precios de PMD: diálogo con consulta a API local (a presupuesto / a CSV)
//   v1.58.0 (31 Mayo 2026) - Crear SimpleQuote: POST a API local /crear_sq con tabla de productos del presupuesto
//   v1.57.0 (31 Mayo 2026) - Nuevos items: Presupuesto → Crear SimpleQuote / Hacer Damex E; Productos → Leer Precios de PMD (placeholders)
//   v1.56.9 (30 Mayo 2026) - Ayuda: sección Novedades actualizada con todos los cambios recientes
//   v1.56.8 (30 Mayo 2026) - Etiquetas S1-S4/TT se persisten en rows al activar estructura (no se revierten al desactivar)
//   v1.56.7 (30 Mayo 2026) - Etiqueta de S1-S4/TT siempre sobrescribe el campo nombre con estructura activa
//   v1.56.6 (30 Mayo 2026) - Etiqueta automática Sn = "TOTAL <texto del Tn anterior>"; TT = "TOTAL <título presupuesto>"
//   v1.56.5 (30 Mayo 2026) - Etiqueta automática "Subtotal N"/"Total" en filas S1-S4/TT con estructura activa
//   v1.56.4 (30 Mayo 2026) - Fix Tarifas: procesa todas las columnas detectadas (pvp_l1, plazoentrega, etc.)
//   v1.56.3 (30 Mayo 2026) - Tarifas: matching por prefijo del nombre real de la columna BD (sin aliases)
//   v1.56.2 (30 Mayo 2026) - Fix Tarifas: preciocoste y grupodescuentospain ahora se reconocen correctamente
//   v1.56.1 (30 Mayo 2026) - Botones de la barra del grid un poco más altos (padding 7px en vez de 4px)
//   v1.56.0 (30 Mayo 2026) - Botones "Copiar filas" y "Pegar filas" en la barra del grid
//   v1.55.3 (30 Mayo 2026) - Leer Producto inserta usando la celda activa: rellena si fila vacía, si no inserta encima
//   v1.55.2 (30 Mayo 2026) - Descuento negativo (recargo) también se muestra en rojo con estructura activa
//   v1.55.1 (30 Mayo 2026) - Calcular Descuento admite negativos; importes negativos en rojo con estructura activa
//   v1.55.0 (30 Mayo 2026) - Calcular Descuento: calcula el dto de cada fila para que el PVP×(1-dto) iguale el valor de la celda seleccionada
//   v1.54.0 (30 Mayo 2026) - Celdas Naturaleza/Representación: input editable con filtro y validación contra lista
//   v1.53.4 (30 Mayo 2026) - Wrap multilínea de ref/producto/desc se mantiene al desactivar estructura
//   v1.53.3 (30 Mayo 2026) - Fix: reset CSS de html/body para eliminar scroll global de la página
//   v1.53.2 (30 Mayo 2026) - Fix: textarea de edición ya no se corta (overflow:visible en celdas de texto largo)
//   v1.53.1 (30 Mayo 2026) - Activar estructura sobreescribe anchos manuales con los auto recalculados
//   v1.53.0 (30 Mayo 2026) - Anchos/altos persistentes y redimensionables a mano (estilo Excel)
//   v1.52.2 (30 Mayo 2026) - Editor de celda: rows auto según contenido (saltos de línea + ancho)
//   v1.52.1 (30 Mayo 2026) - Fix ancho auto Neto Unit./Pos. (columnas calc): formateado real con € y %
//   v1.52.0 (30 Mayo 2026) - Celdas → "Juntar celdas en una": concatena texto de varias celdas en la primera
//   v1.51.4 (29 Mayo 2026) - Fix: paste de números con separador de miles ("3.948, 45 €" → 3948.45)
//   v1.51.3 (29 Mayo 2026) - Presupuesto en blanco se crea con 50 filas en vez de 10
//   v1.51.2 (29 Mayo 2026) - Buscar datos por Referencia: fuerza naturaleza=PD y cantidad mínima 1
//   v1.51.1 (29 Mayo 2026) - Mantenimiento Grupos Descuento: selector SÍ/NO en vez de checkbox (estilo Tarifas)
//   v1.51.0 (29 Mayo 2026) - Mantenimiento Grupos Descuento: recuadro de "Mapeo de columnas detectado" con tags y errores (estilo Tarifas)
//   v1.50.5 (29 Mayo 2026) - Mantenimiento Grupos Descuento: reconoce DGL1 y DGL2 al importar Excel
//   v1.50.4 (28 Mayo 2026) - Mantenimiento Grupos Descuento: soporta importar DGL1 y DGL2 desde Excel
//   v1.50.3 (28 Mayo 2026) - Aplicar descuentos: DGL1 ahora en amarillo (no naranja)
//   v1.50.2 (28 Mayo 2026) - Aplicar descuentos: dto actual en naranja si > DGL1, rojo si > DGL2
//   v1.50.1 (28 Mayo 2026) - Aplicar descuentos: añadidas columnas DGL1 y DGL2 desde gruposdescuento
//   v1.50.0 (28 Mayo 2026) - Edición de celda: textarea flotante más alto para texto largo + selección/copiar/pegar dentro del campo
//   v1.49.2 (28 Mayo 2026) - Fix: IIFE del display de celda mal envuelta en llaves dentro del ternario
//   v1.49.1 (28 Mayo 2026) - Fix: sintaxis del map de cabeceras th (IIFE mal envuelta)
//   v1.49.0 (28 Mayo 2026) - Estructura: anchos auto, multilínea ref/producto/desc, apartados sin 0, TT total, sin límite 50 en Producto
//   v1.48.1 (28 Mayo 2026) - Aplicar descuentos: muestra el dto leído directo (no calculado)
//   v1.48.0 (28 Mayo 2026) - Aplicar descuentos ahora agrupa por Grupo Descuento + Dto actual (no por subfamilia)
//   v1.47.0 (27 Mayo 2026) - Ayuda: añadida sección "Novedades recientes" con resumen por áreas
//   v1.46.1 (27 Mayo 2026) - Fix LoginDialog: iconos duplicados fuera del input absoluto
//   v1.46.0 (27 Mayo 2026) - Mantenimiento BD Subfamilia con editar/borrar (check gruposdescuento)
//   v1.45.0 (27 Mayo 2026) - Mantenimiento BD Familia: diálogo con editar/borrar (con check de uso)
//   v1.44.0 (27 Mayo 2026) - Pastilla usuario clickable con menú Iniciar/Cerrar sesión + LoginDialog reutilizable
//   v1.43.3 (27 Mayo 2026) - Pastilla siempre visible (solo icono User) hasta login OK
//   v1.43.2 (27 Mayo 2026) - Usuario por defecto vacío; pastilla solo se ve tras login OK
// ─────────────────────────────────────────────────────────────────────
import * as XLSX from "xlsx-js-style"; // Fork de SheetJS con soporte de estilos de celda (colores, negrita, bordes)
import { FileText, FolderOpen, Download, Upload, Printer, BarChart3, Palette, Grid3x3, Search, Eraser, MoreHorizontal, Calculator, Link2, Eye, Trash2, X, Scale, Square, MessageSquare, Plus, FileInput, Edit3, TrendingUp, Scissors, CornerDownLeft, DollarSign, Database, Repeat, Bot, HelpCircle, Settings, Percent, Users, Target, Hash, Save, RefreshCw, Home, FileSpreadsheet, MousePointer, Layers, Package, Wrench, ArrowLeft, Check, Copy, FileUp, ClipboardCheck, User, Lock, LogIn, LogOut } from "lucide-react";

// Helper para iconos outline pequeños del menú
const Icon = ({ as: Component, size = 14, color = "currentColor" }) => Component ? (
  <Component size={size} color={color} strokeWidth={1.6} style={{ flexShrink: 0 }} />
) : null;

// Helper para descargar Excel en cualquier entorno (CodeSandbox, navegador, etc.)
function descargarXLSX(wb, nombreFichero) {
  // Generar buffer del workbook en memoria
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  // Crear blob y forzar descarga
  const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombreFichero;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 100);
}

// Helper: contenido de botón con icono + texto alineados
const BtnContent = ({ icon, children, iconColor, iconSize = 14 }) => (
  <span style={{ display: "inline-flex", alignItems: "center", gap: 6, justifyContent: "center" }}>
    <Icon as={icon} size={iconSize} color={iconColor || "currentColor"} />
    <span>{children}</span>
  </span>
);

// URL base de la API en Render. Aparte de esto, hay una API local en 127.0.0.1:8000 (ver API_LOCAL_URL)
const API_URL = "https://presupuestos-api-x11d.onrender.com"
// URL de la API local (CdM_Presupuestos_API_local.py): unifica funciones de PMD y crear_sq
const API_LOCAL_URL = "http://127.0.0.1:8000";

// ── Pantalla de Opciones ──
const FUENTES_DISPONIBLES = [
  "Segoe UI", "Arial", "Helvetica", "Calibri", "Times New Roman",
  "Georgia", "Verdana", "Tahoma", "Trebuchet MS", "Courier New",
  "Consolas", "Monaco", "Comic Sans MS",
];

const NATURALEZAS_CON_ESTILO = [
  { key: "T1", label: "T1 - Título Nivel 1" },
  { key: "T2", label: "T2 - Título Nivel 2" },
  { key: "T3", label: "T3 - Título Nivel 3" },
  { key: "T4", label: "T4 - Título Nivel 4" },
  { key: "S1", label: "S1 - Subtotal Nivel 1" },
  { key: "S2", label: "S2 - Subtotal Nivel 2" },
  { key: "S3", label: "S3 - Subtotal Nivel 3" },
  { key: "S4", label: "S4 - Subtotal Nivel 4" },
  { key: "TT", label: "TT - Total Presupuesto" },
  { key: "CM", label: "CM - Comentario" },
  { key: "VERDE", label: "VERDE - Línea resaltada verde" },
  { key: "GRIS",  label: "GRIS - Línea resaltada gris" },
  { key: "PD",    label: "PD - Producto normal" },
  { key: "CONF",  label: "CONF - Confirmar por el cliente" },
];

const OPCIONES_MENU = [
  { id: "estilos",       label: "Configurar Estilos", icon: Palette, tooltip: "Define colores y fuentes de cada naturaleza (T1-T4, S1-S4, etc.)" },
  { id: "tarifas",       label: "Actualizar Tarifas", icon: TrendingUp, tooltip: "Importa precios y datos de productos desde un Excel de tarifa" },
  { id: "mantenimiento", label: "Mantenimiento BD",   icon: Database, tooltip: "Mantiene tablas (productos, clientes, contactos) y comprueba integridad" },
  { id: "usuarios",      label: "Gestión de Usuarios", icon: Users, tooltip: "Crea usuarios, contraseñas y permisos de la aplicación" },
  { id: "varias",        label: "Configuraciones Varias", icon: Settings, tooltip: "Ajustes de SimpleQuote (destinatario/CC) y otras opciones" },
];

// ── Sección Actualizar Tarifas dentro de Opciones ──
// Columnas válidas para tarifa
// Columnas reales de la tabla productos que la importación de tarifas puede actualizar.
// El matching se hace por prefijo del nombre real (sin aliases).
// Ej: "pla" → "plazoentrega" (único campo que empieza por "pla")
//     "precio" → "preciocoste" (único campo de la tabla que empieza por "precio"; "pvp" no empieza por "precio")
//     "fecha" → AMBIGUO (fechapvp y fechapreciocoste) → error
const TARIFA_COLS_VALIDAS = [
  { key: "referencia" },
  { key: "nombre" },
  { key: "descripcion" },
  { key: "pvp" },
  { key: "preciocoste" },
  { key: "subfamilia" },
  { key: "familia" },
  { key: "grupodescuento" },
  { key: "plazoentrega" },
  { key: "imagen" },
  { key: "masusado" },
  { key: "fechapvp" },
  { key: "fechapreciocoste" },
  { key: "pvp_l1" },
];

function matchTarifaHeaders(headers) {
  const mapping = {};
  const errores = [];
  // Normaliza: minúsculas, sin acentos, sin espacios (pegamos para comparar nombres tipo BD)
  const normalizar = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "").trim();
  headers.forEach((header, idx) => {
    const norm = normalizar(header);
    if (!norm) return;
    // Coincidencia exacta tiene prioridad
    const exacta = TARIFA_COLS_VALIDAS.find(col => normalizar(col.key) === norm);
    if (exacta) { mapping[idx] = exacta.key; return; }
    // Si no, buscar campos cuyo nombre real EMPIECE por el texto de la cabecera
    const candidatos = TARIFA_COLS_VALIDAS.filter(col => normalizar(col.key).startsWith(norm));
    if (candidatos.length === 1) mapping[idx] = candidatos[0].key;
    else if (candidatos.length > 1) errores.push(`Columna "${header}" es ambigua, podría ser: ${candidatos.map(c => c.key).join(", ")}`);
    else errores.push(`Columna "${header}" no se reconoce`);
  });
  return { mapping, errores };
}

function TarifasSection({ setStatus }) {
  const [fileData, setFileData] = useState(null); // { headers, rows }
  const [mapping, setMapping] = useState({});
  const [erroresMapeo, setErroresMapeo] = useState([]);
  const [log, setLog] = useState([]);
  const [procesando, setProcesando] = useState(false);
  const [resumen, setResumen] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [subfamilias, setSubfamilias] = useState([]);
  const [sobreescribir, setSobreescribir] = useState(true);
  const logRef = useRef(null);
  const stopRequested = useRef(false);

  // Cargar subfamilias al montar
  useEffect(() => {
    fetch(`${API_URL}/subfamilias/`)
      .then(r => r.ok ? r.json() : [])
      .then(setSubfamilias)
      .catch(() => setSubfamilias([]));
  }, []);

  // Auto-scroll al final del log
  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight;
  }, [log]);

  const handleFile = async (file) => {
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    procesar(aoa, file.name);
  };

  
  const procesar = (aoa, origen) => {
    if (aoa.length === 0) { setErroresMapeo(["El fichero está vacío"]); return; }
    const headers = aoa[0].map(h => String(h || "").trim());
    const rows = aoa.slice(1).filter(r => r.some(c => String(c || "").trim() !== ""));
    const { mapping, errores } = matchTarifaHeaders(headers);
    // Validar columnas: solo la referencia es obligatoria.
    // Para crear productos nuevos hace falta también descripcion y grupodescuento,
    // pero esa validación se hace por fila durante el procesado en aplicarTarifa.
    const tieneRef = Object.values(mapping).includes("referencia");
    const numColumnas = Object.keys(mapping).length;
    const erroresExtra = [];
    if (!tieneRef) erroresExtra.push("Falta columna 'Referencia' (obligatoria)");
    if (numColumnas < 2) erroresExtra.push("El fichero debe tener al menos 2 columnas (referencia + otra)");
    setFileData({ headers, rows, origen });
    setMapping(mapping);
    setErroresMapeo([...errores, ...erroresExtra]);
    setLog([]);
    setResumen(null);
  };

  const addLog = (texto, tipo = "info") => {
    setLog(l => [...l, { texto, tipo, hora: new Date().toLocaleTimeString("es-ES") }]);
  };

  const aplicarTarifa = async () => {
    if (!fileData || erroresMapeo.length > 0 || procesando) return;

    // Validar columnas del Excel
    const idxPorClave = {};
    Object.entries(mapping).forEach(([idx, key]) => { idxPorClave[key] = Number(idx); });
    const columnasDetectadas = Object.keys(idxPorClave);

    if (!columnasDetectadas.includes("referencia")) {
      addLog("ERROR: el fichero debe tener una columna 'Referencia'. Se cancela el proceso.", "error");
      setStatus && setStatus("Falta columna 'Referencia'", "error");
      return;
    }
    if (columnasDetectadas.length < 2) {
      addLog("ERROR: el fichero debe tener al menos 2 columnas (referencia + otra). Se cancela el proceso.", "error");
      setStatus && setStatus("El fichero debe tener al menos 2 columnas", "error");
      return;
    }

    stopRequested.current = false;
    setProcesando(true);
    setLog([]);
    setResumen(null);
    setStatus && setStatus("Procesando tarifa...", "working");
    addLog(`Procesando fichero "${fileData.origen || "(desconocido)"}"`, "info");
    addLog(`Iniciando: ${fileData.rows.length} filas. Columnas detectadas: ${columnasDetectadas.join(", ")}`, "info");

    // Cargar grupos descuento para resolver código → id (si la columna grupodescuento está)
    let gruposDesc = [];
    if (idxPorClave.grupodescuento !== undefined) {
      try {
        const r = await fetch(`${API_URL}/gruposdescuento/`);
        if (r.ok) {
          gruposDesc = await r.json();
          addLog(`Cargados ${gruposDesc.length} grupos descuento de BD`, "info");
        }
      } catch (e) {
        addLog("AVISO: no se pudieron cargar grupos descuento de BD: " + e.message, "warning");
      }
    }

    // Helper: resolver código de grupodescuento ("GD123") → id numérico
    const resolverIdGrupo = (codigo) => {
      if (!codigo) return null;
      const c = String(codigo).toUpperCase().trim();
      const g = gruposDesc.find(x => x.grupodescuentospain && String(x.grupodescuentospain).toUpperCase().trim() === c);
      return g ? g.id : undefined; // undefined si no se encuentra
    };

    let actualizados = 0, nuevos = 0, errores = 0;

    for (let i = 0; i < fileData.rows.length; i++) {
      if (stopRequested.current) {
        addLog(`⏹ Detenido por el usuario en la fila ${i + 1} de ${fileData.rows.length}`, "warning");
        break;
      }
      const row = fileData.rows[i];
      const referencia = String(row[idxPorClave.referencia] ?? "").trim();
      if (!referencia) {
        addLog(`Fila ${i + 1}: sin referencia, se omite`, "warning");
        errores++;
        continue;
      }

      // Construir un objeto con SOLO los campos que vienen en el Excel (ignorando familia/subfamilia)
      const datosExcel = {};
      const camposAportados = [];
      const tryCol = (clave, conversor = v => v) => {
        if (idxPorClave[clave] === undefined) return;
        const raw = row[idxPorClave[clave]];
        const valTxt = String(raw ?? "").trim();
        if (valTxt === "") return; // celda vacía → no aporta
        const conv = conversor(valTxt);
        if (conv === null || conv === undefined) return;
        datosExcel[clave] = conv;
        camposAportados.push(clave);
      };
      // Conversores por tipo de campo
      const parseNum = (v) => {
        const n = parseFloat(String(v).replace(",", "."));
        return isNaN(n) ? null : n;
      };
      const parseInt0 = (v) => {
        const n = parseInt(String(v).replace(",", ".").replace(/\..*$/, ""), 10);
        return isNaN(n) ? null : n;
      };
      // Conversores específicos por clave (los que no estén aquí se tratan como texto)
      const CONVERSORES = {
        pvp: parseNum,
        preciocoste: parseNum,
        pvp_l1: parseNum,
        plazoentrega: parseInt0,
        masusado: parseInt0,
        // familia y subfamilia en la nueva tabla son integer
        familia: parseInt0,
        subfamilia: parseInt0,
      };
      // Procesar TODAS las columnas detectadas (excepto referencia, que ya se leyó)
      // y grupodescuento (necesita resolución especial de código→id)
      let errorGrupo = false;
      Object.keys(idxPorClave).forEach(clave => {
        if (clave === "referencia" || clave === "grupodescuento") return;
        const conv = CONVERSORES[clave] || (v => v);
        tryCol(clave, conv);
      });
      // grupodescuento: convertir código → id
      if (idxPorClave.grupodescuento !== undefined) {
        const codigo = String(row[idxPorClave.grupodescuento] ?? "").trim();
        if (codigo) {
          const idg = resolverIdGrupo(codigo);
          if (idg === undefined) {
            addLog(`Fila ${i + 1}: ${referencia} → ERROR grupo descuento "${codigo}" no existe en BD`, "error");
            errorGrupo = true;
          } else {
            datosExcel.grupodescuento = idg;
            camposAportados.push("grupodescuento");
          }
        }
      }
      if (errorGrupo) { errores++; continue; }

      try {
        const resBuscar = await fetch(`${API_URL}/productos/referencia/${encodeURIComponent(referencia)}`);

        if (resBuscar.ok) {
          // EXISTE
          if (!sobreescribir) {
            addLog(`Fila ${i + 1}: ${referencia} → ya existe, se omite (sobreescribir = NO)`, "info");
            continue;
          }
          // EXISTE y sobreescribir = SÍ: actualizar los campos aportados (excepto referencia)
          const prodBd = await resBuscar.json();
          if (Object.keys(datosExcel).length === 0) {
            addLog(`Fila ${i + 1}: ${referencia} → solo se aportó referencia, no hay nada que actualizar`, "info");
            continue;
          }
          const resPatch = await fetch(`${API_URL}/productos/${prodBd.id}`, {
            method: "PATCH",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(datosExcel),
          });
          if (!resPatch.ok) {
            const err = await resPatch.json().catch(() => null);
            throw new Error(err?.detail || `HTTP ${resPatch.status}`);
          }
          addLog(`Fila ${i + 1}: ${referencia} → ACTUALIZADO (${camposAportados.join(", ")})`, "success");
          actualizados++;
        } else if (resBuscar.status === 404) {
          // NO EXISTE: dar de alta. Requiere referencia + descripcion + grupodescuento
          if (!datosExcel.descripcion) {
            addLog(`Fila ${i + 1}: ${referencia} → ERROR para crear hay que aportar descripción`, "error");
            errores++;
            continue;
          }
          if (datosExcel.grupodescuento === undefined) {
            addLog(`Fila ${i + 1}: ${referencia} → ERROR para crear hay que aportar grupodescuento`, "error");
            errores++;
            continue;
          }
          // Si no hay nombre, usar primeros 50 caracteres de la descripción
          if (!datosExcel.nombre) {
            datosExcel.nombre = String(datosExcel.descripcion).slice(0, 50);
          }
          // POST con referencia + campos aportados (sin familia/subfamilia)
          const body = { referencia, ...datosExcel };
          const resCrear = await fetch(`${API_URL}/productos/`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
          });
          if (!resCrear.ok) {
            const err = await resCrear.json().catch(() => null);
            throw new Error(err?.detail || `HTTP ${resCrear.status}`);
          }
          const data = await resCrear.json();
          addLog(`Fila ${i + 1}: ${referencia} → NUEVO (ID ${data.id}, campos: ${camposAportados.join(", ")})`, "success");
          nuevos++;
        } else {
          throw new Error("HTTP " + resBuscar.status);
        }
      } catch (e) {
        addLog(`Fila ${i + 1}: ${referencia} → ERROR ${e.message}`, "error");
        errores++;
      }
    }

    setResumen({ actualizados, nuevos, errores, avisos: 0, total: fileData.rows.length });
    setProcesando(false);
    const modo = sobreescribir ? "" : " (sobreescribir = NO)";

    // Añadir resumen detallado al log para que viaje en el copiar log
    addLog("════════════════════════════════════════════════════", "info");
    addLog(`RESULTADO DEL PROCESO${modo}`, "info");
    addLog(`  • Total filas en el fichero:    ${fileData.rows.length}`, "info");
    addLog(`  • Productos actualizados:       ${actualizados}`, actualizados > 0 ? "success" : "info");
    addLog(`  • Productos nuevos creados:     ${nuevos}`, nuevos > 0 ? "success" : "info");
    addLog(`  • Errores:                      ${errores}`, errores > 0 ? "error" : "info");
    const procesadas = actualizados + nuevos + errores;
    const omitidas = fileData.rows.length - procesadas;
    if (omitidas > 0) {
      addLog(`  • Filas omitidas (sin cambios): ${omitidas}`, "info");
    }
    addLog("════════════════════════════════════════════════════", "info");

    setStatus && setStatus(`Tarifa procesada: ${actualizados} actualizados, ${nuevos} nuevos, ${errores} errores${modo}`, errores > 0 ? "error" : "success");
  };

  return (
    <div>
      <h2 style={{ fontSize: 18, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, borderBottom: "2px solid #2563eb", paddingBottom: 8 }}>Actualizar Tarifas</h2>
      <p style={{ color: "#475569", lineHeight: 1.6, marginBottom: 16, fontSize: 12 }}>
        Carga un fichero Excel de tarifa con productos. Para cada fila se comprueba si el producto existe:
        si existe se actualiza su PVP, si no existe se da de alta. La primera fila debe contener los nombres de columna.
        Columnas obligatorias: <strong>Referencia</strong>, <strong>Descripción</strong>, y o bien <strong>SubFamilia</strong> o bien <strong>GruposDescuento</strong> (si no hay subfamilia, se busca la subfamilia que contiene ese grupo de descuento).
      </p>

      {/* Drag & drop fichero */}
      <div style={{ marginBottom: 12 }}>
        <div
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={e => {
            e.preventDefault();
            setDragOver(false);
            const file = e.dataTransfer.files[0];
            if (file) handleFile(file);
          }}
          onClick={() => document.getElementById("tarifas-file-input")?.click()}
          style={{ padding: "20px", border: `2px dashed ${dragOver ? "#2563eb" : "#d4d4d4"}`, borderRadius: 8, background: dragOver ? "#eff6ff" : "#fafafa", textAlign: "center", fontSize: 12, color: "#525252", cursor: "pointer" }}>
          <Icon as={Download} size={24} color="#2563eb" />
          <div style={{ marginTop: 6 }}>
            {dragOver
              ? <strong>Suelta el fichero aquí</strong>
              : <>Arrastra un Excel/CSV (.xlsx, .xls, .csv), o haz click para elegir uno.</>}
          </div>
        </div>
        <input id="tarifas-file-input" type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
          onChange={e => { if (e.target.files?.[0]) handleFile(e.target.files[0]); e.target.value = ""; }} />
      </div>

      {/* Selector sobreescribir + info recomendación */}
      <div style={{ marginBottom: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 10 }}>
          <label style={{ fontSize: 12, color: "#171717", fontWeight: 500 }}>Si existe el producto, sobreescribir:</label>
          <div style={{ display: "inline-flex", border: "1px solid #d4d4d4", borderRadius: 6, overflow: "hidden" }}>
            <button onClick={() => setSobreescribir(true)}
              style={{ padding: "4px 14px", fontSize: 11, border: "none", cursor: "pointer",
                background: sobreescribir ? "#171717" : "#fff",
                color: sobreescribir ? "#fff" : "#171717",
                fontWeight: sobreescribir ? 600 : 400 }}>
              SÍ
            </button>
            <button onClick={() => setSobreescribir(false)}
              style={{ padding: "4px 14px", fontSize: 11, border: "none", borderLeft: "1px solid #d4d4d4", cursor: "pointer",
                background: !sobreescribir ? "#171717" : "#fff",
                color: !sobreescribir ? "#fff" : "#171717",
                fontWeight: !sobreescribir ? 600 : 400 }}>
              NO
            </button>
          </div>
        </div>

        <div style={{ padding: "10px 12px", background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 6, fontSize: 11, color: "#0369a1", lineHeight: 1.5 }}>
          <strong>Recomendación para actualizar tarifas</strong>
          <ol style={{ margin: "4px 0 0 18px", padding: 0 }}>
            <li>Carga una Excel con <code>referencia, pvp, descripcion, grupodescuento, nombre</code> y el selector <strong>"Si existe el producto, sobreescribir: NO"</strong>. Esto creará los productos nuevos con toda la información y dejará los existentes intactos.</li>
            <li>Vuelve a pasar la misma Excel (o solo con <code>referencia, pvp, grupodescuento</code>) y el selector <strong>"Si existe el producto, sobreescribir: SÍ"</strong>. Esto actualizará los PVP y los grupos descuento de los productos existentes sin sobreescribir nombre ni descripción.</li>
          </ol>
        </div>
      </div>
      {/* Mapeo */}
      {fileData && (
        <div style={{ marginBottom: 12, padding: "10px 12px", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 6 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", marginBottom: 6 }}>
            Mapeo de columnas detectado <span style={{ color: "#64748b", fontWeight: 400 }}>({fileData.rows.length} filas en "{fileData.origen}")</span>:
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {fileData.headers.map((h, i) => (
              <div key={i} style={{ fontSize: 11, padding: "3px 8px", borderRadius: 4, background: mapping[i] ? "#dcfce7" : "#fee2e2", border: mapping[i] ? "1px solid #86efac" : "1px solid #fca5a5", color: mapping[i] ? "#14532d" : "#991b1b" }}>
                <strong>{h}</strong> {mapping[i] ? "→ " + mapping[i] : "(no reconocida)"}
              </div>
            ))}
          </div>
          {erroresMapeo.length > 0 && (
            <div style={{ marginTop: 8, padding: "6px 10px", background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 4, fontSize: 11, color: "#991b1b" }}>
              <strong>Errores:</strong>
              <ul style={{ margin: "4px 0 0 16px", padding: 0 }}>
                {erroresMapeo.map((e, i) => <li key={i}>{e}</li>)}
              </ul>
            </div>
          )}
        </div>
      )}

      {/* Botones de acción */}
      {fileData && (
        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          <button onClick={aplicarTarifa} disabled={procesando || erroresMapeo.length > 0}
            style={{ padding: "7px 18px", borderRadius: 6, border: (!procesando && erroresMapeo.length === 0) ? "1px solid #16a34a" : "1px solid #d4d4d4", background: (!procesando && erroresMapeo.length === 0) ? "#dcfce7" : "#f5f5f5", color: (!procesando && erroresMapeo.length === 0) ? "#14532d" : "#737373", cursor: (!procesando && erroresMapeo.length === 0) ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={procesando ? RefreshCw : Check} iconColor={(!procesando && erroresMapeo.length === 0) ? "#14532d" : "#737373"}>{procesando ? "Procesando..." : "Aplicar tarifa a BD"}</BtnContent>
          </button>
          {procesando && (
            <button onClick={() => { stopRequested.current = true; setStatus && setStatus("Solicitando parada del procesado...", "working"); }}
              style={{ padding: "7px 18px", borderRadius: 6, border: "1px solid #dc2626", background: "#fef2f2", color: "#991b1b", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
              <BtnContent icon={X} iconColor="#dc2626">Parar</BtnContent>
            </button>
          )}
        </div>
      )}

      {/* Resumen */}
      {resumen && (
        <div style={{ background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 6, padding: "10px 14px", marginBottom: 10, fontSize: 12, color: "#0369a1" }}>
          <strong>Resumen:</strong> {resumen.total} filas procesadas — <span style={{ color: "#16a34a" }}>{resumen.actualizados} actualizados</span>, <span style={{ color: "#2563eb" }}>{resumen.nuevos} nuevos</span>, <span style={{ color: "#dc2626" }}>{resumen.errores} errores</span>, <span style={{ color: "#d97706" }}>{resumen.avisos} avisos</span>
        </div>
      )}

      {/* Log: siempre visible para ver el progreso */}
      <div style={{ marginTop: 8 }}>
        <div style={{ fontSize: 12, fontWeight: 600, color: "#171717", marginBottom: 4, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span>Progreso ({log.length} entrada{log.length !== 1 ? "s" : ""})</span>
          {log.length > 0 && (
            <div style={{ display: "inline-flex", gap: 6 }}>
              <button onClick={() => {
                  const texto = log.map(l => {
                    const prefijo = l.tipo === "error" ? "[ERROR]" : l.tipo === "warning" ? "[AVISO]" : l.tipo === "success" ? "[OK]" : "[INFO]";
                    return `[${l.hora}] ${prefijo} ${l.texto}`;
                  }).join("\n");
                  if (navigator.clipboard) {
                    navigator.clipboard.writeText(texto)
                      .then(() => setStatus && setStatus(`Log (${log.length} líneas) copiado al portapapeles`, "success"))
                      .catch(() => setStatus && setStatus("No se pudo copiar al portapapeles", "error"));
                  } else {
                    // Fallback: textarea + execCommand
                    const ta = document.createElement("textarea");
                    ta.value = texto;
                    document.body.appendChild(ta);
                    ta.select();
                    try { document.execCommand("copy"); setStatus && setStatus(`Log (${log.length} líneas) copiado al portapapeles`, "success"); }
                    catch { setStatus && setStatus("No se pudo copiar al portapapeles", "error"); }
                    document.body.removeChild(ta);
                  }
                }}
                title="Copiar todo el log al portapapeles"
                style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                <BtnContent icon={Copy} iconColor="#475569">Copiar log</BtnContent>
              </button>
              <button onClick={() => exportarLogCSV(log, "log")}
                title="Exportar el log a un fichero CSV"
                style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                <BtnContent icon={Download} iconColor="#475569">Exportar CSV</BtnContent>
              </button>
              <button onClick={() => setLog([])}
                style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                <BtnContent icon={Trash2} iconColor="#475569">Limpiar log</BtnContent>
              </button>
            </div>
          )}
        </div>
        <div ref={logRef} style={{ height: 320, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, background: "#0f172a", color: "#e2e8f0", fontFamily: "monospace", fontSize: 11, padding: "8px 12px", lineHeight: 1.5 }}>
          {log.length === 0 ? (
            <div style={{ color: "#64748b", textAlign: "center", padding: "20px 0", fontStyle: "italic" }}>
              Sin actividad. Carga un fichero Excel y pulsa "Aplicar tarifa a BD" para ver el progreso aquí.
            </div>
          ) : log.map((l, i) => (
            <div key={i} style={{ padding: "1px 0",
              color: l.tipo === "error" ? "#fca5a5"
                : l.tipo === "warning" ? "#fcd34d"
                : l.tipo === "success" ? "#86efac"
                : "#cbd5e1",
              fontWeight: l.tipo === "error" ? 600 : 400 }}>
              <span style={{ color: "#64748b" }}>[{l.hora}]</span>
              {l.tipo === "error" && <span style={{ color: "#fca5a5", fontWeight: 700 }}> ❌ </span>}
              {l.tipo === "warning" && <span style={{ color: "#fcd34d", fontWeight: 700 }}> ⚠ </span>}
              {l.tipo === "success" && <span style={{ color: "#86efac" }}> ✓ </span>}
              {l.tipo === "info" && " "}
              {l.texto}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── Sección Mantenimiento BD ──
// ── Sección Gestión de Usuarios ──
function UsuariosSection({ setStatus }) {
  const [usuarios, setUsuarios] = useState([]);
  const [cargando, setCargando] = useState(false);
  const [seleccionado, setSeleccionado] = useState(null); // id del usuario en edición
  const [draft, setDraft] = useState(null); // datos del formulario en edición
  const [busqueda, setBusqueda] = useState("");
  const [confirmBorrar, setConfirmBorrar] = useState(null); // {id, usuario}
  const [error, setError] = useState(null);

  const USUARIO_VACIO = {
    usuario: "",
    codigopresupuestos: "",
    password: "",
    permisocrearpresupuesto: false,
    permisoeditarpresupuesto: false,
    permisoborrarpresupuesto: false,
    permisocrearcliente: false,
    permisoeditarcliente: false,
    permisoborrarcliente: false,
    permisocrearproducto: false,
    permisoeditarproducto: false,
    permisoborrarproducto: false,
  };

  const cargarUsuarios = async () => {
    setCargando(true);
    setError(null);
    try {
      const res = await fetch(`${API_URL}/usuarios/`);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      setUsuarios(data);
    } catch (e) {
      setError(e.message);
    } finally {
      setCargando(false);
    }
  };

  useEffect(() => { cargarUsuarios(); }, []);

  const usuariosFiltrados = usuarios.filter(u =>
    !busqueda.trim() ||
    String(u.usuario || "").toLowerCase().includes(busqueda.toLowerCase()) ||
    String(u.codigopresupuestos || "").toLowerCase().includes(busqueda.toLowerCase())
  );

  const editar = (u) => {
    setSeleccionado(u.id);
    setDraft({ ...USUARIO_VACIO, ...u, password: "" });
  };

  const nuevo = () => {
    setSeleccionado("nuevo");
    setDraft({ ...USUARIO_VACIO });
  };

  const cancelar = () => {
    setSeleccionado(null);
    setDraft(null);
  };

  const guardar = async () => {
    if (!draft) return;
    if (!draft.usuario || !draft.usuario.trim()) {
      setStatus && setStatus("El campo Usuario es obligatorio", "error");
      return;
    }
    setStatus && setStatus("Guardando usuario...", "working");
    try {
      const payload = { ...draft };
      // No enviar password vacío al actualizar (mantiene el actual)
      if (seleccionado !== "nuevo" && !payload.password) {
        delete payload.password;
      }
      const url = seleccionado === "nuevo" ? `${API_URL}/usuarios/` : `${API_URL}/usuarios/${seleccionado}`;
      const method = seleccionado === "nuevo" ? "POST" : "PUT";
      const res = await fetch(url, {
        method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "HTTP " + res.status);
      }
      setStatus && setStatus(`Usuario "${draft.usuario}" ${seleccionado === "nuevo" ? "creado" : "actualizado"}`, "success");
      setSeleccionado(null);
      setDraft(null);
      cargarUsuarios();
    } catch (e) {
      setStatus && setStatus("Error guardando: " + e.message, "error");
    }
  };

  const borrar = async (id, nombre) => {
    setStatus && setStatus(`Borrando usuario "${nombre}"...`, "working");
    try {
      const res = await fetch(`${API_URL}/usuarios/${id}`, { method: "DELETE" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      setStatus && setStatus(`Usuario "${nombre}" borrado`, "success");
      setConfirmBorrar(null);
      if (seleccionado === id) { setSeleccionado(null); setDraft(null); }
      cargarUsuarios();
    } catch (e) {
      setStatus && setStatus("Error borrando: " + e.message, "error");
    }
  };

  // Toggle de permiso visual
  const Toggle = ({ label, checked, onChange }) => (
    <label style={{ display: "flex", alignItems: "center", gap: 8, padding: "4px 0", cursor: "pointer", fontSize: 12, color: "#171717" }}>
      <input type="checkbox" checked={!!checked} onChange={e => onChange(e.target.checked)} />
      {label}
    </label>
  );

  return (
    <div>
      <h2 style={{ fontSize: 18, fontWeight: 700, color: "#171717", margin: 0, marginBottom: 8, borderBottom: "2px solid #2563eb", paddingBottom: 8 }}>
        Gestión de Usuarios
      </h2>
      <p style={{ color: "#525252", lineHeight: 1.6, marginBottom: 16, fontSize: 12 }}>
        Crear, editar y borrar usuarios de la aplicación. Cada usuario tiene su propio código de presupuestos y permisos para crear, editar o borrar presupuestos, clientes y productos.
      </p>

      <div style={{ display: "grid", gridTemplateColumns: draft ? "1fr 1fr" : "1fr", gap: 16 }}>
        {/* Lista de usuarios */}
        <div style={{ border: "1px solid #e5e5e5", borderRadius: 8, padding: 12, background: "#fff" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <input value={busqueda} onChange={e => setBusqueda(e.target.value)}
              placeholder="Buscar usuario..."
              style={{ flex: 1, padding: "5px 10px", border: "1px solid #d4d4d4", borderRadius: 6, fontSize: 12 }} />
            <button onClick={nuevo}
              style={{ marginLeft: 8, padding: "5px 12px", borderRadius: 6, border: "1px solid #16a34a", background: "#dcfce7", color: "#14532d", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
              <BtnContent icon={Plus} iconColor="#14532d">Nuevo</BtnContent>
            </button>
            <button onClick={cargarUsuarios} title="Recargar lista"
              style={{ marginLeft: 6, padding: "5px 10px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
              <BtnContent icon={RefreshCw} iconColor="#475569" />
            </button>
          </div>

          {error && <div style={{ padding: 10, color: "#991b1b", background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 6, fontSize: 12, marginBottom: 8 }}>Error: {error}</div>}

          {cargando ? (
            <div style={{ padding: 16, textAlign: "center", color: "#737373", fontSize: 12 }}>Cargando...</div>
          ) : usuariosFiltrados.length === 0 ? (
            <div style={{ padding: 16, textAlign: "center", color: "#737373", fontSize: 12 }}>
              {usuarios.length === 0 ? "No hay usuarios registrados" : "Ningún usuario coincide con la búsqueda"}
            </div>
          ) : (
            <div style={{ maxHeight: 480, overflowY: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#fafafa", color: "#171717", fontSize: 11 }}>
                    <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Usuario</th>
                    <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Código</th>
                    <th style={{ padding: "6px 8px", textAlign: "center", borderBottom: "1px solid #e5e5e5", fontWeight: 600, width: 60 }}>Pwd</th>
                    <th style={{ padding: "6px 8px", textAlign: "right", borderBottom: "1px solid #e5e5e5", fontWeight: 600, width: 110 }}></th>
                  </tr>
                </thead>
                <tbody>
                  {usuariosFiltrados.map(u => (
                    <tr key={u.id} style={{ borderBottom: "1px solid #f1f5f9",
                        background: seleccionado === u.id ? "#eff6ff" : "transparent" }}>
                      <td style={{ padding: "6px 8px", fontWeight: 500, color: "#171717" }}>{u.usuario}</td>
                      <td style={{ padding: "6px 8px", color: "#525252", fontFamily: "monospace" }}>{u.codigopresupuestos || "—"}</td>
                      <td style={{ padding: "6px 8px", textAlign: "center" }}>
                        {u.tiene_password ? <span style={{ color: "#16a34a" }}>✓</span> : <span style={{ color: "#d4d4d4" }}>—</span>}
                      </td>
                      <td style={{ padding: "4px 6px", textAlign: "right", whiteSpace: "nowrap" }}>
                        <button onClick={() => editar(u)} title="Editar"
                          style={{ padding: "3px 8px", borderRadius: 4, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 11, marginRight: 4 }}>
                          <BtnContent icon={Edit3} iconColor="#475569" />
                        </button>
                        <button onClick={() => setConfirmBorrar({ id: u.id, usuario: u.usuario })} title="Borrar"
                          style={{ padding: "3px 8px", borderRadius: 4, border: "1px solid #fca5a5", background: "#fef2f2", color: "#991b1b", cursor: "pointer", fontSize: 11 }}>
                          <BtnContent icon={Trash2} iconColor="#dc2626" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* Formulario de edición */}
        {draft && (
          <div style={{ border: "1px solid #e5e5e5", borderRadius: 8, padding: 14, background: "#fff" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, paddingBottom: 8, borderBottom: "1px solid #e5e5e5" }}>
              <h3 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: "#171717" }}>
                {seleccionado === "nuevo" ? "Nuevo usuario" : `Editar: ${draft.usuario}`}
              </h3>
            </div>

            {/* Campos básicos */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
              <div>
                <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 3 }}>Usuario *</label>
                <input value={draft.usuario || ""} onChange={e => setDraft({ ...draft, usuario: e.target.value })}
                  style={{ width: "100%", padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
              </div>
              <div>
                <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 3 }}>Código presupuestos</label>
                <input value={draft.codigopresupuestos || ""} onChange={e => setDraft({ ...draft, codigopresupuestos: e.target.value })}
                  maxLength={10} placeholder="ej. CDM"
                  style={{ width: "100%", padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12, fontFamily: "monospace", boxSizing: "border-box" }} />
              </div>
              <div style={{ gridColumn: "1 / -1" }}>
                <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 3 }}>
                  Contraseña {seleccionado !== "nuevo" && <span style={{ color: "#737373", fontStyle: "italic" }}>(deja vacío para no cambiarla)</span>}
                </label>
                <input type="password" value={draft.password || ""} onChange={e => setDraft({ ...draft, password: e.target.value })}
                  placeholder="••••••••"
                  style={{ width: "100%", padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
              </div>
            </div>

            {/* Permisos */}
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: "#171717", marginBottom: 6, paddingBottom: 4, borderBottom: "1px solid #f1f5f9" }}>Permisos</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6 }}>
                <div>
                  <div style={{ fontSize: 11, fontWeight: 600, color: "#525252", marginBottom: 2 }}>Presupuestos</div>
                  <Toggle label="Crear" checked={draft.permisocrearpresupuesto} onChange={v => setDraft({ ...draft, permisocrearpresupuesto: v })} />
                  <Toggle label="Editar" checked={draft.permisoeditarpresupuesto} onChange={v => setDraft({ ...draft, permisoeditarpresupuesto: v })} />
                  <Toggle label="Borrar" checked={draft.permisoborrarpresupuesto} onChange={v => setDraft({ ...draft, permisoborrarpresupuesto: v })} />
                </div>
                <div>
                  <div style={{ fontSize: 11, fontWeight: 600, color: "#525252", marginBottom: 2 }}>Clientes</div>
                  <Toggle label="Crear" checked={draft.permisocrearcliente} onChange={v => setDraft({ ...draft, permisocrearcliente: v })} />
                  <Toggle label="Editar" checked={draft.permisoeditarcliente} onChange={v => setDraft({ ...draft, permisoeditarcliente: v })} />
                  <Toggle label="Borrar" checked={draft.permisoborrarcliente} onChange={v => setDraft({ ...draft, permisoborrarcliente: v })} />
                </div>
                <div>
                  <div style={{ fontSize: 11, fontWeight: 600, color: "#525252", marginBottom: 2 }}>Productos</div>
                  <Toggle label="Crear" checked={draft.permisocrearproducto} onChange={v => setDraft({ ...draft, permisocrearproducto: v })} />
                  <Toggle label="Editar" checked={draft.permisoeditarproducto} onChange={v => setDraft({ ...draft, permisoeditarproducto: v })} />
                  <Toggle label="Borrar" checked={draft.permisoborrarproducto} onChange={v => setDraft({ ...draft, permisoborrarproducto: v })} />
                </div>
              </div>
            </div>

            {/* Botones */}
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", paddingTop: 10, borderTop: "1px solid #e5e5e5" }}>
              <button onClick={cancelar}
                style={{ padding: "6px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
                <BtnContent icon={X}>Cancelar</BtnContent>
              </button>
              <button onClick={guardar}
                style={{ padding: "6px 18px", borderRadius: 6, border: "1px solid #16a34a", background: "#dcfce7", color: "#14532d", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                <BtnContent icon={Save} iconColor="#14532d">{seleccionado === "nuevo" ? "Crear" : "Guardar"}</BtnContent>
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Diálogo confirmar borrado */}
      {confirmBorrar && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}
          onClick={() => setConfirmBorrar(null)}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 420 }} onClick={e => e.stopPropagation()}>
            <h3 style={{ margin: 0, marginBottom: 12, fontSize: 15, fontWeight: 700, color: "#171717" }}>¿Borrar usuario?</h3>
            <p style={{ fontSize: 13, color: "#525252", marginBottom: 16, lineHeight: 1.5 }}>
              ¿Seguro que quieres borrar el usuario <strong style={{ color: "#171717" }}>{confirmBorrar.usuario}</strong>? Esta acción no se puede deshacer.
            </p>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setConfirmBorrar(null)}
                style={{ padding: "6px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
                <BtnContent icon={X}>Cancelar</BtnContent>
              </button>
              <button onClick={() => borrar(confirmBorrar.id, confirmBorrar.usuario)}
                style={{ padding: "6px 14px", borderRadius: 6, border: "1px solid #dc2626", background: "#fef2f2", color: "#991b1b", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                <BtnContent icon={Trash2} iconColor="#dc2626">Sí, borrar</BtnContent>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Diálogo Mantenimiento BD Subfamilia ──
function MantenimientoSubfamiliaDialog({ onClose, setStatus }) {
  const [subfamilias, setSubfamilias] = useState([]);
  const [familias, setFamilias] = useState([]); // para el selector al editar
  const [cargando, setCargando] = useState(false);
  const [busqueda, setBusqueda] = useState("");
  const [seleccionada, setSeleccionada] = useState(null);
  const [editar, setEditar] = useState(null);
  const [guardando, setGuardando] = useState(false);
  const [confirmBorrar, setConfirmBorrar] = useState(null);
  const [borrando, setBorrando] = useState(false);
  const [comprobandoUso, setComprobandoUso] = useState(false);

  const cargar = async () => {
    setCargando(true);
    try {
      const res = await fetch(`${API_URL}/subfamilias/`);
      if (!res.ok) throw new Error("HTTP " + res.status);
      setSubfamilias(await res.json());
    } catch (e) {
      setStatus && setStatus("Error cargando subfamilias: " + e.message, "error");
    } finally {
      setCargando(false);
    }
  };

  useEffect(() => {
    cargar();
    // Cargar familias en paralelo para el selector de edición
    fetch(`${API_URL}/familias/`)
      .then(r => r.ok ? r.json() : [])
      .then(setFamilias)
      .catch(() => setFamilias([]));
  }, []);

  const filtradas = subfamilias.filter(sf =>
    !busqueda.trim() ||
    String(sf.subfamilia || "").toLowerCase().includes(busqueda.toLowerCase()) ||
    String(sf.descripcion || "").toLowerCase().includes(busqueda.toLowerCase()) ||
    String(sf.familia || "").toLowerCase().includes(busqueda.toLowerCase())
  );

  const iniciarEdicion = () => {
    if (!seleccionada) return;
    setEditar({
      subfamilia: seleccionada.subfamilia || "",
      descripcion: seleccionada.descripcion || "",
      idfamiliarelacionada: seleccionada.idfamilia || null,
    });
  };

  const cancelarEdicion = () => setEditar(null);

  const guardarEdicion = async () => {
    if (!seleccionada || !editar) return;
    if (!editar.subfamilia || !editar.subfamilia.trim()) {
      setStatus && setStatus("El nombre de la subfamilia no puede estar vacío", "error");
      return;
    }
    setGuardando(true);
    try {
      const res = await fetch(`${API_URL}/subfamilias/${seleccionada.id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          subfamilia: editar.subfamilia.trim(),
          descripcion: editar.descripcion,
          idfamiliarelacionada: editar.idfamiliarelacionada,
        }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "HTTP " + res.status);
      }
      setStatus && setStatus(`Subfamilia "${editar.subfamilia}" actualizada`, "success");
      setEditar(null);
      cargar();
    } catch (e) {
      setStatus && setStatus("Error guardando: " + e.message, "error");
    } finally {
      setGuardando(false);
    }
  };

  const comprobarYBorrar = async () => {
    if (!seleccionada) return;
    setComprobandoUso(true);
    setStatus && setStatus(`Comprobando referencias de "${seleccionada.subfamilia}"...`, "working");
    try {
      const res = await fetch(`${API_URL}/subfamilias/${seleccionada.id}/uso`);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      if (!data.borrable) {
        setStatus && setStatus(`No se puede borrar "${seleccionada.subfamilia}": está referenciada en ${data.en_gruposdescuento} grupo(s) de descuento.`, "error");
        return;
      }
      setStatus && setStatus("Subfamilia sin referencias, listo para borrar", "info");
      setConfirmBorrar({ id: seleccionada.id, subfamilia: seleccionada.subfamilia });
    } catch (e) {
      setStatus && setStatus("Error comprobando uso: " + e.message, "error");
    } finally {
      setComprobandoUso(false);
    }
  };

  const borrarConfirmado = async () => {
    if (!confirmBorrar) return;
    setBorrando(true);
    try {
      const res = await fetch(`${API_URL}/subfamilias/${confirmBorrar.id}`, { method: "DELETE" });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "HTTP " + res.status);
      }
      setStatus && setStatus(`Subfamilia "${confirmBorrar.subfamilia}" borrada`, "success");
      setConfirmBorrar(null);
      setSeleccionada(null);
      cargar();
    } catch (e) {
      setStatus && setStatus("Error borrando: " + e.message, "error");
    } finally {
      setBorrando(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }} onClick={() => !editar && !borrando && !guardando && onClose()}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 880, maxHeight: "85vh", display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, paddingBottom: 10, borderBottom: "1px solid #e5e5e5" }}>
          <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Wrench} size={18} color="#1e3a5f" /> Mantenimiento BD Subfamilia
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10 }}>
          <input value={busqueda} onChange={e => setBusqueda(e.target.value)}
            placeholder="Buscar subfamilia, descripción o familia..."
            style={{ flex: 1, padding: "6px 10px", border: "1px solid #cbd5e1", borderRadius: 6, fontSize: 12 }} />
          <span style={{ fontSize: 11, color: "#737373" }}>{filtradas.length} subfamilia{filtradas.length !== 1 ? "s" : ""}</span>
          <button onClick={cargar} title="Recargar"
            style={{ padding: "6px 10px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", cursor: "pointer", fontSize: 12 }}>
            <BtnContent icon={RefreshCw} iconColor="#475569" />
          </button>
        </div>

        <div style={{ flex: 1, overflow: "auto", border: "1px solid #e5e5e5", borderRadius: 6, marginBottom: 10, minHeight: 200 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 1, borderBottom: "1px solid #e5e5e5" }}>
              <tr>
                <th style={{ padding: "6px 10px", textAlign: "right", fontWeight: 600, color: "#171717", width: 50 }}>ID</th>
                <th style={{ padding: "6px 10px", textAlign: "left", fontWeight: 600, color: "#171717", width: "22%" }}>Subfamilia</th>
                <th style={{ padding: "6px 10px", textAlign: "left", fontWeight: 600, color: "#171717" }}>Descripción</th>
                <th style={{ padding: "6px 10px", textAlign: "left", fontWeight: 600, color: "#171717", width: "22%" }}>Familia</th>
              </tr>
            </thead>
            <tbody>
              {cargando && <tr><td colSpan={4} style={{ padding: 20, textAlign: "center", color: "#737373" }}>Cargando...</td></tr>}
              {!cargando && filtradas.length === 0 && <tr><td colSpan={4} style={{ padding: 20, textAlign: "center", color: "#94a3b8" }}>Sin resultados</td></tr>}
              {!cargando && filtradas.map(sf => {
                const sel = seleccionada?.id === sf.id;
                const editando = sel && editar;
                return (
                  <tr key={sf.id}
                    onClick={() => !editar && setSeleccionada(sf)}
                    style={{ background: sel ? "#dbeafe" : "transparent", cursor: editar ? "default" : "pointer", borderBottom: "1px solid #f1f5f9" }}>
                    <td style={{ padding: "5px 10px", textAlign: "right", color: "#1e3a5f", fontWeight: 600 }}>{sf.id}</td>
                    <td style={{ padding: "5px 10px", color: "#171717" }}>
                      {editando ? (
                        <input autoFocus value={editar.subfamilia} onChange={e => setEditar({ ...editar, subfamilia: e.target.value })}
                          style={{ width: "100%", padding: "3px 6px", border: "1px solid #2563eb", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
                      ) : sf.subfamilia}
                    </td>
                    <td style={{ padding: "5px 10px", color: "#525252" }}>
                      {editando ? (
                        <input value={editar.descripcion || ""} onChange={e => setEditar({ ...editar, descripcion: e.target.value })}
                          style={{ width: "100%", padding: "3px 6px", border: "1px solid #2563eb", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
                      ) : (sf.descripcion || "—")}
                    </td>
                    <td style={{ padding: "5px 10px", color: "#525252" }}>
                      {editando ? (
                        <select value={editar.idfamiliarelacionada || ""}
                          onChange={e => setEditar({ ...editar, idfamiliarelacionada: e.target.value ? parseInt(e.target.value) : null })}
                          style={{ width: "100%", padding: "3px 6px", border: "1px solid #2563eb", borderRadius: 4, fontSize: 12, background: "#fff", boxSizing: "border-box" }}>
                          <option value="">— Sin familia —</option>
                          {familias.map(f => <option key={f.id} value={f.id}>{f.familia}</option>)}
                        </select>
                      ) : (sf.familia || "—")}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        <div style={{ display: "flex", gap: 6, justifyContent: "space-between", alignItems: "center", paddingTop: 10, borderTop: "1px solid #e5e5e5" }}>
          <div style={{ display: "flex", gap: 6 }}>
            <button onClick={comprobarYBorrar} disabled={!seleccionada || comprobandoUso || !!editar}
              style={{ padding: "7px 14px", borderRadius: 6, border: (seleccionada && !editar) ? "1px solid #fca5a5" : "1px solid #d4d4d4", background: (seleccionada && !editar) ? "#fef2f2" : "#f5f5f5", color: (seleccionada && !editar) ? "#991b1b" : "#737373", cursor: (seleccionada && !editar && !comprobandoUso) ? "pointer" : "default", fontSize: 12 }}>
              <BtnContent icon={comprobandoUso ? RefreshCw : Trash2} iconColor={(seleccionada && !editar) ? "#dc2626" : "#737373"}>
                {comprobandoUso ? "Comprobando..." : "Borrar"}
              </BtnContent>
            </button>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            {!editar ? (
              <button onClick={iniciarEdicion} disabled={!seleccionada}
                style={{ padding: "7px 14px", borderRadius: 6, border: seleccionada ? "1px solid #cbd5e1" : "1px solid #d4d4d4", background: seleccionada ? "#fff" : "#f5f5f5", color: seleccionada ? "#171717" : "#737373", cursor: seleccionada ? "pointer" : "default", fontSize: 12 }}>
                <BtnContent icon={Edit3} iconColor={seleccionada ? "#475569" : "#737373"}>Editar</BtnContent>
              </button>
            ) : (
              <>
                <button onClick={cancelarEdicion} disabled={guardando}
                  style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: guardando ? "default" : "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={guardarEdicion} disabled={guardando}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #16a34a", background: "#dcfce7", color: "#14532d", cursor: guardando ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={guardando ? RefreshCw : Check} iconColor="#14532d">{guardando ? "Guardando..." : "Aceptar"}</BtnContent>
                </button>
              </>
            )}
            {!editar && (
              <button onClick={onClose}
                style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
                <BtnContent icon={X}>Cerrar</BtnContent>
              </button>
            )}
          </div>
        </div>

        {/* Confirmar borrar */}
        {confirmBorrar && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }} onClick={() => !borrando && setConfirmBorrar(null)}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 450 }} onClick={e => e.stopPropagation()}>
              <h3 style={{ margin: 0, marginBottom: 12, fontSize: 15, fontWeight: 700, display: "inline-flex", alignItems: "center", gap: 8 }}>
                <Icon as={Trash2} size={18} color="#dc2626" /> ¿Borrar subfamilia?
              </h3>
              <p style={{ fontSize: 13, color: "#525252", marginBottom: 12 }}>
                Vas a borrar la subfamilia <strong>"{confirmBorrar.subfamilia}"</strong>.
              </p>
              <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "10px 12px", fontSize: 12, color: "#991b1b", marginBottom: 16 }}>
                <strong>Atención:</strong> esta acción no se puede deshacer.
              </div>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmBorrar(null)} disabled={borrando}
                  style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", cursor: borrando ? "default" : "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={borrarConfirmado} disabled={borrando}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #dc2626", background: "#fef2f2", color: "#991b1b", cursor: borrando ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={borrando ? RefreshCw : Trash2} iconColor="#dc2626">{borrando ? "Borrando..." : "Sí, borrar"}</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Diálogo Mantenimiento BD Familia ──
function MantenimientoFamiliaDialog({ onClose, setStatus }) {
  const [familias, setFamilias] = useState([]);
  const [cargando, setCargando] = useState(false);
  const [busqueda, setBusqueda] = useState("");
  const [seleccionada, setSeleccionada] = useState(null);
  const [editar, setEditar] = useState(null); // {familia, descripcion} en edición
  const [guardando, setGuardando] = useState(false);
  const [confirmBorrar, setConfirmBorrar] = useState(null); // {id, familia, en_subfamilias}
  const [borrando, setBorrando] = useState(false);
  const [comprobandoUso, setComprobandoUso] = useState(false);

  const cargar = async () => {
    setCargando(true);
    try {
      const res = await fetch(`${API_URL}/familias/`);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      setFamilias(data);
    } catch (e) {
      setStatus && setStatus("Error cargando familias: " + e.message, "error");
    } finally {
      setCargando(false);
    }
  };

  useEffect(() => { cargar(); }, []);

  const filtradas = familias.filter(f =>
    !busqueda.trim() ||
    String(f.familia || "").toLowerCase().includes(busqueda.toLowerCase()) ||
    String(f.descripcion || "").toLowerCase().includes(busqueda.toLowerCase())
  );

  const iniciarEdicion = () => {
    if (!seleccionada) return;
    setEditar({ familia: seleccionada.familia || "", descripcion: seleccionada.descripcion || "" });
  };

  const cancelarEdicion = () => setEditar(null);

  const guardarEdicion = async () => {
    if (!seleccionada || !editar) return;
    if (!editar.familia || !editar.familia.trim()) {
      setStatus && setStatus("El nombre de la familia no puede estar vacío", "error");
      return;
    }
    setGuardando(true);
    try {
      const res = await fetch(`${API_URL}/familias/${seleccionada.id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ familia: editar.familia.trim(), descripcion: editar.descripcion }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "HTTP " + res.status);
      }
      setStatus && setStatus(`Familia "${editar.familia}" actualizada`, "success");
      setEditar(null);
      setSeleccionada({ ...seleccionada, familia: editar.familia.trim(), descripcion: editar.descripcion });
      cargar();
    } catch (e) {
      setStatus && setStatus("Error guardando: " + e.message, "error");
    } finally {
      setGuardando(false);
    }
  };

  const comprobarYBorrar = async () => {
    if (!seleccionada) return;
    setComprobandoUso(true);
    setStatus && setStatus(`Comprobando referencias de "${seleccionada.familia}"...`, "working");
    try {
      const res = await fetch(`${API_URL}/familias/${seleccionada.id}/uso`);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      if (!data.borrable) {
        setStatus && setStatus(`No se puede borrar "${seleccionada.familia}": está referenciada en ${data.en_subfamilias} subfamilia(s).`, "error");
        return;
      }
      setStatus && setStatus("Familia sin referencias, listo para borrar", "info");
      setConfirmBorrar({ id: seleccionada.id, familia: seleccionada.familia });
    } catch (e) {
      setStatus && setStatus("Error comprobando uso: " + e.message, "error");
    } finally {
      setComprobandoUso(false);
    }
  };

  const borrarConfirmado = async () => {
    if (!confirmBorrar) return;
    setBorrando(true);
    try {
      const res = await fetch(`${API_URL}/familias/${confirmBorrar.id}`, { method: "DELETE" });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "HTTP " + res.status);
      }
      setStatus && setStatus(`Familia "${confirmBorrar.familia}" borrada`, "success");
      setConfirmBorrar(null);
      setSeleccionada(null);
      cargar();
    } catch (e) {
      setStatus && setStatus("Error borrando: " + e.message, "error");
    } finally {
      setBorrando(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }} onClick={() => !editar && !borrando && !guardando && onClose()}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 760, maxHeight: "85vh", display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, paddingBottom: 10, borderBottom: "1px solid #e5e5e5" }}>
          <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Wrench} size={18} color="#1e3a5f" /> Mantenimiento BD Familia
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10 }}>
          <input value={busqueda} onChange={e => setBusqueda(e.target.value)}
            placeholder="Buscar familia..."
            style={{ flex: 1, padding: "6px 10px", border: "1px solid #cbd5e1", borderRadius: 6, fontSize: 12 }} />
          <span style={{ fontSize: 11, color: "#737373" }}>{filtradas.length} familia{filtradas.length !== 1 ? "s" : ""}</span>
          <button onClick={cargar} title="Recargar"
            style={{ padding: "6px 10px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", cursor: "pointer", fontSize: 12 }}>
            <BtnContent icon={RefreshCw} iconColor="#475569" />
          </button>
        </div>

        <div style={{ flex: 1, overflow: "auto", border: "1px solid #e5e5e5", borderRadius: 6, marginBottom: 10, minHeight: 200 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 1, borderBottom: "1px solid #e5e5e5" }}>
              <tr>
                <th style={{ padding: "6px 10px", textAlign: "right", fontWeight: 600, color: "#171717", width: 50 }}>ID</th>
                <th style={{ padding: "6px 10px", textAlign: "left", fontWeight: 600, color: "#171717", width: "30%" }}>Familia</th>
                <th style={{ padding: "6px 10px", textAlign: "left", fontWeight: 600, color: "#171717" }}>Descripción</th>
              </tr>
            </thead>
            <tbody>
              {cargando && <tr><td colSpan={3} style={{ padding: 20, textAlign: "center", color: "#737373" }}>Cargando...</td></tr>}
              {!cargando && filtradas.length === 0 && <tr><td colSpan={3} style={{ padding: 20, textAlign: "center", color: "#94a3b8" }}>Sin resultados</td></tr>}
              {!cargando && filtradas.map(f => {
                const sel = seleccionada?.id === f.id;
                const editando = sel && editar;
                return (
                  <tr key={f.id}
                    onClick={() => !editar && setSeleccionada(f)}
                    style={{ background: sel ? "#dbeafe" : "transparent", cursor: editar ? "default" : "pointer", borderBottom: "1px solid #f1f5f9" }}>
                    <td style={{ padding: "5px 10px", textAlign: "right", color: "#1e3a5f", fontWeight: 600 }}>{f.id}</td>
                    <td style={{ padding: "5px 10px", color: "#171717" }}>
                      {editando ? (
                        <input autoFocus value={editar.familia} onChange={e => setEditar({ ...editar, familia: e.target.value })}
                          style={{ width: "100%", padding: "3px 6px", border: "1px solid #2563eb", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
                      ) : f.familia}
                    </td>
                    <td style={{ padding: "5px 10px", color: "#525252" }}>
                      {editando ? (
                        <input value={editar.descripcion || ""} onChange={e => setEditar({ ...editar, descripcion: e.target.value })}
                          style={{ width: "100%", padding: "3px 6px", border: "1px solid #2563eb", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
                      ) : (f.descripcion || "—")}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        <div style={{ display: "flex", gap: 6, justifyContent: "space-between", alignItems: "center", paddingTop: 10, borderTop: "1px solid #e5e5e5" }}>
          <div style={{ display: "flex", gap: 6 }}>
            <button onClick={comprobarYBorrar} disabled={!seleccionada || comprobandoUso || !!editar}
              style={{ padding: "7px 14px", borderRadius: 6, border: (seleccionada && !editar) ? "1px solid #fca5a5" : "1px solid #d4d4d4", background: (seleccionada && !editar) ? "#fef2f2" : "#f5f5f5", color: (seleccionada && !editar) ? "#991b1b" : "#737373", cursor: (seleccionada && !editar && !comprobandoUso) ? "pointer" : "default", fontSize: 12 }}>
              <BtnContent icon={comprobandoUso ? RefreshCw : Trash2} iconColor={(seleccionada && !editar) ? "#dc2626" : "#737373"}>
                {comprobandoUso ? "Comprobando..." : "Borrar"}
              </BtnContent>
            </button>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            {!editar ? (
              <button onClick={iniciarEdicion} disabled={!seleccionada}
                style={{ padding: "7px 14px", borderRadius: 6, border: seleccionada ? "1px solid #cbd5e1" : "1px solid #d4d4d4", background: seleccionada ? "#fff" : "#f5f5f5", color: seleccionada ? "#171717" : "#737373", cursor: seleccionada ? "pointer" : "default", fontSize: 12 }}>
                <BtnContent icon={Edit3} iconColor={seleccionada ? "#475569" : "#737373"}>Editar</BtnContent>
              </button>
            ) : (
              <>
                <button onClick={cancelarEdicion} disabled={guardando}
                  style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: guardando ? "default" : "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={guardarEdicion} disabled={guardando}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #16a34a", background: "#dcfce7", color: "#14532d", cursor: guardando ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={guardando ? RefreshCw : Check} iconColor="#14532d">{guardando ? "Guardando..." : "Aceptar"}</BtnContent>
                </button>
              </>
            )}
            {!editar && (
              <button onClick={onClose}
                style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
                <BtnContent icon={X}>Cerrar</BtnContent>
              </button>
            )}
          </div>
        </div>

        {/* Confirmar borrar */}
        {confirmBorrar && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }} onClick={() => !borrando && setConfirmBorrar(null)}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 450 }} onClick={e => e.stopPropagation()}>
              <h3 style={{ margin: 0, marginBottom: 12, fontSize: 15, fontWeight: 700, display: "inline-flex", alignItems: "center", gap: 8 }}>
                <Icon as={Trash2} size={18} color="#dc2626" /> ¿Borrar familia?
              </h3>
              <p style={{ fontSize: 13, color: "#525252", marginBottom: 12 }}>
                Vas a borrar la familia <strong>"{confirmBorrar.familia}"</strong>.
              </p>
              <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "10px 12px", fontSize: 12, color: "#991b1b", marginBottom: 16 }}>
                <strong>Atención:</strong> esta acción no se puede deshacer.
              </div>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmBorrar(null)} disabled={borrando}
                  style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", cursor: borrando ? "default" : "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={borrarConfirmado} disabled={borrando}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #dc2626", background: "#fef2f2", color: "#991b1b", cursor: borrando ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={borrando ? RefreshCw : Trash2} iconColor="#dc2626">{borrando ? "Borrando..." : "Sí, borrar"}</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Sección reutilizable: importar/actualizar una tabla desde Excel ──
// config = {
//   titulo, endpoint, icono, color,
//   columnas: [{ claves:[alias...], destino, label, obligatoria, tipo }],
//   buscarExistente: (fila, registros) => registroExistente | null,
//   construirNuevo: (datos) => bodyParaPOST,
//   etiquetaRegistro: (datos) => "texto",
// }
function ImportTablaSection({ setStatus, config }) {
  const [fileData, setFileData] = useState(null);   // { headers, rows, origen }
  const [mapping, setMapping] = useState({});        // idxColumna -> destino
  const [erroresMapeo, setErroresMapeo] = useState([]);
  const [log, setLog] = useState([]);
  const [procesando, setProcesando] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [sobreescribir, setSobreescribir] = useState(true);
  const logRef = useRef(null);
  const stopRequested = useRef(false);

  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight;
  }, [log]);

  const addLog = (texto, tipo = "info") => {
    setLog(l => [...l, { texto, tipo, hora: new Date().toLocaleTimeString("es-ES") }]);
  };

  // Mapea cabeceras del Excel a destinos según los alias de cada columna
  const mapearCabeceras = (headers) => {
    const map = {};
    headers.forEach((h, idx) => {
      const hNorm = String(h || "").trim().toLowerCase();
      if (!hNorm) return;
      for (const col of config.columnas) {
        if (col.claves.some(a => a.toLowerCase() === hNorm)) {
          map[idx] = col.destino;
          break;
        }
      }
    });
    return map;
  };

  const handleFile = async (file) => {
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (aoa.length === 0) { setErroresMapeo(["El fichero está vacío"]); return; }
    const headers = aoa[0].map(h => String(h || "").trim());
    const rows = aoa.slice(1).filter(r => r.some(c => String(c || "").trim() !== ""));
    const map = mapearCabeceras(headers);
    const destinos = Object.values(map);
    const errores = [];
    config.columnas.filter(c => c.obligatoria).forEach(c => {
      if (!destinos.includes(c.destino)) errores.push(`Falta la columna obligatoria "${c.label}"`);
    });
    setFileData({ headers, rows, origen: file.name });
    setMapping(map);
    setErroresMapeo(errores);
    setLog([]);
  };

  const onDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) handleFile(file);
  };

  const procesar = async () => {
    if (!fileData || erroresMapeo.length > 0 || procesando) return;
    stopRequested.current = false;
    setProcesando(true);
    setLog([]);
    setStatus && setStatus(`Procesando ${config.titulo}...`, "working");

    const idxPorDestino = {};
    Object.entries(mapping).forEach(([idx, destino]) => { idxPorDestino[destino] = Number(idx); });

    addLog(`Fichero "${fileData.origen}" — ${fileData.rows.length} filas. Columnas: ${Object.values(mapping).join(", ")}`, "info");

    // Cargar registros existentes y datos auxiliares
    let contexto = {};
    try {
      contexto = await config.cargarContexto ? await config.cargarContexto() : {};
    } catch (e) {
      addLog("AVISO cargando datos auxiliares: " + e.message, "warning");
    }

    let nuevos = 0, actualizados = 0, omitidos = 0, errores = 0;

    for (let i = 0; i < fileData.rows.length; i++) {
      if (stopRequested.current) {
        addLog(`⏹ Detenido por el usuario en la fila ${i + 1}`, "warning");
        break;
      }
      const row = fileData.rows[i];
      // Construir objeto datos a partir de columnas mapeadas
      const datos = {};
      Object.entries(idxPorDestino).forEach(([destino, idx]) => {
        const val = String(row[idx] ?? "").trim();
        if (val !== "") datos[destino] = val;
      });

      // Validar obligatorias por fila
      const faltan = config.columnas.filter(c => c.obligatoria && !datos[c.destino]);
      if (faltan.length > 0) {
        addLog(`Fila ${i + 1}: faltan campos obligatorios (${faltan.map(c => c.label).join(", ")}), se omite`, "warning");
        errores++;
        continue;
      }

      try {
        const resultado = await config.procesarFila(datos, contexto, sobreescribir, addLog, i + 1);
        if (resultado === "nuevo") nuevos++;
        else if (resultado === "actualizado") actualizados++;
        else if (resultado === "omitido") omitidos++;
        else if (resultado === "error") errores++;
      } catch (e) {
        addLog(`Fila ${i + 1}: ERROR ${e.message}`, "error");
        errores++;
      }
    }

    addLog(`Terminado: ${nuevos} nuevo(s), ${actualizados} actualizado(s), ${omitidos} omitido(s), ${errores} error(es)`, "success");
    setStatus && setStatus(`${config.titulo}: ${nuevos} nuevos, ${actualizados} actualizados`, "success");
    setProcesando(false);
  };

  return (
    <div style={{ marginBottom: 20, padding: "14px 18px", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8 }}>
      <h3 style={{ fontSize: 14, fontWeight: 700, color: "#171717", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
        <Icon as={config.icono} size={16} color={config.color || "#2563eb"} /> {config.titulo}
      </h3>
      <p style={{ fontSize: 12, color: "#525252", marginBottom: 12, lineHeight: 1.5 }}>
        {config.descripcion}
      </p>

      {/* Zona de arrastrar fichero */}
      <div
        onDragOver={e => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={onDrop}
        style={{ border: `2px dashed ${dragOver ? "#2563eb" : "#cbd5e1"}`, borderRadius: 8, padding: "18px", textAlign: "center", background: dragOver ? "#eff6ff" : "#f8fafc", marginBottom: 12 }}>
        <Icon as={Upload} size={20} color="#94a3b8" />
        <p style={{ fontSize: 12, color: "#64748b", margin: "6px 0" }}>
          Arrastra aquí el fichero Excel o
          <label style={{ color: "#2563eb", cursor: "pointer", fontWeight: 600 }}>
            {" "}selecciónalo
            <input type="file" accept=".xlsx,.xls" style={{ display: "none" }}
              onChange={e => handleFile(e.target.files[0])} />
          </label>
        </p>
        {fileData && <p style={{ fontSize: 11, color: "#16a34a", margin: 0 }}>✓ {fileData.origen} ({fileData.rows.length} filas)</p>}
      </div>

      {/* Mapeo de columnas detectado (estilo Tarifas) */}
      {fileData && (
        <div style={{ marginBottom: 12, padding: "10px 12px", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 6 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", marginBottom: 6 }}>
            Mapeo de columnas detectado <span style={{ color: "#64748b", fontWeight: 400 }}>({fileData.rows.length} filas en "{fileData.origen}")</span>:
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {fileData.headers.map((h, i) => (
              <div key={i} style={{ fontSize: 11, padding: "3px 8px", borderRadius: 4, background: mapping[i] ? "#dcfce7" : "#fee2e2", border: mapping[i] ? "1px solid #86efac" : "1px solid #fca5a5", color: mapping[i] ? "#14532d" : "#991b1b" }}>
                <strong>{h}</strong> {mapping[i] ? "→ " + mapping[i] : "(no reconocida)"}
              </div>
            ))}
          </div>
          {erroresMapeo.length > 0 && (
            <div style={{ marginTop: 8, padding: "6px 10px", background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 4, fontSize: 11, color: "#991b1b" }}>
              <strong>Errores:</strong>
              <ul style={{ margin: "4px 0 0 16px", padding: 0 }}>
                {erroresMapeo.map((e, i) => <li key={i}>{e}</li>)}
              </ul>
            </div>
          )}
        </div>
      )}

      {/* Selector sobreescribir + botón */}
      <div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 12, flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <label style={{ fontSize: 12, color: "#171717", fontWeight: 500 }}>Actualizar existentes:</label>
          <div style={{ display: "inline-flex", border: "1px solid #d4d4d4", borderRadius: 6, overflow: "hidden" }}>
            <button onClick={() => setSobreescribir(true)}
              style={{ padding: "4px 14px", fontSize: 11, border: "none", cursor: "pointer",
                background: sobreescribir ? "#171717" : "#fff",
                color: sobreescribir ? "#fff" : "#171717",
                fontWeight: sobreescribir ? 600 : 400 }}>
              SÍ
            </button>
            <button onClick={() => setSobreescribir(false)}
              style={{ padding: "4px 14px", fontSize: 11, border: "none", borderLeft: "1px solid #d4d4d4", cursor: "pointer",
                background: !sobreescribir ? "#171717" : "#fff",
                color: !sobreescribir ? "#fff" : "#171717",
                fontWeight: !sobreescribir ? 600 : 400 }}>
              NO
            </button>
          </div>
        </div>
        <button onClick={procesando ? () => { stopRequested.current = true; } : procesar}
          disabled={!fileData || erroresMapeo.length > 0}
          style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #2563eb", background: procesando ? "#fee2e2" : (fileData && erroresMapeo.length === 0 ? "#eff6ff" : "#f1f5f9"), color: procesando ? "#991b1b" : (fileData && erroresMapeo.length === 0 ? "#1e40af" : "#cbd5e1"), cursor: (fileData && erroresMapeo.length === 0) ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
          <BtnContent icon={procesando ? X : Upload} iconColor={procesando ? "#991b1b" : "#2563eb"}>
            {procesando ? "Detener" : "Procesar fichero"}
          </BtnContent>
        </button>
      </div>

      {/* Log de progreso detallado (siempre visible, como en Actualizar Tarifas) */}
      <div style={{ fontSize: 12, fontWeight: 600, color: "#171717", marginBottom: 4, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span>Progreso ({log.length} entrada{log.length !== 1 ? "s" : ""})</span>
        {log.length > 0 && (
          <div style={{ display: "inline-flex", gap: 6 }}>
            <button onClick={() => {
                const texto = log.map(l => {
                  const prefijo = l.tipo === "error" ? "[ERROR]" : l.tipo === "warning" ? "[AVISO]" : l.tipo === "success" ? "[OK]" : "[INFO]";
                  return `[${l.hora || ""}] ${prefijo} ${l.texto}`;
                }).join("\n");
                if (navigator.clipboard) navigator.clipboard.writeText(texto).then(() => setStatus && setStatus("Log copiado al portapapeles", "success")).catch(() => {});
              }}
              title="Copiar todo el log al portapapeles"
              style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
              <BtnContent icon={Copy} iconColor="#475569">Copiar log</BtnContent>
            </button>
            <button onClick={() => exportarLogCSV(log, "mantenimiento")}
              title="Exportar el log a un fichero CSV"
              style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
              <BtnContent icon={Download} iconColor="#475569">Exportar CSV</BtnContent>
            </button>
            <button onClick={() => setLog([])}
              style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
              <BtnContent icon={Trash2} iconColor="#475569">Limpiar log</BtnContent>
            </button>
          </div>
        )}
      </div>
      <div ref={logRef} style={{ height: 280, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, background: "#0f172a", color: "#e2e8f0", fontFamily: "monospace", fontSize: 11, padding: "8px 12px", lineHeight: 1.5 }}>
        {log.length === 0 ? (
          <div style={{ color: "#64748b", textAlign: "center", padding: "20px 0", fontStyle: "italic" }}>
            Sin actividad. Carga un fichero Excel y pulsa "Procesar fichero" para ver el progreso aquí.
          </div>
        ) : log.map((l, i) => (
          <div key={i} style={{ padding: "1px 0",
            color: l.tipo === "error" ? "#fca5a5"
              : l.tipo === "warning" ? "#fcd34d"
              : l.tipo === "success" ? "#86efac"
              : "#cbd5e1",
            fontWeight: l.tipo === "error" ? 600 : 400 }}>
            <span style={{ color: "#64748b" }}>[{l.hora}]</span>
            {l.tipo === "error" && <span style={{ color: "#fca5a5", fontWeight: 700 }}> ❌ </span>}
            {l.tipo === "warning" && <span style={{ color: "#fcd34d", fontWeight: 700 }}> ⚠ </span>}
            {l.tipo === "success" && <span style={{ color: "#86efac" }}> ✓ </span>}
            {l.tipo === "info" && " "}
            {l.texto}
          </div>
        ))}
      </div>
    </div>
  );
}

function MantenimientoSection({ setStatus }) {
  const [logsDialog, setLogsDialog] = useState(false);
  const [logs, setLogs] = useState([]);
  const [showFamiliasDialog, setShowFamiliasDialog] = useState(false);
  const [showSubfamiliasDialog, setShowSubfamiliasDialog] = useState(false);
  const [actualizandoMasusado, setActualizandoMasusado] = useState(false);
  const [ultimoResultadoMasusado, setUltimoResultadoMasusado] = useState(null);

  // ─── Comprobar integridad referencial de la BD ───
  const [integridadLog, setIntegridadLog] = useState([]);
  const [comprobandoIntegridad, setComprobandoIntegridad] = useState(false);
  const integridadLogRef = useRef(null);
  const integridadAddLog = (texto, tipo = "info") => {
    const hora = new Date().toLocaleTimeString("es-ES");
    setIntegridadLog(l => [...l, { texto, tipo, hora }]);
    setTimeout(() => { if (integridadLogRef.current) integridadLogRef.current.scrollTop = integridadLogRef.current.scrollHeight; }, 0);
  };

  const comprobarIntegridad = async () => {
    setComprobandoIntegridad(true);
    setIntegridadLog([]);
    setStatus && setStatus("Comprobando integridad de datos en BD...", "working");
    integridadAddLog("Iniciando comprobación de integridad referencial...", "info");
    try {
      const r = await fetch(`${API_URL}/integridad/comprobar`);
      if (!r.ok) throw new Error("HTTP " + r.status);
      const data = await r.json();
      integridadAddLog(`Comprobadas ${data.total_relaciones} relaciones entre tablas`, "info");
      integridadAddLog("", "info");

      (data.resultados || []).forEach(res => {
        if (res.error) {
          integridadAddLog(`⚠ ${res.relacion}: no se pudo comprobar (${res.error})`, "warning");
        } else if (res.huerfanos === 0) {
          integridadAddLog(`✓ ${res.relacion}: correcto`, "success");
        } else {
          integridadAddLog(`❌ ${res.relacion}: ${res.huerfanos} registro(s) huérfano(s) — ${res.tabla_origen}.${res.campo} apunta a ${res.tabla_destino} inexistente`, "error");
          res.muestra.forEach(m => {
            integridadAddLog(`      · ${res.tabla_origen} id=${m.id_origen} → ${res.campo}=${m.valor_fk} (no existe en ${res.tabla_destino})`, "error");
          });
          if (res.huerfanos > res.muestra.length) {
            integridadAddLog(`      · ... y ${res.huerfanos - res.muestra.length} más`, "error");
          }
        }
      });

      integridadAddLog("", "info");
      if (data.integro) {
        integridadAddLog(`✓ INTEGRIDAD CORRECTA: no se encontraron registros huérfanos`, "success");
        setStatus && setStatus("Integridad correcta: no hay datos huérfanos", "success");
      } else {
        integridadAddLog(`❌ Se encontraron ${data.total_huerfanos} registro(s) huérfano(s) en total`, "error");
        setStatus && setStatus(`Integridad: ${data.total_huerfanos} registros huérfanos encontrados`, "error");
      }
    } catch (e) {
      integridadAddLog("Error al comprobar integridad: " + e.message, "error");
      setStatus && setStatus("Error al comprobar integridad: " + e.message, "error");
    } finally {
      setComprobandoIntegridad(false);
    }
  };

  // ─── Asignar grupo descuento a productos por raíz ───
  const [asignarProcesando, setAsignarProcesando] = useState(false);
  const [asignarLog, setAsignarLog] = useState([]);
  const asignarLogRef = useRef(null);
  const asignarStopRef = useRef(false);
  const [confirmRaiz, setConfirmRaiz] = useState(null);  // { codigo, idgrupo, raiz, count, resolve }
  const asignarAddLog = (texto, tipo = "info") => {
    const hora = new Date().toLocaleTimeString("es-ES");
    setAsignarLog(l => [...l, { texto, tipo, hora }]);
    setTimeout(() => { if (asignarLogRef.current) asignarLogRef.current.scrollTop = asignarLogRef.current.scrollHeight; }, 0);
  };

  const ejecutarAsignacionGrupos = async () => {
    asignarStopRef.current = false;
    setAsignarProcesando(true);
    setStatus && setStatus("Cargando grupos de descuento...", "working");

    try {
      // 1. Cargar todos los grupos
      const resGrupos = await fetch(`${API_URL}/gruposdescuento/`);
      if (!resGrupos.ok) throw new Error("Error cargando grupos: " + resGrupos.status);
      const grupos = await resGrupos.json();
      asignarAddLog(`Cargados ${grupos.length} grupos de descuento`, "info");

      let totalActualizados = 0;
      let combinaciones = 0;
      let confirmadas = 0;
      let canceladas = 0;
      let canceladoTodo = false;

      for (const g of grupos) {
        if (asignarStopRef.current || canceladoTodo) break;
        if (!g.raizproductos || !String(g.raizproductos).trim()) continue;

        const raices = String(g.raizproductos).split(",").map(r => r.trim()).filter(Boolean);
        if (raices.length === 0) continue;

        for (const raiz of raices) {
          if (asignarStopRef.current || canceladoTodo) break;
          combinaciones++;

          // Contar previo
          let count = 0;
          try {
            const resC = await fetch(`${API_URL}/productos/contar-por-raiz?raiz=${encodeURIComponent(raiz)}`);
            if (resC.ok) {
              const d = await resC.json();
              count = d.productos || 0;
            }
          } catch {}

          if (count === 0) {
            asignarAddLog(`${g.grupodescuentospain} / raíz "${raiz}" → 0 productos coinciden, se omite`, "info");
            continue;
          }

          // Pedir confirmación por cada combinación grupo+raíz
          const decision = await new Promise(resolve => {
            setConfirmRaiz({
              codigo: g.grupodescuentospain,
              idgrupo: g.id,
              raiz,
              count,
              resolve,
            });
          });
          setConfirmRaiz(null);

          if (decision === "cancel_all") {
            asignarAddLog("⏹ Proceso cancelado por el usuario", "warning");
            canceladoTodo = true;
            break;
          }
          if (decision === "skip") {
            asignarAddLog(`${g.grupodescuentospain} / raíz "${raiz}" (${count} productos) → CANCELADO por el usuario`, "warning");
            canceladas++;
            continue;
          }
          // decision === "ok"
          confirmadas++;

          try {
            const resApl = await fetch(`${API_URL}/productos/aplicar-grupo-por-raiz`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ raiz, idgrupo: g.id }),
            });
            if (!resApl.ok) {
              const err = await resApl.json().catch(() => null);
              throw new Error(err?.detail || `HTTP ${resApl.status}`);
            }
            const d = await resApl.json();
            asignarAddLog(`${g.grupodescuentospain} / raíz "${raiz}" → ${d.actualizados} productos actualizados`, "success");
            totalActualizados += d.actualizados || 0;
          } catch (err) {
            asignarAddLog(`${g.grupodescuentospain} / raíz "${raiz}" → ERROR: ${err.message}`, "error");
          }
        }
      }

      asignarAddLog(`Terminado: ${combinaciones} combinaciones (${confirmadas} aplicadas, ${canceladas} canceladas). Total productos actualizados: ${totalActualizados}`, "success");
      setStatus && setStatus(`Asignación completada: ${totalActualizados} productos actualizados`, "success");
    } catch (e) {
      asignarAddLog("Error: " + e.message, "error");
      setStatus && setStatus("Error: " + e.message, "error");
    } finally {
      setAsignarProcesando(false);
      asignarStopRef.current = false;
    }
  };

  // ─── Grupos Descuento ───
  const [gdFileData, setGdFileData] = useState(null); // { headers, rows, mapping }
  const [gdSobreescribir, setGdSobreescribir] = useState(false);
  const [gdProcesando, setGdProcesando] = useState(false);
  const [gdDragOver, setGdDragOver] = useState(false);
  const [gdLog, setGdLog] = useState([]);
  const gdLogRef = useRef(null);
  const gdStopRef = useRef(false);
  const gdAddLog = (texto, tipo = "info") => {
    const hora = new Date().toLocaleTimeString("es-ES");
    setGdLog(l => [...l, { texto, tipo, hora }]);
    setTimeout(() => { if (gdLogRef.current) gdLogRef.current.scrollTop = gdLogRef.current.scrollHeight; }, 0);
  };

  // Columnas válidas con sus aliases (los mismos campos de la tabla gruposdescuento)
  const GD_COLS = [
    { key: "grupodescuentospain",      aliases: ["grupodescuentospain", "grupo descuento spain", "gd spain", "gdspain", "grupo descuento españa", "gd"] },
    { key: "grupodescuentocasamatriz", aliases: ["grupodescuentocasamatriz", "grupo descuento casa matriz", "gd casa matriz", "gdcasamatriz", "casa matriz"] },
    { key: "descripcion",              aliases: ["descripcion", "descripción", "desc"] },
    { key: "dtoreferencia",            aliases: ["dtoreferencia", "dto referencia", "dto ref", "descuento referencia"] },
    { key: "raizproductos",            aliases: ["raizproductos", "raíz productos", "raiz productos", "raiz", "raíz"] },
    { key: "subfamilia",               aliases: ["subfamilia", "sub familia", "sf"] },
    { key: "dgl1",                     aliases: ["dgl1", "dgl 1", "limite 1", "límite 1", "umbral 1", "umbral amarillo"] },
    { key: "dgl2",                     aliases: ["dgl2", "dgl 2", "limite 2", "límite 2", "umbral 2", "umbral rojo"] },
  ];

  const detectarColumnasGD = (headers) => {
    const norm = (s) => String(s || "").toLowerCase().replace(/[áéíóú]/g, c => ({á:"a",é:"e",í:"i",ó:"o",ú:"u"}[c])).replace(/[._-]/g, " ").trim();
    const mapping = {};
    headers.forEach((h, idx) => {
      const hn = norm(h);
      for (const col of GD_COLS) {
        if (col.aliases.some(a => norm(a) === hn)) {
          mapping[idx] = col.key;
          break;
        }
      }
    });
    return mapping;
  };

  const cargarGdFile = (file) => {
    const nombreFichero = file.name;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx", "xls", "csv", "tsv", "txt"].includes(ext)) {
      gdAddLog(`Tipo de fichero "${ext}" no soportado`, "error");
      return;
    }
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = ev.target.result;
        const wb = XLSX.read(data, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
        if (!aoa || aoa.length < 2) { gdAddLog("Fichero vacío o sin datos", "error"); return; }
        const headers = aoa[0].map(h => String(h || "").trim());
        const rows = aoa.slice(1).filter(r => r.some(c => String(c || "").trim() !== ""));
        const mapping = detectarColumnasGD(headers);
        setGdFileData({ headers, rows, mapping, nombreFichero });
        const detectadas = Object.values(mapping);
        gdAddLog(`Excel cargado: ${rows.length} filas. Columnas detectadas: ${detectadas.join(", ")}`, "success");
        if (!detectadas.includes("grupodescuentospain")) {
          gdAddLog("⚠ Falta columna obligatoria: grupodescuentospain", "error");
        }
        if (!detectadas.includes("subfamilia")) {
          gdAddLog("ℹ Sin columna 'subfamilia': solo se podrán actualizar grupos existentes, no crear nuevos", "warning");
        }
      } catch (err) {
        gdAddLog("Error leyendo fichero: " + err.message, "error");
      }
    };
    reader.readAsBinaryString(file);
  };

  const aplicarGd = async () => {
    if (!gdFileData || gdProcesando) return;

    // Validación inicial: referencia (grupodescuentospain) obligatoria + al menos 2 columnas
    const idxPorClave = {};
    Object.entries(gdFileData.mapping).forEach(([idx, clave]) => { idxPorClave[clave] = parseInt(idx); });
    const columnasDetectadas = Object.keys(idxPorClave);

    if (!columnasDetectadas.includes("grupodescuentospain")) {
      gdAddLog("ERROR: el fichero debe tener una columna 'grupodescuentospain'. Se cancela el proceso.", "error");
      setStatus && setStatus("Falta columna 'grupodescuentospain'", "error");
      return;
    }
    if (columnasDetectadas.length < 2) {
      gdAddLog("ERROR: el fichero debe tener al menos 2 columnas (grupodescuentospain + otra). Se cancela el proceso.", "error");
      setStatus && setStatus("El fichero debe tener al menos 2 columnas", "error");
      return;
    }

    gdStopRef.current = false;
    setGdProcesando(true);
    setStatus && setStatus("Procesando grupos de descuento...", "working");
    gdAddLog(`Procesando fichero "${gdFileData.nombreFichero || "(desconocido)"}"`, "info");
    gdAddLog(`Iniciando: ${gdFileData.rows.length} filas. Columnas detectadas: ${columnasDetectadas.join(", ")}`, "info");

    // Cargar subfamilias solo si la columna 'subfamilia' está presente (para resolver ids)
    let subfamilias = [];
    if (idxPorClave.subfamilia !== undefined) {
      try {
        const r = await fetch(`${API_URL}/subfamilias/`);
        subfamilias = await r.json();
        gdAddLog(`Cargadas ${subfamilias.length} subfamilias de BD para resolver ids`, "info");
      } catch (e) {
        gdAddLog("AVISO: no se pudieron cargar subfamilias de BD: " + e.message, "warning");
      }
    }

    let nuevos = 0, actualizados = 0, saltados = 0, errores = 0;

    for (let i = 0; i < gdFileData.rows.length; i++) {
      if (gdStopRef.current) {
        gdAddLog(`⏹ Detenido por el usuario en la fila ${i + 1} de ${gdFileData.rows.length}`, "warning");
        break;
      }
      const row = gdFileData.rows[i];
      const codigoGD = String(row[idxPorClave.grupodescuentospain] ?? "").trim();
      if (!codigoGD) { gdAddLog(`Fila ${i + 1}: sin grupodescuentospain, se omite`, "warning"); errores++; continue; }

      // Construir payload con SOLO los campos aportados en el Excel (no vacíos)
      const payload = {};
      const aportados = [];

      // Resolver subfamilia → idsubfamiliarelacionada (si la columna está y tiene valor)
      let idSubfamRel = null;
      if (idxPorClave.subfamilia !== undefined) {
        const subfamTxt = String(row[idxPorClave.subfamilia] ?? "").trim();
        if (subfamTxt) {
          const sf = subfamilias.find(s => s.subfamilia && s.subfamilia.toUpperCase() === subfamTxt.toUpperCase());
          if (!sf) {
            gdAddLog(`Fila ${i + 1}: ${codigoGD} → ERROR subfamilia "${subfamTxt}" no existe en BD`, "error");
            errores++; continue;
          }
          idSubfamRel = sf.id;
          payload.idsubfamiliarelacionada = sf.id;
          aportados.push("idsubfamiliarelacionada");
        }
      }

      if (idxPorClave.grupodescuentocasamatriz !== undefined) {
        const v = String(row[idxPorClave.grupodescuentocasamatriz] ?? "").trim();
        if (v) { payload.grupodescuentocasamatriz = v; aportados.push("grupodescuentocasamatriz"); }
      }
      if (idxPorClave.descripcion !== undefined) {
        const v = String(row[idxPorClave.descripcion] ?? "").trim();
        if (v) { payload.descripcion = v; aportados.push("descripcion"); }
      }
      if (idxPorClave.dtoreferencia !== undefined) {
        const raw = String(row[idxPorClave.dtoreferencia] ?? "").trim().replace(",", ".");
        const num = parseFloat(raw);
        if (!isNaN(num)) { payload.dtoreferencia = num; aportados.push("dtoreferencia"); }
      }
      if (idxPorClave.raizproductos !== undefined) {
        const v = String(row[idxPorClave.raizproductos] ?? "").trim();
        if (v) { payload.raizproductos = v; aportados.push("raizproductos"); }
      }
      if (idxPorClave.dgl1 !== undefined) {
        const raw = String(row[idxPorClave.dgl1] ?? "").trim().replace(",", ".");
        const num = parseFloat(raw);
        if (!isNaN(num)) { payload.dgl1 = num; aportados.push("dgl1"); }
      }
      if (idxPorClave.dgl2 !== undefined) {
        const raw = String(row[idxPorClave.dgl2] ?? "").trim().replace(",", ".");
        const num = parseFloat(raw);
        if (!isNaN(num)) { payload.dgl2 = num; aportados.push("dgl2"); }
      }

      try {
        // Buscar si existe
        const resBuscar = await fetch(`${API_URL}/gruposdescuento/${encodeURIComponent(codigoGD)}`);
        if (resBuscar.status === 404) {
          // NO EXISTE: dar de alta. Requiere grupodescuentospain + subfamilia (resuelta a id)
          if (idSubfamRel === null) {
            gdAddLog(`Fila ${i + 1}: ${codigoGD} → ERROR para crear hay que aportar columna 'subfamilia'`, "error");
            errores++; continue;
          }
          const body = { grupodescuentospain: codigoGD, ...payload };
          const res = await fetch(`${API_URL}/gruposdescuento/`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
          });
          if (!res.ok) {
            const err = await res.json().catch(() => null);
            throw new Error(err?.detail || `HTTP ${res.status}`);
          }
          gdAddLog(`Fila ${i + 1}: ${codigoGD} → ALTA (campos: ${aportados.join(", ")})`, "success");
          nuevos++;
        } else if (resBuscar.ok) {
          const existente = await resBuscar.json();
          if (!gdSobreescribir) {
            gdAddLog(`Fila ${i + 1}: ${codigoGD} → ya existe, se omite (sobreescribir = NO)`, "info");
            saltados++;
            continue;
          }
          // Actualizar — solo los campos aportados (vacíos no sobreescriben)
          if (aportados.length === 0) {
            gdAddLog(`Fila ${i + 1}: ${codigoGD} → solo se aportó código, nada que actualizar`, "info");
            continue;
          }
          // Body para PUT necesita el grupodescuentospain también
          const body = { grupodescuentospain: codigoGD, ...payload };
          const qsAportados = [...aportados, "grupodescuentospain"];
          const qs = qsAportados.map(c => `solo_aportados=${encodeURIComponent(c)}`).join("&");
          const res = await fetch(`${API_URL}/gruposdescuento/${existente.id}?${qs}`, {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
          });
          if (!res.ok) {
            const err = await res.json().catch(() => null);
            throw new Error(err?.detail || `HTTP ${res.status}`);
          }
          gdAddLog(`Fila ${i + 1}: ${codigoGD} → ACTUALIZADO (campos: ${aportados.join(", ")})`, "success");
          actualizados++;
        } else {
          throw new Error(`HTTP ${resBuscar.status} buscando ${codigoGD}`);
        }
      } catch (err) {
        gdAddLog(`Fila ${i + 1}: ${codigoGD} → ERROR ${err.message}`, "error");
        errores++;
      }
    }

    gdAddLog(`Terminado: ${nuevos} nuevos, ${actualizados} actualizados, ${saltados} omitidos, ${errores} errores`, "success");
    setStatus && setStatus(`Grupos descuento procesados: ${nuevos + actualizados} OK, ${errores} errores`, errores > 0 ? "error" : "success");
    setGdProcesando(false);
  };


  const actualizarMasusado = async () => {
    setActualizandoMasusado(true);
    setUltimoResultadoMasusado(null);
    setStatus && setStatus("Recalculando campo masusado de los productos...", "working");
    try {
      const res = await fetch(`${API_URL}/productos/actualizar-masusado`, { method: "POST" });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || `Error ${res.status}`);
      }
      const data = await res.json();
      setUltimoResultadoMasusado(data);
      setStatus && setStatus(`masusado recalculado: ${data.productos_actualizados} productos actualizados de ${data.referencias_procesadas} referencias procesadas`, "success");
    } catch (e) {
      setStatus && setStatus("Error recalculando masusado: " + e.message, "error");
    } finally {
      setActualizandoMasusado(false);
    }
  };

  const [cargandoLogs, setCargandoLogs] = useState(false);
  const [errorLogs, setErrorLogs] = useState(null);
  const [fechaCorte, setFechaCorte] = useState(() => {
    // Por defecto, hace 30 días
    const d = new Date();
    d.setDate(d.getDate() - 30);
    return d.toISOString().slice(0, 10);
  });
  const [confirmBorrarLogs, setConfirmBorrarLogs] = useState(false);
  const [borrando, setBorrando] = useState(false);
  const [exportandoSubfam, setExportandoSubfam] = useState(false);

  const cargarLogs = async () => {
    setCargandoLogs(true);
    setErrorLogs(null);
    try {
      const res = await fetch(`${API_URL}/logs/`);
      if (!res.ok) throw new Error("Error " + res.status);
      const data = await res.json();
      setLogs(data);
    } catch (e) {
      setErrorLogs(e.message);
      setLogs([]);
    } finally {
      setCargandoLogs(false);
    }
  };

  const abrirLogs = () => {
    setLogsDialog(true);
    cargarLogs();
  };

  const borrarLogsAnteriores = async () => {
    setBorrando(true);
    setStatus && setStatus(`Borrando logs anteriores a ${fechaCorte}...`, "working");
    try {
      const res = await fetch(`${API_URL}/logs/?antes_de=${fechaCorte}`, { method: "DELETE" });
      if (!res.ok) throw new Error("Error " + res.status);
      const data = await res.json();
      setStatus && setStatus(`${data.borrados ?? 0} logs borrados`, "success");
      setConfirmBorrarLogs(false);
      cargarLogs();
    } catch (e) {
      setStatus && setStatus("Error borrando logs: " + e.message, "error");
    } finally {
      setBorrando(false);
    }
  };

  const exportarSubfamilias = async () => {
    setExportandoSubfam(true);
    setStatus && setStatus("Generando Excel de Familia/SubFamilia...", "working");
    try {
      const res = await fetch(`${API_URL}/subfamilias/`);
      if (!res.ok) throw new Error("Error " + res.status);
      const subfamilias = await res.json();

      // Construir Excel: id, subfamilia, familia, idfamilia, descuento
      const aoa = [];
      aoa.push(["LISTADO DE SUBFAMILIAS Y FAMILIAS"]);
      aoa.push([`Fecha exportación: ${new Date().toLocaleDateString("es-ES", { day: "2-digit", month: "long", year: "numeric" })}`]);
      aoa.push([`Total: ${subfamilias.length} subfamilias`]);
      aoa.push([]);
      aoa.push(["Id SubFamilia", "SubFamilia", "Id Familia", "Familia", "Descuento %"]);
      // Ordenar por familia, luego subfamilia
      const sorted = [...subfamilias].sort((a, b) => {
        const fa = (a.familia || "").localeCompare(b.familia || "");
        if (fa !== 0) return fa;
        return (a.subfamilia || "").localeCompare(b.subfamilia || "");
      });
      sorted.forEach(s => {
        aoa.push([
          s.id,
          s.subfamilia || "",
          s.idfamilia,
          s.familia || "",
          s.descuento != null ? Number(s.descuento) : 0,
        ]);
      });

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws["!cols"] = [
        { wch: 14 }, { wch: 30 }, { wch: 12 }, { wch: 30 }, { wch: 12 },
      ];
      ws["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
      ];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "SubFamilias");
      descargarXLSX(wb, "SubFamilias.xlsx");

      setStatus && setStatus(`Excel generado con ${subfamilias.length} subfamilias`, "success");
    } catch (e) {
      setStatus && setStatus("Error exportando: " + e.message, "error");
    } finally {
      setExportandoSubfam(false);
    }
  };

  const formatearFechaLog = (f) => {
    if (!f) return "—";
    try {
      const d = new Date(f);
      return d.toLocaleString("es-ES", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit" });
    } catch { return String(f); }
  };

  return (
    <div>
      <h2 style={{ fontSize: 18, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, borderBottom: "2px solid #2563eb", paddingBottom: 8 }}>Mantenimiento BD</h2>
      <p style={{ color: "#475569", lineHeight: 1.6, marginBottom: 20, fontSize: 12 }}>
        Herramientas de mantenimiento de la base de datos.
      </p>

      {/* Apartado Log */}
      <div style={{ marginBottom: 20, padding: "14px 18px", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8 }}>
        <h3 style={{ fontSize: 14, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
          <Icon as={FileText} size={16} color="#2563eb" /> Mantenimiento tabla Log
        </h3>
        <p style={{ fontSize: 12, color: "#64748b", marginBottom: 12, lineHeight: 1.5 }}>
          Consulta el histórico de eventos de la aplicación y borra los registros antiguos para liberar espacio en la base de datos.
        </p>
        <button onClick={abrirLogs}
          style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #2563eb", background: "#eff6ff", color: "#1e40af", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
          <BtnContent icon={FileText} iconColor="#2563eb">Ver tabla de logs</BtnContent>
        </button>
      </div>

      {/* Apartado Productos: actualizar masusado */}
      <div style={{ marginBottom: 20, padding: "14px 18px", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8 }}>
        <h3 style={{ fontSize: 14, fontWeight: 700, color: "#171717", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
          <Icon as={Package} size={16} color="#2563eb" /> Mantenimiento tabla Productos
        </h3>
        <p style={{ fontSize: 12, color: "#525252", marginBottom: 12, lineHeight: 1.5 }}>
          Actualiza el campo <strong>masusado</strong> de cada producto contando cuántas veces aparece su referencia en la tabla <code>detallepresupuestos</code>. Solo se incrementa el valor si el nuevo número es mayor que el actual (nunca se disminuye, así no se pierde el histórico si se han borrado presupuestos antiguos).
        </p>
        <button onClick={actualizarMasusado} disabled={actualizandoMasusado}
          style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #2563eb", background: "#eff6ff", color: "#1e40af", cursor: actualizandoMasusado ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
          <BtnContent icon={actualizandoMasusado ? RefreshCw : TrendingUp} iconColor="#2563eb">
            {actualizandoMasusado ? "Calculando..." : "Recalcular campo masusado"}
          </BtnContent>
        </button>
        {ultimoResultadoMasusado && (
          <div style={{ marginTop: 12, padding: "10px 14px", background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 6, fontSize: 12, color: "#0369a1" }}>
            <div><strong>Resultado:</strong></div>
            <ul style={{ margin: "4px 0 0 16px", padding: 0, lineHeight: 1.6 }}>
              <li>Referencias distintas procesadas: <strong>{ultimoResultadoMasusado.referencias_procesadas}</strong></li>
              <li>Productos actualizados (nuevo valor mayor): <strong style={{ color: "#15803d" }}>{ultimoResultadoMasusado.productos_actualizados}</strong></li>
              <li>Sin cambio (valor actual ya era igual o mayor): <strong>{ultimoResultadoMasusado.sin_cambio}</strong></li>
              <li>Referencias sin producto en BD: <strong style={{ color: "#d97706" }}>{ultimoResultadoMasusado.referencias_sin_producto}</strong></li>
            </ul>
          </div>
        )}

      </div>

      {/* Apartado Familia/Subfamilia */}
      <div style={{ marginBottom: 20, padding: "14px 18px", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8 }}>
        <h3 style={{ fontSize: 14, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
          <Icon as={Layers} size={16} color="#16a34a" /> Mantenimiento tablas Familia / SubFamilia
        </h3>
        <p style={{ fontSize: 12, color: "#64748b", marginBottom: 12, lineHeight: 1.5 }}>
          Genera un Excel con el listado completo de subfamilias indicando a qué familia pertenecen y su descuento. Útil para revisar la estructura de productos.
        </p>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button onClick={exportarSubfamilias} disabled={exportandoSubfam}
            style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #16a34a", background: "#f0fdf4", color: "#14532d", cursor: exportandoSubfam ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={exportandoSubfam ? RefreshCw : Download} iconColor="#16a34a">{exportandoSubfam ? "Generando..." : "Exportar listado a Excel"}</BtnContent>
          </button>
          <button onClick={() => setShowFamiliasDialog(true)}
            style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={Wrench} iconColor="#475569">Mantenimiento BD Familia</BtnContent>
          </button>
          <button onClick={() => setShowSubfamiliasDialog(true)}
            style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={Wrench} iconColor="#475569">Mantenimiento BD Subfamilia</BtnContent>
          </button>
        </div>
        {showFamiliasDialog && <MantenimientoFamiliaDialog onClose={() => setShowFamiliasDialog(false)} setStatus={setStatus} />}
        {showSubfamiliasDialog && <MantenimientoSubfamiliaDialog onClose={() => setShowSubfamiliasDialog(false)} setStatus={setStatus} />}
      </div>

      {/* Apartado Mantenimiento tabla Clientes (importar Excel) */}
      <ImportTablaSection
        setStatus={setStatus}
        config={{
          titulo: "Mantenimiento tabla Clientes",
          descripcion: "Importa o actualiza clientes desde un Excel. Mínimo: Razón social O Nombre común (al menos uno). Otras columnas reconocidas: NIF, IFA, Dirección, Población, CP, Teléfono, Provincia. Al crear un cliente con solo razón social, se usa también como nombre común.",
          icono: Users,
          color: "#0891b2",
          columnas: [
            { claves: ["nombre comun", "nombrecomun", "nombre común", "nombre"], destino: "nombrecomun", label: "Nombre común", obligatoria: false },
            { claves: ["razon social", "razonsocial", "razón social"], destino: "razonsocial", label: "Razón social", obligatoria: false },
            { claves: ["nif", "cif"], destino: "nif", label: "NIF", obligatoria: false },
            { claves: ["ifa"], destino: "ifa", label: "IFA", obligatoria: false },
            { claves: ["direccion", "dirección"], destino: "direccion", label: "Dirección", obligatoria: false },
            { claves: ["poblacion", "población"], destino: "poblacion", label: "Población", obligatoria: false },
            { claves: ["cp", "codigo postal", "código postal", "c.p."], destino: "cp", label: "CP", obligatoria: false },
            { claves: ["telefono", "teléfono", "telefono1", "tlf", "tel"], destino: "telefono1", label: "Teléfono", obligatoria: false },
            { claves: ["provincia"], destino: "_provincia_nombre", label: "Provincia", obligatoria: false },
          ],
          cargarContexto: async () => {
            const ctx = { clientes: [], provincias: [] };
            try {
              const rc = await fetch(`${API_URL}/clientes/`);
              if (rc.ok) ctx.clientes = await rc.json();
            } catch {}
            try {
              const rp = await fetch(`${API_URL}/provincias/`);
              if (rp.ok) ctx.provincias = await rp.json();
            } catch {}
            return ctx;
          },
          procesarFila: async (datos, ctx, sobreescribir, addLog, nFila) => {
            const razon = String(datos.razonsocial || "").trim();
            const nombreC = String(datos.nombrecomun || "").trim();
            // Validación: al menos uno de los dos
            if (!razon && !nombreC) {
              addLog(`Fila ${nFila}: sin razón social ni nombre común, se omite`, "warning");
              return "error";
            }
            // Resolver provincia (nombre → id) si viene
            let idprovincia = null;
            if (datos._provincia_nombre) {
              const txt = datos._provincia_nombre.toLowerCase();
              let prov = (ctx.provincias || []).find(p => String(p.provincia).toLowerCase() === txt);
              if (!prov) prov = (ctx.provincias || []).find(p => String(p.provincia).toLowerCase().startsWith(txt));
              idprovincia = prov ? prov.id : null;
              if (!prov) addLog(`Fila ${nFila}: provincia "${datos._provincia_nombre}" no encontrada, se deja vacía`, "warning");
            }
            // cp e ifa son INTEGER en BD → convertir o null
            const toInt = (v) => {
              if (v == null || String(v).trim() === "") return null;
              const n = parseInt(String(v).replace(/[^0-9-]/g, ""), 10);
              return isNaN(n) ? null : n;
            };

            // Buscar existente: por razón social primero, luego por nombre común (case-insensitive)
            let existente = null;
            if (razon) {
              existente = (ctx.clientes || []).find(c => String(c.razonsocial || "").trim().toLowerCase() === razon.toLowerCase());
            }
            if (!existente && nombreC) {
              existente = (ctx.clientes || []).find(c => String(c.nombrecomun || "").trim().toLowerCase() === nombreC.toLowerCase());
            }
            const etiqueta = razon || nombreC;

            if (existente) {
              if (!sobreescribir) {
                addLog(`Fila ${nFila}: "${etiqueta}" → ya existe, se omite (actualizar = NO)`, "info");
                return "omitido";
              }
              // Actualizar: respetar valores existentes si la celda del Excel viene vacía
              const body = {
                nombrecomun: nombreC || existente.nombrecomun || razon || null,
                razonsocial: razon || existente.razonsocial || null,
                nif: datos.nif || existente.nif || null,
                ifa: datos.ifa != null && String(datos.ifa).trim() !== "" ? toInt(datos.ifa) : (existente.ifa ?? null),
                direccion: datos.direccion || existente.direccion || null,
                poblacion: datos.poblacion || existente.poblacion || null,
                cp: datos.cp != null && String(datos.cp).trim() !== "" ? toInt(datos.cp) : (existente.cp ?? null),
                telefono1: datos.telefono1 || existente.telefono1 || null,
                idprovincia: idprovincia != null ? idprovincia : (existente.idprovincia ?? null),
              };
              const r = await fetch(`${API_URL}/clientes/${existente.id}`, {
                method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
              });
              if (!r.ok) { const e = await r.json().catch(() => null); throw new Error(e?.detail || "HTTP " + r.status); }
              addLog(`Fila ${nFila}: "${etiqueta}" → ACTUALIZADO`, "success");
              return "actualizado";
            } else {
              // Crear nuevo. Si solo hay razón social, el nombre común será la razón social (y viceversa)
              const body = {
                nombrecomun: nombreC || razon || null,
                razonsocial: razon || nombreC || null,
                nif: datos.nif || null,
                ifa: toInt(datos.ifa),
                direccion: datos.direccion || null,
                poblacion: datos.poblacion || null,
                cp: toInt(datos.cp),
                telefono1: datos.telefono1 || null,
                idprovincia,
              };
              const r = await fetch(`${API_URL}/clientes/`, {
                method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
              });
              if (!r.ok) { const e = await r.json().catch(() => null); throw new Error(e?.detail || "HTTP " + r.status); }
              const nuevo = await r.json();
              if (nuevo && nuevo.id) ctx.clientes.push(nuevo); // para detectar duplicados dentro del mismo Excel
              addLog(`Fila ${nFila}: "${etiqueta}" → CREADO`, "success");
              return "nuevo";
            }
          },
        }}
      />

      {/* Apartado Mantenimiento tabla Contactos (importar Excel) */}
      <ImportTablaSection
        setStatus={setStatus}
        config={{
          titulo: "Mantenimiento tabla Contactos",
          descripcion: "Importa o actualiza contactos desde un Excel. Columnas mínimas obligatorias: Nombre y Cliente (nombre común del cliente). Otras columnas reconocidas: Email, Cargo, Teléfono 1, Teléfono 2.",
          icono: User,
          color: "#0891b2",
          columnas: [
            { claves: ["nombre"], destino: "nombre", label: "Nombre", obligatoria: true },
            { claves: ["cliente", "nombre cliente", "nombre comun cliente"], destino: "_cliente_nombre", label: "Cliente", obligatoria: true },
            { claves: ["email", "correo", "e-mail"], destino: "email", label: "Email", obligatoria: false },
            { claves: ["cargo", "puesto"], destino: "cargo", label: "Cargo", obligatoria: false },
            { claves: ["telefono1", "telefono", "teléfono", "teléfono 1", "telefono 1", "tlf", "tel", "tel1"], destino: "telefono1", label: "Teléfono 1", obligatoria: false },
            { claves: ["telefono2", "teléfono 2", "telefono 2", "tlf2", "tel2", "movil", "móvil"], destino: "telefono2", label: "Teléfono 2", obligatoria: false },
          ],
          cargarContexto: async () => {
            const ctx = { contactos: [], clientes: [] };
            try {
              const rc = await fetch(`${API_URL}/contactos/`);
              if (rc.ok) ctx.contactos = await rc.json();
            } catch {}
            try {
              const rcl = await fetch(`${API_URL}/clientes/`);
              if (rcl.ok) ctx.clientes = await rcl.json();
            } catch {}
            return ctx;
          },
          procesarFila: async (datos, ctx, sobreescribir, addLog, nFila) => {
            // Resolver cliente (nombre común → id)
            const txt = String(datos._cliente_nombre).trim().toLowerCase();
            let cli = (ctx.clientes || []).find(c => String(c.nombrecomun || "").trim().toLowerCase() === txt);
            if (!cli) cli = (ctx.clientes || []).find(c => String(c.razonsocial || "").trim().toLowerCase() === txt);
            if (!cli) {
              addLog(`Fila ${nFila}: cliente "${datos._cliente_nombre}" no encontrado, se omite`, "error");
              return "error";
            }
            const body = {
              nombre: datos.nombre || null,
              email: datos.email || null,
              cargo: datos.cargo || null,
              telefono1: datos.telefono1 || null,
              telefono2: datos.telefono2 || null,
              idcliente: cli.id,
            };
            // Buscar contacto existente por nombre + idcliente
            const existente = (ctx.contactos || []).find(c =>
              String(c.nombre || "").trim().toLowerCase() === String(datos.nombre).trim().toLowerCase() &&
              String(c.idcliente) === String(cli.id)
            );
            if (existente) {
              if (!sobreescribir) {
                addLog(`Fila ${nFila}: "${datos.nombre}" (${cli.nombrecomun}) → ya existe, se omite (actualizar = NO)`, "info");
                return "omitido";
              }
              const r = await fetch(`${API_URL}/contactos/${existente.id}`, {
                method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
              });
              if (!r.ok) { const e = await r.json().catch(() => null); throw new Error(e?.detail || "HTTP " + r.status); }
              addLog(`Fila ${nFila}: "${datos.nombre}" (${cli.nombrecomun}) → ACTUALIZADO`, "success");
              return "actualizado";
            } else {
              const r = await fetch(`${API_URL}/contactos/`, {
                method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
              });
              if (!r.ok) { const e = await r.json().catch(() => null); throw new Error(e?.detail || "HTTP " + r.status); }
              const nuevo = await r.json();
              if (nuevo && nuevo.id) ctx.contactos.push({ ...body, id: nuevo.id });
              addLog(`Fila ${nFila}: "${datos.nombre}" (${cli.nombrecomun}) → CREADO`, "success");
              return "nuevo";
            }
          },
        }}
      />

      {/* Apartado Comprobar integridad de datos */}
      <div style={{ marginBottom: 20, padding: "14px 18px", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8 }}>
        <h3 style={{ fontSize: 14, fontWeight: 700, color: "#171717", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
          <Icon as={ClipboardCheck} size={16} color="#dc2626" /> Comprobar integridad datos en BD
        </h3>
        <p style={{ fontSize: 12, color: "#525252", marginBottom: 12, lineHeight: 1.5 }}>
          Revisa todas las relaciones entre tablas y detecta registros "huérfanos": campos que apuntan a un id que ya no existe en otra tabla
          (por ejemplo, un grupo descuento con una subfamilia relacionada que no existe, un contacto con un cliente borrado, un detalle de presupuesto sin su presupuesto, etc.).
        </p>
        <button onClick={comprobarIntegridad} disabled={comprobandoIntegridad}
          style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #dc2626", background: comprobandoIntegridad ? "#f5f5f5" : "#fef2f2", color: comprobandoIntegridad ? "#737373" : "#991b1b", cursor: comprobandoIntegridad ? "default" : "pointer", fontSize: 12, fontWeight: 600, marginBottom: 12 }}>
          <BtnContent icon={comprobandoIntegridad ? RefreshCw : ClipboardCheck} iconColor={comprobandoIntegridad ? "#737373" : "#dc2626"}>
            {comprobandoIntegridad ? "Comprobando..." : "Comprobar integridad"}
          </BtnContent>
        </button>
        {integridadLog.length > 0 && (
          <div style={{ fontSize: 12, fontWeight: 600, color: "#171717", margin: "8px 0 4px", display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 6 }}>
            <button onClick={() => {
                const texto = integridadLog.map(l => `[${l.hora || ""}] ${l.texto}`).join("\n");
                if (navigator.clipboard) navigator.clipboard.writeText(texto).then(() => setStatus && setStatus("Log copiado al portapapeles", "success")).catch(() => {});
              }}
              title="Copiar todo el log al portapapeles"
              style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
              <BtnContent icon={Copy} iconColor="#475569">Copiar log</BtnContent>
            </button>
            <button onClick={() => exportarLogCSV(integridadLog, "integridad")}
              title="Exportar el log a un fichero CSV"
              style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
              <BtnContent icon={Download} iconColor="#475569">Exportar CSV</BtnContent>
            </button>
            <button onClick={() => setIntegridadLog([])}
              style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
              <BtnContent icon={Trash2} iconColor="#475569">Limpiar log</BtnContent>
            </button>
          </div>
        )}
        <div ref={integridadLogRef} style={{ height: 280, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, background: "#0f172a", color: "#e2e8f0", fontFamily: "monospace", fontSize: 11, padding: "8px 12px", lineHeight: 1.5 }}>
          {integridadLog.length === 0 ? (
            <div style={{ color: "#64748b", textAlign: "center", padding: "20px 0", fontStyle: "italic" }}>
              Sin actividad. Pulsa "Comprobar integridad" para revisar las relaciones entre tablas.
            </div>
          ) : integridadLog.map((l, i) => (
            <div key={i} style={{ padding: "1px 0", whiteSpace: "pre-wrap",
              color: l.tipo === "error" ? "#fca5a5"
                : l.tipo === "warning" ? "#fcd34d"
                : l.tipo === "success" ? "#86efac"
                : "#cbd5e1",
              fontWeight: l.tipo === "error" && !l.texto.startsWith("      ") ? 600 : 400 }}>
              {l.texto ? <><span style={{ color: "#64748b" }}>[{l.hora}]</span> {l.texto}</> : " "}
            </div>
          ))}
        </div>
      </div>

      {/* Apartado Grupos Descuento */}
      <div style={{ marginBottom: 20, padding: "14px 18px", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8 }}>
        <h3 style={{ fontSize: 14, fontWeight: 700, color: "#171717", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
          <Icon as={Percent} size={16} color="#7c3aed" /> Mantenimiento tabla Grupos Descuento
        </h3>
        <p style={{ fontSize: 12, color: "#525252", marginBottom: 12, lineHeight: 1.5 }}>
          Carga un Excel para crear o actualizar grupos de descuento. La primera fila debe tener los nombres de los campos. Obligatorios: <strong>grupodescuentospain</strong>. Para <em>crear nuevos</em> grupos, también es necesaria la columna <strong>subfamilia</strong>. Opcionales: <code>grupodescuentocasamatriz</code>, <code>descripcion</code>, <code>dtoreferencia</code>, <code>raizproductos</code>, <code>dgl1</code>, <code>dgl2</code>, <code>dgl1</code>, <code>dgl2</code>. Las celdas vacías del Excel NO sobreescriben valores existentes.
        </p>

        {/* Drag & drop */}
        <div
          onDragOver={e => { e.preventDefault(); setGdDragOver(true); }}
          onDragLeave={e => { e.preventDefault(); setGdDragOver(false); }}
          onDrop={e => {
            e.preventDefault();
            setGdDragOver(false);
            const file = e.dataTransfer.files?.[0];
            if (file) cargarGdFile(file);
          }}
          style={{
            border: `2px dashed ${gdDragOver ? "#7c3aed" : "#d4d4d4"}`,
            borderRadius: 8,
            padding: "20px",
            textAlign: "center",
            background: gdDragOver ? "#f5f3ff" : "#fafafa",
            marginBottom: 12,
            cursor: "pointer",
          }}
          onClick={() => document.getElementById("gd-file-input")?.click()}
        >
          <input id="gd-file-input" type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" style={{ display: "none" }}
            onChange={e => { if (e.target.files?.[0]) cargarGdFile(e.target.files[0]); e.target.value = ""; }} />
          <Icon as={Download} size={24} color="#7c3aed" />
          <div style={{ marginTop: 6, fontSize: 12, color: "#525252" }}>
            {gdFileData
              ? <><strong>{gdFileData.rows.length} filas</strong> cargadas. Pulsa el botón "Aplicar" para procesar.</>
              : <>Arrastra aquí un Excel/CSV, o haz click para elegir uno.</>}
          </div>
          {gdFileData && (
            <div style={{ marginTop: 6, fontSize: 11, color: "#737373" }}>
              Columnas detectadas: {Object.values(gdFileData.mapping).join(", ") || "ninguna"}
            </div>
          )}
        </div>

        {/* Mapeo de columnas con tags verde/rojo (igual estilo que Actualizar Tarifas) */}
        {gdFileData && (() => {
          const erroresMapeoGD = [];
          // Cabeceras no reconocidas
          gdFileData.headers.forEach((h, i) => {
            if (!gdFileData.mapping[i] && String(h || "").trim() !== "") {
              erroresMapeoGD.push(`Columna "${h}" no se reconoce`);
            }
          });
          // Obligatoria: grupodescuentospain
          const detectadas = Object.values(gdFileData.mapping);
          if (!detectadas.includes("grupodescuentospain")) {
            erroresMapeoGD.push("Falta columna 'grupodescuentospain' (obligatoria)");
          }
          // Al menos 2 columnas reconocidas
          if (detectadas.length < 2) {
            erroresMapeoGD.push("El fichero debe tener al menos 2 columnas reconocidas (grupodescuentospain + otra)");
          }
          return (
            <div style={{ marginBottom: 12, padding: "10px 12px", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 6 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", marginBottom: 6 }}>
                Mapeo de columnas detectado <span style={{ color: "#64748b", fontWeight: 400 }}>({gdFileData.rows.length} filas en "{gdFileData.nombreFichero}")</span>:
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {gdFileData.headers.map((h, i) => (
                  <div key={i} style={{ fontSize: 11, padding: "3px 8px", borderRadius: 4, background: gdFileData.mapping[i] ? "#dcfce7" : "#fee2e2", border: gdFileData.mapping[i] ? "1px solid #86efac" : "1px solid #fca5a5", color: gdFileData.mapping[i] ? "#14532d" : "#991b1b" }}>
                    <strong>{h}</strong> {gdFileData.mapping[i] ? "→ " + gdFileData.mapping[i] : "(no reconocida)"}
                  </div>
                ))}
              </div>
              {erroresMapeoGD.length > 0 && (
                <div style={{ marginTop: 8, padding: "6px 10px", background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 4, fontSize: 11, color: "#991b1b" }}>
                  <strong>Errores:</strong>
                  <ul style={{ margin: "4px 0 0 16px", padding: 0 }}>
                    {erroresMapeoGD.map((e, i) => <li key={i}>{e}</li>)}
                  </ul>
                </div>
              )}
            </div>
          );
        })()}

        {/* Sobreescribir + Aplicar + Parar */}
        <div style={{ display: "flex", gap: 12, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
          <label style={{ fontSize: 12, color: "#171717", fontWeight: 500 }}>Si el Grupo de descuento existe, sobreescribir:</label>
          <div style={{ display: "inline-flex", border: "1px solid #d4d4d4", borderRadius: 6, overflow: "hidden" }}>
            <button onClick={() => setGdSobreescribir(true)}
              style={{ padding: "4px 14px", fontSize: 11, border: "none", cursor: "pointer",
                background: gdSobreescribir ? "#171717" : "#fff",
                color: gdSobreescribir ? "#fff" : "#171717",
                fontWeight: gdSobreescribir ? 600 : 400 }}>
              SÍ
            </button>
            <button onClick={() => setGdSobreescribir(false)}
              style={{ padding: "4px 14px", fontSize: 11, border: "none", borderLeft: "1px solid #d4d4d4", cursor: "pointer",
                background: !gdSobreescribir ? "#171717" : "#fff",
                color: !gdSobreescribir ? "#fff" : "#171717",
                fontWeight: !gdSobreescribir ? 600 : 400 }}>
              NO
            </button>
          </div>
          {gdFileData && (
            <>
              <button onClick={aplicarGd} disabled={gdProcesando}
                style={{ padding: "7px 18px", borderRadius: 6, border: gdProcesando ? "1px solid #d4d4d4" : "1px solid #16a34a", background: gdProcesando ? "#f5f5f5" : "#dcfce7", color: gdProcesando ? "#737373" : "#14532d", cursor: gdProcesando ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
                <BtnContent icon={gdProcesando ? RefreshCw : Check} iconColor={gdProcesando ? "#737373" : "#14532d"}>
                  {gdProcesando ? "Procesando..." : "Aplicar a BD"}
                </BtnContent>
              </button>
              {gdProcesando && (
                <button onClick={() => { gdStopRef.current = true; setStatus && setStatus("Solicitando parada...", "working"); }}
                  style={{ padding: "7px 18px", borderRadius: 6, border: "1px solid #dc2626", background: "#fef2f2", color: "#991b1b", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={X} iconColor="#dc2626">Parar</BtnContent>
                </button>
              )}
              <button onClick={() => { setGdFileData(null); setGdLog([]); }}
                style={{ padding: "5px 10px", fontSize: 11, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                <BtnContent icon={X} iconColor="#475569">Limpiar fichero</BtnContent>
              </button>
            </>
          )}
        </div>

        {/* Log */}
        <div>
          <div style={{ fontSize: 12, fontWeight: 600, color: "#171717", marginBottom: 4, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span>Progreso ({gdLog.length} entrada{gdLog.length !== 1 ? "s" : ""})</span>
            {gdLog.length > 0 && (
              <div style={{ display: "inline-flex", gap: 6 }}>
                <button onClick={() => {
                    const texto = gdLog.map(l => {
                      const prefijo = l.tipo === "error" ? "[ERROR]" : l.tipo === "warning" ? "[AVISO]" : l.tipo === "success" ? "[OK]" : "[INFO]";
                      return `[${l.hora}] ${prefijo} ${l.texto}`;
                    }).join("\n");
                    if (navigator.clipboard) {
                      navigator.clipboard.writeText(texto)
                        .then(() => setStatus && setStatus(`Log (${gdLog.length} líneas) copiado al portapapeles`, "success"))
                        .catch(() => setStatus && setStatus("No se pudo copiar", "error"));
                    }
                  }}
                  style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                  <BtnContent icon={Copy} iconColor="#475569">Copiar log</BtnContent>
                </button>
                <button onClick={() => exportarLogCSV(gdLog, "grupos-descuento")}
                title="Exportar el log a un fichero CSV"
                style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                <BtnContent icon={Download} iconColor="#475569">Exportar CSV</BtnContent>
              </button>
              <button onClick={() => setGdLog([])}
                  style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                  <BtnContent icon={Trash2} iconColor="#475569">Limpiar log</BtnContent>
                </button>
              </div>
            )}
          </div>
          <div ref={gdLogRef} style={{ height: 280, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, background: "#0f172a", color: "#e2e8f0", fontFamily: "monospace", fontSize: 11, padding: "8px 12px", lineHeight: 1.5 }}>
            {gdLog.length === 0 ? (
              <div style={{ color: "#64748b", textAlign: "center", padding: "20px 0", fontStyle: "italic" }}>
                Sin actividad. Carga un Excel y pulsa "Aplicar a BD" para ver el progreso.
              </div>
            ) : gdLog.map((l, i) => (
              <div key={i} style={{ padding: "1px 0",
                color: l.tipo === "error" ? "#fca5a5"
                  : l.tipo === "warning" ? "#fcd34d"
                  : l.tipo === "success" ? "#86efac"
                  : "#cbd5e1",
                fontWeight: l.tipo === "error" ? 600 : 400 }}>
                <span style={{ color: "#64748b" }}>[{l.hora}]</span>
                {l.tipo === "error" && <span style={{ color: "#fca5a5", fontWeight: 700 }}> ❌ </span>}
                {l.tipo === "warning" && <span style={{ color: "#fcd34d", fontWeight: 700 }}> ⚠ </span>}
                {l.tipo === "success" && <span style={{ color: "#86efac" }}> ✓ </span>}
                {l.tipo === "info" && " "}
                {l.texto}
              </div>
            ))}
          </div>
        </div>
        {/* --- Asignar grupodescuento por raíz --- */}
        <div style={{ marginTop: 16, paddingTop: 14, borderTop: "1px dashed #e5e5e5" }}>
          <h4 style={{ fontSize: 13, fontWeight: 600, color: "#171717", marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
            <Icon as={Percent} size={14} color="#7c3aed" /> Asignar grupo descuento por raíz
          </h4>
          <p style={{ fontSize: 12, color: "#525252", marginBottom: 10, lineHeight: 1.5 }}>
            Para cada grupo de descuento, lee el campo <code>raizproductos</code> (raíces separadas por comas) y asigna ese grupo a todos los productos cuya referencia empiece por cada raíz. Se sobreescribe el grupo previo del producto si tenía uno asignado. <strong>Se pedirá confirmación por cada combinación grupo + raíz.</strong>
          </p>
          <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
            <button onClick={ejecutarAsignacionGrupos} disabled={asignarProcesando}
              style={{ padding: "7px 16px", borderRadius: 6, border: asignarProcesando ? "1px solid #d4d4d4" : "1px solid #7c3aed", background: asignarProcesando ? "#f5f5f5" : "#f5f3ff", color: asignarProcesando ? "#737373" : "#5b21b6", cursor: asignarProcesando ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
              <BtnContent icon={asignarProcesando ? RefreshCw : Percent} iconColor={asignarProcesando ? "#737373" : "#7c3aed"}>
                {asignarProcesando ? "Procesando..." : "Asignar grupos descuento por raíz"}
              </BtnContent>
            </button>
            {asignarProcesando && (
              <button onClick={() => { asignarStopRef.current = true; setStatus && setStatus("Solicitando parada...", "working"); }}
                style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #dc2626", background: "#fef2f2", color: "#991b1b", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                <BtnContent icon={X} iconColor="#dc2626">Parar</BtnContent>
              </button>
            )}
          </div>

          {/* Log */}
          <div style={{ fontSize: 12, fontWeight: 600, color: "#171717", marginBottom: 4, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span>Progreso ({asignarLog.length} entrada{asignarLog.length !== 1 ? "s" : ""})</span>
            {asignarLog.length > 0 && (
              <div style={{ display: "inline-flex", gap: 6 }}>
                <button onClick={() => {
                    const texto = asignarLog.map(l => {
                      const prefijo = l.tipo === "error" ? "[ERROR]" : l.tipo === "warning" ? "[AVISO]" : l.tipo === "success" ? "[OK]" : "[INFO]";
                      return `[${l.hora}] ${prefijo} ${l.texto}`;
                    }).join("\n");
                    if (navigator.clipboard) {
                      navigator.clipboard.writeText(texto)
                        .then(() => setStatus && setStatus(`Log (${asignarLog.length} líneas) copiado al portapapeles`, "success"))
                        .catch(() => setStatus && setStatus("No se pudo copiar", "error"));
                    }
                  }}
                  style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                  <BtnContent icon={Copy} iconColor="#475569">Copiar log</BtnContent>
                </button>
                <button onClick={() => exportarLogCSV(asignarLog, "asignar-grupos")}
                title="Exportar el log a un fichero CSV"
                style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                <BtnContent icon={Download} iconColor="#475569">Exportar CSV</BtnContent>
              </button>
              <button onClick={() => setAsignarLog([])}
                  style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                  <BtnContent icon={Trash2} iconColor="#475569">Limpiar log</BtnContent>
                </button>
              </div>
            )}
          </div>
          <div ref={asignarLogRef} style={{ height: 220, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, background: "#0f172a", color: "#e2e8f0", fontFamily: "monospace", fontSize: 11, padding: "8px 12px", lineHeight: 1.5 }}>
            {asignarLog.length === 0 ? (
              <div style={{ color: "#64748b", textAlign: "center", padding: "20px 0", fontStyle: "italic" }}>
                Sin actividad. Pulsa "Asignar grupos descuento por raíz" para empezar.
              </div>
            ) : asignarLog.map((l, i) => (
              <div key={i} style={{ padding: "1px 0",
                color: l.tipo === "error" ? "#fca5a5"
                  : l.tipo === "warning" ? "#fcd34d"
                  : l.tipo === "success" ? "#86efac"
                  : "#cbd5e1",
                fontWeight: l.tipo === "error" ? 600 : 400 }}>
                <span style={{ color: "#64748b" }}>[{l.hora}]</span>
                {l.tipo === "error" && <span style={{ color: "#fca5a5", fontWeight: 700 }}> ❌ </span>}
                {l.tipo === "warning" && <span style={{ color: "#fcd34d", fontWeight: 700 }}> ⚠ </span>}
                {l.tipo === "success" && <span style={{ color: "#86efac" }}> ✓ </span>}
                {l.tipo === "info" && " "}
                {l.texto}
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Diálogo confirmación asignación grupo por raíz */}
      {confirmRaiz && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}
          onClick={() => { confirmRaiz.resolve("skip"); }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 500, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, borderBottom: "1px solid #e5e5e5", paddingBottom: 10 }}>
              <h3 style={{ fontSize: 14, fontWeight: 700, margin: 0, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
                <Icon as={Percent} size={16} color="#7c3aed" /> Confirmar asignación
              </h3>
            </div>
            <p style={{ fontSize: 13, color: "#171717", lineHeight: 1.6, marginBottom: 8 }}>
              ¿Asignar el grupo descuento <strong style={{ color: "#7c3aed" }}>{confirmRaiz.codigo}</strong> a los productos cuya referencia empiece por <code style={{ background: "#f5f3ff", padding: "1px 6px", borderRadius: 4 }}>{confirmRaiz.raiz}</code>?
            </p>
            <p style={{ fontSize: 12, color: "#525252", marginBottom: 16 }}>
              Se actualizarán <strong style={{ color: "#16a34a" }}>{confirmRaiz.count}</strong> producto{confirmRaiz.count !== 1 ? "s" : ""}. Si tenían un grupo previo, será sobreescrito.
            </p>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", flexWrap: "wrap" }}>
              <button onClick={() => confirmRaiz.resolve("cancel_all")}
                style={{ padding: "6px 14px", borderRadius: 6, border: "1px solid #fca5a5", background: "#fff", color: "#991b1b", cursor: "pointer", fontSize: 12 }}>
                <BtnContent icon={X} iconColor="#dc2626">Cancelar todo</BtnContent>
              </button>
              <button onClick={() => confirmRaiz.resolve("skip")}
                style={{ padding: "6px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
                <BtnContent icon={X}>Saltar esta</BtnContent>
              </button>
              <button onClick={() => confirmRaiz.resolve("ok")}
                style={{ padding: "6px 18px", borderRadius: 6, border: "1px solid #16a34a", background: "#dcfce7", color: "#14532d", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                <BtnContent icon={Check} iconColor="#14532d">Aceptar y aplicar</BtnContent>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Diálogo Logs */}
      {logsDialog && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 1100, maxHeight: "90vh", display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, borderBottom: "1px solid #e2e8f0", paddingBottom: 12, flexShrink: 0 }}>
              <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}>
                <Icon as={FileText} size={18} color="#1e3a5f" /> Tabla de Log
              </h2>
              <button onClick={() => setLogsDialog(false)} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
                <Icon as={X} size={18} />
              </button>
            </div>

            {/* Toolbar */}
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12, flexShrink: 0, flexWrap: "wrap" }}>
              <span style={{ fontSize: 12, color: "#64748b" }}>Total: <strong style={{ color: "#1e3a5f" }}>{logs.length}</strong> registros</span>
              <button onClick={cargarLogs}
                style={{ padding: "5px 12px", fontSize: 12, borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", cursor: "pointer" }}>
                <BtnContent icon={RefreshCw}>Refrescar</BtnContent>
              </button>
              <div style={{ flex: 1 }} />
              <label style={{ fontSize: 12, color: "#64748b", display: "inline-flex", alignItems: "center", gap: 6 }}>
                Borrar logs anteriores a:
                <input type="date" value={fechaCorte} onChange={e => setFechaCorte(e.target.value)}
                  style={{ padding: "4px 8px", fontSize: 12, borderRadius: 4, border: "1px solid #cbd5e1", background: "#f8fafc" }} />
              </label>
              <button onClick={() => setConfirmBorrarLogs(true)} disabled={!fechaCorte || borrando}
                style={{ padding: "6px 14px", fontSize: 12, borderRadius: 6, border: "1px solid #fca5a5", background: fechaCorte ? "#fff" : "#fafafa", color: fechaCorte ? "#dc2626" : "#cbd5e1", cursor: fechaCorte ? "pointer" : "default", fontWeight: 500 }}>
                <BtnContent icon={Trash2} iconColor={fechaCorte ? "#dc2626" : "#cbd5e1"}>Borrar anteriores</BtnContent>
              </button>
            </div>

            {errorLogs && (
              <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#991b1b", marginBottom: 10, flexShrink: 0 }}>
                {errorLogs}
              </div>
            )}

            {/* Tabla logs */}
            <div style={{ flex: 1, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, marginBottom: 12, minHeight: 250 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 1, borderBottom: "1px solid #e5e5e5" }}>
                  <tr>
                    <th style={{ padding: "8px 10px", textAlign: "right", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>ID</th>
                    <th style={{ padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>Fecha</th>
                    <th style={{ padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>Usuario</th>
                    <th style={{ padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Texto</th>
                  </tr>
                </thead>
                <tbody>
                  {cargandoLogs && (
                    <tr><td colSpan={4} style={{ padding: "30px", textAlign: "center", color: "#64748b" }}>Cargando...</td></tr>
                  )}
                  {!cargandoLogs && logs.length === 0 && !errorLogs && (
                    <tr><td colSpan={4} style={{ padding: "30px", textAlign: "center", color: "#94a3b8" }}>No hay registros de log.</td></tr>
                  )}
                  {!cargandoLogs && logs.map((l, i) => (
                    <tr key={l.id || i} style={{ background: i % 2 === 0 ? "#fff" : "#f8fafc", borderBottom: "1px solid #f1f5f9" }}>
                      <td style={{ padding: "5px 10px", textAlign: "right", color: "#64748b", fontFamily: "monospace", whiteSpace: "nowrap" }}>{l.id}</td>
                      <td style={{ padding: "5px 10px", whiteSpace: "nowrap", color: "#475569" }}>{formatearFechaLog(l.fecha)}</td>
                      <td style={{ padding: "5px 10px", color: "#1e3a5f", whiteSpace: "nowrap" }}>{l.usuario || "—"}</td>
                      <td style={{ padding: "5px 10px", color: "#475569" }}>{l.texto || ""}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div style={{ display: "flex", justifyContent: "flex-end", paddingTop: 10, borderTop: "1px solid #e2e8f0", flexShrink: 0 }}>
              <button onClick={() => setLogsDialog(false)}
                style={{ padding: "7px 18px", borderRadius: 6, border: "none", background: "#2563eb", color: "#fff", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                <BtnContent icon={X} iconColor="#fff">Cerrar</BtnContent>
              </button>
            </div>

            {/* Confirmación borrado de logs */}
            {confirmBorrarLogs && (
              <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }}>
                <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", maxWidth: 440, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                    <Icon as={Trash2} size={22} color="#dc2626" />
                    <h3 style={{ fontSize: 15, fontWeight: 700, margin: 0 }}>Borrar logs antiguos</h3>
                  </div>
                  <p style={{ fontSize: 13, color: "#475569", lineHeight: 1.6, margin: "0 0 16px" }}>
                    Se borrarán todos los registros de la tabla Log con fecha <strong>anterior</strong> al {new Date(fechaCorte).toLocaleDateString("es-ES")}.
                    <br /><span style={{ color: "#dc2626", fontSize: 12 }}>Esta acción no se puede deshacer.</span>
                  </p>
                  <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                    <button onClick={() => setConfirmBorrarLogs(false)} disabled={borrando}
                      style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>
                      <BtnContent icon={X}>Cancelar</BtnContent>
                    </button>
                    <button onClick={borrarLogsAnteriores} disabled={borrando}
                      style={{ padding: "7px 16px", borderRadius: 6, border: "none", background: "#dc2626", color: "#fff", cursor: borrando ? "default" : "pointer", fontSize: 13, fontWeight: 600 }}>
                      <BtnContent icon={borrando ? RefreshCw : Check} iconColor="#fff">{borrando ? "Borrando..." : "Sí, borrar"}</BtnContent>
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Sección Configuraciones Varias dentro de Opciones ──
function ConfiguracionesVariasSection({ config, setConfig, setStatus }) {
  const [draft, setDraft] = useState(config);
  const fileInputRef = useRef(null);

  const hayCambios = JSON.stringify(draft) !== JSON.stringify(config);

  const guardarCambios = () => {
    setConfig(draft);
    // Además, descargar fichero de configuración JSON
    try {
      const data = JSON.stringify(draft, null, 2);
      const blob = new Blob([data], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "CdM_configuracion.json";
      a.click();
      URL.revokeObjectURL(url);
      setStatus("Configuración guardada y descargada en CdM_configuracion.json", "success");
    } catch (e) {
      setStatus("Configuración aplicada (no se pudo descargar el fichero: " + e.message + ")", "warning");
    }
  };

  const leerFichero = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const obj = JSON.parse(ev.target.result);
        const nuevo = {
          sqDestinatario: obj.sqDestinatario || "",
          sqCC: obj.sqCC || "",
        };
        setDraft(nuevo);
        setConfig(nuevo);
        setStatus("Configuración cargada desde fichero", "success");
      } catch (err) {
        setStatus("El fichero no es una configuración válida (JSON)", "error");
      }
    };
    reader.readAsText(file);
    e.target.value = ""; // permite recargar el mismo fichero
  };

  const campo = (label, key, placeholder) => (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: "block", fontSize: 12, color: "#525252", fontWeight: 500, marginBottom: 4 }}>{label}</label>
      <input
        type="text"
        value={draft[key] || ""}
        onChange={e => setDraft(d => ({ ...d, [key]: e.target.value }))}
        placeholder={placeholder}
        style={{ width: "100%", maxWidth: 480, padding: "7px 10px", border: "1px solid #cbd5e1", borderRadius: 6, fontSize: 13, boxSizing: "border-box" }}
      />
    </div>
  );

  return (
    <div>
      {/* Cabecera con botones */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16, borderBottom: "2px solid #2563eb", paddingBottom: 8 }}>
        <h2 style={{ fontSize: 18, fontWeight: 700, color: "#171717", margin: 0 }}>Configuraciones Varias</h2>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button onClick={() => fileInputRef.current && fileInputRef.current.click()}
            style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
            <BtnContent icon={Upload} iconColor="#475569">Leer fichero configuración</BtnContent>
          </button>
          <button onClick={guardarCambios}
            style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #16a34a", background: hayCambios ? "#dcfce7" : "#fff", color: "#14532d", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={Save} iconColor="#16a34a">Guardar cambios</BtnContent>
          </button>
          <input ref={fileInputRef} type="file" accept=".json,application/json" onChange={leerFichero} style={{ display: "none" }} />
        </div>
      </div>

      {/* Apartado Crear SimpleQuote */}
      <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, padding: "16px 20px", maxWidth: 560 }}>
        <h3 style={{ fontSize: 14, fontWeight: 700, color: "#1e3a5f", marginTop: 0, marginBottom: 12, display: "flex", alignItems: "center", gap: 8 }}>
          <Icon as={FileSpreadsheet} size={16} color="#0369a1" /> Crear SimpleQuote
        </h3>
        {campo("Destinatario del correo", "sqDestinatario", "simplequote-rfq.industry@siemens.com")}
        {campo("CC (copia, opcional)", "sqCC", "correo@ejemplo.com")}
        <p style={{ fontSize: 11, color: "#94a3b8", margin: "8px 0 0 0" }}>
          Estos valores se usan al ejecutar Presupuesto → Crear SimpleQuote.
        </p>
      </div>
    </div>
  );
}

function OpcionesScreen({ estilos, setEstilos, configVarias, setConfigVarias, onVolver, setStatus, statusMessage }) {
  const [seccion, setSeccion] = useState("estilos");
  const [draft, setDraft] = useState(estilos);
  const [confirmReset, setConfirmReset] = useState(false);

  const aplicar = () => {
    setEstilos(draft);
  };

  const resetear = () => {
    setConfirmReset(true);
  };

  const confirmarReset = () => {
    const def = JSON.parse(JSON.stringify(ESTILOS_DEFAULT));
    setDraft(def);
    setEstilos(def);
    setConfirmReset(false);
  };

  const hayCambios = JSON.stringify(draft) !== JSON.stringify(estilos);

  const actualizar = (nat, campo, valor) => {
    setDraft(d => ({ ...d, [nat]: { ...d[nat], [campo]: valor } }));
  };

  return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", fontSize: 13, color: "#1e293b", height: "100vh", display: "flex", flexDirection: "column", background: "#f8fafc" }}>
      <div style={{ background: "#f5f5f5", color: "#171717", padding: "8px 16px", display: "flex", alignItems: "center", gap: 12, flexShrink: 0, borderBottom: "1px solid #e5e5e5" }}>
        <button onClick={onVolver} style={{ background: "#fff", border: "1px solid #d4d4d4", color: "#171717", borderRadius: 6, padding: "4px 12px", cursor: "pointer", fontSize: 12 }}>
          <BtnContent icon={ArrowLeft}>Volver</BtnContent>
        </button>
        <span style={{ fontWeight: 700, fontSize: 15, display: "inline-flex", alignItems: "center", gap: 8 }}>
          <Icon as={Settings} size={18} color="#171717" /> Opciones
        </span>

      </div>

      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
        {/* Menú izquierda */}
        <div style={{ width: 220, flexShrink: 0, background: "#fff", borderRight: "1px solid #e2e8f0", padding: "8px 0" }}>
          <div style={{ padding: "10px 12px 6px", fontSize: 11, fontWeight: 700, color: "#94a3b8", letterSpacing: 1, textTransform: "uppercase" }}>Configuración</div>
          {OPCIONES_MENU.map(opt => (
            <div key={opt.id}
              title={opt.tooltip || undefined}
              onClick={() => setSeccion(opt.id)}
              style={{ padding: "8px 14px", cursor: "pointer", fontSize: 13, display: "flex", alignItems: "center", gap: 8,
                background: seccion === opt.id ? "#eff6ff" : "transparent",
                color: seccion === opt.id ? "#2563eb" : "#475569",
                fontWeight: seccion === opt.id ? 600 : 400,
                borderLeft: seccion === opt.id ? "3px solid #2563eb" : "3px solid transparent",
                transition: "all 0.15s" }}
              onMouseEnter={e => { if (seccion !== opt.id) e.currentTarget.style.background = "#f8fafc"; }}
              onMouseLeave={e => { if (seccion !== opt.id) e.currentTarget.style.background = "transparent"; }}>
              <Icon as={opt.icon} size={14} color={seccion === opt.id ? "#2563eb" : "#64748b"} />
              <span>{opt.label}</span>
            </div>
          ))}
        </div>

        {/* Contenido derecha */}
        <div style={{ flex: 1, overflowY: "auto", padding: "28px 40px" }}>
          {seccion === "tarifas" && <TarifasSection setStatus={setStatus} />}
          {seccion === "mantenimiento" && <MantenimientoSection setStatus={setStatus} />}
          {seccion === "usuarios" && <UsuariosSection setStatus={setStatus} />}
          {seccion === "varias" && <ConfiguracionesVariasSection config={configVarias} setConfig={setConfigVarias} setStatus={setStatus} />}
          {seccion === "estilos" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8, borderBottom: "2px solid #2563eb", paddingBottom: 8 }}>
                <h2 style={{ fontSize: 18, fontWeight: 700, color: "#171717", margin: 0 }}>Configurar Estilos</h2>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <button onClick={() => {
                      // Exportar estilos a fichero JSON
                      const data = JSON.stringify(draft, null, 2);
                      const blob = new Blob([data], { type: "application/json" });
                      const url = URL.createObjectURL(blob);
                      const a = document.createElement("a");
                      const fecha = new Date().toISOString().slice(0, 10);
                      a.href = url;
                      a.download = `estilos-presupuestos-${fecha}.json`;
                      document.body.appendChild(a);
                      a.click();
                      document.body.removeChild(a);
                      URL.revokeObjectURL(url);
                      setStatus && setStatus("Estilos exportados a fichero JSON", "success");
                    }}
                    title="Descargar los estilos actuales en un fichero JSON"
                    style={{ padding: "5px 12px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
                    <BtnContent icon={Download} iconColor="#475569">Exportar JSON</BtnContent>
                  </button>
                  <button onClick={() => document.getElementById("import-estilos-input")?.click()}
                    title="Cargar estilos desde un fichero JSON exportado previamente"
                    style={{ padding: "5px 12px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
                    <BtnContent icon={FileUp} iconColor="#475569">Importar JSON</BtnContent>
                  </button>
                  <input id="import-estilos-input" type="file" accept=".json,application/json" style={{ display: "none" }}
                    onChange={e => {
                      const file = e.target.files?.[0];
                      e.target.value = "";
                      if (!file) return;
                      const reader = new FileReader();
                      reader.onload = ev => {
                        try {
                          const parsed = JSON.parse(ev.target.result);
                          if (!parsed || typeof parsed !== "object") throw new Error("Formato inválido");
                          // Fusionar con defaults para no perder ninguna clave
                          const fusionado = {};
                          Object.keys(ESTILOS_DEFAULT).forEach(k => {
                            fusionado[k] = { ...ESTILOS_DEFAULT[k], ...(parsed[k] || {}) };
                          });
                          setDraft(fusionado);
                          setStatus && setStatus(`Estilos importados desde "${file.name}". Pulsa "Aplicar y guardar" para aplicar los cambios.`, "success");
                        } catch (err) {
                          setStatus && setStatus("Error importando estilos: " + err.message, "error");
                        }
                      };
                      reader.readAsText(file);
                    }} />
                  <button onClick={resetear}
                    style={{ padding: "5px 12px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
                    <BtnContent icon={RefreshCw}>Restaurar por defecto</BtnContent>
                  </button>
                  <button onClick={aplicar} disabled={!hayCambios}
                    style={{ padding: "5px 14px", borderRadius: 6, border: hayCambios ? "1px solid #16a34a" : "1px solid #d4d4d4", background: hayCambios ? "#dcfce7" : "#f5f5f5", color: hayCambios ? "#14532d" : "#737373", cursor: hayCambios ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
                    <BtnContent icon={Save} iconColor={hayCambios ? "#14532d" : "#737373"}>Aplicar y guardar</BtnContent>
                  </button>
                </div>
              </div>
              <p style={{ color: "#475569", lineHeight: 1.6, marginBottom: 20, fontSize: 12 }}>
                Personaliza la apariencia de cada tipo de línea (naturaleza). Los cambios se guardan en el navegador (localStorage) y se aplican al activar "Aplicar Estructura" en el presupuesto.
              </p>

              <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 12 }}>
                {NATURALEZAS_CON_ESTILO.map(nat => {
                  const cfg = draft[nat.key] || ESTILOS_DEFAULT[nat.key];
                  return (
                    <div key={nat.key} style={{ border: "1px solid #e2e8f0", borderRadius: 8, padding: "12px 16px", background: "#fff" }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                        <h3 style={{ fontSize: 13, fontWeight: 600, color: "#1e3a5f", margin: 0 }}>{nat.label}</h3>
                        {/* Vista previa */}
                        <div style={{ padding: "4px 14px", borderRadius: 4, background: cfg.bg || "transparent", color: cfg.color, fontFamily: cfg.fontFamily, fontWeight: cfg.fontWeight, fontSize: cfg.fontSize, border: !cfg.bg ? "1px dashed #cbd5e1" : "none" }}>
                          Vista previa — {nat.key}
                        </div>
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: "1.4fr 0.8fr 1fr 1fr 0.9fr", gap: 10 }}>
                        {/* Fuente */}
                        <div>
                          <label style={{ fontSize: 11, color: "#64748b", display: "block", marginBottom: 3 }}>Tipo de letra</label>
                          <select value={cfg.fontFamily} onChange={e => actualizar(nat.key, "fontFamily", e.target.value)}
                            style={{ width: "100%", padding: "4px 6px", fontSize: 12, borderRadius: 4, border: "1px solid #cbd5e1", background: "#f8fafc" }}>
                            {FUENTES_DISPONIBLES.map(f => <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>)}
                          </select>
                        </div>
                        {/* Tamaño */}
                        <div>
                          <label style={{ fontSize: 11, color: "#64748b", display: "block", marginBottom: 3 }}>Tamaño</label>
                          <input type="number" min="8" max="32" value={cfg.fontSize} onChange={e => actualizar(nat.key, "fontSize", parseInt(e.target.value) || 12)}
                            style={{ width: "100%", padding: "4px 6px", fontSize: 12, borderRadius: 4, border: "1px solid #cbd5e1", background: "#f8fafc", textAlign: "right" }} />
                        </div>
                        {/* Peso */}
                        <div>
                          <label style={{ fontSize: 11, color: "#64748b", display: "block", marginBottom: 3 }}>Negrita</label>
                          <select value={cfg.fontWeight} onChange={e => actualizar(nat.key, "fontWeight", parseInt(e.target.value))}
                            style={{ width: "100%", padding: "4px 6px", fontSize: 12, borderRadius: 4, border: "1px solid #cbd5e1", background: "#f8fafc" }}>
                            <option value={400}>Normal</option>
                            <option value={500}>Semi-negrita</option>
                            <option value={600}>Negrita</option>
                            <option value={700}>Extra-negrita</option>
                          </select>
                        </div>
                        {/* Color fondo */}
                        <div>
                          <label style={{ fontSize: 11, color: "#64748b", display: "block", marginBottom: 3 }}>Color fondo</label>
                          <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                            <input type="color" value={cfg.bg} onChange={e => actualizar(nat.key, "bg", e.target.value)}
                              style={{ width: 36, height: 26, padding: 0, border: "1px solid #cbd5e1", borderRadius: 4, cursor: "pointer", background: "#f8fafc" }} />
                            <input type="text" value={cfg.bg} onChange={e => actualizar(nat.key, "bg", e.target.value)}
                              style={{ flex: 1, padding: "4px 6px", fontSize: 11, borderRadius: 4, border: "1px solid #cbd5e1", background: "#f8fafc", fontFamily: "monospace" }} />
                          </div>
                        </div>
                        {/* Color texto */}
                        <div>
                          <label style={{ fontSize: 11, color: "#64748b", display: "block", marginBottom: 3 }}>Color texto</label>
                          <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                            <input type="color" value={cfg.color} onChange={e => actualizar(nat.key, "color", e.target.value)}
                              style={{ width: 36, height: 26, padding: 0, border: "1px solid #cbd5e1", borderRadius: 4, cursor: "pointer", background: "#f8fafc" }} />
                            <input type="text" value={cfg.color} onChange={e => actualizar(nat.key, "color", e.target.value)}
                              style={{ flex: 1, padding: "4px 6px", fontSize: 11, borderRadius: 4, border: "1px solid #cbd5e1", background: "#f8fafc", fontFamily: "monospace" }} />
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      </div>

      {confirmReset && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", minWidth: 380, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
              <Icon as={RefreshCw} size={22} color="#d97706" />
              <h2 style={{ fontSize: 15, fontWeight: 600, margin: 0 }}>Restaurar valores por defecto</h2>
            </div>
            <p style={{ fontSize: 13, color: "#475569", margin: "0 0 20px", lineHeight: 1.6 }}>
              Se restaurarán todos los estilos a sus valores originales y se borrarán los cambios guardados.
              <br /><span style={{ color: "#94a3b8", fontSize: 12 }}>Esta acción se aplica inmediatamente y no se puede deshacer.</span>
            </p>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setConfirmReset(false)} style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>
                <BtnContent icon={X}>Cancelar</BtnContent>
              </button>
              <button onClick={confirmarReset}
                style={{ padding: "7px 16px", borderRadius: 6, border: "none", background: "#d97706", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 500 }}>
                <BtnContent icon={RefreshCw} iconColor="#fff">Sí, restaurar</BtnContent>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Barra de estado */}
      {statusMessage && (
        <div style={{
          background: statusMessage.type === "error" ? "#7f1d1d" : statusMessage.type === "success" ? "#14532d" : statusMessage.type === "working" ? "#1e40af" : "#0f172a",
          color: "#fff", padding: "5px 16px", fontSize: 12, flexShrink: 0,
          display: "flex", alignItems: "center", gap: 8, borderTop: "1px solid rgba(255,255,255,0.1)",
          transition: "background 0.2s"
        }}>
          <Icon as={
            statusMessage.type === "error" ? X
            : statusMessage.type === "success" ? Check
            : statusMessage.type === "working" ? RefreshCw
            : Check
          } size={13} color="#fff" />
          <span>{statusMessage.text}</span>
          {statusMessage.timestamp && (
            <span style={{ marginLeft: "auto", color: "rgba(255,255,255,0.7)", fontSize: 11 }}>
              {new Date(statusMessage.timestamp).toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
            </span>
          )}
        </div>
      )}
    </div>
  );
}

// ── Tree Node component (PrimeReact Tree style) ──
function TreeNode({ node, activeId, onSelect, depth = 0 }) {
  const hasChildren = node.children && node.children.length > 0;
  const isChildActive = hasChildren && node.children.some(c => c.id === activeId);
  const [expanded, setExpanded] = useState(hasChildren);
  const isActive = activeId === node.id;

  return (
    <div>
      <div
        onClick={() => { if (hasChildren) setExpanded(e => !e); else onSelect(node.id); }}
        onMouseEnter={e => { if (!isActive) e.currentTarget.style.background = "#f1f5f9"; }}
        onMouseLeave={e => { if (!isActive) e.currentTarget.style.background = isChildActive ? "#f8fafc" : "transparent"; }}
        style={{ display: "flex", alignItems: "center", padding: `5px 8px 5px ${8 + depth * 18}px`, cursor: "pointer", borderRadius: 4, margin: "1px 4px", background: isActive ? "#eff6ff" : isChildActive ? "#f8fafc" : "transparent", transition: "background 0.12s" }}>
        <span style={{ width: 18, height: 18, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, color: "#94a3b8", fontSize: 9, transform: hasChildren ? (expanded ? "rotate(90deg)" : "rotate(0deg)") : "none", opacity: hasChildren ? 1 : 0, transition: "transform 0.15s" }}>▶</span>
        <span style={{ marginRight: 6, flexShrink: 0, display: "inline-flex" }}><Icon as={node.icon} size={14} color={isActive ? "#2563eb" : "#64748b"} /></span>
        <span style={{ fontSize: 12, flex: 1, lineHeight: 1.4, color: isActive ? "#2563eb" : "#374151", fontWeight: isActive ? 600 : isChildActive ? 500 : 400 }}>{node.label}</span>
        {isActive && <span style={{ width: 6, height: 6, borderRadius: "50%", background: "#2563eb", flexShrink: 0 }} />}
      </div>
      {hasChildren && expanded && (
        <div style={{ position: "relative" }}>
          <div style={{ position: "absolute", left: 8 + depth * 18 + 17, top: 0, bottom: 4, width: 1, background: "#e2e8f0" }} />
          {node.children.map(child => (
            <div key={child.id} style={{ position: "relative" }}>
              <div style={{ position: "absolute", left: 8 + depth * 18 + 17, top: "50%", width: 12, height: 1, background: "#e2e8f0" }} />
              <div
                onClick={() => onSelect(child.id)}
                onMouseEnter={e => { if (activeId !== child.id) e.currentTarget.style.background = "#f1f5f9"; }}
                onMouseLeave={e => { if (activeId !== child.id) e.currentTarget.style.background = "transparent"; }}
                style={{ display: "flex", alignItems: "center", padding: `5px 8px 5px ${8 + (depth + 1) * 18}px`, cursor: "pointer", borderRadius: 4, margin: "1px 4px", background: activeId === child.id ? "#eff6ff" : "transparent", transition: "background 0.12s" }}>
                <span style={{ width: 18, flexShrink: 0 }} />
                <span style={{ marginRight: 6, flexShrink: 0, display: "inline-flex" }}><Icon as={child.icon} size={14} color={activeId === child.id ? "#2563eb" : "#64748b"} /></span>
                <span style={{ fontSize: 12, flex: 1, color: activeId === child.id ? "#2563eb" : "#374151", fontWeight: activeId === child.id ? 600 : 400 }}>{child.label}</span>
                {activeId === child.id && <span style={{ width: 6, height: 6, borderRadius: "50%", background: "#2563eb", flexShrink: 0 }} />}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

const MENU_STRUCTURE = [
  { id: "presupuesto", icon: FileText, label: "Presupuesto", tooltip: "Guardar, leer, importar/exportar y operar con presupuestos", items: [
    { label: "Guardar Presupuesto", action: "GuardarPresupuesto", icon: Save, tooltip: "Guarda el presupuesto actual (cabecera y líneas) en la base de datos" },
    { label: "Comprobar Presupuesto", action: "ComprobarPresupuesto", icon: ClipboardCheck, tooltip: "Verifica referencias, precios y datos de las líneas contra la BD" },
    { label: "Leer Presupuestos", action: "LeerPresupuestos", icon: FolderOpen, tooltip: "Abre la lista de presupuestos guardados para cargar uno" },
    { label: "Importar", action: "Importar", icon: Download, tooltip: "Carga las líneas del presupuesto desde un fichero Excel" },
    { label: "Exportar", action: "Exportar", icon: FileUp, tooltip: "Vuelca el presupuesto actual a un fichero Excel con estilos" },
    { label: "Imprimir", action: "Imprimir", icon: Printer, tooltip: "Genera la versión imprimible del presupuesto" },
    { label: "Formato Simple Quote", action: "FormatoSimpleQuote", icon: Printer, tooltip: "Prepara el presupuesto con el formato de tabla SimpleQuote" },
    { label: "Crear SimpleQuote", action: "CrearSimpleQuote", icon: FileSpreadsheet, tooltip: "Envía la oferta a SimpleQuote (Siemens) vía la API local" },
    { label: "Hacer Damex E", action: "HacerDamexE", icon: FileSpreadsheet, tooltip: "Crea el pedido Damex E en Siemens con los datos del cliente final" },
    { label: "Resumen", action: "Resumen", icon: BarChart3, tooltip: "Muestra el resumen por familia/subfamilia con importes y márgenes" },
    { label: "Aplicar Estructura", action: "AplicarEstructura", icon: Palette, tooltip: "Aplica colores y estilos de títulos, subtotales y comentarios" },
    { label: "Borrar Presupuesto actual", action: "BorrarPresupuestoActual", icon: X, tooltip: "Vacía el presupuesto en pantalla (no borra el guardado en BD)" },
    { label: "Comparar Presupuestos", action: "CompararPresupuestos", icon: Scale, tooltip: "Compara dos presupuestos línea a línea" },
  ]},
  { id: "celdas", icon: Grid3x3, label: "Celdas", tooltip: "Operaciones sobre celdas y filas del grid", items: [
    { label: "Comprobar Celda", action: "ComprobarCelda", icon: Search, tooltip: "Muestra nº de caracteres, color de fondo y color de tinta de la celda seleccionada" },
    { label: "Borrar filas vacías", action: "BorrarFilas", icon: Trash2, tooltip: "Elimina las filas completamente vacías del grid" },
    { label: "Borrar filas con 0", action: "BorrarFilasCero", icon: Trash2, tooltip: "Borra las filas que tienen 0 en la columna seleccionada (rectángulo azul, una sola columna)" },
    { label: "Seleccionar Celdas", action: "SeleccionarCeldas", icon: Square, tooltip: "Selecciona un rango de celdas para operar sobre ellas" },
    { label: "Convertir selección en comentario", action: "ConvertirEnComentario", icon: MessageSquare, tooltip: "Convierte las filas seleccionadas en líneas de comentario (CM)" },
    { label: "Juntar celdas en una", action: "JuntarCeldas", icon: Link2, tooltip: "Une el contenido de varias celdas en una sola" },
  ]},
  { id: "elementos", icon: Layers, label: "Elementos", tooltip: "Guardar y reutilizar conjuntos de líneas (elementos)", items: [
    { label: "Guardar Elemento", action: "GuardarElemento", icon: Save, tooltip: "Guarda las líneas seleccionadas como un elemento reutilizable" },
    { label: "Leer Elemento", action: "LeerElemento", icon: Download, tooltip: "Inserta en el presupuesto las líneas de un elemento guardado" },
  ]},
  { id: "productos", icon: Package, label: "Productos", tooltip: "Buscar, guardar y procesar productos del catálogo", items: [
    { label: "Guardar Producto", action: "GuardarProducto", icon: Save, tooltip: "Crea o actualiza un producto en el catálogo de la BD" },
    { label: "Leer Producto", action: "LeerProducto", icon: Download, tooltip: "Busca un producto del catálogo y vuelca sus datos" },
    { label: "Buscar datos por Referencia", action: "BuscarDatosProductos", icon: Search, tooltip: "Rellena nombre, PVP y datos de cada línea por su referencia" },
    { label: "Juntar productos duplicados", action: "JuntarDuplicados", icon: Layers, tooltip: "Agrupa en una sola línea los productos repetidos sumando cantidades" },
    { label: "---" },
    { label: "Quitar caracteres Referencia", action: "QuitarEspacios", icon: Scissors, tooltip: "Elimina espacios y caracteres sobrantes de las referencias" },
    { label: "Quitar Saltos de línea", action: "QuitarCR", icon: CornerDownLeft, tooltip: "Elimina los saltos de línea del texto de las celdas" },
    { label: "Leer Precios de PMD", action: "LeerPreciosPMD", icon: Download, tooltip: "Consulta precios actualizados en PMD vía la API local" },
    { label: "Buscar Referencia SIEMENS", action: "BuscarReferencia", icon: Search, tooltip: "Detecta referencias Siemens (MLFB) en el texto y las extrae a celdas" },
    { label: "Buscar equivalencia Competencia", action: "BuscarEquivalencia", icon: Repeat, tooltip: "Busca el producto de competencia equivalente a la referencia" },
    { label: "Calcular PVP a partir de GA", action: "CalcularPVP", icon: Calculator, tooltip: "Calcula el PVP a partir del coste GA y el margen" },
    { label: "Asistente Referencias", action: "Asistente", icon: Bot, tooltip: "Asistente para completar y corregir referencias" },
  ]},
  { id: "clientes", icon: Users, label: "Clientes", tooltip: "Gestionar clientes y contactos", items: [
    { label: "Gestionar Clientes", action: "GestionarClientes", icon: Users, tooltip: "Abre la tabla de clientes para crear, editar o borrar" },
    { label: "Gestionar Contactos", action: "GestionarContactos", icon: User, tooltip: "Abre la tabla de contactos para crear, editar o borrar" },
  ]},
  { id: "descuentos", icon: Percent, label: "Descuentos", tooltip: "Aplicar y calcular descuentos y precios", items: [
    { label: "Aplicar descuentos", action: "AplicarDescuentos", icon: DollarSign, tooltip: "Aplica un porcentaje de descuento a las celdas seleccionadas" },
    { label: "Calcular Descuento", action: "CalcularDescuento", icon: Calculator, tooltip: "Calcular el descuento para obtener precio neto seleccionado en las filas seleccionadas" },
    { label: "Fijar el precio total", action: "FijarPrecioTotal", icon: Target, tooltip: "Ajusta los descuentos para alcanzar un precio total objetivo" },
    { label: "Fijar precio de celdas", action: "FijarPrecioCeldas", icon: Hash, tooltip: "Fija el precio neto de las celdas seleccionadas" },
    { label: "---" },
    { label: "Gestionar Estrategias Descuento", action: "GestionarEstrategias", icon: Percent, tooltip: "Consulta, crea, edita y aplica estrategias de descuento por grupo" },
  ]},
  { id: "otros", icon: MoreHorizontal, label: "Otros", tooltip: "Ayuda, opciones y bases de datos auxiliares", items: [
    { label: "Ayuda", action: "Ayuda", icon: HelpCircle, tooltip: "Muestra la ayuda con atajos y acciones del grid" },
    { label: "Opciones", action: "Opciones", icon: Settings, tooltip: "Estilos, tarifas, mantenimiento de BD, usuarios y configuración" },
    { label: "---" },
    { label: "Gestionar BD Competencia", action: "GestionarProductosCompetencia", icon: Database, tooltip: "Mantiene la base de datos de productos de la competencia" },
  ]},
];

const REPRESENTACION_OPCIONES = [
  { id: 1, representacion: "TP",       descripcion: "Total Posición - En la oferta aparecera un total con 1 unidad" },
  { id: 2, representacion: "Opcional", descripcion: "Posición Opcional - No sumada en el total del presupuesto." },
  { id: 3, representacion: "Conf",     descripcion: "Confirmar Posición. Sale resaltada la linea en la oferta." },
];

const NATURALEZA_OPCIONES = [
  { id: 1,  naturaleza: "NP",    descripcion: "Numero Presupuesto" },
  { id: 2,  naturaleza: "R",     descripcion: "Numero Revisión" },
  { id: 3,  naturaleza: "T",     descripcion: "Titulo del Presupuesto" },
  { id: 4,  naturaleza: "T1",    descripcion: "Titulo Apartado Nivel 1" },
  { id: 5,  naturaleza: "T2",    descripcion: "Titulo Apartado Nivel 2" },
  { id: 6,  naturaleza: "T3",    descripcion: "Titulo Apartado Nivel 3" },
  { id: 7,  naturaleza: "T4",    descripcion: "Titulo Apartado Nivel 4" },
  { id: 8,  naturaleza: "S1",    descripcion: "SubTotal Apartado Nivel 1" },
  { id: 9,  naturaleza: "S2",    descripcion: "SubTotal Apartado Nivel 2" },
  { id: 10, naturaleza: "S3",    descripcion: "SubTotal Apartado Nivel 3" },
  { id: 11, naturaleza: "S4",    descripcion: "SubTotal Apartado Nivel 4" },
  { id: 12, naturaleza: "TT",    descripcion: "Total del Presupuesto" },
  { id: 13, naturaleza: "CM",    descripcion: "Linea de Comentario" },
  { id: 14, naturaleza: "CL",    descripcion: "Cliente" },
  { id: 15, naturaleza: "PD",    descripcion: "Producto independiente" },
  { id: 16, naturaleza: "PE",    descripcion: "Producto pertenece a Elemento" },
  { id: 17, naturaleza: "E",     descripcion: "Elemento" },
  { id: 18, naturaleza: "VERDE", descripcion: "Resaltamos linea en Verde" },
  { id: 19, naturaleza: "GRIS",  descripcion: "Resaltamos linea en Gris" },
];

const COLUMNS = [
  { key: "representacion",      label: "Repres.",             tooltip: "Representación", width: 50,  type: "select", opciones: "representacion", align: "center" },
  { key: "naturaleza",          label: "Nat.",                tooltip: "Naturaleza", width: 50,  type: "select", opciones: "naturaleza", align: "center" },
  { key: "posicion",            label: "Apartado",            width: 70 },
  { key: "cantidad",            label: "Cant.",               tooltip: "Cantidad", width: 38, type: "number", align: "center" },
  { key: "referencia",          label: "Referencia",          width: 145, align: "center" },
  { key: "nombre",              label: "Producto",            tooltip: "Producto", width: 320 },
  { key: "pvp",                 label: ["PVP", "Unitario"],   tooltip: "PVP Unitario", width: 80, type: "number" },
  { key: "dtoaplicado",         label: "Dto. %",              width: 60, type: "number" },
  { key: "precionetounitario",  label: ["Neto", "Unitario"],  tooltip: "Neto Unitario", width: 80, type: "calc" },
  { key: "precionetoposicion",  label: ["Neto", "Posición"],  tooltip: "Neto Posición", width: 90, type: "calc" },
  { key: "descripcion",         label: "Descripción",         width: 260, wrap: true },
  { key: "grupodescuento",      label: ["Grupo", "Descuento"], tooltip: "Grupo Descuento", width: 75 },
  { key: "familia",             label: "Familia",             width: 90 },
  { key: "subfamilia",          label: "SubFamilia",          width: 90 },
  { key: "preciocosteunitario", label: ["Coste Unit.", "(GA)"], tooltip: "Coste Unitario (GA)", width: 80, type: "number" },
  { key: "costeposicion",       label: ["Coste", "Posición"], tooltip: "Coste Posición", width: 80, type: "calc" },
  { key: "margen",              label: "Margen %",            width: 70, type: "calc" },
  { key: "idposicion",          label: ["Id", "Posición"],    tooltip: "Id Posición", width: 60 },
  { key: "imagen",              label: "Imagen",              width: 70 },
  { key: "precionetounitario2", label: ["Neto Unit.", "2"],   tooltip: "Neto Unitario 2", width: 80, type: "number" },
];

const COLORES = ["#2563eb","#16a34a","#dc2626","#d97706","#7c3aed","#0891b2","#be185d","#65a30d","#ea580c","#0284c7"];

const COLS_TITULO   = ["naturaleza", "posicion", "nombre"];
const COLS_SUBTOTAL = ["naturaleza", "posicion", "nombre", "dtoaplicado", "precionetoposicion"];
const COLS_COMENT   = ["naturaleza", "nombre"];
const SUBTOTAL_NIVELES = { "S1": 1, "S2": 2, "S3": 3, "S4": 4 };

// Condiciones particulares de suministro que se vuelcan en la pestaña "Condiciones Comerciales"
// del Excel de Imprimir. Cada entrada: { texto, negrita?, enlace? }.
const CONDICIONES_PARTICULARES_SUMINISTRO = [
  { texto: "1. CONDICIONES PARTICULARES DE SUMINISTRO", negrita: true },
  { texto: "" },
  { texto: "1.1 PRECIOS", negrita: true },
  { texto: "Se entienden netos por material puesto en punto de destino en península, sin descarga, e incluyen, si procede, embalaje terrestre normal." },
  { texto: "Estos precios se mantendrán fijos dentro del plazo de validez de la presente oferta, siempre y cuando su pedido tenga un importe mínimo de 1.000 €." },
  { texto: "Dichos precios no incluyen ningún tipo de montaje ni puesta en marcha del Alcance del Contrato." },
  { texto: "En los precios indicados no está incluido el I.V.A., que será reflejado por separado en factura." },
  { texto: "Una vez formalizado el Contrato o realizado el suministro total o parcial del mismo, no se aceptarán anulaciones de pedido o devoluciones de los productos suministrados." },
  { texto: "" },
  { texto: "1.2 PLAZO DE ENTREGA", negrita: true },
  { texto: "El plazo de entrega mostrado para cada posición del pedido se entiende orientativo, y para el caso de que el Contrato se formalizase en la fecha de emisión de la oferta." },
  { texto: "El plazo se recalculará a partir de la fecha de recepción de su pedido, una vez esté técnica y comercialmente aclarado en todos sus puntos y, en su caso, una vez hecho efectivo el anticipo o el primer hito de pago, y/o aportadas las garantías de pago solicitadas." },
  { texto: "El plazo de entrega final se confirmará con la formalización del Contrato." },
  { texto: "" },
  { texto: "1.3 CONDICIONES DE PAGO", negrita: true },
  { texto: "Hito 1: 25% del precio del Contrato. Pago mediante transferencia al contado." },
  { texto: "Hito 2: 75% a la entrega o puesta a disposición del Alcance del Contrato, admitiéndose suministros, facturaciones y cobros parciales." },
  { texto: "Los pagos correspondientes al hito 2 se realizarán mediante transferencia o confirming a 60 días fecha devengo." },
  { texto: "En el caso de existir unas condiciones de pago acordadas previamente por las partes, éstas prevalecerán." },
  { texto: "" },
  { texto: "1.4 RETRASOS EN LA ENTREGA POR CAUSAS IMPUTABLES AL CLIENTE", negrita: true },
  { texto: "Si no fuese posible el suministro por causas ajenas a Siemens, se llevará a cabo la facturación una vez notificada al Cliente la disponibilidad del Alcance para su envío." },
  { texto: "Se traspasará al Cliente el riesgo de los suministros y Siemens se reserva el derecho a repercutir los gastos que ello origine." },
  { texto: "" },
  { texto: "1.5 PERSONALIDAD JURÍDICA DEL DESTINATARIO FINAL", negrita: true },
  { texto: "Si la persona jurídica que realiza el pedido (Contrato) no es la misma a la que se dirige esta oferta (Alcance), Siemens podrá modificarla, cancelarla y/o exigir garantías de pago adicionales." },
  { texto: "" },
  { texto: "1.6 PUESTA EN MARCHA", negrita: true },
  { texto: "La puesta en marcha no se encuentra incluida en el Alcance del Contrato." },
  { texto: "" },
  { texto: "1.7 POLÍTICA ANTICORRUPCIÓN", negrita: true },
  { texto: "El Cliente declara y garantiza que cumple con todas las normas y leyes aplicables a su negocio." },
  { texto: "No realiza pagos u ofrece nada de valor, directa o indirectamente, a ningún gobierno o empleado administrativo para conseguir o conservar un negocio." },
  { texto: "Respeta los derechos humanos básicos de sus empleados, así como su salud y seguridad." },
  { texto: "Prohíbe el trabajo infantil y fomenta la protección del medio ambiente." },
  { texto: "Las Partes se comprometen a incorporar en el Contrato cláusulas globales que garanticen el cumplimiento normativo." },
  { texto: "Para Siemens, esta condición previa es necesaria para la validez del Contrato." },
  { texto: "" },
  { texto: "2. CONDICIONES CONTRACTUALES", negrita: true },
  { texto: "Para todo lo no previsto expresamente en el presente Contrato, regirán nuestras Condiciones Contractuales, disponibles en los siguientes enlaces:" },
  { texto: "" },
  { texto: "Condiciones Básicas España", enlace: "https://www.siemens.com/sts-base-terms-esp" },
  { texto: "Condiciones Complementarias de Hardware", enlace: "https://www.siemens.com/sts-st-hardware" },
  { texto: "Condiciones Complementarias de Servicios", enlace: "https://www.siemens.com/sts-st-services" },
  { texto: "Condiciones Complementarias de Soluciones", enlace: "https://www.siemens.com/sts-st-solutions" },
  { texto: "Condiciones Complementarias Específicas para Software local del negocio de Infraestructura e Industria", enlace: "https://www.siemens.com/sts-st-software" },
  { texto: "Condiciones Generales de Software y Complementarias para Servicios en la Nube", enlace: "https://www.siemens.com/sts-st-cloud" },
];

const ESTILOS_DEFAULT = {
  T1:    { bg: "#dcfce7", color: "#14532d", fontFamily: "Segoe UI", fontWeight: 700, fontSize: 11 },
  T2:    { bg: "#fee2e2", color: "#7c2d12", fontFamily: "Segoe UI", fontWeight: 600, fontSize: 11 },
  T3:    { bg: "#ffedd5", color: "#7c2d12", fontFamily: "Segoe UI", fontWeight: 600, fontSize: 11 },
  T4:    { bg: "#ffffff", color: "#111827", fontFamily: "Segoe UI", fontWeight: 700, fontSize: 11 },
  S1:    { bg: "#dcfce7", color: "#14532d", fontFamily: "Segoe UI", fontWeight: 700, fontSize: 11 },
  S2:    { bg: "#fee2e2", color: "#7c2d12", fontFamily: "Segoe UI", fontWeight: 600, fontSize: 11 },
  S3:    { bg: "#ffedd5", color: "#7c2d12", fontFamily: "Segoe UI", fontWeight: 600, fontSize: 11 },
  S4:    { bg: "#ffffff", color: "#111827", fontFamily: "Segoe UI", fontWeight: 600, fontSize: 11 },
  TT:    { bg: "#dbeafe", color: "#1e3a8a", fontFamily: "Segoe UI", fontWeight: 700, fontSize: 11 },
  CM:    { bg: "#f2f2f2", color: "#111827", fontFamily: "Segoe UI", fontWeight: 600, fontSize: 11 },
  VERDE: { bg: "#dcfce7", color: "#14532d", fontFamily: "Segoe UI", fontWeight: 400, fontSize: 11 },
  GRIS:  { bg: "#cfcfcf", color: "#111111", fontFamily: "Segoe UI", fontWeight: 700, fontSize: 10 },
  PD:    { bg: "#ffffff", color: "#1e293b", fontFamily: "Segoe UI", fontWeight: 400, fontSize: 11 },
  CONF:  { bg: "#fef9c3", color: "#854d0e", fontFamily: "Segoe UI", fontWeight: 600, fontSize: 11 },
};

// Cargar estilos guardados en localStorage o usar los por defecto
function cargarEstilos() {
  try {
    const guardado = typeof localStorage !== "undefined" ? localStorage.getItem("estilosNaturaleza") : null;
    if (guardado) {
      const parsed = JSON.parse(guardado);
      // Mezcla con defaults para campos que falten
      const result = {};
      Object.keys(ESTILOS_DEFAULT).forEach(k => {
        result[k] = { ...ESTILOS_DEFAULT[k], ...(parsed[k] || {}) };
      });
      return result;
    }
  } catch (e) { /* ignore */ }
  return JSON.parse(JSON.stringify(ESTILOS_DEFAULT));
}

function guardarEstilos(estilos) {
  try {
    if (typeof localStorage !== "undefined") {
      localStorage.setItem("estilosNaturaleza", JSON.stringify(estilos));
    }
  } catch (e) { /* ignore */ }
}

// Cargar presupuesto guardado en localStorage
function cargarPresupuestoLocal() {
  try {
    if (typeof localStorage !== "undefined") {
      const data = localStorage.getItem("presupuestoActual");
      if (data) return JSON.parse(data);
    }
  } catch (e) { /* ignore */ }
  return null;
}

function guardarPresupuestoLocal(presupuesto, rows) {
  try {
    if (typeof localStorage !== "undefined") {
      localStorage.setItem("presupuestoActual", JSON.stringify({ presupuesto, rows, fecha: new Date().toISOString() }));
      return true;
    }
  } catch (e) { /* ignore */ }
  return false;
}

// Borrar presupuesto guardado en localStorage
function borrarPresupuestoLocal() {
  try {
    if (typeof localStorage !== "undefined") {
      localStorage.removeItem("presupuestoActual");
      return true;
    }
  } catch (e) { /* ignore */ }
  return false;
}

// Compara presupuesto actual con el de BD y devuelve lista de diferencias
function compararPresupuestos(actual, rowsActual, bd) {
  if (!bd) return ["No hay versión guardada en BD para comparar"];
  const diffs = [];

  // Comparar cabecera
  const cabBd = bd.cabecera || {};
  const camposCab = [
    { key: "titulo", label: "Título" },
    { key: "revision", label: "Revisión" },
    { key: "idcliente", label: "Cliente" },
  ];
  camposCab.forEach(({ key, label }) => {
    const a = String(actual[key === "idcliente" ? "idcliente" : key] ?? "");
    const b = String(cabBd[key] ?? "");
    if (a !== b) diffs.push(`Cabecera "${label}": "${b}" → "${a}"`);
  });

  // Comparar detalle
  const detalleBd = bd.detalle || [];
  // Conservamos el índice ORIGINAL (posición en el grid del presupuesto actual)
  // para que los mensajes de diferencia indiquen la fila tal como la ve el usuario
  const filasActuales = rowsActual
    .map((r, idxOriginal) => ({ r, idxOriginal }))
    .filter(({ r }) => r.naturaleza || r.referencia || r.nombre || r.descripcion || (r.cantidad && r.cantidad !== 0));

  if (filasActuales.length !== detalleBd.length) {
    diffs.push(`Número de filas: BD tiene ${detalleBd.length}, actual tiene ${filasActuales.length}`);
  }

  const n = Math.max(filasActuales.length, detalleBd.length);
  const campos = ["naturaleza","cantidad","referencia","nombre","pvp","dtoaplicado","descripcion","familia","subfamilia"];
  for (let i = 0; i < n; i++) {
    const aWrap = filasActuales[i];
    const b = detalleBd[i];
    // Si existe la fila en el presupuesto actual, el "número de fila" que mostramos es el del grid actual (1-based)
    // Si no, usamos el índice de comparación como referencia (no hay fila actual a la que apuntar)
    const filaActualNum = aWrap ? (aWrap.idxOriginal + 1) : (i + 1);
    const a = aWrap ? aWrap.r : null;
    if (!a) { diffs.push(`Fila ${filaActualNum}: solo en BD (${b.naturaleza || ""} ${b.referencia || b.nombre || ""})`); continue; }
    if (!b) { diffs.push(`Fila ${filaActualNum}: solo en actual (${a.naturaleza || ""} ${a.referencia || a.nombre || ""})`); continue; }
    campos.forEach(c => {
      const va = String(a[c] ?? "");
      const vb = String(b[c] ?? "");
      if (va !== vb) {
        diffs.push(`Fila ${filaActualNum} [${c}]: "${vb}" → "${va}"`);
      }
    });
  }

  return diffs;
}

// Genera 50 filas vacías para un presupuesto en blanco
function filasVacias() {
  return Array.from({ length: 50 }, (_, i) => ({
    id: i + 1, representacion: "", naturaleza: "", posicion: "", cantidad: 0,
    referencia: "", nombre: "", pvp: 0, dtoaplicado: 0, descripcion: "",
    familia: "", subfamilia: "", preciocosteunitario: 0, idposicion: "",
    imagen: "", precionetounitario2: 0, grupodescuento: "",
  }));
}

// Obtener el siguiente número desde la API (filtrado por año actual)
async function obtenerSiguienteNumero(ano) {
  try {
    const yr = ano || getAnoFiscal();
    const res = await fetch(`${API_URL}/presupuestos/?busqueda=`);
    if (!res.ok) throw new Error("Error " + res.status);
    const lista = await res.json();
    // Filtrar solo los del año indicado
    const delAno = lista.filter(p => Number(p.anopresupuesto) === Number(yr));
    const maxNum = delAno.reduce((m, p) => {
      const n = parseInt(p.numero, 10);
      return !isNaN(n) && n > m ? n : m;
    }, 0);
    return maxNum + 1;
  } catch (e) {
    throw e;
  }
}

// Calcula anchos automáticos (fit-to-content) para columnas cuando estructura activa.
// Columnas que se ajustan al contenido: representacion, naturaleza, cantidad, pvp, dtoaplicado, precionetounitario
function calcAnchosAuto(rows, apartados) {
  // Estimación: ~7px por carácter + padding. Mínimo el width base de la columna.
  const charPx = 7;
  const padding = 18;
  const cols = ["representacion", "naturaleza", "cantidad", "pvp", "dtoaplicado", "precionetounitario", "precionetoposicion", "posicion"];
  // Mínimos por columna para que quepa el contenido formateado (€, %, etc.) más una pequeña holgura
  const minPorKey = {
    representacion: 34, naturaleza: 34, cantidad: 38, posicion: 50,
    pvp: 75, dtoaplicado: 60, precionetounitario: 85, precionetoposicion: 95,
  };
  // Cabeceras (para que el título no se corte)
  const labelLen = {
    representacion: 7, naturaleza: 5, cantidad: 5, posicion: 8,
    pvp: 12, dtoaplicado: 6, precionetounitario: 12, precionetoposicion: 12,
  };
  const anchos = {};
  cols.forEach(key => {
    let maxLen = labelLen[key] || 0;
    rows.forEach(row => {
      let s = "";
      if (key === "posicion") {
        s = String(apartados && apartados[row.id] ? apartados[row.id] : (row.posicion ?? ""));
      } else if (key === "precionetounitario") {
        // Columna calculada: pvp × (1 - dto/100)
        const v = (row.pvp || 0) * (1 - (row.dtoaplicado || 0) / 100);
        s = v ? v.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €" : "";
      } else if (key === "precionetoposicion") {
        // Columna calculada: netoUnit × cantidad
        const netoU = (row.pvp || 0) * (1 - (row.dtoaplicado || 0) / 100);
        const v = netoU * (row.cantidad || 0);
        s = v ? v.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €" : "";
      } else if (key === "pvp") {
        const v = Number(row.pvp) || 0;
        s = v ? v.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €" : "";
      } else if (key === "dtoaplicado") {
        const v = Number(row.dtoaplicado) || 0;
        s = v ? v.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "%" : "";
      } else {
        s = String(row[key] ?? "");
      }
      if (s.length > maxLen) maxLen = s.length;
    });
    anchos[key] = Math.max(maxLen * charPx + padding, minPorKey[key] || 34);
  });
  return anchos;
}

function getEstiloFila(nat, estilos) {
  const cfg = estilos && estilos[nat];
  if (cfg) {
    const isTitle = ["T1","T2","T3","T4"].includes(nat);
    const isSubtotal = ["S1","S2","S3","S4","TT"].includes(nat);
    const isComment = nat === "CM";
    return { bg: cfg.bg, color: cfg.color, fontFamily: cfg.fontFamily, fontWeight: cfg.fontWeight, fontSize: cfg.fontSize, isTitle, isSubtotal, isComment };
  }
  return { bg: null, color: "#1e293b", fontFamily: "Segoe UI", fontWeight: 400, fontSize: 12 };
}

// Versión legacy (sigue funcionando si no se pasa el segundo parámetro)
// (getEstiloFilaLegacy removed - using estilos config now)

function calcNetoUnit(row) { return (row.pvp || 0) * (1 - (row.dtoaplicado || 0) / 100); }
function calcNetoPos(row)  { return calcNetoUnit(row) * (row.cantidad || 0); }
function calcCostePos(row) { return (row.preciocosteunitario || 0) * (row.cantidad || 0); }
function calcMargen(row)   { const n = calcNetoPos(row), c = calcCostePos(row); return n ? ((n - c) / n) * 100 : 0; }
function fmt(n) { return (n || 0).toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: "always" }); }
// Devuelve true si la fecha (ISO o Date) tiene más de un año de antigüedad
function tieneMasDeUnAno(fecha) {
  if (!fecha) return false;
  const d = new Date(fecha);
  if (isNaN(d.getTime())) return false;
  const hace1Ano = new Date();
  hace1Ano.setFullYear(hace1Ano.getFullYear() - 1);
  return d < hace1Ano;
}
// Formatea una fecha ISO a dd/mm/aaaa (o "" si no válida)
function fmtFecha(fecha) {
  if (!fecha) return "";
  const d = new Date(fecha);
  if (isNaN(d.getTime())) return "";
  return d.toLocaleDateString("es-ES", { day: "2-digit", month: "2-digit", year: "numeric" });
}
// Exporta un array de entradas de log a un fichero CSV (columnas: Hora, Tipo, Mensaje).
// Cada entrada puede ser { texto, tipo, hora } o un string suelto.
function exportarLogCSV(log, nombreBase = "log") {
  if (!log || log.length === 0) return;
  const esc = (s) => '"' + String(s ?? "").replace(/"/g, '""') + '"';
  const cab = ["Hora", "Tipo", "Mensaje"].join(";");
  const filas = log.map(l => {
    if (typeof l === "string") return [esc(""), esc("info"), esc(l)].join(";");
    return [esc(l.hora || ""), esc(l.tipo || "info"), esc(l.texto ?? "")].join(";");
  });
  // BOM para que Excel reconozca UTF-8
  const contenido = "\ufeff" + [cab, ...filas].join("\r\n");
  const blob = new Blob([contenido], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const fecha = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `${nombreBase}-${fecha}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
// Formato monetario con €. Si el valor es 0 devuelve "" (campo vacío)
function fmtEur(n) {
  const num = Number(n) || 0;
  if (num === 0) return "";
  return num.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: "always" }) + " €";
}

function calcSubtotal(rows, rowIdx, nivel) {
  const titleNat = "T" + nivel, subNat = "S" + nivel;
  let neto = 0, pvpTotal = 0;
  let i = rowIdx - 1;
  while (i >= 0) {
    const r = rows[i];
    if (r.naturaleza === titleNat) break;
    if (r.naturaleza === subNat) { i = -1; break; }
    if (["PD", "PE", "E"].includes(r.naturaleza)) {
      neto += calcNetoPos(r);
      pvpTotal += (r.pvp || 0) * (r.cantidad || 0);
    }
    i--;
  }
  return { neto, dto: pvpTotal > 0 ? ((pvpTotal - neto) / pvpTotal) * 100 : 0 };
}

// Total de TODO el presupuesto (todas las líneas PD/PE/E)
function calcTotalPresupuesto(rows) {
  let neto = 0, pvpTotal = 0;
  rows.forEach(r => {
    if (["PD", "PE", "E"].includes(r.naturaleza)) {
      neto += calcNetoPos(r);
      pvpTotal += (r.pvp || 0) * (r.cantidad || 0);
    }
  });
  return { neto, dto: pvpTotal > 0 ? ((pvpTotal - neto) / pvpTotal) * 100 : 0 };
}

function calcApartados(rows) {
  const cnt = [0, 0, 0, 0], elem = [0, 0, 0, 0];
  const result = {};
  // Une los niveles asegurando que ninguna cifra sea 0 (mínimo 1).
  // Ej: si no hay T1, el primer T2 sería 0.1 → se corrige a 1.1
  const join = (...parts) => parts.map(n => (n < 1 ? 1 : n)).join(".");
  rows.forEach(row => {
    const nat = row.naturaleza;
    if (nat === "T1") {
      cnt[0]++; cnt[1]=0; cnt[2]=0; cnt[3]=0; elem[0]=0; elem[1]=0; elem[2]=0; elem[3]=0;
      result[row.id] = join(cnt[0]);
    } else if (nat === "T2") {
      elem[0]++; cnt[1]=elem[0]; cnt[2]=0; cnt[3]=0; elem[1]=0; elem[2]=0; elem[3]=0;
      result[row.id] = join(cnt[0], cnt[1]);
    } else if (nat === "T3") {
      elem[1]++; cnt[2]=elem[1]; cnt[3]=0; elem[2]=0; elem[3]=0;
      result[row.id] = join(cnt[0], cnt[1], cnt[2]);
    } else if (nat === "T4") {
      elem[2]++; cnt[3]=elem[2]; elem[3]=0;
      result[row.id] = join(cnt[0], cnt[1], cnt[2], cnt[3]);
    } else if (["PD", "PE", "E"].includes(nat)) {
      if (cnt[3] > 0)      { elem[3]++; result[row.id] = join(cnt[0], cnt[1], cnt[2], cnt[3], elem[3]); }
      else if (cnt[2] > 0) { elem[2]++; result[row.id] = join(cnt[0], cnt[1], cnt[2], elem[2]); }
      else if (cnt[1] > 0) { elem[1]++; result[row.id] = join(cnt[0], cnt[1], elem[1]); }
      else if (cnt[0] > 0) { elem[0]++; result[row.id] = join(cnt[0], elem[0]); }
      else { elem[0]++; result[row.id] = join(1, elem[0]); }
    } else if (["S1","S2","S3","S4","TT","CM"].includes(nat)) {
      // Subtotales, total y comentarios: la columna "apartado" debe quedar vacía
      result[row.id] = "";
    } else {
      result[row.id] = row.posicion || "";
    }
  });
  return result;
}

const INITIAL_ROWS = Array.from({ length: 50 }, (_, i) => ({
  id: i + 1, representacion: "", naturaleza: "", posicion: "", cantidad: 0,
  referencia: "", nombre: "", pvp: 0, dtoaplicado: 0, descripcion: "",
  familia: "", subfamilia: "", preciocosteunitario: 0, idposicion: "",
  imagen: "", precionetounitario2: 0, grupodescuento: "",
}));

const AYUDA_TREE = [
  { id: "intro",     label: "¿Qué es esta app?",  icon: Home, children: [] },
  { id: "novedades", label: "Novedades recientes", icon: RefreshCw, children: [] },
  { id: "cabecera",  label: "Cabecera",            icon: FileInput, children: [] },
  { id: "columnas",  label: "Columnas de la grid", icon: BarChart3, children: [] },
  { id: "acciones",  label: "Acciones de la grid", icon: MousePointer, children: [] },
  { id: "statusbar", label: "Barra de estado",     icon: Check, children: [] },
  { id: "autoref",   label: "Autocompletado",      icon: Search, children: [] },
  { id: "menus",     label: "Menús",               icon: FolderOpen, children: [
    { id: "m-presupuesto", label: "Presupuesto", icon: FileText },
    { id: "m-celdas",      label: "Celdas",      icon: Grid3x3 },
    { id: "m-elementos",   label: "Elementos",   icon: Layers },
    { id: "m-productos",   label: "Productos",   icon: Package },
    { id: "m-clientes",    label: "Clientes",    icon: Users },
    { id: "m-descuentos",  label: "Descuentos",  icon: Percent },
    { id: "m-otros",       label: "Otros",       icon: Wrench },
  ]},
];


// ── Exportar presupuesto a Excel ──
// Requiere xlsx-js-style (no la librería xlsx estándar) para aplicar colores y estilos.
function exportToExcel(presupuesto, rows, apartados, estructuraActiva, estilos) {
  // Helper para convertir color CSS (#xxxxxx) a formato RGB hexadecimal (sin #) para xlsx-js-style
  const cssToRgb = (color, fallback = "FFFFFF") => {
    if (!color) return fallback;
    let s = String(color).trim();
    if (s.startsWith("#")) s = s.slice(1);
    if (s.length === 3) s = s.split("").map(c => c + c).join("");
    if (s.length === 6) return s.toUpperCase();
    return fallback;
  };
  const calcNetoUnit = (row) => (row.pvp || 0) * (1 - (row.dtoaplicado || 0) / 100);
  const calcNetoPos  = (row) => calcNetoUnit(row) * (row.cantidad || 0);
  const fmtNum = (n) => isNaN(n) ? 0 : Math.round(n * 100) / 100;

  // Nombre del fichero (limpiar caracteres no válidos para Windows)
  const numCompletoFich = presupuesto.numerocompleto || presupuesto.np || "";
  const nombreFichero = `Oferta SIEMENS ${numCompletoFich} Rev ${presupuesto.revision} - ${presupuesto.titulo} - ${presupuesto.cliente}`
    .replace(/[\\/:*?"<>|]/g, "-");

  // Fecha actual en formato largo español
  const fechaHoy = new Date().toLocaleDateString("es-ES", { day: "2-digit", month: "long", year: "numeric" });

  // ── Paleta de colores ──
  const C = {
    T1: "1F497D",    // azul oscuro
    T2: "4F81BD",    // azul medio
    T3: "DCE6F1",    // azul claro
    T4: "EAF1FA",    // azul muy claro
    S:  "C6D9F0",    // subtotales gris-azulado
    TT: "1F497D",    // total general
    CM: "FEFCE8",    // amarillo claro
    VERDE: "DCFCE7",
    GRIS:  "F1F5F9",
    HEADER: "FAFAFA",  // Supabase: gris muy claro para cabeceras
    LABEL:  "F5F5F5",  // Supabase: gris claro para etiquetas
    HEADER_TEXT: "171717",  // Texto oscuro sobre fondo claro
    WHITE:  "FFFFFF",
    BORDER: "E5E5E5",  // Borde gris suave
    DARKBLUE: "1F497D",
    BROWN:  "713F12",
    GREEN:  "14532D",
    GRAY:   "475569",
  };

  // Borde estándar
  const border = {
    top:    { style: "thin", color: { rgb: C.BORDER } },
    bottom: { style: "thin", color: { rgb: C.BORDER } },
    left:   { style: "thin", color: { rgb: C.BORDER } },
    right:  { style: "thin", color: { rgb: C.BORDER } },
  };

  // Devuelve el estilo de fila según la naturaleza
  const styleByNat = (nat) => {
    // Usar los estilos configurables del usuario (Opciones → Configurar Estilos)
    const cfg = estilos && estilos[nat];
    if (cfg) {
      return {
        fill:  cssToRgb(cfg.bg, null),
        color: cssToRgb(cfg.color, "000000"),
        size:  cfg.fontSize || 11,
        bold:  (cfg.fontWeight || 400) >= 600,
        font:  cfg.fontFamily || "Calibri",
      };
    }
    // Fallback a colores por defecto si no hay configuración
    switch (nat) {
      case "T1": return { fill: C.T1, color: C.WHITE,    size: 13, bold: true, font: "Calibri" };
      case "T2": return { fill: C.T2, color: C.WHITE,    size: 12, bold: true, font: "Calibri" };
      case "T3": return { fill: C.T3, color: C.DARKBLUE, size: 12, bold: true, font: "Calibri" };
      case "T4": return { fill: C.T4, color: C.DARKBLUE, size: 11, bold: true, font: "Calibri" };
      case "S1": return { fill: C.S,  color: C.DARKBLUE, size: 12, bold: true, font: "Calibri" };
      case "S2": return { fill: C.S,  color: C.DARKBLUE, size: 11, bold: true, font: "Calibri" };
      case "S3": return { fill: C.S,  color: C.DARKBLUE, size: 11, bold: true, font: "Calibri" };
      case "S4": return { fill: C.S,  color: C.DARKBLUE, size: 11, bold: true, font: "Calibri" };
      case "TT": return { fill: C.TT, color: C.WHITE,    size: 14, bold: true, font: "Calibri" };
      case "CM": return { fill: C.CM, color: C.BROWN,    size: 11, bold: false, font: "Calibri" };
      case "VERDE": return { fill: C.VERDE, color: C.GREEN, size: 11, bold: false, font: "Calibri" };
      case "GRIS":  return { fill: C.GRIS,  color: C.GRAY,  size: 11, bold: false, font: "Calibri" };
      default:   return { fill: null, color: "000000", size: 11, bold: false, font: "Calibri" };
    }
  };

  // Sangría según nivel
  const indentByNat = { T1: 0, T2: 1, T3: 2, T4: 3, S1: 0, S2: 1, S3: 2, S4: 3, TT: 0 };

  // ── Construir matriz de datos ──
  const aoa = [];
  const numCompleto = presupuesto.numerocompleto || presupuesto.np || "";
  aoa.push([null, `OFERTA SIEMENS — ${numCompleto} Revisión ${presupuesto.revision}`]);
  aoa.push([]);
  aoa.push([null, null, null, "Presupuesto Descripción", presupuesto.titulo]);
  aoa.push([null, null, null, "Cliente",                 presupuesto.cliente]);
  aoa.push([null, null, null, "Presupuesto Número",      `${numCompleto} Revisión ${presupuesto.revision}`]);
  aoa.push([]);
  aoa.push([null, null, null, "Fecha",                   fechaHoy]);
  aoa.push([]);
  aoa.push([null, "Apartado", "Cantidad", "Referencia", "Producto", "Neto Unitario", "Neto Posición", "Descripción"]);
  aoa.push([]);

  const headerRowIdx = 8;  // 0-based: la fila de cabecera de columnas
  const dataStartIdx = aoa.length;  // primera fila de datos (0-based)

  // Metadatos para aplicar estilos después
  const rowMeta = [];

  rows.forEach(row => {
    const nat = row.naturaleza;
    // Con estructura: usar la numeración calculada tal cual (vacía en S1-S4/TT/CM).
    // Sin estructura: usar el apartado persistido en posicion.
    const ap = estructuraActiva ? (apartados[row.id] != null ? apartados[row.id] : "") : (row.posicion || "");
    const indent = indentByNat[nat] !== undefined ? indentByNat[nat] : 0;
    const esProducto = ["PD","PE","E"].includes(nat);
    const esTitulo   = ["T1","T2","T3","T4"].includes(nat);
    const esSubtotal = ["S1","S2","S3","S4","TT"].includes(nat);

    if (esTitulo) {
      aoa.push([null, ap, null, null, row.nombre || "", null, null, null]);
    } else if (esSubtotal) {
      const nivel = { S1:1, S2:2, S3:3, S4:4, TT:0 }[nat];
      let suma = 0;
      if (nat === "TT") {
        suma = rows.reduce((s, r) => s + calcNetoPos(r), 0);
      } else {
        const titleNat = "T" + nivel;
        const subNat = "S" + nivel;
        const idx = rows.findIndex(r => r.id === row.id);
        for (let i = idx - 1; i >= 0; i--) {
          if (rows[i].naturaleza === titleNat) break;
          if (rows[i].naturaleza === subNat) break;
          if (["PD","PE","E"].includes(rows[i].naturaleza)) suma += calcNetoPos(rows[i]);
        }
      }
      aoa.push([null, null, null, null, row.nombre || "", null, fmtNum(suma), null]);
    } else if (esProducto) {
      // TP = "Total Posición": en la impresión, cantidad 1 y el neto unitario pasa a ser
      // el total de la posición (cantidad x neto unitario original). No afecta a pantalla ni BD.
      const esTP = String(row.representacion || "").trim().toUpperCase() === "TP";
      if (esTP) {
        const totalPos = calcNetoPos(row);
        aoa.push([
          null, ap, 1, row.referencia || "",
          row.nombre || "", fmtNum(totalPos), fmtNum(totalPos),
          row.descripcion || "",
        ]);
      } else {
        aoa.push([
          null, ap, row.cantidad || 0, row.referencia || "",
          row.nombre || "", fmtNum(calcNetoUnit(row)), fmtNum(calcNetoPos(row)),
          row.descripcion || "",
        ]);
      }
    } else {
      aoa.push([null, ap, null, null, row.nombre || "", null, null, row.descripcion || ""]);
    }
    const esConf = String(row.representacion || "").trim().toLowerCase() === "conf";
    rowMeta.push({ excelRow: aoa.length - 1, naturaleza: nat, indent, esConf });
  });

  // ── Crear worksheet ──
  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Anchos
  ws["!cols"] = [
    { wch: 4  }, { wch: 10 }, { wch: 9  }, { wch: 27 },
    { wch: 60 }, { wch: 13 }, { wch: 16 }, { wch: 60 },
    { wch: 18 },  // Columna I: "A confirmar por el cliente"
  ];

  // Alturas
  ws["!rows"] = [];
  ws["!rows"][0] = { hpt: 30 };
  ws["!rows"][2] = { hpt: 22 };
  ws["!rows"][3] = { hpt: 22 };
  ws["!rows"][4] = { hpt: 22 };
  ws["!rows"][6] = { hpt: 22 };
  ws["!rows"][8] = { hpt: 28 };

  // Merges
  ws["!merges"] = [
    { s: { r: 0, c: 1 }, e: { r: 0, c: 7 } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 7 } },
    { s: { r: 3, c: 4 }, e: { r: 3, c: 7 } },
    { s: { r: 4, c: 4 }, e: { r: 4, c: 7 } },
    { s: { r: 6, c: 4 }, e: { r: 6, c: 7 } },
  ];

  // Helper para obtener/crear celda y aplicar estilo
  const setStyle = (r, c, style) => {
    const addr = XLSX.utils.encode_cell({ r, c });
    if (!ws[addr]) ws[addr] = { v: "", t: "s" };
    ws[addr].s = style;
  };

  // ── Estilo título principal (Supabase: gris claro + texto oscuro) ──
  for (let c = 1; c <= 7; c++) {
    setStyle(0, c, {
      font: { name: "Calibri", sz: 16, bold: true, color: { rgb: C.HEADER_TEXT } },
      fill: { patternType: "solid", fgColor: { rgb: C.HEADER } },
      alignment: { horizontal: "center", vertical: "center" },
      border,
    });
  }

  // ── Cabecera del presupuesto (filas 2, 3, 4, 6) ──
  // La etiqueta va ahora en la columna D (índice 3); el valor sigue en E..H (merge 4..7)
  [2, 3, 4, 6].forEach(r => {
    setStyle(r, 3, {
      font: { name: "Calibri", sz: 11, bold: true, color: { rgb: C.HEADER_TEXT } },
      fill: { patternType: "solid", fgColor: { rgb: C.LABEL } },
      alignment: { horizontal: "right", vertical: "center", indent: 1 },
      border,
    });
    // Valor merge (columnas 4..7)
    for (let c = 4; c <= 7; c++) {
      setStyle(r, c, {
        font: { name: "Calibri", sz: 11, bold: false, color: { rgb: "000000" } },
        fill: { patternType: "solid", fgColor: { rgb: C.WHITE } },
        alignment: { horizontal: "left", vertical: "center", indent: 1, wrapText: true },
        border: c === 4 ? border : { right: border.right, top: border.top, bottom: border.bottom },
      });
    }
  });

  // ── Cabecera de columnas (Supabase: gris claro + texto oscuro) ──
  for (let c = 1; c <= 7; c++) {
    setStyle(headerRowIdx, c, {
      font: { name: "Calibri", sz: 10, bold: true, color: { rgb: C.HEADER_TEXT } },
      fill: { patternType: "solid", fgColor: { rgb: C.HEADER } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border,
    });
  }

  // ── Filas de datos ──
  // Para títulos T1-T4: aplicar el color de fondo solo a las columnas Apartado(1) y Producto(4)
  // Para subtotales S1-S4 y TT: aplicar el color de fondo solo a Producto(4), Neto Unitario(5) y Neto Posición(6)
  // Para productos PD/PE/E y otros: aplicar a toda la fila (como ahora)
  rowMeta.forEach(meta => {
    const r = meta.excelRow;
    // Si la representación es "conf", la fila usa el estilo CONF (Confirmar por el cliente)
    const st = meta.esConf ? styleByNat("CONF") : styleByNat(meta.naturaleza);
    const esProducto = ["PD","PE","E"].includes(meta.naturaleza);
    const esTitulo   = ["T1","T2","T3","T4"].includes(meta.naturaleza);
    const esSubtotal = ["S1","S2","S3","S4","TT"].includes(meta.naturaleza);

    // Columnas en las que aplicar el FONDO/color de fuente del estilo de la naturaleza
    let columnasConEstilo;
    if (esTitulo) {
      columnasConEstilo = new Set([1, 4]); // Apartado, Producto
    } else if (esSubtotal) {
      columnasConEstilo = new Set([4, 6]); // Producto, Neto Posición
    } else {
      columnasConEstilo = new Set([1, 2, 3, 4, 5, 6, 7]); // toda la fila (productos PD/PE/E, CM, VERDE, GRIS, etc.)
    }

    for (let c = 1; c <= 7; c++) {
      // Columnas B(1), C(2) y E(4) → centradas horizontal y verticalmente
      const isCentrada = c === 1 || c === 2 || c === 4;
      // Columnas de importe en euros (Neto Unitario=5, Neto Posición=6) → a la derecha
      const isEuro = c === 5 || c === 6;
      const isLeftAlign = c === 3 || c === 7;
      const isProductoCol = c === 4;
      const aplicarEstilo = columnasConEstilo.has(c);

      const cellStyle = {
        font: aplicarEstilo
          ? { name: st.font || "Calibri", sz: st.size, bold: st.bold, color: { rgb: st.color } }
          : { name: "Calibri", sz: 11, bold: false, color: { rgb: "171717" } },
        alignment: {
          horizontal: isEuro ? "right" : (isCentrada ? "center" : (isLeftAlign ? "left" : "center")),
          vertical: "center",
          wrapText: true,
          // En la columna Producto, si va centrada no aplicamos sangría
          indent: (isProductoCol && !isCentrada) ? meta.indent : 0,
        },
        border,
      };
      if (aplicarEstilo && st.fill) {
        cellStyle.fill = { patternType: "solid", fgColor: { rgb: st.fill } };
      }
      setStyle(r, c, cellStyle);

      // Formato numérico
      const addr = XLSX.utils.encode_cell({ r, c });
      if ((c === 5 || c === 6) && ws[addr] && typeof ws[addr].v === "number") {
        // Moneda: 2 decimales, separador de miles y símbolo € a la derecha
        ws[addr].z = '#,##0.00" \u20AC"';
      }
      if (c === 2 && ws[addr] && typeof ws[addr].v === "number") {
        ws[addr].z = "#,##0";
      }
      // En subtotales destacar el importe en negrita
      if (esSubtotal && c === 6 && ws[addr]) {
        ws[addr].s.font.bold = true;
      }
    }

    // Columna I (índice 8): texto "A confirmar por el cliente" en las filas con representación conf,
    // con el estilo de la fila (CONF) y justificado en varias líneas dentro de la celda.
    if (meta.esConf) {
      const addrI = XLSX.utils.encode_cell({ r, c: 8 });
      ws[addrI] = { t: "s", v: "A confirmar por el cliente" };
      ws[addrI].s = {
        font: { name: st.font || "Calibri", sz: st.size, bold: st.bold, color: { rgb: st.color } },
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        border,
      };
      if (st.fill) ws[addrI].s.fill = { patternType: "solid", fgColor: { rgb: st.fill } };
      // Asegurar que el rango de la hoja incluye la columna I
      const ref = XLSX.utils.decode_range(ws["!ref"]);
      if (ref.e.c < 8) { ref.e.c = 8; ws["!ref"] = XLSX.utils.encode_range(ref); }
    }
  });

  // ── Workbook ──
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Hoja1");

  // Segunda hoja: condiciones comerciales (a partir de CONDICIONES_PARTICULARES_SUMINISTRO)
  const condAoa = CONDICIONES_PARTICULARES_SUMINISTRO.map(c => [c.texto || ""]);
  const wsCond = XLSX.utils.aoa_to_sheet(condAoa);
  wsCond["!cols"] = [{ wch: 120 }];
  const condRows = [];
  CONDICIONES_PARTICULARES_SUMINISTRO.forEach((c, i) => {
    const addr = XLSX.utils.encode_cell({ r: i, c: 0 });
    const cell = wsCond[addr];
    const esTitulo = c.negrita && /^\d/.test(String(c.texto).trim()); // "1.", "1.1", "2." → título
    const esEnlace = !!c.enlace;
    if (!cell) { condRows.push({ hpt: 15 }); return; }
    cell.s = {
      font: {
        name: "Calibri",
        sz: esTitulo ? (/^\d\.\s/.test(c.texto) || /^\d\.$/.test(c.texto.trim()) ? 12 : 11) : 10,
        bold: !!c.negrita || esEnlace,
        color: { rgb: esEnlace ? "0563C1" : "000000" },
        underline: esEnlace,
      },
      alignment: { horizontal: "left", vertical: "center", wrapText: true, indent: esTitulo ? 0 : 1 },
    };
    // Hipervínculo clicable
    if (esEnlace) {
      cell.l = { Target: c.enlace, Tooltip: c.enlace };
    }
    // Altura de fila: títulos algo más altos; líneas largas se ajustan
    condRows.push({ hpt: c.texto === "" ? 8 : (esTitulo ? 20 : 15) });
  });
  wsCond["!rows"] = condRows;
  XLSX.utils.book_append_sheet(wb, wsCond, "Condiciones Comerciales");

  descargarXLSX(wb, nombreFichero + ".xlsx");
}

// ── Diálogo Importar ──
// Columnas válidas y sus alias
const COLUMNAS_VALIDAS = [
  { key: "representacion",      aliases: ["representacion", "repres", "rep", "representación"] },
  { key: "naturaleza",          aliases: ["naturaleza", "nat"] },
  { key: "posicion",            aliases: ["apartado", "posicion", "pos", "apart"] },
  { key: "cantidad",            aliases: ["cantidad", "can", "cant", "qty"] },
  { key: "referencia",          aliases: ["referencia", "ref", "refer"] },
  { key: "nombre",              aliases: ["producto", "elemento", "nombre", "prod", "prod/elem", "producto/elemento"] },
  { key: "pvp",                 aliases: ["pvp", "pvp unitario", "pvpunitario", "precio"] },
  { key: "dtoaplicado",         aliases: ["dto", "descuento", "dto%", "dto.%"] },
  { key: "precionetounitario",  aliases: ["neto unitario", "netounitario", "netounit"] },
  { key: "precionetoposicion",  aliases: ["neto posicion", "netoposicion", "neto pos"] },
  { key: "descripcion",         aliases: ["descripcion", "desc", "descr", "descripción"] },
  { key: "familia",             aliases: ["familia", "fam"] },
  { key: "subfamilia",          aliases: ["subfamilia", "subfam", "sfam"] },
  { key: "preciocosteunitario", aliases: ["coste unitario", "costeunitario", "coste", "costeunit", "coste unit", "ga"] },
  { key: "costeposicion",       aliases: ["coste posicion", "costeposicion", "coste pos"] },
  { key: "margen",              aliases: ["margen", "marg", "margen%"] },
  { key: "idposicion",          aliases: ["idposicion", "id posicion", "idpos"] },
  { key: "imagen",              aliases: ["imagen", "img"] },
  { key: "precionetounitario2", aliases: ["neto unitario 2", "netounitario2", "neto2", "neto unit 2"] },
  { key: "grupodescuento",      aliases: ["grupo descuento", "grupodescuento", "grupodto", "gdto"] },
];

// Normaliza un texto: minúsculas, sin acentos, sin espacios extra
function normalizarHeader(h) {
  return String(h || "")
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

// Para cada cabecera del fichero, busca a qué columna válida corresponde.
// Devuelve { mapping: { headerIdx: keyDeColumna }, errores: [] }
function matchHeaders(headers) {
  const mapping = {};
  const errores = [];

  headers.forEach((header, idx) => {
    const normalizado = normalizarHeader(header);
    if (!normalizado) return;

    // Buscar coincidencias: alias que empiece con el texto del header,
    // o header que empiece con un alias
    const candidatos = COLUMNAS_VALIDAS.filter(col =>
      col.aliases.some(alias => {
        const a = normalizarHeader(alias);
        return a === normalizado || a.startsWith(normalizado) || normalizado.startsWith(a);
      })
    );

    // Si hay coincidencia exacta con un alias, esa gana
    const exacta = COLUMNAS_VALIDAS.filter(col =>
      col.aliases.some(alias => normalizarHeader(alias) === normalizado)
    );

    if (exacta.length === 1) {
      mapping[idx] = exacta[0].key;
    } else if (exacta.length > 1) {
      errores.push(`Columna "${header}" coincide exactamente con varias columnas válidas`);
    } else if (candidatos.length === 1) {
      mapping[idx] = candidatos[0].key;
    } else if (candidatos.length > 1) {
      const nombres = candidatos.map(c => c.key).join(", ");
      errores.push(`Columna "${header}" es ambigua, podría ser: ${nombres}`);
    } else {
      errores.push(`Columna "${header}" no se reconoce como columna válida`);
    }
  });

  return { mapping, errores };
}

// Parsea texto pegado (TSV o CSV) en array de arrays
function parsePastedText(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
  return lines.map(line => {
    // Detectar separador: tabulador, punto y coma o coma
    if (line.includes("\t")) return line.split("\t");
    if (line.includes(";")) return line.split(";");
    return line.split(",");
  });
}

function ImportarDialog({ onClose, onImport }) {
  const [fileData, setFileData] = useState(null);
  const [previewData, setPreviewData] = useState(null);
  const [mapping, setMapping] = useState({});
  const [errores, setErrores] = useState([]);
  const [step, setStep] = useState("input"); // "input" | "preview"
  const [pastedText, setPastedText] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [log, setLog] = useState([]);
  const logRef = useRef(null);
  const addLog = (texto, tipo = "info") => {
    const hora = new Date().toLocaleTimeString("es-ES");
    setLog(l => [...l, { texto, tipo, hora }]);
    setTimeout(() => { if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight; }, 0);
  };

  const handleFile = async (file) => {
    if (!file) return;
    addLog(`Cargando fichero "${file.name}" (${(file.size / 1024).toFixed(1)} KB)...`, "info");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const hoja = wb.SheetNames[0];
      addLog(`Fichero leído. Hoja activa: "${hoja}"`, "info");
      const ws = wb.Sheets[hoja];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      procesar(aoa, file.name);
    } catch (e) {
      addLog("Error leyendo fichero: " + e.message, "error");
    }
  };

  const handlePasted = () => {
    if (!pastedText.trim()) return;
    addLog(`Procesando texto pegado (${pastedText.length} caracteres)...`, "info");
    const aoa = parsePastedText(pastedText);
    procesar(aoa, "(texto pegado)");
  };

  const procesar = (aoa, origen = "fichero") => {
    if (aoa.length === 0) {
      addLog("ERROR: el fichero está vacío", "error");
      setErrores(["El fichero está vacío"]);
      return;
    }
    const headers = aoa[0].map(h => String(h || "").trim());
    const dataRows = aoa.slice(1).filter(r => r.some(c => String(c || "").trim() !== ""));
    addLog(`Detectadas ${headers.length} columnas y ${dataRows.length} filas de datos`, "info");
    addLog(`Cabeceras: ${headers.join(", ")}`, "info");
    const { mapping, errores } = matchHeaders(headers);
    const reconocidas = Object.values(mapping);
    addLog(`Columnas reconocidas: ${reconocidas.join(", ") || "ninguna"}`, reconocidas.length > 0 ? "success" : "warning");
    if (errores.length > 0) {
      errores.forEach(e => addLog("AVISO: " + e, "warning"));
    } else {
      addLog("Validación correcta. Pulsa 'Confirmar e insertar' para añadir las filas al presupuesto.", "success");
    }
    setPreviewData({ headers, rows: dataRows });
    setMapping(mapping);
    setErrores(errores);
    setStep("preview");
  };

  const confirmar = () => {
    if (errores.length > 0) {
      addLog("No se puede confirmar mientras haya errores", "error");
      return;
    }
    addLog(`Insertando ${previewData.rows.length} filas en el presupuesto...`, "info");
    const nuevasFilas = previewData.rows.map(srcRow => {
      const nueva = {
        representacion: "", naturaleza: "PD", posicion: "", cantidad: 0,
        referencia: "", nombre: "", pvp: 0, dtoaplicado: 0, descripcion: "",
        familia: "", subfamilia: "", preciocosteunitario: 0,
        idposicion: "", imagen: "", precionetounitario2: 0, grupodescuento: "",
      };
      let netoUnitImportado = null;
      Object.entries(mapping).forEach(([headerIdx, key]) => {
        const valor = srcRow[Number(headerIdx)];
        if (valor === undefined || valor === "") return;
        // Convertir a número los campos numéricos
        if (key === "cantidad") {
          const num = parseInt(String(valor).replace(",", "."), 10);
          nueva[key] = isNaN(num) ? 0 : num;
        } else if (["pvp","dtoaplicado","preciocosteunitario","precionetounitario2"].includes(key)) {
          const num = parseFloat(String(valor).replace(",", "."));
          nueva[key] = isNaN(num) ? 0 : num;
        } else if (key === "precionetounitario") {
          // Campo calculado: capturar el valor para usarlo como PVP con Dto=0
          const num = parseFloat(String(valor).replace(",", "."));
          if (!isNaN(num)) netoUnitImportado = num;
        } else if (key === "precionetoposicion" || key === "costeposicion" || key === "margen") {
          // Otros campos calculados: ignorar, se recalcularán
        } else {
          nueva[key] = String(valor);
        }
      });
      // El neto unitario tiene prioridad sobre el PVP:
      // si nos dieron neto, lo guardamos como PVP con dto=0 (aunque hubiera PVP en el fichero)
      if (netoUnitImportado !== null) {
        nueva.pvp = netoUnitImportado;
        nueva.dtoaplicado = 0;
      }
      return nueva;
    });
    addLog(`Insertadas ${nuevasFilas.length} filas correctamente`, "success");
    onImport(nuevasFilas);
    onClose();
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "90%", maxWidth: 900, maxHeight: "90vh", overflow: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>

        {/* Cabecera */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, borderBottom: "1px solid #e2e8f0", paddingBottom: 12 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}><Icon as={Download} size={18} color="#1e3a5f" /> Importar presupuesto desde Excel</h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}><Icon as={X} size={18} /></button>
        </div>

        {step === "input" && (<>
          {/* Opción 1: Drag & drop */}
          <div style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", display: "block", marginBottom: 6 }}>1. Arrastra o selecciona un fichero Excel</label>
            <div
              onDragOver={e => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={e => {
                e.preventDefault();
                setDragOver(false);
                const file = e.dataTransfer.files?.[0];
                if (file) handleFile(file);
              }}
              onClick={() => document.getElementById("importar-file-input")?.click()}
              style={{ padding: "20px", border: `2px dashed ${dragOver ? "#2563eb" : "#d4d4d4"}`, borderRadius: 8, background: dragOver ? "#eff6ff" : "#fafafa", textAlign: "center", fontSize: 12, color: "#525252", cursor: "pointer" }}>
              <Icon as={Download} size={24} color="#2563eb" />
              <div style={{ marginTop: 6 }}>
                {dragOver
                  ? <strong>Suelta el fichero aquí</strong>
                  : <>Arrastra un Excel/CSV (.xlsx, .xls, .csv), o haz click para elegir uno.</>}
              </div>
            </div>
            <input id="importar-file-input" type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
              onChange={e => { if (e.target.files?.[0]) handleFile(e.target.files[0]); e.target.value = ""; }} />
          </div>

          <div style={{ textAlign: "center", color: "#94a3b8", fontSize: 11, margin: "12px 0" }}>— O —</div>

          {/* Opción 2: Pegar */}
          <div style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", display: "block", marginBottom: 6 }}>2. Pegar contenido (copiar/pegar desde Excel)</label>
            <textarea value={pastedText} onChange={e => setPastedText(e.target.value)}
              placeholder="Pega aquí las celdas copiadas desde Excel..."
              style={{ width: "100%", minHeight: 100, padding: "8px 10px", border: "1px solid #cbd5e1", borderRadius: 6, fontSize: 12, fontFamily: "monospace", resize: "vertical", background: "#f8fafc" }} />
            <button onClick={handlePasted} disabled={!pastedText.trim()}
              style={{ marginTop: 8, padding: "6px 14px", borderRadius: 6, border: "none", background: pastedText.trim() ? "#2563eb" : "#cbd5e1", color: "#fff", cursor: pastedText.trim() ? "pointer" : "default", fontSize: 12 }}>
              Procesar texto pegado
            </button>
          </div>

          {/* Ayuda */}
          <div style={{ background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 6, padding: "10px 12px", fontSize: 11, color: "#0369a1", marginBottom: 12 }}>
            <strong>Formato esperado:</strong> primera fila con nombres de columna (referencia, cantidad, descripcion...).
            Se permiten abreviaturas (Ref, Can, Desc) pero no ambiguas: "re" sería ambigua entre <em>referencia</em> y <em>representación</em>.
          </div>
        </>)}

        {step === "preview" && previewData && (<>
          {/* Errores */}
          {errores.length > 0 && (
            <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "10px 12px", fontSize: 12, color: "#991b1b", marginBottom: 12 }}>
              <strong>Se han detectado errores:</strong>
              <ul style={{ margin: "6px 0 0 16px", padding: 0 }}>
                {errores.map((e, i) => <li key={i}>{e}</li>)}
              </ul>
            </div>
          )}

          {/* Mapeo */}
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", marginBottom: 6 }}>Mapeo de columnas detectado:</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {previewData.headers.map((h, i) => (
                <div key={i} style={{ fontSize: 11, padding: "4px 8px", borderRadius: 4, background: mapping[i] ? "#dcfce7" : "#fee2e2", border: mapping[i] ? "1px solid #86efac" : "1px solid #fca5a5", color: mapping[i] ? "#14532d" : "#991b1b" }}>
                  <strong>{h}</strong> {mapping[i] ? "→ " + mapping[i] : "(no reconocida)"}
                </div>
              ))}
            </div>
          </div>

          {/* Preview tabla */}
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", marginBottom: 6 }}>Vista previa ({previewData.rows.length} filas):</div>
            <div style={{ maxHeight: 200, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                <thead style={{ position: "sticky", top: 0, background: "#fafafa", borderBottom: "1px solid #e5e5e5" }}>
                  <tr>
                    {previewData.headers.map((h, i) => (
                      <th key={i} style={{ padding: "5px 8px", color: "#171717", fontWeight: 600, textAlign: "left", borderRight: "1px solid #e5e5e5" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {previewData.rows.slice(0, 50).map((row, ri) => (
                    <tr key={ri} style={{ background: ri % 2 === 0 ? "#fff" : "#f8fafc" }}>
                      {previewData.headers.map((_, ci) => (
                        <td key={ci} style={{ padding: "4px 8px", border: "1px solid #f1f5f9", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 150 }}>
                          {String(row[ci] ?? "")}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {previewData.rows.length > 50 && (
                <div style={{ padding: "6px 12px", fontSize: 11, color: "#64748b", background: "#f1f5f9", textAlign: "center" }}>
                  ... y {previewData.rows.length - 50} filas más
                </div>
              )}
            </div>
          </div>

          {/* Botones */}
          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 16, paddingTop: 12, borderTop: "1px solid #e2e8f0" }}>
            <button onClick={() => setStep("input")} style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 12 }}><BtnContent icon={ArrowLeft}>← Volver</BtnContent></button>
            <button onClick={onClose} style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 12 }}><BtnContent icon={X}>Cancelar</BtnContent></button>
            <button onClick={confirmar} disabled={errores.length > 0}
              style={{ padding: "7px 20px", borderRadius: 6, border: "none", background: errores.length === 0 ? "#16a34a" : "#cbd5e1", color: "#fff", cursor: errores.length === 0 ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
              <BtnContent icon={Check} iconColor="#fff">Importar {previewData.rows.length} filas</BtnContent>
            </button>
          </div>
        </>)}

        {/* Log de progreso */}
        <div style={{ marginTop: 8 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: "#171717", marginBottom: 4, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span>Progreso ({log.length} entrada{log.length !== 1 ? "s" : ""})</span>
            {log.length > 0 && (
              <div style={{ display: "inline-flex", gap: 6 }}>
                <button onClick={() => {
                    const texto = log.map(l => {
                      const prefijo = l.tipo === "error" ? "[ERROR]" : l.tipo === "warning" ? "[AVISO]" : l.tipo === "success" ? "[OK]" : "[INFO]";
                      return `[${l.hora}] ${prefijo} ${l.texto}`;
                    }).join("\n");
                    if (navigator.clipboard) navigator.clipboard.writeText(texto);
                  }}
                  style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                  <BtnContent icon={Copy} iconColor="#475569">Copiar log</BtnContent>
                </button>
                <button onClick={() => exportarLogCSV(log, "importacion")}
                  title="Exportar el log a un fichero CSV"
                  style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                  <BtnContent icon={Download} iconColor="#475569">Exportar CSV</BtnContent>
                </button>
                <button onClick={() => setLog([])}
                  style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                  <BtnContent icon={Trash2} iconColor="#475569">Limpiar</BtnContent>
                </button>
              </div>
            )}
          </div>
          <div ref={logRef} style={{ height: 160, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, background: "#0f172a", color: "#e2e8f0", fontFamily: "monospace", fontSize: 11, padding: "8px 12px", lineHeight: 1.5 }}>
            {log.length === 0 ? (
              <div style={{ color: "#64748b", textAlign: "center", padding: "20px 0", fontStyle: "italic" }}>
                Sin actividad. Carga un fichero o pega texto para empezar.
              </div>
            ) : log.map((l, i) => (
              <div key={i} style={{ padding: "1px 0",
                color: l.tipo === "error" ? "#fca5a5"
                  : l.tipo === "warning" ? "#fcd34d"
                  : l.tipo === "success" ? "#86efac"
                  : "#cbd5e1",
                fontWeight: l.tipo === "error" ? 600 : 400 }}>
                <span style={{ color: "#64748b" }}>[{l.hora}]</span>
                {l.tipo === "error" && <span style={{ color: "#fca5a5", fontWeight: 700 }}> ❌ </span>}
                {l.tipo === "warning" && <span style={{ color: "#fcd34d", fontWeight: 700 }}> ⚠ </span>}
                {l.tipo === "success" && <span style={{ color: "#86efac" }}> ✓ </span>}
                {l.tipo === "info" && " "}
                {l.texto}
              </div>
            ))}
          </div>
        </div>

      </div>
    </div>
  );
}

// ── Exportar datos crudos del presupuesto a Excel ──
function exportarDatosExcel(presupuesto, rows) {
  const calcNetoUnit = (row) => (row.pvp || 0) * (1 - (row.dtoaplicado || 0) / 100);
  const calcNetoPos  = (row) => calcNetoUnit(row) * (row.cantidad || 0);
  const calcCostePos = (row) => (row.preciocosteunitario || 0) * (row.cantidad || 0);
  const calcMargen   = (row) => { const n = calcNetoPos(row), c = calcCostePos(row); return n ? ((n - c) / n) * 100 : 0; };
  const fmt = (n) => isNaN(n) ? 0 : Math.round((n || 0) * 100) / 100;

  const numCompletoFich = presupuesto.numerocompleto || presupuesto.np || "";
  const nombreFichero = `Datos Presupuesto ${numCompletoFich} - ${presupuesto.titulo} - ${presupuesto.cliente}`
    .replace(/[\\/:*?"<>|]/g, "-");

  // Filtrar solo filas de productos
  const productos = rows.filter(r => ["PD", "PE", "E"].includes(r.naturaleza));

  // ── HOJA 1: Datos del presupuesto ──
  const aoa = [];

  // Cabecera del presupuesto (un campo por fila)
  aoa.push(["DATOS DEL PRESUPUESTO"]);
  aoa.push([]);
  aoa.push(["ID:",              presupuesto.id]);
  aoa.push(["Nº Presupuesto:",  presupuesto.np]);
  aoa.push(["Revisión:",        presupuesto.revision]);
  aoa.push(["Título:",          presupuesto.titulo]);
  aoa.push(["Cliente:",         presupuesto.cliente]);
  aoa.push(["Fecha exportación:", new Date().toLocaleDateString("es-ES", { day: "2-digit", month: "long", year: "numeric" })]);
  aoa.push(["Nº líneas producto:", productos.length]);
  aoa.push([]);

  // Cabecera de columnas (todas las columnas internas + número de fila)
  aoa.push([
    "Nº Fila",
    "Repres.",
    "Naturaleza",
    "Apartado",
    "Cantidad",
    "Referencia",
    "Producto/Elemento",
    "PVP Unitario",
    "Dto. %",
    "Neto Unitario",
    "Neto Posición",
    "Descripción",
    "Familia",
    "SubFamilia",
    "Coste Unit. (GA)",
    "Coste Posición",
    "Margen %",
    "Id Posición",
    "Imagen",
    "Neto Unitario 2",
    "Grupo Descuento",
  ]);

  // Filas de productos
  productos.forEach((row, idx) => {
    aoa.push([
      idx + 1,
      row.representacion || "",
      row.naturaleza || "",
      row.posicion || "",
      row.cantidad || 0,
      row.referencia || "",
      row.nombre || "",
      fmt(row.pvp),
      fmt(row.dtoaplicado),
      fmt(calcNetoUnit(row)),
      fmt(calcNetoPos(row)),
      row.descripcion || "",
      row.familia || "",
      row.subfamilia || "",
      fmt(row.preciocosteunitario),
      fmt(calcCostePos(row)),
      fmt(calcMargen(row)),
      row.idposicion || "",
      row.imagen || "",
      fmt(row.precionetounitario2),
      row.grupodescuento || "",
    ]);
  });

  // Totales (última fila)
  const totalCantidad = productos.reduce((s, r) => s + (r.cantidad || 0), 0);
  const totalPVPpos   = productos.reduce((s, r) => s + (r.pvp || 0) * (r.cantidad || 0), 0);
  const totalNeto     = productos.reduce((s, r) => s + calcNetoPos(r), 0);
  const totalCoste    = productos.reduce((s, r) => s + calcCostePos(r), 0);
  const margenTotal   = totalNeto ? ((totalNeto - totalCoste) / totalNeto) * 100 : 0;
  const dtoMedio      = totalPVPpos ? ((totalPVPpos - totalNeto) / totalPVPpos) * 100 : 0;

  aoa.push([]);
  aoa.push([
    "TOTAL",
    "", "", "",
    totalCantidad,
    "", "",
    "",                  // PVP unitario (no aplica al total)
    fmt(dtoMedio),       // Dto. medio
    "",                  // Neto unitario (no aplica)
    fmt(totalNeto),      // Neto Posición total
    "",                  // Descripción
    "", "",              // Familia, SubFamilia
    "",                  // Coste unit.
    fmt(totalCoste),     // Coste Posición total
    fmt(margenTotal),    // Margen %
    "", "", "", "",
  ]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Anchos
  ws["!cols"] = [
    { wch: 8  },  // Nº Fila
    { wch: 10 },  // Repres.
    { wch: 12 },  // Naturaleza
    { wch: 10 },  // Apartado
    { wch: 9  },  // Cantidad
    { wch: 22 },  // Referencia
    { wch: 38 },  // Producto/Elemento
    { wch: 12 },  // PVP
    { wch: 8  },  // Dto.
    { wch: 13 },  // Neto Unit
    { wch: 14 },  // Neto Pos
    { wch: 40 },  // Descripción
    { wch: 16 },  // Familia
    { wch: 16 },  // SubFamilia
    { wch: 14 },  // Coste Unit
    { wch: 14 },  // Coste Pos
    { wch: 10 },  // Margen %
    { wch: 12 },  // Id Pos
    { wch: 12 },  // Imagen
    { wch: 12 },  // Neto Unit 2
    { wch: 14 },  // Grupo Dto
  ];

  // Merge título
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 20 } },
  ];

  // Formato numérico
  const headerDataRow = 10; // 0-based: la fila de cabecera de columnas
  const colsDecimales = [7, 8, 9, 10, 14, 15, 16, 19]; // PVP, Dto, Neto Unit, Neto Pos, Coste Unit, Coste Pos, Margen, Neto Unit 2
  const colCantidad = 4;
  for (let r = headerDataRow + 1; r < aoa.length; r++) {
    for (const c of colsDecimales) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (ws[addr] && typeof ws[addr].v === "number") {
        ws[addr].z = "#,##0.00";
      }
    }
    // Cantidad: formato entero
    const addrCant = XLSX.utils.encode_cell({ r, c: colCantidad });
    if (ws[addrCant] && typeof ws[addrCant].v === "number") {
      ws[addrCant].z = "#,##0";
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Datos Presupuesto");
  descargarXLSX(wb, nombreFichero + ".xlsx");
}

// ── Diálogo Borrar Filas Vacías ──
function BorrarFilasVaciasDialog({ onClose, onAceptar, lineasSeleccionadas }) {
  const [columna, setColumna] = useState("");
  const [confirmar, setConfirmar] = useState(false);
  const [cantidadBorrar, setCantidadBorrar] = useState(0);

  const aceptar = () => {
    if (!columna) return;
    // Contar cuántas se borrarían
    const n = lineasSeleccionadas.filter(row => {
      const val = row[columna];
      return val === null || val === undefined || val === "" || val === 0;
    }).length;
    setCantidadBorrar(n);
    setConfirmar(true);
  };

  const confirmarBorrado = () => {
    onAceptar(columna);
    onClose();
  };

  const labelColumna = COLUMNS.find(c => c.key === columna)?.label || columna;

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", minWidth: 440, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, borderBottom: "1px solid #e2e8f0", paddingBottom: 12 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Trash2} size={18} color="#1e3a5f" /> Borrar filas vacías
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        <p style={{ fontSize: 12, color: "#64748b", margin: "0 0 14px", lineHeight: 1.5 }}>
          Selecciona una columna. Se borrarán todas las filas seleccionadas (con el checkbox marcado) cuyo valor en esa columna esté vacío.
        </p>

        <div style={{ background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 4, padding: "8px 12px", marginBottom: 14, fontSize: 12, color: "#0369a1" }}>
          <strong>{lineasSeleccionadas.length}</strong> fila{lineasSeleccionadas.length !== 1 ? "s" : ""} marcada{lineasSeleccionadas.length !== 1 ? "s" : ""} con el checkbox
        </div>

        <div style={{ marginBottom: 14 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", display: "block", marginBottom: 6 }}>Columna a comprobar</label>
          <select value={columna} onChange={e => setColumna(e.target.value)} autoFocus
            style={{ width: "100%", padding: "7px 10px", fontSize: 13, borderRadius: 4, border: "1px solid #cbd5e1", background: "#f8fafc" }}>
            <option value="">-- Seleccionar columna --</option>
            {COLUMNS.filter(c => c.type !== "calc").map(c => (
              <option key={c.key} value={c.key}>{c.label}</option>
            ))}
          </select>
          <div style={{ fontSize: 11, color: "#64748b", marginTop: 4 }}>
            Solo se muestran columnas editables (las calculadas se omiten).
          </div>
        </div>

        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", paddingTop: 12, borderTop: "1px solid #e2e8f0" }}>
          <button onClick={onClose}
            style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 12 }}>
            <BtnContent icon={X}>Cancelar</BtnContent>
          </button>
          <button onClick={aceptar} disabled={!columna || lineasSeleccionadas.length === 0}
            style={{ padding: "7px 20px", borderRadius: 6, border: "none", background: (columna && lineasSeleccionadas.length > 0) ? "#dc2626" : "#cbd5e1", color: "#fff", cursor: (columna && lineasSeleccionadas.length > 0) ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={Trash2} iconColor="#fff">Borrar</BtnContent>
          </button>
        </div>

        {/* Confirmación */}
        {confirmar && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", maxWidth: 460, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                <Icon as={Trash2} size={22} color="#dc2626" />
                <h3 style={{ fontSize: 15, fontWeight: 700, color: "#1e3a5f", margin: 0 }}>Confirmar borrado</h3>
              </div>
              <p style={{ fontSize: 13, color: "#475569", lineHeight: 1.6, margin: "0 0 16px" }}>
                Se borrarán <strong>{cantidadBorrar}</strong> fila{cantidadBorrar !== 1 ? "s" : ""} que tienen la columna <strong>{labelColumna}</strong> vacía.
                {cantidadBorrar === 0 && <><br /><span style={{ color: "#94a3b8" }}>No hay filas con esa columna vacía entre las seleccionadas.</span></>}
              </p>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmar(false)} style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={confirmarBorrado} disabled={cantidadBorrar === 0}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "none", background: cantidadBorrar > 0 ? "#dc2626" : "#cbd5e1", color: "#fff", cursor: cantidadBorrar > 0 ? "pointer" : "default", fontSize: 13, fontWeight: 600 }}>
                  <BtnContent icon={Check} iconColor="#fff">Sí, borrar</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Diálogo Fijar Precio Total / Conjunto ──
function FijarPrecioDialog({ titulo, descripcion, lineasAfectadas, onClose, onAplicar }) {
  const [importe, setImporte] = useState("");
  const [error, setError] = useState(null);

  // Calcular el PVP total y neto actual de las líneas afectadas
  const totales = (() => {
    let pvpTotal = 0;
    let netoTotal = 0;
    lineasAfectadas.forEach(row => {
      const cantidad = row.cantidad || 0;
      const pvp = row.pvp || 0;
      const dto = row.dtoaplicado || 0;
      pvpTotal += pvp * cantidad;
      netoTotal += pvp * (1 - dto / 100) * cantidad;
    });
    return { pvpTotal, netoTotal };
  })();

  const fmt = (n) => (n || 0).toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const aplicar = () => {
    setError(null);
    const objetivo = parseFloat(String(importe).replace(",", "."));
    if (isNaN(objetivo) || objetivo <= 0) {
      setError("Indica un importe válido mayor que cero");
      return;
    }
    if (totales.netoTotal === 0) {
      setError("El neto actual de las líneas seleccionadas es 0, no se puede repartir");
      return;
    }
    // Factor de proporcionalidad respecto al neto actual de cada línea
    const factor = objetivo / totales.netoTotal;
    // Para cada línea: newNetoUnit = netoUnitActual * factor
    //                  newDto = (1 - newNetoUnit / pvp) * 100
    const nuevosDescuentos = {};
    let alguno_fuera_rango = false;
    lineasAfectadas.forEach(row => {
      const pvp = row.pvp || 0;
      const dto = row.dtoaplicado || 0;
      const netoUnit = pvp * (1 - dto / 100);
      const newNetoUnit = netoUnit * factor;
      if (pvp === 0) {
        nuevosDescuentos[row.id] = dto; // sin PVP no podemos calcular, dejar igual
        return;
      }
      const newDto = (1 - newNetoUnit / pvp) * 100;
      if (newDto < -100 || newDto > 100) alguno_fuera_rango = true;
      nuevosDescuentos[row.id] = newDto;
    });
    if (alguno_fuera_rango) {
      setError("El reparto generaría descuentos fuera del rango -100% a 100% en alguna línea");
      return;
    }
    onAplicar(nuevosDescuentos);
    onClose();
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", minWidth: 460, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, borderBottom: "1px solid #e2e8f0", paddingBottom: 12 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Target} size={18} color="#1e3a5f" /> {titulo}
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        <p style={{ fontSize: 12, color: "#64748b", margin: "0 0 16px", lineHeight: 1.5 }}>
          {descripcion}
        </p>

        {/* Resumen actual */}
        <div style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 6, padding: "10px 14px", marginBottom: 14 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12 }}>
            <div><span style={{ color: "#64748b" }}>Líneas afectadas:</span> <strong>{lineasAfectadas.length}</strong></div>
            <div><span style={{ color: "#64748b" }}>PVP total:</span> <strong>{fmt(totales.pvpTotal)} €</strong></div>
            <div><span style={{ color: "#64748b" }}>Neto actual:</span> <strong style={{ color: "#0369a1" }}>{fmt(totales.netoTotal)} €</strong></div>
            <div><span style={{ color: "#64748b" }}>Dto. medio actual:</span> <strong>{totales.pvpTotal > 0 ? fmt((totales.pvpTotal - totales.netoTotal) / totales.pvpTotal * 100) : "0,00"}%</strong></div>
          </div>
        </div>

        {/* Importe objetivo */}
        <div style={{ marginBottom: 14 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", display: "block", marginBottom: 6 }}>Importe neto objetivo (€)</label>
          <input type="number" min="0" step="0.01" value={importe} onChange={e => setImporte(e.target.value)}
            onKeyDown={e => { if (e.key === "Enter") aplicar(); }}
            placeholder={String(Math.round(totales.netoTotal * 100) / 100)}
            autoFocus
            style={{ width: "100%", padding: "8px 12px", fontSize: 14, borderRadius: 6, border: "1px solid #cbd5e1", textAlign: "right", background: "#f8fafc", fontWeight: 600 }} />
          <div style={{ fontSize: 11, color: "#64748b", marginTop: 4 }}>
            Se ajustará el descuento de cada línea para que su neto se reparta proporcionalmente respecto al neto actual y el total sea exactamente este importe.
          </div>
        </div>

        {error && (
          <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#991b1b", marginBottom: 12 }}>
            {error}
          </div>
        )}

        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", paddingTop: 12, borderTop: "1px solid #e2e8f0" }}>
          <button onClick={onClose}
            style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 12 }}>
            <BtnContent icon={X}>Cancelar</BtnContent>
          </button>
          <button onClick={aplicar} disabled={!importe || lineasAfectadas.length === 0}
            style={{ padding: "7px 20px", borderRadius: 6, border: "none", background: (importe && lineasAfectadas.length > 0) ? "#16a34a" : "#cbd5e1", color: "#fff", cursor: (importe && lineasAfectadas.length > 0) ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={Check} iconColor="#fff">Aplicar</BtnContent>
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Diálogo Aplicar Descuentos por SubFamilia ──
// ── Diálogo Aplicar Descuentos por Grupo Descuento ──
function AplicarDescuentosDialog({ rows, onClose, onApply }) {
  // Carga de gruposdescuento (para mostrar DGL1 y DGL2 por código)
  const [gruposBD, setGruposBD] = useState([]);
  useEffect(() => {
    fetch(`${API_URL}/gruposdescuento/`)
      .then(r => r.ok ? r.json() : [])
      .then(setGruposBD)
      .catch(() => setGruposBD([]));
  }, []);
  const dglFor = (codigo) => {
    if (!codigo) return { dgl1: null, dgl2: null };
    const cod = String(codigo).trim().toUpperCase();
    const g = gruposBD.find(x => String(x.grupodescuentospain || "").trim().toUpperCase() === cod);
    return g ? { dgl1: g.dgl1, dgl2: g.dgl2 } : { dgl1: null, dgl2: null };
  };

  // Recopilar combinaciones únicas de GrupoDescuento + Dto actual (solo PD/PE/E)
  const grupos = (() => {
    const mapa = {};
    rows.forEach(row => {
      if (!["PD", "PE", "E"].includes(row.naturaleza)) return;
      const gd = (row.grupodescuento || "").trim() || "(Sin grupo)";
      const dto = Number(row.dtoaplicado) || 0;
      // Clave: grupo descuento + descuento aplicado (redondeado a 2 decimales)
      const dtoKey = Math.round(dto * 100) / 100;
      const key = gd + "||" + dtoKey;
      if (!mapa[key]) {
        mapa[key] = {
          grupodescuento: gd,
          dtoActual: dtoKey,
          familia: row.familia || "(Sin familia)",
          subfamilia: row.subfamilia || "(Sin subfamilia)",
          lineas: 0,
          importeNeto: 0,
          importePvp: 0,
        };
      }
      const netoUnit = (row.pvp || 0) * (1 - dto / 100);
      const netoPos = netoUnit * (row.cantidad || 0);
      const pvpPos = (row.pvp || 0) * (row.cantidad || 0);
      mapa[key].lineas++;
      mapa[key].importeNeto += netoPos;
      mapa[key].importePvp += pvpPos;
    });
    return Object.values(mapa)
      .sort((a, b) => b.importeNeto - a.importeNeto);
  })();

  const [seleccionada, setSeleccionada] = useState(null);
  const [nuevoDto, setNuevoDto] = useState("");

  // Cuando cambia la selección, inicializar el descuento con el actual
  useEffect(() => {
    if (seleccionada) {
      setNuevoDto(String(seleccionada.dtoActual));
    } else {
      setNuevoDto("");
    }
  }, [seleccionada]);

  const fmt = (n) => (n || 0).toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const aplicar = () => {
    if (!seleccionada) return;
    const dto = parseFloat(String(nuevoDto).replace(",", "."));
    if (isNaN(dto) || dto < -100 || dto > 100) {
      alert("Introduce un descuento válido entre -100 y 100 (negativo = recargo)");
      return;
    }
    // Aplica a las líneas con ese grupo descuento Y ese descuento actual
    onApply(seleccionada.grupodescuento, seleccionada.dtoActual, dto);
    onClose();
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "90%", maxWidth: 820, maxHeight: "90vh", overflow: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>

        {/* Cabecera */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, borderBottom: "1px solid #e2e8f0", paddingBottom: 12 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}><Icon as={DollarSign} size={18} color="#1e3a5f" /> Aplicar descuentos por Grupo Descuento</h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}><Icon as={X} size={18} /></button>
        </div>

        {/* Lista grupos */}
        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", marginBottom: 8 }}>
            Grupos de descuento en el presupuesto ({grupos.length}):
          </div>
          {grupos.length === 0 ? (
            <div style={{ padding: "20px", textAlign: "center", color: "#94a3b8", fontSize: 13, background: "#f8fafc", borderRadius: 6, border: "1px dashed #cbd5e1" }}>
              No hay productos con grupo descuento en el presupuesto.
            </div>
          ) : (
            <div style={{ maxHeight: 340, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 1, borderBottom: "1px solid #e5e5e5" }}>
                  <tr>
                    <th style={{ width: 32, padding: "7px 8px", color: "#171717" }}></th>
                    <th style={{ padding: "7px 10px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Grupo Descuento</th>
                    <th style={{ padding: "7px 10px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Familia</th>
                    <th style={{ padding: "7px 10px", textAlign: "left", color: "#171717", fontWeight: 600 }}>SubFamilia</th>
                    <th style={{ padding: "7px 10px", textAlign: "right", color: "#171717", fontWeight: 600 }}>Líneas</th>
                    <th style={{ padding: "7px 10px", textAlign: "right", color: "#171717", fontWeight: 600 }}>Dto. actual</th>
                    <th style={{ padding: "7px 10px", textAlign: "right", color: "#171717", fontWeight: 600 }}>DGL1</th>
                    <th style={{ padding: "7px 10px", textAlign: "right", color: "#171717", fontWeight: 600 }}>DGL2</th>
                    <th style={{ padding: "7px 10px", textAlign: "right", color: "#171717", fontWeight: 600 }}>Importe Neto</th>
                  </tr>
                </thead>
                <tbody>
                  {grupos.map((g, i) => {
                    const isSel = seleccionada && seleccionada.grupodescuento === g.grupodescuento && seleccionada.dtoActual === g.dtoActual;
                    return (
                      <tr key={i}
                        onClick={() => setSeleccionada(g)}
                        style={{ background: isSel ? "#dbeafe" : (i % 2 === 0 ? "#fff" : "#f8fafc"), cursor: "pointer", borderBottom: "1px solid #f1f5f9" }}>
                        <td style={{ padding: "5px 8px", textAlign: "center" }}>
                          <input type="radio" checked={isSel} onChange={() => setSeleccionada(g)} />
                        </td>
                        <td style={{ padding: "6px 10px", fontWeight: 600, color: "#1e3a5f" }}>{g.grupodescuento}</td>
                        <td style={{ padding: "6px 10px" }}>{g.familia}</td>
                        <td style={{ padding: "6px 10px" }}>{g.subfamilia}</td>
                        <td style={{ padding: "6px 10px", textAlign: "right" }}>{g.lineas}</td>
                        {(() => {
                          const d = dglFor(g.grupodescuento);
                          // Color del dto. actual: rojo si > DGL2, naranja si > DGL1, normal si no
                          let dtoColor = "#475569";
                          let dtoWeight = 400;
                          if (d.dgl2 != null && g.dtoActual > Number(d.dgl2)) { dtoColor = "#dc2626"; dtoWeight = 700; }
                          else if (d.dgl1 != null && g.dtoActual > Number(d.dgl1)) { dtoColor = "#ca8a04"; dtoWeight = 600; }
                          return <>
                            <td style={{ padding: "6px 10px", textAlign: "right", color: dtoColor, fontWeight: dtoWeight }}>{fmt(g.dtoActual)}%</td>
                            <td style={{ padding: "6px 10px", textAlign: "right", color: "#475569" }}>{d.dgl1 == null ? "—" : fmt(d.dgl1) + "%"}</td>
                            <td style={{ padding: "6px 10px", textAlign: "right", color: "#475569" }}>{d.dgl2 == null ? "—" : fmt(d.dgl2) + "%"}</td>
                          </>;
                        })()}
                        <td style={{ padding: "6px 10px", textAlign: "right", fontWeight: 500, color: "#0369a1" }}>{fmt(g.importeNeto)} €</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* Descuento a aplicar */}
        {seleccionada && (
          <div style={{ background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 6, padding: "12px 14px", marginBottom: 12 }}>
            <div style={{ fontSize: 12, color: "#1e3a5f", marginBottom: 8 }}>
              Grupo seleccionado: <strong>{seleccionada.grupodescuento}</strong> con dto. actual <strong>{fmt(seleccionada.dtoActual)}%</strong> ({seleccionada.lineas} línea{seleccionada.lineas !== 1 ? "s" : ""})
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f" }}>Nuevo descuento:</label>
              <input
                type="number" min="0" max="100" step="1"
                value={nuevoDto} onChange={e => setNuevoDto(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter") aplicar(); }}
                autoFocus
                style={{ width: 90, padding: "5px 8px", fontSize: 13, borderRadius: 4, border: "1px solid #cbd5e1", textAlign: "right", background: "#fff" }} />
              <span style={{ fontSize: 13, color: "#475569" }}>%</span>
              <span style={{ fontSize: 11, color: "#64748b", marginLeft: 8 }}>
                (actualmente {fmt(seleccionada.dtoActual)}%)
              </span>
            </div>
          </div>
        )}

        {/* Botones */}
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 12, paddingTop: 12, borderTop: "1px solid #e2e8f0" }}>
          <button onClick={onClose} style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 12 }}><BtnContent icon={X}>Cancelar</BtnContent></button>
          <button onClick={aplicar} disabled={!seleccionada || nuevoDto === ""}
            style={{ padding: "7px 20px", borderRadius: 6, border: "none", background: (seleccionada && nuevoDto !== "") ? "#16a34a" : "#cbd5e1", color: "#fff", cursor: (seleccionada && nuevoDto !== "") ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
            Aplicar descuento
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Diálogo Leer Presupuestos ──
function LeerPresupuestosDialog({ onClose, onCargar, setStatus }) {
  const [busqueda, setBusqueda] = useState("");
  const [presupuestos, setPresupuestos] = useState([]);
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState(null);
  const [seleccionado, setSeleccionado] = useState(null);
  const [confirmacion, setConfirmacion] = useState(false);
  const [confirmBorrar, setConfirmBorrar] = useState(false);
  const [borrando, setBorrando] = useState(false);
  const [tab, setTab] = useState("lista"); // "lista" o "detalle"
  const [detalle, setDetalle] = useState(null);
  const [cargandoDetalle, setCargandoDetalle] = useState(false);

  // Cargar el detalle cuando se selecciona un presupuesto y se va a la pestaña detalle
  useEffect(() => {
    if (tab !== "detalle" || !seleccionado) { setDetalle(null); return; }
    setCargandoDetalle(true);
    setDetalle(null);
    fetch(`${API_URL}/presupuestos/${seleccionado.id}/completo`)
      .then(r => r.ok ? r.json() : Promise.reject(new Error("Error " + r.status)))
      .then(d => setDetalle(d))
      .catch(e => setStatus && setStatus("Error cargando detalle: " + e.message, "error"))
      .finally(() => setCargandoDetalle(false));
  }, [tab, seleccionado]);

  // Cargar lista al abrir y cuando cambia búsqueda (con debounce)
  useEffect(() => {
    const t = setTimeout(() => {
      cargarLista();
    }, 300);
    return () => clearTimeout(t);
  }, [busqueda]);

  const cargarLista = async () => {
    setCargando(true);
    setError(null);
    setStatus && setStatus(busqueda.trim() ? `Buscando "${busqueda.trim()}" en presupuestos...` : "Cargando lista de presupuestos...", "working");
    try {
      const url = busqueda.trim()
        ? `${API_URL}/presupuestos/?busqueda=${encodeURIComponent(busqueda.trim())}`
        : `${API_URL}/presupuestos/`;
      const res = await fetch(url);
      if (!res.ok) throw new Error("Error " + res.status);
      const data = await res.json();
      setPresupuestos(data);
      setStatus && setStatus(`${data.length} presupuesto${data.length !== 1 ? "s" : ""} encontrado${data.length !== 1 ? "s" : ""}`, "success");
    } catch (e) {
      setError(e.message || "Error cargando presupuestos");
      setPresupuestos([]);
      setStatus && setStatus("Error al cargar presupuestos: " + e.message, "error");
    } finally {
      setCargando(false);
    }
  };

  const borrarPresupuesto = async () => {
    if (!seleccionado) return;
    setBorrando(true);
    setStatus && setStatus(`Borrando presupuesto ${seleccionado.numerocompleto || seleccionado.numero}...`, "working");
    try {
      const res = await fetch(`${API_URL}/presupuestos/${seleccionado.id}`, { method: "DELETE" });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "HTTP " + res.status);
      }
      setStatus && setStatus(`Presupuesto ${seleccionado.numerocompleto || seleccionado.numero} borrado de la base de datos`, "success");
      setConfirmBorrar(false);
      setSeleccionado(null);
      setDetalle(null);
      setTab("lista");
      cargarLista();
    } catch (e) {
      setStatus && setStatus("Error borrando presupuesto: " + e.message, "error");
    } finally {
      setBorrando(false);
    }
  };

  const leerSeleccionado = async () => {
    if (!seleccionado) return;
    setCargando(true);
    setError(null);
    setStatus && setStatus(`Leyendo presupuesto ${seleccionado.numerocompleto || seleccionado.numero}...`, "working");
    try {
      const res = await fetch(`${API_URL}/presupuestos/${seleccionado.id}/completo`);
      if (!res.ok) throw new Error("Error " + res.status);
      const data = await res.json();
      onCargar(data);
      onClose();
    } catch (e) {
      setError(e.message || "Error leyendo presupuesto");
      setStatus && setStatus("Error al leer presupuesto: " + e.message, "error");
    } finally {
      setCargando(false);
    }
  };

  const formatearFecha = (f) => {
    if (!f) return "";
    try {
      const d = new Date(f);
      return d.toLocaleDateString("es-ES", { day: "2-digit", month: "2-digit", year: "numeric" });
    } catch { return String(f); }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "90%", maxWidth: 1000, height: "85vh", maxHeight: "92vh", minWidth: 480, minHeight: 320, display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.2)", resize: "both", overflow: "auto" }} onClick={e => e.stopPropagation()}>

        {/* Cabecera */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, borderBottom: "1px solid #e2e8f0", paddingBottom: 12, flexShrink: 0 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}><Icon as={FolderOpen} size={18} color="#1e3a5f" /> Leer presupuesto desde la base de datos</h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}><Icon as={X} size={18} /></button>
        </div>

        {/* Pestañas */}
        <div style={{ display: "flex", borderBottom: "1px solid #e5e5e5", marginBottom: 12, flexShrink: 0 }}>
          <button onClick={() => setTab("lista")}
            style={{ padding: "8px 18px", border: "none", background: "transparent", fontSize: 12, fontWeight: 600, cursor: "pointer",
              color: tab === "lista" ? "#171717" : "#737373",
              borderBottom: tab === "lista" ? "2px solid #171717" : "2px solid transparent",
              marginBottom: "-1px" }}>
            <BtnContent icon={FolderOpen} iconColor={tab === "lista" ? "#171717" : "#737373"}>Lista de presupuestos</BtnContent>
          </button>
          <button onClick={() => setTab("detalle")} disabled={!seleccionado}
            style={{ padding: "8px 18px", border: "none", background: "transparent", fontSize: 12, fontWeight: 600, cursor: seleccionado ? "pointer" : "default",
              color: !seleccionado ? "#cbd5e1" : tab === "detalle" ? "#171717" : "#737373",
              borderBottom: tab === "detalle" ? "2px solid #171717" : "2px solid transparent",
              marginBottom: "-1px" }}>
            <BtnContent icon={FileText} iconColor={!seleccionado ? "#cbd5e1" : tab === "detalle" ? "#171717" : "#737373"}>
              Detalle del presupuesto {seleccionado ? `(${seleccionado.numerocompleto || seleccionado.numero})` : ""}
            </BtnContent>
          </button>
        </div>

        {/* TAB: Lista */}
        {tab === "lista" && <>
        {/* Filtro */}
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, flexShrink: 0 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", whiteSpace: "nowrap" }}>Buscar:</label>
          <input type="text" value={busqueda} onChange={e => setBusqueda(e.target.value)}
            placeholder="Título, nombre del cliente o número de presupuesto..."
            autoFocus
            style={{ flex: 1, padding: "6px 10px", fontSize: 12, borderRadius: 6, border: "1px solid #cbd5e1", background: "#f8fafc" }} />
          <button onClick={cargarLista} style={{ padding: "6px 12px", fontSize: 12, borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", cursor: "pointer" }}>
            <BtnContent icon={RefreshCw}>Refrescar</BtnContent>
          </button>
          <span style={{ fontSize: 11, color: "#64748b" }}>{presupuestos.length} resultado{presupuestos.length !== 1 ? "s" : ""}</span>
        </div>

        {/* Error */}
        {error && (
          <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#991b1b", marginBottom: 10, flexShrink: 0 }}>
            {error}
          </div>
        )}

        {/* Tabla */}
        <div style={{ flex: 1, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, marginBottom: 12, minHeight: 200 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 1, borderBottom: "1px solid #e5e5e5" }}>
              <tr>
                <th style={{ width: 32, padding: "8px 6px", color: "#171717" }}></th>
                <th style={{ padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600, whiteSpace: "nowrap", minWidth: 140 }}>Nº Completo</th>
                <th style={{ padding: "8px 10px", textAlign: "center", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>Rev.</th>
                <th style={{ padding: "8px 10px", textAlign: "center", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>Fecha modificación</th>
                <th style={{ padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Título</th>
                <th style={{ padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Cliente</th>
                <th style={{ padding: "8px 10px", textAlign: "right", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>Total</th>
                <th style={{ padding: "8px 10px", textAlign: "right", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>ID</th>
              </tr>
            </thead>
            <tbody>
              {cargando && (
                <tr><td colSpan={7} style={{ padding: "30px", textAlign: "center", color: "#64748b" }}>⏳ Cargando...</td></tr>
              )}
              {!cargando && presupuestos.length === 0 && !error && (
                <tr><td colSpan={7} style={{ padding: "30px", textAlign: "center", color: "#94a3b8" }}>
                  {busqueda ? "No hay presupuestos que coincidan con la búsqueda." : "No hay presupuestos en la base de datos."}
                </td></tr>
              )}
              {!cargando && presupuestos.map(p => {
                const isSel = seleccionado?.id === p.id;
                return (
                  <tr key={p.id}
                    onClick={() => setSeleccionado(p)}
                    onDoubleClick={() => { setSeleccionado(p); setConfirmacion(true); }}
                    style={{ background: isSel ? "#dbeafe" : "transparent", cursor: "pointer", borderBottom: "1px solid #f1f5f9" }}>
                    <td style={{ padding: "5px 6px", textAlign: "center" }}>
                      <input type="radio" checked={isSel} onChange={() => setSeleccionado(p)} />
                    </td>
                    <td style={{ padding: "6px 10px", fontWeight: 600, color: "#1e3a5f", whiteSpace: "nowrap" }}>{p.numerocompleto || p.numero}</td>
                    <td style={{ padding: "6px 10px", textAlign: "center", color: "#475569" }}>{p.revision ?? 0}</td>
                    <td style={{ padding: "6px 10px", textAlign: "center", color: "#64748b", fontSize: 11, whiteSpace: "nowrap" }}>{formatearFecha(p.fecha)}</td>
                    <td style={{ padding: "6px 10px" }}>{p.titulo}</td>
                    <td style={{ padding: "6px 10px", color: "#475569" }}>{p.nombrecomun || p.razonsocial}</td>
                    <td style={{ padding: "6px 10px", textAlign: "right", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>{(Number(p.totalpresupuesto) || 0).toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: "always" })} €</td>
                    <td style={{ padding: "6px 10px", textAlign: "right", color: "#64748b", fontFamily: "monospace", whiteSpace: "nowrap" }}>{p.id}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        </>}

        {/* TAB: Detalle */}
        {tab === "detalle" && (
          <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden", marginBottom: 12 }}>
            {!seleccionado ? (
              <div style={{ padding: "40px", textAlign: "center", color: "#94a3b8", fontSize: 13 }}>
                No hay presupuesto seleccionado. Vuelve a la lista y selecciona uno.
              </div>
            ) : cargandoDetalle ? (
              <div style={{ padding: "40px", textAlign: "center", color: "#64748b" }}>
                <Icon as={RefreshCw} size={18} color="#64748b" /> Cargando detalle...
              </div>
            ) : detalle ? (
              <>
                <div style={{ padding: "10px 14px", background: "#f8fafc", border: "1px solid #e5e5e5", borderRadius: 6, marginBottom: 10, display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, fontSize: 12 }}>
                  <div><strong style={{ color: "#737373", fontSize: 10, textTransform: "uppercase" }}>Nº Completo</strong><br /><span style={{ color: "#171717", fontWeight: 600 }}>{detalle.cabecera.numerocompleto || detalle.cabecera.numero}</span></div>
                  <div><strong style={{ color: "#737373", fontSize: 10, textTransform: "uppercase" }}>Revisión</strong><br /><span style={{ color: "#171717" }}>{detalle.cabecera.revision ?? 0}</span></div>
                  <div><strong style={{ color: "#737373", fontSize: 10, textTransform: "uppercase" }}>Título</strong><br /><span style={{ color: "#171717" }}>{detalle.cabecera.titulo || "—"}</span></div>
                  <div><strong style={{ color: "#737373", fontSize: 10, textTransform: "uppercase" }}>Cliente</strong><br /><span style={{ color: "#171717" }}>{detalle.cabecera.nombrecomun || detalle.cabecera.razonsocial || "—"}</span></div>
                </div>

                <div style={{ flex: 1, overflow: "auto", border: "1px solid #e5e5e5", borderRadius: 6 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                    <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 1, borderBottom: "1px solid #e5e5e5" }}>
                      <tr>
                        <th style={{ padding: "6px 8px", textAlign: "right", color: "#171717", fontWeight: 600 }}>Pos.</th>
                        <th style={{ padding: "6px 8px", textAlign: "center", color: "#171717", fontWeight: 600 }} title="Naturaleza">Nat.</th>
                        <th style={{ padding: "6px 8px", textAlign: "center", color: "#171717", fontWeight: 600 }}>Cant.</th>
                        <th style={{ padding: "6px 8px", textAlign: "left", color: "#171717", fontWeight: 600, whiteSpace: "nowrap", width: 115 }}>Referencia</th>
                        <th style={{ padding: "6px 8px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Nombre</th>
                        <th style={{ padding: "6px 8px", textAlign: "right", color: "#171717", fontWeight: 600 }}>Dto %</th>
                          <th style={{ padding: "6px 8px", textAlign: "right", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>Neto Pos.</th>
                        <th style={{ padding: "6px 8px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Familia</th>
                        <th style={{ padding: "6px 8px", textAlign: "left", color: "#171717", fontWeight: 600 }}>SubFamilia</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(detalle.detalle || []).length === 0 && (
                        <tr><td colSpan={9} style={{ padding: 30, textAlign: "center", color: "#94a3b8" }}>Este presupuesto no tiene líneas.</td></tr>
                      )}
                      {(detalle.detalle || []).map((d, i) => (
                        <tr key={d.id || i} style={{ background: i % 2 === 0 ? "#fff" : "#fafafa", borderBottom: "1px solid #f1f5f9" }}>
                          <td style={{ padding: "4px 8px", textAlign: "right", color: "#64748b", fontFamily: "monospace" }}>{d.posicion}</td>
                          <td style={{ padding: "4px 8px", textAlign: "center", color: "#475569" }}>{d.naturaleza || ""}</td>
                          <td style={{ padding: "4px 8px", textAlign: "center", color: "#475569" }}>{d.cantidad || 0}</td>
                          <td style={{ padding: "4px 8px", color: "#1e3a5f", fontFamily: "monospace", whiteSpace: "nowrap", fontWeight: 600, width: 115 }}>{d.referencia || ""}</td>
                          <td style={{ padding: "4px 8px", color: "#171717" }}>{d.nombre || ""}</td>
                          <td style={{ padding: "4px 8px", textAlign: "right", color: "#525252" }}>{(Number(d.dtoaplicado) || 0).toLocaleString("es-ES", { minimumFractionDigits: 1 })}</td>
                          <td style={{ padding: "4px 8px", textAlign: "right", color: "#0369a1", fontWeight: 600, whiteSpace: "nowrap" }}>{(Number(d.precionetoposicion) || 0).toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: "always" })} €</td>
                          <td style={{ padding: "4px 8px", color: "#737373" }}>{d.familia || ""}</td>
                          <td style={{ padding: "4px 8px", color: "#737373" }}>{d.subfamilia || ""}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div style={{ fontSize: 11, color: "#64748b", marginTop: 6, textAlign: "right" }}>
                  {(detalle.detalle || []).length} línea{(detalle.detalle || []).length !== 1 ? "s" : ""}
                </div>
              </>
            ) : null}
          </div>
        )}

        {/* Botones */}
        <div style={{ display: "flex", gap: 8, justifyContent: "space-between", alignItems: "center", paddingTop: 10, borderTop: "1px solid #e2e8f0", flexShrink: 0 }}>
          <button onClick={() => setConfirmBorrar(true)} disabled={!seleccionado}
            style={{ padding: "7px 14px", borderRadius: 6, border: seleccionado ? "1px solid #fca5a5" : "1px solid #d4d4d4", background: seleccionado ? "#fef2f2" : "#f5f5f5", color: seleccionado ? "#991b1b" : "#737373", cursor: seleccionado ? "pointer" : "default", fontSize: 12 }}>
            <BtnContent icon={Trash2} iconColor={seleccionado ? "#dc2626" : "#737373"}>Borrar Presupuesto</BtnContent>
          </button>
          <div style={{ display: "flex", gap: 8 }}>
          <button onClick={onClose}
            style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
            <BtnContent icon={X}>Cancelar</BtnContent>
          </button>
          <button onClick={() => setConfirmacion(true)} disabled={!seleccionado || cargando}
            style={{ padding: "7px 20px", borderRadius: 6, border: (seleccionado && !cargando) ? "1px solid #16a34a" : "1px solid #d4d4d4", background: (seleccionado && !cargando) ? "#dcfce7" : "#f5f5f5", color: (seleccionado && !cargando) ? "#14532d" : "#737373", cursor: (seleccionado && !cargando) ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={Check} iconColor={(seleccionado && !cargando) ? "#14532d" : "#737373"}>Leer Presupuesto</BtnContent>
          </button>
          </div>
        </div>

        {/* Diálogo de confirmación borrar */}
        {confirmBorrar && seleccionado && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100001 }} onClick={() => !borrando && setConfirmBorrar(false)}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 470, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
              <h3 style={{ margin: 0, marginBottom: 12, fontSize: 15, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
                <Icon as={Trash2} size={18} color="#dc2626" /> ¿Borrar presupuesto?
              </h3>
              <p style={{ fontSize: 13, color: "#525252", marginBottom: 8, lineHeight: 1.5 }}>
                Vas a borrar el presupuesto <strong style={{ color: "#171717" }}>{seleccionado.numerocompleto || seleccionado.numero}</strong> Rev.{seleccionado.revision ?? 0} de la base de datos.
              </p>
              <div style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#1e293b", marginBottom: 10 }}>
                <div><strong>Título:</strong> {seleccionado.titulo}</div>
                <div><strong>Cliente:</strong> {seleccionado.nombrecomun || seleccionado.razonsocial}</div>
              </div>
              <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "10px 12px", fontSize: 12, color: "#991b1b", marginBottom: 16 }}>
                <strong>Atención:</strong> esta acción borrará el presupuesto completo y todas sus líneas de detalle. No se puede deshacer.
              </div>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmBorrar(false)} disabled={borrando}
                  style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: borrando ? "default" : "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={borrarPresupuesto} disabled={borrando}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #dc2626", background: "#fef2f2", color: "#991b1b", cursor: borrando ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={borrando ? RefreshCw : Trash2} iconColor="#dc2626">{borrando ? "Borrando..." : "Sí, borrar"}</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Diálogo de confirmación */}
        {confirmacion && seleccionado && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", maxWidth: 460, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                <Icon as={HelpCircle} size={24} color="#d97706" />
                <h3 style={{ fontSize: 15, fontWeight: 700, color: "#1e3a5f", margin: 0 }}>Confirmar lectura del presupuesto</h3>
              </div>
              <p style={{ fontSize: 13, color: "#475569", lineHeight: 1.6, margin: "0 0 8px" }}>
                Vas a leer el presupuesto <strong>{seleccionado.numerocompleto || seleccionado.numero}</strong> Rev.{seleccionado.revision ?? 0}:
              </p>
              <div style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#1e293b", marginBottom: 12 }}>
                <div><strong>Título:</strong> {seleccionado.titulo}</div>
                <div><strong>Cliente:</strong> {seleccionado.nombrecomun || seleccionado.razonsocial}</div>
              </div>
              <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "10px 12px", fontSize: 12, color: "#991b1b", marginBottom: 16 }}>
                <strong>Atención:</strong> esta acción <strong>sobreescribirá todos los datos del presupuesto en pantalla</strong>.
              </div>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmacion(false)} style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}><BtnContent icon={X}>Cancelar</BtnContent></button>
                <button onClick={leerSeleccionado} disabled={cargando}
                  onMouseEnter={e => { if (!cargando) { e.currentTarget.style.background = "#bbf7d0"; e.currentTarget.style.borderColor = "#15803d"; } }}
                  onMouseLeave={e => { if (!cargando) { e.currentTarget.style.background = "#dcfce7"; e.currentTarget.style.borderColor = "#16a34a"; } }}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #16a34a", background: cargando ? "#f1f5f9" : "#dcfce7", color: "#14532d", cursor: cargando ? "default" : "pointer", fontSize: 13, fontWeight: 600, transition: "all 0.15s ease" }}>
                  <BtnContent icon={cargando ? RefreshCw : Check} iconColor="#14532d">{cargando ? "Cargando..." : "Sí, leer y sobreescribir"}</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}

// ── Diálogo Gestión de Clientes ──
const CLIENTE_COLS = [
  { key: "id",          label: "ID",           width: 70,  readonly: true },
  { key: "razonsocial", label: "Razón Social", width: 200 },
  { key: "nombrecomun", label: "Nombre Común", width: 180 },
  { key: "nif",         label: "NIF",          width: 100 },
  { key: "ifa",         label: "IFA",          width: 90,  type: "number" },
  { key: "direccion",   label: "Dirección",    width: 200 },
  { key: "poblacion",   label: "Población",    width: 140 },
  { key: "cp",          label: "C.P.",         width: 70,  type: "number" },
  { key: "telefono1",   label: "Teléfono",     width: 120 },
  { key: "idprovincia", label: "Provincia",    width: 160, type: "provincia" },
];

// ── Diálogo Gestionar Contactos ──
function ContactosDialog({ onClose, setStatus }) {
  const [contactos, setContactos] = useState([]);
  const [clientes, setClientes] = useState([]);
  const [filtroClienteId, setFiltroClienteId] = useState(null); // null = todos
  const [busqueda, setBusqueda] = useState("");
  const [cargando, setCargando] = useState(false);
  const [seleccionado, setSeleccionado] = useState(null); // id contacto en edición
  const [draft, setDraft] = useState(null);
  const [confirmBorrar, setConfirmBorrar] = useState(null);
  const [error, setError] = useState(null);

  const cargar = async () => {
    setCargando(true); setError(null);
    try {
      const params = filtroClienteId ? `?idcliente=${filtroClienteId}` : "";
      const res = await fetch(`${API_URL}/contactos/${params}`);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      setContactos(data);
      setStatus && setStatus(`${data.length} contacto${data.length !== 1 ? "s" : ""} cargado${data.length !== 1 ? "s" : ""}`, "success");
    } catch (e) {
      setError(e.message);
      setStatus && setStatus("Error cargando contactos: " + e.message, "error");
    } finally {
      setCargando(false);
    }
  };

  useEffect(() => { cargar(); }, [filtroClienteId]);

  // Cargar lista de clientes una vez para el selector
  useEffect(() => {
    fetch(`${API_URL}/clientes/`)
      .then(r => r.ok ? r.json() : [])
      .then(setClientes)
      .catch(() => setClientes([]));
  }, []);

  const filtrados = contactos.filter(c =>
    !busqueda.trim() ||
    String(c.nombre || "").toLowerCase().includes(busqueda.toLowerCase()) ||
    String(c.email || "").toLowerCase().includes(busqueda.toLowerCase()) ||
    String(c.cargo || "").toLowerCase().includes(busqueda.toLowerCase())
  );

  const editar = (c) => {
    setSeleccionado(c.id);
    setDraft({ ...c });
  };

  const nuevo = () => {
    setSeleccionado("nuevo");
    setDraft({
      nombre: "",
      email: "",
      cargo: "",
      telefono1: "",
      telefono2: "",
      idcliente: filtroClienteId || 1, // si hay filtro de cliente activo lo usa, si no 1 (no definido)
    });
  };

  const cancelar = () => { setSeleccionado(null); setDraft(null); };

  const guardar = async () => {
    if (!draft) return;
    if (!draft.nombre || !draft.nombre.trim()) {
      setStatus && setStatus("El nombre es obligatorio", "error");
      return;
    }
    setStatus && setStatus("Guardando contacto...", "working");
    try {
      const url = seleccionado === "nuevo" ? `${API_URL}/contactos/` : `${API_URL}/contactos/${seleccionado}`;
      const method = seleccionado === "nuevo" ? "POST" : "PUT";
      const res = await fetch(url, {
        method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(draft),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "HTTP " + res.status);
      }
      setStatus && setStatus(`Contacto "${draft.nombre}" ${seleccionado === "nuevo" ? "creado" : "actualizado"}`, "success");
      setSeleccionado(null); setDraft(null);
      cargar();
    } catch (e) {
      setStatus && setStatus("Error guardando: " + e.message, "error");
    }
  };

  const borrar = async (id, nombre) => {
    setStatus && setStatus(`Borrando "${nombre}"...`, "working");
    try {
      const res = await fetch(`${API_URL}/contactos/${id}`, { method: "DELETE" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      setStatus && setStatus(`Contacto "${nombre}" borrado`, "success");
      setConfirmBorrar(null);
      if (seleccionado === id) { setSeleccionado(null); setDraft(null); }
      cargar();
    } catch (e) {
      setStatus && setStatus("Error borrando: " + e.message, "error");
    }
  };

  const nombreCliente = (id) => {
    const cl = clientes.find(c => c.id === id);
    return cl ? (cl.nombrecomun || cl.razonsocial || `Cliente ${id}`) : `Cliente ${id}`;
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 1100, height: "85vh", maxHeight: "92vh", minWidth: 480, minHeight: 320, overflow: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.2)", resize: "both" }} onClick={e => e.stopPropagation()}>
        {/* Cabecera */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, paddingBottom: 12, borderBottom: "1px solid #e5e5e5" }}>
          <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={User} size={18} color="#171717" /> Gestionar Contactos
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        {/* Filtro por cliente + búsqueda */}
        <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
          <label style={{ fontSize: 12, color: "#525252", fontWeight: 500 }}>Cliente:</label>
          <select value={filtroClienteId || ""} onChange={e => setFiltroClienteId(e.target.value ? parseInt(e.target.value) : null)}
            style={{ padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 6, fontSize: 12, background: "#fff", minWidth: 200 }}>
            <option value="">— Todos los clientes —</option>
            {clientes.map(c => (
              <option key={c.id} value={c.id}>{c.nombrecomun || c.razonsocial || `Cliente ${c.id}`}</option>
            ))}
          </select>
          <input value={busqueda} onChange={e => setBusqueda(e.target.value)}
            placeholder="Buscar nombre, email o cargo..."
            style={{ flex: 1, padding: "5px 10px", border: "1px solid #d4d4d4", borderRadius: 6, fontSize: 12, minWidth: 200 }} />
          <button onClick={nuevo}
            style={{ padding: "5px 12px", borderRadius: 6, border: "1px solid #16a34a", background: "#dcfce7", color: "#14532d", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={Plus} iconColor="#14532d">Nuevo</BtnContent>
          </button>
          <button onClick={cargar} title="Recargar"
            style={{ padding: "5px 10px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", cursor: "pointer", fontSize: 12 }}>
            <BtnContent icon={RefreshCw} iconColor="#475569" />
          </button>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: draft ? "1fr 380px" : "1fr", gap: 14 }}>
          {/* Lista */}
          <div style={{ border: "1px solid #e5e5e5", borderRadius: 8, overflow: "hidden", background: "#fff" }}>
            {error && <div style={{ padding: 10, color: "#991b1b", background: "#fef2f2", borderBottom: "1px solid #fecaca", fontSize: 12 }}>Error: {error}</div>}
            {cargando ? (
              <div style={{ padding: 16, textAlign: "center", color: "#737373", fontSize: 12 }}>Cargando...</div>
            ) : filtrados.length === 0 ? (
              <div style={{ padding: 16, textAlign: "center", color: "#737373", fontSize: 12 }}>
                {contactos.length === 0 ? "No hay contactos" : "Ninguno coincide con la búsqueda"}
              </div>
            ) : (
              <div style={{ maxHeight: 450, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 1 }}>
                    <tr style={{ color: "#171717", fontSize: 11 }}>
                      <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Nombre</th>
                      <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Cargo</th>
                      <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Email</th>
                      <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Tel. 1</th>
                      <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Tel. 2</th>
                      <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Cliente</th>
                      <th style={{ padding: "6px 8px", textAlign: "right", borderBottom: "1px solid #e5e5e5", fontWeight: 600, width: 90 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtrados.map(c => (
                      <tr key={c.id} style={{ borderBottom: "1px solid #f1f5f9",
                          background: seleccionado === c.id ? "#eff6ff" : "transparent" }}>
                        <td style={{ padding: "6px 8px", fontWeight: 500, color: "#171717" }}>{c.nombre}</td>
                        <td style={{ padding: "6px 8px", color: "#525252" }}>{c.cargo || "—"}</td>
                        <td style={{ padding: "6px 8px", color: "#525252", fontSize: 11 }}>{c.email || "—"}</td>
                        <td style={{ padding: "6px 8px", color: "#525252", fontSize: 11 }}>{c.telefono1 || "—"}</td>
                        <td style={{ padding: "6px 8px", color: "#525252", fontSize: 11 }}>{c.telefono2 || "—"}</td>
                        <td style={{ padding: "6px 8px", color: "#737373", fontSize: 11 }}>{c.cliente_nombre || c.cliente_razonsocial || "—"}</td>
                        <td style={{ padding: "4px 6px", textAlign: "right", whiteSpace: "nowrap" }}>
                          <button onClick={() => editar(c)} title="Editar"
                            style={{ padding: "3px 7px", borderRadius: 4, border: "1px solid #d4d4d4", background: "#fff", cursor: "pointer", fontSize: 11, marginRight: 4 }}>
                            <BtnContent icon={Edit3} iconColor="#475569" />
                          </button>
                          <button onClick={() => setConfirmBorrar({ id: c.id, nombre: c.nombre })} title="Borrar"
                            style={{ padding: "3px 7px", borderRadius: 4, border: "1px solid #fca5a5", background: "#fef2f2", cursor: "pointer", fontSize: 11 }}>
                            <BtnContent icon={Trash2} iconColor="#dc2626" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>

          {/* Formulario */}
          {draft && (
            <div style={{ border: "1px solid #e5e5e5", borderRadius: 8, padding: 14, background: "#fff" }}>
              <h3 style={{ margin: 0, marginBottom: 10, paddingBottom: 8, borderBottom: "1px solid #e5e5e5", fontSize: 14, fontWeight: 700, color: "#171717" }}>
                {seleccionado === "nuevo" ? "Nuevo contacto" : `Editar: ${draft.nombre}`}
              </h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                <div>
                  <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 3 }}>Nombre *</label>
                  <input value={draft.nombre || ""} onChange={e => setDraft({ ...draft, nombre: e.target.value })}
                    style={{ width: "100%", padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 3 }}>Cargo</label>
                  <input value={draft.cargo || ""} onChange={e => setDraft({ ...draft, cargo: e.target.value })}
                    style={{ width: "100%", padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 3 }}>Email</label>
                  <input type="email" value={draft.email || ""} onChange={e => setDraft({ ...draft, email: e.target.value })}
                    style={{ width: "100%", padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 3 }}>Teléfono 1</label>
                  <input value={draft.telefono1 || ""} onChange={e => setDraft({ ...draft, telefono1: e.target.value })}
                    style={{ width: "100%", padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 3 }}>Teléfono 2</label>
                  <input value={draft.telefono2 || ""} onChange={e => setDraft({ ...draft, telefono2: e.target.value })}
                    style={{ width: "100%", padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12, boxSizing: "border-box" }} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 3 }}>Cliente</label>
                  <select value={draft.idcliente || 1} onChange={e => setDraft({ ...draft, idcliente: parseInt(e.target.value) })}
                    style={{ width: "100%", padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12, background: "#fff", boxSizing: "border-box" }}>
                    {clientes.map(c => (
                      <option key={c.id} value={c.id}>{c.nombrecomun || c.razonsocial || `Cliente ${c.id}`}</option>
                    ))}
                  </select>
                </div>
              </div>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 14, paddingTop: 10, borderTop: "1px solid #e5e5e5" }}>
                <button onClick={cancelar}
                  style={{ padding: "6px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", cursor: "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={guardar}
                  style={{ padding: "6px 18px", borderRadius: 6, border: "1px solid #16a34a", background: "#dcfce7", color: "#14532d", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={Save} iconColor="#14532d">{seleccionado === "nuevo" ? "Crear" : "Guardar"}</BtnContent>
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Confirmar borrado */}
        {confirmBorrar && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }} onClick={() => setConfirmBorrar(null)}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", maxWidth: 420 }} onClick={e => e.stopPropagation()}>
              <h3 style={{ margin: 0, marginBottom: 12, fontSize: 15, fontWeight: 700 }}>¿Borrar contacto?</h3>
              <p style={{ fontSize: 13, color: "#525252", marginBottom: 16 }}>¿Seguro que quieres borrar <strong>{confirmBorrar.nombre}</strong>?</p>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmBorrar(null)} style={{ padding: "6px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", cursor: "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={() => borrar(confirmBorrar.id, confirmBorrar.nombre)} style={{ padding: "6px 14px", borderRadius: 6, border: "1px solid #dc2626", background: "#fef2f2", color: "#991b1b", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={Trash2} iconColor="#dc2626">Sí, borrar</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Diálogo selector de cliente ──
function SelectorClienteDialog({ onClose, onSelect, setStatus }) {
  const [clientes, setClientes] = useState([]);
  const [busqueda, setBusqueda] = useState("");
  const [cargando, setCargando] = useState(false);

  useEffect(() => {
    setCargando(true);
    fetch(`${API_URL}/clientes/`)
      .then(r => r.ok ? r.json() : [])
      .then(setClientes)
      .catch(() => setClientes([]))
      .finally(() => setCargando(false));
  }, []);

  const filtrados = clientes.filter(c => {
    const q = busqueda.toLowerCase().trim();
    if (!q) return true;
    return (
      String(c.nombrecomun || "").toLowerCase().includes(q) ||
      String(c.razonsocial || "").toLowerCase().includes(q) ||
      String(c.ifa || "").toLowerCase().includes(q) ||
      String(c.poblacion || "").toLowerCase().includes(q)
    );
  });

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 820, maxHeight: "85vh", overflow: "hidden", boxShadow: "0 20px 60px rgba(0,0,0,0.2)", display: "flex", flexDirection: "column" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, paddingBottom: 12, borderBottom: "1px solid #e5e5e5" }}>
          <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={User} size={18} color="#171717" /> Seleccionar cliente
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        <div style={{ marginBottom: 12 }}>
          <input value={busqueda} onChange={e => setBusqueda(e.target.value)}
            placeholder="Buscar por nombre, razón social, NIF, población..."
            autoFocus
            style={{ width: "100%", padding: "6px 10px", border: "1px solid #d4d4d4", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} />
        </div>

        <div style={{ flex: 1, overflowY: "auto", border: "1px solid #e5e5e5", borderRadius: 8 }}>
          {cargando ? (
            <div style={{ padding: 20, textAlign: "center", color: "#737373", fontSize: 12 }}>Cargando...</div>
          ) : filtrados.length === 0 ? (
            <div style={{ padding: 20, textAlign: "center", color: "#737373", fontSize: 12 }}>Sin resultados</div>
          ) : (
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead style={{ position: "sticky", top: 0, background: "#fafafa" }}>
                <tr style={{ color: "#171717", fontSize: 11 }}>
                  <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Nombre común</th>
                  <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Razón social</th>
                  <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>NIF</th>
                  <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Población</th>
                </tr>
              </thead>
              <tbody>
                {filtrados.map(c => (
                  <tr key={c.id}
                    onDoubleClick={() => onSelect(c)}
                    onClick={() => onSelect(c)}
                    style={{ borderBottom: "1px solid #f1f5f9", cursor: "pointer" }}
                    onMouseEnter={e => e.currentTarget.style.background = "#f1f5f9"}
                    onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                    <td style={{ padding: "5px 8px", color: "#171717", fontWeight: 500 }}>{c.nombrecomun || "—"}</td>
                    <td style={{ padding: "5px 8px", color: "#475569" }}>{c.razonsocial || "—"}</td>
                    <td style={{ padding: "5px 8px", color: "#525252", fontFamily: "monospace" }}>{c.ifa || "—"}</td>
                    <td style={{ padding: "5px 8px", color: "#525252" }}>{c.poblacion || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
        <div style={{ marginTop: 10, fontSize: 11, color: "#94a3b8", textAlign: "right" }}>
          {filtrados.length} cliente{filtrados.length !== 1 ? "s" : ""} · doble clic o un clic para seleccionar
        </div>
      </div>
    </div>
  );
}

// ── Diálogo selector de contacto (para "A la atención de") ──
function SelectorContactoDialog({ onClose, onSelect, setStatus }) {
  const [contactos, setContactos] = useState([]);
  const [clientes, setClientes] = useState([]);
  const [filtroClienteId, setFiltroClienteId] = useState(null);
  const [busqueda, setBusqueda] = useState("");
  const [cargando, setCargando] = useState(false);

  useEffect(() => {
    setCargando(true);
    const params = filtroClienteId ? `?idcliente=${filtroClienteId}` : "";
    fetch(`${API_URL}/contactos/${params}`)
      .then(r => r.ok ? r.json() : [])
      .then(setContactos)
      .catch(() => setContactos([]))
      .finally(() => setCargando(false));
  }, [filtroClienteId]);

  useEffect(() => {
    fetch(`${API_URL}/clientes/`).then(r => r.ok ? r.json() : []).then(setClientes).catch(() => setClientes([]));
  }, []);

  const filtrados = contactos.filter(c =>
    !busqueda.trim() ||
    String(c.nombre || "").toLowerCase().includes(busqueda.toLowerCase()) ||
    String(c.cargo || "").toLowerCase().includes(busqueda.toLowerCase()) ||
    String(c.email || "").toLowerCase().includes(busqueda.toLowerCase())
  );

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 760, maxHeight: "85vh", overflow: "hidden", boxShadow: "0 20px 60px rgba(0,0,0,0.2)", display: "flex", flexDirection: "column" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, paddingBottom: 12, borderBottom: "1px solid #e5e5e5" }}>
          <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={User} size={18} color="#171717" /> Seleccionar contacto
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
          <label style={{ fontSize: 12, color: "#525252", fontWeight: 500 }}>Cliente:</label>
          <select value={filtroClienteId || ""} onChange={e => setFiltroClienteId(e.target.value ? parseInt(e.target.value) : null)}
            style={{ padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 6, fontSize: 12, background: "#fff", minWidth: 200 }}>
            <option value="">— Todos —</option>
            {clientes.map(c => (
              <option key={c.id} value={c.id}>{c.nombrecomun || c.razonsocial || `Cliente ${c.id}`}</option>
            ))}
          </select>
          <input value={busqueda} onChange={e => setBusqueda(e.target.value)}
            placeholder="Buscar..."
            style={{ flex: 1, padding: "5px 10px", border: "1px solid #d4d4d4", borderRadius: 6, fontSize: 12, minWidth: 180 }} />
        </div>

        <div style={{ flex: 1, overflowY: "auto", border: "1px solid #e5e5e5", borderRadius: 8 }}>
          {cargando ? (
            <div style={{ padding: 20, textAlign: "center", color: "#737373", fontSize: 12 }}>Cargando...</div>
          ) : filtrados.length === 0 ? (
            <div style={{ padding: 20, textAlign: "center", color: "#737373", fontSize: 12 }}>Sin resultados</div>
          ) : (
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead style={{ position: "sticky", top: 0, background: "#fafafa" }}>
                <tr style={{ color: "#171717", fontSize: 11 }}>
                  <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Nombre</th>
                  <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Cargo</th>
                  <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5", fontWeight: 600 }}>Cliente</th>
                </tr>
              </thead>
              <tbody>
                {filtrados.map(c => (
                  <tr key={c.id}
                    onDoubleClick={() => onSelect(c)}
                    onClick={() => onSelect(c)}
                    style={{ borderBottom: "1px solid #f1f5f9", cursor: "pointer" }}
                    onMouseEnter={e => e.currentTarget.style.background = "#eff6ff"}
                    onMouseLeave={e => e.currentTarget.style.background = ""}>
                    <td style={{ padding: "6px 8px", fontWeight: 500, color: "#171717" }}>{c.nombre}</td>
                    <td style={{ padding: "6px 8px", color: "#525252" }}>{c.cargo || "—"}</td>
                    <td style={{ padding: "6px 8px", color: "#737373", fontSize: 11 }}>{c.cliente_nombre || c.cliente_razonsocial || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 12, paddingTop: 10, borderTop: "1px solid #e5e5e5" }}>
          <button onClick={onClose}
            style={{ padding: "6px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", cursor: "pointer", fontSize: 12 }}>
            <BtnContent icon={X}>Cancelar</BtnContent>
          </button>
        </div>
      </div>
    </div>
  );
}

function ClientesDialog({ onClose, setStatus, onAsignarPresupuesto }) {
  const [clientes, setClientes] = useState([]);
  const [original, setOriginal] = useState([]); // snapshot para detectar cambios
  const [busqueda, setBusqueda] = useState("");
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState(null);
  const [seleccionado, setSeleccionado] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [confirmGuardar, setConfirmGuardar] = useState(false);
  const [guardando, setGuardando] = useState(false);
  const [provincias, setProvincias] = useState([]); // [{id, provincia, acronimo, cp}]
  const [confirmBorrarCliente, setConfirmBorrarCliente] = useState(null); // {id, nombre} | null
  const [borrandoCliente, setBorrandoCliente] = useState(false);
  const [comprobandoUsoCliente, setComprobandoUsoCliente] = useState(false);

  // ID temporal para clientes nuevos (negativos para distinguir)
  const nextTempId = useRef(-1);

  useEffect(() => { cargar(); }, []);

  // Cargar provincias para el desplegable del campo idprovincia
  useEffect(() => {
    fetch(`${API_URL}/provincias/`)
      .then(r => r.ok ? r.json() : [])
      .then(data => { if (Array.isArray(data)) setProvincias(data); })
      .catch(() => {});
  }, []);

  // Helper: nombre de provincia a partir de su id
  const nombreProvincia = (id) => {
    if (id == null || id === "") return "";
    const p = provincias.find(pv => String(pv.id) === String(id));
    return p ? p.provincia : "";
  };

  const cargar = async () => {
    setCargando(true); setError(null);
    setStatus && setStatus("Cargando clientes desde la base de datos...", "working");
    try {
      const res = await fetch(`${API_URL}/clientes/`);
      if (!res.ok) throw new Error("Error " + res.status);
      const data = await res.json();
      setClientes(data);
      setOriginal(JSON.parse(JSON.stringify(data)));
      setStatus && setStatus(`${data.length} cliente${data.length !== 1 ? "s" : ""} cargado${data.length !== 1 ? "s" : ""}`, "success");
    } catch (e) {
      setError(e.message);
      setStatus && setStatus("Error cargando clientes: " + e.message, "error");
    } finally {
      setCargando(false);
    }
  };

  // Filtrar
  const filtrados = clientes.filter(c => {
    if (!busqueda.trim()) return true;
    const q = busqueda.toLowerCase();
    return Object.values(c).some(v => String(v ?? "").toLowerCase().includes(q));
  });

  // Detectar cambios
  const cambios = (() => {
    const result = [];
    clientes.forEach(c => {
      // Cliente nuevo (id negativo) → siempre se incluye
      if (c.id < 0) { result.push(c); return; }
      const orig = original.find(o => o.id === c.id);
      if (!orig) return;
      const cambiado = Object.keys(c).some(k => String(c[k] ?? "") !== String(orig[k] ?? ""));
      if (cambiado) result.push(c);
    });
    return result;
  })();

  // Edición de celda
  const iniciarEdicion = (id, key) => {
    const col = CLIENTE_COLS.find(c => c.key === key);
    if (col?.readonly) return;
    const cliente = clientes.find(c => c.id === id);
    if (col?.type === "provincia") {
      // Precargar el NOMBRE de la provincia (no el id) para editar/buscar
      setEditingCell({ id, key });
      setEditValue(nombreProvincia(cliente?.[key]));
    } else if (key === "cp") {
      // CP: precargar con 5 cifras (ceros a la izquierda) para editar
      const n = String(cliente?.[key] ?? "").trim();
      setEditingCell({ id, key });
      setEditValue(n !== "" ? n.padStart(5, "0") : "");
    } else {
      setEditingCell({ id, key });
      setEditValue(String(cliente?.[key] ?? ""));
    }
  };

  const guardarCelda = () => {
    if (!editingCell) return;
    const col = CLIENTE_COLS.find(c => c.key === editingCell.key);
    let valor = editValue;
    if (col?.type === "provincia") {
      // editValue contiene el nombre escrito/elegido; buscamos el id correspondiente
      const txt = String(valor).trim().toLowerCase();
      if (!txt) {
        valor = null;
      } else {
        // coincidencia exacta primero, luego "empieza por"
        let prov = provincias.find(p => String(p.provincia).toLowerCase() === txt);
        if (!prov) prov = provincias.find(p => String(p.provincia).toLowerCase().startsWith(txt));
        valor = prov ? prov.id : null;
      }
    } else if (col?.type === "number") {
      const n = parseFloat(String(valor).replace(",", "."));
      valor = isNaN(n) ? null : Math.round(n);
    } else if (valor === "") {
      valor = null;
    }
    setClientes(cs => cs.map(c => c.id === editingCell.id ? { ...c, [editingCell.key]: valor } : c));
    setEditingCell(null);
  };

  const copiarCliente = () => {
    if (!seleccionado) return;
    const src = clientes.find(c => c.id === seleccionado);
    if (!src) return;
    const copia = { ...src, id: nextTempId.current--, razonsocial: (src.razonsocial || "") + " (copia)" };
    setClientes(cs => [copia, ...cs]);
    setSeleccionado(copia.id);
  };

  const comprobarYBorrarCliente = async () => {
    if (!seleccionado) return;
    const cli = clientes.find(c => c.id === seleccionado);
    if (!cli) return;
    const nombre = cli.nombrecomun || cli.razonsocial || `Cliente ${cli.id}`;
    // Si es un cliente nuevo sin guardar (id negativo), se elimina de la lista local sin tocar BD
    if (cli.id < 0) {
      setClientes(cs => cs.filter(c => c.id !== cli.id));
      setSeleccionado(null);
      setStatus && setStatus(`Cliente nuevo "${nombre}" descartado`, "info");
      return;
    }
    setComprobandoUsoCliente(true);
    setStatus && setStatus(`Comprobando uso de "${nombre}"...`, "working");
    try {
      const res = await fetch(`${API_URL}/clientes/${cli.id}/uso`);
      if (res.ok) {
        const data = await res.json();
        if (!data.borrable) {
          setStatus && setStatus(`No se puede borrar "${nombre}": está referenciado en ${data.en_presupuestos || 0} presupuesto(s) y ${data.en_contactos || 0} contacto(s).`, "error");
          return;
        }
      }
      // Si el endpoint de uso no existe (404), seguimos igualmente y dejamos que el DELETE decida
      setConfirmBorrarCliente({ id: cli.id, nombre });
    } catch (e) {
      // Si falla la comprobación, permitimos intentar el borrado igualmente
      setConfirmBorrarCliente({ id: cli.id, nombre });
    } finally {
      setComprobandoUsoCliente(false);
    }
  };

  const borrarClienteConfirmado = async () => {
    if (!confirmBorrarCliente) return;
    setBorrandoCliente(true);
    try {
      const res = await fetch(`${API_URL}/clientes/${confirmBorrarCliente.id}`, { method: "DELETE" });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "HTTP " + res.status);
      }
      setStatus && setStatus(`Cliente "${confirmBorrarCliente.nombre}" borrado`, "success");
      setConfirmBorrarCliente(null);
      setSeleccionado(null);
      await cargar();
    } catch (e) {
      setStatus && setStatus("Error borrando cliente: " + e.message, "error");
    } finally {
      setBorrandoCliente(false);
    }
  };

  const guardarTodo = async () => {
    setGuardando(true); setError(null);
    setStatus && setStatus(`Guardando ${cambios.length} cambio${cambios.length > 1 ? "s" : ""} en la base de datos...`, "working");
    try {
      const res = await fetch(`${API_URL}/clientes/batch-update`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(cambios),
      });
      if (!res.ok) throw new Error("Error " + res.status);
      await res.json();
      setStatus && setStatus(`${cambios.length} cambio${cambios.length > 1 ? "s" : ""} guardado${cambios.length > 1 ? "s" : ""} correctamente`, "success");
      // Limpiar estados de selección/edición: los ids temporales ya no existen tras recargar
      setSeleccionado(null);
      setEditingCell(null);
      setEditValue("");
      setConfirmGuardar(false);
      await cargar();
    } catch (e) {
      setError(e.message);
      setStatus && setStatus("Error guardando cambios: " + e.message, "error");
    } finally {
      setGuardando(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 1280, height: "88vh", maxHeight: "94vh", minWidth: 520, minHeight: 340, display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.2)", resize: "both", overflow: "auto" }} onClick={e => e.stopPropagation()}>

        {/* Datalist único para autocompletar provincias */}
        <datalist id="lista-provincias">
          {provincias.map(p => <option key={p.id} value={p.provincia} />)}
        </datalist>

        {/* Cabecera */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, borderBottom: "1px solid #e2e8f0", paddingBottom: 12, flexShrink: 0 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}><Icon as={Users} size={18} color="#1e3a5f" /> Gestión de clientes</h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}><Icon as={X} size={18} /></button>
        </div>

        {/* Toolbar */}
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, flexShrink: 0, flexWrap: "wrap" }}>
          <input type="text" value={busqueda} onChange={e => setBusqueda(e.target.value)}
            placeholder="Buscar en cualquier columna..."
            style={{ flex: 1, minWidth: 200, padding: "6px 10px", fontSize: 12, borderRadius: 6, border: "1px solid #cbd5e1", background: "#f8fafc" }} />
          <button onClick={copiarCliente} disabled={!seleccionado}
            style={{ padding: "6px 14px", fontSize: 12, borderRadius: 6, border: "1px solid #cbd5e1", background: seleccionado ? "#fff" : "#fafafa", color: seleccionado ? "#1e293b" : "#cbd5e1", cursor: seleccionado ? "pointer" : "default" }}>
            <BtnContent icon={Copy}>Copiar cliente</BtnContent>
          </button>
          <button onClick={comprobarYBorrarCliente} disabled={!seleccionado || comprobandoUsoCliente}
            style={{ padding: "6px 14px", fontSize: 12, borderRadius: 6, border: "1px solid #fecaca", background: seleccionado && !comprobandoUsoCliente ? "#fff" : "#fafafa", color: seleccionado && !comprobandoUsoCliente ? "#dc2626" : "#cbd5e1", cursor: seleccionado && !comprobandoUsoCliente ? "pointer" : "default" }}>
            <BtnContent icon={Trash2} iconColor={seleccionado && !comprobandoUsoCliente ? "#dc2626" : "#cbd5e1"}>Borrar cliente</BtnContent>
          </button>
          <button onClick={cargar} style={{ padding: "6px 12px", fontSize: 12, borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", cursor: "pointer" }}><BtnContent icon={RefreshCw}>Refrescar</BtnContent></button>
          <div style={{ flex: 1 }} />
          <button onClick={() => {
              if (!seleccionado) return;
              const cliente = clientes.find(c => c.id === seleccionado);
              if (!cliente) return;
              onAsignarPresupuesto && onAsignarPresupuesto(cliente);
              setStatus && setStatus(`Cliente "${cliente.nombrecomun || cliente.razonsocial}" asignado al presupuesto`, "success");
              onClose();
            }}
            disabled={!seleccionado}
            title="Asigna el cliente seleccionado al presupuesto actual"
            style={{
              padding: "5px 12px",
              fontSize: 11,
              lineHeight: 1.15,
              borderRadius: 6,
              border: "1px solid " + (seleccionado ? "#16a34a" : "#cbd5e1"),
              background: seleccionado ? "#dcfce7" : "#fafafa",
              color: seleccionado ? "#14532d" : "#cbd5e1",
              cursor: seleccionado ? "pointer" : "default",
              fontWeight: 600,
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              whiteSpace: "normal",
              textAlign: "left",
              maxWidth: 220,
            }}>
            <Icon as={Check} size={16} color={seleccionado ? "#16a34a" : "#cbd5e1"} />
            <span>Asignar Cliente a<br />Presupuesto actual</span>
          </button>
          {cambios.length > 0 && (
            <span style={{ fontSize: 11, padding: "3px 10px", borderRadius: 99, background: "#fef3c7", color: "#92400e", border: "1px solid #fcd34d" }}>
              ● {cambios.length} cambio{cambios.length > 1 ? "s" : ""} sin guardar
            </span>
          )}
        </div>

        {/* Error */}
        {error && (
          <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#991b1b", marginBottom: 10, flexShrink: 0 }}>
            {error}
          </div>
        )}

        {/* Tabla */}
        <div style={{ flex: 1, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, marginBottom: 12, minHeight: 250 }}>
          <table style={{ borderCollapse: "collapse", fontSize: 12, width: "max-content", minWidth: "100%" }}>
            <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 2, borderBottom: "1px solid #e5e5e5" }}>
              <tr>
                <th style={{ width: 28, padding: "8px 6px", borderBottom: "1px solid #e5e5e5", borderRight: "1px solid #e5e5e5" }}></th>
                {CLIENTE_COLS.map(col => (
                  <th key={col.key} style={{ width: col.width, minWidth: col.width, maxWidth: col.width, padding: "8px 10px", textAlign: col.type === "number" || col.key === "id" ? "right" : "left", color: "#171717", fontWeight: 600, borderBottom: "1px solid #e5e5e5", borderRight: "1px solid #e5e5e5" }}>
                    {col.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {cargando && (
                <tr><td colSpan={CLIENTE_COLS.length + 1} style={{ padding: "30px", textAlign: "center", color: "#64748b" }}>⏳ Cargando...</td></tr>
              )}
              {!cargando && filtrados.length === 0 && (
                <tr><td colSpan={CLIENTE_COLS.length + 1} style={{ padding: "30px", textAlign: "center", color: "#94a3b8" }}>
                  {busqueda ? "No hay clientes que coincidan." : "No hay clientes."}
                </td></tr>
              )}
              {!cargando && filtrados.map((c, i) => {
                const isSel = seleccionado === c.id;
                const orig = original.find(o => o.id === c.id);
                const isNuevo = c.id < 0;
                const isModificado = orig && Object.keys(c).some(k => String(c[k] ?? "") !== String(orig[k] ?? ""));
                return (
                  <tr key={c.id}
                    style={{ background: isSel ? "#dbeafe" : (isNuevo ? "#ecfdf5" : (isModificado ? "#fef3c7" : (i % 2 === 0 ? "#fff" : "#f8fafc"))), borderBottom: "1px solid #f1f5f9" }}>
                    <td style={{ padding: "5px 6px", textAlign: "center", border: "1px solid #e2e8f0" }}>
                      <input type="radio" checked={isSel} onChange={() => setSeleccionado(c.id)} />
                    </td>
                    {CLIENTE_COLS.map(col => {
                      const isEditing = editingCell?.id === c.id && editingCell?.key === col.key;
                      const valor = c[col.key];
                      let display;
                      if (isNuevo && col.key === "id") display = "NUEVO";
                      else if (col.type === "provincia") display = nombreProvincia(valor);
                      else if (col.key === "cp") {
                        // Código postal: siempre 5 cifras, rellenando con ceros a la izquierda (1250 → 01250)
                        const n = String(valor ?? "").trim();
                        display = n !== "" ? n.padStart(5, "0") : "";
                      }
                      else display = (valor ?? "");
                      return (
                        <td key={col.key}
                          onDoubleClick={() => !col.readonly && iniciarEdicion(c.id, col.key)}
                          style={{ padding: 0, width: col.width, minWidth: col.width, maxWidth: col.width, border: "1px solid #e2e8f0", cursor: col.readonly ? "default" : "cell", background: col.readonly ? "#f8fafc" : "inherit" }}>
                          {isEditing ? (
                            col.type === "provincia" ? (
                              <input autoFocus value={editValue}
                                list="lista-provincias"
                                placeholder="Escribe para buscar..."
                                onChange={e => setEditValue(e.target.value)}
                                onBlur={guardarCelda}
                                onKeyDown={e => { if (e.key === "Enter") guardarCelda(); if (e.key === "Escape") setEditingCell(null); }}
                                style={{ width: "100%", border: "none", outline: "none", padding: "5px 8px", fontSize: 12, background: "#fff" }} />
                            ) : (
                              <input autoFocus value={editValue}
                                onChange={e => setEditValue(e.target.value)}
                                onBlur={guardarCelda}
                                onKeyDown={e => { if (e.key === "Enter") guardarCelda(); if (e.key === "Escape") setEditingCell(null); }}
                                style={{ width: "100%", border: "none", outline: "none", padding: "5px 8px", fontSize: 12, background: "#fff", textAlign: col.type === "number" ? "right" : "left" }} />
                            )
                          ) : (
                            <div style={{ padding: "5px 8px", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", textAlign: col.type === "number" || col.key === "id" ? "right" : "left", color: col.readonly ? (isNuevo ? "#16a34a" : "#94a3b8") : "#1e293b", fontWeight: col.key === "id" ? 600 : 400 }}>
                              {display}
                            </div>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Leyenda */}
        <div style={{ display: "flex", gap: 16, fontSize: 11, color: "#64748b", marginBottom: 8, flexShrink: 0 }}>
          <span><span style={{ display: "inline-block", width: 10, height: 10, background: "#fef3c7", border: "1px solid #fcd34d", marginRight: 4, verticalAlign: "middle" }} /> Modificado</span>
          <span><span style={{ display: "inline-block", width: 10, height: 10, background: "#ecfdf5", border: "1px solid #86efac", marginRight: 4, verticalAlign: "middle" }} /> Nuevo (sin guardar)</span>
          <span style={{ marginLeft: "auto" }}>Doble click en una celda para editar. {filtrados.length} cliente{filtrados.length !== 1 ? "s" : ""}.</span>
        </div>

        {/* Botones */}
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", paddingTop: 10, borderTop: "1px solid #e2e8f0", flexShrink: 0 }}>
          <button onClick={onClose} style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 12 }}><BtnContent icon={X}>Cerrar</BtnContent></button>
          <button onClick={() => setConfirmGuardar(true)} disabled={cambios.length === 0 || guardando}
            style={{ padding: "7px 20px", borderRadius: 6, border: "none", background: cambios.length > 0 && !guardando ? "#16a34a" : "#cbd5e1", color: "#fff", cursor: cambios.length > 0 && !guardando ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={Save} iconColor="#fff">Guardar cambios{cambios.length > 0 ? ` (${cambios.length})` : ""}</BtnContent>
          </button>
        </div>

        {/* Confirmación de guardado */}
        {confirmGuardar && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", maxWidth: 460, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                <Icon as={Save} size={24} color="#16a34a" />
                <h3 style={{ fontSize: 15, fontWeight: 700, color: "#1e3a5f", margin: 0 }}>Confirmar guardado</h3>
              </div>
              <p style={{ fontSize: 13, color: "#475569", lineHeight: 1.6, margin: "0 0 16px" }}>
                Se aplicarán <strong>{cambios.length}</strong> cambio{cambios.length > 1 ? "s" : ""} en la base de datos:
              </p>
              <ul style={{ fontSize: 12, color: "#475569", marginBottom: 16, paddingLeft: 18, maxHeight: 200, overflow: "auto" }}>
                {cambios.map(c => (
                  <li key={c.id}>
                    {c.id < 0 ? <strong style={{ color: "#16a34a" }}>NUEVO: </strong> : <strong>ID {c.id}: </strong>}
                    {c.razonsocial || "(sin razón social)"} — {c.nombrecomun || ""}
                  </li>
                ))}
              </ul>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmGuardar(false)} disabled={guardando} style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}><BtnContent icon={X}>Cancelar</BtnContent></button>
                <button onClick={guardarTodo} disabled={guardando}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "none", background: "#16a34a", color: "#fff", cursor: guardando ? "default" : "pointer", fontSize: 13, fontWeight: 600 }}>
                  <BtnContent icon={guardando ? RefreshCw : Save} iconColor="#fff">{guardando ? "Guardando..." : "Sí, guardar"}</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}

        {confirmBorrarCliente && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", maxWidth: 440, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                <Icon as={Trash2} size={24} color="#dc2626" />
                <h3 style={{ fontSize: 15, fontWeight: 700, color: "#991b1b", margin: 0 }}>Borrar cliente</h3>
              </div>
              <p style={{ fontSize: 13, color: "#475569", lineHeight: 1.6, margin: "0 0 16px" }}>
                ¿Seguro que quieres borrar el cliente <strong>{confirmBorrarCliente.nombre}</strong>? Esta acción no se puede deshacer.
              </p>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmBorrarCliente(null)} disabled={borrandoCliente} style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}><BtnContent icon={X}>Cancelar</BtnContent></button>
                <button onClick={borrarClienteConfirmado} disabled={borrandoCliente}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "none", background: "#dc2626", color: "#fff", cursor: borrandoCliente ? "default" : "pointer", fontSize: 13, fontWeight: 600 }}>
                  <BtnContent icon={borrandoCliente ? RefreshCw : Trash2} iconColor="#fff">{borrandoCliente ? "Borrando..." : "Sí, borrar"}</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}

// ── Diálogo Seleccionar Filas ──
function SeleccionarCeldasDialog({ onClose, onAceptar, totalFilas }) {
  const [desdeFila, setDesdeFila] = useState(1);
  const [hastaFila, setHastaFila] = useState(totalFilas);
  const [error, setError] = useState(null);

  const aceptar = () => {
    setError(null);
    const fIni = parseInt(desdeFila);
    const fFin = parseInt(hastaFila);
    if (isNaN(fIni) || isNaN(fFin) || fIni < 1 || fFin < 1) {
      setError("Indica filas válidas (mínimo 1)");
      return;
    }
    if (fIni > totalFilas || fFin > totalFilas) {
      setError(`El presupuesto solo tiene ${totalFilas} filas`);
      return;
    }
    const min = Math.min(fIni, fFin);
    const max = Math.max(fIni, fFin);
    onAceptar({ desde: min, hasta: max });
    onClose();
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", minWidth: 420, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, borderBottom: "1px solid #e2e8f0", paddingBottom: 12 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Square} size={18} color="#1e3a5f" /> Seleccionar filas
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        <p style={{ fontSize: 12, color: "#64748b", margin: "0 0 16px" }}>
          Indica el rango de filas a marcar. Se seleccionarán los checkbox de todas las filas dentro del rango.
        </p>

        <div style={{ display: "flex", gap: 12, alignItems: "center", marginBottom: 16 }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 4, flex: 1 }}>
            <label style={{ fontSize: 11, fontWeight: 600, color: "#1e3a5f" }}>Desde la fila</label>
            <input type="number" min="1" max={totalFilas} value={desdeFila} onChange={e => setDesdeFila(e.target.value)} autoFocus
              style={{ padding: "6px 10px", fontSize: 13, borderRadius: 4, border: "1px solid #cbd5e1", textAlign: "right", background: "#f8fafc" }} />
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4, flex: 1 }}>
            <label style={{ fontSize: 11, fontWeight: 600, color: "#1e3a5f" }}>Hasta la fila</label>
            <input type="number" min="1" max={totalFilas} value={hastaFila} onChange={e => setHastaFila(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter") aceptar(); }}
              style={{ padding: "6px 10px", fontSize: 13, borderRadius: 4, border: "1px solid #cbd5e1", textAlign: "right", background: "#f8fafc" }} />
          </div>
        </div>

        <div style={{ fontSize: 11, color: "#64748b", marginBottom: 12, padding: "6px 10px", background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 4 }}>
          El presupuesto tiene <strong>{totalFilas} fila{totalFilas !== 1 ? "s" : ""}</strong>.
        </div>

        {error && (
          <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#991b1b", marginBottom: 12 }}>
            {error}
          </div>
        )}

        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", paddingTop: 12, borderTop: "1px solid #e2e8f0" }}>
          <button onClick={onClose}
            style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 12 }}>
            <BtnContent icon={X}>Cancelar</BtnContent>
          </button>
          <button onClick={aceptar}
            style={{ padding: "7px 20px", borderRadius: 6, border: "none", background: "#16a34a", color: "#fff", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={Check} iconColor="#fff">Aceptar selección</BtnContent>
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Componente AutocompleteReferencia ──
// Input con sugerencias del backend al teclear la referencia
function AutocompleteReferencia({ value, onChange, onConfirm, onCancel, align }) {
  const [text, setText] = useState(value || "");
  const [suggestions, setSuggestions] = useState([]);
  const [open, setOpen] = useState(false);
  const [highlighted, setHighlighted] = useState(0);
  const [loading, setLoading] = useState(false);
  const skipBlur = useRef(false);
  const fetchSeq = useRef(0);

  useEffect(() => {
    const txt = String(text || "").trim();
    if (txt.length < 2) { setSuggestions([]); setOpen(false); return; }
    setLoading(true);
    const seq = ++fetchSeq.current;
    const t = setTimeout(async () => {
      try {
        const url = API_URL + "/productos/?busqueda=" + encodeURIComponent(txt) + "&campo=referencia&limite=5";
        const res = await fetch(url);
        if (!res.ok) throw new Error("HTTP " + res.status);
        const data = await res.json();
        if (seq !== fetchSeq.current) return;
        const txtUp = txt.toUpperCase();
        const ordenados = [...data].sort((a, b) => {
          const aStarts = String(a.referencia || "").toUpperCase().startsWith(txtUp) ? 0 : 1;
          const bStarts = String(b.referencia || "").toUpperCase().startsWith(txtUp) ? 0 : 1;
          if (aStarts !== bStarts) return aStarts - bStarts;
          return String(a.referencia || "").localeCompare(String(b.referencia || ""));
        });
        setSuggestions(ordenados.slice(0, 5));
        setOpen(ordenados.length > 0);
        setHighlighted(0);
      } catch {
        setSuggestions([]);
        setOpen(false);
      } finally {
        if (seq === fetchSeq.current) setLoading(false);
      }
    }, 250);
    return () => clearTimeout(t);
  }, [text]);

  const elegir = (s) => {
    skipBlur.current = true;
    setText(s.referencia);
    setOpen(false);
    onConfirm && onConfirm(s.referencia);
  };

  return (
    <div style={{ position: "relative", width: "100%" }}>
      <input autoFocus value={text}
        onChange={e => { setText(e.target.value); onChange && onChange(e.target.value); }}
        onBlur={() => {
          if (skipBlur.current) { skipBlur.current = false; return; }
          setTimeout(() => onConfirm && onConfirm(text), 120);
        }}
        onKeyDown={e => {
          if (e.key === "ArrowDown" && open) { e.preventDefault(); setHighlighted(h => Math.min(h + 1, suggestions.length - 1)); }
          else if (e.key === "ArrowUp" && open) { e.preventDefault(); setHighlighted(h => Math.max(h - 1, 0)); }
          else if (e.key === "Enter") {
            e.preventDefault();
            if (open && suggestions[highlighted]) elegir(suggestions[highlighted]);
            else onConfirm && onConfirm(text);
          }
          else if (e.key === "Escape") { setOpen(false); onCancel && onCancel(); }
          else if (e.key === "Tab") {
            if (open && suggestions[highlighted]) elegir(suggestions[highlighted]);
            else onConfirm && onConfirm(text);
          }
        }}
        style={{ width: "100%", border: "none", outline: "none", padding: "5px 8px", fontSize: 12, background: "#fff", textAlign: align || "left", fontFamily: "monospace" }}
      />
      {open && suggestions.length > 0 && (
        <div style={{ position: "absolute", top: "100%", left: 0, zIndex: 9999, background: "#fff", border: "1px solid #d4d4d4", borderRadius: 4, boxShadow: "0 4px 12px rgba(0,0,0,0.12)", minWidth: 380, maxWidth: 520, marginTop: 2 }}>
          {loading && <div style={{ padding: "6px 10px", fontSize: 11, color: "#737373" }}>Buscando...</div>}
          {suggestions.map((s, i) => (
            <div key={s.id}
              onMouseDown={() => elegir(s)}
              onMouseEnter={() => setHighlighted(i)}
              style={{
                padding: "5px 10px",
                fontSize: 11,
                background: i === highlighted ? "#f5f5f5" : "#fff",
                cursor: "pointer",
                borderBottom: i < suggestions.length - 1 ? "1px solid #f1f5f9" : "none",
                display: "grid",
                gridTemplateColumns: "auto 1fr auto",
                gap: 8,
                alignItems: "center",
              }}>
              <span style={{ fontFamily: "monospace", fontWeight: 600, color: "#1e3a5f", whiteSpace: "nowrap" }}>{s.referencia}</span>
              <span style={{ color: "#525252", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }} title={s.nombre || s.descripcion}>{s.nombre || s.descripcion || ""}</span>
              <span style={{ color: "#0369a1", fontWeight: 500, whiteSpace: "nowrap" }}>{(Number(s.pvp) || 0).toLocaleString("es-ES", { minimumFractionDigits: 2 })} €</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Diálogo Leer Producto ──
function LeerProductoDialog({ onClose, onInsertar, setStatus }) {
  const [busqueda, setBusqueda] = useState("");
  const [campo, setCampo] = useState("referencia"); // "referencia" o "todos"
  const [productos, setProductos] = useState([]);
  const [hayMas, setHayMas] = useState(false); // true si el servidor pudo devolver el límite máximo
  const [seleccionado, setSeleccionado] = useState(null);
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState(null);
  const [copiado, setCopiado] = useState(false);
  const [confirmBorrar, setConfirmBorrar] = useState(false);
  const [borrando, setBorrando] = useState(false);
  const [comprobandoUso, setComprobandoUso] = useState(false);

  // Anchos de columna redimensionables (clave → px). Por defecto los iniciales.
  const ANCHOS_INI = { masusado: 70, referencia: 150, pvp: 90, grupodescuento: 140, descripcion: 320, id: 60 };
  const [anchosCol, setAnchosCol] = useState(ANCHOS_INI);
  const resizeRef = useRef(null);
  const onResizeCol = (e, key) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startW = anchosCol[key] ?? 100;
    resizeRef.current = { key, startX, startW };
    const onMove = (ev) => {
      if (!resizeRef.current) return;
      const { key: k, startX: sx, startW: sw } = resizeRef.current;
      const nuevo = Math.max(40, sw + (ev.clientX - sx));
      setAnchosCol(prev => ({ ...prev, [k]: nuevo }));
    };
    const onUp = () => {
      resizeRef.current = null;
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";
  };
  // Componente del tirador de redimensión en el borde derecho de una cabecera
  const Resizer = ({ colKey }) => (
    <div
      onMouseDown={(e) => onResizeCol(e, colKey)}
      onDoubleClick={(e) => { e.stopPropagation(); setAnchosCol(prev => ({ ...prev, [colKey]: ANCHOS_INI[colKey] })); }}
      title="Arrastra para redimensionar (doble clic para restablecer)"
      onMouseEnter={e => { e.currentTarget.style.background = "#2563eb"; }}
      onMouseLeave={e => { e.currentTarget.style.background = "transparent"; }}
      style={{ position: "absolute", top: 0, right: -4, width: 9, height: "100%", cursor: "col-resize", zIndex: 6, transition: "background 0.1s" }}
    />
  );

  useEffect(() => {
    const t = setTimeout(() => cargarLista(), 300);
    return () => clearTimeout(t);
  }, [busqueda, campo]);

  const cargarLista = async () => {
    setCargando(true);
    setError(null);
    try {
      const LIMITE = 200;
      const params = busqueda.trim() ? `busqueda=${encodeURIComponent(busqueda.trim())}&campo=${campo}&limite=${LIMITE}` : `limite=${LIMITE}`;
      const url = `${API_URL}/productos/?${params}`;
      const res = await fetch(url);
      if (!res.ok) throw new Error("Error " + res.status);
      const data = await res.json();
      // Detectar si el servidor pudo estar capado por el límite (ya filtrado por backend)
      setHayMas(data.length >= LIMITE);
      // Ordenar por referencia ascendente
      data.sort((a, b) => String(a.referencia || "").localeCompare(String(b.referencia || "")));
      setProductos(data);
    } catch (e) {
      setError(e.message);
      setProductos([]);
    } finally {
      setCargando(false);
    }
  };

  const copiarReferencia = async () => {
    if (!seleccionado) return;
    try {
      await navigator.clipboard.writeText(seleccionado.referencia || "");
      setCopiado(true);
      setStatus && setStatus(`Referencia ${seleccionado.referencia} copiada al portapapeles`, "success");
      setTimeout(() => setCopiado(false), 1500);
    } catch (e) {
      setStatus && setStatus("No se pudo copiar la referencia", "error");
    }
  };

  const insertar = () => {
    if (!seleccionado) return;
    onInsertar(seleccionado);
    setStatus && setStatus(`Producto ${seleccionado.referencia} insertado`, "success");
    onClose();
  };

  const borrarProducto = async () => {
    if (!seleccionado) return;
    setBorrando(true);
    setStatus && setStatus(`Borrando producto "${seleccionado.referencia}"...`, "working");
    try {
      const res = await fetch(`${API_URL}/productos/${seleccionado.id}`, { method: "DELETE" });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "HTTP " + res.status);
      }
      setStatus && setStatus(`Producto "${seleccionado.referencia}" borrado de la base de datos`, "success");
      setConfirmBorrar(false);
      setSeleccionado(null);
      cargarLista();
    } catch (e) {
      setStatus && setStatus("Error borrando producto: " + e.message, "error");
    } finally {
      setBorrando(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 1100, height: "85vh", maxHeight: "92vh", minWidth: 480, minHeight: 320, display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.2)", resize: "both", overflow: "auto" }} onClick={e => e.stopPropagation()}>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, borderBottom: "1px solid #e2e8f0", paddingBottom: 12, flexShrink: 0 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Package} size={18} color="#1e3a5f" /> Leer producto
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        {/* Filtro */}
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, flexShrink: 0 }}>
          <input type="text" value={busqueda} onChange={e => setBusqueda(e.target.value)}
            placeholder={campo === "referencia" ? "Buscar por referencia..." : "Buscar en referencia, nombre y descripción..."} autoFocus
            style={{ flex: 1, padding: "6px 10px", fontSize: 12, borderRadius: 6, border: "1px solid #cbd5e1", background: "#f8fafc" }} />
          <div style={{ display: "inline-flex", border: "1px solid #d4d4d4", borderRadius: 6, overflow: "hidden" }}>
            <button onClick={() => setCampo("referencia")}
              style={{ padding: "5px 12px", fontSize: 11, border: "none", cursor: "pointer",
                background: campo === "referencia" ? "#171717" : "#fff",
                color: campo === "referencia" ? "#fff" : "#171717",
                fontWeight: campo === "referencia" ? 600 : 400 }}>
              Solo referencia
            </button>
            <button onClick={() => setCampo("todos")}
              style={{ padding: "5px 12px", fontSize: 11, border: "none", borderLeft: "1px solid #d4d4d4", cursor: "pointer",
                background: campo === "todos" ? "#171717" : "#fff",
                color: campo === "todos" ? "#fff" : "#171717",
                fontWeight: campo === "todos" ? 600 : 400 }}>
              Todos los campos
            </button>
          </div>
          <span style={{ fontSize: 11, color: hayMas ? "#d97706" : "#64748b", display: "inline-flex", alignItems: "center", gap: 4, whiteSpace: "nowrap" }}>
            {hayMas && <Icon as={HelpCircle} size={12} color="#d97706" />}
            {productos.length} resultado{productos.length !== 1 ? "s" : ""}
            {hayMas && " (límite alcanzado, hay más)"}
          </span>
        </div>

        {/* Referencia + descripción del seleccionado */}
        <div style={{ display: "grid", gridTemplateColumns: "auto 1fr", gap: 10, marginBottom: 12, padding: "10px 12px", background: seleccionado ? "#f0f9ff" : "#f8fafc", border: "1px solid " + (seleccionado ? "#bae6fd" : "#e2e8f0"), borderRadius: 6, flexShrink: 0 }}>
          {/* Referencia con botón copiar */}
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 11, fontWeight: 600, color: "#0369a1" }}>Referencia</label>
            <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
              <input type="text" readOnly value={seleccionado?.referencia || ""}
                style={{ minWidth: 180, padding: "5px 8px", fontSize: 12, fontFamily: "monospace", fontWeight: 600, color: "#1e3a5f", border: "1px solid #cbd5e1", borderRadius: 4, background: "#fff" }} />
              <button onClick={copiarReferencia} disabled={!seleccionado}
                title="Copiar al portapapeles"
                style={{ padding: "5px 8px", border: "1px solid " + (copiado ? "#16a34a" : "#cbd5e1"), borderRadius: 4, background: copiado ? "#dcfce7" : "#fff", cursor: seleccionado ? "pointer" : "default", color: seleccionado ? "#1e3a5f" : "#cbd5e1" }}>
                <Icon as={copiado ? Check : Copy} size={14} color={copiado ? "#16a34a" : seleccionado ? "#475569" : "#cbd5e1"} />
              </button>
            </div>
          </div>
          {/* Descripción completa */}
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 11, fontWeight: 600, color: "#0369a1" }}>Descripción completa</label>
            <textarea readOnly value={seleccionado?.descripcion || ""}
              style={{ width: "100%", padding: "5px 8px", fontSize: 11, color: "#475569", border: "1px solid #cbd5e1", borderRadius: 4, background: "#fff", minHeight: 50, maxHeight: 100, resize: "vertical", fontFamily: "inherit", lineHeight: 1.4 }} />
          </div>
        </div>

        {error && (
          <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#991b1b", marginBottom: 10, flexShrink: 0 }}>
            {error}
          </div>
        )}

        {/* Tabla */}
        <div style={{ flex: 1, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, marginBottom: 12, minHeight: 250 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed" }}>
            <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 1, borderBottom: "1px solid #e5e5e5" }}>
              <tr>
                <th style={{ width: 32, padding: "8px 6px", color: "#171717" }}></th>
                <th style={{ position: "relative", width: anchosCol.masusado, padding: "8px 10px", textAlign: "center", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }} title="Más usado">Más Us.<Resizer colKey="masusado" /></th>
                <th style={{ position: "relative", width: anchosCol.referencia, padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>Referencia<Resizer colKey="referencia" /></th>
                <th style={{ position: "relative", width: anchosCol.pvp, padding: "8px 10px", textAlign: "right", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>PVP<Resizer colKey="pvp" /></th>
                <th style={{ position: "relative", width: anchosCol.grupodescuento, padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>Grupo Dto.<Resizer colKey="grupodescuento" /></th>
                <th style={{ position: "relative", width: anchosCol.descripcion, padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Descripción<Resizer colKey="descripcion" /></th>
                <th style={{ position: "relative", width: anchosCol.id, padding: "8px 10px", textAlign: "right", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>ID</th>
              </tr>
            </thead>
            <tbody>
              {cargando && (
                <tr><td colSpan={7} style={{ padding: "30px", textAlign: "center", color: "#64748b" }}>Cargando...</td></tr>
              )}
              {!cargando && productos.length === 0 && !error && (
                <tr><td colSpan={7} style={{ padding: "30px", textAlign: "center", color: "#94a3b8" }}>
                  {busqueda ? "No hay productos que coincidan." : "No hay productos en la base de datos."}
                </td></tr>
              )}
              {!cargando && productos.map(p => {
                const isSel = seleccionado?.id === p.id;
                return (
                  <tr key={p.id}
                    onClick={() => setSeleccionado(p)}
                    onDoubleClick={() => { setSeleccionado(p); insertar(); }}
                    style={{ background: isSel ? "#dbeafe" : "transparent", cursor: "pointer", borderBottom: "1px solid #f1f5f9" }}>
                    <td style={{ padding: "5px 6px", textAlign: "center" }}>
                      <input type="radio" checked={isSel} onChange={() => setSeleccionado(p)} />
                    </td>
                    <td style={{ padding: "5px 10px", textAlign: "center", color: "#475569", fontWeight: 500 }}>{p.masusado ?? 0}</td>
                    <td style={{ padding: "5px 10px", fontWeight: 600, color: "#1e3a5f", fontFamily: "monospace", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }} title={p.referencia}>{p.referencia}</td>
                    <td style={{ padding: "5px 10px", textAlign: "right", color: "#0369a1", fontWeight: 500, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                      {(Number(p.pvp) || 0).toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} €
                    </td>
                    <td style={{ padding: "5px 10px", color: "#475569", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}
                      title={p.descripciongrupodescuento || p.grupodescuento || ""}>
                      {p.grupodescuento || ""}
                    </td>
                    <td style={{ padding: "5px 10px", color: "#475569", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}
                      title={p.descripcion || ""}>
                      {p.descripcion || p.nombre || ""}
                    </td>
                    <td style={{ padding: "5px 10px", textAlign: "right", color: "#94a3b8", fontFamily: "monospace", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.id}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        <div style={{ display: "flex", gap: 8, justifyContent: "space-between", alignItems: "center", paddingTop: 10, borderTop: "1px solid #e2e8f0", flexShrink: 0 }}>
          <button onClick={async () => {
              if (!seleccionado) return;
              setComprobandoUso(true);
              setStatus && setStatus(`Comprobando referencias del producto...`, "working");
              try {
                const res = await fetch(`${API_URL}/productos/${seleccionado.id}/uso`);
                if (!res.ok) throw new Error("HTTP " + res.status);
                const data = await res.json();
                if (!data.borrable) {
                  setStatus && setStatus(`No se puede borrar "${seleccionado.referencia}": está referenciado en ${data.en_presupuestos} línea(s) de presupuesto y ${data.en_elementos} elemento(s).`, "error");
                  return;
                }
                setStatus && setStatus("Producto sin referencias, listo para borrar", "info");
                setConfirmBorrar(true);
              } catch (e) {
                setStatus && setStatus("Error comprobando uso: " + e.message, "error");
              } finally {
                setComprobandoUso(false);
              }
            }} disabled={!seleccionado || comprobandoUso}
            style={{ padding: "7px 14px", borderRadius: 6, border: (seleccionado && !comprobandoUso) ? "1px solid #fca5a5" : "1px solid #d4d4d4", background: (seleccionado && !comprobandoUso) ? "#fef2f2" : "#f5f5f5", color: (seleccionado && !comprobandoUso) ? "#991b1b" : "#737373", cursor: (seleccionado && !comprobandoUso) ? "pointer" : "default", fontSize: 12 }}>
            <BtnContent icon={comprobandoUso ? RefreshCw : Trash2} iconColor={(seleccionado && !comprobandoUso) ? "#dc2626" : "#737373"}>{comprobandoUso ? "Comprobando..." : "Borrar Producto"}</BtnContent>
          </button>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={onClose} style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 12 }}>
              <BtnContent icon={X}>Cancelar</BtnContent>
            </button>
            <button onClick={insertar} disabled={!seleccionado}
              style={{ padding: "7px 20px", borderRadius: 6, border: "none", background: seleccionado ? "#16a34a" : "#cbd5e1", color: "#fff", cursor: seleccionado ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
              <BtnContent icon={Check} iconColor="#fff">Insertar producto</BtnContent>
            </button>
          </div>
        </div>

        {/* Diálogo confirmar borrar producto */}
        {confirmBorrar && seleccionado && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100001 }} onClick={() => !borrando && setConfirmBorrar(false)}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 470, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
              <h3 style={{ margin: 0, marginBottom: 12, fontSize: 15, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
                <Icon as={Trash2} size={18} color="#dc2626" /> ¿Borrar producto?
              </h3>
              <p style={{ fontSize: 13, color: "#525252", marginBottom: 8, lineHeight: 1.5 }}>
                Vas a borrar el producto <strong style={{ color: "#171717", fontFamily: "monospace" }}>{seleccionado.referencia}</strong> de la base de datos.
              </p>
              <div style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#1e293b", marginBottom: 10 }}>
                <div><strong>Nombre:</strong> {seleccionado.nombre || "—"}</div>
                {seleccionado.descripcion && <div><strong>Descripción:</strong> {String(seleccionado.descripcion).substring(0, 100)}{seleccionado.descripcion.length > 100 ? "..." : ""}</div>}
              </div>
              <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "10px 12px", fontSize: 12, color: "#991b1b", marginBottom: 16 }}>
                <strong>Atención:</strong> esta acción borrará el producto de la base de datos. No se puede deshacer.
              </div>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmBorrar(false)} disabled={borrando}
                  style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: borrando ? "default" : "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={borrarProducto} disabled={borrando}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #dc2626", background: "#fef2f2", color: "#991b1b", cursor: borrando ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={borrando ? RefreshCw : Trash2} iconColor="#dc2626">{borrando ? "Borrando..." : "Sí, borrar"}</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Diálogo Guardar Elemento ──
function GuardarElementoDialog({ filasSeleccionadas, onClose, setStatus }) {
  const [nombre, setNombre] = useState("");
  const [descripcion, setDescripcion] = useState("");
  const [guardando, setGuardando] = useState(false);
  const [error, setError] = useState(null);
  const [resultado, setResultado] = useState(null); // datos devueltos por el backend tras guardar

  // Solo productos válidos para guardar
  const productos = filasSeleccionadas.filter(f =>
    ["PD","PE","E"].includes(f.naturaleza) && f.referencia
  );

  const guardar = async () => {
    setError(null);
    if (!nombre.trim()) {
      setError("Indica un nombre para el elemento");
      return;
    }
    if (productos.length === 0) {
      setError("No hay filas con referencia válida entre las seleccionadas");
      return;
    }
    setGuardando(true);
    setStatus && setStatus(`Guardando elemento "${nombre}" con ${productos.length} producto${productos.length > 1 ? "s" : ""}...`, "working");
    try {
      const res = await fetch(`${API_URL}/elementos/`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          nombre: nombre.trim(),
          descripcion: descripcion.trim(),
          idgrupo: 1,
          productos: productos.map(p => ({ referencia: p.referencia, cantidadproducto: p.cantidad || 1 })),
        }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || `Error ${res.status}`);
      }
      const data = await res.json();
      setResultado(data);
      setStatus && setStatus(`Elemento "${nombre}" guardado con ${data.productos_insertados} productos`, "success");
    } catch (e) {
      setError(e.message);
      setStatus && setStatus("Error guardando elemento: " + e.message, "error");
    } finally {
      setGuardando(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "90%", maxWidth: 540, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, borderBottom: "1px solid #e2e8f0", paddingBottom: 12 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Save} size={18} color="#1e3a5f" /> Guardar elemento
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        {!resultado ? (<>
          <div style={{ background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 4, padding: "8px 12px", marginBottom: 14, fontSize: 12, color: "#0369a1" }}>
            <strong>{productos.length}</strong> línea{productos.length !== 1 ? "s" : ""} válida{productos.length !== 1 ? "s" : ""} (productos + comentarios) de las {filasSeleccionadas.length} fila{filasSeleccionadas.length !== 1 ? "s" : ""} seleccionada{filasSeleccionadas.length !== 1 ? "s" : ""}.
          </div>

          <div style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", display: "block", marginBottom: 6 }}>Nombre del elemento *</label>
            <input value={nombre} onChange={e => setNombre(e.target.value)} autoFocus
              onKeyDown={e => { if (e.key === "Enter") guardar(); }}
              placeholder="Ej: Conjunto variador G120 55kW"
              style={{ width: "100%", padding: "7px 10px", fontSize: 13, borderRadius: 4, border: "1px solid #cbd5e1", background: "#f8fafc" }} />
          </div>

          <div style={{ marginBottom: 14 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", display: "block", marginBottom: 6 }}>Descripción (opcional)</label>
            <textarea value={descripcion} onChange={e => setDescripcion(e.target.value)}
              placeholder="Notas o descripción del elemento..."
              style={{ width: "100%", padding: "7px 10px", fontSize: 12, borderRadius: 4, border: "1px solid #cbd5e1", background: "#f8fafc", minHeight: 60, resize: "vertical", fontFamily: "inherit" }} />
          </div>

          {/* Vista previa de productos */}
          <div style={{ marginBottom: 14 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#1e3a5f", marginBottom: 6 }}>Productos a guardar:</div>
            <div style={{ maxHeight: 150, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 4, fontSize: 11 }}>
              {productos.length === 0 ? (
                <div style={{ padding: "12px", textAlign: "center", color: "#94a3b8" }}>No hay filas con referencia válida</div>
              ) : productos.map((p, i) => (
                <div key={i} style={{ padding: "5px 10px", display: "flex", gap: 12, borderBottom: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#f8fafc" }}>
                  <span style={{ fontWeight: 500, color: "#1e3a5f", minWidth: 80, textAlign: "center" }}>{p.cantidad || 1} ×</span>
                  <strong style={{ color: "#1e3a5f", minWidth: 150 }}>{p.referencia}</strong>
                  <span style={{ color: "#475569", flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.nombre}</span>
                </div>
              ))}
            </div>
          </div>

          {error && (
            <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#991b1b", marginBottom: 12 }}>
              {error}
            </div>
          )}

          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", paddingTop: 12, borderTop: "1px solid #e2e8f0" }}>
            <button onClick={onClose} style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 12 }}>
              <BtnContent icon={X}>Cancelar</BtnContent>
            </button>
            <button onClick={guardar} disabled={guardando || !nombre.trim() || productos.length === 0}
              style={{ padding: "7px 20px", borderRadius: 6, border: "none", background: (!guardando && nombre.trim() && productos.length > 0) ? "#16a34a" : "#cbd5e1", color: "#fff", cursor: (!guardando && nombre.trim() && productos.length > 0) ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
              <BtnContent icon={guardando ? RefreshCw : Save} iconColor="#fff">{guardando ? "Guardando..." : "Guardar"}</BtnContent>
            </button>
          </div>
        </>) : (<>
          {/* Resultado del guardado */}
          <div style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 6, padding: "12px 16px", marginBottom: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
              <Icon as={Check} size={18} color="#16a34a" />
              <strong style={{ color: "#14532d", fontSize: 14 }}>Elemento guardado correctamente</strong>
            </div>
            <div style={{ fontSize: 12, color: "#15803d" }}>
              ID: <strong>{resultado.id}</strong> — Nombre: <strong>{resultado.nombre}</strong><br />
              Productos insertados: <strong>{resultado.productos_insertados}</strong>
            </div>
          </div>

          {resultado.referencias_no_encontradas && resultado.referencias_no_encontradas.length > 0 && (
            <div style={{ background: "#fef3c7", border: "1px solid #fcd34d", borderRadius: 6, padding: "10px 14px", fontSize: 12, color: "#92400e", marginBottom: 14 }}>
              <strong>⚠ Atención:</strong> las siguientes referencias no se encontraron en la BD y no se guardaron:
              <ul style={{ margin: "6px 0 0 16px", padding: 0 }}>
                {resultado.referencias_no_encontradas.map((r, i) => <li key={i}><code>{r}</code></li>)}
              </ul>
            </div>
          )}

          <div style={{ display: "flex", justifyContent: "flex-end", paddingTop: 12, borderTop: "1px solid #e2e8f0" }}>
            <button onClick={onClose}
              style={{ padding: "7px 20px", borderRadius: 6, border: "none", background: "#2563eb", color: "#fff", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
              <BtnContent icon={Check} iconColor="#fff">Cerrar</BtnContent>
            </button>
          </div>
        </>)}
      </div>
    </div>
  );
}

// ── Diálogo Leer Elemento ──
function LeerElementoDialog({ onClose, onInsertar, setStatus }) {
  const [busqueda, setBusqueda] = useState("");
  const [elementos, setElementos] = useState([]);
  const [grupos, setGrupos] = useState([]); // [{id, grupo}]
  const [selIds, setSelIds] = useState(new Set()); // Set de ids seleccionados
  const [ultimoSelIdx, setUltimoSelIdx] = useState(null); // para shift+click
  const [cargando, setCargando] = useState(false);
  const [insertando, setInsertando] = useState(false);
  const [confirmBorrar, setConfirmBorrar] = useState(false);
  const [borrando, setBorrando] = useState(false);
  const [confirmGrupo, setConfirmGrupo] = useState(false);
  const [grupoNuevoId, setGrupoNuevoId] = useState(null);
  const [aplicandoGrupo, setAplicandoGrupo] = useState(false);
  const [editar, setEditar] = useState(null); // {nombre, descripcion} para editar
  const [guardandoEditar, setGuardandoEditar] = useState(false);
  const [error, setError] = useState(null);
  const [tab, setTab] = useState("lista");
  const [detalle, setDetalle] = useState(null);
  const [cargandoDetalle, setCargandoDetalle] = useState(false);

  // Elemento único seleccionado (si hay solo 1) - útil para acciones que requieren uno solo
  const seleccionado = selIds.size === 1
    ? elementos.find(e => e.id === [...selIds][0]) || null
    : null;
  const cantSel = selIds.size;

  // Cargar el detalle del elemento al cambiar a la pestaña detalle
  useEffect(() => {
    if (tab !== "detalle" || !seleccionado) { setDetalle(null); return; }
    setCargandoDetalle(true);
    setDetalle(null);
    fetch(`${API_URL}/elementos/${seleccionado.id}`)
      .then(r => r.ok ? r.json() : Promise.reject(new Error("Error " + r.status)))
      .then(d => setDetalle(d))
      .catch(e => setStatus && setStatus("Error cargando detalle del elemento: " + e.message, "error"))
      .finally(() => setCargandoDetalle(false));
  }, [tab, seleccionado?.id]);

  useEffect(() => {
    const t = setTimeout(() => cargarLista(), 300);
    return () => clearTimeout(t);
  }, [busqueda]);

  // Cargar grupos de elementos una vez al montar
  useEffect(() => {
    fetch(`${API_URL}/gruposelementos/`)
      .then(r => r.ok ? r.json() : [])
      .then(setGrupos)
      .catch(() => setGrupos([]));
  }, []);

  const grupoNombre = (idgrupo) => {
    if (!idgrupo) return "—";
    const g = grupos.find(x => x.id === idgrupo);
    return g ? g.grupo : `#${idgrupo}`;
  };

  const cargarLista = async () => {
    setCargando(true);
    setError(null);
    try {
      const url = busqueda.trim()
        ? `${API_URL}/elementos/?busqueda=${encodeURIComponent(busqueda.trim())}`
        : `${API_URL}/elementos/`;
      const res = await fetch(url);
      if (!res.ok) throw new Error("Error " + res.status);
      setElementos(await res.json());
    } catch (e) {
      setError(e.message);
      setElementos([]);
    } finally {
      setCargando(false);
    }
  };

  // Selección
  const toggleSel = (el, evt, idx) => {
    setSelIds(prev => {
      const next = new Set(prev);
      if (evt.shiftKey && ultimoSelIdx !== null && idx !== null) {
        // Rango entre el último y el actual
        const ini = Math.min(ultimoSelIdx, idx), fin = Math.max(ultimoSelIdx, idx);
        for (let i = ini; i <= fin; i++) {
          if (elementos[i]) next.add(elementos[i].id);
        }
      } else if (evt.ctrlKey || evt.metaKey) {
        // Toggle individual
        if (next.has(el.id)) next.delete(el.id); else next.add(el.id);
      } else {
        // Click normal: reemplaza la selección por solo este elemento
        next.clear();
        next.add(el.id);
      }
      return next;
    });
    setUltimoSelIdx(idx);
  };

  const insertar = async () => {
    if (cantSel !== 1 || !seleccionado) return;
    setInsertando(true);
    setStatus && setStatus(`Cargando elemento "${seleccionado.nombre}"...`, "working");
    try {
      const res = await fetch(`${API_URL}/elementos/${seleccionado.id}`);
      if (!res.ok) throw new Error("Error " + res.status);
      const data = await res.json();
      onInsertar(data);
      setStatus && setStatus(`Elemento "${seleccionado.nombre}" insertado con ${data.productos.length} productos`, "success");
      onClose();
    } catch (e) {
      setError(e.message);
      setStatus && setStatus("Error insertando elemento: " + e.message, "error");
    } finally {
      setInsertando(false);
    }
  };

  const borrarElementos = async () => {
    if (selIds.size === 0) return;
    setBorrando(true);
    const ids = [...selIds];
    setStatus && setStatus(`Borrando ${ids.length} elemento${ids.length !== 1 ? "s" : ""}...`, "working");
    let ok = 0, errs = 0;
    for (const id of ids) {
      try {
        const res = await fetch(`${API_URL}/elementos/${id}`, { method: "DELETE" });
        if (!res.ok) throw new Error("HTTP " + res.status);
        ok++;
      } catch (e) {
        errs++;
        console.error("Error borrando elemento", id, e);
      }
    }
    setStatus && setStatus(`Borrado: ${ok} OK, ${errs} con error`, errs > 0 ? "error" : "success");
    setConfirmBorrar(false);
    setSelIds(new Set());
    setBorrando(false);
    cargarLista();
  };

  const aplicarGrupo = async () => {
    if (selIds.size === 0 || grupoNuevoId == null) return;
    setAplicandoGrupo(true);
    const ids = [...selIds];
    setStatus && setStatus(`Cambiando grupo a ${ids.length} elemento${ids.length !== 1 ? "s" : ""}...`, "working");
    let ok = 0, errs = 0;
    for (const id of ids) {
      try {
        const res = await fetch(`${API_URL}/elementos/${id}`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ idgrupo: grupoNuevoId }),
        });
        if (!res.ok) throw new Error("HTTP " + res.status);
        ok++;
      } catch (e) {
        errs++;
        console.error("Error cambiando grupo de elemento", id, e);
      }
    }
    setStatus && setStatus(`Grupo cambiado: ${ok} OK, ${errs} con error`, errs > 0 ? "error" : "success");
    setConfirmGrupo(false);
    setGrupoNuevoId(null);
    setAplicandoGrupo(false);
    cargarLista();
  };

  const guardarEdicion = async () => {
    if (!editar || cantSel !== 1) return;
    setGuardandoEditar(true);
    setStatus && setStatus(`Guardando cambios...`, "working");
    try {
      const payload = {};
      if (editar.nombre !== undefined) payload.nombre = editar.nombre;
      if (editar.descripcion !== undefined) payload.descripcion = editar.descripcion;
      const res = await fetch(`${API_URL}/elementos/${seleccionado.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "HTTP " + res.status);
      }
      setStatus && setStatus(`Elemento actualizado`, "success");
      setEditar(null);
      cargarLista();
    } catch (e) {
      setStatus && setStatus("Error guardando: " + e.message, "error");
    } finally {
      setGuardandoEditar(false);
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "90%", maxWidth: 1000, height: "85vh", maxHeight: "92vh", minWidth: 480, minHeight: 320, display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.2)", resize: "both", overflow: "auto" }} onClick={e => e.stopPropagation()}>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, borderBottom: "1px solid #e2e8f0", paddingBottom: 12, flexShrink: 0 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Download} size={18} color="#1e3a5f" /> Leer Elemento
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        {/* Pestañas */}
        <div style={{ display: "flex", borderBottom: "1px solid #e5e5e5", marginBottom: 12, flexShrink: 0 }}>
          <button onClick={() => setTab("lista")}
            style={{ padding: "8px 18px", border: "none", background: "transparent", fontSize: 12, fontWeight: 600, cursor: "pointer",
              color: tab === "lista" ? "#171717" : "#737373",
              borderBottom: tab === "lista" ? "2px solid #171717" : "2px solid transparent",
              marginBottom: "-1px" }}>
            <BtnContent icon={FolderOpen} iconColor={tab === "lista" ? "#171717" : "#737373"}>Lista de elementos</BtnContent>
          </button>
          <button onClick={() => setTab("detalle")} disabled={cantSel !== 1}
            style={{ padding: "8px 18px", border: "none", background: "transparent", fontSize: 12, fontWeight: 600, cursor: cantSel === 1 ? "pointer" : "default",
              color: cantSel !== 1 ? "#cbd5e1" : tab === "detalle" ? "#171717" : "#737373",
              borderBottom: tab === "detalle" ? "2px solid #171717" : "2px solid transparent",
              marginBottom: "-1px" }}>
            <BtnContent icon={Package} iconColor={cantSel !== 1 ? "#cbd5e1" : tab === "detalle" ? "#171717" : "#737373"}>
              Detalle del elemento {cantSel === 1 && seleccionado ? `(${seleccionado.nombre})` : ""}
            </BtnContent>
          </button>
        </div>

        {/* TAB: Lista */}
        {tab === "lista" && <>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, flexShrink: 0 }}>
          <input type="text" value={busqueda} onChange={e => setBusqueda(e.target.value)}
            placeholder="Buscar por nombre o descripción..." autoFocus
            style={{ flex: 1, padding: "6px 10px", fontSize: 12, borderRadius: 6, border: "1px solid #cbd5e1", background: "#f8fafc" }} />
          <span style={{ fontSize: 11, color: "#64748b" }}>
            {elementos.length} resultado{elementos.length !== 1 ? "s" : ""}
            {cantSel > 0 && <> · <strong style={{ color: "#1e3a5f" }}>{cantSel} seleccionado{cantSel !== 1 ? "s" : ""}</strong></>}
          </span>
        </div>
        <div style={{ fontSize: 10, color: "#94a3b8", marginBottom: 6 }}>
          Click para seleccionar uno. <kbd style={{ fontFamily: "monospace", background: "#f1f5f9", padding: "1px 4px", borderRadius: 3, border: "1px solid #cbd5e1" }}>Ctrl+Click</kbd> para añadir/quitar. <kbd style={{ fontFamily: "monospace", background: "#f1f5f9", padding: "1px 4px", borderRadius: 3, border: "1px solid #cbd5e1" }}>Shift+Click</kbd> para rango.
        </div>

        {error && (
          <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#991b1b", marginBottom: 10, flexShrink: 0 }}>
            {error}
          </div>
        )}

        <div style={{ flex: 1, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6, marginBottom: 12, minHeight: 200 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 1, borderBottom: "1px solid #e5e5e5" }}>
              <tr>
                <th style={{ width: 32, padding: "8px 6px", color: "#171717" }}></th>
                <th style={{ padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600, width: "32%" }}>Nombre</th>
                <th style={{ padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Grupo Elementos</th>
                <th style={{ padding: "8px 10px", textAlign: "center", color: "#171717", fontWeight: 600, whiteSpace: "nowrap" }}>Fecha modificación</th>
                <th style={{ padding: "8px 10px", textAlign: "left", color: "#171717", fontWeight: 600, width: "20%" }}>Descripción</th>
                <th style={{ padding: "8px 10px", textAlign: "center", color: "#171717", fontWeight: 600 }}>Productos</th>
                <th style={{ padding: "8px 10px", textAlign: "right", color: "#171717", fontWeight: 600 }}>ID</th>
              </tr>
            </thead>
            <tbody>
              {cargando && (
                <tr><td colSpan={7} style={{ padding: "30px", textAlign: "center", color: "#64748b" }}>Cargando...</td></tr>
              )}
              {!cargando && elementos.length === 0 && !error && (
                <tr><td colSpan={7} style={{ padding: "30px", textAlign: "center", color: "#94a3b8" }}>
                  {busqueda ? "No hay elementos que coincidan." : "No hay elementos en la base de datos."}
                </td></tr>
              )}
              {!cargando && elementos.map((el, idx) => {
                const isSel = selIds.has(el.id);
                return (
                  <tr key={el.id}
                    onClick={(e) => toggleSel(el, e, idx)}
                    onDoubleClick={() => { setSelIds(new Set([el.id])); setTimeout(() => insertar(), 0); }}
                    style={{ background: isSel ? "#dbeafe" : "transparent", cursor: "pointer", borderBottom: "1px solid #f1f5f9", userSelect: "none" }}>
                    <td style={{ padding: "5px 6px", textAlign: "center" }}>
                      <input type="checkbox" checked={isSel} readOnly />
                    </td>
                    <td style={{ padding: "6px 10px", fontWeight: 500 }}>{el.nombre}</td>
                    <td style={{ padding: "6px 10px", color: "#475569" }}>{el.grupo || grupoNombre(el.idgrupo)}</td>
                    <td style={{ padding: "6px 10px", textAlign: "center", color: "#64748b", fontSize: 11, whiteSpace: "nowrap" }}>
                      {el.fechamodificacion ? new Date(el.fechamodificacion).toLocaleDateString("es-ES", { day: "2-digit", month: "2-digit", year: "numeric" }) : "—"}
                    </td>
                    <td style={{ padding: "6px 10px", color: "#475569" }}>{el.descripcion || "—"}</td>
                    <td style={{ padding: "6px 10px", textAlign: "center", fontWeight: 500, color: "#0369a1" }}>{el.numproductos || 0}</td>
                    <td style={{ padding: "6px 10px", textAlign: "right", fontWeight: 600, color: "#1e3a5f" }}>{el.id}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        </>}

        {/* TAB: Detalle */}
        {tab === "detalle" && (
          <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden", marginBottom: 12 }}>
            {cantSel !== 1 ? (
              <div style={{ padding: "40px", textAlign: "center", color: "#94a3b8", fontSize: 13 }}>
                Selecciona un único elemento para ver su detalle.
              </div>
            ) : cargandoDetalle ? (
              <div style={{ padding: "40px", textAlign: "center", color: "#64748b" }}>
                <Icon as={RefreshCw} size={18} color="#64748b" /> Cargando detalle...
              </div>
            ) : detalle ? (
              <>
                <div style={{ padding: "10px 14px", background: "#f8fafc", border: "1px solid #e5e5e5", borderRadius: 6, marginBottom: 10, display: "grid", gridTemplateColumns: "auto 1fr", gap: 10, fontSize: 12, columnGap: 20 }}>
                  <div><strong style={{ color: "#737373", fontSize: 10, textTransform: "uppercase" }}>Nombre</strong><br /><span style={{ color: "#171717", fontWeight: 600 }}>{detalle.cabecera.nombre}</span></div>
                  <div><strong style={{ color: "#737373", fontSize: 10, textTransform: "uppercase" }}>Descripción</strong><br /><span style={{ color: "#171717" }}>{detalle.cabecera.descripcion || "—"}</span></div>
                </div>

                <div style={{ flex: 1, overflow: "auto", border: "1px solid #e5e5e5", borderRadius: 6 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                    <thead style={{ position: "sticky", top: 0, background: "#fafafa", zIndex: 1, borderBottom: "1px solid #e5e5e5" }}>
                      <tr>
                        <th style={{ padding: "6px 8px", textAlign: "center", color: "#171717", fontWeight: 600 }}>Cant.</th>
                        <th style={{ padding: "6px 8px", textAlign: "left", color: "#171717", fontWeight: 600, whiteSpace: "nowrap", width: 115 }}>Referencia</th>
                        <th style={{ padding: "6px 8px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Nombre</th>
                        <th style={{ padding: "6px 8px", textAlign: "right", color: "#171717", fontWeight: 600 }}>PVP</th>
                        <th style={{ padding: "6px 8px", textAlign: "left", color: "#171717", fontWeight: 600 }}>Familia</th>
                        <th style={{ padding: "6px 8px", textAlign: "left", color: "#171717", fontWeight: 600 }}>SubFamilia</th>
                        <th style={{ padding: "6px 8px", textAlign: "right", color: "#171717", fontWeight: 600 }}>ID Producto</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(detalle.productos || []).length === 0 && (
                        <tr><td colSpan={7} style={{ padding: 30, textAlign: "center", color: "#94a3b8" }}>Este elemento no tiene productos.</td></tr>
                      )}
                      {(detalle.productos || []).map((p, i) => {
                        const esCom = p.es_comentario || p.idproducto === 118852;
                        if (esCom) {
                          return (
                            <tr key={"cm-" + i} style={{ background: "#fef3c7", borderBottom: "1px solid #fcd34d" }}>
                              <td colSpan={7} style={{ padding: "4px 12px", color: "#92400e", fontStyle: "italic" }}>
                                📝 <strong>Comentario:</strong> {p.texto || ""}
                              </td>
                            </tr>
                          );
                        }
                        return (
                          <tr key={p.idproducto || i} style={{ background: i % 2 === 0 ? "#fff" : "#fafafa", borderBottom: "1px solid #f1f5f9" }}>
                            <td style={{ padding: "4px 8px", textAlign: "center", color: "#475569" }}>{p.cantidad || 0}</td>
                            <td style={{ padding: "4px 8px", color: "#1e3a5f", fontFamily: "monospace", whiteSpace: "nowrap", fontWeight: 600 }}>{p.referencia || ""}</td>
                            <td style={{ padding: "4px 8px", color: "#171717" }}>{p.nombre || ""}</td>
                            <td style={{ padding: "4px 8px", textAlign: "right", color: "#0369a1" }}>{(Number(p.pvp) || 0).toLocaleString("es-ES", { minimumFractionDigits: 2 })}</td>
                            <td style={{ padding: "4px 8px", color: "#737373" }}>{p.familia || ""}</td>
                            <td style={{ padding: "4px 8px", color: "#737373" }}>{p.subfamilia || ""}</td>
                            <td style={{ padding: "4px 8px", textAlign: "right", color: "#94a3b8", fontFamily: "monospace" }}>{p.idproducto}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                <div style={{ fontSize: 11, color: "#64748b", marginTop: 6, textAlign: "right" }}>
                  {(detalle.productos || []).length} producto{(detalle.productos || []).length !== 1 ? "s" : ""}
                </div>
              </>
            ) : null}
          </div>
        )}

        <div style={{ display: "flex", gap: 6, justifyContent: "space-between", alignItems: "center", paddingTop: 10, borderTop: "1px solid #e2e8f0", flexShrink: 0, flexWrap: "wrap" }}>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            <button onClick={() => setConfirmBorrar(true)} disabled={cantSel === 0}
              style={{ padding: "7px 12px", borderRadius: 6, border: cantSel > 0 ? "1px solid #fca5a5" : "1px solid #d4d4d4", background: cantSel > 0 ? "#fef2f2" : "#f5f5f5", color: cantSel > 0 ? "#991b1b" : "#737373", cursor: cantSel > 0 ? "pointer" : "default", fontSize: 12 }}>
              <BtnContent icon={Trash2} iconColor={cantSel > 0 ? "#dc2626" : "#737373"}>Borrar Elemento{cantSel > 1 ? "s" : ""}</BtnContent>
            </button>
            <button onClick={() => { setGrupoNuevoId(null); setConfirmGrupo(true); }} disabled={cantSel === 0}
              style={{ padding: "7px 12px", borderRadius: 6, border: cantSel > 0 ? "1px solid #cbd5e1" : "1px solid #d4d4d4", background: cantSel > 0 ? "#fff" : "#f5f5f5", color: cantSel > 0 ? "#171717" : "#737373", cursor: cantSel > 0 ? "pointer" : "default", fontSize: 12 }}>
              <BtnContent icon={FolderOpen} iconColor={cantSel > 0 ? "#475569" : "#737373"}>Cambiar grupo</BtnContent>
            </button>
            <button onClick={() => setEditar({ nombre: seleccionado?.nombre || "", descripcion: seleccionado?.descripcion || "" })} disabled={cantSel !== 1}
              style={{ padding: "7px 12px", borderRadius: 6, border: cantSel === 1 ? "1px solid #cbd5e1" : "1px solid #d4d4d4", background: cantSel === 1 ? "#fff" : "#f5f5f5", color: cantSel === 1 ? "#171717" : "#737373", cursor: cantSel === 1 ? "pointer" : "default", fontSize: 12 }}>
              <BtnContent icon={Edit3} iconColor={cantSel === 1 ? "#475569" : "#737373"}>Cambiar nombre / descripción</BtnContent>
            </button>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={onClose}
              style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
              <BtnContent icon={X}>Cancelar</BtnContent>
            </button>
            <button onClick={insertar} disabled={cantSel !== 1 || insertando}
              style={{ padding: "7px 20px", borderRadius: 6, border: (cantSel === 1 && !insertando) ? "1px solid #16a34a" : "1px solid #d4d4d4", background: (cantSel === 1 && !insertando) ? "#dcfce7" : "#f5f5f5", color: (cantSel === 1 && !insertando) ? "#14532d" : "#737373", cursor: (cantSel === 1 && !insertando) ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
              <BtnContent icon={insertando ? RefreshCw : Check} iconColor={(cantSel === 1 && !insertando) ? "#14532d" : "#737373"}>{insertando ? "Leyendo..." : "Leer Elemento"}</BtnContent>
            </button>
          </div>
        </div>

        {/* Diálogo confirmar borrar elemento(s) */}
        {confirmBorrar && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }} onClick={() => !borrando && setConfirmBorrar(false)}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 470, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
              <h3 style={{ margin: 0, marginBottom: 12, fontSize: 15, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
                <Icon as={Trash2} size={18} color="#dc2626" /> ¿Borrar elemento{cantSel > 1 ? "s" : ""}?
              </h3>
              <p style={{ fontSize: 13, color: "#525252", marginBottom: 8, lineHeight: 1.5 }}>
                {cantSel === 1
                  ? <>Vas a borrar el elemento <strong style={{ color: "#171717" }}>"{seleccionado?.nombre}"</strong> de la base de datos.</>
                  : <>Vas a borrar <strong style={{ color: "#171717" }}>{cantSel} elementos</strong> de la base de datos.</>}
              </p>
              <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "10px 12px", fontSize: 12, color: "#991b1b", marginBottom: 16 }}>
                <strong>Atención:</strong> esta acción borra los elementos completos y todos sus productos asociados. No se puede deshacer.
              </div>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmBorrar(false)} disabled={borrando}
                  style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: borrando ? "default" : "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={borrarElementos} disabled={borrando}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #dc2626", background: "#fef2f2", color: "#991b1b", cursor: borrando ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={borrando ? RefreshCw : Trash2} iconColor="#dc2626">{borrando ? "Borrando..." : "Sí, borrar"}</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Diálogo cambiar grupo */}
        {confirmGrupo && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }} onClick={() => !aplicandoGrupo && setConfirmGrupo(false)}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 460, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
              <h3 style={{ margin: 0, marginBottom: 12, fontSize: 15, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
                <Icon as={FolderOpen} size={18} color="#1e3a5f" /> Cambiar grupo al Elemento
              </h3>
              <p style={{ fontSize: 13, color: "#525252", marginBottom: 12, lineHeight: 1.5 }}>
                Se asignará el grupo seleccionado a <strong style={{ color: "#171717" }}>{cantSel} elemento{cantSel !== 1 ? "s" : ""}</strong>.
              </p>
              <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 4 }}>Nuevo grupo:</label>
              <select value={grupoNuevoId ?? ""} onChange={e => setGrupoNuevoId(e.target.value ? parseInt(e.target.value) : null)}
                style={{ width: "100%", padding: "7px 10px", border: "1px solid #d4d4d4", borderRadius: 6, fontSize: 13, background: "#fff", marginBottom: 16 }}>
                <option value="">— Selecciona un grupo —</option>
                {grupos.map(g => (
                  <option key={g.id} value={g.id}>{g.grupo}</option>
                ))}
              </select>
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setConfirmGrupo(false)} disabled={aplicandoGrupo}
                  style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: aplicandoGrupo ? "default" : "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={aplicarGrupo} disabled={aplicandoGrupo || grupoNuevoId == null}
                  style={{ padding: "7px 16px", borderRadius: 6, border: (grupoNuevoId != null && !aplicandoGrupo) ? "1px solid #16a34a" : "1px solid #d4d4d4", background: (grupoNuevoId != null && !aplicandoGrupo) ? "#dcfce7" : "#f5f5f5", color: (grupoNuevoId != null && !aplicandoGrupo) ? "#14532d" : "#737373", cursor: (grupoNuevoId != null && !aplicandoGrupo) ? "pointer" : "default", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={aplicandoGrupo ? RefreshCw : Check} iconColor={(grupoNuevoId != null && !aplicandoGrupo) ? "#14532d" : "#737373"}>{aplicandoGrupo ? "Aplicando..." : "Aplicar grupo"}</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Diálogo editar nombre/descripcion */}
        {editar && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }} onClick={() => !guardandoEditar && setEditar(null)}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 520, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
              <h3 style={{ margin: 0, marginBottom: 12, fontSize: 15, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
                <Icon as={Edit3} size={18} color="#1e3a5f" /> Cambiar nombre / descripción
              </h3>
              <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 4 }}>Nombre</label>
              <input value={editar.nombre} onChange={e => setEditar({ ...editar, nombre: e.target.value })}
                style={{ width: "100%", padding: "7px 10px", border: "1px solid #d4d4d4", borderRadius: 6, fontSize: 13, marginBottom: 12, boxSizing: "border-box" }} />
              <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 4 }}>Descripción</label>
              <textarea value={editar.descripcion} onChange={e => setEditar({ ...editar, descripcion: e.target.value })}
                rows={4}
                style={{ width: "100%", padding: "7px 10px", border: "1px solid #d4d4d4", borderRadius: 6, fontSize: 13, marginBottom: 16, boxSizing: "border-box", resize: "vertical", fontFamily: "inherit" }} />
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button onClick={() => setEditar(null)} disabled={guardandoEditar}
                  style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: guardandoEditar ? "default" : "pointer", fontSize: 12 }}>
                  <BtnContent icon={X}>Cancelar</BtnContent>
                </button>
                <button onClick={guardarEdicion} disabled={guardandoEditar}
                  style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #16a34a", background: "#dcfce7", color: "#14532d", cursor: guardandoEditar ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
                  <BtnContent icon={guardandoEditar ? RefreshCw : Save} iconColor="#14532d">{guardandoEditar ? "Guardando..." : "Guardar"}</BtnContent>
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function MenuItem({ item, onAction }) {
  const [open, setOpen] = useState(false);
  if (item.label === "---") return <div style={{ height: 1, background: "#e2e8f0", margin: "4px 0" }} />;
  if (item.submenu) return (
    <div style={{ position: "relative" }} onMouseEnter={() => setOpen(true)} onMouseLeave={() => setOpen(false)}>
      <div title={item.tooltip || undefined} style={{ padding: "7px 14px", cursor: "pointer", display: "flex", justifyContent: "space-between", gap: 24, fontSize: 13, color: "#1e293b", borderRadius: 4, background: open ? "#f1f5f9" : "transparent", whiteSpace: "nowrap" }}>
        <span>{item.label}</span><span style={{ fontSize: 10, color: "#94a3b8" }}>▶</span>
      </div>
      {open && (
        <div style={{ position: "absolute", left: "100%", top: 0, background: "#fff", border: "1px solid #e2e8f0", borderRadius: 6, boxShadow: "0 8px 24px rgba(0,0,0,0.12)", minWidth: 220, zIndex: 9999, padding: "4px 0" }}>
          {item.submenu.map((sub, i) => <MenuItem key={i} item={sub} onAction={onAction} />)}
        </div>
      )}
    </div>
  );
  return (
    <div title={item.tooltip || undefined} style={{ padding: "7px 14px", cursor: "pointer", fontSize: 13, color: "#1e293b", borderRadius: 4, display: "flex", alignItems: "center", gap: 8, whiteSpace: "nowrap" }}
      onMouseEnter={e => e.currentTarget.style.background = "#f1f5f9"}
      onMouseLeave={e => e.currentTarget.style.background = "transparent"}
      onClick={() => onAction(item.action, item.label)}>
      <Icon as={item.icon} size={14} color="#64748b" /><span>{item.label}</span>
    </div>
  );
}

function MenuGroup({ group, onAction }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ position: "relative" }} onMouseEnter={() => setOpen(true)} onMouseLeave={() => setOpen(false)}>
      <div title={group.tooltip || undefined} style={{ padding: "6px 14px", cursor: "pointer", fontSize: 13, fontWeight: 500, color: open ? "#2563eb" : "#1e293b", borderBottom: open ? "2px solid #2563eb" : "2px solid transparent", userSelect: "none", whiteSpace: "nowrap", display: "inline-flex", alignItems: "center", gap: 6 }}>
        <Icon as={group.icon} size={14} color={open ? "#2563eb" : "#475569"} />
        <span>{group.label}</span>
      </div>
      {open && (
        <div style={{ position: "absolute", top: "100%", left: 0, background: "#fff", border: "1px solid #e2e8f0", borderRadius: 6, boxShadow: "0 8px 24px rgba(0,0,0,0.12)", minWidth: 220, zIndex: 9998, padding: "4px 0" }}>
          {group.items.map((item, i) => <MenuItem key={i} item={item} onAction={onAction} />)}
        </div>
      )}
    </div>
  );
}

function ModalComponent({ action, label, onClose }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "2rem", minWidth: 380, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h2 style={{ fontSize: 16, fontWeight: 600, margin: 0 }}>{label}</h2>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}><Icon as={X} size={18} /></button>
        </div>
        <div style={{ background: "#f8fafc", borderRadius: 8, padding: "1rem", border: "1px solid #e2e8f0", marginBottom: 16 }}>
          <p style={{ margin: 0, fontSize: 13, color: "#64748b", fontFamily: "monospace" }}>Componente: <strong style={{ color: "#2563eb" }}>{action}</strong></p>
        </div>
        <button onClick={onClose} style={{ width: "100%", padding: "8px 16px", borderRadius: 6, background: "#2563eb", color: "#fff", border: "none", cursor: "pointer", fontSize: 13 }}><BtnContent icon={X}>Cerrar</BtnContent></button>
      </div>
    </div>
  );
}

function DonutChart({ data, total, onHover, onLeave }) {
  const size = 260, cx = 130, cy = 130, r = 90, inner = 50;
  let angulo = -Math.PI / 2;
  const segs = data.map((item, i) => {
    const pct = total ? item.importe / total : 0;
    const sweep = pct * 2 * Math.PI;
    const x1 = cx + r * Math.cos(angulo), y1 = cy + r * Math.sin(angulo);
    angulo += sweep;
    const x2 = cx + r * Math.cos(angulo), y2 = cy + r * Math.sin(angulo);
    const xi1 = cx + inner * Math.cos(angulo - sweep), yi1 = cy + inner * Math.sin(angulo - sweep);
    const xi2 = cx + inner * Math.cos(angulo), yi2 = cy + inner * Math.sin(angulo);
    const large = sweep > Math.PI ? 1 : 0;
    return { path: `M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} L ${xi2} ${yi2} A ${inner} ${inner} 0 ${large} 0 ${xi1} ${yi1} Z`, color: COLORES[i % COLORES.length], item, pct };
  });
  return (
    <svg width={size} height={size} style={{ filter: "drop-shadow(0 4px 12px rgba(0,0,0,0.10))" }}>
      {segs.map((s, i) => (
        <path key={i} d={s.path} fill={s.color} opacity={0.85} style={{ cursor: "pointer" }}
          onMouseEnter={e => onHover({ i, item: s.item, pct: s.pct, x: e.clientX, y: e.clientY })}
          onMouseMove={e => onHover(t => ({ ...t, x: e.clientX, y: e.clientY }))}
          onMouseLeave={onLeave} />
      ))}
      <text x={cx} y={cy - 8} textAnchor="middle" fontSize={12} fill="#64748b">Total</text>
      <text x={cx} y={cy + 10} textAnchor="middle" fontSize={13} fontWeight="600" fill="#1e293b">{fmt(total)}</text>
      <text x={cx} y={cy + 26} textAnchor="middle" fontSize={11} fill="#94a3b8">€</text>
    </svg>
  );
}

// Calcular el año fiscal (del 1 Octubre del año N-1 al 30 Septiembre del año N)
// Si la fecha actual está entre Oct-Dic → año fiscal = año natural + 1
// Si está entre Ene-Sep → año fiscal = año natural
function getAnoFiscal(fecha) {
  const d = fecha || new Date();
  const año = d.getFullYear();
  const mes = d.getMonth(); // 0 = enero, 9 = octubre
  return mes >= 9 ? año + 1 : año;
}

// Construir numerocompleto: CODIGO-NUMERO-AÑO (ej. "CDM-123-2026")
function buildNumeroCompleto(codigo, numero, ano) {
  const c = String(codigo || "").trim();
  const n = String(numero || "").trim();
  const a = String(ano || "").trim();
  if (!n) return "";
  return [c, n, a].filter(Boolean).join("-");
}

// ────────────────────────────────────────────────────────────────
// Detector de referencias SIEMENS
// Formato base:        9AA9999-9AA99-9AA9      (ej. 6ES7531-7KF00-0AB0)
// Versiones reducidas:  9AA9999-9AA99
//                      9AA9999-9AA9
//                      9AA9999-9AA
//                      9AA9999-9
//                      9AA9999
// Sin guiones:         9AA99999AA999AA9 (10-16 chars alfanuméricos)
//
// Devuelve un array de objetos: { ref: "6ES7531-7KF00-0AB0", original: "...", index: 12 }
// donde original es el match exacto y index la posición en el texto.
// ────────────────────────────────────────────────────────────────
function parseSiemensRefs(textoOriginal) {
  if (!textoOriginal) return [];
  const original = String(textoOriginal);

  // Trabajamos sobre una versión SIN espacios (ni tabs/saltos) para que las referencias
  // partidas por espacios (ej. "6SL3210 -1KE21 -3UF1") se detecten igualmente.
  // Mantenemos un mapa de cada posición del texto-sin-espacios a su posición en el original,
  // para poder resaltar correctamente la referencia dentro del texto que ve el usuario.
  const texto = [];        // caracteres sin espacios
  const mapaIdx = [];      // mapaIdx[i] = posición en 'original' del carácter texto[i]
  for (let i = 0; i < original.length; i++) {
    if (/\s/.test(original[i])) continue; // saltar cualquier espacio en blanco
    texto.push(original[i]);
    mapaIdx.push(i);
  }
  const textoSinEsp = texto.join("");
  if (!textoSinEsp) return [];

  const encontradas = [];
  const yaUsados = new Set();   // intervalos ya consumidos (en índices de textoSinEsp)

  // Devuelve {iniOrig, finOrig, original} mapeando un match de textoSinEsp al texto original
  const mapearAOriginal = (iniSinEsp, finSinEsp) => {
    const iniOrig = mapaIdx[iniSinEsp];
    const finOrig = mapaIdx[finSinEsp - 1] + 1; // posición original del último carácter + 1
    return { iniOrig, finOrig, original: original.substring(iniOrig, finOrig) };
  };

  // Patrón con guiones: 1 dígito, 2 letras, 4 dígitos seguido opcionalmente de más bloques
  const reConGuion = /\d[A-Z]{2}\d{4}(?:-\d(?:[A-Z]{1,2}(?:\d{1,2}(?:-\d(?:[A-Z]{1,2}(?:\d)?)?)?)?)?)?/g;
  let m;
  while ((m = reConGuion.exec(textoSinEsp)) !== null) {
    const ini = m.index, fin = m.index + m[0].length;
    // Si solo capturó el bloque base (7 chars, sin guiones) y justo después hay más
    // caracteres alfanuméricos, es una ref pegada sin guiones: la dejamos para el patrón siguiente.
    const soloBase = !m[0].includes("-") && m[0].length === 7;
    const siguienteEsAlfaNum = fin < textoSinEsp.length && /[0-9A-Z]/.test(textoSinEsp[fin]);
    if (soloBase && siguienteEsAlfaNum) continue;
    const mp = mapearAOriginal(ini, fin);
    encontradas.push({ ref: m[0].toUpperCase(), original: mp.original, index: mp.iniOrig });
    for (let k = ini; k < fin; k++) yaUsados.add(k);
  }

  // Patrón sin guiones: 1 dígito + 2 letras + 4 dígitos + caracteres alfanuméricos
  const reSinGuion = /\d[A-Z]{2}\d{4}[0-9A-Z]{0,11}/g;
  while ((m = reSinGuion.exec(textoSinEsp)) !== null) {
    const ini = m.index, fin = m.index + m[0].length;
    let solapa = false;
    for (let k = ini; k < fin; k++) { if (yaUsados.has(k)) { solapa = true; break; } }
    if (solapa) continue;
    if (m[0].includes("-")) continue;
    if (m[0].length >= 10) {
      const mp = mapearAOriginal(ini, fin);
      encontradas.push({ ref: m[0].toUpperCase(), original: mp.original, index: mp.iniOrig });
      for (let k = ini; k < fin; k++) yaUsados.add(k);
    }
  }

  // Ordenar por posición de aparición y deduplicar (misma ref = misma posición de inicio)
  encontradas.sort((a, b) => a.index - b.index);
  const vistas = new Set();
  return encontradas.filter(e => {
    const k = `${e.index}:${e.ref}`;
    if (vistas.has(k)) return false;
    vistas.add(k);
    return true;
  });
}

// ── Diálogo Leer Precios de PMD ──
function LeerPreciosPMDDialog({ dialog, onClose, setStatus, aplicarAFila }) {
  const [referencia, setReferencia] = useState(dialog.referencia || "");
  const [cargando, setCargando] = useState(false);
  const [accion, setAccion] = useState(""); // "presupuesto" o "csv"

  const obtenerYAplicar = async () => {
    const ref = referencia.trim();
    if (!ref) {
      setStatus("Introduce una referencia", "error");
      return;
    }
    setCargando(true);
    setAccion("presupuesto");
    setStatus(`Consultando PMD para "${ref}"...`, "working");
    try {
      const r = await fetch(`${API_LOCAL_URL}/precio?mlfb=${encodeURIComponent(ref)}`);
      if (!r.ok) throw new Error("HTTP " + r.status);
      const data = await r.json();
      // Aplicar a la fila objetivo (si existe)
      if (dialog.rowIdx >= 0) {
        aplicarAFila(dialog.rowIdx, { ...data, referencia: ref });
        setStatus(`Datos de PMD aplicados a la fila ${dialog.rowIdx + 1} (${ref})`, "success");
        onClose();
      } else {
        setStatus(`Datos recibidos pero no hay fila seleccionada para aplicar`, "info");
      }
    } catch (e) {
      setStatus(`Error consultando PMD: ${e.message}`, "error");
    } finally {
      setCargando(false);
      setAccion("");
    }
  };

  const guardarCSV = async () => {
    const ref = referencia.trim();
    if (!ref) {
      setStatus("Introduce una referencia", "error");
      return;
    }
    setCargando(true);
    setAccion("csv");
    setStatus(`Generando CSV de PMD para "${ref}"...`, "working");
    try {
      const r = await fetch(`${API_LOCAL_URL}/precio_csv?mlfb=${encodeURIComponent(ref)}`);
      if (!r.ok) throw new Error("HTTP " + r.status);
      const data = await r.json();
      if (data.error) {
        setStatus(`CSV generado con aviso: ${data.error} (${data.csv_path || ""})`, "warning");
      } else {
        setStatus(`CSV generado: ${data.csv_path}`, "success");
      }
    } catch (e) {
      setStatus(`Error generando CSV: ${e.message}`, "error");
    } finally {
      setCargando(false);
      setAccion("");
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }} onClick={() => !cargando && onClose()}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 520, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, paddingBottom: 10, borderBottom: "1px solid #e5e5e5" }}>
          <h2 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Download} size={18} color="#1e3a5f" /> Leer Precios de PMD
          </h2>
          <button onClick={onClose} disabled={cargando} style={{ border: "none", background: "none", cursor: cargando ? "default" : "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        <div style={{ marginBottom: 14 }}>
          <label htmlFor="pmd-ref" style={{ display: "block", fontSize: 11, color: "#525252", fontWeight: 500, marginBottom: 4 }}>Referencia (MLFB)</label>
          <input
            id="pmd-ref"
            type="text"
            value={referencia}
            onChange={e => setReferencia(e.target.value)}
            autoFocus
            placeholder="ej. 6ES7511-1AL03-0AB0"
            style={{ width: "100%", padding: "8px 10px", border: "1px solid #cbd5e1", borderRadius: 6, fontSize: 13, background: "#fff", boxSizing: "border-box", fontFamily: "monospace" }}
          />
        </div>

        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", flexWrap: "wrap", paddingTop: 8, borderTop: "1px solid #e5e5e5" }}>
          <button onClick={onClose} disabled={cargando}
            style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: cargando ? "default" : "pointer", fontSize: 12 }}>
            <BtnContent icon={X}>Cancelar</BtnContent>
          </button>
          <button onClick={guardarCSV} disabled={cargando || !referencia.trim()}
            style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#171717", cursor: (cargando || !referencia.trim()) ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={(cargando && accion === "csv") ? RefreshCw : FileSpreadsheet} iconColor="#475569">
              {(cargando && accion === "csv") ? "Generando..." : "datos de PMD a CSV"}
            </BtnContent>
          </button>
          <button onClick={obtenerYAplicar} disabled={cargando || !referencia.trim()}
            style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #16a34a", background: "#dcfce7", color: "#14532d", cursor: (cargando || !referencia.trim()) ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={(cargando && accion === "presupuesto") ? RefreshCw : Check} iconColor="#14532d">
              {(cargando && accion === "presupuesto") ? "Consultando..." : "datos de PMD a Presupuesto"}
            </BtnContent>
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Diálogo de Login (reutilizable) ──
function LoginDialog({ onClose, onLoginOk }) {
  const [usuario, setUsuario] = useState("");
  const [password, setPassword] = useState("");
  const [aviso, setAviso] = useState(null);
  const [verificando, setVerificando] = useState(false);

  const submit = async () => {
    if (!usuario.trim()) {
      setAviso({ tipo: "error", texto: "Introduce un usuario" });
      return;
    }
    setVerificando(true);
    setAviso(null);
    try {
      const res = await fetch(`${API_URL}/usuarios/login`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ usuario: usuario.trim(), password: password || "" }),
      });
      if (res.ok) {
        const data = await res.json();
        if (data.ok) {
          setAviso({ tipo: "ok", texto: `Login OK — bienvenido ${data.usuario}` });
          setVerificando(false);
          setTimeout(() => onLoginOk(data), 700);
          return;
        } else {
          setAviso({ tipo: "error", texto: `Login KO — ${data.motivo || "credenciales incorrectas"}` });
        }
      } else {
        setAviso({ tipo: "error", texto: `Error servidor (${res.status})` });
      }
    } catch (err) {
      setAviso({ tipo: "error", texto: `No se pudo verificar: ${err.message}` });
    }
    setVerificando(false);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }} onClick={() => !verificando && onClose()}>
      <div style={{ background: "rgba(255,255,255,0.97)", borderRadius: 12, padding: "1.8rem 2rem", width: "95%", maxWidth: 380, boxShadow: "0 20px 60px rgba(0,0,0,0.3)" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, paddingBottom: 10, borderBottom: "1px solid #e5e5e5" }}>
          <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={LogIn} size={18} color="#1e3a5f" /> Iniciar sesión
          </h3>
          <button onClick={onClose} disabled={verificando} style={{ border: "none", background: "none", cursor: verificando ? "default" : "pointer", padding: 4, color: "#94a3b8", display: "inline-flex" }}>
            <Icon as={X} size={18} />
          </button>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 4, marginBottom: 12 }}>
          <label htmlFor="ld-user" style={{ fontSize: 11, color: "#525252", fontWeight: 500 }}>Usuario</label>
          <div style={{ position: "relative" }}>
            <input
              id="ld-user"
              type="text"
              value={usuario}
              onChange={e => setUsuario(e.target.value)}
              autoFocus
              style={{
                width: "100%",
                padding: "8px 10px 8px 32px",
                border: "1px solid #cbd5e1",
                borderRadius: 6,
                fontSize: 13,
                background: "#fff",
                boxSizing: "border-box",
              }}
            />
            <div style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }}>
              <Icon as={User} size={14} color="#94a3b8" />
            </div>
          </div>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 4, marginBottom: 14 }}>
          <label htmlFor="ld-pass" style={{ fontSize: 11, color: "#525252", fontWeight: 500 }}>Contraseña</label>
          <div style={{ position: "relative" }}>
            <input
              id="ld-pass"
              type="password"
              value={password}
              onChange={e => setPassword(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter" && !verificando) submit(); }}
              style={{
                width: "100%",
                padding: "8px 10px 8px 32px",
                border: "1px solid #cbd5e1",
                borderRadius: 6,
                fontSize: 13,
                background: "#fff",
                boxSizing: "border-box",
              }}
            />
            <div style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }}>
              <Icon as={Lock} size={14} color="#94a3b8" />
            </div>
          </div>
        </div>

        {aviso && (
          <div style={{
            marginBottom: 12,
            padding: "8px 12px",
            borderRadius: 6,
            fontSize: 12,
            fontWeight: 500,
            textAlign: "center",
            background: aviso.tipo === "ok" ? "#dcfce7" : "#fef2f2",
            border: `1px solid ${aviso.tipo === "ok" ? "#86efac" : "#fca5a5"}`,
            color: aviso.tipo === "ok" ? "#14532d" : "#991b1b",
          }}>
            {aviso.texto}
          </div>
        )}

        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
          <button onClick={onClose} disabled={verificando}
            style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: verificando ? "default" : "pointer", fontSize: 12 }}>
            <BtnContent icon={X}>Cancelar</BtnContent>
          </button>
          <button onClick={submit} disabled={verificando}
            style={{ padding: "7px 18px", borderRadius: 6, border: "1px solid #16a34a", background: verificando ? "#f5f5f5" : "#dcfce7", color: "#14532d", cursor: verificando ? "default" : "pointer", fontSize: 12, fontWeight: 600 }}>
            <BtnContent icon={verificando ? RefreshCw : Check} iconColor="#14532d">{verificando ? "Verificando..." : "Iniciar sesión"}</BtnContent>
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Pantalla de bienvenida / login ──
// Constantes del control de intentos de login
const LOGIN_MAX_INTENTOS = 5;
const LOGIN_BLOQUEO_MS = 10 * 60 * 1000; // 10 minutos
const LOGIN_STORAGE_KEY = "loginIntentos"; // { fallos: n, bloqueadoHasta: timestamp }

function leerEstadoLogin() {
  try {
    const raw = localStorage.getItem(LOGIN_STORAGE_KEY);
    if (!raw) return { fallos: 0, bloqueadoHasta: 0 };
    const obj = JSON.parse(raw);
    return { fallos: obj.fallos || 0, bloqueadoHasta: obj.bloqueadoHasta || 0 };
  } catch { return { fallos: 0, bloqueadoHasta: 0 }; }
}
function guardarEstadoLogin(estado) {
  try { localStorage.setItem(LOGIN_STORAGE_KEY, JSON.stringify(estado)); } catch {}
}

function WelcomeScreen({ onLogin }) {
  const [usuario, setUsuario] = useState("");
  const [password, setPassword] = useState("");

  const [aviso, setAviso] = useState(null); // { tipo: "ok"|"error", texto: string }
  const [verificando, setVerificando] = useState(false);

  // Control de intentos / bloqueo temporal
  const [estadoLogin, setEstadoLogin] = useState(() => leerEstadoLogin());
  const [ahora, setAhora] = useState(Date.now());

  // Si está bloqueado, actualizar el reloj cada segundo para mostrar la cuenta atrás
  useEffect(() => {
    if (estadoLogin.bloqueadoHasta > Date.now()) {
      const t = setInterval(() => setAhora(Date.now()), 1000);
      return () => clearInterval(t);
    }
  }, [estadoLogin.bloqueadoHasta]);

  const bloqueado = estadoLogin.bloqueadoHasta > ahora;
  const minutosRestantes = bloqueado ? Math.ceil((estadoLogin.bloqueadoHasta - ahora) / 60000) : 0;
  const segundosRestantes = bloqueado ? Math.ceil((estadoLogin.bloqueadoHasta - ahora) / 1000) : 0;

  const registrarFallo = () => {
    const fallos = estadoLogin.fallos + 1;
    let nuevoEstado;
    if (fallos >= LOGIN_MAX_INTENTOS) {
      nuevoEstado = { fallos, bloqueadoHasta: Date.now() + LOGIN_BLOQUEO_MS };
    } else {
      nuevoEstado = { fallos, bloqueadoHasta: 0 };
    }
    setEstadoLogin(nuevoEstado);
    guardarEstadoLogin(nuevoEstado);
    return nuevoEstado;
  };

  const submit = async (e) => {
    if (e && e.preventDefault) e.preventDefault();
    // Si está bloqueado, no permitir intentar
    if (estadoLogin.bloqueadoHasta > Date.now()) {
      const min = Math.ceil((estadoLogin.bloqueadoHasta - Date.now()) / 60000);
      setAviso({ tipo: "error", texto: `Demasiados intentos fallidos. Inténtalo de nuevo en ${min} minuto(s).` });
      return;
    }
    // El usuario es obligatorio para acceder
    if (!usuario.trim()) {
      setAviso({ tipo: "error", texto: "Introduce tu usuario y contraseña para acceder." });
      return;
    }
    setVerificando(true);
    setAviso(null);
    try {
      const res = await fetch(`${API_URL}/usuarios/login`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ usuario: usuario.trim(), password: password || "" }),
      });
      let loginOk = null; // datos del usuario si fue OK
      if (res.ok) {
        const data = await res.json();
        if (data.ok) {
          // Login correcto: limpiar contador y entrar
          const limpio = { fallos: 0, bloqueadoHasta: 0 };
          setEstadoLogin(limpio);
          guardarEstadoLogin(limpio);
          setAviso({ tipo: "ok", texto: `Login OK — bienvenido ${data.usuario}` });
          loginOk = data;
          setVerificando(false);
          setTimeout(() => onLogin(loginOk), 700);
          return;
        }
      }
      // Login fallido (credenciales incorrectas o error de servidor)
      setVerificando(false);
      const nuevoEstado = registrarFallo();
      const restantes = LOGIN_MAX_INTENTOS - nuevoEstado.fallos;
      if (nuevoEstado.bloqueadoHasta > Date.now()) {
        setAviso({ tipo: "error", texto: `Demasiados intentos fallidos. Acceso bloqueado durante 10 minutos.` });
      } else {
        setAviso({ tipo: "error", texto: `Credenciales incorrectas. Te quedan ${restantes} intento(s).` });
      }
    } catch (err) {
      // Error de red: NO consume intento (no es culpa de credenciales)
      setVerificando(false);
      setAviso({ tipo: "error", texto: `No se pudo verificar el login: ${err.message}` });
    }
  };

  return (
    <div style={{
      fontFamily: "'Segoe UI', system-ui, sans-serif",
      height: "100vh",
      width: "100vw",
      backgroundImage: "url('/inicio.png')",
      backgroundSize: "cover",
      backgroundPosition: "center",
      backgroundRepeat: "no-repeat",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      position: "relative",
    }}>
      {/* Título */}
      <div style={{
        position: "absolute",
        top: "15%",
        left: 0,
        right: 0,
        textAlign: "center",
      }}>
        <h1 style={{
          margin: 0,
          fontSize: "clamp(40px, 7vw, 88px)",
          fontWeight: 700,
          color: "#525252",
          letterSpacing: -1,
          textShadow: [
            "-2px -2px 0 #fff",
            "2px -2px 0 #fff",
            "-2px 2px 0 #fff",
            "2px 2px 0 #fff",
          ].join(", "),
        }}>
          CdM_Presupuestos
        </h1>
      </div>

      {/* Caja de login centrada */}
      <div style={{
        background: "rgba(255,255,255,0.92)",
        backdropFilter: "blur(8px)",
        border: "1px solid #e5e5e5",
        borderRadius: 14,
        padding: "28px 32px",
        boxShadow: "0 20px 40px rgba(0,0,0,0.15), 0 2px 8px rgba(0,0,0,0.08)",
        width: "100%",
        maxWidth: 360,
        display: "flex",
        flexDirection: "column",
        gap: 14,
      }}>
        <h2 style={{
          margin: 0,
          fontSize: 18,
          fontWeight: 700,
          color: "#171717",
          textAlign: "center",
          paddingBottom: 6,
          borderBottom: "1px solid #e5e5e5",
        }}>
          Iniciar sesión
        </h2>

        {/* Usuario */}
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <label htmlFor="login-user" style={{ fontSize: 11, color: "#525252", fontWeight: 500 }}>Usuario</label>
          <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 10px", border: "1px solid #d4d4d4", borderRadius: 6, background: "#fff" }}>
            <Icon as={User} size={14} color="#737373" />
            <input
              id="login-user"
              type="text"
              value={usuario}
              autoFocus
              autoComplete="username"
              onChange={e => setUsuario(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter") submit(); }}
              placeholder="Tu usuario"
              style={{ flex: 1, border: "none", outline: "none", fontSize: 13, color: "#171717", background: "transparent" }}
            />
          </div>
        </div>

        {/* Contraseña */}
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <label htmlFor="login-pass" style={{ fontSize: 11, color: "#525252", fontWeight: 500 }}>Contraseña</label>
          <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 10px", border: "1px solid #d4d4d4", borderRadius: 6, background: "#fff" }}>
            <Icon as={Lock} size={14} color="#737373" />
            <input
              id="login-pass"
              type="password"
              value={password}
              autoComplete="current-password"
              onChange={e => setPassword(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter") submit(); }}
              placeholder="••••••••"
              style={{ flex: 1, border: "none", outline: "none", fontSize: 13, color: "#171717", background: "transparent" }}
            />
          </div>
        </div>

        {/* Aviso de login */}
        {aviso && (
          <div style={{
            marginTop: 4,
            padding: "8px 12px",
            borderRadius: 6,
            fontSize: 12,
            fontWeight: 500,
            textAlign: "center",
            background: aviso.tipo === "ok" ? "#dcfce7" : "#fef2f2",
            border: `1px solid ${aviso.tipo === "ok" ? "#86efac" : "#fca5a5"}`,
            color: aviso.tipo === "ok" ? "#14532d" : "#991b1b",
          }}>
            {aviso.texto}
          </div>
        )}
        {/* Botón login */}
        <button
          type="button"
          onClick={submit} disabled={verificando || bloqueado}
          onMouseEnter={e => { if (!bloqueado && !verificando) { e.currentTarget.style.background = "#bbf7d0"; e.currentTarget.style.borderColor = "#15803d"; } }}
          onMouseLeave={e => { if (!bloqueado && !verificando) { e.currentTarget.style.background = "#dcfce7"; e.currentTarget.style.borderColor = "#16a34a"; } }}
          style={{
            marginTop: 4,
            padding: "9px 16px",
            fontSize: 13,
            fontWeight: 600,
            color: bloqueado ? "#991b1b" : "#14532d",
            background: bloqueado ? "#fee2e2" : "#dcfce7",
            border: `1px solid ${bloqueado ? "#fca5a5" : "#16a34a"}`,
            borderRadius: 6,
            cursor: (bloqueado || verificando) ? "default" : "pointer",
            transition: "all 0.15s ease",
            display: "inline-flex",
            alignItems: "center",
            justifyContent: "center",
            gap: 6,
          }}>
          <Icon as={verificando ? RefreshCw : (bloqueado ? Lock : Check)} size={14} color={bloqueado ? "#991b1b" : "#14532d"} />
          {verificando ? "Verificando..." : bloqueado ? (minutosRestantes > 1 ? `Bloqueado (${minutosRestantes} min)` : `Bloqueado (${segundosRestantes} s)`) : "Login"}
        </button>
      </div>
    </div>
  );
}

// ── Diálogo de referencias SIEMENS detectadas ──
function SiemensRefsDialog({ celdasTrabajadas, onClose, onComplete, setStatus }) {
  // celdasTrabajadas = [{ rowId, colKey, texto, refs: [{ref, index}, ...] }]
  const [posCelda, setPosCelda] = useState(0);
  const [posRef, setPosRef] = useState(0);
  const [aceptadas, setAceptadas] = useState([]);  // [{ rowId, colKey, ref, esPrimera }] o nuevas filas
  const [busqBD, setBusqBD] = useState({ buscando: false, existe: null });

  const celdaActual = celdasTrabajadas[posCelda];
  const refActual = celdaActual?.refs[posRef];
  const totalCeldas = celdasTrabajadas.length;
  const totalRefsCelda = celdaActual?.refs.length || 0;

  const refsAceptadasCelda = aceptadas.filter(a => a.rowId === celdaActual?.rowId && a.colKey === celdaActual?.colKey).length;

  // avanzar recibe opcionalmente la lista de aceptadas actualizada (para no depender del
  // estado asíncrono cuando se acepta la última referencia)
  const avanzar = (listaAceptadas) => {
    setBusqBD({ buscando: false, existe: null });
    if (posRef + 1 < totalRefsCelda) {
      setPosRef(posRef + 1);
    } else if (posCelda + 1 < totalCeldas) {
      setPosCelda(posCelda + 1);
      setPosRef(0);
    } else {
      // Hemos terminado → usar la lista pasada si existe, si no el estado actual
      onComplete(listaAceptadas || aceptadas);
    }
  };

  const aceptar = () => {
    const esPrimera = refsAceptadasCelda === 0;
    const nuevaLista = [...aceptadas, {
      rowId: celdaActual.rowId,
      colKey: celdaActual.colKey,
      ref: refActual.ref,
      esPrimera,
    }];
    setAceptadas(nuevaLista);
    avanzar(nuevaLista);
  };

  const saltar = () => avanzar();

  const cancelar = () => {
    setStatus && setStatus("Búsqueda SIEMENS cancelada", "info");
    onClose();
  };

  const buscarEnBD = async () => {
    if (!refActual) return;
    setBusqBD({ buscando: true, existe: null });
    try {
      const res = await fetch(`${API_URL}/productos/referencia/${encodeURIComponent(refActual.ref)}`);
      if (res.ok) {
        setBusqBD({ buscando: false, existe: true });
      } else if (res.status === 404) {
        setBusqBD({ buscando: false, existe: false });
      } else {
        throw new Error("HTTP " + res.status);
      }
    } catch (e) {
      setBusqBD({ buscando: false, existe: null, error: e.message });
    }
  };

  if (!celdaActual || !refActual) {
    return null;
  }

  // Resaltar la referencia dentro del texto original
  const textoOriginal = celdaActual.texto || "";
  const ini = refActual.index;
  const fin = ini + refActual.original.length;
  const antes = textoOriginal.substring(0, ini);
  const elMatch = textoOriginal.substring(ini, fin);
  const despues = textoOriginal.substring(fin);

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }} onClick={cancelar}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 620, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
        {/* Cabecera */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, paddingBottom: 10, borderBottom: "1px solid #e5e5e5" }}>
          <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Search} size={16} color="#171717" /> Referencia SIEMENS detectada
          </h3>
          <span style={{ fontSize: 11, color: "#737373" }}>
            Celda {posCelda + 1} de {totalCeldas} · Ref {posRef + 1} de {totalRefsCelda}
          </span>
        </div>

        {/* Texto original con la ref resaltada */}
        <div style={{ marginBottom: 10 }}>
          <div style={{ fontSize: 11, color: "#525252", fontWeight: 500, marginBottom: 4 }}>Texto de la celda:</div>
          <div style={{ padding: "10px 12px", background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 6, fontSize: 12, color: "#1e293b", lineHeight: 1.5, maxHeight: 100, overflow: "auto", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
            {antes}
            <strong style={{ background: "#fef3c7", color: "#92400e", padding: "1px 4px", borderRadius: 3, border: "1px solid #fcd34d" }}>{elMatch}</strong>
            {despues}
          </div>
        </div>

        {/* Referencia detectada */}
        <div style={{ marginBottom: 14, padding: "10px 14px", background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 6 }}>
          <div style={{ fontSize: 11, color: "#1e3a5f", fontWeight: 500, marginBottom: 4 }}>Referencia detectada (normalizada en mayúsculas):</div>
          <div style={{ fontSize: 18, fontFamily: "monospace", fontWeight: 700, color: "#1e3a5f", letterSpacing: 0.5 }}>{refActual.ref}</div>
        </div>

        {/* Resultado de búsqueda en BD */}
        {busqBD.existe !== null && (
          <div style={{ marginBottom: 12, padding: "8px 12px", borderRadius: 6, fontSize: 12,
            background: busqBD.existe ? "#dcfce7" : "#fef2f2",
            border: `1px solid ${busqBD.existe ? "#86efac" : "#fca5a5"}`,
            color: busqBD.existe ? "#14532d" : "#991b1b" }}>
            {busqBD.existe
              ? <><strong>✓ Existe en BD</strong> — el producto con esta referencia está registrado.</>
              : <><strong>✗ No existe en BD</strong> — esta referencia no está registrada como producto.</>}
          </div>
        )}
        {busqBD.error && (
          <div style={{ marginBottom: 12, padding: "8px 12px", borderRadius: 6, fontSize: 12, background: "#fef2f2", border: "1px solid #fca5a5", color: "#991b1b" }}>
            Error consultando BD: {busqBD.error}
          </div>
        )}

        {/* Botones */}
        <div style={{ display: "flex", gap: 8, justifyContent: "space-between", paddingTop: 12, borderTop: "1px solid #e5e5e5" }}>
          <button onClick={buscarEnBD} disabled={busqBD.buscando}
            style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: busqBD.buscando ? "default" : "pointer", fontSize: 12 }}>
            <BtnContent icon={busqBD.buscando ? RefreshCw : Database} iconColor="#475569">
              {busqBD.buscando ? "Buscando..." : "Buscar en BD"}
            </BtnContent>
          </button>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={cancelar}
              style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
              <BtnContent icon={X}>Cancelar</BtnContent>
            </button>
            <button onClick={saltar}
              style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
              <BtnContent icon={CornerDownLeft} iconColor="#475569">Saltar</BtnContent>
            </button>
            <button onClick={aceptar}
              onMouseEnter={e => { e.currentTarget.style.background = "#bbf7d0"; e.currentTarget.style.borderColor = "#15803d"; }}
              onMouseLeave={e => { e.currentTarget.style.background = "#dcfce7"; e.currentTarget.style.borderColor = "#16a34a"; }}
              style={{ padding: "7px 18px", borderRadius: 6, border: "1px solid #16a34a", background: "#dcfce7", color: "#14532d", cursor: "pointer", fontSize: 12, fontWeight: 600, transition: "all 0.15s ease" }}>
              <BtnContent icon={Check} iconColor="#14532d">Aceptar</BtnContent>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}


// ── Diálogo Actualizar Productos existentes ──
// Permite elegir qué columnas actualizar (global) y si pedir confirmación por fila.
function ActualizarProductosDialog({ datos, onClose, setStatus }) {
  const { existentes } = datos;
  // Columnas que se pueden actualizar
  const COLS = [
    { key: "pvp", label: "PVP" },
    { key: "descripcion", label: "Descripción" },
    { key: "nombre", label: "Nombre" },
    { key: "preciocoste", label: "Precio coste" },
    { key: "grupodescuento", label: "Grupo descuento" },
  ];
  // Solo ofrecer columnas que tengan valor en al menos una fila
  const colsDisponibles = COLS.filter(c =>
    existentes.some(e => {
      const v = e.nuevos[c.key];
      return v !== null && v !== undefined && String(v).trim() !== "";
    })
  );
  const [seleccion, setSeleccion] = useState(() => {
    const s = {}; colsDisponibles.forEach(c => { s[c.key] = true; }); return s;
  });
  const [pedirConfirmacion, setPedirConfirmacion] = useState(false);
  const [fase, setFase] = useState("config"); // "config" | "porFila" | "procesando"
  const [idxActual, setIdxActual] = useState(0);
  const [resultados, setResultados] = useState({ actualizados: 0, omitidos: 0, errores: 0 });
  const [log, setLog] = useState([]);

  const colsElegidas = colsDisponibles.filter(c => seleccion[c.key]);

  const valorMostrar = (key, val) => {
    if (val === null || val === undefined || val === "") return "(vacío)";
    if (key === "pvp" || key === "preciocoste") return fmtEur(Number(val));
    return String(val);
  };

  // Aplica la actualización a un producto (PATCH). Devuelve "ok"|"error"
  const actualizarUno = async (item) => {
    const payload = {};
    colsElegidas.forEach(c => {
      const v = item.nuevos[c.key];
      if (v !== null && v !== undefined && String(v).trim() !== "") payload[c.key] = v;
    });
    if (Object.keys(payload).length === 0) return "omitido";
    try {
      const r = await fetch(`${API_URL}/productos/${item.productoBD.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!r.ok) { const e = await r.json().catch(() => null); throw new Error(e?.detail || "HTTP " + r.status); }
      return "ok";
    } catch (e) {
      return "error:" + e.message;
    }
  };

  // Modo sin confirmación: actualizar todos de golpe
  const procesarTodos = async () => {
    setFase("procesando");
    let act = 0, om = 0, err = 0;
    const nuevoLog = [];
    for (const item of existentes) {
      const ref = item.fila.referencia;
      const res = await actualizarUno(item);
      if (res === "ok") { act++; nuevoLog.push(`✓ ${ref}: actualizado`); }
      else if (res === "omitido") { om++; nuevoLog.push(`⊘ ${ref}: sin columnas que actualizar`); }
      else { err++; nuevoLog.push(`✗ ${ref}: ${res.replace("error:", "")}`); }
    }
    setResultados({ actualizados: act, omitidos: om, errores: err });
    setLog(nuevoLog);
    setStatus(`Actualizar productos: ${act} actualizado(s), ${om} omitido(s), ${err} error(es)`, err > 0 ? "error" : "success");
    setFase("fin");
  };

  // Iniciar el flujo según el toggle de confirmación
  const iniciar = () => {
    if (colsElegidas.length === 0) {
      setStatus("Selecciona al menos una columna para actualizar", "error");
      return;
    }
    if (pedirConfirmacion) {
      setIdxActual(0);
      setFase("porFila");
    } else {
      procesarTodos();
    }
  };

  // Decisión por fila (modo confirmación SÍ)
  const decidirFila = async (aplicar) => {
    const item = existentes[idxActual];
    const ref = item.fila.referencia;
    if (aplicar) {
      const res = await actualizarUno(item);
      setResultados(r => ({
        ...r,
        actualizados: r.actualizados + (res === "ok" ? 1 : 0),
        errores: r.errores + (res.startsWith && res.startsWith("error") ? 1 : 0),
        omitidos: r.omitidos + (res === "omitido" ? 1 : 0),
      }));
      setLog(l => [...l, res === "ok" ? `✓ ${ref}: actualizado` : res === "omitido" ? `⊘ ${ref}: sin columnas` : `✗ ${ref}: ${String(res).replace("error:", "")}`]);
    } else {
      setResultados(r => ({ ...r, omitidos: r.omitidos + 1 }));
      setLog(l => [...l, `⊘ ${ref}: omitido por el usuario`]);
    }
    // Avanzar
    if (idxActual + 1 < existentes.length) {
      setIdxActual(idxActual + 1);
    } else {
      setFase("fin");
      setStatus("Actualización por fila terminada", "success");
    }
  };

  const item = existentes[idxActual];

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100002 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "92%", maxWidth: 600, maxHeight: "88vh", overflow: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>

        {fase === "config" && (
          <>
            <h3 style={{ fontSize: 16, fontWeight: 700, color: "#171717", margin: "0 0 6px" }}>Actualizar productos existentes</h3>
            <p style={{ fontSize: 13, color: "#475569", margin: "0 0 16px", lineHeight: 1.5 }}>
              {existentes.length} producto(s) seleccionado(s) ya existen en el catálogo. Elige qué columnas actualizar.
            </p>
            <div style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: "#171717", marginBottom: 8 }}>Columnas a actualizar:</div>
              {colsDisponibles.length === 0 ? (
                <p style={{ fontSize: 12, color: "#dc2626" }}>Las filas no tienen valores en columnas actualizables.</p>
              ) : colsDisponibles.map(c => (
                <label key={c.key} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", fontSize: 13, cursor: "pointer" }}>
                  <input type="checkbox" checked={!!seleccion[c.key]}
                    onChange={e => setSeleccion(s => ({ ...s, [c.key]: e.target.checked }))} />
                  {c.label}
                </label>
              ))}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 20 }}>
              <span style={{ fontSize: 13, color: "#171717", fontWeight: 500 }}>Pedir confirmación por cada fila:</span>
              <div style={{ display: "inline-flex", border: "1px solid #d4d4d4", borderRadius: 6, overflow: "hidden" }}>
                <button onClick={() => setPedirConfirmacion(true)}
                  style={{ padding: "4px 14px", fontSize: 12, border: "none", cursor: "pointer", background: pedirConfirmacion ? "#171717" : "#fff", color: pedirConfirmacion ? "#fff" : "#171717", fontWeight: pedirConfirmacion ? 600 : 400 }}>SÍ</button>
                <button onClick={() => setPedirConfirmacion(false)}
                  style={{ padding: "4px 14px", fontSize: 12, border: "none", borderLeft: "1px solid #d4d4d4", cursor: "pointer", background: !pedirConfirmacion ? "#171717" : "#fff", color: !pedirConfirmacion ? "#fff" : "#171717", fontWeight: !pedirConfirmacion ? 600 : 400 }}>NO</button>
              </div>
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <button onClick={onClose} style={{ padding: "8px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>Cancelar</button>
              <button onClick={iniciar} disabled={colsElegidas.length === 0}
                style={{ padding: "8px 16px", borderRadius: 6, border: "none", background: colsElegidas.length === 0 ? "#cbd5e1" : "#2563eb", color: "#fff", cursor: colsElegidas.length === 0 ? "default" : "pointer", fontSize: 13, fontWeight: 600 }}>
                {pedirConfirmacion ? "Revisar fila a fila" : "Actualizar todos"}
              </button>
            </div>
          </>
        )}

        {fase === "porFila" && item && (
          <>
            <h3 style={{ fontSize: 16, fontWeight: 700, color: "#171717", margin: "0 0 4px" }}>
              Confirmar actualización ({idxActual + 1} de {existentes.length})
            </h3>
            <p style={{ fontSize: 13, color: "#475569", margin: "0 0 14px" }}>
              Producto <strong>{item.fila.referencia}</strong> (id {item.productoBD.id})
            </p>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, marginBottom: 18 }}>
              <thead>
                <tr style={{ background: "#f8fafc" }}>
                  <th style={{ textAlign: "left", padding: "6px 8px", border: "1px solid #e2e8f0" }}>Columna</th>
                  <th style={{ textAlign: "left", padding: "6px 8px", border: "1px solid #e2e8f0" }}>Valor actual (BD)</th>
                  <th style={{ textAlign: "left", padding: "6px 8px", border: "1px solid #e2e8f0" }}>Nuevo valor</th>
                </tr>
              </thead>
              <tbody>
                {colsElegidas.map(c => {
                  const actualBD = c.key === "grupodescuento" ? (item.productoBD.idgrupodescuento ?? item.productoBD.grupodescuento) : item.productoBD[c.key];
                  const nuevo = c.key === "grupodescuento" ? item.nuevos.grupoCod : item.nuevos[c.key];
                  return (
                    <tr key={c.key}>
                      <td style={{ padding: "6px 8px", border: "1px solid #e2e8f0", fontWeight: 600 }}>{c.label}</td>
                      <td style={{ padding: "6px 8px", border: "1px solid #e2e8f0", color: "#64748b" }}>{valorMostrar(c.key, actualBD)}</td>
                      <td style={{ padding: "6px 8px", border: "1px solid #e2e8f0", color: "#16a34a", fontWeight: 600 }}>{valorMostrar(c.key, nuevo)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <button onClick={() => decidirFila(false)} style={{ padding: "8px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>No actualizar</button>
              <button onClick={() => decidirFila(true)} style={{ padding: "8px 16px", borderRadius: 6, border: "none", background: "#16a34a", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>Actualizar este</button>
            </div>
          </>
        )}

        {(fase === "procesando" || fase === "fin") && (
          <>
            <h3 style={{ fontSize: 16, fontWeight: 700, color: "#171717", margin: "0 0 10px" }}>
              {fase === "procesando" ? "Actualizando..." : "Actualización terminada"}
            </h3>
            <div style={{ fontSize: 13, color: "#475569", marginBottom: 8, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span>{resultados.actualizados} actualizado(s), {resultados.omitidos} omitido(s), {resultados.errores} error(es)</span>
              {log.length > 0 && (
                <div style={{ display: "inline-flex", gap: 6 }}>
                  <button onClick={() => exportarLogCSV(log, "actualizar-productos")}
                    title="Exportar el log a un fichero CSV"
                    style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                    <BtnContent icon={Download} iconColor="#475569">Exportar CSV</BtnContent>
                  </button>
                  <button onClick={() => setLog([])}
                    style={{ padding: "3px 10px", fontSize: 10, border: "1px solid #d4d4d4", borderRadius: 4, background: "#fff", color: "#171717", cursor: "pointer" }}>
                    <BtnContent icon={Trash2} iconColor="#475569">Limpiar log</BtnContent>
                  </button>
                </div>
              )}
            </div>
            <div style={{ maxHeight: 240, overflow: "auto", background: "#0f172a", borderRadius: 6, padding: "8px 12px", fontFamily: "monospace", fontSize: 11 }}>
              {log.map((l, i) => (
                <div key={i} style={{ color: l.startsWith("✓") ? "#86efac" : l.startsWith("✗") ? "#fca5a5" : "#fcd34d", padding: "1px 0" }}>{l}</div>
              ))}
            </div>
            {fase === "fin" && (
              <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
                <button onClick={onClose} style={{ padding: "8px 20px", borderRadius: 6, border: "none", background: "#2563eb", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>Cerrar</button>
              </div>
            )}
          </>
        )}

      </div>
    </div>
  );
}

// ── Error Boundary: evita que un fallo de render deje la app en blanco ──
// ── Diálogo Asistente de Referencias ──
// Permite elegir un producto base y combinar opciones que se insertan en posiciones
// concretas de la referencia. Con la referencia final montada, consulta la tabla
// productos para traer PVP, nombre y descripción.
function AsistenteReferenciasDialog({ onClose, onInsertar, setStatus }) {
  const [productos, setProductos] = useState([]);
  const [cargandoProd, setCargandoProd] = useState(true);
  const [prodSel, setProdSel] = useState(null);        // producto seleccionado
  const [opciones, setOpciones] = useState([]);        // opciones del producto
  const [cargandoOpc, setCargandoOpc] = useState(false);
  const [opcionesSel, setOpcionesSel] = useState({});  // { posicion: opcion } una por posición

  const [refBase, setRefBase] = useState("");          // referencia base del producto
  const [refFinal, setRefFinal] = useState("");        // referencia montada

  // Datos consultados de la tabla productos
  const [datosBD, setDatosBD] = useState(null);        // producto elegido de las coincidencias | null
  const [buscandoBD, setBuscandoBD] = useState(false);
  const [noEncontrado, setNoEncontrado] = useState(false);
  const [coincidencias, setCoincidencias] = useState([]); // productos de BD cuya referencia empieza por refFinal
  const buscarSeq = useRef(0);

  // Cargar productos del asistente al abrir
  useEffect(() => {
    setCargandoProd(true);
    fetch(`${API_URL}/asistentes/productos`)
      .then(r => r.ok ? r.json() : [])
      .then(data => { setProductos(Array.isArray(data) ? data : []); })
      .catch(() => setStatus && setStatus("No se pudieron cargar los productos del asistente", "error"))
      .finally(() => setCargandoProd(false));
  }, []);

  // Al seleccionar producto: fijar referencia base y cargar opciones
  const seleccionarProducto = (p) => {
    setProdSel(p);
    setRefBase(p.referencia || "");
    setOpcionesSel({});
    setDatosBD(null);
    setNoEncontrado(false);
    setCargandoOpc(true);
    fetch(`${API_URL}/asistentes/opciones/${p.id}`)
      .then(r => r.ok ? r.json() : [])
      .then(data => {
        const arr = Array.isArray(data) ? [...data] : [];
        // Ordenar por posición y, dentro de la misma posición, por referencia
        arr.sort((a, b) => {
          // posición -1 (final) se muestra al final de la lista
          const pa = Number(a.posicion) === -1 ? Infinity : (Number(a.posicion) || 0);
          const pb = Number(b.posicion) === -1 ? Infinity : (Number(b.posicion) || 0);
          if (pa !== pb) return pa - pb;
          return String(a.referencia || "").localeCompare(String(b.referencia || ""));
        });
        setOpciones(arr);
      })
      .catch(() => setOpciones([]))
      .finally(() => setCargandoOpc(false));
  };

  // Alternar una opción. Una por cada posición fija; las de posición -1 (final) pueden ser varias.
  const toggleOpcion = (opc) => {
    setOpcionesSel(prev => {
      const next = { ...prev };
      const key = Number(opc.posicion) === -1 ? `fin:${opc.referencia}` : String(opc.posicion);
      if (next[key] && next[key].referencia === opc.referencia) {
        delete next[key]; // deseleccionar
      } else {
        next[key] = opc;
      }
      return next;
    });
    setDatosBD(null);
    setNoEncontrado(false);
  };

  // Montar la referencia final: cada opción SOBRESCRIBE los caracteres de la referencia
  // base a partir de su posición (1-based). posicion = -1 → se añade al final.
  // Ejemplo: base "3RT....-...." + pos4 "2025-" → "3RT2025-...." ; + pos9 "1" → "3RT2025-1..."
  useEffect(() => {
    let ref = refBase || "";
    const todas = Object.values(opcionesSel).filter(o => o && o.referencia);
    // 1) Opciones con posición fija: sobrescribir a partir de (posicion - 1)
    const fijas = todas.filter(o => Number(o.posicion) !== -1).sort((a, b) => (Number(a.posicion) || 0) - (Number(b.posicion) || 0));
    fijas.forEach(o => {
      const pos = Math.max(1, Number(o.posicion) || 1);
      const idx = pos - 1; // índice 0-based donde empieza a sobrescribir
      const opref = o.referencia;
      // Si la referencia base es más corta, se rellena con la opción a partir de idx
      const inicio = ref.slice(0, idx);
      const relleno = inicio.length < idx ? " ".repeat(idx - inicio.length) : "";
      const resto = ref.slice(idx + opref.length); // lo que queda tras los caracteres sobrescritos
      ref = inicio + relleno + opref + resto;
    });
    // 2) Opciones de final (posición -1): se añaden al final
    const finales = todas.filter(o => Number(o.posicion) === -1);
    finales.forEach(o => { ref = ref + o.referencia; });
    setRefFinal(ref);
  }, [refBase, opcionesSel]);

  // Búsqueda automática: cuando cambia la referencia montada, buscar productos cuya
  // referencia empiece por ella (igual que el autocompletado de la celda del grid).
  useEffect(() => {
    const ref = String(refFinal || "").trim();
    setDatosBD(null);
    setNoEncontrado(false);
    if (ref.length < 2) { setCoincidencias([]); return; }
    setBuscandoBD(true);
    const seq = ++buscarSeq.current;
    const t = setTimeout(async () => {
      try {
        const url = API_URL + "/productos/?busqueda=" + encodeURIComponent(ref) + "&campo=referencia&limite=15";
        const res = await fetch(url);
        if (!res.ok) throw new Error("HTTP " + res.status);
        const data = await res.json();
        if (seq !== buscarSeq.current) return;
        const refUp = ref.toUpperCase();
        // Solo las que EMPIEZAN por la referencia montada, ordenadas
        const empiezan = (Array.isArray(data) ? data : [])
          .filter(p => String(p.referencia || "").toUpperCase().startsWith(refUp))
          .sort((a, b) => String(a.referencia || "").localeCompare(String(b.referencia || "")));
        setCoincidencias(empiezan);
        setNoEncontrado(empiezan.length === 0);
      } catch {
        if (seq === buscarSeq.current) { setCoincidencias([]); }
      } finally {
        if (seq === buscarSeq.current) setBuscandoBD(false);
      }
    }, 250);
    return () => clearTimeout(t);
  }, [refFinal]);

  const insertar = () => {
    if (!datosBD) { setStatus && setStatus("Selecciona un producto de la lista de coincidencias", "error"); return; }
    onInsertar(datosBD);
    setStatus && setStatus(`Producto ${datosBD.referencia} insertado`, "success");
    onClose();
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "92%", maxWidth: 1000, height: "88vh", maxHeight: "94vh", minWidth: 560, minHeight: 400, display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.2)", resize: "both", overflow: "auto" }} onClick={e => e.stopPropagation()}>

        {/* Cabecera */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
          <h2 style={{ fontSize: 17, fontWeight: 700, color: "#171717", margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Bot} size={20} color="#2563eb" /> Asistente de Referencias
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "transparent", cursor: "pointer", color: "#64748b" }}>
            <Icon as={X} size={20} />
          </button>
        </div>

        {/* Dos listas */}
        <div style={{ display: "flex", gap: 14, flex: "1 1 auto", minHeight: 180, marginBottom: 14 }}>
          {/* Lista de productos */}
          <div style={{ flex: 1, display: "flex", flexDirection: "column", border: "1px solid #e2e8f0", borderRadius: 8, overflow: "hidden" }}>
            <div style={{ padding: "6px 10px", background: "#f8fafc", borderBottom: "1px solid #e2e8f0", fontSize: 12, fontWeight: 600, color: "#171717" }}>
              Productos {cargandoProd ? "(cargando...)" : `(${productos.length})`}
            </div>
            <div style={{ flex: 1, overflow: "auto" }}>
              {productos.map(p => (
                <div key={p.id} onClick={() => seleccionarProducto(p)}
                  style={{ padding: "6px 10px", cursor: "pointer", fontSize: 12, borderBottom: "1px solid #f1f5f9",
                    background: prodSel?.id === p.id ? "#dbeafe" : "transparent" }}>
                  <div style={{ fontWeight: 600, color: "#171717" }}>{p.nombre}</div>
                  <div style={{ fontFamily: "monospace", color: "#2563eb", fontSize: 13, fontWeight: 600 }}>{p.referencia}</div>
                </div>
              ))}
              {!cargandoProd && productos.length === 0 && (
                <div style={{ padding: 16, color: "#94a3b8", fontSize: 12, textAlign: "center" }}>No hay productos en el asistente</div>
              )}
            </div>
          </div>

          {/* Lista de opciones */}
          <div style={{ flex: 1, display: "flex", flexDirection: "column", border: "1px solid #e2e8f0", borderRadius: 8, overflow: "hidden" }}>
            <div style={{ padding: "6px 10px", background: "#f8fafc", borderBottom: "1px solid #e2e8f0", fontSize: 12, fontWeight: 600, color: "#171717" }}>
              Opciones {cargandoOpc ? "(cargando...)" : prodSel ? `(${opciones.length})` : ""}
            </div>
            <div style={{ flex: 1, overflow: "auto" }}>
              {!prodSel && <div style={{ padding: 16, color: "#94a3b8", fontSize: 12, textAlign: "center" }}>Selecciona un producto</div>}
              {opciones.map((o, i) => {
                const key = Number(o.posicion) === -1 ? `fin:${o.referencia}` : String(o.posicion);
                const seleccionada = opcionesSel[key] && opcionesSel[key].referencia === o.referencia;
                return (
                  <div key={i} onClick={() => toggleOpcion(o)}
                    style={{ padding: "6px 10px", cursor: "pointer", fontSize: 12, borderBottom: "1px solid #f1f5f9",
                      background: seleccionada ? "#dcfce7" : "transparent" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                      <span style={{ color: "#171717" }}>{o.descripcion}</span>
                      <span style={{ fontFamily: "monospace", color: "#2563eb", fontWeight: 600 }}>{o.referencia}</span>
                    </div>
                    <div style={{ color: "#94a3b8", fontSize: 10 }}>{Number(o.posicion) === -1 ? "Posición: final" : `Posición ${o.posicion}`}{seleccionada ? " · seleccionada" : ""}</div>
                  </div>
                );
              })}
              {prodSel && !cargandoOpc && opciones.length === 0 && (
                <div style={{ padding: 16, color: "#94a3b8", fontSize: 12, textAlign: "center" }}>Este producto no tiene opciones</div>
              )}
            </div>
          </div>
        </div>

        {/* Campos inferiores */}
        <div style={{ borderTop: "1px solid #e2e8f0", paddingTop: 12, flexShrink: 0 }}>
          <div style={{ marginBottom: 10 }}>
            <label style={{ fontSize: 11, color: "#525252", fontWeight: 500, display: "block", marginBottom: 3 }}>
              Referencia montada {buscandoBD ? "(buscando...)" : ""}
            </label>
            <input value={refFinal} onChange={e => { setRefBase(e.target.value); setOpcionesSel({}); }}
              style={{ width: "100%", padding: "6px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 14, fontFamily: "monospace", fontWeight: 700, boxSizing: "border-box" }} />
          </div>

          {/* Lista de coincidencias en BD (referencias que empiezan por la montada) */}
          <div style={{ marginBottom: 10 }}>
            <div style={{ fontSize: 11, color: "#525252", fontWeight: 600, marginBottom: 4 }}>
              Productos en la base de datos {coincidencias.length > 0 ? `(${coincidencias.length})` : ""}
            </div>
            <div style={{ maxHeight: 140, overflow: "auto", border: "1px solid #e2e8f0", borderRadius: 6 }}>
              {noEncontrado && !buscandoBD && (
                <div style={{ padding: 12, color: "#b91c1c", fontSize: 12, textAlign: "center" }}>
                  No hay productos cuya referencia empiece por "{refFinal}".
                </div>
              )}
              {!noEncontrado && coincidencias.length === 0 && (
                <div style={{ padding: 12, color: "#94a3b8", fontSize: 12, textAlign: "center" }}>
                  Selecciona opciones para montar la referencia y ver las coincidencias.
                </div>
              )}
              {coincidencias.map((p, i) => (
                <div key={i} onClick={() => setDatosBD(p)}
                  style={{ padding: "6px 10px", cursor: "pointer", fontSize: 12, borderBottom: "1px solid #f1f5f9",
                    background: datosBD && datosBD.id === p.id ? "#dbeafe" : "transparent",
                    display: "flex", justifyContent: "space-between", gap: 10 }}>
                  <span style={{ fontFamily: "monospace", fontWeight: 600, color: "#1e3a5f", whiteSpace: "nowrap" }}>{p.referencia}</span>
                  <span style={{ color: "#475569", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>{p.nombre || p.descripcion || ""}</span>
                  <span style={{ color: "#0369a1", fontWeight: 600, whiteSpace: "nowrap" }}>{fmtEur(Number(p.pvp) || 0)}</span>
                </div>
              ))}
            </div>
          </div>

          {/* Datos del producto seleccionado de la lista */}
          <div style={{ display: "grid", gridTemplateColumns: "120px 1fr", gap: 8, alignItems: "center", marginBottom: 6 }}>
            <span style={{ fontSize: 12, fontWeight: 600, color: "#525252" }}>Referencia</span>
            <span style={{ fontSize: 13, color: datosBD ? "#1e3a5f" : "#94a3b8", fontFamily: "monospace", fontWeight: 600 }}>{datosBD ? datosBD.referencia : "—"}</span>
            <span style={{ fontSize: 12, fontWeight: 600, color: "#525252" }}>PVP</span>
            <span style={{ fontSize: 13, color: datosBD ? "#0369a1" : "#94a3b8", fontWeight: 600 }}>{datosBD ? fmtEur(Number(datosBD.pvp) || 0) : "—"}</span>
            <span style={{ fontSize: 12, fontWeight: 600, color: "#525252" }}>Nombre</span>
            <span style={{ fontSize: 13, color: datosBD ? "#171717" : "#94a3b8" }}>{datosBD ? (datosBD.nombre || "—") : "—"}</span>
            <span style={{ fontSize: 12, fontWeight: 600, color: "#525252" }}>Descripción</span>
            <span style={{ fontSize: 13, color: datosBD ? "#475569" : "#94a3b8" }}>{datosBD ? (datosBD.descripcion || "—") : "—"}</span>
          </div>

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 12 }}>
            <button onClick={onClose} style={{ padding: "8px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>Cerrar</button>
            <button onClick={insertar} disabled={!datosBD}
              style={{ padding: "8px 16px", borderRadius: 6, border: "none", background: !datosBD ? "#cbd5e1" : "#16a34a", color: "#fff", cursor: !datosBD ? "default" : "pointer", fontSize: 13, fontWeight: 600 }}>
              Insertar en presupuesto
            </button>
          </div>
        </div>

      </div>
    </div>
  );
}

// ── Diálogo Gestionar Estrategias de Descuento ──
// Lista de estrategias (tabla descuentos) + su detalle (detalledescuentos con grupo descuento).
// Permite consultar, crear, editar, borrar y aplicar la estrategia al presupuesto.
function EstrategiasDescuentoDialog({ onClose, onAplicar, setStatus }) {
  const [estrategias, setEstrategias] = useState([]);
  const [cargando, setCargando] = useState(true);
  const [sel, setSel] = useState(null);            // estrategia seleccionada {id, nombre, descripcion}
  const [detalle, setDetalle] = useState([]);      // [{idgrupodescuento, grupodescuento, descripciongrupo, descuento}]
  const [cargandoDet, setCargandoDet] = useState(false);
  const [modo, setModo] = useState("ver");         // "ver" | "editar" | "nuevo"
  const [grupos, setGrupos] = useState([]);        // catálogo de grupos descuento para añadir
  const [editNombre, setEditNombre] = useState("");
  const [editDesc, setEditDesc] = useState("");
  const [editDetalle, setEditDetalle] = useState([]); // edición del detalle
  const [confirmBorrar, setConfirmBorrar] = useState(false);

  const cargarLista = () => {
    setCargando(true);
    fetch(`${API_URL}/descuentos/`)
      .then(r => r.ok ? r.json() : [])
      .then(data => setEstrategias(Array.isArray(data) ? data : []))
      .catch(() => setStatus && setStatus("No se pudieron cargar las estrategias", "error"))
      .finally(() => setCargando(false));
  };

  useEffect(() => { cargarLista(); }, []);
  // Catálogo de grupos descuento (para el editor)
  useEffect(() => {
    fetch(`${API_URL}/gruposdescuento/`).then(r => r.ok ? r.json() : []).then(d => setGrupos(Array.isArray(d) ? d : [])).catch(() => {});
  }, []);

  const seleccionar = (e) => {
    setSel(e); setModo("ver"); setCargandoDet(true);
    fetch(`${API_URL}/descuentos/${e.id}`)
      .then(r => r.ok ? r.json() : null)
      .then(data => { setDetalle(data ? data.detalle : []); })
      .catch(() => setDetalle([]))
      .finally(() => setCargandoDet(false));
  };

  const nuevaEstrategia = () => {
    setSel(null); setModo("nuevo"); setEditNombre(""); setEditDesc(""); setEditDetalle([]);
  };

  const editarEstrategia = () => {
    if (!sel) return;
    setModo("editar"); setEditNombre(sel.nombre || ""); setEditDesc(sel.descripcion || "");
    setEditDetalle(detalle.map(d => ({ idgrupodescuento: d.idgrupodescuento, grupodescuento: d.grupodescuento, descripciongrupo: d.descripciongrupo, descuento: d.descuento })));
  };

  const addLineaEdit = () => setEditDetalle(d => [...d, { idgrupodescuento: null, grupodescuento: "", descripciongrupo: "", descuento: 0 }]);
  const setLineaGrupo = (i, idg) => setEditDetalle(d => {
    const g = grupos.find(x => x.id === Number(idg));
    const next = [...d];
    next[i] = { ...next[i], idgrupodescuento: Number(idg), grupodescuento: g ? g.grupodescuentospain : "", descripciongrupo: g ? g.descripcion : "" };
    return next;
  });
  const setLineaDto = (i, val) => setEditDetalle(d => { const next = [...d]; next[i] = { ...next[i], descuento: val }; return next; });
  const delLineaEdit = (i) => setEditDetalle(d => d.filter((_, idx) => idx !== i));

  const guardar = async () => {
    if (!editNombre.trim()) { setStatus && setStatus("La estrategia necesita un nombre", "error"); return; }
    const payload = {
      nombre: editNombre.trim(),
      descripcion: editDesc.trim() || null,
      detalle: editDetalle.filter(l => l.idgrupodescuento).map(l => ({
        idgrupodescuento: l.idgrupodescuento,
        descuento: Number(String(l.descuento).replace(",", ".")) || 0,
      })),
    };
    try {
      const url = modo === "nuevo" ? `${API_URL}/descuentos/` : `${API_URL}/descuentos/${sel.id}`;
      const method = modo === "nuevo" ? "POST" : "PUT";
      const r = await fetch(url, { method, headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      if (!r.ok) throw new Error("HTTP " + r.status);
      setStatus && setStatus(modo === "nuevo" ? "Estrategia creada" : "Estrategia actualizada", "success");
      setModo("ver"); cargarLista();
      // Recargar el detalle si era edición
      if (modo === "editar" && sel) seleccionar(sel);
    } catch (e) {
      setStatus && setStatus("Error al guardar la estrategia: " + e.message, "error");
    }
  };

  const borrar = async () => {
    if (!sel) return;
    try {
      const r = await fetch(`${API_URL}/descuentos/${sel.id}`, { method: "DELETE" });
      if (!r.ok) throw new Error("HTTP " + r.status);
      setStatus && setStatus("Estrategia borrada", "success");
      setSel(null); setDetalle([]); setConfirmBorrar(false); cargarLista();
    } catch (e) {
      setStatus && setStatus("Error al borrar: " + e.message, "error");
    }
  };

  const aplicar = () => {
    if (!detalle || detalle.length === 0) { setStatus && setStatus("La estrategia no tiene descuentos", "error"); return; }
    onAplicar(detalle);
    onClose();
  };

  const lineasMostrar = modo === "ver" ? detalle : editDetalle;

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "92%", maxWidth: 980, height: "86vh", maxHeight: "94vh", minWidth: 560, minHeight: 400, display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.2)", resize: "both", overflow: "auto" }} onClick={e => e.stopPropagation()}>

        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
          <h2 style={{ fontSize: 17, fontWeight: 700, color: "#171717", margin: 0, display: "inline-flex", alignItems: "center", gap: 8 }}>
            <Icon as={Percent} size={20} color="#2563eb" /> Estrategias de Descuento
          </h2>
          <button onClick={onClose} style={{ border: "none", background: "transparent", cursor: "pointer", color: "#64748b" }}><Icon as={X} size={20} /></button>
        </div>

        <div style={{ display: "flex", gap: 14, flex: "1 1 auto", minHeight: 200 }}>
          {/* Lista de estrategias */}
          <div style={{ width: 280, display: "flex", flexDirection: "column", border: "1px solid #e2e8f0", borderRadius: 8, overflow: "hidden" }}>
            <div style={{ padding: "6px 10px", background: "#f8fafc", borderBottom: "1px solid #e2e8f0", fontSize: 12, fontWeight: 600, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span>Estrategias {cargando ? "(...)" : `(${estrategias.length})`}</span>
              <button onClick={nuevaEstrategia} title="Nueva estrategia" style={{ border: "none", background: "#2563eb", color: "#fff", borderRadius: 4, cursor: "pointer", padding: "2px 8px", fontSize: 11 }}>+ Nueva</button>
            </div>
            <div style={{ flex: 1, overflow: "auto" }}>
              {estrategias.map(e => (
                <div key={e.id} onClick={() => seleccionar(e)}
                  style={{ padding: "7px 10px", cursor: "pointer", fontSize: 12, borderBottom: "1px solid #f1f5f9", background: sel?.id === e.id ? "#dbeafe" : "transparent" }}>
                  <div style={{ fontWeight: 600, color: "#171717" }}>{e.nombre}</div>
                  {e.descripcion && <div style={{ color: "#64748b", fontSize: 11 }}>{e.descripcion}</div>}
                </div>
              ))}
              {!cargando && estrategias.length === 0 && <div style={{ padding: 16, color: "#94a3b8", fontSize: 12, textAlign: "center" }}>No hay estrategias</div>}
            </div>
          </div>

          {/* Detalle / editor */}
          <div style={{ flex: 1, display: "flex", flexDirection: "column", border: "1px solid #e2e8f0", borderRadius: 8, overflow: "hidden" }}>
            {(modo === "editar" || modo === "nuevo") ? (
              <div style={{ padding: "10px 12px", borderBottom: "1px solid #e2e8f0", display: "flex", gap: 8 }}>
                <input value={editNombre} onChange={e => setEditNombre(e.target.value)} placeholder="Nombre de la estrategia"
                  style={{ flex: 1, padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12 }} />
                <input value={editDesc} onChange={e => setEditDesc(e.target.value)} placeholder="Descripción"
                  style={{ flex: 1, padding: "5px 8px", border: "1px solid #d4d4d4", borderRadius: 4, fontSize: 12 }} />
              </div>
            ) : (
              <div style={{ padding: "8px 12px", background: "#f8fafc", borderBottom: "1px solid #e2e8f0", fontSize: 12, fontWeight: 600 }}>
                {sel ? `Descuentos de "${sel.nombre}"` : "Selecciona una estrategia"}
              </div>
            )}

            <div style={{ flex: 1, overflow: "auto" }}>
              {cargandoDet && <div style={{ padding: 16, color: "#94a3b8", fontSize: 12, textAlign: "center" }}>Cargando...</div>}
              {!cargandoDet && modo === "ver" && !sel && <div style={{ padding: 16, color: "#94a3b8", fontSize: 12, textAlign: "center" }}>Selecciona una estrategia para ver sus descuentos</div>}
              {!cargandoDet && (
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#fafafa" }}>
                    <tr>
                      <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5" }}>Grupo Dto.</th>
                      <th style={{ padding: "6px 8px", textAlign: "left", borderBottom: "1px solid #e5e5e5" }}>Descripción</th>
                      <th style={{ padding: "6px 8px", textAlign: "right", borderBottom: "1px solid #e5e5e5", width: 90 }}>Descuento</th>
                      {(modo === "editar" || modo === "nuevo") && <th style={{ width: 36, borderBottom: "1px solid #e5e5e5" }}></th>}
                    </tr>
                  </thead>
                  <tbody>
                    {lineasMostrar.map((d, i) => (
                      <tr key={i} style={{ borderBottom: "1px solid #f1f5f9" }}>
                        <td style={{ padding: "5px 8px" }}>
                          {(modo === "editar" || modo === "nuevo") ? (
                            <select value={d.idgrupodescuento || ""} onChange={e => setLineaGrupo(i, e.target.value)}
                              style={{ width: "100%", padding: "3px", fontSize: 11, border: "1px solid #d4d4d4", borderRadius: 4 }}>
                              <option value="">— grupo —</option>
                              {grupos.map(g => <option key={g.id} value={g.id}>{g.grupodescuentospain}</option>)}
                            </select>
                          ) : <span style={{ fontFamily: "monospace", fontWeight: 600, color: "#1e3a5f" }}>{d.grupodescuento || d.idgrupodescuento}</span>}
                        </td>
                        <td style={{ padding: "5px 8px", color: "#475569" }}>{d.descripciongrupo || ""}</td>
                        <td style={{ padding: "5px 8px", textAlign: "right" }}>
                          {(modo === "editar" || modo === "nuevo") ? (
                            <input value={d.descuento ?? ""} onChange={e => setLineaDto(i, e.target.value)}
                              style={{ width: 60, padding: "3px", fontSize: 11, textAlign: "right", border: "1px solid #d4d4d4", borderRadius: 4 }} />
                          ) : <span style={{ fontWeight: 600, color: "#0369a1" }}>{fmt(Number(d.descuento) || 0)} %</span>}
                        </td>
                        {(modo === "editar" || modo === "nuevo") && (
                          <td style={{ textAlign: "center" }}>
                            <button onClick={() => delLineaEdit(i)} style={{ border: "none", background: "transparent", cursor: "pointer", color: "#dc2626" }}><Icon as={Trash2} size={14} /></button>
                          </td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
              {(modo === "editar" || modo === "nuevo") && (
                <button onClick={addLineaEdit} style={{ margin: 10, padding: "5px 12px", fontSize: 12, border: "1px dashed #94a3b8", borderRadius: 6, background: "#f8fafc", color: "#475569", cursor: "pointer" }}>+ Añadir grupo descuento</button>
              )}
            </div>
          </div>
        </div>

        {/* Botonera */}
        <div style={{ display: "flex", justifyContent: "space-between", gap: 8, marginTop: 14, flexShrink: 0 }}>
          <div style={{ display: "flex", gap: 8 }}>
            {modo === "ver" && sel && (
              <>
                <button onClick={editarEstrategia} style={{ padding: "8px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>Editar</button>
                <button onClick={() => setConfirmBorrar(true)} style={{ padding: "8px 14px", borderRadius: 6, border: "1px solid #fca5a5", background: "#fff", color: "#b91c1c", cursor: "pointer", fontSize: 13 }}>Borrar</button>
              </>
            )}
            {(modo === "editar" || modo === "nuevo") && (
              <>
                <button onClick={guardar} style={{ padding: "8px 14px", borderRadius: 6, border: "none", background: "#16a34a", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>Guardar</button>
                <button onClick={() => { setModo("ver"); if (sel) seleccionar(sel); }} style={{ padding: "8px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>Cancelar</button>
              </>
            )}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={onClose} style={{ padding: "8px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>Cerrar</button>
            {modo === "ver" && sel && (
              <button onClick={aplicar} style={{ padding: "8px 16px", borderRadius: 6, border: "none", background: "#2563eb", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>Aplicar al presupuesto</button>
            )}
          </div>
        </div>

        {confirmBorrar && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100001 }} onClick={() => setConfirmBorrar(false)}>
            <div style={{ background: "#fff", borderRadius: 10, padding: "1.5rem", maxWidth: 380, textAlign: "center" }} onClick={e => e.stopPropagation()}>
              <p style={{ fontSize: 14, color: "#171717", marginBottom: 16 }}>¿Borrar la estrategia "{sel?.nombre}"?</p>
              <div style={{ display: "flex", justifyContent: "center", gap: 8 }}>
                <button onClick={() => setConfirmBorrar(false)} style={{ padding: "8px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", cursor: "pointer", fontSize: 13 }}>Cancelar</button>
                <button onClick={borrar} style={{ padding: "8px 16px", borderRadius: 6, border: "none", background: "#dc2626", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>Borrar</button>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}

class ErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }
  componentDidCatch(error, info) {
    // Registrar para diagnóstico en consola
    console.error("[ErrorBoundary] Se ha capturado un error de render:", error, info);
  }
  reset = () => this.setState({ hasError: false, error: null });
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ position: "fixed", inset: 0, background: "#f8fafc", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200000, padding: 24 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "2rem", maxWidth: 520, boxShadow: "0 20px 60px rgba(0,0,0,0.2)", textAlign: "center" }}>
            <div style={{ fontSize: 18, fontWeight: 700, color: "#b91c1c", marginBottom: 10 }}>Se ha producido un error</div>
            <p style={{ fontSize: 13, color: "#475569", lineHeight: 1.6, marginBottom: 8 }}>
              La aplicación encontró un problema al dibujar la pantalla, pero tus datos no se han perdido. Pulsa "Reintentar" para volver a la edición.
            </p>
            <p style={{ fontSize: 11, color: "#94a3b8", fontFamily: "monospace", background: "#f1f5f9", borderRadius: 6, padding: "6px 10px", marginBottom: 16, wordBreak: "break-word" }}>
              {this.state.error ? String(this.state.error.message || this.state.error) : "Error desconocido"}
            </p>
            <button onClick={this.reset}
              style={{ padding: "9px 24px", borderRadius: 6, border: "none", background: "#2563eb", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
              Reintentar
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

function AppInner() {
  // Reset global de html/body para que no aparezca scroll de la página completa
  useEffect(() => {
    const styleEl = document.createElement("style");
    styleEl.innerHTML = `
      html, body, #root { margin: 0; padding: 0; height: 100%; overflow: hidden; }
      *, *::before, *::after { box-sizing: border-box; }
    `;
    document.head.appendChild(styleEl);
    return () => { document.head.removeChild(styleEl); };
  }, []);

  // Cargar países al iniciar (para selector "País cliente final")
  useEffect(() => {
    fetch(`${API_URL}/paises/`).then(r => r.ok ? r.json() : []).then(data => {
      if (Array.isArray(data)) setPaisesList(data);
    }).catch(() => {});
  }, []);


    // Pantalla de bienvenida una vez por sesión
  // Datos del usuario actual: usuario (nombre) y codigoUsuario (codigopresupuestos, para construir el nº completo)
  const [usuarioActual, setUsuarioActual] = useState(""); // Nombre del usuario en la barra (vacío hasta login OK)
  const [codigoUsuario, setCodigoUsuario] = useState("CDM"); // Código de presupuestos del usuario activo (se carga al hacer login)
  const [userMenuOpen, setUserMenuOpen] = useState(false);
  const [showLoginDialog, setShowLoginDialog] = useState(false);
  const [showWelcome, setShowWelcome] = useState(() => {
    try { return sessionStorage.getItem("welcomeShown") !== "1"; } catch { return true; }
  });
  const [rows, setRows] = useState(() => {
    const saved = cargarPresupuestoLocal();
    return (saved && Array.isArray(saved.rows) && saved.rows.length > 0) ? saved.rows : INITIAL_ROWS;
  });
  const [selectedCell, setSelectedCell] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const escapeEditRef = useRef(false); // true cuando se pulsa Escape para descartar la edición sin guardar
  const [editValue, setEditValue] = useState("");
  const [selectedRows, setSelectedRows] = useState(new Set());
  const [selectionRange, setSelectionRange] = useState(null); // { startRowIdx, startColIdx, endRowIdx, endColIdx }
  const [isSelecting, setIsSelecting] = useState(false);

  // Listener global para detectar mouseup fuera de la tabla
  useEffect(() => {
    const onUp = () => setIsSelecting(false);
    window.addEventListener("mouseup", onUp);
    return () => window.removeEventListener("mouseup", onUp);
  }, []);

  // Auto-scroll del grid mientras se selecciona con el ratón (estilo Excel)
  // Funciona en vertical (arriba/abajo) y horizontal (izquierda/derecha)
  useEffect(() => {
    if (!isSelecting) return;
    let raf = null;
    let speedY = 0; // px/frame: positivo = abajo, negativo = arriba
    let speedX = 0; // px/frame: positivo = derecha, negativo = izquierda
    const onMove = (e) => {
      const el = tableContainerRef.current;
      if (!el) return;
      const rect = el.getBoundingClientRect();
      const EDGE = 40; // px desde el borde donde se activa el scroll
      const MAX = 25;  // velocidad máxima en px/frame
      // Vertical
      if (e.clientY > rect.bottom - EDGE) {
        const dist = Math.min(EDGE, e.clientY - (rect.bottom - EDGE));
        speedY = Math.ceil((dist / EDGE) * MAX);
      } else if (e.clientY < rect.top + EDGE) {
        const dist = Math.min(EDGE, (rect.top + EDGE) - e.clientY);
        speedY = -Math.ceil((dist / EDGE) * MAX);
      } else {
        speedY = 0;
      }
      // Horizontal
      if (e.clientX > rect.right - EDGE) {
        const dist = Math.min(EDGE, e.clientX - (rect.right - EDGE));
        speedX = Math.ceil((dist / EDGE) * MAX);
      } else if (e.clientX < rect.left + EDGE) {
        const dist = Math.min(EDGE, (rect.left + EDGE) - e.clientX);
        speedX = -Math.ceil((dist / EDGE) * MAX);
      } else {
        speedX = 0;
      }
    };
    const tick = () => {
      const el = tableContainerRef.current;
      if (el) {
        if (speedY !== 0) el.scrollTop += speedY;
        if (speedX !== 0) el.scrollLeft += speedX;
      }
      raf = requestAnimationFrame(tick);
    };
    window.addEventListener("mousemove", onMove);
    raf = requestAnimationFrame(tick);
    return () => {
      window.removeEventListener("mousemove", onMove);
      if (raf) cancelAnimationFrame(raf);
    };
  }, [isSelecting]);

  // ── Pegar (paste) desde el portapapeles al grid ──
  useEffect(() => {
    const onPaste = (e) => {
      // Si estamos editando una celda, dejar al input gestionar el paste
      if (editingCell) return;
      // Si el target del paste es un input/textarea (cliente, título, etc), dejarlo
      const tag = (e.target?.tagName || "").toLowerCase();
      if (tag === "input" || tag === "textarea" || tag === "select") return;

      const text = (e.clipboardData || window.clipboardData)?.getData("text");
      if (!text) {
        console.log("[Paste] No hay texto en el portapapeles");
        return;
      }

      // Si no hay celda seleccionada, avisar al usuario
      if (!selectedCell) {
        e.preventDefault();
        setStatus("Para pegar, primero haz click en una celda del grid (verás un rectángulo azul)", "warning");
        return;
      }

      e.preventDefault();
      console.log("[Paste] Pegando en", selectedCell, "longitud texto:", text.length);

      // Detectar si es tabular: contiene saltos de línea o tabs
      const lines = text.replace(/\r\n/g, "\n").split("\n");
      // Quitar última línea vacía si existe (típico al copiar de Excel)
      while (lines.length > 0 && lines[lines.length - 1].trim() === "") lines.pop();
      if (lines.length === 0) return;

      const matrix = lines.map(l => l.includes("\t") ? l.split("\t") : [l]);
      const isTabular = matrix.length > 1 || matrix[0].length > 1;

      // Localizar fila y columna actuales
      const rowIdx = rows.findIndex(r => r.id === selectedCell.rowId);
      const colIdx = COLUMNS.findIndex(c => c.key === selectedCell.colKey);
      if (rowIdx === -1 || colIdx === -1) return;

      // Helper para convertir string a tipo correcto
      // Maneja formatos europeo "3.948,45 €" y americano "3,948.45"
      const parseNumeroFlex = (valor) => {
        let s = String(valor).trim();
        if (s === "") return NaN;
        // Quitar símbolos de moneda y espacios
        s = s.replace(/[€$£¥\s]/g, "");
        const tieneComa = s.includes(",");
        const tienePunto = s.includes(".");
        if (tieneComa && tienePunto) {
          // El último separador es el decimal; el otro es separador de miles
          const lastComa = s.lastIndexOf(",");
          const lastPunto = s.lastIndexOf(".");
          if (lastComa > lastPunto) {
            // Europeo: "3.948,45" → quitar puntos, coma a punto
            s = s.replace(/\./g, "").replace(",", ".");
          } else {
            // Americano: "3,948.45" → quitar comas
            s = s.replace(/,/g, "");
          }
        } else if (tieneComa) {
          // Solo coma: la tratamos como decimal (formato europeo)
          // Salvo que haya más de una (entonces son miles): "1,234,567"
          const n = (s.match(/,/g) || []).length;
          if (n > 1) s = s.replace(/,/g, "");
          else s = s.replace(",", ".");
        }
        // Si solo hay puntos lo dejamos: parseFloat ya entiende "3.45" o "3.948" (este último se interpreta como número entero 3948 si es separador de miles, pero parseFloat lo lee como 3.948 = decimal; en ese caso ambiguo respetamos la lectura natural)
        return parseFloat(s);
      };
      const convertir = (col, valor) => {
        if (col.type === "calc") return null; // ignorar calculadas
        if (col.key === "cantidad") {
          const n = parseNumeroFlex(valor);
          return isNaN(n) ? 0 : Math.round(n);
        }
        if (col.type === "number") {
          const n = parseNumeroFlex(valor);
          return isNaN(n) ? 0 : n;
        }
        return String(valor);
      };

      setRows(rs => {
        const next = [...rs];
        let pegadas = 0, omitidas = 0;
        for (let i = 0; i < matrix.length; i++) {
          const r = rowIdx + i;
          if (r >= next.length) break; // no más filas disponibles
          const fila = { ...next[r] };
          for (let j = 0; j < matrix[i].length; j++) {
            const c = colIdx + j;
            if (c >= COLUMNS.length) break;
            const col = COLUMNS[c];
            if (col.type === "calc") { omitidas++; continue; }
            const valorConvertido = convertir(col, matrix[i][j]);
            if (valorConvertido !== null) {
              fila[col.key] = valorConvertido;
              pegadas++;
            }
          }
          next[r] = fila;
        }
        setTimeout(() => {
          if (isTabular) {
            setStatus(`Pegadas ${pegadas} celdas (${matrix.length} fila${matrix.length > 1 ? "s" : ""} × ${matrix[0].length} col${matrix[0].length > 1 ? "s" : ""})${omitidas > 0 ? `, ${omitidas} columnas calculadas omitidas` : ""}`, "success");
          } else {
            setStatus(`Valor pegado en ${selectedCell.colKey} fila ${rowIdx + 1}`, "success");
          }
        }, 0);
        return next;
      });
    };

    window.addEventListener("paste", onPaste);
    return () => window.removeEventListener("paste", onPaste);
  }, [selectedCell, editingCell, rows]);

  // ── Copiar (copy) celda/rango seleccionado al portapapeles ──
  useEffect(() => {
    const onCopy = (e) => {
      // Si se está editando una celda, dejar al input gestionar el copy normal
      if (editingCell) return;
      const tag = (e.target?.tagName || "").toLowerCase();
      if (tag === "input" || tag === "textarea" || tag === "select") return;

      // Si el usuario tiene texto seleccionado con el ratón, dejarlo
      const sel = window.getSelection?.();
      if (sel && sel.toString && sel.toString().length > 0) return;

      let texto = "";

      if (selectionRange) {
        const minRow = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
        const maxRow = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
        const minCol = Math.min(selectionRange.startColIdx, selectionRange.endColIdx);
        const maxCol = Math.max(selectionRange.startColIdx, selectionRange.endColIdx);
        const lines = [];
        for (let r = minRow; r <= maxRow; r++) {
          const row = rows[r];
          if (!row) continue;
          const cells = [];
          for (let c = minCol; c <= maxCol; c++) {
            const col = COLUMNS[c];
            if (!col) { cells.push(""); continue; }
            let val = row[col.key];
            if (col.key === "precionetounitario") val = calcNetoUnit(row);
            else if (col.key === "precionetoposicion") val = calcNetoPos(row);
            else if (col.key === "costeposicion") val = (Number(row.preciocosteunitario) || 0) * (Number(row.cantidad) || 0);
            else if (col.key === "margen") val = calcMargen(row);
            if (val === null || val === undefined) val = "";
            cells.push(String(val));
          }
          lines.push(cells.join("\t"));
        }
        texto = lines.join("\n");
      } else if (selectedCell) {
        const row = rows.find(r => r.id === selectedCell.rowId);
        if (!row) return;
        const col = COLUMNS.find(c => c.key === selectedCell.colKey);
        if (!col) return;
        let val = row[col.key];
        if (col.key === "precionetounitario") val = calcNetoUnit(row);
        else if (col.key === "precionetoposicion") val = calcNetoPos(row);
        else if (col.key === "costeposicion") val = (Number(row.preciocosteunitario) || 0) * (Number(row.cantidad) || 0);
        else if (col.key === "margen") val = calcMargen(row);
        if (val === null || val === undefined) val = "";
        texto = String(val);
      } else {
        return;
      }

      e.preventDefault();
      if (e.clipboardData) {
        e.clipboardData.setData("text/plain", texto);
      } else if (navigator.clipboard) {
        navigator.clipboard.writeText(texto).catch(() => {});
      }

      if (selectionRange) {
        const rows_n = Math.abs(selectionRange.endRowIdx - selectionRange.startRowIdx) + 1;
        const cols_n = Math.abs(selectionRange.endColIdx - selectionRange.startColIdx) + 1;
        setStatus(`Copiado ${rows_n}×${cols_n} (${rows_n * cols_n} celdas) al portapapeles`, "success");
      } else {
        setStatus(`Copiada celda ${selectedCell.colKey} al portapapeles`, "success");
      }
    };

    window.addEventListener("copy", onCopy);
    return () => window.removeEventListener("copy", onCopy);
  }, [selectedCell, selectionRange, editingCell, rows]);

  // Al abrir, si no hay número de presupuesto, pedir a la API el siguiente número
  useEffect(() => {
    if (presupuesto.np) return; // ya hay número (cargado o asignado)
    let cancel = false;
    (async () => {
      setStatus("Obteniendo siguiente número de presupuesto...", "working");
      try {
        const ano = presupuesto.anopresupuesto || getAnoFiscal();
        const sig = await obtenerSiguienteNumero(ano);
        if (cancel) return;
        setPresupuesto(p => {
          const newP = { ...p, np: String(sig), anopresupuesto: ano };
          newP.numerocompleto = buildNumeroCompleto(codigoUsuario, newP.np, newP.anopresupuesto);
          return newP;
        });
        setStatus(`Siguiente número asignado: ${buildNumeroCompleto(codigoUsuario, sig, ano)}`, "success");
      } catch (e) {
        if (cancel) return;
        setStatus("No se ha podido obtener el siguiente número de la BD: " + e.message, "error");
      }
    })();
    return () => { cancel = true; };
  }, []);
  const [modal, setModal] = useState(null);
  const [numFilas, setNumFilas] = useState(1);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [confirmClearRows, setConfirmClearRows] = useState(false);
  const [vista, setVista] = useState("grid");
  const [donutTooltip, setDonutTooltip] = useState(null);
  // Anchos y altos persistentes: una vez se aplica estructura o el usuario redimensiona,
  // estos valores se mantienen aunque se desactive estructura
  const [anchosManual, setAnchosManual] = useState({}); // { colKey: number }
  const [altosManual, setAltosManual] = useState({});   // { rowId: number }
  // Una vez se ha aplicado estructura, mantenemos el wrap multilínea en ref/producto/desc
  // aunque se desactive estructura (igual que persistimos los anchos)
  const [wrapPersistente, setWrapPersistente] = useState(false);
  const [filasCopiadas, setFilasCopiadas] = useState([]); // array de objetos fila (sin id)
  const [leerPmdDialog, setLeerPmdDialog] = useState(null); // {referencia, rowIdx}
  const [paisesList, setPaisesList] = useState([]); // [{id, codigo_iso2, pais, prefijo_telefonico}]
  const [apiLocalViva, setApiLocalViva] = useState(null); // null = sin comprobar, true = viva, false = no responde
  // Mapas nombre→descripción para los tooltips de las columnas grupodescuento/familia/subfamilia
  const [descGrupos, setDescGrupos] = useState({});       // { "NOMBRE": "descripción" }
  const [descFamilias, setDescFamilias] = useState({});
  const [descSubfamilias, setDescSubfamilias] = useState({});
  const [showApiLocalAviso, setShowApiLocalAviso] = useState(false); // diálogo de aviso cuando la API local no responde
  // Configuraciones varias (persistidas en fichero JSON que el usuario guarda/lee)
  const [configVarias, setConfigVarias] = useState({
    sqDestinatario: "simplequote-rfq.industry@siemens.com",
    sqCC: "josecarlos.demena@siemens.com",
  });
  // Estado durante drag de redimensión
  const resizingRef = useRef(null); // { type:'col'|'row', key, startPos, startSize }
  const [estructuraActiva, setEstructuraActiva] = useState(false);
  const [showImportar, setShowImportar] = useState(false);
  const [showDescuentos, setShowDescuentos] = useState(false);
  const [showEstrategias, setShowEstrategias] = useState(false);
  const [showLeer, setShowLeer] = useState(false);
  const [showClientes, setShowClientes] = useState(false);
  const [showContactos, setShowContactos] = useState(false);
  const [showSelecContacto, setShowSelecContacto] = useState(false);
  const [showSelecCliente, setShowSelecCliente] = useState(false);
  const [siemensDialog, setSiemensDialog] = useState(null); // null o { celdasTrabajadas: [...] }
  const [juntarDupConfirm, setJuntarDupConfirm] = useState(null); // {duplicados, totalFilasAfectadas, totalGrupos, rowIni, rowFin}
  const [showSeleccionar, setShowSeleccionar] = useState(false);
  const [showFijarPrecio, setShowFijarPrecio] = useState(null); // null | "total" | "celdas"
  const [showBorrarVacias, setShowBorrarVacias] = useState(false);
  const [infoCeldaDialog, setInfoCeldaDialog] = useState(null); // { caracteres, bg, color, colKey, fila } | null
  const [showGuardarElem, setShowGuardarElem] = useState(false);
  const [showLeerElem, setShowLeerElem] = useState(false);
  const [showLeerProducto, setShowLeerProducto] = useState(false);
  const [showAsistente, setShowAsistente] = useState(false);
  const [actualizarProductos, setActualizarProductos] = useState(null); // { existentes: [{fila, productoBD}], gruposCache } | null
  const [confirmBorrarPresup, setConfirmBorrarPresup] = useState(false);
  const [comprobarDialog, setComprobarDialog] = useState(null); // { tipo: "iguales" | "diferentes" | "no_existe", diffs: [], comparado: presupuestoBD }
  const [confirmSobreescribir, setConfirmSobreescribir] = useState(false);
  const [guardandoPresup, setGuardandoPresup] = useState(false);
  const [estilos, setEstilos] = useState(cargarEstilos);
  const [descCorta, setDescCorta] = useState(false);
  const [ayudaSeccion, setAyudaSeccion] = useState("intro");

  // ── StatusBar ──
  const [statusMessage, setStatusMessage] = useState({ text: "Listo", type: "info", timestamp: null });
  // Helper para escribir en la barra de estado. type: info | working | success | error
  const setStatus = useCallback((text, type = "info") => {
    const ts = Date.now();
    setStatusMessage({ text, type, timestamp: ts });
    // Auto-borrar mensajes success/error después de 5 segundos
    if (type === "success" || type === "error") {
      setTimeout(() => {
        setStatusMessage(s => s.timestamp === ts ? { text: "Listo", type: "info", timestamp: null } : s);
      }, 5000);
    }
  }, []);

  // Cargar descripciones de grupos descuento, familias y subfamilias (para tooltips del grid)
  useEffect(() => {
    const norm = s => String(s || "").trim().toUpperCase();
    fetch(`${API_URL}/gruposdescuento/`).then(r => r.ok ? r.json() : []).then(data => {
      if (Array.isArray(data)) {
        const m = {};
        data.forEach(g => { if (g.grupodescuentospain) m[norm(g.grupodescuentospain)] = g.descripcion || ""; });
        setDescGrupos(m);
      }
    }).catch(() => {});
    fetch(`${API_URL}/familias/`).then(r => r.ok ? r.json() : []).then(data => {
      if (Array.isArray(data)) {
        const m = {};
        data.forEach(f => { if (f.familia) m[norm(f.familia)] = f.descripcion || ""; });
        setDescFamilias(m);
      }
    }).catch(() => {});
    fetch(`${API_URL}/subfamilias/`).then(r => r.ok ? r.json() : []).then(data => {
      if (Array.isArray(data)) {
        const m = {};
        data.forEach(sf => { if (sf.subfamilia) m[norm(sf.subfamilia)] = sf.descripcion || ""; });
        setDescSubfamilias(m);
      }
    }).catch(() => {});
  }, []);

  // Comprobar al iniciar que la API local (CdM_Presupuestos_API_local.py) está viva
  useEffect(() => {
    fetch(`${API_LOCAL_URL}/`)
      .then(r => r.ok ? r.json() : null)
      .then(data => {
        const viva = !!(data && data.estado === "ok");
        setApiLocalViva(viva);
        if (viva) {
          setStatus(`API local conectada (${data.servicio || "CdM Presupuestos API local"})`, "success");
        } else {
          setStatus("La API local respondió pero con un estado inesperado", "warning");
          setShowApiLocalAviso(true);
        }
      })
      .catch(() => {
        setApiLocalViva(false);
        setStatus("API local no disponible: funciones PMD, SimpleQuote y Damex no funcionarán hasta arrancarla", "warning");
        setShowApiLocalAviso(true);
      });
  }, [setStatus]);
  const [presupuesto, setPresupuesto] = useState(() => {
    const saved = cargarPresupuestoLocal();
    if (saved && saved.presupuesto) return saved.presupuesto;
    return { id: "", np: "", numerocompleto: "", anopresupuesto: getAnoFiscal(), revision: "0", titulo: "", cliente: "", idcliente: null, alaatencion: null, alaatencion_nombre: "", alaatencion_cliente: "" };
  });
  const nextId = useRef(Math.max(...rows.map(r => r.id || 0), 0) + 1);
  const tableContainerRef = useRef(null);

  const apartados = calcApartados(rows);
  const anchosAuto = estructuraActiva ? calcAnchosAuto(rows, apartados) : {};
  // Columnas que se ajustan al contenido cuando estructura activa
  const COLS_FIT = ["representacion", "naturaleza", "cantidad", "pvp", "dtoaplicado", "precionetounitario", "precionetoposicion", "posicion"];
  // Columnas que pasan a multilínea (wrap) cuando estructura activa (salvo desc. corta)
  const COLS_WRAP_ESTRUCTURA = ["referencia", "nombre", "descripcion"];
  // Función que da el ancho efectivo de una columna.
  // Prioridad: 1) ancho manual del usuario, 2) ancho auto si estructura activa, 3) ancho base
  const getColWidth = (col) => {
    if (anchosManual[col.key] != null) return anchosManual[col.key];
    if (estructuraActiva && COLS_FIT.includes(col.key) && anchosAuto[col.key]) return anchosAuto[col.key];
    return col.width;
  };
  // Cuando se activa estructura, sobreescribimos los anchos manuales con los auto calculados
  // y volcamos las etiquetas automáticas de S1-S4/TT al campo nombre de cada fila.
  // Estos cambios se persisten en rows para que se mantengan al desactivar estructura.
  useEffect(() => {
    if (estructuraActiva) {
      setAnchosManual(prev => {
        const next = { ...prev };
        COLS_FIT.forEach(key => {
          if (anchosAuto[key]) next[key] = anchosAuto[key];
        });
        return next;
      });
      setWrapPersistente(true);
      // Sobrescribir nombre de subtotales y total, y persistir la numeración de apartado
      const tituloPres = String(presupuesto.titulo || "").trim();
      setRows(rs => {
        const aps = calcApartados(rs); // numeración jerárquica con las rows actuales
        let cambios = false;
        const next = rs.map((row, idx) => {
          const nat = row.naturaleza;
          let nuevoNombre = null;
          if (["S1","S2","S3","S4"].includes(nat)) {
            const nivel = nat[1];
            const tituloNat = "T" + nivel;
            let tituloTexto = "";
            for (let k = idx - 1; k >= 0; k--) {
              if (rs[k] && rs[k].naturaleza === tituloNat) {
                tituloTexto = String(rs[k].nombre || "").trim();
                break;
              }
            }
            nuevoNombre = tituloTexto ? `TOTAL ${tituloTexto}` : "TOTAL";
          } else if (nat === "TT") {
            nuevoNombre = tituloPres ? `TOTAL ${tituloPres}` : "TOTAL";
          }
          // Apartado a persistir: el calculado (vacío para S1-S4/TT/CM)
          const nuevoApartado = aps[row.id] != null ? aps[row.id] : (row.posicion || "");
          const cambiaNombre = nuevoNombre != null && nuevoNombre !== row.nombre;
          const cambiaApartado = nuevoApartado !== (row.posicion || "");
          // Al aplicar estructura se limpia la marca de PVP vencido (rojo)
          const limpiaVencido = !!row.pvpVencido;
          if (cambiaNombre || cambiaApartado || limpiaVencido) {
            cambios = true;
            return {
              ...row,
              nombre: cambiaNombre ? nuevoNombre : row.nombre,
              posicion: nuevoApartado,
              pvpVencido: false,
            };
          }
          return row;
        });
        return cambios ? next : rs;
      });
    }
  }, [estructuraActiva]);

  // ─ Drag de redimensión columna/fila ─
  const onResizeStart = (e, type, key, currentSize) => {
    e.preventDefault();
    e.stopPropagation();
    const startPos = type === "col" ? e.clientX : e.clientY;
    resizingRef.current = { type, key, startPos, startSize: currentSize };
    const onMove = (ev) => {
      if (!resizingRef.current) return;
      const { type: t, key: k, startPos: sp, startSize: ss } = resizingRef.current;
      const pos = t === "col" ? ev.clientX : ev.clientY;
      const delta = pos - sp;
      const nuevo = Math.max(24, ss + delta);
      if (t === "col") setAnchosManual(prev => ({ ...prev, [k]: nuevo }));
      else setAltosManual(prev => ({ ...prev, [k]: nuevo }));
    };
    const onUp = () => {
      resizingRef.current = null;
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    document.body.style.cursor = type === "col" ? "col-resize" : "row-resize";
    document.body.style.userSelect = "none";
  };

  // Busca el presupuesto en BD por número y revisión actuales
  const buscarPresupuestoEnBD = async () => {
    // Identificamos un presupuesto en BD por: numerocompleto + revision
    const numComp = String(presupuesto.numerocompleto || "").trim();
    if (!numComp) return null;
    try {
      // 1) Buscar por numerocompleto
      const res = await fetch(`${API_URL}/presupuestos/?busqueda=${encodeURIComponent(numComp)}`);
      if (!res.ok) throw new Error("Error " + res.status);
      const lista = await res.json();
      // Filtrar por numerocompleto exacto y revisión
      const enc = lista.find(p =>
        String(p.numerocompleto || "").trim() === numComp
        && String(p.revision ?? 0) === String(presupuesto.revision ?? 0)
      );
      if (!enc) return null;
      // 2) Leer su contenido completo
      const res2 = await fetch(`${API_URL}/presupuestos/${enc.id}/completo`);
      if (!res2.ok) throw new Error("Error " + res2.status);
      return await res2.json();
    } catch (e) {
      throw e;
    }
  };

  // Guarda el presupuesto en BD (POST o PUT según exista)
  const guardarPresupuestoEnBD = async (idExistente) => {
    const detalle = rows
      .filter(r => r.naturaleza || r.referencia || r.nombre || r.descripcion || (r.cantidad && r.cantidad !== 0))
      .map((r, i) => {
        const pvp = Number(r.pvp) || 0;
        const dto = Number(r.dtoaplicado) || 0;
        const cant = parseInt(r.cantidad, 10) || 0;
        const netoUnit = pvp * (1 - dto / 100);
        const netoPos = netoUnit * cant;
        return {
          posicion: i + 1,
          representacion: r.representacion || "",
          naturaleza: r.naturaleza || "",
          cantidad: cant,
          referencia: r.referencia || "",
          nombre: r.nombre || "",
          pvp,
          dtoaplicado: dto,
          precionetounitario: Math.round(netoUnit * 10000) / 10000,
          precionetoposicion: Math.round(netoPos * 10000) / 10000,
          descripcion: r.descripcion || "",
          familia: r.familia || "",
          subfamilia: r.subfamilia || "",
          preciocosteunitario: Number(r.preciocosteunitario) || 0,
          idposicion: (() => {
            const v = String(r.idposicion || "").trim();
            if (!v) return null;
            const n = parseInt(v, 10);
            return isNaN(n) ? null : n;
          })(),
          imagen: r.imagen || "",
          precionetounitario2: Number(r.precionetounitario2) || 0,
          grupodescuento: r.grupodescuento || "",
        };
      });
    // Total del presupuesto = suma de los neto posición de líneas tipo producto (PD, PE, E)
    const NATS_SUMA = new Set(["PD", "PE", "E"]);
    const totalPres = detalle
      .filter(d => NATS_SUMA.has(String(d.naturaleza || "").toUpperCase()))
      .reduce((s, d) => s + (Number(d.precionetoposicion) || 0), 0);
    const body = {
      numero: parseInt(presupuesto.np, 10) || null,
      numerocompleto: presupuesto.numerocompleto || presupuesto.np || "",
      revision: parseInt(presupuesto.revision, 10) || 0,
      titulo: presupuesto.titulo || "",
      idcliente: presupuesto.idcliente || null,
      anopresupuesto: parseInt(presupuesto.anopresupuesto, 10) || getAnoFiscal(),
      alaatencion: presupuesto.alaatencion || null,
      totalpresupuesto: Math.round(totalPres * 100) / 100,
      idpaisclientefinal: presupuesto.idpaisclientefinal || null,
      clientefinal: presupuesto.clientefinal || null,
      detalle,
    };
    const url = idExistente ? `${API_URL}/presupuestos/${idExistente}` : `${API_URL}/presupuestos/`;
    const method = idExistente ? "PUT" : "POST";
    const res = await fetch(url, {
      method,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => null);
      throw new Error(err?.detail || `Error ${res.status}`);
    }
    return await res.json();
  };

  const handleAction = useCallback((action, label) => {
    if (action === "Resumen") { setVista("resumen"); setStatus("Mostrando resumen del presupuesto", "info"); return; }
    if (action === "Ayuda") { setVista("ayuda"); setStatus("Mostrando pantalla de ayuda", "info"); return; }
    if (action === "Opciones") { setVista("opciones"); setStatus("Configurando opciones", "info"); return; }
    if (action === "AplicarEstructura") {
      setEstructuraActiva(e => {
        setStatus(e ? "Estructura desactivada" : "Estructura aplicada al presupuesto", "success");
        return !e;
      });
      return;
    }
    if (action === "Imprimir") {
      setStatus("Generando fichero Excel para impresión...", "working");
      try {
        exportToExcel(presupuesto, rows, apartados, estructuraActiva, estilos);
        setStatus("Fichero Excel generado y descargado correctamente", "success");
      } catch (err) {
        setStatus("Error al generar el Excel: " + (err.message || err), "error");
      }
      return;
    }
    if (action === "FormatoSimpleQuote") {
      setStatus("Generando fichero Simple Quote...", "working");
      try {
        // Solo productos: PD, PE, E (sin títulos, subtotales ni comentarios)
        const productos = rows.filter(r => ["PD", "PE", "E"].includes(r.naturaleza) && r.referencia && String(r.referencia).trim() !== "");
        if (productos.length === 0) {
          setStatus("No hay productos en el presupuesto para exportar", "error");
          return;
        }
        const calcNetoUnit = (row) => (Number(row.pvp) || 0) * (1 - (Number(row.dtoaplicado) || 0) / 100);
        // Construir AOA: A=referencia, B=cantidad, C=null, D=null, E=null, F=netoUnit, G=1, H="G1"
        const aoa = productos.map(p => [
          String(p.referencia || "").trim(), // A
          Number(p.cantidad) || 0,           // B
          null,                              // C
          null,                              // D
          null,                              // E
          Math.round(calcNetoUnit(p) * 100) / 100, // F
          1,                                 // G
          "G1",                              // H
        ]);
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "SimpleQuote");
        const nombreFichero = `SimpleQuote ${presupuesto.numerocompleto || presupuesto.np || ""} - ${presupuesto.titulo || "presupuesto"}`
          .replace(/[\\/:*?"<>|]/g, "").trim() + ".xlsx";
        descargarXLSX(wb, nombreFichero);
        setStatus(`Simple Quote generado con ${productos.length} producto${productos.length !== 1 ? "s" : ""}`, "success");
      } catch (err) {
        setStatus("Error al generar Simple Quote: " + (err.message || err), "error");
      }
      return;
    }
    if (action === "Exportar") {
      setStatus("Exportando datos del presupuesto a Excel...", "working");
      try {
        exportarDatosExcel(presupuesto, rows);
        setStatus("Datos exportados correctamente", "success");
      } catch (err) {
        setStatus("Error al exportar: " + (err.message || err), "error");
      }
      return;
    }
    if (action === "Importar") { setShowImportar(true); setStatus("Abriendo diálogo de importación", "info"); return; }
    if (action === "AplicarDescuentos") { setShowDescuentos(true); setStatus("Abriendo gestor de descuentos por grupo descuento", "info"); return; }
    if (action === "GestionarEstrategias") { setShowEstrategias(true); return; }
    if (action === "LeerPresupuestos") { setShowLeer(true); setStatus("Cargando lista de presupuestos desde la base de datos...", "working"); return; }
    if (action === "GestionarClientes") { setShowClientes(true); setStatus("Cargando lista de clientes desde la base de datos...", "working"); return; }
    if (action === "GestionarContactos") { setShowContactos(true); setStatus("Cargando lista de contactos desde la base de datos...", "working"); return; }
    if (action === "SeleccionarCeldas") { setShowSeleccionar(true); return; }
    if (action === "FijarPrecioTotal") { setShowFijarPrecio("total"); return; }
    if (action === "FijarPrecioCeldas") { setShowFijarPrecio("celdas"); return; }
    if (action === "ComprobarCelda") {
      // Muestra info de la celda seleccionada: nº de caracteres, color de fondo y color de tinta.
      if (!selectedCell) {
        setStatus("Selecciona una celda (rectángulo azul) para comprobarla", "error");
        return;
      }
      const row = rows.find(r => r.id === selectedCell.rowId);
      if (!row) { setStatus("No se encuentra la fila seleccionada", "error"); return; }
      const col = COLUMNS.find(c => c.key === selectedCell.colKey);
      // Valor mostrado de la celda (texto) para contar caracteres
      let valorTexto = row[selectedCell.colKey];
      if (valorTexto == null) valorTexto = "";
      const caracteres = String(valorTexto).length;
      // Estilo aplicado a la celda (depende de estructura activa y naturaleza)
      const estilo = estructuraActiva
        ? getEstiloFila(row.naturaleza, estilos)
        : { bg: null, color: "#1e293b" };
      const bg = estilo.bg || "#ffffff";
      const color = estilo.color || "#1e293b";
      const filaNum = rows.findIndex(r => r.id === selectedCell.rowId) + 1;
      setInfoCeldaDialog({
        caracteres,
        bg,
        color,
        colKey: col ? col.label : selectedCell.colKey,
        fila: filaNum,
        contenido: String(valorTexto),
      });
      return;
    }
    if (action === "BorrarFilasCero") {
      // Borra las filas que tengan 0 en la columna seleccionada (rectángulo azul).
      // La selección debe ser de una sola columna.
      let colKey = null;
      let colIdx = -1;
      let filasIdx = [];
      if (selectionRange) {
        const colIniS = Math.min(selectionRange.startColIdx, selectionRange.endColIdx);
        const colFinS = Math.max(selectionRange.startColIdx, selectionRange.endColIdx);
        if (colIniS !== colFinS) {
          setStatus("Selecciona celdas de una sola columna para borrar filas con 0", "error");
          return;
        }
        colIdx = colIniS;
        colKey = COLUMNS[colIdx]?.key;
        const rIni = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
        const rFin = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
        for (let r = rIni; r <= rFin; r++) if (rows[r]) filasIdx.push(r);
      } else if (selectedCell) {
        colKey = selectedCell.colKey;
        colIdx = COLUMNS.findIndex(c => c.key === colKey);
        const rIdx = rows.findIndex(r => r.id === selectedCell.rowId);
        if (rIdx >= 0) filasIdx.push(rIdx);
      } else {
        setStatus("Selecciona una celda o columna (rectángulo azul) para borrar filas con 0", "error");
        return;
      }
      if (!colKey) {
        setStatus("No se pudo determinar la columna seleccionada", "error");
        return;
      }
      // Valor numérico de la columna para cada fila (incluye columnas calculadas)
      const valorCol = (row) => {
        let v = row[colKey];
        if (colKey === "precionetounitario") v = calcNetoUnit(row);
        else if (colKey === "precionetoposicion") v = calcNetoPos(row);
        else if (colKey === "costeposicion") v = calcCostePos(row);
        else if (colKey === "margen") v = calcMargen(row);
        const n = Number(String(v).replace(",", "."));
        return isNaN(n) ? null : n;
      };
      // IDs de las filas a borrar: las seleccionadas cuyo valor en la columna sea 0
      const idsBorrar = new Set();
      filasIdx.forEach(r => {
        const row = rows[r];
        if (!row) return;
        if (valorCol(row) === 0) idsBorrar.add(row.id);
      });
      if (idsBorrar.size === 0) {
        setStatus(`No hay filas con 0 en la columna "${colKey}" dentro de la selección`, "info");
        return;
      }
      setRows(rs => rs.filter(r => !idsBorrar.has(r.id)));
      setSelectedRows(new Set());
      setSelectionRange(null);
      setSelectedCell(null);
      setStatus(`${idsBorrar.size} fila(s) con 0 en "${colKey}" borrada(s)`, "success");
      return;
    }
    if (action === "BorrarFilas") { setShowBorrarVacias(true); return; }
    if (action === "BorrarPresupuestoActual") { setConfirmBorrarPresup(true); return; }
    if (action === "ComprobarPresupuesto") {
      setStatus("Comparando con el presupuesto guardado en BD...", "working");
      buscarPresupuestoEnBD().then(bd => {
        if (!bd) {
          setComprobarDialog({ tipo: "no_existe" });
          setStatus("No existe este presupuesto en la BD", "info");
          return;
        }
        const diffs = compararPresupuestos(presupuesto, rows, bd);
        if (diffs.length === 0) {
          setComprobarDialog({ tipo: "iguales", comparado: bd });
          setStatus("El presupuesto actual es idéntico al guardado en BD", "success");
        } else {
          setComprobarDialog({ tipo: "diferentes", diffs, comparado: bd });
          setStatus(`Se han detectado ${diffs.length} diferencia${diffs.length > 1 ? "s" : ""} con la BD`, "info");
        }
      }).catch(e => {
        setStatus("Error comparando con BD: " + e.message, "error");
      });
      return;
    }
    if (action === "JuntarCeldas") {
      // Junta el texto de varias celdas en la primera, vaciando las demás.
      // Solo opera sobre columnas de texto (sin type number/calc/select).
      const COLS_TEXTO = ["posicion", "referencia", "nombre", "descripcion", "grupodescuento", "familia", "subfamilia", "idposicion", "imagen"];
      const cols = COLUMNS.map(c => c.key);
      const celdasSel = []; // [{rowIdx, colIdx, key, value}]
      if (selectionRange) {
        const colIni = Math.min(selectionRange.startColIdx, selectionRange.endColIdx);
        const colFin = Math.max(selectionRange.startColIdx, selectionRange.endColIdx);
        const rowIni = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
        const rowFin = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
        for (let r = rowIni; r <= rowFin; r++) {
          for (let c = colIni; c <= colFin; c++) {
            const row = rows[r];
            const key = cols[c];
            if (!row || !key) continue;
            celdasSel.push({ rowIdx: r, colIdx: c, key, value: row[key] });
          }
        }
      } else if (selectedCell) {
        const rowIdx = rows.findIndex(r => r.id === selectedCell.rowId);
        const colIdx = COLUMNS.findIndex(c => c.key === selectedCell.colKey);
        if (rowIdx >= 0 && colIdx >= 0) {
          celdasSel.push({ rowIdx, colIdx, key: selectedCell.colKey, value: rows[rowIdx][selectedCell.colKey] });
        }
      }
      if (celdasSel.length < 2) {
        setStatus("Selecciona al menos 2 celdas para juntarlas", "error");
        return;
      }
      // Filtrar solo columnas de texto
      const celdasTexto = celdasSel.filter(c => COLS_TEXTO.includes(c.key));
      if (celdasTexto.length < 2) {
        setStatus("Esta función solo opera sobre celdas de texto (referencia, producto, descripción...). Selecciona al menos 2 celdas de texto.", "error");
        return;
      }
      // Ordenar por columna izq→der y dentro de cada columna por fila arr→ab
      celdasTexto.sort((a, b) => a.colIdx - b.colIdx || a.rowIdx - b.rowIdx);
      // Concatenar valores no vacíos
      const textoConcat = celdasTexto
        .map(c => String(c.value ?? "").trim())
        .filter(v => v !== "")
        .join(" ");
      if (!textoConcat) {
        setStatus("Las celdas seleccionadas están vacías", "error");
        return;
      }
      // Primera celda recibe el texto, las demás se vacían
      const primera = celdasTexto[0];
      setRows(rs => {
        const next = [...rs];
        celdasTexto.forEach((c, i) => {
          const fila = { ...next[c.rowIdx] };
          fila[c.key] = (i === 0) ? textoConcat : "";
          next[c.rowIdx] = fila;
        });
        return next;
      });
      setStatus(`Texto unido en ${primera.key} (fila ${primera.rowIdx + 1}); ${celdasTexto.length - 1} celda${celdasTexto.length - 1 !== 1 ? "s" : ""} vaciada${celdasTexto.length - 1 !== 1 ? "s" : ""}`, "success");
      return;
    }
    if (action === "ConvertirEnComentario") {
      // 1) Determinar las celdas seleccionadas (priorizando rango, luego filas marcadas, luego celda activa)
      const celdasSel = []; // [{rowIdx, colIdx, key, value}]
      const cols = COLUMNS.map(c => c.key);
      if (selectionRange) {
        const colIni = Math.min(selectionRange.startColIdx, selectionRange.endColIdx);
        const colFin = Math.max(selectionRange.startColIdx, selectionRange.endColIdx);
        const rowIni = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
        const rowFin = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
        for (let r = rowIni; r <= rowFin; r++) {
          for (let c = colIni; c <= colFin; c++) {
            const row = rows[r];
            const key = cols[c];
            if (!row || !key) continue;
            celdasSel.push({ rowIdx: r, colIdx: c, key, value: row[key] });
          }
        }
      } else if (selectedCell) {
        const rowIdx = rows.findIndex(r => r.id === selectedCell.rowId);
        const colIdx = COLUMNS.findIndex(c => c.key === selectedCell.colKey);
        if (rowIdx >= 0 && colIdx >= 0) {
          celdasSel.push({ rowIdx, colIdx, key: selectedCell.colKey, value: rows[rowIdx][selectedCell.colKey] });
        }
      }
      if (celdasSel.length === 0) {
        setStatus("Selecciona al menos una celda para convertir en comentario", "error");
        return;
      }
      // 2) Ordenar por columna (izquierda a derecha) y dentro de la misma columna por fila (arriba a abajo)
      celdasSel.sort((a, b) => a.colIdx - b.colIdx || a.rowIdx - b.rowIdx);
      // 3) Construir el texto del comentario (solo valores no vacíos)
      const textoComentario = celdasSel
        .map(c => String(c.value ?? "").trim())
        .filter(v => v !== "")
        .join(" - ");
      if (!textoComentario) {
        setStatus("Las celdas seleccionadas están vacías", "error");
        return;
      }
      // 4) Determinar los rowIdx únicos que están seleccionados
      const filasAfectadasIdx = [...new Set(celdasSel.map(c => c.rowIdx))].sort((a, b) => a - b);
      const idsAEliminar = new Set(filasAfectadasIdx.map(idx => rows[idx].id));
      const idxInsercion = filasAfectadasIdx[0]; // antes de la primera fila afectada
      // 5) Crear la nueva fila comentario
      const filaComentario = {
        id: nextId.current++,
        representacion: "",
        naturaleza: "CM",
        posicion: "",
        cantidad: 0,
        referencia: "",
        nombre: textoComentario,
        pvp: 0, dtoaplicado: 0,
        precionetounitario: 0, precionetoposicion: 0,
        descripcion: "", familia: "", subfamilia: "",
        preciocosteunitario: 0, idposicion: "", imagen: "",
        precionetounitario2: 0, grupodescuento: "",
      };
      // 6) Insertar antes de filasAfectadasIdx[0] y eliminar las filas afectadas
      setRows(prev => {
        const result = [];
        prev.forEach((r, idx) => {
          if (idx === idxInsercion) {
            result.push(filaComentario);
          }
          if (!idsAEliminar.has(r.id)) {
            result.push(r);
          }
        });
        return result;
      });
      // Limpiar selecciones de celda/rango (las filas marcadas no se tocan)
      setSelectionRange(null);
      setSelectedCell(null);
      setStatus(`Comentario creado uniendo ${celdasSel.filter(c => String(c.value ?? "").trim() !== "").length} valor${celdasSel.length !== 1 ? "es" : ""} de ${filasAfectadasIdx.length} fila${filasAfectadasIdx.length !== 1 ? "s" : ""}`, "success");
      return;
    }
    if (action === "GuardarPresupuesto") {
      if (!presupuesto.idcliente) {
        setStatus("Selecciona un cliente antes de guardar el presupuesto", "error");
        return;
      }
      setStatus("Comprobando si ya existe en BD...", "working");
      buscarPresupuestoEnBD().then(async bd => {
        if (!bd) {
          // No existe, crear
          setGuardandoPresup(true);
          setStatus("Guardando presupuesto nuevo en BD...", "working");
          try {
            const r = await guardarPresupuestoEnBD(null);
            setPresupuesto(p => ({ ...p, id: String(r.id) }));
            setStatus(`Presupuesto guardado en BD (ID ${r.id}, ${r.lineas_insertadas} líneas)`, "success");
          } catch (e) {
            setStatus("Error al guardar: " + e.message, "error");
          } finally {
            setGuardandoPresup(false);
          }
          return;
        }
        // Existe: comparar
        const diffs = compararPresupuestos(presupuesto, rows, bd);
        if (diffs.length === 0) {
          setStatus("El presupuesto ya está guardado en BD sin cambios", "info");
          setComprobarDialog({ tipo: "iguales", comparado: bd });
          return;
        }
        // Hay diferencias, pedir confirmación
        setConfirmSobreescribir({ idBd: bd.cabecera.id, diffs, lineasBd: bd.detalle.length });
      }).catch(e => {
        setStatus("Error: " + e.message, "error");
      });
      return;
    }
    if (action === "GuardarElemento") {
      if (selectedRows.size === 0) {
        setStatus("Selecciona al menos una fila con un producto para guardar como elemento", "error");
        return;
      }
      setShowGuardarElem(true);
      return;
    }
    if (action === "LeerElemento") { setShowLeerElem(true); return; }
    if (action === "JuntarDuplicados") {
      // Solo opera sobre el rango seleccionado
      if (!selectionRange) {
        setStatus("Selecciona un rango de filas para juntar duplicados", "error");
        return;
      }
      const rowIni = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
      const rowFin = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
      // Agrupar por referencia (normalizada)
      const grupos = new Map(); // refNormalizada -> array de índices de fila
      for (let r = rowIni; r <= rowFin; r++) {
        const row = rows[r];
        if (!row) continue;
        const ref = String(row.referencia || "").trim().toUpperCase();
        if (!ref) continue; // saltar filas sin referencia
        if (!grupos.has(ref)) grupos.set(ref, []);
        grupos.get(ref).push(r);
      }
      // Sólo grupos con más de una aparición
      const duplicados = [...grupos.entries()].filter(([_, idxs]) => idxs.length > 1);
      if (duplicados.length === 0) {
        setStatus("No se han encontrado referencias duplicadas en el rango seleccionado", "info");
        return;
      }
      const totalFilasAfectadas = duplicados.reduce((sum, [_, idxs]) => sum + idxs.length, 0);
      const totalGrupos = duplicados.length;
      // Pedir al usuario qué hacer con los duplicados
      setJuntarDupConfirm({ duplicados, totalFilasAfectadas, totalGrupos, rowIni, rowFin });
      return;
    }
    if (action === "CrearSimpleQuote") {
      // Envía un email a SimpleQuote (Siemens) llamando a la API local /crear_sq.
      // Cuerpo: indicamos cliente (razón social) + tabla de productos PD/PE/E del presupuesto.
      (async () => {
        try {
          // Obtener email del contacto del presupuesto (a la atención de), si existe
          let emailContacto = "";
          if (presupuesto.alaatencion) {
            try {
              const rc = await fetch(`${API_URL}/contactos/${presupuesto.alaatencion}`);
              if (rc.ok) {
                const cont = await rc.json();
                emailContacto = String(cont.email || "").trim();
              }
            } catch {}
          }

          // Filtrar líneas de producto (PD, PE, E) con cantidad y referencia
          const productos = rows.filter(r => ["PD", "PE", "E"].includes(r.naturaleza))
            .filter(r => String(r.referencia || "").trim() !== "" && Number(r.cantidad) > 0);
          if (productos.length === 0) {
            setStatus("No hay líneas de producto con referencia y cantidad en el presupuesto", "error");
            return;
          }

          // Construir tabla en TEXTO PLANO con columnas alineadas por espacios.
          // Outlook a veces destroza el HTML; texto plano se ve fiable.
          const filas = productos.map(p => ({
            cant: String(Number(p.cantidad) || 0),
            ref: String(p.referencia || "").trim(),
          }));
          // Ancho de la columna cantidad = max(longitud del valor más largo, longitud encabezado)
          const anchoCant = Math.max(8, ...filas.map(f => f.cant.length));
          const sep = "    "; // 4 espacios de separación entre columnas
          const lineaCab = "Cantidad".padStart(anchoCant) + sep + "Referencia";
          const lineaSep = "-".repeat(anchoCant) + sep + "-".repeat(20);
          const lineasTabla = filas.map(f => f.cant.padStart(anchoCant) + sep + f.ref).join("\n");
          let cuerpo = `${lineaCab}\n${lineaSep}\n${lineasTabla}\n`;
          if (emailContacto) {
            cuerpo += `\nDe: ${emailContacto}\n`;
          }

          const payload = {
            destinatario: (configVarias.sqDestinatario || "").trim() || "simplequote-rfq.industry@siemens.com",
            asunto: String(presupuesto.titulo || "").trim() || "(presupuesto sin título)",
            cuerpo,
            cc: (configVarias.sqCC || "").trim(),
            adjuntos: [],
          };

          setStatus("Enviando SimpleQuote...", "working");
          const res = await fetch(`${API_LOCAL_URL}/crear_sq`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
          });
          if (!res.ok) {
            const txt = await res.text().catch(() => "");
            throw new Error(`HTTP ${res.status} ${txt.slice(0, 200)}`);
          }
          setStatus(`SimpleQuote enviada con ${productos.length} línea${productos.length !== 1 ? "s" : ""}`, "success");
        } catch (e) {
          setStatus(`Error al enviar SimpleQuote: ${e.message}. Comprueba que la API local en localhost:8000 esté en marcha.`, "error");
        }
      })();
      return;
    }
    if (action === "HacerDamexE") {
      // Validar requisitos
      const faltan = [];
      const titulo = String(presupuesto.titulo || "").trim();
      if (!titulo) faltan.push("Título del proyecto");
      if (!presupuesto.idcliente) faltan.push("Cliente");
      if (!presupuesto.idpaisclientefinal) faltan.push("País del cliente final");
      const clienteFinal = String(presupuesto.clientefinal || "").trim();
      if (!clienteFinal) faltan.push("Cliente final");
      if (faltan.length > 0) {
        setStatus(`No se puede ejecutar Damex E. Faltan: ${faltan.join(", ")}`, "error");
        return;
      }
      // Resolver código ISO del país desde paisesList
      const paisObj = paisesList.find(p => p.id === presupuesto.idpaisclientefinal);
      if (!paisObj) {
        setStatus(`No se encuentra el país con id ${presupuesto.idpaisclientefinal} en la lista de países`, "error");
        return;
      }
      // Nombre del cliente (nombrecomun o razonsocial — tenemos solo presupuesto.cliente como string)
      const nombreCliente = String(presupuesto.cliente || "").trim();
      // Total del presupuesto (suma de neto posición de líneas PD/PE/E)
      const totalObj = calcTotalPresupuesto(rows); // { neto, dto }
      const totalNeto = Number(totalObj?.neto) || 0;
      const valorPedido = (Math.round(totalNeto * 100) / 100).toString();

      const payload = {
        cliente: nombreCliente,
        pais_cliente: paisObj.codigo_iso2,
        nombre_cliente_final: clienteFinal,
        pedido: titulo,
        nombre_proyecto: titulo,
        valor_pedido: valorPedido,
      };
      setStatus(`Ejecutando Damex E...`, "working");
      (async () => {
        try {
          const r = await fetch(`${API_LOCAL_URL}/damex/ejecutar`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
          });
          if (!r.ok) throw new Error("HTTP " + r.status);
          const datos = await r.json();
          const resultado = String(datos.resultado || "").trim().toUpperCase();
          // Volcar al log los pasos ejecutados si los hay (independientemente del éxito)
          if (Array.isArray(datos.logs)) {
            datos.logs.forEach(line => console.log("[Damex E]", line));
          }
          if (resultado === "NN") {
            setStatus(`Damex E completado correctamente (${resultado})`, "success");
          } else {
            setStatus(`Damex E ha fallado: resultado "${resultado || "(vacío)"}" (esperado "NN")`, "error");
          }
        } catch (e) {
          setStatus(`Error ejecutando Damex E: ${e.message}`, "error");
        }
      })();
      return;
    }
    if (action === "LeerPreciosPMD") {
      // Determinar referencia inicial desde la fila de la celda activa
      let refInicial = "";
      let rowIdxObjetivo = -1;
      if (selectedCell) {
        rowIdxObjetivo = rows.findIndex(r => r.id === selectedCell.rowId);
        if (rowIdxObjetivo >= 0) refInicial = String(rows[rowIdxObjetivo].referencia || "").trim();
      }
      setLeerPmdDialog({ referencia: refInicial, rowIdx: rowIdxObjetivo });
      return;
    }
    if (action === "CalcularDescuento") {
      // Para cada celda seleccionada, calcular el descuento que aplicado al PVP
      // de su fila da el valor objetivo de la celda. Solo aplica a celdas numéricas.
      // Recoge celdas: prioridad rango → celda activa
      const cols = COLUMNS.map(c => c.key);
      const celdas = []; // [{rowIdx, key, valor}]
      if (selectionRange) {
        const colIni = Math.min(selectionRange.startColIdx, selectionRange.endColIdx);
        const colFin = Math.max(selectionRange.startColIdx, selectionRange.endColIdx);
        const rowIni = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
        const rowFin = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
        for (let r = rowIni; r <= rowFin; r++) {
          for (let c = colIni; c <= colFin; c++) {
            const row = rows[r];
            const key = cols[c];
            if (!row || !key) continue;
            celdas.push({ rowIdx: r, key, valor: row[key] });
          }
        }
      } else if (selectedCell) {
        const rowIdx = rows.findIndex(r => r.id === selectedCell.rowId);
        if (rowIdx >= 0) celdas.push({ rowIdx, key: selectedCell.colKey, valor: rows[rowIdx][selectedCell.colKey] });
      }
      if (celdas.length === 0) {
        setStatus("Selecciona al menos una celda con un valor numérico para calcular el descuento", "error");
        return;
      }
      // Por cada fila única afectada, tomar la PRIMERA celda numérica como objetivo
      const objetivoPorFila = new Map(); // rowIdx → valor numérico objetivo
      celdas.forEach(c => {
        if (objetivoPorFila.has(c.rowIdx)) return;
        const n = Number(c.valor);
        if (!isNaN(n)) objetivoPorFila.set(c.rowIdx, n);
      });
      if (objetivoPorFila.size === 0) {
        setStatus("Las celdas seleccionadas no contienen valores numéricos válidos", "error");
        return;
      }
      let aplicadas = 0, omitidas = 0;
      const detalles = [];
      setRows(rs => {
        const next = [...rs];
        objetivoPorFila.forEach((objetivo, rowIdx) => {
          const fila = next[rowIdx];
          if (!fila) { omitidas++; return; }
          const pvp = Number(fila.pvp) || 0;
          if (pvp <= 0) { omitidas++; detalles.push(`Fila ${rowIdx + 1}: PVP es 0 o negativo, no se puede calcular`); return; }
          // dto = (1 - objetivo/pvp) * 100
          const dto = (1 - objetivo / pvp) * 100;
          // Limitar a rango razonable: 0..100 (no permitimos descuentos negativos ni > 100%)
          // Se admiten descuentos negativos (recargo) sin tope superior fijo
          next[rowIdx] = { ...fila, dtoaplicado: Math.round(dto * 100) / 100 };
          aplicadas++;
        });
        return next;
      });
      if (aplicadas > 0 && omitidas === 0) {
        setStatus(`Descuento recalculado en ${aplicadas} fila${aplicadas !== 1 ? "s" : ""}`, "success");
      } else if (aplicadas > 0) {
        setStatus(`${aplicadas} fila${aplicadas !== 1 ? "s" : ""} ajustada${aplicadas !== 1 ? "s" : ""}, ${omitidas} omitida${omitidas !== 1 ? "s" : ""}${detalles.length ? ": " + detalles.slice(0, 2).join("; ") : ""}`, "warning");
      } else {
        setStatus(`No se pudo calcular: ${detalles[0] || "ninguna fila válida"}`, "error");
      }
      return;
    }
    if (action === "BuscarReferencia") {
      // 1) Determinar qué celdas procesar: la celda activa, las del rango, las de las filas marcadas
      const celdasObjetivo = []; // [{rowId, colKey}]
      if (selectionRange) {
        const cols = COLUMNS.map(c => c.key);
        const colIni = Math.min(selectionRange.startColIdx, selectionRange.endColIdx);
        const colFin = Math.max(selectionRange.startColIdx, selectionRange.endColIdx);
        const rowIni = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
        const rowFin = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
        for (let r = rowIni; r <= rowFin; r++) {
          const row = rows[r];
          if (!row) continue;
          for (let c = colIni; c <= colFin; c++) {
            const colKey = cols[c];
            if (!colKey) continue;
            celdasObjetivo.push({ rowId: row.id, colKey });
          }
        }
      } else if (selectedCell) {
        celdasObjetivo.push({ rowId: selectedCell.rowId, colKey: selectedCell.colKey });
      } else if (selectedRows.size > 0) {
        // Filas marcadas → toda la fila (todas las columnas de texto)
        rows.forEach(r => {
          if (selectedRows.has(r.id)) {
            ["referencia", "nombre", "descripcion"].forEach(k => celdasObjetivo.push({ rowId: r.id, colKey: k }));
          }
        });
      }
      if (celdasObjetivo.length === 0) {
        setStatus("Selecciona al menos una celda o fila para buscar referencias SIEMENS", "error");
        return;
      }
      // 2) Para cada celda extraer las refs SIEMENS
      const celdasTrabajadas = [];
      celdasObjetivo.forEach(({ rowId, colKey }) => {
        const row = rows.find(r => r.id === rowId);
        if (!row) return;
        const texto = String(row[colKey] ?? "");
        const refs = parseSiemensRefs(texto);
        if (refs.length > 0) {
          celdasTrabajadas.push({ rowId, colKey, texto, refs });
        }
      });
      if (celdasTrabajadas.length === 0) {
        setStatus("No se encontraron referencias SIEMENS en las celdas seleccionadas", "info");
        return;
      }
      const totalRefs = celdasTrabajadas.reduce((sum, c) => sum + c.refs.length, 0);
      setStatus(`Encontradas ${totalRefs} referencia${totalRefs !== 1 ? "s" : ""} SIEMENS en ${celdasTrabajadas.length} celda${celdasTrabajadas.length !== 1 ? "s" : ""}`, "info");
      setSiemensDialog({ celdasTrabajadas });
      return;
    }
    if (action === "GuardarProducto") {
      // 1) Determinar filas a procesar
      const idsProcesar = new Set();
      if (selectionRange) {
        const rowIni = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
        const rowFin = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
        for (let r = rowIni; r <= rowFin; r++) {
          const row = rows[r];
          if (row) idsProcesar.add(row.id);
        }
      }
      if (selectedRows.size > 0) {
        selectedRows.forEach(id => idsProcesar.add(id));
      }
      if (idsProcesar.size === 0 && selectedCell) {
        idsProcesar.add(selectedCell.rowId);
      }
      if (idsProcesar.size === 0) {
        setStatus("Selecciona al menos una fila para guardar como producto", "error");
        return;
      }
      const filasObjetivo = rows.filter(r => idsProcesar.has(r.id));

      (async () => {
        setStatus(`Guardando ${filasObjetivo.length} producto${filasObjetivo.length !== 1 ? "s" : ""}...`, "working");
        // Cargar tabla gruposdescuento una sola vez para resolver códigos → ids
        let gruposCache = [];
        try {
          const r = await fetch(`${API_URL}/gruposdescuento/`);
          if (r.ok) gruposCache = await r.json();
        } catch (e) {
          setStatus("Error cargando grupos descuento: " + e.message, "error");
          return;
        }
        const resolverGrupoId = (codigo) => {
          if (!codigo) return null;
          const cod = String(codigo).trim().toUpperCase();
          const g = gruposCache.find(x => String(x.grupodescuentospain || "").trim().toUpperCase() === cod);
          return g ? g.id : null;
        };

        let creados = 0, saltadosExistentes = 0, errores = 0, saltadosFaltan = 0, saltadosGrupo = 0;
        const detalles = []; // mensajes detalle para el log
        const existentes = []; // productos que ya existen → se gestionarán en el diálogo de actualización

        for (const fila of filasObjetivo) {
          const ref = String(fila.referencia || "").trim();
          const pvp = Number(fila.pvp) || 0;
          const descripcion = String(fila.descripcion || "").trim();
          const grupoCod = String(fila.grupodescuento || "").trim();
          const nombreOriginal = String(fila.nombre || "").trim();

          // Validar campos obligatorios
          if (!ref || pvp <= 0 || !descripcion || !grupoCod) {
            saltadosFaltan++;
            detalles.push(`✗ ${ref || "(sin ref)"}: faltan campos obligatorios (ref, pvp, descripción, grupodescuento)`);
            continue;
          }
          // Resolver el id del grupo
          const idGrupo = resolverGrupoId(grupoCod);
          if (idGrupo === null) {
            saltadosGrupo++;
            detalles.push(`✗ ${ref}: grupo descuento "${grupoCod}" no existe en BD`);
            continue;
          }
          // Calcular nombre: si nombre está vacío → primeros 50 chars de descripción
          const nombre = nombreOriginal || descripcion.substring(0, 50);

          // Construir payload (solo campos no vacíos)
          const payload = { referencia: ref, nombre, descripcion, pvp, grupodescuento: idGrupo };
          const preciocoste = Number(fila.preciocosteunitario);
          if (!isNaN(preciocoste) && preciocoste > 0) payload.preciocoste = preciocoste;

          // Comprobar si ya existe por referencia
          try {
            const rExist = await fetch(`${API_URL}/productos/referencia/${encodeURIComponent(ref)}`);
            if (rExist.status === 404) {
              // No existe → seguimos al INSERT
            } else if (rExist.ok) {
              // 2xx: verificar que de verdad devuelve un producto con la misma referencia
              const data = await rExist.json().catch(() => null);
              const refDevuelta = data && data.referencia ? String(data.referencia).trim().toUpperCase() : "";
              if (data && data.id && refDevuelta === ref.toUpperCase()) {
                // Existe: lo recopilamos para el diálogo de actualización (no se omite sin más)
                existentes.push({
                  fila,
                  productoBD: data,
                  // valores nuevos candidatos a actualizar (de la fila del presupuesto)
                  nuevos: {
                    pvp,
                    descripcion,
                    nombre,
                    preciocoste: (!isNaN(Number(fila.preciocosteunitario)) && Number(fila.preciocosteunitario) > 0) ? Number(fila.preciocosteunitario) : null,
                    grupodescuento: idGrupo,
                    grupoCod,
                  },
                });
                saltadosExistentes++;
                continue;
              }
              // 2xx pero no es un producto válido → seguimos al INSERT (lo trataremos como inexistente)
              detalles.push(`! ${ref}: respuesta 2xx inesperada (sin id/ref), se intenta crear igualmente`);
            } else {
              throw new Error("HTTP " + rExist.status);
            }
          } catch (e) {
            errores++;
            detalles.push(`✗ ${ref}: error comprobando existencia (${e.message})`);
            continue;
          }

          // Crear el producto
          try {
            const rCreate = await fetch(`${API_URL}/productos/`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(payload),
            });
            if (!rCreate.ok) {
              const err = await rCreate.json().catch(() => null);
              throw new Error(err?.detail || "HTTP " + rCreate.status);
            }
            creados++;
            detalles.push(`✓ ${ref}: creado (grupo ${grupoCod})`);
          } catch (e) {
            errores++;
            detalles.push(`✗ ${ref}: error creando (${e.message})`);
          }
        }

        // Resumen de la fase de creación
        const partes = [];
        if (creados > 0) partes.push(`${creados} creado${creados !== 1 ? "s" : ""}`);
        if (saltadosFaltan > 0) partes.push(`${saltadosFaltan} sin campos obligatorios`);
        if (saltadosGrupo > 0) partes.push(`${saltadosGrupo} con grupo inválido`);
        if (errores > 0) partes.push(`${errores} con error`);
        console.log("[GuardarProducto] Detalles:");
        detalles.forEach(d => console.log("  " + d));

        // Si hay productos existentes, abrir el diálogo de actualización
        if (existentes.length > 0) {
          const resumenCrear = partes.length > 0 ? partes.join(", ") + ". " : "";
          setStatus(`${resumenCrear}${existentes.length} producto(s) ya existen: configura qué actualizar`, "info");
          setActualizarProductos({ existentes, gruposCache });
        } else {
          const resumen = partes.join(", ");
          const tipo = errores > 0 ? "error" : (creados > 0 ? "success" : "info");
          setStatus(`Guardar Producto: ${resumen || "nada que procesar"}`, tipo);
        }
      })();
      return;
    }
    if (action === "LeerProducto") { setShowLeerProducto(true); return; }
    if (action === "Asistente") { setShowAsistente(true); return; }
    if (action === "BuscarDatosProductos") {
      // Determinar qué filas procesar:
      // 1. Si hay rango seleccionado (Shift+click/arrastrar) → todas las filas del rango
      // 2. Si hay filas con checkbox marcado → esas filas
      // 3. Si solo hay una celda activa → la fila de esa celda
      const idsProcesar = new Set();
      if (selectionRange) {
        const minRow = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
        const maxRow = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
        for (let i = minRow; i <= maxRow; i++) {
          if (rows[i]) idsProcesar.add(rows[i].id);
        }
      } else if (selectedRows.size > 0) {
        selectedRows.forEach(id => idsProcesar.add(id));
      } else if (selectedCell) {
        idsProcesar.add(selectedCell.rowId);
      } else {
        setStatus("Selecciona una celda o un rango de filas", "error");
        return;
      }

      // Filtrar filas que tengan referencia
      const filasConRef = rows.filter(r =>
        idsProcesar.has(r.id) && r.referencia && String(r.referencia).trim() !== ""
      );

      if (filasConRef.length === 0) {
        setStatus("Ninguna de las filas seleccionadas tiene referencia", "error");
        return;
      }

      setStatus(`Buscando ${filasConRef.length} referencia${filasConRef.length !== 1 ? "s" : ""} en BD...`, "working");

      (async () => {
        let ok = 0, noEncontradas = 0, errores = 0;
        const actualizaciones = new Map(); // id → objeto con campos a actualizar

        for (const row of filasConRef) {
          const ref = String(row.referencia).trim();
          try {
            const res = await fetch(`${API_URL}/productos/referencia/${encodeURIComponent(ref)}`);
            if (res.status === 404) {
              noEncontradas++;
              continue;
            }
            if (!res.ok) { errores++; continue; }
            const prod = await res.json();
            const descuento = Number(prod.dtoreferenciagrupo) || 0;
            // Naturaleza: forzar PD
            // Cantidad: si está vacía, no es número o es ≤ 0 → 1; si es número > 0 se respeta
            const cantNum = Number(row.cantidad);
            const cantidadFinal = (isNaN(cantNum) || cantNum <= 0) ? 1 : cantNum;
            actualizaciones.set(row.id, {
              naturaleza: "PD",
              cantidad: cantidadFinal,
              referencia: prod.referencia || "",
              nombre: prod.nombre || "",
              descripcion: prod.descripcion || "",
              pvp: Number(prod.pvp) || 0,
              dtoaplicado: descuento,
              familia: prod.familia || "",
              subfamilia: prod.subfamilia || "",
              descripcionsubfamilia: prod.descripcionsubfamilia || "",
              grupodescuento: prod.grupodescuento || "",
              descripciongrupodescuento: prod.descripciongrupodescuento || "",
              preciocosteunitario: Number(prod.preciocoste) || row.preciocosteunitario || 0,
              idposicion: prod.id || "",
              fechapvp: prod.fechapvp || null,
              fechapreciocoste: prod.fechapreciocoste || null,
              pvpVencido: tieneMasDeUnAno(prod.fechapvp),
            });
            ok++;
          } catch (e) {
            errores++;
          }
        }

        // Aplicar todas las actualizaciones en una sola llamada a setRows
        if (actualizaciones.size > 0) {
          setRows(rs => rs.map(r => actualizaciones.has(r.id) ? { ...r, ...actualizaciones.get(r.id) } : r));
        }

        // Mensaje resumen
        const partes = [];
        if (ok > 0) partes.push(`${ok} encontrado${ok !== 1 ? "s" : ""}`);
        if (noEncontradas > 0) partes.push(`${noEncontradas} no encontrado${noEncontradas !== 1 ? "s" : ""}`);
        if (errores > 0) partes.push(`${errores} error${errores !== 1 ? "es" : ""}`);
        const tipo = errores > 0 ? "error" : noEncontradas > 0 ? "info" : "success";
        setStatus(`Buscar datos terminado: ${partes.join(", ")}`, tipo);
      })();
      return;
    }
    setModal({ action, label });
    setStatus("Función \"" + label + "\" sin implementar", "info");
  }, [presupuesto, rows, apartados, estructuraActiva, setStatus, selectedRows]);

  const startEdit = (rowId, colKey, val) => { escapeEditRef.current = false; setEditingCell({ rowId, colKey }); setEditValue(String(val ?? "")); };
  const commitEdit = (override) => {
    if (!editingCell) return;
    // Si se acaba de pulsar Escape, descartar sin guardar
    if (escapeEditRef.current) {
      escapeEditRef.current = false;
      setEditingCell(null);
      return;
    }
    // Capturar la celda en edición ahora, por si el estado cambia durante el commit
    const celda = editingCell;
    const rawValue = (override !== undefined && override !== null) ? override : editValue;
    try {
      setRows(r => r.map(row => {
        if (row.id !== celda.rowId) return row;
        let value = rawValue;
        // Columnas numéricas: aceptan coma o punto como separador decimal
        const COLS_NUMERICAS = ["pvp", "dtoaplicado", "preciocosteunitario", "precionetounitario2"];
        if (celda.colKey === "cantidad") {
          const n = parseInt(String(rawValue).replace(",", "."), 10);
          value = isNaN(n) ? 0 : n;
        } else if (celda.colKey === "referencia") {
          // Mantener como string siempre, no convertir a Number aunque sea numérico
          value = String(rawValue ?? "");
        } else if (COLS_NUMERICAS.includes(celda.colKey)) {
          // Normalizar coma decimal → punto antes de convertir
          const limpio = String(rawValue).replace(/\s/g, "").replace(",", ".");
          const n = parseFloat(limpio);
          value = isNaN(n) ? 0 : n;
        } else if (!isNaN(rawValue) && rawValue !== "") {
          value = Number(rawValue);
        }
        return { ...row, [celda.colKey]: value };
      }));
    } catch (err) {
      console.error("[commitEdit] error al guardar la celda:", err);
    }
    setEditingCell(null);
  };

  // Comprueba si una celda está dentro del rango de selección
  const isCellInRange = (rowIdx, colIdx) => {
    if (!selectionRange) return false;
    const minRow = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
    const maxRow = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
    const minCol = Math.min(selectionRange.startColIdx, selectionRange.endColIdx);
    const maxCol = Math.max(selectionRange.startColIdx, selectionRange.endColIdx);
    return rowIdx >= minRow && rowIdx <= maxRow && colIdx >= minCol && colIdx <= maxCol;
  };

  const totalNeto  = rows.reduce((s, r) => s + calcNetoPos(r), 0);

  // Resumen de la selección (rectángulo azul): nº de celdas y suma de las numéricas,
  // al estilo de Excel. Funciona tanto para un rango como para una celda única.
  const resumenSeleccion = (() => {
    const valorCelda = (row, col) => {
      if (!row || !col) return null;
      let val = row[col.key];
      if (col.key === "precionetounitario") val = calcNetoUnit(row);
      else if (col.key === "precionetoposicion") val = calcNetoPos(row);
      else if (col.key === "costeposicion") val = calcCostePos(row);
      else if (col.key === "margen") val = calcMargen(row);
      const n = Number(val);
      return isNaN(n) ? null : n;
    };
    let celdas = 0;     // total de celdas en la selección
    let numericas = 0;  // celdas con valor numérico
    let suma = 0;
    if (selectionRange) {
      const minRow = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
      const maxRow = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
      const minCol = Math.min(selectionRange.startColIdx, selectionRange.endColIdx);
      const maxCol = Math.max(selectionRange.startColIdx, selectionRange.endColIdx);
      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          celdas++;
          const n = valorCelda(rows[r], COLUMNS[c]);
          if (n !== null) { numericas++; suma += n; }
        }
      }
    } else if (selectedCell) {
      celdas = 1;
      const row = rows.find(rr => rr.id === selectedCell.rowId);
      const col = COLUMNS.find(cc => cc.key === selectedCell.colKey);
      const n = valorCelda(row, col);
      if (n !== null) { numericas++; suma += n; }
    }
    return { celdas, numericas, suma };
  })();
  const totalCoste = rows.reduce((s, r) => s + calcCostePos(r), 0);
  const margenTotal = totalNeto ? ((totalNeto - totalCoste) / totalNeto * 100) : 0;

  const resumenData = (() => {
    const mapa = {};
    rows.forEach(row => {
      if (!row.familia && !row.subfamilia) return;
      const key = (row.familia || "Sin familia") + "||" + (row.subfamilia || "Sin subfamilia");
      if (!mapa[key]) mapa[key] = { familia: row.familia || "Sin familia", subfamilia: row.subfamilia || "Sin subfamilia", articulos: 0, importe: 0 };
      mapa[key].articulos += (row.cantidad || 0);
      mapa[key].importe += calcNetoPos(row);
    });
    return Object.values(mapa).sort((a, b) => b.importe - a.importe);
  })();
  const totalResumen = resumenData.reduce((s, r) => s + r.importe, 0);

  const scrollTo = (id) => {
    setAyudaSeccion(id);
    const el = document.getElementById("ayuda-" + id);
    if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  // ── VISTA OPCIONES ──
  if (vista === "opciones") return (
    <OpcionesScreen
      estilos={estilos}
      setEstilos={(nuevos) => { setEstilos(nuevos); guardarEstilos(nuevos); }}
      configVarias={configVarias}
      setConfigVarias={setConfigVarias}
      setStatus={setStatus}
      statusMessage={statusMessage}
      onVolver={() => setVista("grid")}
    />
  );

  // ── VISTA AYUDA ──
  if (vista === "ayuda") return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", fontSize: 13, color: "#1e293b", height: "100vh", display: "flex", flexDirection: "column" }}>
      <div style={{ background: "#f5f5f5", color: "#171717", padding: "8px 16px", display: "flex", alignItems: "center", gap: 12, flexShrink: 0, borderBottom: "1px solid #e5e5e5" }}>
        <button onClick={() => setVista("grid")} style={{ background: "#fff", border: "1px solid #d4d4d4", color: "#171717", borderRadius: 6, padding: "4px 12px", cursor: "pointer", fontSize: 12 }}><BtnContent icon={ArrowLeft}>← Volver</BtnContent></button>
        <span style={{ fontWeight: 700, fontSize: 15, display: "inline-flex", alignItems: "center", gap: 8 }}><Icon as={HelpCircle} size={18} color="#171717" /> Ayuda — Manual de uso</span>
        <span style={{ color: "#737373", fontSize: 12 }}>v2.00.0 (9 Junio 2026)</span>
      </div>
      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
        {/* ÁRBOL IZQUIERDA */}
        <div style={{ width: 230, flexShrink: 0, background: "#fff", borderRight: "1px solid #e2e8f0", overflowY: "auto", padding: "8px 0" }}>
          <div style={{ padding: "10px 12px 6px", fontSize: 11, fontWeight: 700, color: "#94a3b8", letterSpacing: 1, textTransform: "uppercase" }}>Índice</div>
          {AYUDA_TREE.map(node => (
            <TreeNode key={node.id} node={node} activeId={ayudaSeccion} onSelect={scrollTo} />
          ))}
        </div>
        {/* CONTENIDO DERECHA */}
        <div style={{ flex: 1, overflowY: "auto", padding: "28px 40px", background: "#f8fafc" }}
          onScroll={e => {
            for (const node of AYUDA_TREE) {
              const ids = node.children && node.children.length > 0 ? node.children.map(c => c.id) : [node.id];
              for (const id of [node.id, ...ids]) {
                const el = document.getElementById("ayuda-" + id);
                if (el) { const rect = el.getBoundingClientRect(); if (rect.top >= 0 && rect.top < 250) { setAyudaSeccion(id); return; } }
              }
            }
          }}>
          <div style={{ maxWidth: 740, display: "flex", flexDirection: "column", gap: 36 }}>

            <section id="ayuda-intro">
              <h2 style={{ fontSize: 18, fontWeight: 700, color: "#1e3a5f", marginBottom: 10, borderBottom: "2px solid #2563eb", paddingBottom: 8 }}>¿Qué es esta aplicación?</h2>
              <p style={{ color: "#475569", lineHeight: 1.8, margin: 0 }}>
                Aplicación para crear y gestionar <strong>presupuestos comerciales</strong>. Cada presupuesto tiene una cabecera con datos generales (ID, número, revisión, título, cliente) y una grid de líneas con productos, elementos, apartados y subtotales. Los importes neto, coste y margen se calculan automáticamente. La numeración de apartados se genera sola al activar la estructura, los subtotales (S1-S4) calculan la suma y descuento medio de su apartado, y los productos se autocompletan al teclear referencias gracias a la conexión con la base de datos. La aplicación permite importar/exportar Excel, leer presupuestos y elementos guardados, gestionar clientes y contactos, y aplicar descuentos por subfamilia o fijar precios totales con reparto proporcional al neto. <strong>El acceso requiere login</strong> (tras 5 intentos fallidos se bloquea 10 minutos). Los valores numéricos usan formato español: coma decimal y punto de miles, y los importes en euros con el símbolo €.
              </p>
            </section>

            <section id="ayuda-novedades">
              <h2 style={{ fontSize: 16, fontWeight: 700, color: "#1e3a5f", marginBottom: 12, borderBottom: "2px solid #e2e8f0", paddingBottom: 8 }}>Novedades recientes</h2>
              <p style={{ color: "#475569", lineHeight: 1.7, margin: "0 0 14px 0" }}>
                Resumen de los cambios más recientes incorporados a la aplicación, agrupados por área.
              </p>

              <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>

                <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, padding: "14px 18px" }}>
                  <h3 style={{ fontSize: 14, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
                    <Icon as={Palette} size={16} color="#16a34a" /> Aplicar estructura
                  </h3>
                  <ul style={{ color: "#475569", lineHeight: 1.7, margin: 0, paddingLeft: 20, fontSize: 13 }}>
                    <li><strong>Numeración jerárquica automática</strong> de apartados (T1 → 1, T2 → 1.1, T3 → 1.1.1, etc.), sin ceros (mínimo 1 en cada nivel).</li>
                    <li>Anchos de columna <strong>auto-ajustados al contenido</strong> (representación, naturaleza, cantidad, PVP, dto, neto unit., neto pos., apartado).</li>
                    <li><strong>Multilínea automático</strong> en Referencia, Producto y Descripción si no caben en una línea. Tooltip con el contenido completo.</li>
                    <li>Importes formateados como <strong>moneda con €</strong> y descuentos con <strong>2 decimales + %</strong>.</li>
                    <li><strong>Subtotales S1-S4 y total TT</strong> calculados automáticamente. Estilos selectivos: títulos centrados, subtotales/total justificados a la izquierda.</li>
                    <li>El campo nombre de <strong>S1-S4 se rellena con "TOTAL &lt;texto del Tn anterior&gt;"</strong>; el de <strong>TT con "TOTAL &lt;título del presupuesto&gt;"</strong>. Persiste al desactivar estructura.</li>
                    <li><strong>Importes negativos en rojo</strong> (PVP, Dto, netos, coste, neto2) cuando estructura activa.</li>
                    <li>Al <strong>desactivar</strong> estructura los formatos quedan tal cual (anchos, multilínea, etiquetas). Al volver a activar se recalcula todo.</li>
                  </ul>
                </div>

                <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, padding: "14px 18px" }}>
                  <h3 style={{ fontSize: 14, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
                    <Icon as={Grid3x3} size={16} color="#7c3aed" /> Grid de líneas
                  </h3>
                  <ul style={{ color: "#475569", lineHeight: 1.7, margin: 0, paddingLeft: 20, fontSize: 13 }}>
                    <li><strong>Redimensionado a mano</strong> (estilo Excel) arrastrando el borde derecho de la cabecera o el borde inferior del número de fila.</li>
                    <li>Los <strong>anchos y altos persisten</strong> al desactivar estructura. Activarla recalcula los anchos auto (sobreescribiendo los manuales hasta el siguiente cambio).</li>
                    <li>Edición de celdas de texto largo (Referencia / Producto / Descripción) con <strong>textarea flotante</strong> que se autoajusta al contenido (filas variables).</li>
                    <li>Selección, copiar/pegar y borrar <strong>funcionan dentro del editor</strong> (no hay interferencia con los handlers globales del grid).</li>
                    <li>Sin <strong>límite de 50 caracteres</strong> en la columna Producto (admite títulos y comentarios largos).</li>
                    <li>Naturaleza y Representación se pueden <strong>editar a mano</strong> (con filtrado y validación contra la lista permitida).</li>
                    <li>Presupuestos nuevos arrancan con <strong>50 filas vacías</strong> en vez de 10.</li>
                    <li>Botones <strong>Copiar filas / Pegar filas</strong> en la barra del grid: copia las filas marcadas, pega a partir de la celda activa (rellena vacías o inserta si están ocupadas).</li>
                    <li>Pegar números con formato europeo (<code>3.948,45 €</code>) en celdas numéricas se interpreta correctamente.</li>
                  </ul>
                </div>

                <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, padding: "14px 18px" }}>
                  <h3 style={{ fontSize: 14, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
                    <Icon as={MessageSquare} size={16} color="#dc2626" /> Celdas
                  </h3>
                  <ul style={{ color: "#475569", lineHeight: 1.7, margin: 0, paddingLeft: 20, fontSize: 13 }}>
                    <li><strong>Juntar celdas en una</strong>: concatena el texto de varias celdas de texto en la primera, vaciando las demás.</li>
                    <li><strong>Convertir selección en comentario</strong>: une las celdas seleccionadas con " - " en una nueva fila CM y elimina las originales.</li>
                  </ul>
                </div>

                <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, padding: "14px 18px" }}>
                  <h3 style={{ fontSize: 14, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
                    <Icon as={Percent} size={16} color="#7c3aed" /> Descuentos
                  </h3>
                  <ul style={{ color: "#475569", lineHeight: 1.7, margin: 0, paddingLeft: 20, fontSize: 13 }}>
                    <li><strong>Calcular descuento</strong>: para cada celda numérica seleccionada, ajusta el dto. de su fila para que el neto unitario coincida con el valor objetivo. Admite descuentos negativos (recargo).</li>
                    <li><strong>Aplicar descuentos por Grupo Descuento</strong>: agrupa por <code>grupo + dto actual</code> (filas separadas por descuento). Muestra <strong>DGL1</strong> y <strong>DGL2</strong> del grupo; resalta el dto. en <strong>amarillo si supera DGL1</strong> y en <strong>rojo si supera DGL2</strong>.</li>
                  </ul>
                </div>

                <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, padding: "14px 18px" }}>
                  <h3 style={{ fontSize: 14, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
                    <Icon as={Package} size={16} color="#0369a1" /> Productos
                  </h3>
                  <ul style={{ color: "#475569", lineHeight: 1.7, margin: 0, paddingLeft: 20, fontSize: 13 }}>
                    <li><strong>Buscar datos por Referencia</strong> rellena además <code>naturaleza = PD</code> y fuerza <code>cantidad ≥ 1</code>.</li>
                    <li><strong>Leer Producto</strong> inserta en la fila de la <strong>celda activa</strong>: si está vacía la rellena; si tiene contenido inserta una fila nueva encima.</li>
                    <li><strong>Juntar productos duplicados</strong>: suma cantidades en la primera fila del grupo (dejar a 0 o borrar duplicadas).</li>
                  </ul>
                </div>

                <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, padding: "14px 18px" }}>
                  <h3 style={{ fontSize: 14, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
                    <Icon as={Wrench} size={16} color="#7c3aed" /> Mantenimiento BD
                  </h3>
                  <ul style={{ color: "#475569", lineHeight: 1.7, margin: 0, paddingLeft: 20, fontSize: 13 }}>
                    <li><strong>Mantenimiento BD Familia / Subfamilia</strong>: editar y borrar (con validación de uso antes de borrar).</li>
                    <li><strong>Mantenimiento Grupos Descuento</strong>: importa Excel con drag&drop. Reconoce las columnas <strong>DGL1</strong> y <strong>DGL2</strong>. Recuadro de mapeo con tags verde/rojo y errores. Selector SÍ/NO para sobreescribir existentes.</li>
                    <li><strong>Actualizar Tarifas</strong>: matching de cabeceras por prefijo del nombre real del campo BD (sin aliases). Ej: <code>pla → plazoentrega</code>, <code>precio → preciocoste</code>, <code>fecha → ambigua</code>. Procesa todas las columnas reconocidas (pvp_l1, plazoentrega, etc.).</li>
                  </ul>
                </div>

                <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, padding: "14px 18px" }}>
                  <h3 style={{ fontSize: 14, fontWeight: 700, color: "#1e3a5f", marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
                    <Icon as={User} size={16} color="#2563eb" /> Sesión, ventana y app
                  </h3>
                  <ul style={{ color: "#475569", lineHeight: 1.7, margin: 0, paddingLeft: 20, fontSize: 13 }}>
                    <li>Pantalla de bienvenida con login real contra el backend (mensaje OK/KO).</li>
                    <li><strong>Pastilla de usuario</strong> en la barra superior con menú Iniciar / Cerrar sesión.</li>
                    <li>El diálogo de Iniciar sesión también permite cambiar de usuario sin salir.</li>
                    <li>Eliminado el <strong>scroll global</strong> de la página: la ventana se dimensiona automáticamente sin barra extra.</li>
                  </ul>
                </div>

              </div>
            </section>

            <section id="ayuda-cabecera">
              <h2 style={{ fontSize: 16, fontWeight: 700, color: "#1e3a5f", marginBottom: 12, borderBottom: "2px solid #e2e8f0", paddingBottom: 8 }}>Cabecera del presupuesto</h2>
              {[
                { campo: "ID",             desc: "Identificador único generado por la BD. No modificable." },
                { campo: "Nº Completo",    desc: "Solo lectura. Se construye automáticamente como CÓDIGO-NÚMERO-AÑO (ej. CDM-123-2026), donde el código viene del usuario activo. Se recalcula al cambiar el Nº o el Año." },
                { campo: "Nº",             desc: "Número de presupuesto. Al crear uno nuevo se asigna automáticamente como (max + 1) de los presupuestos del año actual." },
                { campo: "Rev.",           desc: "Número de revisión. Empieza en 0." },
                { campo: "Año",            desc: "Año fiscal del presupuesto. Por defecto el actual (año fiscal va del 1 octubre del año N-1 al 30 septiembre del año N). Se puede editar a mano para registrar presupuestos de años anteriores." },
                { campo: "Título",         desc: "Descripción breve del proyecto." },
                { campo: "Cliente",        desc: "Nombre del cliente destinatario." },
              ].map((r, i) => (
                <div key={i} style={{ display: "flex", gap: 16, padding: "8px", borderBottom: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#f8fafc" }}>
                  <span style={{ minWidth: 140, fontWeight: 600, color: "#1e3a5f", fontSize: 12 }}>{r.campo}</span>
                  <span style={{ color: "#475569", lineHeight: 1.6, fontSize: 12 }}>{r.desc}</span>
                </div>
              ))}
            </section>

            <section id="ayuda-columnas">
              <h2 style={{ fontSize: 16, fontWeight: 700, color: "#1e3a5f", marginBottom: 12, borderBottom: "2px solid #e2e8f0", paddingBottom: 8 }}>Columnas de la grid</h2>
              <p style={{ color: "#475569", marginBottom: 10, fontSize: 12 }}>Doble click para editar. Las columnas en azul son calculadas automáticamente.</p>
              {[
                { col: "Repres.",             desc: "Seleccionable: TP (Total Posición), Opcional (no suma al total), Conf (Confirmar). Puede dejarse vacío." },
                { col: "Nat.",                desc: "Naturaleza de la línea: PD/PE (Productos), E (Elemento), T1-T4 (Títulos), S1-S4 (Subtotales), CM (Comentario), TT (Total), VERDE/GRIS (líneas resaltadas). Obligatorio." },
                { col: "Apartado",            desc: "Numeración jerárquica automática al activar Estructura. T1→1, sus PD→1.1/1.2, T2→1.3..." },
                { col: "Cant.",               desc: "Número de unidades." },
                { col: "Referencia",          desc: "Código de referencia del fabricante. Editable con autocompletado conectado a BD." },
                { col: "Producto",            desc: "Nombre del producto, elemento, título o texto del comentario (en filas CM)." },
                { col: "PVP Unitario",        desc: "Precio de tarifa oficial unitario. Con Estructura activa: formato 1.234,56 €." },
                { col: "Dto. %",              desc: "Descuento aplicado. En S1-S4 muestra el descuento medio ponderado del apartado." },
                { col: "Neto Unitario",       desc: "Calculado: PVP × (1 - Dto% / 100). Con Estructura activa: formato 1.234,56 €." },
                { col: "Neto Posición",       desc: "Calculado: Neto Unitario × Cantidad. En S1-S4 muestra la suma del apartado. Con Estructura activa: formato 1.234,56 €." },
                { col: "Descripción",         desc: "Texto descriptivo adicional. Wrap automático (varias líneas)." },
                { col: "Grupo Descuento",     desc: "Código del grupo descuento del producto (ej. GD123). Al pasar el ratón muestra la descripción completa del grupo." },
                { col: "Familia",             desc: "Familia del producto (deducida de la subfamilia)." },
                { col: "SubFamilia",          desc: "Subfamilia del producto. Al pasar el ratón muestra la descripción completa." },
                { col: "Coste Unit. (GA)",    desc: "Precio de coste unitario (precio GA). Con Estructura activa: formato 1.234,56 €." },
                { col: "Coste Posición",      desc: "Calculado: Coste Unitario × Cantidad. Con Estructura activa: formato 1.234,56 €." },
                { col: "Margen %",            desc: "Calculado: (Neto Pos. - Coste Pos.) / Neto Pos. × 100." },
                { col: "Id Posición",         desc: "Identificador del producto en BD (idproducto). Útil para trazabilidad." },
                { col: "Imagen",              desc: "Ruta o referencia a la imagen del producto." },
                { col: "Neto Unit. 2",        desc: "Precio neto alternativo (segunda tarifa). Con Estructura activa: formato 1.234,56 €." },
              ].map((r, i) => (
                <div key={i} style={{ display: "flex", gap: 16, padding: "8px", borderBottom: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#f8fafc" }}>
                  <span style={{ minWidth: 165, fontWeight: 600, color: "#1e3a5f", fontSize: 12, flexShrink: 0 }}>{r.col}</span>
                  <span style={{ color: "#475569", lineHeight: 1.6, fontSize: 12 }}>{r.desc}</span>
                </div>
              ))}
            </section>

            <section id="ayuda-acciones">
              <h2 style={{ fontSize: 16, fontWeight: 700, color: "#1e3a5f", marginBottom: 12, borderBottom: "2px solid #e2e8f0", paddingBottom: 8 }}>Acciones de la grid</h2>
              {[
                { accion: "Nueva fila",             desc: "Inserta encima de la fila con la celda seleccionada (rectángulo azul), empujando el resto hacia abajo. Si no hay celda seleccionada, da error." },
                { accion: "Campo numérico de filas", desc: "Indica cuántas filas insertar (por defecto 1, máximo 100)." },
                { accion: "Borrar seleccionadas",    desc: "Borra filas con checkbox marcado. Pide confirmación. Deshabilitado sin selección." },
                { accion: "Checkbox de fila",        desc: "Selecciona/deselecciona una fila. Se pueden marcar varias." },
                { accion: "Checkbox cabecera",       desc: "Selecciona o deselecciona todas las filas." },
                { accion: "Click en celda",          desc: "Selecciona la celda (rectángulo azul). Lista para Ctrl+C." },
                { accion: "Arrastrar con el ratón",  desc: "Selecciona un rango rectangular de celdas. Si arrastras hasta el borde del grid se hace auto-scroll en las 4 direcciones (estilo Excel)." },
                { accion: "Doble click en celda",    desc: "Activa edición. Repres. y Nat. abren desplegable. Referencia abre autocompletado." },
                { accion: "Enter / clic fuera",      desc: "Confirma la edición." },
                { accion: "Escape",                  desc: "Cancela la edición sin guardar." },
                { accion: "Ctrl+C",                  desc: "Copia al portapapeles la celda o rango seleccionado. Las columnas calculadas exportan su valor calculado actual. Formato TSV compatible con Excel." },
                { accion: "Ctrl+V",                  desc: "Pega desde el portapapeles. Si el contenido es tabular (de Excel u otra app) se pega en una matriz desde la celda actual hacia abajo y la derecha. Las columnas calculadas se omiten." },
                { accion: "Selector Larga/Corta",    desc: "En la cabecera del presupuesto. \"Larga\" muestra descripciones multilínea. \"Corta\" limita las celdas a una línea con tooltip al pasar el ratón para ver el texto completo." },
              ].map((r, i) => (
                <div key={i} style={{ display: "flex", gap: 16, padding: "9px 8px", borderBottom: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#f8fafc" }}>
                  <span style={{ minWidth: 190, fontWeight: 600, color: "#1e3a5f", fontSize: 12, flexShrink: 0 }}>{r.accion}</span>
                  <span style={{ color: "#475569", lineHeight: 1.6, fontSize: 12 }}>{r.desc}</span>
                </div>
              ))}
            </section>

            <section id="ayuda-statusbar">
              <h2 style={{ fontSize: 16, fontWeight: 700, color: "#1e3a5f", marginBottom: 12, borderBottom: "2px solid #e2e8f0", paddingBottom: 8 }}>Barra de estado</h2>
              <p style={{ color: "#475569", lineHeight: 1.7, marginBottom: 12 }}>
                En la parte inferior de la pantalla hay una barra que muestra mensajes de las acciones que ejecutas. Cambia de color según el tipo de mensaje:
              </p>
              {[
                { tipo: "Información (negro)",    color: "#0f172a", desc: "Estados generales: 'Listo', 'Mostrando resumen', etc." },
                { tipo: "Trabajando (azul)",      color: "#1e40af", desc: "Operación en curso: 'Cargando lista de presupuestos...', 'Generando Excel...'" },
                { tipo: "Éxito (verde)",          color: "#14532d", desc: "Operación completada. Se auto-borra a los 5 segundos." },
                { tipo: "Error (rojo)",           color: "#7f1d1d", desc: "Error en la operación. Se auto-borra a los 5 segundos." },
              ].map((r, i) => (
                <div key={i} style={{ display: "flex", gap: 12, padding: "8px", borderBottom: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#f8fafc", alignItems: "center" }}>
                  <span style={{ width: 12, height: 12, borderRadius: 2, background: r.color, flexShrink: 0 }} />
                  <span style={{ minWidth: 170, fontWeight: 600, color: "#1e3a5f", fontSize: 12, flexShrink: 0 }}>{r.tipo}</span>
                  <span style={{ color: "#475569", lineHeight: 1.6, fontSize: 12 }}>{r.desc}</span>
                </div>
              ))}
              <p style={{ color: "#475569", lineHeight: 1.7, marginTop: 14 }}>
                Justo encima hay una <strong>línea de información</strong> tipo Excel que muestra: <strong>Filas</strong> (total de líneas), <strong>Marcadas</strong> (filas con checkbox), <strong>Celdas sel.</strong> (nº de celdas del rectángulo azul) y <strong>Suma</strong> (sumatorio de los valores numéricos de las celdas seleccionadas).
              </p>
            </section>

            <section id="ayuda-autoref">
              <h2 style={{ fontSize: 16, fontWeight: 700, color: "#1e3a5f", marginBottom: 12, borderBottom: "2px solid #e2e8f0", paddingBottom: 8 }}>Autocompletado de referencias</h2>
              <p style={{ color: "#475569", lineHeight: 1.7, margin: "0 0 12px" }}>
                Al editar la columna <strong>Referencia</strong> con doble click, se activa el autocompletado conectado a la base de datos:
              </p>
              <ul style={{ color: "#475569", lineHeight: 1.7, paddingLeft: 20, marginBottom: 0 }}>
                <li>Al teclear <strong>2 o más caracteres</strong> aparece un tooltip con hasta 5 sugerencias mostrando referencia, nombre y PVP.</li>
                <li>Navega con las flechas <code>↑</code>/<code>↓</code> y selecciona con <code>Enter</code>.</li>
                <li>Al confirmar (Enter, click en sugerencia o salir del campo) se buscan los datos completos del producto.</li>
                <li>Si la referencia existe en BD, se rellenan automáticamente: <strong>nombre, descripción, PVP, familia, subfamilia y descuento</strong> (de la subfamilia, o de la familia si no hay).</li>
                <li>Si la referencia no existe, se guarda solo el texto sin sobrescribir nada (te deja crear referencias nuevas a mano).</li>
                <li><code>Escape</code> cancela la edición.</li>
              </ul>
            </section>

            {[
              { id: "m-presupuesto", titulo: "Menú Presupuesto", color: "#171717", items: [
                { op: "Guardar Presupuesto",    desc: "Guarda el presupuesto actual (cabecera y líneas) en la base de datos. Identifica el presupuesto por numerocompleto + revisión; si ya existe pregunta si sobrescribir." },
                { op: "Comprobar Presupuesto",  desc: "Verifica las referencias, precios y datos de las líneas contra la BD y marca las diferencias encontradas." },
                { op: "Leer Presupuestos",      desc: "Abre la lista de presupuestos guardados (buscar por título, cliente o número) y carga uno (sustituye el actual). Añade 10 filas en blanco al final del presupuesto cargado." },
                { op: "Importar",               desc: "Importa filas desde un Excel (fichero o copiar/pegar). Detecta las columnas por su nombre (acepta abreviaturas como Ref, Can, PVP, Dto). Si hay columnas ambiguas avisa con error." },
                { op: "Exportar",               desc: "Exporta el presupuesto a Excel con todos los datos crudos. El nombre del fichero usa el numerocompleto. Las opciones de la columna Representación NO se aplican aquí." },
                { op: "Imprimir",               desc: "Genera el Excel formato Siemens para el cliente. Etiquetas de cabecera en columna D; columnas B/C/E centradas; Neto Unitario/Posición con formato moneda €. Las representaciones \"conf\" (estilo CONF + texto \"A confirmar por el cliente\" en columna I) y \"TP\" (Total Posición: cantidad 1 y unitario = total) se aplican SOLO aquí. Nombre del fichero con numerocompleto + revisión." },
                { op: "Formato Simple Quote",   desc: "Prepara el presupuesto con el formato de tabla SimpleQuote. El nombre del fichero usa el numerocompleto. Las representaciones especiales NO se aplican aquí." },
                { op: "Crear SimpleQuote",      desc: "Envía la oferta a SimpleQuote (Siemens) a través de la API local (127.0.0.1:8000)." },
                { op: "Hacer Damex E",          desc: "Crea el pedido Damex E en Siemens con los datos del cliente final, vía la API local." },
                { op: "Resumen",                desc: "Muestra el desglose por Familia/SubFamilia con tabla y gráfico donut interactivo." },
                { op: "Aplicar Estructura",     desc: "Activa/desactiva el modo estructura: colorea las filas según su naturaleza (estilos configurables), calcula la numeración jerárquica de apartados (T1→1, productos→1.1/1.2, T2→1.3...), persiste esa numeración en la columna Apartado y la deja vacía en S1-S4/TT/CM. En S1-S4 calcula la suma del apartado y el descuento medio. Limpia la marca de PVP caducado (rojo). Al desactivar, la numeración y etiquetas se mantienen." },
                { op: "Borrar Presupuesto actual", desc: "Vacía todas las líneas dejando la cabecera. No borra el presupuesto guardado en la BD." },
                { op: "Comparar Presupuestos",  desc: "Compara dos presupuestos línea a línea." },
              ]},
              { id: "m-celdas", titulo: "Menú Celdas", color: "#171717", items: [
                { op: "Comprobar Celda",              desc: "Muestra el nº de caracteres, el color de fondo y el color de tinta de la celda seleccionada, con una vista previa." },
                { op: "Borrar filas vacías",          desc: "Abre un diálogo donde eliges una columna; borra las filas marcadas con checkbox que tengan esa columna vacía, pidiendo confirmación." },
                { op: "Borrar filas con 0",           desc: "Usa la columna seleccionada con el rectángulo azul (debe ser una sola columna) y borra las filas que tengan 0 en esa columna dentro de la selección." },
                { op: "Seleccionar Celdas",           desc: "Marca el checkbox de un rango de filas (desde la fila X hasta la fila Y)." },
                { op: "Convertir selección en comentario", desc: "Cambia la naturaleza de las filas seleccionadas a CM (comentario)." },
                { op: "Juntar celdas en una",         desc: "Une el contenido de varias celdas seleccionadas en una sola." },
              ]},
              { id: "m-elementos", titulo: "Menú Elementos", color: "#171717", items: [
                { op: "Guardar Elemento",       desc: "Toma las filas seleccionadas y las guarda en la BD como elemento reutilizable (pide nombre y descripción). Avisa si alguna referencia no existe en BD." },
                { op: "Leer Elemento",          desc: "Abre la lista de elementos guardados. Al insertarlo añade una fila de comentario (CM) con el nombre del elemento y debajo todas sus líneas de producto." },
              ]},
              { id: "m-productos", titulo: "Menú Productos", color: "#171717", items: [
                { op: "Guardar Producto",            desc: "Crea en el catálogo los productos de las filas seleccionadas (referencia, PVP>0, descripción y grupo descuento obligatorios). Si un producto YA EXISTE, abre un diálogo para elegir qué columnas actualizar (PVP, Descripción, Nombre, Precio coste, Grupo descuento) con opción de confirmar por cada fila (SÍ/NO)." },
                { op: "Leer Producto",               desc: "Busca un producto del catálogo y vuelca sus datos en la fila. La lista muestra columnas redimensionables (incluida Grupo Dto. con su descripción en tooltip). El PVP se marca en rojo si su fecha tiene más de 1 año." },
                { op: "Buscar datos por Referencia", desc: "Rellena nombre, PVP, descripción y datos de cada línea por su referencia. Marca el PVP en rojo si la fecha de actualización supera 1 año (hasta que se aplique estructura)." },
                { op: "Juntar productos duplicados", desc: "Agrupa en una sola línea los productos con la misma referencia, sumando las cantidades." },
                { op: "Quitar caracteres Referencia", desc: "Elimina espacios y caracteres no válidos de las referencias." },
                { op: "Quitar Saltos de línea",      desc: "Limpia los saltos de línea del texto de las celdas." },
                { op: "Leer Precios de PMD",         desc: "Consulta precios actualizados en PMD a través de la API local." },
                { op: "Buscar Referencia SIEMENS",   desc: "Detecta referencias Siemens (MLFB) dentro de un texto pegado (ignorando espacios y guiones) y las extrae a celdas." },
                { op: "Buscar equivalencia Competencia", desc: "Busca el producto de competencia equivalente a la referencia." },
                { op: "Calcular PVP a partir de GA", desc: "Calcula el PVP a partir del precio de coste GA y el margen." },
                { op: "Asistente Referencias",       desc: "Asistente para completar y corregir referencias." },
              ]},
              { id: "m-clientes", titulo: "Menú Clientes", color: "#171717", items: [
                { op: "Gestionar Clientes",     desc: "Tabla de clientes editable: crear, editar (cualquier campo salvo el ID), copiar y borrar. El borrado se bloquea si el cliente está en algún presupuesto o tiene contactos. El código postal se muestra siempre con 5 cifras (ceros a la izquierda)." },
                { op: "Gestionar Contactos",    desc: "Tabla de contactos: crear, editar y borrar. Cada contacto tiene nombre, cargo, email, Teléfono 1, Teléfono 2 y cliente asociado." },
              ]},
              { id: "m-descuentos", titulo: "Menú Descuentos", color: "#171717", items: [
                { op: "Aplicar descuentos",       desc: "Lista las combinaciones únicas de Grupo Descuento + descuento actual de las líneas (PD/PE/E), mostrando grupo, familia, subfamilia, nº de líneas, descuento e importe neto. Permite aplicar un nuevo descuento a toda una combinación." },
                { op: "Calcular Descuento",       desc: "Calcula, para cada celda numérica seleccionada, el descuento de su fila para que el neto unitario coincida con el valor objetivo. Admite descuentos negativos (recargo)." },
                { op: "Fijar el precio total",    desc: "Pide el importe neto total objetivo y reparte el descuento entre todas las líneas proporcionalmente a su neto actual." },
                { op: "Fijar precio de celdas",   desc: "Pide el importe neto objetivo y lo reparte solo entre las filas seleccionadas (por checkbox o rango) proporcionalmente a su neto actual." },
              ]},
              { id: "m-otros", titulo: "Menú Otros", color: "#171717", items: [
                { op: "Ayuda",                    desc: "Muestra este manual." },
                { op: "Opciones",                 desc: "Pantalla con menú lateral: (1) Configurar Estilos: personaliza color, fuente, tamaño y peso de cada naturaleza, incluidos PD (producto normal) y CONF (Confirmar por el cliente), con Exportar/Importar JSON. (2) Actualizar Tarifas: carga un Excel para crear/actualizar productos, con selector Sobrescribir SÍ/NO y log. (3) Mantenimiento BD: tablas de Clientes y Contactos (importar Excel, los contactos incluyen Teléfono 1 y 2), Log, recálculo de masusado, exportaciones y comprobación de integridad de datos. (4) Gestión de Usuarios: crear, editar y borrar usuarios con permisos y contraseña. (5) Configuraciones Varias." },
                { op: "Gestionar Estrategias Descuento", desc: "Consulta, crea, edita, borra y aplica estrategias de descuento (conjuntos de grupo descuento + % guardados) sobre el presupuesto." },
                { op: "Gestionar BD Competencia", desc: "Gestor de los productos de la competencia y sus equivalencias." },
              ]},
                        ].map(menu => (
              <section key={menu.id} id={"ayuda-" + menu.id}>
                <h2 style={{ fontSize: 16, fontWeight: 700, color: menu.color, marginBottom: 12, borderBottom: `2px solid ${menu.color}`, paddingBottom: 8 }}>{menu.titulo}</h2>
                {menu.items.map((r, i) => (
                  <div key={i} style={{ display: "flex", gap: 16, padding: "8px", borderBottom: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#f8fafc" }}>
                    <span style={{ minWidth: 210, fontWeight: 600, color: "#1e293b", fontSize: 12, flexShrink: 0 }}>{r.op}</span>
                    <span style={{ color: "#475569", lineHeight: 1.6, fontSize: 12 }}>{r.desc}</span>
                  </div>
                ))}
              </section>
            ))}

            <div style={{ background: "#f1f5f9", borderRadius: 8, padding: "12px 16px", fontSize: 12, color: "#64748b", textAlign: "center" }}>
              Para soporte técnico contacta con el administrador del sistema.
            </div>
          </div>
        </div>
      </div>
    </div>
  );

  // ── VISTA RESUMEN ──
  if (vista === "resumen") return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", fontSize: 13, color: "#1e293b", minHeight: "100vh", background: "#f8fafc" }}>
      <div style={{ background: "#f5f5f5", color: "#171717", padding: "8px 16px", display: "flex", alignItems: "center", gap: 12, borderBottom: "1px solid #e5e5e5" }}>
        <button onClick={() => setVista("grid")} style={{ background: "#fff", border: "1px solid #d4d4d4", color: "#171717", borderRadius: 6, padding: "4px 12px", cursor: "pointer", fontSize: 12 }}><BtnContent icon={ArrowLeft}>← Volver</BtnContent></button>
        <span style={{ fontWeight: 700, fontSize: 15, display: "inline-flex", alignItems: "center", gap: 8 }}><Icon as={BarChart3} size={18} color="#171717" /> Resumen del Presupuesto</span>
        <span style={{ color: "#737373", fontSize: 12 }}>{presupuesto.np} — {presupuesto.titulo}</span>
      </div>
      <div style={{ padding: "24px", display: "flex", gap: 32, flexWrap: "wrap", alignItems: "flex-start" }}>
        <div style={{ flex: 1, minWidth: 380 }}>
          <h3 style={{ fontSize: 14, fontWeight: 600, marginBottom: 12 }}>Desglose por Familia / SubFamilia</h3>
          <div style={{ border: "1px solid #e2e8f0", borderRadius: 8, overflow: "hidden" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr style={{ background: "#fafafa", borderBottom: "1px solid #e5e5e5" }}>
                  {["Familia", "SubFamilia", "Artículos", "Importe Neto", "%"].map((h, i) => (
                    <th key={i} style={{ padding: "8px 12px", textAlign: i > 1 ? "right" : "left", color: "#e2e8f0", fontSize: 12, fontWeight: 600 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {resumenData.map((row, i) => (
                  <tr key={i} style={{ borderBottom: "1px solid #e2e8f0", background: i % 2 === 0 ? "#fff" : "#f8fafc" }}>
                    <td style={{ padding: "7px 12px", display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{ width: 10, height: 10, borderRadius: "50%", background: COLORES[i % COLORES.length], display: "inline-block", flexShrink: 0 }} />{row.familia}
                    </td>
                    <td style={{ padding: "7px 12px", color: "#475569" }}>{row.subfamilia}</td>
                    <td style={{ padding: "7px 12px", textAlign: "right", fontWeight: 500 }}>{row.articulos}</td>
                    <td style={{ padding: "7px 12px", textAlign: "right", fontWeight: 500, color: "#0369a1" }}>{fmt(row.importe)} €</td>
                    <td style={{ padding: "7px 12px", textAlign: "right", color: "#64748b" }}>{totalResumen ? (row.importe / totalResumen * 100).toFixed(1) : 0}%</td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={{ background: "#f1f5f9", borderTop: "2px solid #cbd5e1" }}>
                  <td colSpan={2} style={{ padding: "8px 12px", fontWeight: 700 }}>TOTAL</td>
                  <td style={{ padding: "8px 12px", textAlign: "right", fontWeight: 700 }}>{resumenData.reduce((s, r) => s + r.articulos, 0)}</td>
                  <td style={{ padding: "8px 12px", textAlign: "right", fontWeight: 700, color: "#0369a1" }}>{fmt(totalResumen)} €</td>
                  <td style={{ padding: "8px 12px", textAlign: "right", fontWeight: 700 }}>100%</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
          <h3 style={{ fontSize: 14, fontWeight: 600, margin: 0 }}>Distribución por SubFamilia</h3>
          <DonutChart data={resumenData} total={totalResumen} onHover={setDonutTooltip} onLeave={() => setDonutTooltip(null)} />
          <div style={{ display: "flex", flexDirection: "column", gap: 5, width: "100%" }}>
            {resumenData.map((row, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 11 }}>
                <span style={{ width: 10, height: 10, borderRadius: 2, background: COLORES[i % COLORES.length], flexShrink: 0 }} />
                <span style={{ color: "#475569", flex: 1 }}>{row.subfamilia}</span>
                <span style={{ fontWeight: 500 }}>{totalResumen ? (row.importe / totalResumen * 100).toFixed(1) : 0}%</span>
              </div>
            ))}
          </div>
        </div>
      </div>
      {donutTooltip && (
        <div style={{ position: "fixed", left: donutTooltip.x + 12, top: donutTooltip.y - 10, background: "#1e293b", color: "#fff", borderRadius: 8, padding: "8px 12px", fontSize: 12, pointerEvents: "none", zIndex: 99999 }}>
          <div style={{ fontWeight: 600 }}>{donutTooltip.item.subfamilia}</div>
          <div style={{ color: "#93c5fd" }}>{donutTooltip.item.familia}</div>
          <div style={{ marginTop: 4 }}>{fmt(donutTooltip.item.importe)} € <span style={{ color: "#94a3b8" }}>({(donutTooltip.pct * 100).toFixed(1)}%)</span></div>
          <div style={{ color: "#94a3b8" }}>{donutTooltip.item.articulos} artículos</div>
        </div>
      )}
    </div>
  );

  // ── VISTA GRID ──
  // ── Pantalla de bienvenida ──
  if (showWelcome) {
    // Solo se entra a la app con un login correcto (userInfo con usuario válido)
    const entrar = (userInfo) => {
      if (!userInfo || !userInfo.usuario) return; // sin login válido no se accede
      try { sessionStorage.setItem("welcomeShown", "1"); } catch {}
      setUsuarioActual(userInfo.usuario);
      if (userInfo.codigopresupuestos) setCodigoUsuario(userInfo.codigopresupuestos);
      setShowWelcome(false);
    };
    return <WelcomeScreen onLogin={entrar} />;
  }

  return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", fontSize: 13, color: "#1e293b", height: "100vh", display: "flex", flexDirection: "column", background: "#f8fafc" }}>
      <div style={{ background: "#f5f5f5", color: "#171717", padding: "8px 16px", display: "flex", alignItems: "center", gap: 12, flexShrink: 0, borderBottom: "1px solid #e5e5e5" }}>
        <span style={{ fontWeight: 700, fontSize: 15, display: "inline-flex", alignItems: "center", gap: 8 }}><Icon as={FileSpreadsheet} size={18} color="#171717" /> Presupuestos</span>
        <span style={{ color: "#737373", fontSize: 12 }}>v2.00.0 (9 Junio 2026)</span>
        <span
          onClick={() => handleAction("AplicarEstructura")}
          title="Pulsa para activar o desactivar la estructura"
          style={{
            background: estructuraActiva ? "#dcfce7" : "#f1f5f9",
            color: estructuraActiva ? "#14532d" : "#64748b",
            fontSize: 11, padding: "2px 8px", borderRadius: 99,
            display: "inline-flex", alignItems: "center", gap: 4,
            border: `1px solid ${estructuraActiva ? "#86efac" : "#cbd5e1"}`,
            cursor: "pointer", userSelect: "none",
          }}>
          <Icon as={Palette} size={12} color={estructuraActiva ? "#14532d" : "#64748b"} />
          {estructuraActiva ? "Estructura activa" : "Estructura desactivada"}
        </span>
        <div style={{ marginLeft: "auto", position: "relative" }}>
          <button
            onClick={() => setUserMenuOpen(o => !o)}
            style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "4px 10px", background: "#fff", border: "1px solid #e5e5e5", borderRadius: 99, cursor: "pointer", fontSize: 12, color: "#171717", fontWeight: 500 }}
          >
            <Icon as={User} size={14} color="#525252" />
            {usuarioActual && <span>{usuarioActual}</span>}
          </button>
          {userMenuOpen && (
            <>
              {/* Overlay para cerrar al hacer click fuera */}
              <div onClick={() => setUserMenuOpen(false)}
                style={{ position: "fixed", inset: 0, zIndex: 99998, background: "transparent" }} />
              <div style={{ position: "absolute", top: "calc(100% + 4px)", right: 0, background: "#fff", border: "1px solid #e5e5e5", borderRadius: 6, boxShadow: "0 4px 12px rgba(0,0,0,0.1)", zIndex: 99999, minWidth: 180, padding: "4px 0" }}>
                <button
                  onClick={() => { setUserMenuOpen(false); setShowLoginDialog(true); }}
                  style={{ width: "100%", textAlign: "left", padding: "8px 14px", border: "none", background: "transparent", cursor: "pointer", fontSize: 12, color: "#171717", display: "flex", alignItems: "center", gap: 8 }}
                  onMouseEnter={e => e.currentTarget.style.background = "#f1f5f9"}
                  onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                >
                  <Icon as={LogIn} size={14} color="#1e3a5f" /> Iniciar sesión
                </button>
                <div style={{ height: 1, background: "#e5e5e5", margin: "2px 0" }} />
                <button
                  onClick={() => { setUserMenuOpen(false); setUsuarioActual(""); setCodigoUsuario("CDM"); setStatus("Sesión cerrada", "info"); }}
                  disabled={!usuarioActual}
                  style={{ width: "100%", textAlign: "left", padding: "8px 14px", border: "none", background: "transparent", cursor: usuarioActual ? "pointer" : "default", fontSize: 12, color: usuarioActual ? "#171717" : "#94a3b8", display: "flex", alignItems: "center", gap: 8 }}
                  onMouseEnter={e => { if (usuarioActual) e.currentTarget.style.background = "#f1f5f9"; }}
                  onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                >
                  <Icon as={LogOut} size={14} color={usuarioActual ? "#dc2626" : "#94a3b8"} /> Cerrar sesión
                </button>
              </div>
            </>
          )}
        </div>
      </div>
      <div style={{ background: "#fff", borderBottom: "1px solid #e2e8f0", display: "flex", alignItems: "stretch", padding: "0 8px", flexShrink: 0, boxShadow: "0 1px 4px rgba(0,0,0,0.06)" }}>
        {MENU_STRUCTURE.map(group => <MenuGroup key={group.id} group={group} onAction={handleAction} />)}
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 16, fontSize: 12, color: "#525252", paddingRight: 8 }}>
          <span>Neto: <strong style={{ color: "#171717" }}>{fmt(totalNeto)} €</strong></span>
          <span>Coste: <strong style={{ color: "#171717" }}>{fmt(totalCoste)} €</strong></span>
          <span>Margen: <strong style={{ color: margenTotal > 30 ? "#16a34a" : "#dc2626" }}>{fmt(margenTotal)}%</strong></span>
        </div>
      </div>
      <div style={{ background: "#fff", borderBottom: "1px solid #e2e8f0", padding: "10px 16px", display: "flex", gap: 24, flexShrink: 0, alignItems: "center", flexWrap: "wrap" }}>
        {[
          { label: "ID", key: "id", width: 60, readonly: true },
          { label: "Nº Completo", key: "numerocompleto", width: 160, readonly: true },
          { label: "Nº", key: "np", width: 70 },
          { label: "Rev.", key: "revision", width: 30 },
          { label: "Año", key: "anopresupuesto", width: 60 },
          { label: "Título", key: "titulo", width: 240 },
        ].map(f => (
          <div key={f.key} style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <label style={{ fontSize: 11, color: "#64748b", fontWeight: 500, whiteSpace: "nowrap" }}>{f.label}</label>
            <input value={presupuesto[f.key] || ""} readOnly={f.readonly}
              onChange={e => {
                if (f.readonly) return;
                const newValue = e.target.value;
                setPresupuesto(p => {
                  const updated = { ...p, [f.key]: newValue };
                  // Si cambia np o anopresupuesto, recalcular numerocompleto
                  if (f.key === "np" || f.key === "anopresupuesto") {
                    updated.numerocompleto = buildNumeroCompleto(codigoUsuario, updated.np, updated.anopresupuesto);
                  }
                  return updated;
                });
              }}
              style={{ width: f.width, padding: "4px 8px", borderRadius: 4, border: "1px solid #cbd5e1", fontSize: 12, color: f.readonly ? "#94a3b8" : "#1e293b", background: f.readonly ? "#f1f5f9" : "#f8fafc", cursor: f.readonly ? "default" : "text", fontWeight: f.readonly ? 600 : 400 }} />
            {f.key === "numerocompleto" && (
              <button
                onClick={() => {
                  const nuevo = buildNumeroCompleto(codigoUsuario, presupuesto.np, presupuesto.anopresupuesto);
                  setPresupuesto(p => ({ ...p, numerocompleto: nuevo }));
                  setStatus(`Número completo recalculado: ${nuevo}`, "success");
                }}
                title="Recalcular número completo (código usuario + nº + año)"
                style={{ padding: "3px 6px", borderRadius: 4, border: "1px solid #d4d4d4", background: "#fff", cursor: "pointer", display: "inline-flex", alignItems: "center" }}>
                <Icon as={Calculator} size={14} color="#1e3a5f" />
              </button>
            )}
          </div>
        ))}
        {/* Cliente */}
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <label style={{ fontSize: 11, color: "#64748b", fontWeight: 500, whiteSpace: "nowrap" }}>Cliente:</label>
          <button onClick={() => setShowSelecCliente(true)}
            title="Seleccionar cliente"
            style={{
              padding: "4px 10px",
              borderRadius: 4,
              border: "1px solid #cbd5e1",
              background: presupuesto.idcliente ? "#eff6ff" : "#f8fafc",
              color: presupuesto.idcliente ? "#1e3a5f" : "#94a3b8",
              cursor: "pointer",
              fontSize: 12,
              width: 220,
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
            }}>
            <Icon as={User} size={12} color={presupuesto.idcliente ? "#1e3a5f" : "#94a3b8"} />
            {presupuesto.idcliente ? (presupuesto.cliente || "Cliente sin nombre") : "Seleccionar cliente"}
          </button>
          {presupuesto.idcliente && (
            <button onClick={() => setPresupuesto(p => ({ ...p, idcliente: null, cliente: "", alaatencion: null, alaatencion_nombre: "", alaatencion_cliente: "" }))}
              title="Quitar cliente (también quita el contacto)"
              style={{ padding: "3px 6px", borderRadius: 4, border: "1px solid #d4d4d4", background: "#fff", cursor: "pointer", fontSize: 11 }}>
              <Icon as={X} size={10} color="#94a3b8" />
            </button>
          )}
        </div>
        {/* A la atención de */}
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <label style={{ fontSize: 11, color: "#64748b", fontWeight: 500, whiteSpace: "nowrap" }}>A la atención de:</label>
          <button onClick={() => setShowSelecContacto(true)}
            title="Seleccionar contacto del cliente"
            style={{
              padding: "4px 10px",
              borderRadius: 4,
              border: "1px solid #cbd5e1",
              background: presupuesto.alaatencion ? "#eff6ff" : "#f8fafc",
              color: presupuesto.alaatencion ? "#1e3a5f" : "#94a3b8",
              cursor: "pointer",
              fontSize: 12,
              maxWidth: 280,
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
            }}>
            <Icon as={User} size={12} color={presupuesto.alaatencion ? "#1e3a5f" : "#94a3b8"} />
            {presupuesto.alaatencion
              ? `${presupuesto.alaatencion_nombre || "Contacto"}${presupuesto.alaatencion_cliente ? " - " + presupuesto.alaatencion_cliente : ""}`
              : "Seleccionar contacto"}
          </button>
          {presupuesto.alaatencion && (
            <button onClick={() => setPresupuesto(p => ({ ...p, alaatencion: null, alaatencion_nombre: "", alaatencion_cliente: "" }))}
              title="Quitar contacto"
              style={{ padding: "3px 6px", borderRadius: 4, border: "1px solid #d4d4d4", background: "#fff", cursor: "pointer", fontSize: 11 }}>
              <Icon as={X} size={10} color="#94a3b8" />
            </button>
          )}
        </div>

        {/* País cliente final */}
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <label style={{ fontSize: 11, color: "#64748b", fontWeight: 500, whiteSpace: "nowrap" }}>País cliente final:</label>
          <select
            value={presupuesto.idpaisclientefinal || ""}
            onChange={e => setPresupuesto(p => ({ ...p, idpaisclientefinal: e.target.value ? Number(e.target.value) : null }))}
            style={{ width: 180, padding: "4px 8px", borderRadius: 4, border: "1px solid #cbd5e1", fontSize: 12, color: "#1e293b", background: "#f8fafc", cursor: "pointer" }}>
            <option value="">— Seleccionar —</option>
            {paisesList.map(p => (
              <option key={p.id} value={p.id}>{p.codigo_iso2} - {p.pais}</option>
            ))}
          </select>
        </div>

        {/* Cliente final */}
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <label style={{ fontSize: 11, color: "#64748b", fontWeight: 500, whiteSpace: "nowrap" }}>Cliente final:</label>
          <input
            value={presupuesto.clientefinal || ""}
            onChange={e => setPresupuesto(p => ({ ...p, clientefinal: e.target.value }))}
            placeholder="Razón social del cliente final"
            style={{ width: 240, padding: "4px 8px", borderRadius: 4, border: "1px solid #cbd5e1", fontSize: 12, color: "#1e293b", background: "#f8fafc" }} />
        </div>

        {/* Selector descripción corta/larga */}
        <div style={{ display: "flex", alignItems: "center", gap: 6, marginLeft: "auto" }}>
          <label style={{ fontSize: 11, color: "#64748b", fontWeight: 500, whiteSpace: "nowrap" }}>Descripción</label>
          <div style={{ display: "inline-flex", border: "1px solid #d4d4d4", borderRadius: 6, overflow: "hidden" }}>
            <button onClick={() => setDescCorta(false)}
              title="Mostrar la descripción completa con varias líneas"
              style={{ padding: "4px 12px", fontSize: 11, border: "none", cursor: "pointer",
                background: !descCorta ? "#171717" : "#fff",
                color: !descCorta ? "#fff" : "#171717",
                fontWeight: !descCorta ? 600 : 400 }}>
              Larga
            </button>
            <button onClick={() => setDescCorta(true)}
              title="Limita las filas a 1 línea (pasa el ratón por encima para ver toda la descripción)"
              style={{ padding: "4px 12px", fontSize: 11, border: "none", borderLeft: "1px solid #d4d4d4", cursor: "pointer",
                background: descCorta ? "#171717" : "#fff",
                color: descCorta ? "#fff" : "#171717",
                fontWeight: descCorta ? 600 : 400 }}>
              Corta
            </button>
          </div>
        </div>
      </div>
      <div ref={tableContainerRef} style={{ flex: 1, overflow: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "max-content", minWidth: "100%" }}>
          <thead style={{ position: "sticky", top: 0, zIndex: 10 }}>
            <tr>
              <th style={{ width: 32, background: "#fafafa", borderBottom: "1px solid #e5e5e5", borderRight: "1px solid #e5e5e5" }}>
                <input type="checkbox" onChange={e => { if (e.target.checked) setSelectedRows(new Set(rows.map(r => r.id))); else setSelectedRows(new Set()); }} />
              </th>
              <th style={{ width: 32, background: "#fafafa", color: "#737373", fontSize: 10, borderBottom: "1px solid #e5e5e5", borderRight: "1px solid #e5e5e5" }}>#</th>
              {COLUMNS.map(col => {
                const w = getColWidth(col);
                return (
                <th key={col.key} title={col.tooltip || (Array.isArray(col.label) ? col.label.join(" ") : col.label)}
                    style={{ position: "relative", width: w, minWidth: w, maxWidth: w, padding: "6px 8px", background: "#fafafa", color: "#171717", fontWeight: 600, fontSize: 11, textAlign: col.align || (col.type === "number" || col.type === "calc" ? "right" : "left"), borderBottom: "1px solid #e5e5e5", borderRight: "1px solid #e5e5e5", whiteSpace: "nowrap", lineHeight: 1.15 }}>
                  {Array.isArray(col.label) ? col.label.map((l, i) => <div key={i}>{l}</div>) : col.label}
                  {/* Resizer del borde derecho */}
                  <div
                    onMouseDown={(e) => onResizeStart(e, "col", col.key, w)}
                    onDoubleClick={(e) => { e.stopPropagation(); setAnchosManual(prev => { const n = { ...prev }; delete n[col.key]; return n; }); }}
                    title="Arrastra para redimensionar la columna (doble clic para restablecer)"
                    onMouseEnter={e => { e.currentTarget.style.background = "#2563eb"; }}
                    onMouseLeave={e => { e.currentTarget.style.background = "transparent"; }}
                    style={{ position: "absolute", top: 0, right: -4, width: 9, height: "100%", cursor: "col-resize", zIndex: 6, transition: "background 0.1s" }}
                  />
                </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rowIdx) => {
              const estilo = estructuraActiva ? getEstiloFila(row.naturaleza, estilos) : { bg: null, color: "#1e293b", fontFamily: "Segoe UI", fontWeight: 400, fontSize: 12 };
              const colsVisibles = estructuraActiva
                ? (estilo.isTitle ? COLS_TITULO : estilo.isSubtotal ? COLS_SUBTOTAL : estilo.isComment ? COLS_COMENT : null)
                : null;
              const esTotal = row.naturaleza === "TT";
              const esSubtotal = SUBTOTAL_NIVELES[row.naturaleza] !== undefined || esTotal;
              const subtotalCalc = esTotal ? calcTotalPresupuesto(rows)
                : (SUBTOTAL_NIVELES[row.naturaleza] !== undefined ? calcSubtotal(rows, rowIdx, SUBTOTAL_NIVELES[row.naturaleza]) : null);
              // Cuando estructura está activa, solo aplicar estilos en columnas específicas:
              //   T1-T4 (títulos): "posicion" y "nombre"
              //   S1-S4 (subtotales): "nombre", "dtoaplicado", "precionetoposicion"
              //   CM (comentario): solo "nombre"
              //   TT, VERDE, GRIS: toda la fila (mantener comportamiento anterior)
              const COLS_ESTILO_TITULO   = ["posicion", "nombre"];
              const COLS_ESTILO_SUBTOTAL = ["nombre", "dtoaplicado", "precionetoposicion"];
              const COLS_ESTILO_COMENT   = ["nombre"];
              const colsConEstilo = estructuraActiva
                ? (estilo.isTitle ? COLS_ESTILO_TITULO
                  : estilo.isSubtotal ? COLS_ESTILO_SUBTOTAL
                  : estilo.isComment ? COLS_ESTILO_COMENT
                  : null)
                : null;
              const aplicarEstiloFila = !colsConEstilo; // null → estilo en toda la fila
              const rowBgDefault = selectedRows.has(row.id) ? "#dbeafe" : (rowIdx % 2 === 0 ? "#fff" : "#f8fafc");
              const rowBg = aplicarEstiloFila ? (selectedRows.has(row.id) ? "#dbeafe" : estilo.bg || rowBgDefault) : rowBgDefault;
              return (
                <tr key={row.id} style={{ background: rowBg, height: altosManual[row.id] || undefined }}>
                  <td style={{ textAlign: "center", border: "1px solid #e2e8f0", padding: "2px 4px", background: rowBg }}>
                    <input type="checkbox" checked={selectedRows.has(row.id)} onChange={e => { const s = new Set(selectedRows); if (e.target.checked) s.add(row.id); else s.delete(row.id); setSelectedRows(s); }} />
                  </td>
                  <td style={{ position: "relative", textAlign: "center", border: "1px solid #e2e8f0", padding: "2px 6px", fontSize: 11, color: estilo.color || "#94a3b8", background: rowBg }}>
                    {rowIdx + 1}
                    {/* Resizer del borde inferior (alto de fila) */}
                    <div
                      onMouseDown={(e) => {
                        const tr = e.currentTarget.closest("tr");
                        const h = tr ? tr.offsetHeight : 24;
                        onResizeStart(e, "row", row.id, h);
                      }}
                      title="Arrastra para redimensionar la altura de la fila"
                      style={{ position: "absolute", left: 0, bottom: -3, width: "100%", height: 6, cursor: "row-resize", zIndex: 5 }}
                    />
                  </td>
                  {COLUMNS.map(col => {
                    const isEditing = editingCell?.rowId === row.id && editingCell?.colKey === col.key;
                    const isSelected = selectedCell?.rowId === row.id && selectedCell?.colKey === col.key;
                    const isCalc = col.type === "calc";
                    const isRight = col.type === "number" || col.type === "calc";
                    const isHidden = colsVisibles && !colsVisibles.includes(col.key);
                    let displayVal = row[col.key] ?? "";
                    if (col.key === "posicion" && estructuraActiva) {
                      // Con estructura: mostrar la numeración jerárquica calculada.
                      // En S1-S4/TT/CM queda vacía (calcApartados devuelve "").
                      displayVal = apartados[row.id] != null ? apartados[row.id] : "";
                    }
                    // Columnas cuyo valor es un importe en euros (coma decimal, punto de miles y símbolo €)
                    const COLS_EURO = ["pvp", "precionetounitario", "precionetoposicion", "preciocosteunitario", "costeposicion", "precionetounitario2"];
                    // Las etiquetas de S1-S4/TT se asignan al activar estructura
                    // y se persisten en rows.nombre (ver useEffect de estructuraActiva).
                    if (col.key === "precionetounitario") displayVal = estructuraActiva ? fmtEur(calcNetoUnit(row)) : fmtEur(calcNetoUnit(row));
                    if (col.key === "precionetoposicion") {
                      const val = esSubtotal ? subtotalCalc.neto : calcNetoPos(row);
                      displayVal = fmtEur(val);
                    }
                    if (col.key === "dtoaplicado" && esSubtotal) displayVal = fmt(subtotalCalc.dto) + " %";
                    else if (col.key === "dtoaplicado" && typeof displayVal === "number") displayVal = fmt(displayVal) + " %";
                    if (col.key === "costeposicion") displayVal = fmtEur(calcCostePos(row));
                    if (col.key === "margen") displayVal = fmt(calcMargen(row)) + " %";
                    if (col.key === "cantidad" && typeof displayVal === "number") displayVal = displayVal.toLocaleString("es-ES");
                    // Importes en euros: siempre con formato € (coma decimal + punto de miles), con o sin estructura
                    if (COLS_EURO.includes(col.key) && typeof displayVal === "number") displayVal = fmtEur(displayVal);
                    else if (col.type === "number" && typeof displayVal === "number") displayVal = fmt(displayVal);
                    // Decidir si esta celda recibe el estilo de la naturaleza
                    const celdaConEstilo = colsConEstilo ? colsConEstilo.includes(col.key) : aplicarEstiloFila;
                    const cellBg = isHidden
                      ? (celdaConEstilo ? (estilo.bg || rowBg) : rowBg)
                      : isCalc && !colsVisibles ? "#f0f9ff"
                      : celdaConEstilo ? (estilo.bg || "inherit") : rowBgDefault;
                    const isNombre = col.key === "nombre";
                    const colIdx = COLUMNS.findIndex(c => c.key === col.key);
                    const inRange = isCellInRange(rowIdx, colIdx);
                    return (
                      <td key={col.key}
                        title={(() => {
                          // Tooltip con la descripción para grupodescuento/familia/subfamilia (si tienen contenido)
                          const v = String(row[col.key] ?? "").trim();
                          if (!v) return undefined;
                          const k = v.toUpperCase();
                          if (col.key === "grupodescuento") return descGrupos[k] ? `${v} — ${descGrupos[k]}` : undefined;
                          if (col.key === "familia") return descFamilias[k] ? `${v} — ${descFamilias[k]}` : undefined;
                          if (col.key === "subfamilia") return descSubfamilias[k] ? `${v} — ${descSubfamilias[k]}` : undefined;
                          if (col.key === "pvp" && row.fechapvp) return `PVP actualizado: ${fmtFecha(row.fechapvp)}` + (row.pvpVencido ? " (hace más de 1 año)" : "");
                          if (col.key === "preciocosteunitario" && row.fechapreciocoste) return `Precio coste actualizado: ${fmtFecha(row.fechapreciocoste)}`;
                          return undefined;
                        })()}
                        onMouseDown={e => {
                          // Si esta celda ya está en edición, no interferir (permitir seleccionar texto en el input)
                          if (isEditing) return;
                          e.preventDefault();
                          setSelectedCell({ rowId: row.id, colKey: col.key });
                          setSelectionRange({ startRowIdx: rowIdx, startColIdx: colIdx, endRowIdx: rowIdx, endColIdx: colIdx });
                          setIsSelecting(true);
                        }}
                        onMouseEnter={() => {
                          if (isSelecting) {
                            setSelectionRange(r => r ? { ...r, endRowIdx: rowIdx, endColIdx: colIdx } : null);
                          }
                        }}
                        onMouseUp={() => setIsSelecting(false)}
                        onDoubleClick={() => { if (!isCalc && !isHidden) startEdit(row.id, col.key, row[col.key]); }}
                        style={{ position: "relative", border: isSelected || inRange ? "2px solid #2563eb" : "1px solid #e2e8f0", padding: 0, width: getColWidth(col), minWidth: getColWidth(col), maxWidth: getColWidth(col), background: inRange && !isSelected ? "#dbeafe" : cellBg, cursor: isCalc || isHidden ? "default" : "cell", verticalAlign: "middle", overflow: (isEditing && (col.type === "select" || col.key === "referencia" || col.key === "nombre" || col.key === "descripcion")) ? "visible" : "hidden" }}>
                        {isHidden ? null : isEditing && col.type === "select" ? (
                          (() => {
                            const opciones = col.opciones === "naturaleza"
                              ? [{ value: "", label: "( En blanco )", desc: "" }, ...NATURALEZA_OPCIONES.map(n => ({ value: n.naturaleza, label: n.naturaleza, desc: n.descripcion }))]
                              : [{ value: "", label: "( Sin representación )", desc: "" }, ...REPRESENTACION_OPCIONES.map(n => ({ value: n.representacion, label: n.representacion, desc: n.descripcion }))];
                            const filtroTxt = String(editValue || "").toUpperCase().trim();
                            const opcionesFiltradas = filtroTxt
                              ? opciones.filter(o => String(o.value || "").toUpperCase().includes(filtroTxt) || String(o.desc || "").toUpperCase().includes(filtroTxt))
                              : opciones;
                            // Confirma: si el texto coincide con un value válido (case insensitive), se acepta; si no, se cancela
                            const tryConfirm = () => {
                              // Si se pulsó Escape, descartar sin aplicar
                              if (escapeEditRef.current) { escapeEditRef.current = false; setEditingCell(null); return; }
                              const exacta = opciones.find(o => String(o.value || "").toUpperCase() === filtroTxt);
                              if (filtroTxt === "") {
                                setRows(rs => rs.map(rr => rr.id === editingCell.rowId ? { ...rr, [editingCell.colKey]: "" } : rr));
                                setEditingCell(null);
                                return;
                              }
                              if (exacta) {
                                setRows(rs => rs.map(rr => rr.id === editingCell.rowId ? { ...rr, [editingCell.colKey]: exacta.value } : rr));
                                setEditingCell(null);
                                return;
                              }
                              // Si solo queda una opción tras filtrar, la aplicamos
                              if (opcionesFiltradas.length === 1 && opcionesFiltradas[0].value) {
                                setRows(rs => rs.map(rr => rr.id === editingCell.rowId ? { ...rr, [editingCell.colKey]: opcionesFiltradas[0].value } : rr));
                                setEditingCell(null);
                                return;
                              }
                              // No válido → cancelar (mantener valor previo)
                              setEditingCell(null);
                            };
                            return (
                              <>
                                {/* Input editable que filtra la lista y valida contra valores permitidos */}
                                <input autoFocus value={editValue ?? ""}
                                  onChange={e => setEditValue(e.target.value.toUpperCase())}
                                  onMouseDown={e => e.stopPropagation()}
                                  onKeyDown={e => {
                                    e.stopPropagation();
                                    if (e.key === "Enter") { e.preventDefault(); tryConfirm(); }
                                    if (e.key === "Escape") { escapeEditRef.current = true; setEditingCell(null); }
                                  }}
                                  onBlur={tryConfirm}
                                  style={{ width: "100%", padding: "4px 8px", fontSize: 12, border: "2px solid #2563eb", outline: "none", background: "#fff", textAlign: col.align || "center", textTransform: "uppercase", boxSizing: "border-box" }} />
                                {/* Capa que cierra al hacer click fuera */}
                                <div onMouseDown={(e) => { e.stopPropagation(); setEditingCell(null); }}
                                  style={{ position: "fixed", inset: 0, zIndex: 999 }} />
                                {/* Lista desplegable (filtrada por el texto del input) */}
                                <div style={{ position: "absolute", top: "100%", left: 0, minWidth: 380, maxHeight: 320, overflow: "auto", background: "#fff", border: "1px solid #2563eb", borderRadius: 4, boxShadow: "0 6px 20px rgba(0,0,0,0.18)", zIndex: 1000 }}>
                                  {opcionesFiltradas.length === 0 ? <div style={{ padding: "8px 12px", fontSize: 11, color: "#94a3b8", fontStyle: "italic" }}>Sin coincidencias</div> : opcionesFiltradas.map(opt => (
                                    <div key={opt.value || "__empty"}
                                      onMouseDown={(e) => {
                                        e.preventDefault();
                                        e.stopPropagation();
                                        setRows(rs => rs.map(rr => rr.id === editingCell.rowId ? { ...rr, [editingCell.colKey]: opt.value } : rr));
                                        setEditingCell(null);
                                      }}
                                      style={{ padding: "6px 12px", fontSize: 12, cursor: "pointer", borderBottom: "1px solid #f1f5f9", display: "flex", alignItems: "center", gap: 8,
                                        background: opt.value === editValue ? "#dbeafe" : "transparent" }}
                                      onMouseEnter={e => { if (opt.value !== editValue) e.currentTarget.style.background = "#f1f5f9"; }}
                                      onMouseLeave={e => { if (opt.value !== editValue) e.currentTarget.style.background = "transparent"; }}>
                                      <strong style={{ minWidth: 60, color: "#1e3a5f" }}>{opt.label}</strong>
                                      <span style={{ color: "#64748b", fontSize: 11 }}>{opt.desc}</span>
                                    </div>
                                  ))}
                                </div>
                              </>
                            );
                          })()
                        ) : isEditing ? (
                          col.key === "referencia" ? (
                            <AutocompleteReferencia
                              value={editValue}
                              onChange={setEditValue}
                              onConfirm={(val) => commitEdit(val)}
                              onCancel={() => setEditingCell(null)}
                              align={col.align || (isRight ? "right" : "left")}
                            />
                          ) : (
                            (() => {
                              // Columnas de texto largo → editor multilínea (textarea) más alto que la celda
                              const esTextoLargo = ["referencia", "nombre", "descripcion"].includes(col.key);
                              if (esTextoLargo) {
                                // Calcular rows según contenido: saltos de línea + estimación por ancho del editor
                                const txt = String(editValue ?? "");
                                const anchoEditor = Math.max(col.width, 280);
                                const charsPorLinea = Math.max(20, Math.floor((anchoEditor - 20) / 6.5));
                                const lineasTxt = txt.split("\n");
                                let lineasEstimadas = 0;
                                lineasTxt.forEach(l => {
                                  lineasEstimadas += Math.max(1, Math.ceil(l.length / charsPorLinea));
                                });
                                const rowsCalc = Math.min(Math.max(3, lineasEstimadas + 1), 20);
                                return (
                                  <textarea autoFocus value={editValue}
                                    onChange={e => setEditValue(e.target.value)} onBlur={commitEdit}
                                    onMouseDown={e => e.stopPropagation()}
                                    onKeyDown={e => {
                                      e.stopPropagation();
                                      // Enter confirma; Shift+Enter inserta salto de línea
                                      if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); commitEdit(); }
                                      if (e.key === "Escape") { escapeEditRef.current = true; setEditingCell(null); }
                                    }}
                                    style={{
                                      position: "absolute", top: 0, left: 0, zIndex: 50,
                                      width: anchoEditor, height: "auto", minHeight: 70,
                                      border: "2px solid #2563eb", outline: "none", padding: "4px 8px",
                                      fontSize: 12, fontFamily: "inherit", background: "#fff",
                                      textAlign: "left", resize: "both", lineHeight: 1.4,
                                      boxShadow: "0 4px 12px rgba(0,0,0,0.15)", borderRadius: 2,
                                    }}
                                    rows={rowsCalc} />
                                );
                              }
                              return (
                                <input autoFocus value={editValue} maxLength={col.maxLength || undefined}
                                  onChange={e => setEditValue(e.target.value)} onBlur={commitEdit}
                                  onMouseDown={e => e.stopPropagation()}
                                  onKeyDown={e => { e.stopPropagation(); if (e.key === "Enter") commitEdit(); if (e.key === "Escape") { escapeEditRef.current = true; setEditingCell(null); } }}
                                  style={{ width: "100%", border: "none", outline: "none", padding: "4px 8px", fontSize: 12, fontFamily: "inherit", background: "#fff", textAlign: col.align || (isRight ? "right" : "left") }} />
                              );
                            })()
                          )
                        ) : (
                          (() => {
                            const esColTextoEstructura = (estructuraActiva || wrapPersistente) && COLS_WRAP_ESTRUCTURA.includes(col.key);
                            const haceWrap = (col.wrap || esColTextoEstructura) && !descCorta;
                            // Importe negativo en rojo cuando estructura activa.
                            // Calculamos el valor numérico real (no el displayVal que ya está formateado)
                            const COLS_IMPORTE = ["pvp", "dtoaplicado", "precionetounitario", "precionetoposicion", "preciocosteunitario", "costeposicion", "precionetounitario2"];
                            let valNum = null;
                            if (estructuraActiva && COLS_IMPORTE.includes(col.key)) {
                              if (col.key === "precionetounitario") valNum = calcNetoUnit(row);
                              else if (col.key === "precionetoposicion") valNum = esSubtotal ? subtotalCalc.neto : calcNetoPos(row);
                              else if (col.key === "costeposicion") valNum = calcCostePos(row);
                              else valNum = Number(row[col.key]) || 0;
                            }
                            const importeNegativo = valNum != null && valNum < 0;
                            const tooltipFull =
                              (col.key === "subfamilia" && row.descripcionsubfamilia) ? ((row.subfamilia || "") + ": " + row.descripcionsubfamilia)
                              : (col.key === "grupodescuento" && row.descripciongrupodescuento) ? ((row.grupodescuento || "") + ": " + row.descripciongrupodescuento)
                              : (["referencia", "nombre", "descripcion"].includes(col.key) ? String(row[col.key] ?? "") : ((col.wrap && descCorta && typeof displayVal === "string") ? displayVal : undefined));
                            return (
                              <div title={tooltipFull} style={{
                                padding: "4px 8px",
                                fontSize: celdaConEstilo && estilo.fontSize ? estilo.fontSize : 12,
                                whiteSpace: haceWrap ? "normal" : "nowrap",
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                wordBreak: haceWrap ? "break-word" : "normal",
                                lineHeight: haceWrap ? 1.4 : 1.2,
                                textAlign: (estructuraActiva && col.key === "nombre" && (estilo.isTitle || estilo.isComment)) ? "center"
                                  : (estructuraActiva && col.key === "nombre" && estilo.isSubtotal) ? "left"
                                  : col.align || (isRight ? "right" : "left"),
                                color: (col.key === "pvp" && row.pvpVencido) ? "#dc2626" : importeNegativo ? "#dc2626" : celdaConEstilo && (estilo.isTitle || estilo.isSubtotal || estilo.isComment) ? estilo.color : isCalc ? "#0369a1" : "#1e293b",
                                fontFamily: celdaConEstilo && estilo.fontFamily ? estilo.fontFamily : "inherit",
                                fontWeight: (col.key === "pvp" && row.pvpVencido) ? 700 : celdaConEstilo && (estilo.isTitle || estilo.isSubtotal) ? estilo.fontWeight : isCalc ? 500 : 400
                              }}>
                                {isHidden ? "" : displayVal}
                              </div>
                            );
                          })()
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {/* Toolbar fija de acciones del grid */}
      <div style={{ background: "#fafafa", borderTop: "1px solid #e5e5e5", padding: "5px 16px", flexShrink: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <button onClick={() => {
                    // Inserta encima de la fila que contiene la celda seleccionada (rectángulo azul).
                    // Si no hay celda seleccionada, da error.
                    if (!selectedCell) {
                      setStatus("Selecciona una celda (clic en una celda) para indicar dónde insertar la nueva fila", "error");
                      return;
                    }
                    const idx = rows.findIndex(row => row.id === selectedCell.rowId);
                    if (idx === -1) {
                      setStatus("No se encuentra la fila de la celda seleccionada", "error");
                      return;
                    }
                    const n = parseInt(numFilas) || 1;
                    const nuevas = Array.from({ length: n }, () => ({ id: nextId.current++, representacion: "", naturaleza: "", posicion: "", cantidad: "", referencia: "", nombre: "", pvp: 0, dtoaplicado: 0, descripcion: "", familia: "", subfamilia: "", preciocosteunitario: 0, idposicion: "", imagen: "", precionetounitario2: 0, grupodescuento: "" }));
                    setRows(r => {
                      const res = [...r];
                      res.splice(idx, 0, ...nuevas);
                      return res;
                    });
                    setStatus(`${n} fila(s) insertada(s) encima de la fila ${idx + 1}`, "success");
                  }} style={{ padding: "7px 10px", fontSize: 12, cursor: "pointer", borderRadius: 4, border: "1px solid #cbd5e1", background: "#fff", color: "#1e293b", whiteSpace: "nowrap" }}>
                    <BtnContent icon={Plus} iconColor="#16a34a">Nueva fila</BtnContent>
                  </button>
                  <input type="number" min="1" max="100" value={numFilas} onChange={e => setNumFilas(e.target.value)}
                    style={{ width: 52, padding: "4px 6px", fontSize: 12, borderRadius: 4, border: "1px solid #cbd5e1", background: "#fff", textAlign: "center" }} />
                  <span style={{ fontSize: 11, color: "#94a3b8" }}>filas</span>
                  <div style={{ width: 1, height: 20, background: "#e2e8f0", margin: "0 4px" }} />
                  <button
                    onClick={() => {
                      // Determinar filas afectadas: rango de celdas o celda activa
                      const idsFila = new Set();
                      if (selectionRange) {
                        const minRow = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
                        const maxRow = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
                        for (let r = minRow; r <= maxRow; r++) {
                          if (rows[r]) idsFila.add(rows[r].id);
                        }
                      } else if (selectedCell) {
                        idsFila.add(selectedCell.rowId);
                      }
                      if (idsFila.size === 0) {
                        setStatus("Selecciona al menos una celda o rango para marcar sus filas", "error");
                        return;
                      }
                      setSelectedRows(prev => {
                        const next = new Set(prev);
                        idsFila.forEach(id => next.add(id));
                        return next;
                      });
                      setStatus(`Marcadas ${idsFila.size} fila${idsFila.size !== 1 ? "s" : ""}`, "success");
                    }}
                    disabled={!selectionRange && !selectedCell}
                    title="Marca el checkbox de las filas de las celdas seleccionadas"
                    style={{ padding: "7px 10px", fontSize: 12, cursor: (!selectionRange && !selectedCell) ? "default" : "pointer", borderRadius: 4, border: "1px solid #cbd5e1", background: (!selectionRange && !selectedCell) ? "#fafafa" : "#fff", color: (!selectionRange && !selectedCell) ? "#cbd5e1" : "#1e293b", whiteSpace: "nowrap" }}>
                    <BtnContent icon={Check} iconColor={(!selectionRange && !selectedCell) ? "#cbd5e1" : "#16a34a"}>Marcar Celdas</BtnContent>
                  </button>
                  <div style={{ width: 1, height: 20, background: "#e2e8f0", margin: "0 4px" }} />
                  <button onClick={() => { if (selectedRows.size === 0) return; setConfirmDelete(true); }}
                    disabled={selectedRows.size === 0}
                    style={{ padding: "7px 10px", fontSize: 12, cursor: selectedRows.size === 0 ? "default" : "pointer", borderRadius: 4, border: "1px solid #fca5a5", background: selectedRows.size === 0 ? "#fafafa" : "#fff", color: selectedRows.size === 0 ? "#cbd5e1" : "#dc2626", whiteSpace: "nowrap" }}>
                    <BtnContent icon={Trash2} iconColor={selectedRows.size === 0 ? "#cbd5e1" : "#dc2626"}>Borrar seleccionadas{selectedRows.size > 0 ? ` (${selectedRows.size})` : ""}</BtnContent>
                  </button>
                  <div style={{ width: 1, height: 20, background: "#e2e8f0", margin: "0 4px" }} />
                  <button onClick={() => {
                      // Si hay filas seleccionadas, pedir confirmación para borrar todas sus celdas
                      if (selectedRows.size > 0) {
                        setConfirmClearRows(true);
                        return;
                      }
                      // Si hay rango de celdas, borrar directamente sin confirmar
                      if (!selectionRange) return;
                      const minRow = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
                      const maxRow = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
                      const minCol = Math.min(selectionRange.startColIdx, selectionRange.endColIdx);
                      const maxCol = Math.max(selectionRange.startColIdx, selectionRange.endColIdx);
                      setRows(rs => rs.map((row, idx) => {
                        if (idx < minRow || idx > maxRow) return row;
                        const updates = {};
                        for (let c = minCol; c <= maxCol; c++) {
                          const col = COLUMNS[c];
                          if (!col || col.type === "calc") continue;
                          updates[col.key] = col.type === "number" ? 0 : "";
                        }
                        return { ...row, ...updates };
                      }));
                    }}
                    disabled={!selectionRange && selectedRows.size === 0}
                    title={selectedRows.size > 0 ? "Borra el contenido de todas las celdas de las filas seleccionadas" : "Borra el contenido de las celdas seleccionadas (no la fila)"}
                    style={{ padding: "7px 10px", fontSize: 12, cursor: (!selectionRange && selectedRows.size === 0) ? "default" : "pointer", borderRadius: 4, border: "1px solid #cbd5e1", background: (!selectionRange && selectedRows.size === 0) ? "#fafafa" : "#fff", color: (!selectionRange && selectedRows.size === 0) ? "#cbd5e1" : "#1e293b", whiteSpace: "nowrap" }}>
                    <BtnContent icon={Eraser} iconColor={(!selectionRange && selectedRows.size === 0) ? "#cbd5e1" : "#475569"}>Borrar contenido celdas{selectedRows.size > 0 ? ` (${selectedRows.size} filas)` : ""}</BtnContent>
                  </button>
                  <div style={{ width: 1, height: 20, background: "#e2e8f0", margin: "0 4px" }} />
                  <button onClick={() => {
                      if (selectedRows.size === 0) {
                        setStatus("Marca con checkbox las filas que quieres copiar", "error");
                        return;
                      }
                      // Conservar el orden actual del grid
                      const copia = rows.filter(r => selectedRows.has(r.id)).map(r => {
                        const { id, ...rest } = r;
                        return { ...rest };
                      });
                      setFilasCopiadas(copia);
                      setStatus(`${copia.length} fila${copia.length !== 1 ? "s" : ""} copiada${copia.length !== 1 ? "s" : ""} al portapapeles interno`, "success");
                    }}
                    disabled={selectedRows.size === 0}
                    title="Copia las filas marcadas con checkbox al portapapeles interno"
                    style={{ padding: "7px 10px", fontSize: 12, cursor: selectedRows.size === 0 ? "default" : "pointer", borderRadius: 4, border: "1px solid #cbd5e1", background: selectedRows.size === 0 ? "#fafafa" : "#fff", color: selectedRows.size === 0 ? "#94a3b8" : "#1e293b", whiteSpace: "nowrap" }}>
                    <BtnContent icon={Copy} iconColor={selectedRows.size === 0 ? "#cbd5e1" : "#0369a1"}>Copiar filas{selectedRows.size > 0 ? ` (${selectedRows.size})` : ""}</BtnContent>
                  </button>
                  <button onClick={() => {
                      if (filasCopiadas.length === 0) {
                        setStatus("No hay filas copiadas. Marca filas con checkbox y pulsa 'Copiar filas'", "error");
                        return;
                      }
                      if (!selectedCell) {
                        setStatus("Selecciona una celda para indicar dónde pegar", "error");
                        return;
                      }
                      const idxDestino = rows.findIndex(r => r.id === selectedCell.rowId);
                      if (idxDestino === -1) {
                        setStatus("La celda seleccionada no tiene fila válida", "error");
                        return;
                      }
                      // Considera vacía una fila si no tiene referencia, nombre, descripción ni naturaleza
                      const esFilaVacia = (row) =>
                        !String(row.referencia || "").trim() &&
                        !String(row.nombre || "").trim() &&
                        !String(row.descripcion || "").trim() &&
                        !String(row.naturaleza || "").trim();

                      setRows(rs => {
                        const n = filasCopiadas.length;
                        const res = [...rs];
                        // Asignar ids únicos a las filas que se van a pegar
                        let nextLocalId = Math.max(...rs.map(x => x.id), 0) + 1;
                        const aPegar = filasCopiadas.map(f => ({ id: nextLocalId++, ...f }));
                        nextId.current = nextLocalId;

                        // Recorremos las n filas siguientes a partir del índice destino
                        // Si todas las que necesitamos están vacías → sobreescribimos.
                        // Si alguna está ocupada → insertamos n filas vacías nuevas y luego pegamos
                        const finDestino = idxDestino + n - 1;
                        let todasVacias = true;
                        for (let i = idxDestino; i <= finDestino; i++) {
                          if (i >= res.length || !esFilaVacia(res[i])) { todasVacias = false; break; }
                        }
                        if (todasVacias) {
                          for (let i = 0; i < n; i++) {
                            res[idxDestino + i] = aPegar[i];
                          }
                          return res;
                        }
                        // Insertar las n filas a partir del índice destino (empujando el resto hacia abajo)
                        res.splice(idxDestino, 0, ...aPegar);
                        return res;
                      });
                      setStatus(`${filasCopiadas.length} fila${filasCopiadas.length !== 1 ? "s" : ""} pegada${filasCopiadas.length !== 1 ? "s" : ""} a partir de la fila ${idxDestino + 1}`, "success");
                    }}
                    disabled={filasCopiadas.length === 0 || !selectedCell}
                    title="Pega las filas copiadas a partir de la fila de la celda seleccionada"
                    style={{ padding: "7px 10px", fontSize: 12, cursor: (filasCopiadas.length === 0 || !selectedCell) ? "default" : "pointer", borderRadius: 4, border: "1px solid #cbd5e1", background: (filasCopiadas.length === 0 || !selectedCell) ? "#fafafa" : "#fff", color: (filasCopiadas.length === 0 || !selectedCell) ? "#94a3b8" : "#1e293b", whiteSpace: "nowrap" }}>
                    <BtnContent icon={ClipboardCheck} iconColor={(filasCopiadas.length === 0 || !selectedCell) ? "#cbd5e1" : "#16a34a"}>Pegar filas{filasCopiadas.length > 0 ? ` (${filasCopiadas.length})` : ""}</BtnContent>
                  </button>
                  <div style={{ width: 1, height: 20, background: "#e2e8f0", margin: "0 4px" }} />
                  <button onClick={() => {
                      const ok = guardarPresupuestoLocal(presupuesto, rows);
                      if (ok) {
                        setStatus(`Presupuesto guardado en el navegador (${rows.length} filas)`, "success");
                      } else {
                        setStatus("Error al guardar el presupuesto en el navegador", "error");
                      }
                    }}
                    title="Guarda el presupuesto actual en el almacenamiento local del navegador para recuperarlo la próxima vez que abras la aplicación"
                    style={{ padding: "7px 10px", fontSize: 12, cursor: "pointer", borderRadius: 4, border: "1px solid #cbd5e1", background: "#fff", color: "#1e293b", whiteSpace: "nowrap" }}>
                    <BtnContent icon={Save} iconColor="#16a34a">Guardar local</BtnContent>
                  </button>
                </div>
      </div>
      {/* Línea de información */}
      <div style={{ background: "#f5f5f5", color: "#525252", padding: "4px 16px", display: "flex", gap: 24, fontSize: 11, flexShrink: 0, borderTop: "1px solid #e5e5e5" }}>
        <span>Filas: <strong style={{ color: "#171717" }}>{rows.length}</strong></span>
        <span>Marcadas: <strong style={{ color: "#171717" }}>{selectedRows.size}</strong></span>
        {resumenSeleccion.celdas > 0 && (
          <span>Celdas sel.: <strong style={{ color: "#171717" }}>{resumenSeleccion.celdas}</strong></span>
        )}
        {resumenSeleccion.numericas > 0 && (
          <span>Suma: <strong style={{ color: "#171717" }}>{fmt(resumenSeleccion.suma)}</strong></span>
        )}
        {selectedCell && <span>Celda: <strong style={{ color: "#171717" }}>{selectedCell.colKey} / fila {rows.findIndex(r => r.id === selectedCell.rowId) + 1}</strong></span>}
        <span style={{ marginLeft: "auto", fontSize: 11, color: "#737373" }}>Doble click para editar</span>
      </div>
      {/* Barra de estado */}
      <div style={{
        background: statusMessage.type === "error" ? "#7f1d1d" : statusMessage.type === "success" ? "#14532d" : statusMessage.type === "working" ? "#1e40af" : "#0f172a",
        color: "#fff", padding: "5px 16px", fontSize: 12, flexShrink: 0,
        display: "flex", alignItems: "center", gap: 8, borderTop: "1px solid rgba(255,255,255,0.1)",
        transition: "background 0.2s"
      }}>
        <Icon as={
          statusMessage.type === "error" ? X
          : statusMessage.type === "success" ? Check
          : statusMessage.type === "working" ? RefreshCw
          : Check
        } size={13} color="#fff" />
        <span>{statusMessage.text}</span>
        {statusMessage.timestamp && (
          <span style={{ marginLeft: "auto", color: "rgba(255,255,255,0.7)", fontSize: 11 }}>
            {new Date(statusMessage.timestamp).toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
          </span>
        )}
        {/* Indicador de estado de la API local */}
        <span title={apiLocalViva === true ? "API local conectada" : apiLocalViva === false ? "API local no disponible" : "Comprobando API local..."}
          style={{ marginLeft: statusMessage.timestamp ? 12 : "auto", display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11, color: "rgba(255,255,255,0.85)", whiteSpace: "nowrap" }}>
          <span style={{ width: 8, height: 8, borderRadius: "50%", background: apiLocalViva === true ? "#22c55e" : apiLocalViva === false ? "#ef4444" : "#facc15", display: "inline-block" }} />
          API local
        </span>
      </div>
      {modal && <ModalComponent action={modal.action} label={modal.label} onClose={() => setModal(null)} />}
      {showFijarPrecio && (() => {
        // Filas afectadas según el modo
        let lineasAfectadas;
        let titulo, descripcion;
        if (showFijarPrecio === "total") {
          lineasAfectadas = rows.filter(r => ["PD","PE","E"].includes(r.naturaleza));
          titulo = "Fijar el precio total del presupuesto";
          descripcion = "Indica el importe neto total objetivo. Se ajustará el descuento de todas las líneas de producto/elemento proporcionalmente a su neto actual.";
        } else {
          // celdas: usar selectedRows o selectionRange
          let idsAfectadas = new Set();
          if (selectedRows.size > 0) {
            idsAfectadas = selectedRows;
          } else if (selectionRange) {
            const minR = Math.min(selectionRange.startRowIdx, selectionRange.endRowIdx);
            const maxR = Math.max(selectionRange.startRowIdx, selectionRange.endRowIdx);
            for (let i = minR; i <= maxR; i++) {
              if (rows[i]) idsAfectadas.add(rows[i].id);
            }
          }
          lineasAfectadas = rows.filter(r => idsAfectadas.has(r.id) && ["PD","PE","E"].includes(r.naturaleza));
          titulo = "Fijar el precio de las celdas seleccionadas";
          descripcion = "Indica el importe neto objetivo para las filas seleccionadas. Se ajustará el descuento de todas ellas proporcionalmente a su neto actual. (Solo afecta a líneas de producto/elemento)";
        }
        return (
          <FijarPrecioDialog
            titulo={titulo}
            descripcion={descripcion}
            lineasAfectadas={lineasAfectadas}
            onClose={() => setShowFijarPrecio(null)}
            onAplicar={(nuevosDescuentos) => {
              setRows(rs => rs.map(r => (
                nuevosDescuentos.hasOwnProperty(r.id) ? { ...r, dtoaplicado: nuevosDescuentos[r.id] } : r
              )));
            }}
          />
        );
      })()}
      {showAsistente && (
        <AsistenteReferenciasDialog
          setStatus={setStatus}
          onClose={() => setShowAsistente(false)}
          onInsertar={(prod) => {
            const datosProd = {
              representacion: "", naturaleza: "PD", posicion: "", cantidad: 1,
              referencia: prod.referencia || "", nombre: prod.nombre || "",
              pvp: Number(prod.pvp) || 0, dtoaplicado: Number(prod.dtoreferenciagrupo) || 0,
              descripcion: prod.descripcion || "", familia: prod.familia || "",
              subfamilia: prod.subfamilia || "", descripcionsubfamilia: prod.descripcionsubfamilia || "",
              preciocosteunitario: Number(prod.preciocoste) || 0, idposicion: prod.id || "",
              imagen: prod.imagen || "", precionetounitario2: 0,
              grupodescuento: prod.grupodescuento || "", descripciongrupodescuento: prod.descripciongrupodescuento || "",
              fechapvp: prod.fechapvp || null, fechapreciocoste: prod.fechapreciocoste || null,
              pvpVencido: tieneMasDeUnAno(prod.fechapvp),
            };
            const esFilaVacia = (row) =>
              !String(row.referencia || "").trim() && !String(row.nombre || "").trim() &&
              !String(row.descripcion || "").trim() && !String(row.naturaleza || "").trim();
            setRows(r => {
              const idx = selectedCell ? r.findIndex(row => row.id === selectedCell.rowId) : -1;
              if (idx === -1) {
                let nextLocalId = Math.max(...r.map(x => x.id), 0) + 1;
                nextId.current = nextLocalId + 1;
                return [...r, { id: nextLocalId, ...datosProd }];
              }
              const filaActual = r[idx];
              if (esFilaVacia(filaActual)) {
                const res = [...r]; res[idx] = { ...filaActual, ...datosProd }; return res;
              }
              let nextLocalId = Math.max(...r.map(x => x.id), 0) + 1;
              nextId.current = nextLocalId + 1;
              const res = [...r]; res.splice(idx, 0, { id: nextLocalId, ...datosProd }); return res;
            });
          }}
        />
      )}
      {confirmBorrarPresup && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", minWidth: 380, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
              <Icon as={X} size={22} color="#dc2626" />
              <h2 style={{ fontSize: 15, fontWeight: 600, margin: 0 }}>Borrar presupuesto actual</h2>
            </div>
            <p style={{ fontSize: 13, color: "#475569", margin: "0 0 20px", lineHeight: 1.6 }}>
              Se va a borrar el presupuesto actual (cabecera y todas las filas) y se preparará uno nuevo en blanco con el siguiente número disponible.
              <br /><span style={{ color: "#dc2626", fontSize: 12 }}>Esta acción no se puede deshacer.</span>
            </p>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setConfirmBorrarPresup(false)}
                style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>
                <BtnContent icon={X}>Cancelar</BtnContent>
              </button>
              <button onClick={async () => {
                  // Borrar también el guardado local
                  borrarPresupuestoLocal();
                  // Reset filas
                  setRows(filasVacias());
                  setSelectedRows(new Set());
                  setSelectionRange(null);
                  setSelectedCell(null);
                  nextId.current = 11;
                  // Reset cabecera y obtener siguiente número
                  const anoActual = getAnoFiscal();
                  setPresupuesto({ id: "", np: "", numerocompleto: "", anopresupuesto: anoActual, revision: "0", titulo: "", cliente: "", idcliente: null, alaatencion: null, alaatencion_nombre: "", alaatencion_cliente: "" });
                  setConfirmBorrarPresup(false);
                  setStatus("Presupuesto borrado, obteniendo siguiente número...", "working");
                  try {
                    const sig = await obtenerSiguienteNumero(anoActual);
                    setPresupuesto(p => {
                      const newP = { ...p, np: String(sig), anopresupuesto: anoActual };
                      newP.numerocompleto = buildNumeroCompleto(codigoUsuario, newP.np, newP.anopresupuesto);
                      return newP;
                    });
                    setStatus(`Presupuesto en blanco. ${buildNumeroCompleto(codigoUsuario, sig, anoActual)}`, "success");
                  } catch (e) {
                    setStatus("Presupuesto borrado, pero no se ha podido obtener el siguiente número: " + e.message, "error");
                  }
                }}
                style={{ padding: "7px 16px", borderRadius: 6, border: "none", background: "#dc2626", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 500 }}>
                <BtnContent icon={Check} iconColor="#fff">Sí, borrar</BtnContent>
              </button>
            </div>
          </div>
        </div>
      )}
      {comprobarDialog && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", minWidth: 480, maxWidth: 720, maxHeight: "85vh", display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, borderBottom: "1px solid #e2e8f0", paddingBottom: 10 }}>
              <Icon as={ClipboardCheck} size={22} color="#1e3a5f" />
              <h2 style={{ fontSize: 15, fontWeight: 600, margin: 0 }}>Comprobación con la base de datos</h2>
            </div>
            {comprobarDialog.tipo === "no_existe" && (
              <p style={{ fontSize: 13, color: "#475569", margin: 0, lineHeight: 1.6 }}>
                <strong>El presupuesto Nº {presupuesto.np} Rev.{presupuesto.revision}</strong> no existe todavía en la base de datos. Si pulsas <em>Guardar Presupuesto</em> se creará por primera vez.
              </p>
            )}
            {comprobarDialog.tipo === "iguales" && (
              <div>
                <div style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 6, padding: "12px 14px", fontSize: 13, color: "#14532d", lineHeight: 1.6 }}>
                  <Icon as={Check} size={18} color="#16a34a" />&nbsp; <strong>El presupuesto actual es idéntico al guardado en BD.</strong>
                  <br /><span style={{ color: "#15803d", fontSize: 12 }}>No hay cambios pendientes.</span>
                </div>
              </div>
            )}
            {comprobarDialog.tipo === "diferentes" && (
              <div style={{ overflow: "auto", flex: 1 }}>
                <div style={{ background: "#fef3c7", border: "1px solid #fcd34d", borderRadius: 6, padding: "10px 14px", marginBottom: 12, fontSize: 12, color: "#92400e" }}>
                  <strong>{comprobarDialog.diffs.length} diferencia{comprobarDialog.diffs.length > 1 ? "s" : ""}</strong> entre el presupuesto actual y el guardado en BD:
                </div>
                <ul style={{ margin: 0, paddingLeft: 18, fontSize: 12, color: "#475569", maxHeight: 350, overflow: "auto" }}>
                  {comprobarDialog.diffs.map((d, i) => (
                    <li key={i} style={{ padding: "3px 0", borderBottom: "1px dotted #f1f5f9", fontFamily: "monospace" }}>{d}</li>
                  ))}
                </ul>
              </div>
            )}
            <div style={{ display: "flex", justifyContent: "flex-end", paddingTop: 14, marginTop: 12, borderTop: "1px solid #e2e8f0" }}>
              <button onClick={() => setComprobarDialog(null)}
                style={{ padding: "7px 20px", borderRadius: 6, border: "none", background: "#2563eb", color: "#fff", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                <BtnContent icon={X} iconColor="#fff">Cerrar</BtnContent>
              </button>
            </div>
          </div>
        </div>
      )}
      {confirmSobreescribir && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", minWidth: 460, maxWidth: 640, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
              <Icon as={HelpCircle} size={22} color="#d97706" />
              <h3 style={{ fontSize: 15, fontWeight: 700, color: "#1e3a5f", margin: 0 }}>Sobrescribir presupuesto existente</h3>
            </div>
            <p style={{ fontSize: 13, color: "#475569", lineHeight: 1.6, margin: "0 0 12px" }}>
              Ya existe un presupuesto <strong>{presupuesto.numerocompleto || presupuesto.np}</strong> Rev.<strong>{presupuesto.revision}</strong> en la base de datos (ID {confirmSobreescribir.idBd}, {confirmSobreescribir.lineasBd} líneas).
            </p>
            <p style={{ fontSize: 13, color: "#475569", lineHeight: 1.6, margin: "0 0 12px" }}>
              Se han detectado <strong>{confirmSobreescribir.diffs.length} diferencia{confirmSobreescribir.diffs.length > 1 ? "s" : ""}</strong> con el presupuesto actual.
            </p>
            <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#991b1b", marginBottom: 16 }}>
              <strong>Atención:</strong> al sobrescribir, los datos guardados en BD se sustituirán por los del presupuesto actual y no se podrán recuperar.
            </div>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setConfirmSobreescribir(false)} disabled={guardandoPresup}
                style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}>
                <BtnContent icon={X}>Cancelar</BtnContent>
              </button>
              <button onClick={async () => {
                  setGuardandoPresup(true);
                  setStatus("Sobrescribiendo presupuesto en BD...", "working");
                  try {
                    const r = await guardarPresupuestoEnBD(confirmSobreescribir.idBd);
                    setPresupuesto(p => ({ ...p, id: String(confirmSobreescribir.idBd) }));
                    setStatus(`Presupuesto actualizado en BD (${r.lineas_actualizadas} líneas)`, "success");
                    setConfirmSobreescribir(false);
                  } catch (e) {
                    setStatus("Error sobrescribiendo: " + e.message, "error");
                  } finally {
                    setGuardandoPresup(false);
                  }
                }} disabled={guardandoPresup}
                style={{ padding: "7px 16px", borderRadius: 6, border: "none", background: "#dc2626", color: "#fff", cursor: guardandoPresup ? "default" : "pointer", fontSize: 13, fontWeight: 600 }}>
                <BtnContent icon={guardandoPresup ? RefreshCw : Save} iconColor="#fff">{guardandoPresup ? "Guardando..." : "Sí, sobrescribir"}</BtnContent>
              </button>
            </div>
          </div>
        </div>
      )}
      {showGuardarElem && (
        <GuardarElementoDialog
          filasSeleccionadas={rows.filter(r => selectedRows.has(r.id))}
          setStatus={setStatus}
          onClose={() => setShowGuardarElem(false)}
        />
      )}
      {actualizarProductos && (
        <ActualizarProductosDialog
          datos={actualizarProductos}
          setStatus={setStatus}
          onClose={() => setActualizarProductos(null)}
        />
      )}
      {showLeerProducto && (
        <LeerProductoDialog
          setStatus={setStatus}
          onClose={() => setShowLeerProducto(false)}
          onInsertar={(prod) => {
            // Construye los datos del producto (sin id, para usar al rellenar o insertar)
            const datosProd = {
              representacion: "",
              naturaleza: "PD",
              posicion: "",
              cantidad: 1,
              referencia: prod.referencia || "",
              nombre: prod.nombre || "",
              pvp: Number(prod.pvp) || 0,
              dtoaplicado: Number(prod.dtoreferenciagrupo) || 0,
              descripcion: prod.descripcion || "",
              familia: prod.familia || "",
              subfamilia: prod.subfamilia || "",
              descripcionsubfamilia: prod.descripcionsubfamilia || "",
              preciocosteunitario: Number(prod.preciocoste) || 0,
              idposicion: prod.id || "",
              imagen: prod.imagen || "",
              precionetounitario2: 0,
              grupodescuento: prod.grupodescuento || "",
              descripciongrupodescuento: prod.descripciongrupodescuento || "",
              fechapvp: prod.fechapvp || null,
              fechapreciocoste: prod.fechapreciocoste || null,
              pvpVencido: tieneMasDeUnAno(prod.fechapvp),
            };
            // Considera vacía una fila si no tiene referencia, ni producto, ni descripción, ni naturaleza
            const esFilaVacia = (row) =>
              !String(row.referencia || "").trim() &&
              !String(row.nombre || "").trim() &&
              !String(row.descripcion || "").trim() &&
              !String(row.naturaleza || "").trim();

            setRows(r => {
              // Buscar la fila de la celda seleccionada (rectángulo azul)
              const idx = selectedCell ? r.findIndex(row => row.id === selectedCell.rowId) : -1;
              if (idx === -1) {
                // Sin celda seleccionada → añadir al final como nueva fila
                let nextLocalId = Math.max(...r.map(x => x.id), 0) + 1;
                nextId.current = nextLocalId + 1;
                return [...r, { id: nextLocalId, ...datosProd }];
              }
              const filaActual = r[idx];
              if (esFilaVacia(filaActual)) {
                // Rellenar la fila vacía manteniendo su id
                const res = [...r];
                res[idx] = { ...filaActual, ...datosProd };
                return res;
              }
              // Insertar nueva fila ENCIMA de la fila actual
              let nextLocalId = Math.max(...r.map(x => x.id), 0) + 1;
              nextId.current = nextLocalId + 1;
              const res = [...r];
              res.splice(idx, 0, { id: nextLocalId, ...datosProd });
              return res;
            });
          }}
        />
      )}
      {showLeerElem && (
        <LeerElementoDialog
          setStatus={setStatus}
          onClose={() => setShowLeerElem(false)}
          onInsertar={(data) => {
            // data tiene { cabecera, productos }
            // Construir fila comentario + filas producto
            let nextLocalId = Math.max(...rows.map(x => x.id), 0) + 1;
            const filaComentario = {
              id: nextLocalId++,
              representacion: "",
              naturaleza: "CM",
              posicion: "",
              cantidad: 0,
              referencia: "",
              nombre: data.cabecera.nombre || "",
              pvp: 0, dtoaplicado: 0,
              descripcion: data.cabecera.descripcion || "",
              familia: "", subfamilia: "",
              preciocosteunitario: 0, idposicion: "", imagen: "",
              precionetounitario2: 0, grupodescuento: "",
            };
            console.log("[LeerElemento] productos recibidos del backend:", data.productos);
            const filasProducto = (data.productos || []).map(p => {
              // Si la línea es un comentario, crear fila CM con el texto en "nombre"
              // Aceptamos varias formas por si el backend devuelve el campo con nombres ligeramente distintos
              const esComentario = p.es_comentario === true || p.escomentario === true || p.eselemento === false || Number(p.idproducto) === 118852;
              if (esComentario) {
                return {
                  id: nextLocalId++,
                  representacion: "",
                  naturaleza: "CM",
                  posicion: "",
                  cantidad: 0,
                  referencia: "",
                  nombre: p.texto || "",
                  pvp: 0, dtoaplicado: 0,
                  descripcion: "",
                  familia: "", subfamilia: "",
                  descripcionsubfamilia: "",
                  preciocosteunitario: 0, idposicion: "", imagen: "",
                  precionetounitario2: 0, grupodescuento: "",
                  descripciongrupodescuento: "",
                };
              }
              // Línea de producto normal
              return {
                id: nextLocalId++,
                representacion: "",
                naturaleza: "PE",
                posicion: "",
                cantidad: parseInt(p.cantidad, 10) || 1,
                referencia: p.referencia || "",
                nombre: p.nombre || "",
                pvp: Number(p.pvp) || 0,
                dtoaplicado: Number(p.dtoreferenciagrupo) || 0,
                descripcion: p.descripcion || "",
                familia: p.familia || "",
                subfamilia: p.subfamilia || "",
                descripcionsubfamilia: p.descripcionsubfamilia || "",
                preciocosteunitario: Number(p.preciocoste) || 0,
                idposicion: p.idproducto || "",
                imagen: "",
                precionetounitario2: 0,
                grupodescuento: p.grupodescuento || "",
                descripciongrupodescuento: p.descripciongrupodescuento || "",
              };
            });
            const nuevas = [filaComentario, ...filasProducto];
            setRows(r => {
              if (selectedRows.size === 0) return [...r, ...nuevas];
              const idx = r.findIndex(row => selectedRows.has(row.id));
              if (idx === -1) return [...r, ...nuevas];
              const res = [...r];
              res.splice(idx, 0, ...nuevas);
              return res;
            });
            nextId.current = nextLocalId;
          }}
        />
      )}
      {showBorrarVacias && (
        <BorrarFilasVaciasDialog
          lineasSeleccionadas={rows.filter(r => selectedRows.has(r.id))}
          onClose={() => setShowBorrarVacias(false)}
          onAceptar={(columna) => {
            setRows(rs => rs.filter(row => {
              if (!selectedRows.has(row.id)) return true; // no seleccionada → conservar
              const val = row[columna];
              const esVacia = val === null || val === undefined || val === "" || val === 0;
              return !esVacia; // si está vacía la quitamos
            }));
            setSelectedRows(new Set());
          }}
        />
      )}
      {showClientes && (
        <ClientesDialog
          setStatus={setStatus}
          onClose={() => setShowClientes(false)}
          onAsignarPresupuesto={(cliente) => {
            setPresupuesto(p => ({
              ...p,
              cliente: cliente.nombrecomun || cliente.razonsocial || "",
              idcliente: cliente.id,
            }));
          }}
        />
      )}
      {showContactos && (
        <ContactosDialog
          setStatus={setStatus}
          onClose={() => setShowContactos(false)}
        />
      )}
      {juntarDupConfirm && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100000 }} onClick={() => setJuntarDupConfirm(null)}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", width: "95%", maxWidth: 520, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
            <h3 style={{ margin: 0, marginBottom: 12, fontSize: 15, fontWeight: 700, color: "#171717", display: "inline-flex", alignItems: "center", gap: 8 }}>
              <Icon as={Layers} size={18} color="#1e3a5f" /> Juntar productos duplicados
            </h3>
            <p style={{ fontSize: 13, color: "#525252", marginBottom: 8, lineHeight: 1.5 }}>
              Se han encontrado <strong style={{ color: "#171717" }}>{juntarDupConfirm.totalGrupos} referencia{juntarDupConfirm.totalGrupos !== 1 ? "s" : ""} duplicada{juntarDupConfirm.totalGrupos !== 1 ? "s" : ""}</strong> en el rango seleccionado, afectando a <strong style={{ color: "#171717" }}>{juntarDupConfirm.totalFilasAfectadas} fila{juntarDupConfirm.totalFilasAfectadas !== 1 ? "s" : ""}</strong>.
            </p>
            <div style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 6, padding: "8px 12px", fontSize: 12, color: "#475569", marginBottom: 12, maxHeight: 140, overflowY: "auto" }}>
              {juntarDupConfirm.duplicados.map(([ref, idxs]) => {
                const totalCant = idxs.reduce((s, i) => s + (Number(rows[i]?.cantidad) || 0), 0);
                return (
                  <div key={ref} style={{ padding: "2px 0", fontFamily: "monospace", display: "flex", justifyContent: "space-between" }}>
                    <span>{ref}</span>
                    <span style={{ color: "#0369a1" }}>{idxs.length} × → cant. total {totalCant}</span>
                  </div>
                );
              })}
            </div>
            <p style={{ fontSize: 12, color: "#525252", marginBottom: 14 }}>
              La cantidad total se sumará en la <strong>primera fila</strong> de cada grupo. ¿Qué hacemos con las filas duplicadas restantes?
            </p>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", flexWrap: "wrap" }}>
              <button onClick={() => setJuntarDupConfirm(null)}
                style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #d4d4d4", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
                <BtnContent icon={X}>Cancelar</BtnContent>
              </button>
              <button onClick={() => {
                  // Dejar a 0 las duplicadas
                  setRows(prev => {
                    const next = [...prev];
                    juntarDupConfirm.duplicados.forEach(([_, idxs]) => {
                      const totalCant = idxs.reduce((s, i) => s + (Number(next[i]?.cantidad) || 0), 0);
                      // La primera fila recibe el total
                      next[idxs[0]] = { ...next[idxs[0]], cantidad: totalCant };
                      // Las demás a 0
                      for (let k = 1; k < idxs.length; k++) {
                        next[idxs[k]] = { ...next[idxs[k]], cantidad: 0 };
                      }
                    });
                    return next;
                  });
                  setStatus(`Juntados ${juntarDupConfirm.totalGrupos} grupo${juntarDupConfirm.totalGrupos !== 1 ? "s" : ""} duplicado${juntarDupConfirm.totalGrupos !== 1 ? "s" : ""}, ${juntarDupConfirm.totalFilasAfectadas - juntarDupConfirm.totalGrupos} fila${(juntarDupConfirm.totalFilasAfectadas - juntarDupConfirm.totalGrupos) !== 1 ? "s" : ""} a cantidad 0`, "success");
                  setJuntarDupConfirm(null);
                }}
                style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#171717", cursor: "pointer", fontSize: 12 }}>
                <BtnContent icon={Check} iconColor="#475569">Dejar a 0</BtnContent>
              </button>
              <button onClick={() => {
                  // Borrar las duplicadas
                  setRows(prev => {
                    const idsAEliminar = new Set();
                    const next = [...prev];
                    juntarDupConfirm.duplicados.forEach(([_, idxs]) => {
                      const totalCant = idxs.reduce((s, i) => s + (Number(next[i]?.cantidad) || 0), 0);
                      next[idxs[0]] = { ...next[idxs[0]], cantidad: totalCant };
                      for (let k = 1; k < idxs.length; k++) {
                        idsAEliminar.add(next[idxs[k]].id);
                      }
                    });
                    return next.filter(r => !idsAEliminar.has(r.id));
                  });
                  setStatus(`Juntados ${juntarDupConfirm.totalGrupos} grupo${juntarDupConfirm.totalGrupos !== 1 ? "s" : ""} duplicado${juntarDupConfirm.totalGrupos !== 1 ? "s" : ""}, ${juntarDupConfirm.totalFilasAfectadas - juntarDupConfirm.totalGrupos} fila${(juntarDupConfirm.totalFilasAfectadas - juntarDupConfirm.totalGrupos) !== 1 ? "s" : ""} eliminada${(juntarDupConfirm.totalFilasAfectadas - juntarDupConfirm.totalGrupos) !== 1 ? "s" : ""}`, "success");
                  setJuntarDupConfirm(null);
                }}
                style={{ padding: "7px 14px", borderRadius: 6, border: "1px solid #fca5a5", background: "#fef2f2", color: "#991b1b", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                <BtnContent icon={Trash2} iconColor="#dc2626">Borrar duplicadas</BtnContent>
              </button>
            </div>
          </div>
        </div>
      )}
      {leerPmdDialog && (
        <LeerPreciosPMDDialog
          dialog={leerPmdDialog}
          onClose={() => setLeerPmdDialog(null)}
          setStatus={setStatus}
          aplicarAFila={(rowIdx, datos) => {
            setRows(rs => {
              if (rowIdx < 0 || rowIdx >= rs.length) return rs;
              const fila = rs[rowIdx];
              const next = [...rs];
              const upd = { ...fila };
              if (datos.referencia) upd.referencia = String(datos.referencia);
              if (datos.descripcion != null) upd.descripcion = String(datos.descripcion);
              if (datos.pvp != null) {
                const n = parseFloat(String(datos.pvp).replace(",", "."));
                if (!isNaN(n)) upd.pvp = n;
              }
              if (datos.preciocoste != null) {
                const n = parseFloat(String(datos.preciocoste).replace(",", "."));
                if (!isNaN(n)) upd.preciocosteunitario = n;
              }
              if (datos.grupodescuento != null) {
                upd.grupodescuento = String(datos.grupodescuento);
              }
              next[rowIdx] = upd;
              return next;
            });
          }}
        />
      )}
      {showLoginDialog && (
        <LoginDialog
          onClose={() => setShowLoginDialog(false)}
          onLoginOk={(userData) => {
            setUsuarioActual(userData.usuario || "");
            if (userData.codigopresupuestos) setCodigoUsuario(userData.codigopresupuestos);
            setShowLoginDialog(false);
            setStatus(`Sesión iniciada como ${userData.usuario}`, "success");
          }}
        />
      )}
      {siemensDialog && (
        <SiemensRefsDialog
          celdasTrabajadas={siemensDialog.celdasTrabajadas}
          setStatus={setStatus}
          onClose={() => setSiemensDialog(null)}
          onComplete={(aceptadas) => {
            // Aplicar cambios: la primera ref de cada celda sobreescribe el contenido, las siguientes crean nuevas filas
            const porCelda = {};
            aceptadas.forEach(a => {
              const k = `${a.rowId}:${a.colKey}`;
              if (!porCelda[k]) porCelda[k] = [];
              porCelda[k].push(a);
            });
            setRows(prev => {
              const nuevasFilas = [];
              const newRows = prev.map(row => {
                let changed = { ...row };
                Object.keys(porCelda).forEach(k => {
                  const [rowIdStr, colKey] = k.split(":");
                  if (String(row.id) !== rowIdStr) return;
                  const lista = porCelda[k];
                  if (lista.length === 0) return;
                  // La primera sobreescribe la celda
                  changed[colKey] = lista[0].ref;
                  // Si la celda no es referencia, mover la ref a la columna referencia
                  if (colKey !== "referencia") {
                    changed[colKey] = lista[0].ref;
                    // Y opcionalmente también copiarla al campo referencia si está vacía
                    if (!changed.referencia) changed.referencia = lista[0].ref;
                  }
                  // Las siguientes generan filas nuevas por debajo
                  for (let i = 1; i < lista.length; i++) {
                    nuevasFilas.push({
                      pivoteRowId: row.id,
                      fila: {
                        id: nextId.current++,
                        representacion: "",
                        naturaleza: "PD",
                        posicion: "",
                        cantidad: 1,
                        referencia: lista[i].ref,
                        nombre: "",
                        pvp: 0, dtoaplicado: 0,
                        precionetounitario: 0, precionetoposicion: 0,
                        descripcion: "", familia: "", subfamilia: "",
                        preciocosteunitario: 0, idposicion: "", imagen: "",
                        precionetounitario2: 0, grupodescuento: "",
                      },
                    });
                  }
                });
                return changed;
              });
              // Insertar las nuevas filas justo después de su pivote
              if (nuevasFilas.length === 0) return newRows;
              const finales = [];
              newRows.forEach(r => {
                finales.push(r);
                const nuevas = nuevasFilas.filter(n => n.pivoteRowId === r.id);
                nuevas.forEach(n => finales.push(n.fila));
              });
              return finales;
            });
            const numAceptadas = aceptadas.length;
            setStatus(`${numAceptadas} referencia${numAceptadas !== 1 ? "s" : ""} SIEMENS aplicada${numAceptadas !== 1 ? "s" : ""} al presupuesto`, "success");
            setSiemensDialog(null);
          }}
        />
      )}
      {showSelecContacto && (
        <SelectorContactoDialog
          setStatus={setStatus}
          onClose={() => setShowSelecContacto(false)}
          onSelect={(contacto) => {
            setPresupuesto(p => ({
              ...p,
              alaatencion: contacto.id,
              alaatencion_nombre: contacto.nombre || "",
              alaatencion_cliente: contacto.cliente_nombre || contacto.cliente_razonsocial || "",
            }));
            setShowSelecContacto(false);
          }}
        />
      )}
      {showSelecCliente && (
        <SelectorClienteDialog
          setStatus={setStatus}
          onClose={() => setShowSelecCliente(false)}
          onSelect={(cliente) => {
            const nombre = cliente.nombrecomun || cliente.razonsocial || `Cliente ${cliente.id}`;
            setPresupuesto(p => ({
              ...p,
              idcliente: cliente.id,
              cliente: nombre,
              // Si cambia el cliente, limpiar el contacto previo
              alaatencion: null,
              alaatencion_nombre: "",
              alaatencion_cliente: "",
            }));
            setShowSelecCliente(false);
            setStatus(`Cliente seleccionado: ${nombre}`, "success");
          }}
        />
      )}
      {infoCeldaDialog && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100001 }}
          onClick={() => setInfoCeldaDialog(null)}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", maxWidth: 420, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}
            onClick={e => e.stopPropagation()}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
              <Icon as={Search} size={22} color="#2563eb" />
              <h3 style={{ fontSize: 15, fontWeight: 700, color: "#171717", margin: 0 }}>Comprobar celda</h3>
            </div>
            <div style={{ fontSize: 13, color: "#475569", lineHeight: 1.8 }}>
              <div><strong>Celda:</strong> {infoCeldaDialog.colKey} / fila {infoCeldaDialog.fila}</div>
              <div><strong>Nº de caracteres:</strong> {infoCeldaDialog.caracteres}</div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <strong>Color de fondo:</strong>
                <span style={{ display: "inline-block", width: 16, height: 16, borderRadius: 3, border: "1px solid #cbd5e1", background: infoCeldaDialog.bg }} />
                <span style={{ fontFamily: "monospace" }}>{infoCeldaDialog.bg}</span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <strong>Color de tinta:</strong>
                <span style={{ display: "inline-block", width: 16, height: 16, borderRadius: 3, border: "1px solid #cbd5e1", background: infoCeldaDialog.color }} />
                <span style={{ fontFamily: "monospace" }}>{infoCeldaDialog.color}</span>
              </div>
              {/* Previsualización del contenido con sus colores */}
              <div style={{ marginTop: 10 }}>
                <strong>Vista:</strong>
                <div style={{ marginTop: 4, padding: "6px 10px", borderRadius: 4, border: "1px solid #cbd5e1", background: infoCeldaDialog.bg, color: infoCeldaDialog.color, whiteSpace: "pre-wrap", wordBreak: "break-word", maxHeight: 120, overflow: "auto" }}>
                  {infoCeldaDialog.contenido || "(vacía)"}
                </div>
              </div>
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
              <button onClick={() => setInfoCeldaDialog(null)}
                style={{ padding: "8px 20px", borderRadius: 6, border: "none", background: "#2563eb", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}
      {showApiLocalAviso && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100001 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", maxWidth: 460, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
              <span style={{ width: 12, height: 12, borderRadius: "50%", background: "#ef4444", display: "inline-block", flexShrink: 0 }} />
              <h3 style={{ fontSize: 15, fontWeight: 700, color: "#991b1b", margin: 0 }}>API local no disponible</h3>
            </div>
            <p style={{ fontSize: 13, color: "#475569", lineHeight: 1.6, margin: "0 0 8px" }}>
              No se ha podido conectar con la API local (<strong>{API_LOCAL_URL}</strong>).
            </p>
            <p style={{ fontSize: 13, color: "#475569", lineHeight: 1.6, margin: "0 0 16px" }}>
              Las funciones <strong>Leer Precios de PMD</strong>, <strong>Crear SimpleQuote</strong> y <strong>Hacer Damex E</strong> no funcionarán hasta que arranques el programa <code style={{ background: "#f1f5f9", padding: "1px 5px", borderRadius: 3, fontSize: 12 }}>CdM_Presupuestos_API_local.py</code> en tu equipo.
            </p>
            <div style={{ display: "flex", justifyContent: "flex-end" }}>
              <button onClick={() => setShowApiLocalAviso(false)}
                style={{ padding: "8px 20px", borderRadius: 6, border: "none", background: "#2563eb", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
                Aceptar
              </button>
            </div>
          </div>
        </div>
      )}
      {showSeleccionar && (
        <SeleccionarCeldasDialog
          totalFilas={rows.length}
          onClose={() => setShowSeleccionar(false)}
          onAceptar={({ desde, hasta }) => {
            // Convertir índices 1-based a IDs de filas
            const ids = new Set();
            for (let i = desde - 1; i <= hasta - 1; i++) {
              if (rows[i]) ids.add(rows[i].id);
            }
            setSelectedRows(ids);
          }}
        />
      )}
      {showLeer && (
        <LeerPresupuestosDialog
          setStatus={setStatus}
          onClose={() => setShowLeer(false)}
          onCargar={(data) => {
            setStatus(`Presupuesto ${data.cabecera.numerocompleto || data.cabecera.numero} cargado: ${(data.detalle || []).length} líneas`, "success");
            // Cargar cabecera
            const cab = data.cabecera;
            setPresupuesto({
              id: String(cab.id),
              np: String(cab.numero || ""),
              numerocompleto: cab.numerocompleto || buildNumeroCompleto(codigoUsuario, cab.numero, cab.anopresupuesto),
              anopresupuesto: cab.anopresupuesto || getAnoFiscal(),
              revision: String(cab.revision ?? 0),
              titulo: cab.titulo || "",
              cliente: cab.nombrecomun || cab.razonsocial || "",
              idcliente: cab.idcliente || null,
              alaatencion: cab.alaatencion || null,
              alaatencion_nombre: cab.alaatencion_nombre || "",
              alaatencion_cliente: cab.alaatencion_cliente_nombre || cab.alaatencion_cliente || "",
            });
            // Cargar líneas, ordenadas por posición
            const lineas = (data.detalle || [])
              .slice()
              .sort((a, b) => (a.posicion || 0) - (b.posicion || 0))
              .map((d, idx) => ({
                id: d.id ?? (Date.now() + idx),
                representacion: d.representacion || "",
                naturaleza: d.naturaleza || "PD",
                posicion: d.posicion != null ? String(d.posicion) : "",
                cantidad: parseInt(d.cantidad, 10) || 0,
                referencia: d.referencia || "",
                nombre: d.nombre || "",
                pvp: Number(d.pvp) || 0,
                dtoaplicado: Number(d.dtoaplicado) || 0,
                descripcion: d.descripcion || "",
                familia: d.familia || "",
                subfamilia: d.subfamilia || "",
                preciocosteunitario: Number(d.preciocosteunitario) || 0,
                idposicion: d.idposicion != null ? String(d.idposicion) : "",
                imagen: d.imagen || "",
                precionetounitario2: Number(d.precionetounitario2) || 0,
                grupodescuento: d.grupodescuento || "",
              }));
            // Añadir 10 filas en blanco al final del presupuesto leído
            let nextBlankId = Math.max(...lineas.map(l => l.id), 0) + 1;
            const filasBlanco = Array.from({ length: 10 }, () => ({
              id: nextBlankId++, representacion: "", naturaleza: "", posicion: "", cantidad: 0,
              referencia: "", nombre: "", pvp: 0, dtoaplicado: 0, descripcion: "",
              familia: "", subfamilia: "", preciocosteunitario: 0, idposicion: "",
              imagen: "", precionetounitario2: 0, grupodescuento: "",
            }));
            const todas = [...lineas, ...filasBlanco];
            setRows(todas);
            setSelectedRows(new Set());
            setSelectionRange(null);
            setSelectedCell(null);
            nextId.current = nextBlankId;
          }}
        />
      )}
      {showEstrategias && (
        <EstrategiasDescuentoDialog
          setStatus={setStatus}
          onClose={() => setShowEstrategias(false)}
          onAplicar={(detalleEstrategia) => {
            // Mapa código de grupo (mayúsculas) → descuento
            const mapa = {};
            detalleEstrategia.forEach(d => {
              const cod = String(d.grupodescuento || "").trim().toUpperCase();
              if (cod) mapa[cod] = Number(d.descuento) || 0;
            });
            let aplicadas = 0;
            setRows(r => r.map(row => {
              if (!["PD", "PE", "E"].includes(row.naturaleza)) return row;
              const gd = String(row.grupodescuento || "").trim().toUpperCase();
              if (gd && mapa[gd] !== undefined) {
                aplicadas++;
                return { ...row, dtoaplicado: mapa[gd] };
              }
              return row;
            }));
            setStatus(`Estrategia aplicada: ${aplicadas} línea(s) actualizada(s)`, aplicadas > 0 ? "success" : "info");
          }}
        />
      )}
      {showDescuentos && (
        <AplicarDescuentosDialog
          rows={rows}
          onClose={() => setShowDescuentos(false)}
          onApply={(grupodescuento, dtoActual, dto) => {
            setRows(r => r.map(row => {
              const gd = (row.grupodescuento || "").trim() || "(Sin grupo)";
              const dtoRow = Math.round((Number(row.dtoaplicado) || 0) * 100) / 100;
              const matchGrupo = gd === grupodescuento;
              const matchDto = dtoRow === dtoActual;
              if (matchGrupo && matchDto && ["PD","PE","E"].includes(row.naturaleza)) {
                return { ...row, dtoaplicado: dto };
              }
              return row;
            }));
          }}
        />
      )}
      {showImportar && (
        <ImportarDialog
          onClose={() => setShowImportar(false)}
          onImport={(nuevas) => {
            setRows(r => {
              if (selectedRows.size === 0) {
                // Insertar al final
                let nextLocalId = Math.max(...r.map(x => x.id), 0) + 1;
                const filas = nuevas.map(n => ({ ...n, id: nextLocalId++ }));
                return [...r, ...filas];
              }
              // Insertar delante de la primera seleccionada (empujando el resto)
              const idx = r.findIndex(row => selectedRows.has(row.id));
              if (idx === -1) {
                let nextLocalId = Math.max(...r.map(x => x.id), 0) + 1;
                const filas = nuevas.map(n => ({ ...n, id: nextLocalId++ }));
                return [...r, ...filas];
              }
              let nextLocalId = Math.max(...r.map(x => x.id), 0) + 1;
              const filas = nuevas.map(n => ({ ...n, id: nextLocalId++ }));
              const res = [...r];
              res.splice(idx, 0, ...filas);
              return res;
            });
          }}
        />
      )}
      {confirmClearRows && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", minWidth: 340, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
              <Icon as={Eraser} size={22} color="#475569" />
              <h2 style={{ fontSize: 15, fontWeight: 600, margin: 0 }}>Confirmar borrado de contenido</h2>
            </div>
            <p style={{ fontSize: 13, color: "#475569", margin: "0 0 20px" }}>
              ¿Desea borrar el contenido de todas las celdas de <strong>{selectedRows.size}</strong> fila{selectedRows.size > 1 ? "s" : ""} seleccionada{selectedRows.size > 1 ? "s" : ""}?
              <br /><span style={{ color: "#94a3b8", fontSize: 12 }}>Las filas se mantendrán pero quedarán vacías.</span>
            </p>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setConfirmClearRows(false)} style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}><BtnContent icon={X}>Cancelar</BtnContent></button>
              <button onClick={() => {
                setRows(rs => rs.map(row => {
                  if (!selectedRows.has(row.id)) return row;
                  const cleared = { ...row };
                  COLUMNS.forEach(col => {
                    if (col.type === "calc") return;
                    cleared[col.key] = col.type === "number" ? 0 : "";
                  });
                  return cleared;
                }));
                setConfirmClearRows(false);
              }}
                style={{ padding: "7px 16px", borderRadius: 6, border: "none", background: "#dc2626", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 500 }}><BtnContent icon={Eraser}>Borrar contenido</BtnContent></button>
            </div>
          </div>
        </div>
      )}
      {confirmDelete && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 99999 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: "1.5rem 2rem", minWidth: 340, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
              <Icon as={Trash2} size={22} color="#dc2626" />
              <h2 style={{ fontSize: 15, fontWeight: 600, margin: 0 }}>Confirmar borrado</h2>
            </div>
            <p style={{ fontSize: 13, color: "#475569", margin: "0 0 20px" }}>
              ¿Desea borrar <strong>{selectedRows.size}</strong> fila{selectedRows.size > 1 ? "s" : ""}?
              <br /><span style={{ color: "#dc2626", fontSize: 12 }}>Esta acción no se puede deshacer.</span>
            </p>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setConfirmDelete(false)} style={{ padding: "7px 16px", borderRadius: 6, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", cursor: "pointer", fontSize: 13 }}><BtnContent icon={X}>Cancelar</BtnContent></button>
              <button onClick={() => { setRows(r => r.filter(row => !selectedRows.has(row.id))); setSelectedRows(new Set()); setConfirmDelete(false); }}
                style={{ padding: "7px 16px", borderRadius: 6, border: "none", background: "#dc2626", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 500 }}><BtnContent icon={Trash2}>Borrar</BtnContent></button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// App envuelta en el Error Boundary para evitar la pantalla en blanco ante un fallo de render
export default function App() {
  return (
    <ErrorBoundary>
      <AppInner />
    </ErrorBoundary>
  );
}
