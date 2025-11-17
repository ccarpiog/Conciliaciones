# Manual de Uso - Sistema de Conciliación Bancaria

Sistema automatizado de conciliación bancaria para Google Sheets que empareja movimientos contables con extractos bancarios utilizando inteligencia artificial para detectar coincidencias.

## Tabla de Contenidos

- [Requisitos Previos](#requisitos-previos)
- [Instalación](#instalación)
- [Configuración Inicial](#configuración-inicial)
- [Uso Básico](#uso-básico)
- [Características Avanzadas](#características-avanzadas)
- [Interpretación de Resultados](#interpretación-de-resultados)
- [Resolución de Conflictos](#resolución-de-conflictos)
- [Configuración Avanzada](#configuración-avanzada)
- [Solución de Problemas](#solución-de-problemas)

---

## Requisitos Previos

- Cuenta de Google con acceso a Google Sheets
- Hoja de cálculo con datos contables y bancarios
- Permisos de edición en la hoja de cálculo

## Instalación

### Paso 1: Preparar el Proyecto

1. Abre tu hoja de cálculo de Google Sheets
2. Ve a **Extensiones** → **Apps Script**
3. Borra el código por defecto (`function myFunction() {}`)

### Paso 2: Añadir el Código

1. Copia el contenido del archivo `Code.js` en el editor de Apps Script
2. Haz clic en el icono **+** junto a "Archivos" para añadir archivos HTML
3. Crea un archivo llamado `ConflictsSidebar` y pega el contenido de `ConflictsSidebar.html`
4. Crea otro archivo llamado `ConfigDialog` y pega el contenido de `ConfigDialog.html`
5. Guarda el proyecto con **Ctrl+S** o **Cmd+S**

### Paso 3: Primera Ejecución

1. Cierra el editor de Apps Script
2. Recarga la hoja de cálculo (F5 o Cmd+R)
3. Aparecerá un nuevo menú llamado **Conciliación** en la barra de menú
4. La primera vez que lo uses, Google te pedirá autorización para ejecutar el script

---

## Configuración Inicial

### Estructura de Datos Requerida

El sistema necesita dos hojas específicas:

#### 1. Hoja "Origen"

Debe contener los siguientes datos:

**Datos Contables (Columnas A-D):**
- **Columna A**: Fecha del movimiento contable
- **Columna B**: Número de asiento
- **Columna C**: Concepto del movimiento
- **Columna D**: Importe

**Datos Bancarios (Columnas F-J):**
- **Columna F**: Fecha del movimiento bancario
- **Columna G**: Fecha valor (opcional)
- **Columna H**: Concepto del banco
- **Columna I**: Datos adicionales del banco
- **Columna J**: Importe

**Ejemplo:**

| A (Fecha) | B (Asiento) | C (Concepto) | D (Importe) | E | F (Fecha) | G (F. Valor) | H (Concepto) | I (Datos Adic.) | J (Importe) |
|-----------|-------------|--------------|-------------|---|-----------|--------------|--------------|-----------------|-------------|
| 01/01/2024 | 1 | Cheque 661112 | -150.00 | | 01/01/2024 | 02/01/2024 | CHQ 1112 | Pago proveedor | -150.00 |
| 05/01/2024 | 2 | Transferencia | 500.00 | | 06/01/2024 | 06/01/2024 | TRANSF | Cliente A | 500.00 |

#### 2. Hoja "Salida"

Crea una hoja vacía llamada "Salida". El sistema la llenará automáticamente con los resultados.

---

## Uso Básico

### Ejecutar Conciliación Automática

1. Ve al menú **Conciliación** → **Ejecutar conciliación automática**
2. El sistema procesará todos los movimientos
3. Aparecerá un resumen con estadísticas:
   - ✓ Movimientos conciliados automáticamente
   - ⚠ Conflictos que requieren revisión manual
   - ✗ Movimientos sin conciliar

### Resultados en la Hoja "Salida"

La hoja de salida mostrará todos los movimientos ordenados por:
1. Fecha contable (de menor a mayor)
2. Número de asiento (dentro de cada fecha)

**Código de Colores:**
- 🟢 **Verde**: Movimientos conciliados automáticamente
- 🟡 **Amarillo**: Conflictos que necesitan revisión
- 🔴 **Rojo**: Movimientos sin conciliar

---

## Características Avanzadas

### Algoritmo de Conciliación

El sistema utiliza múltiples técnicas para encontrar coincidencias:

#### 1. Coincidencia Exacta de Importes (Obligatoria)
- Los importes deben ser exactamente iguales
- Se redondea a 2 decimales para evitar errores de precisión

#### 2. Tolerancia de Fechas
- Por defecto: ±3 días
- Configurable en **Conciliación** → **Configurar parámetros**
- Ejemplo: Un movimiento del 10/01 puede coincidir con uno del 12/01

#### 3. Similitud de Conceptos (70% del peso)

El sistema detecta similitudes en los conceptos usando:

**a) Coincidencia Exacta (100%)**
- Los conceptos son idénticos (ignorando mayúsculas/minúsculas)

**b) Contención (80%)**
- Un concepto contiene al otro completamente
- Ejemplo: "Pago Factura 123" contiene "Factura 123"

**c) Números Coincidentes (60-70%)**
- Detecta números comunes en ambos conceptos
- **Números exactos**: "Cheque 661112" y "CHQ 661112" → coincidencia alta
- **Números parciales**: "Cheque 661112" y "1112" → coincidencia media
- **Sufijos comunes**: "661112" termina con "1112" → coincidencia alta

**d) Tokens Comunes (30-70%)**
- Busca palabras comunes entre conceptos
- Ejemplo: "Transferencia bancaria" y "TRANSF BANC" → coincidencia media

**e) Distancia de Levenshtein (0-50%)**
- Para textos cortos, calcula similitud por caracteres
- Útil para detectar errores tipográficos

#### 4. Criterios de Decisión

**Conciliación Automática:**
- Solo hay un candidato con el mismo importe, O
- El mejor candidato tiene:
  - Puntuación > 30% (configurable)
  - Al menos 20 puntos más que el segundo mejor

**Conflicto (Revisión Manual):**
- Múltiples candidatos con puntuaciones similares
- Puntuación del mejor candidato < 30%

---

## Interpretación de Resultados

### Columnas en la Hoja "Salida"

| Columna | Descripción |
|---------|-------------|
| Fecha Cont. | Fecha del movimiento contable |
| Asiento | Número de asiento contable |
| Concepto Cont. | Descripción contable |
| Importe | Cantidad del movimiento |
| Estado | ✓ Conciliado / ⚠ Conflicto / ✗ Sin conciliar |
| Fecha Banco | Fecha del movimiento bancario |
| Fecha Valor | Fecha valor del banco (si aplica) |
| Concepto Banco | Descripción bancaria |
| Datos Adic. | Información adicional del banco |
| Puntuación | Confianza de la coincidencia (0-100%) |

### Sección de Movimientos Bancarios No Conciliados

Al final de la hoja aparecerá una sección en rojo con todos los movimientos bancarios que no pudieron conciliarse con ningún movimiento contable. Estos pueden indicar:
- Movimientos registrados en el banco pero no en contabilidad
- Errores en los importes
- Movimientos pendientes de registro

---

## Resolución de Conflictos

### ¿Qué es un Conflicto?

Un conflicto ocurre cuando:
- Varios movimientos bancarios tienen el mismo importe
- Las puntuaciones de similitud son parecidas
- El sistema no puede decidir automáticamente

### Revisar Conflictos

1. Ve a **Conciliación** → **Revisar conflictos**
2. Se abrirá un panel lateral con todos los conflictos
3. Para cada movimiento contable verás:
   - Datos del movimiento contable
   - Lista de candidatos bancarios ordenados por puntuación
   - Puntuación de confianza (código de colores)

**Código de Colores de Puntuación:**
- 🟢 **Verde**: ≥ 70% (alta confianza)
- 🟡 **Amarillo**: 40-69% (confianza media)
- 🔴 **Rojo**: < 40% (baja confianza)

### Resolver un Conflicto

1. **Seleccionar Candidato**: Haz clic en el movimiento bancario correcto
2. **Confirmar**: Presiona "Confirmar Conciliación"
3. El conflicto se marca como resuelto (actualmente visual)

### Aplicar Todas las Conciliaciones

Si has revisado varios conflictos:
1. Selecciona el candidato correcto en cada uno
2. Haz clic en "Aplicar Todas las Conciliaciones" al final
3. Confirma la acción

### Omitir un Conflicto

Si no estás seguro o necesitas más información:
- Haz clic en "Omitir"
- El conflicto permanecerá para revisión posterior

---

## Configuración Avanzada

### Acceder a la Configuración

**Conciliación** → **Configurar parámetros**

### Parámetros Disponibles

#### 1. Tolerancia de Fechas (0-10 días)

Define cuántos días de diferencia se permiten entre fechas contables y bancarias.

- **0 días**: Las fechas deben coincidir exactamente
- **3 días** (recomendado): Permite diferencias de hasta 3 días
- **7-10 días**: Para casos donde hay retrasos frecuentes

**Ejemplo:**
- Tolerancia: 3 días
- Movimiento contable: 10/01/2024
- Movimiento bancario: 12/01/2024
- ✓ Se considera coincidencia (2 días de diferencia)

#### 2. Puntuación Mínima de Similitud (0-100%)

Establece el umbral mínimo para conciliación automática.

- **10-20%**: Muy permisivo, más conciliaciones automáticas (riesgo de errores)
- **30-40%** (recomendado): Balance entre automatización y precisión
- **50-70%**: Conservador, menos automático pero más preciso

**Recomendaciones:**
- Datos bien estructurados: 30-40%
- Conceptos muy variables: 20-30%
- Máxima precisión: 50-60%

### Criterios de Conciliación (No Configurable)

El sistema siempre usa estos pesos:
- **Fecha**: 30% del total
- **Concepto**: 70% del total

---

## Solución de Problemas

### Error: "No se encontraron las hojas 'Origen' o 'Salida'"

**Causa**: Las hojas necesarias no existen o tienen nombres incorrectos.

**Solución:**
1. Verifica que exista una hoja llamada exactamente "Origen"
2. Crea una hoja llamada exactamente "Salida" (puede estar vacía)
3. Los nombres distinguen mayúsculas/minúsculas

### No se Concilian Movimientos Obvios

**Posibles causas:**
1. **Importes diferentes**: Verifica que sean exactamente iguales (incluyendo decimales)
2. **Tolerancia de fechas insuficiente**: Aumenta la tolerancia en configuración
3. **Puntuación mínima muy alta**: Reduce el umbral de similitud

**Solución:**
1. Ve a **Configurar parámetros**
2. Aumenta "Tolerancia de fechas" a 5-7 días
3. Reduce "Puntuación mínima" a 20-30%
4. Ejecuta de nuevo la conciliación

### Demasiados Conflictos

**Causa**: Los conceptos son muy diferentes entre contabilidad y banco.

**Solución:**
1. Reduce la "Puntuación mínima de similitud" a 20-25%
2. Usa "Revisar conflictos" para resolver manualmente
3. Considera estandarizar los conceptos en origen

### Movimientos Bancarios No Conciliados

**Causa**: No existe movimiento contable con el mismo importe.

**Acciones recomendadas:**
1. Revisa la sección roja al final de "Salida"
2. Verifica si falta registrar algo en contabilidad
3. Comprueba si hay errores en los importes
4. Confirma que todos los datos están en la hoja "Origen"

### Error de Autorización

**Causa**: Google necesita permisos para ejecutar el script.

**Solución:**
1. La primera vez que uses el menú, aparecerá una ventana de autorización
2. Haz clic en "Revisar permisos"
3. Selecciona tu cuenta de Google
4. Haz clic en "Avanzado" → "Ir a [nombre del proyecto] (no seguro)"
5. Haz clic en "Permitir"

### El Menú "Conciliación" No Aparece

**Solución:**
1. Recarga la página (F5 o Cmd+R)
2. Espera unos segundos para que cargue el script
3. Si persiste, ve a **Extensiones** → **Apps Script**
4. Verifica que el código esté guardado correctamente
5. Ejecuta manualmente `onOpen()` desde el editor de Apps Script

---

## Limpieza y Mantenimiento

### Limpiar Hoja de Salida

Si necesitas ejecutar la conciliación de nuevo desde cero:

**Conciliación** → **Limpiar hoja de salida**

Esto borrará todos los resultados anteriores. Los datos en "Origen" no se modifican.

### Actualizar Conflictos

Si has modificado datos en "Origen" y quieres revisar conflictos actualizados:

1. Abre **Revisar conflictos**
2. Haz clic en "Actualizar" al final del panel
3. Se recargarán los conflictos con los datos actuales

---

## Consejos y Mejores Prácticas

### Preparación de Datos

1. **Encabezados**: Asegúrate de que la fila 1 contenga encabezados
2. **Fechas**: Usa formato de fecha de Google Sheets (no texto)
3. **Importes**: Formato numérico, sin símbolos de moneda en celdas
4. **Celdas vacías**: Evita filas con celdas críticas vacías (fecha o importe)

### Estrategia de Conciliación

1. **Primera ejecución**: Usa configuración por defecto (3 días, 30%)
2. **Revisar resultados**: Verifica movimientos conciliados automáticamente
3. **Ajustar parámetros**: Si hay muchos errores o conflictos, ajusta configuración
4. **Resolver conflictos**: Usa el panel lateral para casos ambiguos
5. **Re-ejecutar**: Después de ajustar configuración

### Optimización

**Para maximizar conciliaciones automáticas:**
- Estandariza conceptos cuando sea posible
- Incluye números de referencia en ambos sistemas
- Mantén consistencia en formatos de fecha
- Verifica importes antes de importar

**Para máxima precisión:**
- Usa tolerancia de fechas baja (0-1 días)
- Mantén puntuación mínima alta (40-50%)
- Revisa manualmente todos los conflictos

---

## Soporte y Contacto

Para problemas técnicos o sugerencias de mejora:
- Revisa este manual primero
- Verifica la sección "Solución de Problemas"
- Consulta con tu administrador de sistemas

---

**Versión del Manual**: 1.0
**Última Actualización**: 2024

---

## Apéndice: Ejemplos de Casos de Uso

### Caso 1: Cheques con Diferentes Formatos

**Contabilidad**: "Cheque 661112"
**Banco**: "CHQ 1112"

✓ **Coincidencia detectada**: 70% (sufijo numérico común "1112")

### Caso 2: Transferencias

**Contabilidad**: "Transferencia Cliente A"
**Banco**: "TRANSF BANCARIA CLIENTE"

✓ **Coincidencia detectada**: ~60% (tokens comunes: "transf", "cliente")

### Caso 3: Domiciliaciones

**Contabilidad**: "Domiciliación Luz 15/01"
**Banco**: "DOMICIL LUZ"

✓ **Coincidencia detectada**: ~65% (tokens "domicil", "luz")

### Caso 4: Fechas con Diferencia

**Contabilidad**: 15/01/2024
**Banco**: 17/01/2024

✓ **Coincidencia**: Si tolerancia ≥ 2 días
✗ **Conflicto**: Si tolerancia < 2 días

### Caso 5: Múltiples Movimientos Mismo Importe

**Contabilidad**: "Proveedor A" - 150.00€
**Banco 1**: "PAGO PROV" - 150.00€ (puntuación: 55%)
**Banco 2**: "FACTURA" - 150.00€ (puntuación: 30%)

⚠ **Conflicto**: Se requiere revisión manual (diferencia de puntuación < 20 puntos)
