# Metodología de análisis — DATA VISTA 1.0

## Detección de variables
La plataforma clasifica columnas mediante heurísticas sobre los valores no vacíos: numérica, temporal, categórica, texto o vacía. La clasificación es una sugerencia automática y debe verificarse en el contexto del estudio.

## Calidad de datos
- Faltantes: valores nulos, indefinidos o cadenas vacías.
- Duplicados: filas idénticas considerando todas las columnas.
- No se imputan ni modifican automáticamente los datos originales.

## Estadística descriptiva
Para cada variable numérica se calcula el número de observaciones válidas, media, mediana, moda cuando existe, desviación estándar poblacional, varianza poblacional, cuartiles, mínimo, máximo y suma.

## Variables categóricas
Se calcula frecuencia absoluta y porcentaje sobre las observaciones no vacías. Se muestran hasta 10 categorías en los gráficos y hasta 8 en tablas por variable.

## Relaciones
La correlación de Pearson se calcula con pares completos numéricos y requiere al menos tres pares válidos. La comparación por grupos resume cantidad de observaciones y promedio. Ninguna de estas medidas permite inferir causalidad por sí sola.

## Territorial
Se buscan encabezados comunes de departamento, municipio, comunidad y coordenadas. La versión 1.0 resume frecuencias territoriales; no valida ni geocodifica ubicaciones.