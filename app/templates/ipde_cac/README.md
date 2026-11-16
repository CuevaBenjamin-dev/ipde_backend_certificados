# IPDE_CAC

El modelo `IPDE_CAC` usa el diseño conjunto del Instituto Peruano de Diplomados y Especializaciones y el Colegio de Abogados del Callao. Están disponibles `CURSO`, `DIPLOMADO` y `PROGRAMA DE ESPECIALIZACIÓN`.

Cada plantilla tiene dos diapositivas: certificado y certificado de estudios. `CURSO` muestra cinco módulos; `DIPLOMADO` y `PROGRAMA DE ESPECIALIZACIÓN` muestran ocho. El backend elige una sola medida de letra para todos los títulos de módulos de cada certificado, según el título más largo, sin reducir filas individualmente. El QR y el código de estudiante se completan por participante.

El archivo de muestra traía una fila `LIBRO N° 003` y otra `CALIFICACIÓN: 19`, sin campos correspondientes en la solicitud actual. Esas filas muestran ahora el código del estudiante y las horas académicas para evitar datos fijos incorrectos en nuevos certificados.
