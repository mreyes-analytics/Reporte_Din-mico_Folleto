# Reporte_Dinámico_Folleto

Este proyecto fue desarrollado como respuesta a la limitante de acceso a datos crudos: permitiendo a los equipos visualizar el performance y cifras de los artículos en folleto para las tiendas de Reynosa.
# Reporte_Folleto – Semi-automatización de Reportes Comerciales en Excel & R

El reporte surge ante la situación de que solo fuera posible extraer cierta información desde reportes de Power BI proporcionados por TI, lo que obstaculiza la automatización directa y eficiente. Por ello, diseñé un **flujo semi-automatizado** que permite consolidar y transformar la información necesaria en pocos pasos, logrando que el reporte quede prácticamente listo al sustituir el rango de datos de la tabla dinámica en Excel.

La solución se implementó usando **RStudio** (para integración, limpieza y cálculo de campos clave) y **Excel** (como plantilla de reporte operativo), incorporando una macro simple para ordenar datos según las necesidades del cuerpo gerencial. El reporte fue solicitado como una solución rápida ante la carga y retrasos habituales del área de TI, permitiendo a los equipos visualizar el performance y cifras de los artículos en folleto para las tiendas de Reynosa.

---

## 🔧 Tecnologías y librerías utilizadas

- **R**: readr, dplyr, stringr, janitor, tidyr, readxl, writexl
- **Excel**: Tablas dinámicas, Macro simple de ordenamiento

---

## 🛠️ **Flujo de trabajo y chunks de código relevantes**

### 1. **Carga y pre-procesamiento del folleto**
```r
Folleto <- read_xlsx("Precios_Ofertas.xlsx")

Folleto <- Folleto %>%
  filter(`TIPO OFERTA` == "FOLLETO") %>%
  mutate(Codigo = as.character(Codigo)) %>%
  mutate(Codigo = if_else(tolower(Codigo) == "varios", "VARIOS", Codigo)) %>%
  mutate(`CÓDIGO COMBO` = if_else(is.na(`CÓDIGO COMBO`), Codigo, `CÓDIGO COMBO`)) %>%
  rename(
    Precio_Normal = DE,
    Precio_Oferta = A
  )
```
2. Carga y filtrado de ventas
```r
Copiar
Editar
Ventas_Folleto <- read_excel("Ventas_Folleto.xlsx")
Ventas_Folleto_filtrado <- Ventas_Folleto %>%
  filter(Sucursal %in% sucursales_interes) %>%
  rename(Codigo = `Codigo de Producto`)
```
3. Join entre ventas y folleto
```r
Copiar
Editar
Ventas_Cruzadas <- Ventas_Folleto_filtrado %>%
  inner_join(
    Folleto %>% select(Codigo, Descripcion, Precio_Normal, Precio_Oferta, `CÓDIGO COMBO`),
    by = c("Codigo" = "Codigo")
  )
```
4. Join con almacén y cálculo de métricas
```r
Copiar
Editar
Ventas_Promocion <- Ventas_Promocion %>%
  left_join(
    Almacen_filtrado %>%
      select(Codigo, Sucursal, Costo_de_Compra),
    by = c("Codigo", "Sucursal")
  ) %>%
  mutate(
    Costo_de_Venta = Cantidad * Costo_de_Compra,
    Utilidad_Oferta = Venta_Precio_Oferta_SinIVA - Costo_de_Venta,
    Margen_Oferta = round(ifelse(Venta_Precio_Oferta_SinIVA == 0, NA, Utilidad_Oferta / Venta_Precio_Oferta_SinIVA * 100), 2),
    Utilidad_Normal = Venta_Precio_Normal_SinIVA - Costo_de_Venta,
    Margen_Normal = round(ifelse(Venta_Precio_Normal_SinIVA == 0, NA, Utilidad_Normal / Venta_Precio_Normal_SinIVA * 100), 2)
  )
```
5. Cálculo de participación y exportación final
```r
Copiar
Editar
ventas_individuales_total <- Ventas_Todas %>%
  filter(Tipo_Registro != "Combo") %>%
  summarise(total = sum(ImporteVta, na.rm = TRUE)) %>%
  pull(total)

Ventas_Todas <- Ventas_Todas %>%
  mutate(
    Participacion_SKU = if_else(
      Tipo_Registro == "Combo",
      round(Venta_Precio_Oferta_ConIVA / ventas_individuales_total * 100, 2),
      Participacion_SKU
    )
  )

write_xlsx(Ventas_Todas, "Reporte_Ventas_Todas5.xlsx")
```

📊 Output final
El resultado es un archivo Excel con datos listos para análisis con tabla dinámica, permitiendo a los gerentes ordenar, filtrar y consultar el performance de artículos en folleto de manera rápida y sencilla.
![Reporte Excel Output](./pivot_macro.png)

📝 Macro UserForm de ordenamiento
Aunque Excel permite ordenar tablas dinámicas desde el menú contextual, este UserForm fue creado para facilitar la experiencia del usuario y agilizar el análisis, permitiendo ordenar la información clave con un solo clic. Esta mejora es especialmente útil en contextos donde el reporte es utilizado por perfiles no técnicos o donde se requiere rapidez y precisión operativa.

El archivo userform_ordenar_pivot.txt contiene el código completo del UserForm para ordenar automáticamente las tablas dinámicas del reporte según diferentes criterios: venta, participación, margen y utilidad.

Instrucciones rápidas:
Abre el Editor de VBA en Excel (ALT + F11)

Inserta un nuevo UserForm y personaliza el fondo (azul navy)

Agrega los botones y nómbralos:

btnOrdenVentaAsc / btnOrdenVentaDesc

btnOrdenPartAsc / btnOrdenPartDesc

btnOrdenMargenAsc / btnOrdenMargenDesc

btnOrdenUtilAsc / btnOrdenUtilDesc

Pega el código del archivo .txt en el módulo del UserForm

Ajusta los nombres de las tablas dinámicas y campos según tu reporte

El UserForm está diseñado para posicionarse automáticamente en la hoja y aplicar el ordenamiento con solo hacer clic en los botones. Puede adaptarse fácilmente a cualquier tabla dinámica cambiando los nombres de los campos y la hoja objetivo.

🧠 Reflexión profesional
Este proyecto ejemplifica cómo, aún ante limitaciones de acceso y automatización, es posible resolver problemas reales del negocio integrando herramientas como R y Excel. Convertí una tarea que dependía del área de TI en un proceso semi-automático, ágil y amigable para el usuario final.
La solución fue muy bien recibida por el área comercial y demuestra la importancia de tener habilidades mixtas de análisis, integración de datos y foco en el usuario de negocio, especialmente en ambientes donde los recursos técnicos pueden ser limitados o lentos.

Nota:
Los datos han sido anonimizados y las imágenes editadas por motivos de confidencialidad. El flujo y la estructura del archivo reflejan la lógica implementada, no datos reales de la empresa.

## 📧 Contacto

reyes061295@gmail.com  
[LinkedIn](https://www.linkedin.com/in/marb951206/) | [GitHub](https://github.com/mreyes-analytics)
