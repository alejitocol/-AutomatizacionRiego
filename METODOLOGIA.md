# Metodología implementada

Este documento describe las ecuaciones, supuestos y decisiones de implementación de la
herramienta. Cada apartado indica la función de `app.py` donde se materializa el cálculo.

---

## 1. Evapotranspiración de referencia — Hargreaves-Samani

**Función:** `calcular_ret_vectorizado()`

$$ET_0 = k_{RS} \cdot R_a \cdot (T_{media} + 17.8) \cdot \sqrt{T_{max} - T_{min}}$$

Donde:

| Símbolo | Descripción | Unidad |
|---|---|---|
| $ET_0$ | Evapotranspiración de referencia | mm·día⁻¹ |
| $k_{RS}$ | Coeficiente de Hargreaves-Samani | — |
| $R_a$ | Radiación extraterrestre expresada en lámina de agua | mm·día⁻¹ |
| $T_{max}, T_{min}$ | Temperatura máxima y mínima del aire a 2 m | °C |

La radiación extraterrestre se calcula según FAO-56:

$$R_a = \frac{24 \cdot 60}{\pi} \cdot G_{sc} \cdot d_r \cdot \left[\omega_s \sin\varphi \sin\delta + \cos\varphi \cos\delta \sin\omega_s\right]$$

con $G_{sc} = 0.0820$ MJ·m⁻²·min⁻¹, y la conversión a lámina mediante el factor $0.408$.

**Selección de $k_{RS}$:**

| Zona climática | $k_{RS}$ | Criterio |
|---|---|---|
| Árida / semiárida (interior) | 0.0023 | Valor original de Hargreaves-Samani (1985) |
| Húmeda / costera | 0.00185 | Corrección FAO-56 (≈ 0.85 × 0.0023) |
| Rango recomendado | 0.0019 – 0.0025 | FAO-56 |

La clasificación se apoya en el índice de aridez $P/ET_0$ y es ajustable por el usuario.

---

## 2. Agregación decadal

**Función:** `agregar_decadas()`

El año se divide en 36 décadas (tres por mes: días 1–10, 11–20 y 21–fin de mes). La
precipitación se agrega por suma y las temperaturas y la $ET_0$ por promedio, de forma
consistente con la resolución temporal nativa de los productos WaPOR v3.

---

## 3. Precipitación confiable al 75 %

Para cada década se ordena la serie histórica de forma descendente y se asigna la
probabilidad de excedencia mediante la fórmula de Blom:

$$P = \frac{i - 0.375}{n + 0.25} \cdot 100$$

donde $i$ es la posición del dato ordenado y $n$ el número de años de registro. Se interpola
el valor correspondiente al 75 % de probabilidad de excedencia, criterio habitual para el
diseño de sistemas de riego.

---

## 4. Precipitación efectiva — SCS-USDA

$$P_e = \begin{cases} P \cdot \dfrac{125 - 0.2P}{125} & P \leq 250\ \text{mm} \\[2ex] 125 + 0.1P & P > 250\ \text{mm} \end{cases}$$

Se aplica sobre la precipitación decadal confiable, no sobre la precipitación media, con el
fin de mantener el criterio conservador de diseño.

---

## 5. Demanda hídrica del cultivo

$$ET_c = K_c \cdot ET_0 \qquad ; \qquad D_{neta} = ET_c - P_e \qquad ; \qquad D_{bruta} = \frac{D_{neta}}{E_a}$$

El coeficiente $K_c$ se interpola linealmente entre las etapas fenológicas (inicial,
desarrollo, media y final) sobre las 36 décadas del año. La eficiencia de aplicación $E_a$
depende del sistema:

| Sistema | Eficiencia típica |
|---|---|
| Goteo | 0.90 – 0.95 |
| Microaspersión | 0.85 |
| Aspersión | 0.75 |
| Gravedad tecnificada | 0.60 |

> **Conversión de unidades.** El caudal de diseño se obtiene como
> $Q\ (\text{L·s}^{-1}) = \dfrac{D_{bruta}\ (\text{mm}) \cdot A\ (\text{ha}) \cdot 10}{t\ (\text{s})}$.
> El uso del factor 86400 (segundos por día) y no 86.4 es determinante: un error en este
> punto altera el caudal en tres órdenes de magnitud.

---

## 6. Dimensionamiento del reservorio — Método de Rippl en simulación continua

**Ubicación:** pestaña 3, bloque de simulación del tránsito.

Para cada década $t$ se resuelve el balance del embalse:

$$V_t = \min\left(V_{t-1} + A_{ing,t} - D_{t} - E_{t} - I_{t},\ V_{max}\right)$$

| Término | Descripción |
|---|---|
| $A_{ing,t}$ | Aporte por escorrentía del área de captación y precipitación directa |
| $D_t$ | Demanda bruta del cultivo en la década |
| $E_t$ | Evaporación del espejo de agua |
| $I_t$ | Pérdidas por infiltración |

El **volumen útil óptimo $V^*$** se determina mediante búsqueda iterativa: se incrementa
$V_{max}$ hasta que la simulación no registre déficit (o hasta alcanzar la confiabilidad
objetivo definida) durante **toda la serie histórica disponible**, y no únicamente durante un
ciclo de cultivo.

Esta diferencia es el aporte central de la herramienta: el dimensionamiento por ciclo único
subestima sistemáticamente el volumen requerido, porque no captura el agotamiento acumulado
en secuencias de años deficitarios asociadas a eventos ENSO.

---

## 7. Curva área–volumen del reservorio excavado

Para un reservorio de sección troncopiramidal con taludes $z:1$:

$$A(h) = (b + 2zh)(l + 2zh) \qquad ; \qquad V(h) = \int_0^h A(\eta)\,d\eta$$

resuelto de forma discreta por el método del prismoide sobre incrementos de elevación.

---

## 8. Supuestos y limitaciones

1. Se asume que el área de captación aporta escorrentía según un coeficiente único, sin
   modelación de tiempo de concentración.
2. La evaporación del espejo de agua se estima a partir de $ET_0$ afectada por un coeficiente
   de tanque; no se modela estratificación térmica.
3. La infiltración se representa mediante una tasa constante, dependiente del tratamiento de
   impermeabilización seleccionado.
4. No se incorpora aporte por caudal base ni por manantiales.
5. La serie de NASA POWER se emplea con la resolución nativa del producto, sin reducción de
   escala; en terreno montañoso esto introduce incertidumbre en la precipitación.

---

## Referencias

- Allen, R. G., Pereira, L. S., Raes, D., & Smith, M. (1998). *Crop evapotranspiration:
  Guidelines for computing crop water requirements* (FAO Irrigation and Drainage Paper 56). FAO.
- Blom, G. (1958). *Statistical estimates and transformed beta-variables*. Wiley.
- FAO. (2024). *WaPOR v3 database methodology*. Organización de las Naciones Unidas para la
  Alimentación y la Agricultura.
- Hargreaves, G. H., & Samani, Z. A. (1985). Reference crop evapotranspiration from
  temperature. *Applied Engineering in Agriculture, 1*(2), 96–99.
- NASA. (2024). *POWER Data Access Viewer: Methodology*. National Aeronautics and Space
  Administration.
- Rippl, W. (1883). The capacity of storage reservoirs for water supply. *Minutes of the
  Proceedings of the Institution of Civil Engineers, 71*, 270–278.
- USDA Soil Conservation Service. (1970). *Irrigation water requirements* (Technical Release 21).
