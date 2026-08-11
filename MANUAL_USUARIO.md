# Manual de usuario

Guía de operación de la Herramienta Computacional ADR. El recorrido emplea como ejemplo el
caso de estudio de maíz en el municipio de Pelaya (Cesar).

---

## Inicio

```bash
source .venv/bin/activate      # Windows: .venv\Scripts\activate
streamlit run app.py
```

La aplicación abre en `http://localhost:8501`.

---

## Pestaña 1 — Datos Agroclimáticos

**Objetivo:** obtener la serie climática del punto de interés.

1. Ingrese la **latitud** y **longitud** del predio en grados decimales
   (ejemplo: `8.848795`, `-73.609039`).
2. Seleccione la fuente:
   - **NASA POWER (API en línea).** Defina fecha inicial y final. La descarga es automática
     y queda en caché durante la sesión.
   - **WaPOR v3 (archivos ráster .ZIP).** Cargue hasta tres archivos comprimidos con los
     GeoTIFF decadales de precipitación, evaporación y evapotranspiración de referencia.
     Los archivos deben conservar la fecha en el nombre (formato `AAAA-MM-DD` o `AAAAMMDD`).
3. Revise el bloque **Análisis de probabilidad** para verificar la precipitación confiable
   al 75 % por década.
4. Descargue la serie diaria o decadal en CSV si requiere respaldo documental.

> **Verificación recomendada.** Antes de continuar, abra el bloque de revisión de valores
> crudos y confirme que no existan décadas con precipitación cero sistemática: esto suele
> indicar que un ráster no fue leído o que el punto cae fuera del área del producto.

---

## Pestaña 2 — Balance Hídrico

**Objetivo:** calcular la demanda hídrica del cultivo.

1. **Ubicación y periodo.** Se heredan de la pestaña 1; confirme la coincidencia.
2. **Parámetros agronómicos.**
   - Cultivo y fecha de siembra.
   - Duración de las etapas fenológicas (inicial, desarrollo, media, final) en días.
   - Valores de $K_c$ por etapa.
   - Profundidad radicular efectiva y área a irrigar.
3. **Configuración del sistema de riego.** Seleccione el sistema (goteo, aspersión,
   microaspersión) y su eficiencia de aplicación.
4. **Coeficiente $k_{RS}$.** Ajústelo según la zona climática; para el valle del Cesar se
   emplea 0.0023.
5. Revise el **Resumen del balance decadal** y la **distribución de $K_c$** en las 36 décadas.

---

## Pestaña 3 — Volúmenes de Riego

**Objetivo:** determinar el volumen útil del almacenamiento.

1. Seleccione la **fuente climática para la simulación** (NASA POWER o WaPOR).
2. Escoja la alternativa:
   - **Reservorio excavado.** Defina dimensiones de fondo, talud, profundidad, coeficiente
     de escorrentía, área de captación e impermeabilización.
   - **Tanque australiano.** Seleccione el modelo del catálogo comercial o defina dimensiones
     propias, junto con el área de techo para cosecha de aguas lluvias.
3. Ejecute la simulación. Obtendrá:
   - Comportamiento del volumen almacenado a lo largo de la serie.
   - Curva área–volumen frente a la elevación.
   - Tabla de simulación decadal.
   - Esquema espacial en planta del proyecto.
4. Consulte el bloque **Recomendación por método de Rippl** para el volumen útil `V*`.
5. Revise el **diagnóstico de resiliencia**: indica el número de décadas con déficit y el
   porcentaje de confiabilidad alcanzado.

> **Interpretación de `V*`.** Es el volumen mínimo que evita el déficit durante toda la
> ventana histórica simulada. Si el análisis de sensibilidad muestra que `V*` sigue creciendo
> al ampliar la ventana, la serie disponible aún no captura el episodio seco crítico.

---

## Pestaña 4 — Reporte comparativo y generación de anexos

**Objetivo:** contrastar fuentes y producir las memorias de cálculo.

1. Revise la **Tabla 7** (volumen óptimo por fuente y ventana histórica) y la **Tabla 8**
   (episodios secos críticos identificados por año).
2. Analice la **matriz año × década**, que muestra el porcentaje de `V*` alcanzado y permite
   identificar visualmente los periodos de agotamiento.
3. Diligencie los datos del proyecto: nombre, departamento, municipio, beneficiario e
   identificador del predio.
4. Genere los anexos:
   - **Anexo 3** — Hidrología básica e información de reservorios (Word).
   - **Anexo 6** — Memoria de demandas hídricas (Word).
   - **Anexo 7** — Memoria de cálculo hidráulico (Word).
   - **Anexo 7a** — Cálculo hidráulico (Excel).

---

## Solución de problemas frecuentes

| Síntoma | Causa probable | Solución |
|---|---|---|
| `ModuleNotFoundError: generadores_anexos` | El módulo no está en la carpeta de `app.py` | Copie `generadores_anexos.py` junto a `app.py` |
| La descarga de NASA POWER falla o se agota el tiempo | Servicio no disponible o periodo demasiado extenso | Reintente o divida el periodo en tramos |
| Todas las décadas WaPOR reportan cero | Nombres de archivo sin fecha reconocible, o punto fuera del ráster | Verifique el formato del nombre y las coordenadas |
| El archivo ZIP no carga | Excede el límite configurado | Ajuste `maxUploadSize` en `.streamlit/config.toml` |
| Los anexos se generan sin gráficas | La simulación no se ejecutó en la sesión actual | Ejecute primero las pestañas 2 y 3 |
