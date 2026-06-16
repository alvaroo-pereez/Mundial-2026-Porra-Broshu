# Porra Mundial 2026

Herramienta profesional en Excel para gestionar porras del Mundial: pronósticos, resultados, puntuación automática y clasificación en tiempo real. Hay **dos grupos independientes** en paralelo:

| Grupo | Jugadores | Excel | Dashboard |
|-------|-----------|-------|-----------|
| **Broshu** | 11 | `output/Porra_Mundial_2026.xlsx` | `output/dashboard_broshu.html` |
| **Papinenes** | 3 (Álvaro, Papá, Diego) | `output/Porra_Mundial_2026_Papinenes.xlsx` | `output/dashboard_papinenes.html` |

La configuración de cada grupo está en **`config/groups/<grupo>.json`** (jugadores, rutas, carpeta de fotos).

## Archivo listo para usar

Abre el Excel de tu grupo con **Microsoft Excel 2010 o superior** (Windows o Mac). Las fórmulas usan `INDEX`/`MATCH`, `SUMIFS`, `LARGE`, `AGGREGATE`, etc.

Tras abrir, recalcula si hace falta: `Ctrl + Alt + F9`.

## Dashboard web (navegador)

Abre el HTML de tu grupo en Chrome/Edge/Firefox. Es una sola página con navegación interna (no necesita servidor) y tres vistas:

- **Inicio** (`#/`): KPIs, clasificación con cortes alternativos (por puntos, por % de acierto y por exactos), gráficos (puntos totales, origen de los puntos, evolución del Top 3, composición de aciertos), últimos resultados y resumen de apuestas especiales.
- **Partidos** (`#/partidos`): los 104 partidos con sus resultados reales, agrupados por fase y filtrables por fase, estado y equipo.
- **Jugador** (`#/jugador/<nombre>`): ficha por jugador con el desglose de puntos, sus apuestas especiales y cada pronóstico coloreado según el nivel de acierto (exacto, diferencia, ganador, clasificado, sin acierto o pendiente), junto al resultado real.

Tras editar el Excel, actualiza el dashboard del grupo:

```powershell
py build_dashboard.py --group broshu
py build_dashboard.py --group papinenes
```

(O se regenera solo al ejecutar `py generate_porra.py --group <grupo>`.)

El dashboard carga los datos desde `output/{grupo}/data.json` (se actualiza al regenerar). En Netlify, el navegador refresca esos datos cada 5 minutos automáticamente.

### Fotos del dashboard

Cada grupo tiene su carpeta de imágenes fuente:

- **Broshu:** `Fotos a usar en dashboard/`
- **Papinenes:** `Fotos a usar en dashboard Papinenes/`

En ambas:
- **`Portada.png`** — cabecera del dashboard.
- **Un PNG por jugador** — ficha personal (`#/jugador/<nombre>`). Sin acentos en el nombre de archivo (`Alvaro.png`, `Papa.png`); excepción Broshu: `Quinte.png` para Quintero.

Al ejecutar `py build_dashboard.py --group <grupo>`, las fotos se copian a `output/<grupo>/` (p.ej. `output/broshu/portada.jpg`, `output/papinenes/photos/diego.jpg`).

## Jugadores por grupo

**Broshu:** Álvaro, Pepe, Patri, Kike, Quintero, Nacho, Luis, Felipe, Simón, Muni, Fer.

**Papinenes:** Álvaro, Papá, Diego.

## Hojas del libro (orden recomendado)

| Hoja | Uso |
|------|-----|
| **Portada** | Presentación del torneo y lista de jugadores. |
| **Instrucciones** | Guía rápida de uso. |
| **Resumen** | Dashboard estilo consulting: KPIs, clasificación completa, análisis y gráficos. |
| **Partidos** | Calendario (104 partidos). Introduce aquí los resultados reales. |
| **Pronosticos** | Vista consolidada (solo lectura); lee de cada pestaña de jugador. |
| **Helpers >>** | Separador visual antes de las pestañas personales y técnicas. |
| **Tu nombre** (una pestaña por jugador) | Cada jugador rellena aquí sus pronósticos y apuestas especiales (celdas amarillas). |
| **Puntuacion** | Tabla editable con los puntos de cada acierto (para ajustar pesos). |
| **_Helpers** | Cálculos internos y tablas para gráficos (sin input de usuario). |

## Cómo usar

1. Abre **la pestaña con tu nombre** y escribe **goles enteros ≥ 0** en *Mis goles local* y *Mis goles visit.*.
2. En grupos, deja *Clasificado* con `-`. En eliminatorias, elige **Local** o **Visitante**.
3. Rellena tus **apuestas especiales** al final de tu propia pestaña.
4. La persona que administre resultados usa **Partidos** y el bloque de apuestas especiales de **Resumen**.
5. Los **puntos**, **Resumen** y **Pronosticos** se actualizan solos.

Regla visual: **amarillo = input**, **azul = fórmula**, **gris = referencia de otra pestaña**.

## Sistema de puntuación

Los puntos se definen en la hoja **Puntuacion** (editable).

### Fase de grupos
- Marcador exacto: **5**
- Diferencia de goles correcta: **3**
- Ganador/empate correcto: **2**
- Fallo: **0**

Solo aplica el **máximo** (no acumulativo).

### Eliminatorias
- Marcador exacto (90 min): **10**
- Diferencia de goles (90 min): **6**
- Clasificado correcto (incluye prórroga/penaltis): **4**
- Fallo: **0**

### Bonus
- Bonus por ronda perfecta (Octavos/Cuartos/Semifinales): **+15** por ronda.

### Apuestas especiales
8 categorías; cada jugador las rellena en su propia pestaña: Campeón del Mundial, Subcampeón, Tercer equipo, Balón de oro de la FIFA, Premio de la FIFA al mejor jugador joven, Guante de oro de la FIFA, Bota de oro de la FIFA y Máximo goleador España. En **Resumen** se introducen los resultados oficiales y se ven los puntos por apuestas especiales. Cada acierto: **+10**.

## Datos de demostración

Los **5 primeros partidos** ya tienen resultado real y pronósticos de ejemplo para comprobar puntuaciones 0, 1 y 2 al abrir el archivo.

## Regenerar el Excel

Si cambias jugadores o partidos:

```powershell
py -m pip install -r requirements.txt
py build_matches.py                    # opcional: regenerar calendario JSON
py generate_porra.py --group broshu    # Excel + dashboard Broshu
py generate_porra.py --group papinenes # Excel + dashboard Papinenes
py generate_porra.py --group all       # ambos grupos
```

Edita `config/groups/<grupo>.json` (jugadores, fotos, nombres de archivo) y `data/matches_2026.json` (calendario compartido) antes de regenerar.

## Cromos Panini (Mundial 2026)

Genera cromos tipo Panini con equipación de España para cada amigo de la porra.

### Cromo de Felipe (listo)

El primer cromo está en **`output/cromos/felipe.png`** (posición **DEFENSA**, dorsal **#08**).

### Generar un cromo

```powershell
py -m pip install -r requirements.txt

# Retrato + cromo completo
py generate_cromos.py --player felipe

# Solo recomponer el marco (si ya tienes el retrato en output/cromos/_portraits/)
py generate_cromos.py --player felipe --compose-only

# Regenerar retrato forzando
py generate_cromos.py --player felipe --force-portrait
```

### Retrato con IA (equipación realista)

Para transformar la foto con camiseta oficial de España, configura tu clave de OpenAI:

```powershell
$env:OPENAI_API_KEY = "sk-..."
py generate_cromos.py --player felipe --force-portrait
```

Si no hay API key, el script usa un **retrato estilizado** (camiseta simplificada sobre la foto). Para máxima calidad, también puedes generar el retrato manualmente (ChatGPT, Gemini, etc.) con el prompt de **`prompts/cromo_portrait.txt`** y guardarlo como `output/cromos/_portraits/<id>_portrait.png`; luego ejecuta con `--compose-only`.

### Añadir los otros 10 jugadores

1. Copia la foto a `assets/photos/<id>.png` (ej. `assets/photos/pepe.png`).
2. Añade una entrada en **`config/cromos.json`**:

```json
{
  "id": "pepe",
  "name": "Pepe",
  "position": "DELANTERO",
  "photo": "assets/photos/pepe.png",
  "number": 1
}
```

3. Ejecuta: `py generate_cromos.py --player pepe`

Los IDs sugeridos para Broshu (según `config/groups/broshu.json`): `alvaro`, `pepe`, `patri`, `kike`, `quintero`, `nacho`, `luis`, `felipe`, `simon`, `muni`, `fer`.

El diseño del marco (borde dorado, bandera, franja roja, badge de posición) es **idéntico para todos**; solo cambian retrato, nombre, posición y número.

## Sync automático (openfootball + Netlify)

Actualiza resultados reales en Excel y en el dashboard **sin intervención manual**. Usa el dataset gratuito [openfootball/worldcup.json](https://github.com/openfootball/worldcup.json) — **sin API key ni registro**.

### Cómo funciona

1. **GitHub Actions** ejecuta `sync_results.py` cada **10 min, 24 h** durante el torneo (y un segundo intento 15 min después si hubo cambios).
2. **1 descarga JSON** por sync desde GitHub (`worldcup.json` 2026).
3. Si hay resultados nuevos → actualiza Excel (Broshu + Papinenes) → regenera `output/{grupo}/data.json`.
4. **Commit + push** solo si hubo cambios → Netlify redeploya (o usa Build Hook).
5. El dashboard carga `data.json` y se **refresca solo cada 2 min** en el navegador.

### Configuración única

**1. GitHub**

- Crea un repo y sube el proyecto (incluye los `.xlsx` de `output/`).
- En **Settings → Secrets → Actions**, añade solo:
  - `NETLIFY_BUILD_HOOK` — URL del Build Hook de Netlify (opcional si Netlify ya está conectado al repo)

**2. Netlify**

- **Add site → Import from Git** (sustituye la subida manual).
- Build command: *(vacío)*
- Publish directory: `output` (o usa [`netlify.toml`](netlify.toml))
- Build Hook: Site settings → Build hooks → copiar URL al secret de GitHub

**3. Verificar que todo está listo**

```powershell
py verify_automation.py
py verify_automation.py --skip-api   # solo Excel (sin llamar a openfootball)
```

**4. Diagnóstico de emparejamiento** (opcional)

```powershell
py bootstrap_fixture_map.py
```

Genera un informe en `data/api_fixture_map.json` (no es obligatorio para el sync).

**5. Probar sync**

```powershell
py sync_results.py --dry-run   # ver qué cambiaría
py sync_results.py             # aplicar + regenerar dashboards
```

Luego en GitHub: **Actions** → **Sync resultados Mundial** → **Run workflow**.

### Archivos relevantes

| Archivo | Uso |
|---------|-----|
| [`worldcup_data.py`](worldcup_data.py) | Cliente openfootball (fetch + emparejamiento) |
| [`verify_automation.py`](verify_automation.py) | Preflight antes del sync |
| [`sync_results.py`](sync_results.py) | Orquestador principal |
| [`bootstrap_fixture_map.py`](bootstrap_fixture_map.py) | Diagnóstico de emparejamiento (opcional) |
| [`config/team_mapping.json`](config/team_mapping.json) | Nombres ES → inglés (openfootball) |
| [`.github/workflows/sync-results.yml`](.github/workflows/sync-results.yml) | Cron automático |
| [`.github/workflows/verify-fixture-map.yml`](.github/workflows/verify-fixture-map.yml) | Diagnóstico manual del mapeo |

### Coste

Gratis: openfootball (JSON público), GitHub Actions y Netlify free tier.

## Edición libre

Todas las hojas están **sin protección**: puedes modificar cualquier celda. Si cambias fórmulas en **_Helpers** o **Clasificacion**, haz una copia de seguridad antes.

## Macros (opcional)

Esta versión **no usa macros**: todo funciona con fórmulas. Si quieres botones (*Resetear pronósticos*, etc.), guarda una copia como `.xlsm` y añade VBA en Excel.

## Limitaciones

- Compatible con Excel 2010+ (no requiere XLOOKUP ni funciones dinámicas).
- Los gráficos pueden necesitar un pequeño ajuste de tamaño al abrir.
- Equipos de eliminatorias son placeholders hasta que se definan los cruces reales.
