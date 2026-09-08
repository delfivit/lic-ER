# Licitaciones Entre Ríos — Granss SRL

Busca licitaciones de obra pública en Entre Ríos y se queda solo con tus rubros:
**Arquitectura · Infraestructura · Áridos · Vehículos**.

## Correrlo a mano

```bash
./correr.sh                  # baja todo, actualiza el Excel y manda el mail
./correr.sh --sin-mail       # sin mandar mail
./correr.sh --solo Boletin   # una sola fuente, para probar
./correr.sh --todo           # incluye las vencidas
```

## Archivos

| Archivo | Para qué |
|---|---|
| `Licitaciones-Entre-Rios.xlsx` | El resultado, ordenado por estado y con filtros |
| `licitaciones.csv` | Para que la Google Sheet lo lea con `=IMPORTDATA(...)` |
| `fuentes.json` | Prender/apagar fuentes y agregar nuevas. **Este es el que vas a tocar** |
| `estado.json` | Qué licitaciones ya viste (evita avisarte dos veces). No lo borres |
| `diagnostico.csv` | Cómo le fue a cada fuente. Lo muestra la hoja Fuentes de la planilla |
| `.env` | Tus credenciales de mail. **Nunca se sube a GitHub** |

## Agregar una fuente nueva

**Desde la planilla** (lo normal): hoja **Fuentes**, escribí una fila con
Activa=SI, Nombre, URL y Parser. El scraper lee esa hoja en cada corrida, así
que al día siguiente ya la busca. No hace falta tocar código.

**Desde el archivo**: editá `fuentes.json`. Si la misma URL está en los dos
lados, manda la planilla.

El `parser` puede ser:

- `wpjson` — cualquier sitio hecho con WordPress. **Probá este primero**, es el que más rinde
- `generico` — sitios sin API; trae de más y filtra por rubro
- `boletin` — el Boletín Oficial (ya está)
- `parana`, `minplan`, `iapv`, `enersa` — parsers a medida

## Fuentes que NO se pueden leer

- **Concordia** (`compras.concordia.gob.ar`) — carga todo por JavaScript
- **Portal de Transparencia provincial** — ídem
- **Chajarí** — el listado viene de un iframe que además necesita JavaScript

Las tres igual te llegan por el **Boletín Oficial**, que es la fuente más completa:
por ley publican ahí todos los organismos y municipios, incluidas las comunas
chicas sin página web (Villa Elisa, Seguí, Colonia Baylina, Distrito Sauce...).
Es además la única fuente que trae **venta y valor del pliego**.


## Si una fuente falla

La hoja **Fuentes** de la planilla muestra el estado de cada una:

- **verde** anduvo bien
- **rojo** dio error
- **naranja** hace 3 corridas o más que no trae nada (puede haber cambiado el sitio)
- **gris** apagada a propósito (dice el motivo)
- **celeste** la agregó el equipo y todavía no se probó

Cuando alguna queda en rojo o naranja, el Resumen lo avisa arriba de todo y
llega por mail. Mientras tanto conviene entrar a esa web a mano para no
perderse licitaciones. Además, si una fuente falla, sus licitaciones de la
corrida anterior **no se borran** de la planilla.
