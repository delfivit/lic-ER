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
| `Licitaciones Entre Ríos.xlsx` | El resultado, ordenado por estado y con filtros |
| `licitaciones.csv` | Para que la Google Sheet lo lea con `=IMPORTDATA(...)` |
| `fuentes.json` | Prender/apagar fuentes y agregar nuevas. **Este es el que vas a tocar** |
| `estado.json` | Qué licitaciones ya viste (evita avisarte dos veces). No lo borres |
| `.env` | Tus credenciales de mail. **Nunca se sube a GitHub** |

## Agregar una fuente nueva

Editá `fuentes.json` y sumá una entrada. El `parser` puede ser:

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
