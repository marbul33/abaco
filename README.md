# Abaco 
**Gestor digital de efectivo para oficinas bancarias**

Desarrollado para resolver un problema real detectado en mi trabajo 
como Gestor Interno en Banco Sabadell: el arqueo diario de monedas y 
bolsas de efectivo se realizaba íntegramente a mano, con papel y boli, 
o excel, generando errores y perdiendo trazabilidad. Abaco digitaliza y 
centraliza ese proceso.

---

## El problema que resuelve

Cada día, los gestores de oficina deben:
- Llevar el recuento de monedas sueltas por denominación
- Registrar las bolsas de efectivo que entregan los clientes
- Agrupar esas bolsas en lotes de envío con fecha
- Tener siempre visible el total de efectivo en caja

Todo esto se hacía mayormente con papel. Abaco lo hace en segundos, sin errores 
de cálculo y con historial persistente.

## Funcionalidades

| Módulo | Descripción |
|---|---|
| **Monedas** | Saldo por denominación (0,01€–2€) con operaciones +/- en tiempo real |
| **Bolsas** | Registro de sobres con importe, contrato y NIF del cliente |
| **Envíos** | Agrupación en lotes por fecha con historial y reversión |
| **Snapshot diario** | Fotografía del estado de caja al inicio de cada jornada |
| **Arqueo en vivo** | Total de monedas + bolsas + gran total siempre actualizado |
| **Exportación** | Historial completo a CSV compatible con Excel |

## Stack técnico

- **Python 3.10+** — sin dependencias externas
- **Tkinter** — interfaz gráfica de escritorio (stdlib)
- **SQLite3** — persistencia local con tres tablas relacionadas (stdlib)
- **PyInstaller** — empaquetado como .exe para distribución interna

## Arquitectura

El proyecto separa explícitamente la capa de datos de la capa de 
interfaz:

**database.py**  →  CRUD sobre SQLite (monedero, bolsas, snapshots)
**ui.py**        →  Interfaz Tkinter, lógica de negocio y validaciones

`database.py` no importa tkinter y no conoce la existencia de la UI. 
Esto permite sustituir la capa de datos o la interfaz de forma 
independiente.

## Decisiones técnicas destacadas

- **UPSERT nativo de SQLite** (`INSERT ... ON CONFLICT DO UPDATE`) 
  para gestionar saldos sin race conditions
- **Aritmética en céntimos** (multiplicar por 100 y operar con 
  enteros) para evitar errores de punto flotante en cálculos monetarios
- **Snapshot idempotente**: la fotografía diaria se captura solo una 
  vez por jornada, independientemente de cuántas veces se reinicie el 
  programa
- **Bloqueo de instancia única** vía socket en localhost, sin 
  dependencias de registro del sistema
- **Validación en origen**: cada campo Entry usa `validatecommand` 
  de Tkinter para bloquear caracteres inválidos tecla a tecla

## Instalación

```bash
git clone https://github.com/marbul33/abaco.git
cd abaco
python src/ui.py
```

La base de datos se crea automáticamente en el primer arranque en 
`%LOCALAPPDATA%\Abacus_Data\arqueo_local.db`.

No se requiere ningún `pip install`.

## Capturas

![Pestaña Monedas](docs/screenshots/abaco_1.png)
![Pestaña Bolsas IN-FIJO](docs/screenshots/abaco_2.png)
![Pestaña Bolsas HISTORIAL](docs/screenshots/abaco_3.png)
---

*Proyecto personal desarrollado para uso formativo y académico. Todos los datos son
fictícios y simulados. En ningún caso se han usado datos reales.*
