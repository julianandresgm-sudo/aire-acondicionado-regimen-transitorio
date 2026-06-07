# Aire Acondicionado - Régimen Transitorio

Repositorio para el desarrollo y análisis de un sistema de adquisición de datos de temperatura con Arduino para el estudio de régimen transitorio en sistemas de climatización.

## Contenido

### Firmware Arduino
- **`Master_v3_PLX_DAQ.ino`** — Sketch para Arduino Mega 2560 que lee 12 termocuplas (6 locales vía MAX6675 + 6 remotas vía I2C) y envía los datos por serial en formato PLX-DAQ para captura en Excel.

### Scripts de adquisición en vivo
- **`excel_en_vivo.py`** — Lector serial en tiempo real que conecta el Arduino (COM8) con Excel mediante `win32com`, con reconexión automática, auto-guardado y mutex para instancia única.

### Generación de informes
- **`build_informe_v4.py`** — Genera automáticamente un informe técnico en formato Word (.docx) a partir de datos de temperatura (.xlsx), incluyendo tablas, gráficos y análisis estadístico para ensayos de aislamiento térmico (Rubatex APA).

### Sesiones
- **`sesion_arduino_excel.json`** — Exportación de la sesión de opencode utilizada para desarrollar y depurar todo el sistema, con 472 mensajes y ~850K tokens de contexto.

## Licencia

MIT
