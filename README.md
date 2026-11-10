# Generador de Designaciones Docentes

Automatización de descarga y organización de designaciones docentes por sede, facultad y carrera mediante Playwright y Excel.

## 🚀 Requisitos

- Python 3.9 o superior
- Navegador Chromium para Playwright

## 📦 Instalación

1. Clonar el repositorio:
   ```bash
   git clone https://github.com/brada-hub/GENERADOR-DESIGNACIONES.git
   cd GENERADOR-DESIGNACIONES
   ```

2. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```

3. Configurar credenciales:
   Copia el archivo `.env.example` como `.env` y coloca tus credenciales:
   ```bash
   cp .env.example .env
   ```
   Edita `.env` con:
   ```ini
   SISTEMA_USER=tu_usuario
   SISTEMA_PASS=tu_contraseña
   ```

## 💻 Uso

Asegúrate de tener el archivo `DESIGNACIONES.xlsx` en la raíz del proyecto y ejecuta:

```bash
python automatizar_designaciones.py
```
