# SLD Generator - Web App (sld-auto)

Browser-based version of the A176LAB DC Single Line Diagram Generator.
Upload your Excel cable schedule, click Generate, and download the DXF - no Python installation required.

---

## Quick Start (Docker)

### 1. Place your template DXF

```
webapp/
+-- template/
    +-- template.dxf    <-- put your template DXF here
```

### 2. Build and run

```bash
cd webapp
docker-compose up --build -d
```

### 3. Open in browser

```
http://localhost:8080
```

---

## Template Configuration

The template DXF is **never uploaded** - it lives on the server and is mounted read-only into the container.

**Default path inside the container:** `/app/template/template.dxf`

To use a different filename or path, set the `TEMPLATE_DXF` environment variable in `docker-compose.yml`:

```yaml
environment:
  - TEMPLATE_DXF=/app/template/my_custom_template.dxf
```

---

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `TEMPLATE_DXF` | `/app/template/template.dxf` | Absolute path to the template DXF inside the container |
| `PYTHONUNBUFFERED` | `1` | Stream logs immediately (do not change) |

---

## Ports

| Host port | Container port | Service |
|---|---|---|
| `8080` | `5000` | Flask web application |

Change the host port in `docker-compose.yml` if 8080 is already in use:

```yaml
ports:
  - "9000:5000"   # use http://localhost:9000
```

---

## Run Without Docker (development)

```bash
cd webapp
pip install -r requirements.txt
set TEMPLATE_DXF=C:\path\to\template.dxf   # Windows
python app.py
```

Open `http://localhost:5000`.

---

## UI Sections

| Section | Fields |
|---|---|
| 1 - Input Files | Excel cable schedule (.xlsx / .xls) |
| 2 - PV System | Panel model, panels/string, inverter model, DC/AC power, temperature, transformer power |
| 3 - Display Options | Show cable lengths, hide string details, show annotation circle |
| 4 - Advanced | Column/row spacing, circle radius, text size, heavy cable styling |

Selecting an inverter model from the dropdown auto-fills the DC and AC power fields.

---

## API Reference

| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/` | Main form |
| `POST` | `/generate` | Start a generation job. Accepts `multipart/form-data`. Returns `{"job_id": "<uuid>"}` |
| `GET` | `/stream/<job_id>` | Server-Sent Events stream of log messages |
| `GET` | `/download/<job_id>` | Download the generated DXF (triggers job cleanup after 10 s) |
| `GET` | `/status/<job_id>` | JSON status: `{"status": "running|done|error", "error": "..."}` |

### SSE message types

| `type` field | Meaning |
|---|---|
| `log` | Generation log line (`msg` field contains text) |
| `done` | Generation completed successfully |
| `error` | Generation failed (`msg` contains the error) |
| `heartbeat` | Keep-alive ping (every 30 s of inactivity) |
| `sentinel` | Stream closing |

---

## Security Notes

- Only `.xlsx` and `.xls` files are accepted for upload (validated server-side)
- Job IDs are validated as strict UUID v4 before any lookup
- Uploaded files and generated DXF are deleted automatically 10 seconds after download
- Jobs older than 2 hours are purged by a background cleanup thread
- Response headers include `X-Content-Type-Options: nosniff` and `X-Frame-Options: DENY`
- The template DXF is mounted read-only and its path is never exposed in error responses

---

## File Structure

```
webapp/
+-- app.py                  # Flask application
+-- sld_core.py             # Loads generation engine via importlib
+-- requirements.txt        # Python dependencies
+-- Dockerfile              # Python 3.11-slim image
+-- docker-compose.yml      # Service: sld-auto, port 8080
+-- .dockerignore
+-- templates/
|   +-- index.html          # Single-page browser UI (Tailwind CSS)
+-- code/
|   +-- smart_sld_generator.py   # Generation engine (copy from /code/)
+-- template/
    +-- README.txt          # Instructions
    +-- template.dxf        # Your template DXF (gitignored, mounted as volume)
```

---

## Stopping and Removing

```bash
docker-compose down          # stop and remove container
docker-compose down -v       # also remove named volumes
docker rmi webapp-sld-auto   # remove the image
```
