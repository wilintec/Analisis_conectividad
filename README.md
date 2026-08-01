# PowerPoint maestro → sitio web en GitHub Pages

Este repositorio convierte automáticamente el **PowerPoint maestro de conectividad** en un sitio HTML responsive y lo publica en GitHub Pages.

## Cómo funciona

1. El archivo `presentation/Analisis_de_la_conectividad_documento_maestro.pptx` es la fuente principal.
2. `scripts/inspect_pptx.py` genera un inventario CSV/JSON de las diapositivas.
3. `scripts/build_site.py` extrae el título, la técnica estadística, la pregunta de investigación y el texto de cada diapositiva.
4. LibreOffice convierte el PowerPoint en PDF y Poppler renderiza cada página.
5. Python genera imágenes WebP, miniaturas, el HTML, el buscador y el manifiesto.
6. `scripts/validate_site.py` comprueba que no falten diapositivas, imágenes ni enlaces locales.
7. GitHub Actions publica automáticamente la carpeta `dist` en GitHub Pages.

El sitio preserva el estilo visual del PowerPoint porque cada diapositiva se renderiza como imagen; simultáneamente, el texto se extrae para navegación, búsqueda y accesibilidad.

## Publicación inicial en GitHub

1. Cree un repositorio nuevo y suba todo el contenido de esta carpeta.
2. Use la rama `main`.
3. En el repositorio, abra **Settings → Pages**.
4. En **Build and deployment → Source**, seleccione **GitHub Actions**.
5. Haga un `push` o ejecute manualmente el workflow desde la pestaña **Actions**.

## Flujo normal de actualización

Cada vez que agregue o modifique diapositivas:

1. Reemplace el archivo `.pptx` dentro de `presentation/` conservando el mismo nombre.
2. Actualice los rangos de capítulos en `site.yml` cuando cambie el orden o el número de diapositivas.
3. Haga `commit` y `push`.
4. GitHub Actions reconstruirá y publicará la página.

## Estructura

```text
.
├── .github/workflows/deploy-pages.yml
├── presentation/
│   └── Analisis_de_la_conectividad_documento_maestro.pptx
├── scripts/
│   ├── build_site.py
│   ├── inspect_pptx.py
│   └── validate_site.py
├── static/
│   ├── app.js
│   └── styles.css
├── templates/
│   └── index.html.j2
├── site.yml
├── requirements.txt
└── README.md
```

## Prueba local en Ubuntu

```bash
sudo apt-get update
sudo apt-get install -y libreoffice-impress poppler-utils fonts-liberation fonts-dejavu-core fonts-noto-core
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python scripts/inspect_pptx.py --config site.yml
python scripts/build_site.py --config site.yml --output dist
python scripts/validate_site.py --site dist
python -m http.server 8000 --directory dist
```

Abra `http://localhost:8000`.

## Prueba local en Windows

La ejecución más sencilla es usar **WSL Ubuntu** con los comandos anteriores. También puede instalar LibreOffice y Poppler para Windows y asegurarse de que `soffice` y `pdftoppm` estén disponibles en `PATH`.

## Convenciones de contenido

Para que la extracción sea precisa, mantenga en las diapositivas los rótulos:

- `Técnica estadística:`
- `PREGUNTA DE INVESTIGACIÓN`

El script usa esos rótulos para estructurar el contenido web.

## Personalización

Edite `site.yml` para cambiar:

- título, autor, institución y colores;
- ubicación y nombre descargable del PowerPoint;
- calidad y resolución de las imágenes;
- capítulos y rangos de diapositivas;
- títulos, preguntas, técnicas o resúmenes específicos mediante `slide_overrides`.

### Visualizaciones interactivas

Copie un HTML interactivo dentro de `static/embeds/` y añada en `site.yml`:

```yaml
slide_overrides:
  17:
    embed: "assets/embeds/mapa_interactivo.html"
```

El iframe se incorporará debajo de la diapositiva correspondiente.

## Nota sobre fidelidad tipográfica

El workflow usa LibreOffice en Linux. Si el PowerPoint contiene fuentes propietarias no disponibles en el runner, LibreOffice puede sustituirlas. Para minimizar diferencias se instalaron Liberation, DejaVu y Noto. Conviene usar fuentes comunes o incluir una exportación PDF de referencia en una evolución posterior del flujo.
