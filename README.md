# HABIC

**Herramienta de Análisis Basada en IA para la Codificación**

> Proyecto en desarrollo.

HABIC es una aplicación de escritorio para hacer análisis de contenido con modelos de lenguaje. Toma un conjunto de textos, por ejemplo publicaciones de actores políticos en X, fragmentos de entrevistas o respuestas abiertas de encuestas, y los codifica de forma binaria (1 o 0) según las categorías de un libro de códigos definido por el investigador. Después permite medir qué tan confiable es esa codificación comparándola con la de un codificador humano.

Nació como herramienta de apoyo para investigaciones sobre discurso político en el Semillero de Investigación en Ciencia Política (SIPOL) de la Universidad Pontificia Bolivariana, Seccional Bucaramanga.

## Qué hace

**1. Recolecta o carga el corpus**
- **Twitter / X:** descarga las publicaciones de una cuenta en un rango de fechas.
- **Cargar Excel:** analiza cualquier corpus propio a partir de una plantilla con tres columnas (`id`, `texto`, `origen`).

**2. Codifica con un libro de códigos**
- Cada categoría se define en Excel con su código, nombre, descripción, criterios de inclusión, ejemplos positivos y exclusiones.
- La codificación se hace por lotes con modelos locales a través de [Ollama](https://ollama.com) o con Claude a través de Claude Code.
- Se pueden hacer varias rondas con distintos modelos sobre el mismo corpus.
- Cuando el modelo no puede decidir con los criterios dados, el texto queda marcado como **N/C (no concluyente)** en vez de forzar un 0 o un 1. Un error de codificación cuenta como dato faltante, no como 0.

**3. Mide la confiabilidad**
- Calcula el tamaño de la muestra de validación con la fórmula de **Lacy y Riffe (1996)**, con un 95 % de confianza.
- Genera los archivos para que un codificador humano codifique esa muestra.
- Calcula el **kappa de Cohen** entre rondas, o entre el modelo y el codificador humano, por categoría, junto con su interpretación.

Todos los resultados se exportan a Excel.

## Requisitos

- Windows con Python 3.10 o superior (la interfaz usa Tkinter, que viene con Python).
- Google Chrome, solo para la descarga desde X.
- [Ollama](https://ollama.com) con al menos un modelo instalado (por ejemplo `llama3.1:8b`), o [Claude Code](https://claude.com/claude-code) si se quiere codificar con Claude.

## Instalación y uso

```bash
pip install -r requirements.txt
python habic.py
```

Desde la aplicación se pueden descargar las plantillas del libro de códigos y del corpus. También hay ejemplos en la carpeta `plantillas/`. Las pruebas automáticas están en `test_habic.py`.

## Aviso sobre el uso de datos de X

La descarga desde X se hace automatizando un navegador con la sesión del propio usuario. Los términos de servicio de X restringen la recolección automatizada, así que cada persona es responsable del uso que le dé a esta función y de los datos que obtenga. Este repositorio no incluye datos recolectados.

## Sobre el desarrollo

HABIC fue construido con apoyo de inteligencia artificial (Claude Code). El diseño metodológico, las decisiones de codificación y las pruebas son del autor.

## Autor

Gabriel Santiago Rodriguez Rojas, estudiante de Ciencia Política y Gobierno, Universidad Pontificia Bolivariana, Seccional Bucaramanga.

## Licencia

[MIT](LICENSE)
