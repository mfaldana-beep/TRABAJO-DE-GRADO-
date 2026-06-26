# Tesis de grado — IA en la administración de justicia con enfoque garantista del derecho penal

Trabajo de grado de pregrado en Derecho — **Universidad Colegio Mayor de Cundinamarca**, Facultad de Derecho.

**Título:** La inteligencia artificial en la administración de justicia con enfoque garantista de los principios del derecho penal. Hacia un modelo de justicia aumentada: parámetros constitucionales y protocolo de transparencia algorítmica para el proceso penal colombiano.

**Autores:** María Fernanda Aldana Rojas · Jesús Esneider Prieto Soto
**Director:** Dr. Daniel Ríos Sarmiento
**Año:** 2026

## Contenido del repositorio

| Archivo | Descripción |
|---|---|
| `Tesis_IA_Justicia_Penal_Garantista.md` | Documento fuente de la tesis (Markdown, normas APA 7). |
| `Tesis_IA_Justicia_Penal_Garantista.docx` | Versión en Word generada a partir del Markdown. |
| `build_docx.py` | Conversor Markdown → DOCX en Python puro (sin dependencias externas). |

## Estructura de la tesis

1. Preliminares (portada, resumen/abstract, tabla de contenido).
2. Introducción.
3. Aspectos generales (problema, justificación, objetivos, hipótesis, variables, metodología).
4. Capítulo 1 — Marco normativo, jurisprudencial y doctrinal.
5. Capítulo 2 — Riesgos de alucinaciones algorítmicas y sesgos.
6. Capítulo 3 — Derecho comparado (UE, EE. UU., Argentina, Brasil).
7. Capítulo 4 — Propuesta: justicia aumentada, test escalonado y Protocolo de Transparencia Algorítmica y Garantías Penales (PTAG).
8. Conclusiones, recomendaciones, impacto esperado, cronograma, referencias y anexos.

## Cómo regenerar el archivo Word

```bash
python3 build_docx.py
```

Genera `Tesis_IA_Justicia_Penal_Garantista.docx` aplicando formato APA (Times New Roman 12, interlineado doble, márgenes de 2,54 cm, numeración de página y tablas con bordes).
