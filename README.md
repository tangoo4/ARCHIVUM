# ARCHIVUM

> Sistema de gestión para el control de temporadas de encuadernación.

## Objetivo
ARCHIVUM nace para reducir el tiempo empleado en la gestión de temporadas, eliminar errores derivados del trabajo manual sobre hojas de cálculo y centralizar la información entre Encuadernación y Taller.

No pretende sustituir el Excel.

Pretende hacerlo más inteligente.

## Estado del proyecto
**Versión actual:** v1.0.0  
**Estado:** 🟢 Estable  
**Próxima versión:** Por definir

## Tecnologías
- Python 3.13
- CustomTkinter
- OpenPyXL
- JSON
- PyInstaller
- Git / GitHub

## Funcionalidades v1.0.0
- Nueva temporada y continuación de temporadas existentes.
- Gestión de Medido y modificación de tomos.
- Medidas numéricas y estado pendiente mediante `?`.
- Añadir foliado y cálculo automático del número de hojas.
- Control integrado, utilizando Medido como fuente de verdad.
- Buscador de matrices por rango.
- Cierre de temporada sin bloquear modificaciones posteriores.
- Respaldos automáticos cada cinco tomos y respaldo FINAL.
- Nomenclatura automática de temporadas y archivos.
- Crecimiento dinámico del Excel.
- Validaciones de Año y Medida.
- Estados visuales normalizados.
- Vista completa del Excel integrada.
- Distribución portable mediante ejecutable.

## Estructura documental
- 00 - Visión del Proyecto
- 01 - Arquitectura
- 02 - Flujo de Trabajo
- 03 - Normas de Desarrollo
- 04 - Roadmap
- 05 - Backlog
- 06 - Manual de Usuario
- 07 - Especificación Funcional
- 08 - Registro de Versiones

## Roadmap
### v1.0.0 — Completada
La primera versión estable ha completado su desarrollo funcional, auditoría y pruebas de aceptación.

Las nuevas funcionalidades quedan fuera del alcance de v1.0.0 y deberán evaluarse desde el Backlog para versiones posteriores.

## Filosofía
> **Pensar menos. Trabajar mejor.**

Cada funcionalidad debe reducir tiempo, errores o trabajo innecesario sin complicar la experiencia del usuario.
