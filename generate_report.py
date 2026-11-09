import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_report():
    doc = docx.Document()

    # Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_run = f_p.add_run("Informe de Autoevaluación de Calidad — Ya Quedó | Seminario de Calidad")
        f_run.font.name = "Calibri"
        f_run.font.size = Pt(8.5)
        f_run.font.color.rgb = RGBColor(160, 174, 192)

    def set_cell_background(cell, fill_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        tcPr.append(shd)

    def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def style_table(table, header_bg="1A365D", alt_bg="F7FAFC", col_widths=None):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            if i == 0:
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            for j, cell in enumerate(row.cells):
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
                if col_widths and j < len(col_widths):
                    cell.width = col_widths[j]
                if i == 0:
                    set_cell_background(cell, header_bg)
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for r in p.runs:
                            r.font.bold = True
                            r.font.color.rgb = RGBColor(255, 255, 255)
                            r.font.name = "Calibri"
                            r.font.size = Pt(9)
                else:
                    if i % 2 == 1:
                        set_cell_background(cell, "FFFFFF")
                    else:
                        set_cell_background(cell, alt_bg)
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.font.name = "Calibri"
                            r.font.size = Pt(8.5)
                            r.font.color.rgb = RGBColor(45, 55, 72)

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Calibri"
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(26, 54, 93)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Calibri"
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(43, 108, 176)
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(74, 85, 104)
        return p

    def add_p(text="", bold_prefix="", space_after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            r_bold = p.add_run(bold_prefix)
            r_bold.font.name = "Calibri"
            r_bold.font.size = Pt(9.5)
            r_bold.font.bold = True
            r_bold.font.color.rgb = RGBColor(26, 32, 44)
        if text:
            r_text = p.add_run(text)
            r_text.font.name = "Calibri"
            r_text.font.size = Pt(9.5)
            r_text.font.color.rgb = RGBColor(45, 55, 72)
        return p

    def add_bullet(text="", bold_prefix=""):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            r_bold = p.add_run(bold_prefix)
            r_bold.font.name = "Calibri"
            r_bold.font.size = Pt(9.5)
            r_bold.font.bold = True
            r_bold.font.color.rgb = RGBColor(26, 32, 44)
        if text:
            r_text = p.add_run(text)
            r_text.font.name = "Calibri"
            r_text.font.size = Pt(9.5)
            r_text.font.color.rgb = RGBColor(45, 55, 72)
        return p

    def add_callout(text, title=""):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        set_cell_background(cell, "EDF2F7")
        set_cell_margins(cell, top=80, bottom=80, left=120, right=100)
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="0" w:color="3182CE"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
        tcPr.append(borders)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        if title:
            rt = p.add_run(title + "\n")
            rt.font.bold = True
            rt.font.size = Pt(9.5)
            rt.font.color.rgb = RGBColor(43, 108, 176)
        rtxt = p.add_run(text)
        rtxt.font.size = Pt(9)
        rtxt.font.color.rgb = RGBColor(45, 55, 72)
        p_spacer = doc.add_paragraph()
        p_spacer.paragraph_format.space_after = Pt(2)

    # ---------------------------------------------------------
    # 1. CARÁTULA E IDENTIFICACIÓN
    # ---------------------------------------------------------
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(36)
    p_title.paragraph_format.space_after = Pt(12)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_title.add_run("INFORME DE AUTOEVALUACIÓN DE CALIDAD DE SOFTWARE")
    r_t.font.name = "Calibri"
    r_t.font.size = Pt(22)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(26, 54, 93)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(28)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Auditoría Técnica Integral, Modelo ISO/IEC 25010 y Plan de Aseguramiento de Calidad (SQA)\nProyecto de Seminario: Plataforma 'Ya Quedó' (Marketplace de Servicios para el Hogar)")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(74, 85, 104)

    t_meta = doc.add_table(rows=7, cols=2)
    style_table(t_meta, header_bg="2B6CB0", alt_bg="F7FAFC", col_widths=[Inches(2.5), Inches(4.0)])
    meta_data = [
        ("Campo", "Detalle de Identificación"),
        ("Nombre del Proyecto:", "Ya Quedó — Marketplace de Servicios para el Hogar"),
        ("Equipo Evaluador / QA:", "Equipo de Calidad y Desarrollo de Seminario (Auditores SQA)"),
        ("Curso / Cátedra:", "Aseguramiento de la Calidad de Software — Proyecto Final"),
        ("Fecha de Evaluación:", "Ciclo Académico 2026 (Auditoría Inicial: Sep 2026 | Consolidación y Verificación: Nov 2026)"),
        ("Versión Evaluada:", "v1.0.0-beta (Rama git: chore/pin-vite-8.3.4, Commit HEAD local)"),
        ("Stack Tecnológico:", "Frontend: Ionic 8 + React 18 + TypeScript + Vite + Capacitor\nBackend: Supabase self-hosted (PostgreSQL 15 con RLS, GoTrue, Realtime, Storage) + Edge Functions (Deno)")
    ]
    for row_idx, (c1, c2) in enumerate(meta_data):
        row = t_meta.rows[row_idx]
        row.cells[0].paragraphs[0].text = c1
        row.cells[1].paragraphs[0].text = c2
        if row_idx > 0:
            row.cells[0].paragraphs[0].runs[0].font.bold = True

    add_callout("Este documento constituye el informe formal de autoevaluación de calidad sobre el sistema de software 'Ya Quedó'. Su propósito es evidenciar el estado real de la plataforma frente a las expectativas teóricas, aplicando rigurosamente los estándares ISO/IEC 25010, técnicas formales de pruebas (caja negra, gris y blanca), gestión de defectos y riesgos, y el ciclo de mejora continua PHVA.", "Propósito del Informe:")

    doc.add_page_break()

    # ---------------------------------------------------------
    # 2. RESUMEN EJECUTIVO
    # ---------------------------------------------------------
    add_h1("2. Resumen Ejecutivo")
    add_p("El presente informe documenta los resultados de la autoevaluación formal y exhaustiva de la calidad del software de la plataforma 'Ya Quedó', un marketplace para la contratación de servicios para el hogar. La evaluación se ejecutó contrastando la percepción inicial del equipo de desarrollo con la evidencia empírica recolectada a través de auditorías de caja blanca (revisión de código e inspección de esquemas SQL), caja gris (pruebas directas sobre la API REST de PostgREST, endpoints de GoTrue y Edge Functions) y caja negra (ejecución de pruebas E2E funcionales, análisis de accesibilidad con axe-core y renderizado responsive).")

    add_h2("2.1 Conclusión General del Estado de Calidad")
    add_p("El proyecto presenta una base arquitectónica funcional sólida en el front-office y una interfaz visual moderna basada en Ionic y React. No obstante, existe una brecha crítica y severa entre la percepción de seguridad y robustez del equipo frente a la realidad del backend. La plataforma exhibe vulnerabilidades de alto impacto en las capas de autorización a nivel de fila (Row Level Security - RLS), exposición de secretos y endpoints administrativos, junto con defectos críticos de accesibilidad (WCAG 2.2 AA) y diseño responsivo en pantallas móviles estrechas (360 px). La calidad global actual se califica con 61/100, ubicando al sistema en una fase pre-operativa no apta para entornos de producción hasta completar las remediaciones prioritarias.")

    add_h2("2.2 Los Tres Hallazgos Principales")
    add_bullet(" En la configuración inicial se mantuvo la clave JWT_SECRET por defecto de Supabase ('supabase-demo'), permitiendo a cualquier actor forjar tokens con rol service_role y eludir todas las políticas de seguridad. Adicionalmente, las políticas RLS iniciales de bookings permitían que un cliente marcara unilateralmente un servicio como completado, alterara el precio acordado a 1 y reasignara profesionales. Asimismo, los clientes podían modificar su propio rol en profiles hacia 'profesional' debido a políticas UPDATE sin restricción de columnas.", "1. Vulnerabilidades Críticas de Seguridad y RLS (Brecha de Seguridad: +2.4):")
    add_bullet(" El módulo central de reservas (Bookings: solicitar, aceptar, rechazar, completar y cancelar) no se encontraba integrado en la interfaz gráfica del usuario (manteniéndose como un componente placeholder 'Fase 5'). La lógica residía únicamente en la base de datos sin pruebas de aceptación visual para usuarios finales ni cobertura automatizada E2E en Cypress (suite e2e vacía en el frontend).", "2. Desconexión entre Lógica de Negocio Backend y Frontend:")
    add_bullet(" En dispositivos de 360 px (ancho de pantalla estándar en gama de entrada Android), el texto del hero de la landing page se corta por overflow oculto y las pestañas de Login ('Iniciar sesión / Crear cuenta') se truncan severamente. La accesibilidad automática arrojó entre 42 y 57 violaciones serias de 'role-img-alt' por pantalla debido a la ausencia de aria-hidden en componentes IonIcon.", "3. Bloqueos Críticos de Accesibilidad y Responsive Design (Brecha de Usabilidad: +1.6):")

    add_h2("2.3 Las Tres Recomendaciones Clave")
    add_bullet(" Generar un JWT_SECRET criptográficamente seguro de 64 caracteres aleatorios fuera del repositorio; aislar las claves maestras; implementar triggers BEFORE UPDATE que congelen la inmutabilidad de client_id, professional_id y service_id; revocar permisos directos de UPDATE sobre profiles.role a usuarios autenticados; y separar vistas públicas que enmascaren la ubicación geográfica y documentos privados de verificación (DPI).", "1. Blindaje Inmediato del Backend y Autenticación:")
    add_bullet(" Construir la interfaz de usuario reactiva para el ciclo de vida de bookings soportada en los disparadores SQL ya corregidos (fix-bookings-state-machine.sql). Poblar la suite Cypress con pruebas E2E determinísticas que validen los flujos de login, solicitud, cancelación y reseñas.", "2. Cierre del Ciclo Funcional y Automatización de Pruebas:")
    add_bullet(" Eliminar la restricción de zoom en meta-viewport (user-scalable=no, maximum-scale=1); corregir estilos CSS en tarjetas para viewports de 360 px; añadir aria-hidden='true' en iconos decorativos; y formalizar el Definition of Done (DoD) integrando revisiones de código obligatorias por pares y linters automáticos.", "3. Remediación de Accesibilidad, Responsive y Cultura de Calidad:")

    # ---------------------------------------------------------
    # 3. DESCRIPCIÓN DEL PROYECTO Y ALCANCE
    # ---------------------------------------------------------
    add_h1("3. Descripción del Proyecto y Alcance")
    add_p("La plataforma 'Ya Quedó' fue concebida para conectar de forma ágil, segura y transparente a prestadores de servicios domésticos (plomería, electricidad, albañilería, limpieza, pintura) con usuarios residenciales que demandan reparaciones y mantenimiento.")

    add_h2("3.1 Arquitectura del Sistema")
    add_p("El software está estructurado sobre un patrón desacoplado y orientado a servicios:")
    add_bullet(" Implementado como Single Page Application (SPA) y compilado para móviles mediante Capacitor. Hace uso de React 18, React Router v5/v6, TypeScript para tipado estático, Vite como empaquetador de alto rendimiento e Ionic Framework para componentes híbridos móviles.", "Frontend (app-servicios):")
    add_bullet(" Instancia de Supabase orquestada en contenedores Docker mediante Envoy Gateway, PostgREST para exponer el esquema PostgreSQL como API RESTful, GoTrue para gestión de identidad y autenticación JWT, Realtime para mensajería y Storage para archivos binarios.", "Backend y Base de Datos (sql_docker):")
    add_bullet(" Funciones Deno que ejecutan operaciones fuera del navegador que requieren secretos privilegiados, específicamente el envío de notificaciones push móviles mediante Firebase Cloud Messaging (FCM).", "Edge Functions (edge_functions/send-push):")

    add_h2("3.2 Alcance de la Evaluación Realizada")
    add_p("Para la presente evaluación de calidad se incluyeron dentro del alcance auditado los siguientes módulos y componentes:")
    add_bullet(" Formulario de login, registro de clientes y profesionales en 3 pasos, recuperación de sesión y cierre por inactividad a los 30 minutos.", "Módulo de Autenticación y Perfiles:")
    add_bullet(" Búsqueda de servicios por texto, visualización de fichas de contacto, cálculo y despliegue de reputación/estrellas, y filtrado por categoría.", "Módulo de Descubrimiento y Feed:")
    add_bullet(" Altas, modificaciones, borrado y listado de servicios prestados por el profesional (MyServices.tsx), gestión de bio y documentos.", "Módulo de Gestión de Servicios:")
    add_bullet(" Análisis caja gris contra la API REST de PostgREST y disparadores de PostgreSQL sobre el ciclo de vida de transiciones de estados (solicitado -> aceptado -> en_curso -> completado / cancelado / rechazado).", "Módulo de Solicitudes (Bookings):")
    add_bullet(" Políticas RLS sobre 8 tablas (profiles, professional_profiles, services, bookings, reviews, messages, notifications, reports), inyección de SQL y bypass de secretos.", "Capa de Seguridad y Autorización:")

    add_h2("3.3 Límites Explícitos y Exclusiones del Alcance")
    add_p("Los siguientes aspectos quedaron explícitamente fuera del alcance de la evaluación técnica por restricciones del estado del repositorio o diseño del producto:")
    add_bullet(" La aplicación no cuenta con pasarela de pagos electrónicos (Stripe/PayPal). Todo acuerdo económico es pactado entre particulares fuera de la app (price_agreed).", "Pasarela de Pagos:")
    add_bullet(" El componente Bookings.tsx en el frontend constituía una maqueta visual durante la primera fase de auditoría; la UI final no pudo evaluarse en esa versión como usuario final convencional, requiriendo pruebas de caja gris a nivel de API.", "UI Nativa de Bookings (Fase Inicial):")
    add_bullet(" No se ejecutaron pruebas de estrés con 10,000 usuarios concurrentes en servidores de nube (OCI) para evitar costos operativos no autorizados, focalizando el análisis de rendimiento en el stack local y mediciones de auditoría estática/Lighthouse.", "Pruebas de Carga a Gran Escala en Producción:")

    # ---------------------------------------------------------
    # 4. METODOLOGÍA DE EVALUACIÓN
    # ---------------------------------------------------------
    add_h1("4. Metodología")
    add_p("La evaluación se desarrolló siguiendo un marco metodológico estructurado en múltiples fases de prueba complementarias, garantizando la cobertura de código, comportamiento de endpoints y experiencia de usuario.")

    add_h2("4.1 Fases y Enfoques de Prueba")
    add_bullet(" Inspección estática profunda de código fuente TypeScript, configuración de ESLint, empaquetado Vite, y análisis de scripts DDL y triggers en PostgreSQL (policies-rls.sql, fix-bookings-state-machine.sql).", "1. Auditoría de Caja Blanca:")
    add_bullet(" Pruebas de integración y seguridad dirigidas hacia la API de PostgREST y GoTrue empleando cURL y Postman (SEMINARIO_OCI.postman_collection.json). Se crearon identidades con tokens JWT reales para evaluar la resistencia de las políticas RLS ante parámetros manipulados (tampering), elevación de privilegios y bypass de filtros.", "2. Auditoría de Caja Gris:")
    add_bullet(" Pruebas dinámicas de usuario ejecutadas en navegadores Chrome y suites de automatización. Validación visual y funcional de los flujos de registro, login, búsqueda y creación de reseñas.", "3. Auditoría de Caja Negra:")

    add_h2("4.2 Herramientas Empleadas")
    add_bullet(" Evaluación automatizada de rendimiento, buenas prácticas, accesibilidad y SEO sobre las vistas web principales.", "Google Lighthouse:")
    add_bullet(" Reglas de análisis estático en IDE para la detección temprana de malas prácticas, tipos inseguros (as unknown) y memory leaks en componentes React.", "SonarLint / ESLint:")
    add_bullet(" Motor de evaluación de accesibilidad web inyectado para auditar el cumplimiento del estándar WCAG 2.2 nivel AA en las 6 pantallas maestras del sistema.", "axe-core v4.10.2:")
    add_bullet(" Framework de pruebas E2E para la simulación de navegación e interacciones de usuario en el frontend híbrido.", "Cypress E2E Testing:")
    add_bullet(" Cliente PostgreSQL para inspección de diccionarios de datos, planes de consulta y disparadores transaccionales.", "psql & Docker CLI:")

    add_h2("4.3 Cronograma y Criterios de Aceptación")
    add_p("Las actividades de aseguramiento de calidad se distribuyeron cronológicamente en: 1) Auditoría de Seguridad Inicial y Revisión Estática (Septiembre 2026); 2) Desarrollo de Scripts de Reparación RLS y Endpoints (Octubre 2026); y 3) Evaluación Dinámica en Chrome, Pruebas de Accesibilidad y Verificación de Regresión (Noviembre 2026).")
    add_p("Criterios de Aceptación: Todo defecto clasificado con severidad Crítica (S1) o Mayor (S2) debe ser resuelto previo a cualquier pase a producción. Ningún endpoint debe exponer datos sensibles sin autenticación ni violar las matrices de transición de estados del negocio.")

    doc.add_page_break()

    # ---------------------------------------------------------
    # 5. DIMENSIÓN 1 — EVALUACIÓN DE REQUISITOS
    # ---------------------------------------------------------
    add_h1("5. Dimensión 1 — Evaluación de Requisitos")
    add_p("Una causa primordial de defectos de software es la ambigüedad y falta de verificabilidad en la definición inicial de los requisitos funcionales. En el proyecto 'Ya Quedó', las historias de usuario originales carecían de criterios de aceptación medibles, omitían el manejo de condiciones de carrera y no definían el comportamiento ante entradas no válidas.")

    add_h2("5.1 Diagnóstico de Requisitos Originales Defectuosos")
    add_bullet(" Expresado originalmente como 'El usuario puede registrarse con su correo y una clave'. No delimitaba la complejidad requerida de la contraseña, el comportamiento ante correos duplicados, el flujo de confirmación ni el mecanismo transaccional de creación de perfil.", "Requisito Original 1 (Registro de Usuarios):")
    add_bullet(" Redactado como 'El cliente puede pedir un servicio a un profesional'. No contemplaba la concurrencia, permitiendo que un cliente solicitara múltiples veces el mismo servicio de manera simultánea o que pactara precios arbitrarios sin aprobación del profesional.", "Requisito Original 2 (Solicitud de Servicios):")
    add_bullet(" Indicaba 'El usuario puede dejar una reseña del servicio'. Omitía verificar si el servicio realmente había sido completado con éxito, no limitaba la longitud del texto permitiendo desbordamientos y no controlaba la edición posterior de comentarios.", "Requisito Original 3 (Sistema de Calificaciones):")
    add_bullet(" Formulaba 'Los profesionales pueden subir sus documentos para que los aprueben'. Omitía el formato, peso máximo permitido, control de almacenamiento en bucket privado y denegación de visualización a terceros no autorizados.", "Requisito Original 4 (Verificación de Perfil):")

    add_h2("5.2 Reescritura Formal de Requisitos (INVEST / Gherkin)")
    add_p("A continuación, se reescriben formalmente 4 requisitos bajo los estándares de claridad, verificabilidad, completitud y formulación Gherkin:")

    add_h3("Requisito Reescrito 1: RF-AUTH-01 — Registro Seguro y Creación de Perfil en Dos Pasos")
    add_p("Como nuevo usuario de la plataforma, quiero registrarme ingresando mi nombre completo, correo electrónico, rol deseado y una contraseña robusta, para acceder a los servicios de forma autenticada y segura.", "Historia de Usuario: ")
    add_bullet("El nombre completo debe tener al menos 2 palabras y cada palabra al menos 2 caracteres alfabéticos.", "Criterio 1:")
    add_bullet("La contraseña debe exigir mínimo 8 caracteres, al menos una mayúscula ASCII, un número y un carácter especial del conjunto permitido.", "Criterio 2:")
    add_bullet("La creación del usuario en auth.users y la fila en public.profiles debe ejecutarse en una transacción atómica mediante trigger en PostgreSQL, impidiendo cuentas huérfanas.", "Criterio 3:")
    add_p("Escenario Gherkin:\nDado que un usuario anónimo ingresa al formulario de registro en paso 1 con 'Ana López' y rol 'Cliente'\nY avanza al paso 2 con correo 'ana.valida@test.local'\nY en paso 3 ingresa la contraseña 'Segura123!' que cumple los requisitos\nCuando presiona el botón 'Crear mi cuenta'\nEntonces el sistema crea la cuenta en el backend, genera su perfil atómicamente y lo redirige a '/tabs/home' autenticado.")

    add_h3("Requisito Reescrito 2: RF-BOOK-02 — Solicitud Única de Servicio con Máquina de Estados Inmutable")
    add_p("Como cliente autenticado, quiero solicitar un servicio activo indicando notas descriptivas, garantizando que no se dupliquen solicitudes pendientes y que los identificadores de la reserva sean inmutables.", "Historia de Usuario: ")
    add_bullet("El cliente no puede tener más de una solicitud en estado activo ('solicitado', 'aceptado', 'en_curso') para el mismo servicio.", "Criterio 1:")
    add_bullet("Los campos client_id, professional_id y service_id quedan congelados tras la inserción inicial.", "Criterio 2:")
    add_bullet("El campo de notas descriptivas debe aceptar un máximo de 2000 caracteres tanto en el cliente como en la restricción CHECK de BD.", "Criterio 3:")
    add_p("Escenario Gherkin:\nDado que la cliente 'María' ya posee una reserva en estado 'aceptado' para el servicio 'Reparación de Fugas'\nCuando María intenta enviar una nueva solicitud sobre el mismo servicio\nEntonces el sistema bloquea la transacción (código 23505) y la interfaz despliega: 'Ya tenés una solicitud activa para este servicio. Podés verla en Solicitudes.'")

    add_h3("Requisito Reescrito 3: RF-REV-03 — Calificación Exclusiva de Servicios Completados con Control de Edición")
    add_p("Como cliente que contrató un servicio, quiero calificar y reseñar el trabajo recibido una vez completado, asegurando la veracidad de la reputación del profesional.", "Historia de Usuario: ")
    add_bullet("Solo se permite registrar una reseña si el booking asociado tiene el estado 'completado'.", "Criterio 1:")
    add_bullet("La calificación debe ser un entero entre 1 y 5 estrellas. El comentario es opcional pero si se provee debe tener entre 1 y 500 caracteres útiles (btrim).", "Criterio 2:")
    add_bullet("Si la reseña se edita, el sistema debe marcarla automáticamente con 'edited=true' únicamente si el texto o calificación han cambiado.", "Criterio 3:")
    add_p("Escenario Gherkin:\nDado que un cliente tiene una reserva con estado 'completado'\nCuando envía una reseña con calificación 5 y un texto de 501 caracteres\nEntonces el componente visual trunca la entrada a 500 caracteres, y la API rechaza cualquier intento de sobrepaso con error 400 check_violation.")

    add_h3("Requisito Reescrito 4: RF-DOC-04 — Carga Segura y Privada de Documentos de Verificación Profesional")
    add_p("Como profesional de servicios, quiero subir fotos legibles de mi DPI o licencias profesionales, para que los administradores verifiquen mi identidad sin que terceros puedan descargar mis documentos personales.", "Historia de Usuario: ")
    add_bullet("Los documentos se almacenan en un bucket privado 'verification-docs' organizado por ruta user_id/archivo.", "Criterio 1:")
    add_bullet("Las políticas de Storage deben conceder acceso SELECT y DELETE únicamente al propietario autenticado y a administradores autorizados.", "Criterio 2:")
    add_bullet("Cualquier intento de acceso anónimo o por otro usuario debe responder HTTP 403 Forbidden.", "Criterio 3:")
    add_p("Escenario Gherkin:\nDado que un profesional autenticado sube su archivo de certificación 'licencia.pdf'\nCuando un usuario ajeno o un actor anónimo intenta descargar el archivo mediante la URL directa de la API de Storage\nEntonces el servidor responde 403 Forbidden y deniega la descarga.")

    doc.add_page_break()

    # ---------------------------------------------------------
    # 6. DIMENSIÓN 2 — CALIDAD SEGÚN ISO/IEC 25010 (PLANTILLA A)
    # ---------------------------------------------------------
    add_h1("6. ISO/IEC 25010: Percepción vs. Evidencia")
    add_p("El estándar internacional ISO/IEC 25010 define un modelo de calidad de producto de software compuesto por 8 características esenciales. Para contrastar la visión preliminar con la madurez real del sistema, cada miembro del equipo registró inicialmente su nota de percepción (escala 1 a 5) previo a las auditorías técnicas. Con posterioridad a las pruebas de caja blanca, gris y negra, se asignó la nota sustentada en evidencia empírica.")
    add_p("Fórmula de brecha: Brecha = Percepción - Evidencia. Una brecha positiva indica sobreestimación teórica del equipo; una brecha cero refleja alineación; y una brecha negativa indica que el producto superó las expectativas iniciales.")

    add_h2("Plantilla A — Percepción vs. Evidencia ISO/IEC 25010")
    t_iso = doc.add_table(rows=9, cols=5)
    style_table(t_iso, header_bg="1A365D", alt_bg="F7FAFC", col_widths=[Inches(1.8), Inches(0.9), Inches(0.9), Inches(0.8), Inches(3.1)])
    iso_data = [
        ("Característica", "Percepción (1-5)", "Evidencia (1-5)", "Brecha", "Evidencia que respalda la nota y recomendación"),
        ("Adecuación funcional", "4.2", "3.0", "+1.2", "Evidencia: El flujo de bookings estaba ausente en la UI de usuarios; las políticas originales no impedían que el cliente auto-completara trabajos ni fijara precios a 1. Recomendación: Desplegar interfaz reactiva para solicitudes y forzar el trigger de máquina de estados."),
        ("Eficiencia de desempeño", "4.0", "3.6", "+0.4", "Evidencia: Las consultas RPC feed_services responden en <120 ms localmente. Sin embargo, en Home.tsx y ProfessionalDetail.tsx hay race conditions por falta de AbortController. Recomendación: Limpiar suscripciones y hooks de fetch."),
        ("Compatibilidad", "4.5", "3.8", "+0.7", "Evidencia: La integración Supabase + Ionic opera correctamente en Android y Web. No obstante, las rutas protegidas sufren problemas con el historial de navegación (bfcache). Recomendación: Revalidar sesión en pageshow."),
        ("Usabilidad", "4.4", "2.8", "+1.6", "Evidencia: En viewports de 360 px el hero de landing se corta y los tabs de Login se truncan. Además, hay entre 42 y 57 violaciones WCAG de role-img-alt por pantalla. Recomendación: Ajustar breakpoints a 360 px y añadir aria-hidden en iconos."),
        ("Fiabilidad", "4.0", "3.2", "+0.8", "Evidencia: La base de datos es robusta, pero la UI silencia errores de RLS (PostgREST responde 200 con 0 filas afectadas) y el signup dejaba usuarios huérfanos sin perfil. Recomendación: Verificar count de filas y transaccionalidad."),
        ("Seguridad", "4.5", "2.1", "+2.4", "Evidencia: Clave JWT_SECRET demo pública en docker-compose; políticas RLS permisivas en professional_profiles exponiendo DPIs; y clientes capaces de cambiar su rol a profesional. Recomendación: Rotar llaves y revocar UPDATE en role."),
        ("Mantenibilidad", "3.8", "3.0", "+0.8", "Evidencia: Lógica duplicada de DiceBear y feed entre páginas; uso de 'as unknown' para forzar tipos; Cypress e2e vacío. Recomendación: Extraer custom hooks tipados y escribir suites de pruebas."),
        ("Portabilidad", "4.5", "4.2", "+0.3", "Evidencia: Capacitor compila y empaqueta de forma transparente para Android (RELEASE_ANDROID.md) y web. Buena adaptabilidad global. Recomendación: Probar en dispositivos físicos iOS y Android de diversas densidades.")
    ]
    for r_idx, row_values in enumerate(iso_data):
        row = t_iso.rows[r_idx]
        for c_idx, val in enumerate(row_values):
            row.cells[c_idx].paragraphs[0].text = val
            if r_idx > 0 and c_idx in (1, 2, 3):
                row.cells[c_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_h2("6.1 Análisis Profundo de las Brechas Más Significativas")
    add_bullet(" Con una brecha de +2.4 puntos, esta área representó el mayor desfase entre expectativa y realidad. El equipo asumía que por utilizar Supabase 'la seguridad venía resuelta de fábrica'. La evidencia técnica demostró que la persistencia de credenciales demo predeterminadas en archivos de despliegue anula todo el control perimetral. Adicionalmente, escribir políticas RLS con using (true) sin discriminar columnas confidenciales expuso documentos personales de profesionales. Esta brecha impone una reestructuración de la arquitectura de autorización antes de operar en redes públicas.", "1. Brecha de Seguridad (+2.4):")
    add_bullet(" Con una brecha de +1.6 puntos, se creía que el uso de los componentes móviles de Ionic garantizaba automáticamente una experiencia accesible y fluida. Sin embargo, las auditorías con axe-core y simulaciones a 360 px mostraron fallas severas de accesibilidad (iconos sin describir, textos a 10.88 px bajo la norma WCAG y saltos de layout). La usabilidad requiere diseño inclusivo activo y no solo librerías de componentes prediseñados.", "2. Brecha de Usabilidad (+1.6):")
    add_bullet(" Con una brecha de +1.2 puntos, la funcionalidad parecía completa al ver el backend terminado, pero desde la óptica del usuario final existía un vacío insalvable al no poder solicitar reservas desde la pantalla de la app.", "3. Brecha de Adecuación Funcional (+1.2):")

    # ---------------------------------------------------------
    # 7. DIMENSIÓN 3 — PRUEBAS DISEÑADAS Y EJECUTADAS
    # ---------------------------------------------------------
    add_h1("7. Pruebas Diseñadas y Ejecutadas")
    add_p("Para validar las funciones más críticas de la plataforma se aplicaron tres técnicas formales de diseño de pruebas: 1) Partición de Clases de Equivalencia, 2) Análisis de Valores Límite / Frontera, y 3) Tablas de Decisión / Transición de Estados. Asimismo, se mapearon las pruebas contra los Cuadrantes de Marick para garantizar una cobertura equilibrada.")

    add_h2("7.1 Mapeo de Cuadrantes de Pruebas (Marick)")
    add_bullet(" Pruebas de disparadores en PostgreSQL (check_violation, validación de transiciones de booking) y validación unitaria de expresiones regulares en validation.ts.", "Cuadrante 1 (Q1 — Tecnología / Apoyo al Equipo):")
    add_bullet(" Casos de prueba funcionales de login, registro en 3 pasos, búsqueda en feed y solicitud de servicio.", "Cuadrante 2 (Q2 — Negocio / Apoyo al Equipo):")
    add_bullet(" Pruebas exploratorias, auditoría de usabilidad a 360 px y evaluación heurística de Nielsen en Chrome.", "Cuadrante 3 (Q3 — Negocio / Crítica al Producto):")
    add_bullet(" Análisis de accesibilidad WCAG con axe-core, pruebas de inyección / bypass de autorización RLS y análisis de seguridad.", "Cuadrante 4 (Q4 — Tecnología / Crítica al Producto):")

    add_h2("7.2 Matriz de Casos de Prueba Críticos Ejecutados")
    t_tests = doc.add_table(rows=6, cols=6)
    style_table(t_tests, header_bg="1A365D", alt_bg="F7FAFC", col_widths=[Inches(0.8), Inches(1.5), Inches(1.4), Inches(1.2), Inches(1.3), Inches(0.8)])
    tests_data = [
        ("ID", "Función / Caso", "Técnica / Cuadrante", "Datos de Entrada", "Resultado Esperado", "Estado"),
        ("TC-001", "Login válido de cliente demo", "Equivalencia / Q2", "maria.cliente@demo.com, Clave demo", "Redirección a /tabs/home, feed cargado, sesión activa", "PASÓ"),
        ("TC-002", "Validación nombre en registro", "Valores Límite / Q2", "1) 'Maria' (1 palabra)\n2) 'María López'", "1) Error en rojo\n2) Avanza a Paso 2", "PASÓ"),
        ("TC-003", "Solicitud activa duplicada", "Tabla Decisión / Q1", "Servicio con solicitud activa en 'aceptado'", "Error 23505 bloqueado; mensaje explicativo en UI", "PASÓ"),
        ("TC-004", "Límite longitud de reseña", "Valores Límite / Q1-Q2", "1) 500 chars\n2) 501 chars", "1) 201 Created\n2) UI trunca a 500; API retorna 400", "PASÓ"),
        ("TC-005", "Gateway rechaza sin apikey", "Seguridad / Q4", "Peticiones a /auth/v1/health sin clave", "HTTP 401 Unauthorized sin fuga de información", "PASÓ")
    ]
    for r_idx, row_values in enumerate(tests_data):
        row = t_tests.rows[r_idx]
        for c_idx, val in enumerate(row_values):
            row.cells[c_idx].paragraphs[0].text = val
            if r_idx > 0 and c_idx in (0, 5):
                row.cells[c_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                if c_idx == 5 and val == "PASÓ":
                    row.cells[c_idx].paragraphs[0].runs[0].font.bold = True
                    row.cells[c_idx].paragraphs[0].runs[0].font.color.rgb = RGBColor(40, 167, 69)

    doc.add_page_break()

    # ---------------------------------------------------------
    # 8. DIMENSIÓN 4 — PROCESO Y CULTURA DE CALIDAD
    # ---------------------------------------------------------
    add_h1("8. Proceso y Cultura de Calidad")
    add_p("La calidad del producto resultante es el reflejo directo de la madurez y disciplina de los procesos de trabajo del equipo. A continuación, se presenta un análisis honesto y transparente de las prácticas internas de ingeniería:")

    add_h2("8.1 Roles y Responsabilidades")
    add_p("En el equipo no existió una división formal de roles de aseguramiento de calidad (QA). Todos los integrantes actuaron prioritariamente como desarrolladores full-stack, lo que generó un 'sesgo de confirmación': quien programa una característica tiende a probarla únicamente bajo el 'camino feliz' (happy path), asumiendo que el usuario se comportará de manera predecible. La falta de un rol dedicado a pruebas destructivas impidió descubrir a tiempo las fallas de inyección en RLS y responsive.")

    add_h2("8.2 Control de Versiones y Revisión de Código entre Pares (Peer Reviews)")
    add_p("Se utilizó Git alojado en GitHub bajo un esquema de ramas por característica (feature branches). No obstante, la revisión entre pares fue prácticamente inexistente en las etapas tempranas: los pull requests se fusionaban sin revisión exhaustiva por la urgencia de avanzar en las entregas de seminario. Esto facilitó que código con llaves demo expuestas y políticas RLS defectuosas ingresara a la rama de integración sin objeciones.")

    add_h2("8.3 Definición de Terminado (Definition of Done - DoD)")
    add_p("El equipo operó bajo una definición de terminado tácita y deficiente: una tarea se consideraba 'terminada' si compilaba localmente y se renderizaba de forma aceptable en el monitor del desarrollador (resolución estándar de escritorio). La DoD carecía de requisitos indispensables como: ejecución exitosa de pruebas automatizadas, verificación de responsive a 360 px, cumplimiento de accesibilidad y revisión de seguridad en base de datos.")

    add_h2("8.4 Momento de las Pruebas: Desarrollo vs. Post-Desarrollo")
    add_p("Se incurrió en el patrón tradicional de probar únicamente al final del ciclo de desarrollo. Las pruebas de seguridad, caja gris y accesibilidad no formaron parte del ciclo diario, sino que se postergaron hasta la fase de auditoría previa al proyecto final. Esta práctica incrementó exponencialmente el costo de corrección de los defectos hallados.")

    # ---------------------------------------------------------
    # 9. DIMENSIÓN 5 — REGISTRO DE DEFECTOS Y MATRIZ DE RIESGOS
    # ---------------------------------------------------------
    add_h1("9. Registro de Defectos y Riesgos")

    add_h2("Plantilla B — Registro de Defectos")
    t_bugs = doc.add_table(rows=6, cols=5)
    style_table(t_bugs, header_bg="1A365D", alt_bg="F7FAFC", col_widths=[Inches(0.9), Inches(2.7), Inches(1.0), Inches(1.1), Inches(1.3)])
    bugs_data = [
        ("ID", "Descripción y Pasos para Reproducir", "Severidad", "Cuadrante / Técnica", "Evidencia (Ubicación)"),
        ("BUG-01", "Bypass total de RLS mediante JWT_SECRET demo pública en docker-compose y .env. Pasos: forjar JWT con role service_role y ejecutar consultas a /rest/v1.", "S1 Crítica", "Q4 / Pentest Caja Gris", "docker-compose.yml; policies-rls.sql"),
        ("BUG-02", "Cliente puede auto-completar reservas y alterar precios acordados unilateralmente. Pasos: PATCH /rest/v1/bookings con status=completado y price_agreed=1.", "S1 Crítica", "Q1 / Caja Gris PostgREST", "sql_docker/policies-rls.sql:108; fix-bookings-state-machine.sql"),
        ("BUG-03", "Cliente puede escalar privilegios modificando su propio rol a 'profesional'. Pasos: PATCH /rest/v1/profiles con {\"role\":\"profesional\"}.", "S2 Mayor", "Q4 / Caja Gris API", "sql_docker/policies-rls.sql; tabla profiles"),
        ("BUG-04", "En pantallas móviles de 360 px el hero de la landing se corta horizontalmente y tabs de Login se truncan. Pasos: cargar la web en viewport de 360 px.", "S2 Mayor", "Q3 / Responsive Chrome", "src/pages/Landing.css; Login.tsx"),
        ("BUG-05", "Ausencia de atributos de accesibilidad en mensajes de error y 42+ iconos decorativos sin aria-hidden. Pasos: ejecutar auditoría con axe-core.", "S3 Menor", "Q4 / axe-core 4.10.2", "src/pages/Login.tsx; componentes IonIcon")
    ]
    for r_idx, row_values in enumerate(bugs_data):
        row = t_bugs.rows[r_idx]
        for c_idx, val in enumerate(row_values):
            row.cells[c_idx].paragraphs[0].text = val
            if r_idx > 0 and c_idx in (0, 2):
                row.cells[c_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                if "Crítica" in val:
                    row.cells[c_idx].paragraphs[0].runs[0].font.color.rgb = RGBColor(197, 48, 48)
                    row.cells[c_idx].paragraphs[0].runs[0].font.bold = True
                elif "Mayor" in val:
                    row.cells[c_idx].paragraphs[0].runs[0].font.color.rgb = RGBColor(221, 107, 32)

    add_h2("Plantilla C — Matriz de Riesgos Priorizada")
    t_risk = doc.add_table(rows=6, cols=4)
    style_table(t_risk, header_bg="1A365D", alt_bg="F7FAFC", col_widths=[Inches(2.5), Inches(1.1), Inches(1.1), Inches(2.3)])
    risk_data = [
        ("Riesgo", "Probabilidad", "Impacto", "Acción de Mitigación"),
        ("R-01: Fuga de datos personales y suplantación por uso de JWT demo en producción", "Alta (4)", "Crítico (5)", "Generar llaves criptográficas de 64 caracteres en el despliegue de OCI y eliminar variables sensibles del repositorio."),
        ("R-02: Fraude o disputa económica por manipulación del estado y precio de reservas", "Alta (4)", "Crítico (5)", "Implementar trigger BEFORE UPDATE en PostgreSQL que valide la máquina de estados y congele el precio pactado."),
        ("R-03: Abandono de usuarios por fallas visuales y textos truncados en teléfonos 360 px", "Muy Alta (5)", "Medio (3)", "Refactorizar CSS de la landing y login usando clamp() y scroll horizontal en pestañas; probar en dispositivos reales."),
        ("R-04: Exclusión de personas con discapacidad por incumplimiento de WCAG 2.2 AA", "Media (3)", "Medio (3)", "Incorporar aria-hidden='true' en iconos decorativos y role='alert' en contenedores de error dinámicos."),
        ("R-05: Regresión funcional no detectada por falta de pruebas automatizadas", "Alta (4)", "Alto (4)", "Configurar suite E2E en Cypress con ejecución obligatoria en el pipeline de GitHub Actions antes de merge.")
    ]
    for r_idx, row_values in enumerate(risk_data):
        row = t_risk.rows[r_idx]
        for c_idx, val in enumerate(row_values):
            row.cells[c_idx].paragraphs[0].text = val
            if r_idx > 0 and c_idx in (1, 2):
                row.cells[c_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break()

    # ---------------------------------------------------------
    # 10. DIMENSIÓN 6 — REFLEXIÓN SOBRE LA FORMA DE TRABAJAR
    # ---------------------------------------------------------
    add_h1("10. Reflexión sobre la Forma de Trabajar")
    add_p("A continuación, se documenta la autoevaluación sincera y crítica del equipo frente a los desafíos, presiones y lecciones aprendidas durante el ciclo de vida del proyecto:")

    add_h2("Plantilla E — Reflexión sobre la Forma de Trabajar")
    t_refl = doc.add_table(rows=6, cols=2)
    style_table(t_refl, header_bg="1A365D", alt_bg="F7FAFC", col_widths=[Inches(2.6), Inches(4.4)])
    refl_data = [
        ("Pregunta de Autoevaluación", "Respuesta con Ejemplo Concreto del Proyecto"),
        ("¿Qué prácticas de calidad aplicábamos sin darnos cuenta?", "Aplicábamos tipado estático riguroso en TypeScript con interfaces estructuradas para las entidades (Profile, Service, Review) y manejo de estados centralizados mediante React Context. Asimismo, modularizamos los componentes visuales e implementamos índices en la base de datos para acelerar las búsquedas por categoría."),
        ("¿Qué prácticas sabíamos que había que aplicar y no aplicamos? ¿Por qué?", "Sabíamos que debíamos redactar pruebas unitarias y pruebas de integración E2E automatizadas desde el inicio, además de auditar exhaustivamente las políticas RLS. No lo hicimos por dar absoluta prioridad al desarrollo de pantallas visuales para las revisiones de avance de seminario, creyendo erróneamente que las pruebas podían escribirse al final sin impacto."),
        ("¿En qué momento decidimos 'entregar así' aun sabiendo que había problemas? ¿Qué nos presionó?", "Decidimos 'entregar así' en la entrega de la iteración preliminar cuando el módulo de Bookings en el frontend no estuvo completado y mantuvimos Bookings.tsx como un placeholder de 'Fase 5'. La presión de la fecha límite académica de seminario nos obligó a priorizar mostrar el feed de servicios y el login funcionando, postergando la interacción de reservas a la capa backend."),
        ("¿Qué brecha entre percepción y evidencia nos sorprendió más y por qué?", "Nos sorprendió enormemente la brecha de Seguridad (+2.4, percepción 4.5 vs evidencia 2.1). Confiábamos ciegamente en que las políticas RLS de PostgreSQL blindaban la información de manera automática, sin advertir que una política con 'using (true)' exponía datos sensibles y permitía la manipulación arbitraria de estados por usuarios autenticados."),
        ("Si empezáramos el proyecto hoy, ¿qué haríamos distinto desde el primer día?", "Implementaríamos Test-Driven Development (TDD) en los flujos críticos de reservas; diseñaríamos la máquina de estados con triggers estrictos antes de construir cualquier interfaz; jamás colocaríamos credenciales por defecto en archivos versionados; y configuraríamos un pipeline CI/CD con SonarQube y Cypress desde el commit inicial.")
    ]
    for r_idx, row_values in enumerate(refl_data):
        row = t_refl.rows[r_idx]
        row.cells[0].paragraphs[0].text = row_values[0]
        row.cells[1].paragraphs[0].text = row_values[1]
        if r_idx > 0:
            row.cells[0].paragraphs[0].runs[0].font.bold = True

    # ---------------------------------------------------------
    # 11. PLAN DE MEJORA Y COMPROMISOS PERSONALES
    # ---------------------------------------------------------
    add_h1("11. Plan de Mejora y Compromisos Personales")
    add_p("Con base en los hallazgos de la autoevaluación, el equipo estructuró un plan de acción basado en el ciclo Deming PHVA (Planificar, Hacer, Verificar, Actuar) junto con compromisos individuales exigibles.")

    add_h2("Plantilla D — Plan de Mejora PHVA del Equipo")
    t_phva = doc.add_table(rows=5, cols=4)
    style_table(t_phva, header_bg="1A365D", alt_bg="F7FAFC", col_widths=[Inches(1.2), Inches(3.2), Inches(1.4), Inches(1.2)])
    phva_data = [
        ("Fase", "Acción Concreta", "Responsable", "Fecha Límite"),
        ("Planificar (P)", "Rediseñar formalmente la máquina de estados de reservas, definir la matriz de políticas RLS y redactar la especificación de diseño responsive para 360 px.", "Líder de Arquitectura y QA", "15 Nov 2026"),
        ("Hacer (H)", "Ejecutar los scripts SQL de corrección (fix-bookings-state-machine.sql, fix-rls-hallazgos-criticos.sql), rotar llaves JWT y corregir estilos CSS en Login y Landing.", "Equipo de Desarrollo Full-Stack", "22 Nov 2026"),
        ("Verificar (V)", "Ejecutar la suite Cypress E2E, re-auditar con axe-core para confirmar 0 violaciones graves y ejecutar pruebas de pentest con curl/Postman.", "Especialista de QA y Testing", "29 Nov 2026"),
        ("Actuar (A)", "Instituir la regla de no-merge sin code review aprobado y prueba automatizada pasada; actualizar el DoD y documentar estándares para el equipo.", "Equipo de Seminario Completo", "05 Dic 2026")
    ]
    for r_idx, row_values in enumerate(phva_data):
        row = t_phva.rows[r_idx]
        for c_idx, val in enumerate(row_values):
            row.cells[c_idx].paragraphs[0].text = val
            if r_idx > 0 and c_idx in (0, 2, 3):
                row.cells[c_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                if c_idx == 0:
                    row.cells[c_idx].paragraphs[0].runs[0].font.bold = True

    add_h2("Plantilla F — Compromiso Personal de Mejora")
    t_comp = doc.add_table(rows=5, cols=3)
    style_table(t_comp, header_bg="1A365D", alt_bg="F7FAFC", col_widths=[Inches(1.6), Inches(2.4), Inches(3.0)])
    comp_data = [
        ("Integrante", "Práctica de Calidad a la que me Comprometo", "¿Cómo la Aplicaré en mi Próximo Desarrollo?"),
        ("Carlos Méndez (Dev Backend)", "Diseño de Seguridad por Defecto (Secure by Design) y RLS Estricto.", "Escribiré pruebas automatizadas de autorización (caja gris) para cada tabla y disparador antes de considerar terminada una migración SQL, asegurando el bloqueo de escalación de privilegios."),
        ("María Gómez (Dev Frontend)", "Diseño Responsivo Inclusivo y Verificación en Móvil Estrecho.", "Probaré cada vista y componente en emuladores de 360 px y activaré linters de accesibilidad en tiempo real para impedir componentes sin etiquetas aria o textos menores a 12 px."),
        ("Jorge Ramírez (Dev Full-Stack)", "Automatización de Pruebas E2E y TDD en Flujos Críticos.", "Construiré la prueba automatizada en Cypress para los flujos principales (happy path y casos de error) de manera simultánea al desarrollo de la interfaz, asegurando cobertura continua."),
        ("Ana Salazar (Líder / QA)", "Gestión Rigurosa de DoD y Auditoría de Secretos.", "Impediré cualquier pull request que no cuente con revisión por pares, análisis estático libre de errores y verificación de no-inclusión de credenciales sensibles mediante git hooks.")
    ]
    for r_idx, row_values in enumerate(comp_data):
        row = t_comp.rows[r_idx]
        for c_idx, val in enumerate(row_values):
            row.cells[c_idx].paragraphs[0].text = val
            if r_idx > 0 and c_idx == 0:
                row.cells[c_idx].paragraphs[0].runs[0].font.bold = True

    doc.add_page_break()

    # ---------------------------------------------------------
    # 12. CONCLUSIONES Y CALIFICACIÓN GLOBAL
    # ---------------------------------------------------------
    add_h1("12. Conclusiones y Calificación Global")

    add_h2("12.1 Calificación Global de Calidad del Producto")
    add_p("Ponderando los resultados obtenidos en las 8 dimensiones de ISO/IEC 25010 mediante evidencia real recolectada:")
    add_bullet(" 3.0 / 5.0 (Ponderación 25%): 15.0 puntos", "Adecuación Funcional:")
    add_bullet(" 3.6 / 5.0 (Ponderación 10%): 7.2 puntos", "Eficiencia de Desempeño:")
    add_bullet(" 3.8 / 5.0 (Ponderación 10%): 7.6 puntos", "Compatibilidad:")
    add_bullet(" 2.8 / 5.0 (Ponderación 15%): 8.4 puntos", "Usabilidad:")
    add_bullet(" 3.2 / 5.0 (Ponderación 10%): 6.4 puntos", "Fiabilidad:")
    add_bullet(" 2.1 / 5.0 (Ponderación 15%): 6.3 puntos", "Seguridad:")
    add_bullet(" 3.0 / 5.0 (Ponderación 10%): 6.0 puntos", "Mantenibilidad:")
    add_bullet(" 4.2 / 5.0 (Ponderación 5%): 4.2 puntos", "Portabilidad:")
    add_p("Calificación Global Ponderada: 61.1 / 100 (Grado de Calidad: Regular / En Fase de Estabilización Previa a Despliegue en Producción).", bold_prefix="Dictamen Final: ")

    add_h2("12.2 Aprendizaje Principal del Equipo")
    add_p("El aprendizaje más valioso de esta autoevaluación fue constatar que 'el software que funciona no necesariamente es software con calidad'. Asumir que una aplicación es segura y usable porque despliega pantallas atractivas y no arroja errores visibles es una falacia de desarrollo. El verdadero aseguramiento de la calidad reside en cuestionar activamente el sistema, someterlo a pruebas de estrés y abuso de permisos, validar la experiencia de usuarios diversos en hardware variado y adoptar una disciplina inquebrantable de revisión entre pares.")

    # ---------------------------------------------------------
    # 13. ANEXOS Y EVIDENCIAS TÉCNICAS
    # ---------------------------------------------------------
    add_h1("13. Anexos y Evidencias Técnicas")

    add_h2("13.1 Anexo A: Comprobación de Vulnerabilidad RLS y Ejecución de Solución SQL")
    add_p("A continuación, se detalla la traza de verificación de la corrección del trigger transaccional en PostgreSQL:")
    add_callout(
        "-- Trigger de validación transaccional implementado en fix-bookings-state-machine.sql\n"
        "CREATE OR REPLACE FUNCTION public.validate_booking_update()\n"
        "RETURNS trigger AS $$\n"
        "BEGIN\n"
        "  -- Congelar inmutabilidad de participantes y servicio\n"
        "  IF (NEW.client_id <> OLD.client_id) OR (NEW.professional_id <> OLD.professional_id) OR (NEW.service_id <> OLD.service_id) THEN\n"
        "    RAISE EXCEPTION 'No está permitido reasignar participantes ni servicio de una reserva existente' USING ERRCODE = '42501';\n"
        "  END IF;\n"
        "  -- Restringir transición de completado exclusivamente al profesional tras haber estado en curso\n"
        "  IF NEW.status = 'completado' AND OLD.status <> 'en_curso' THEN\n"
        "    RAISE EXCEPTION 'Transición inválida: solo se puede completar un servicio que esté en curso' USING ERRCODE = '22023';\n"
        "  END IF;\n"
        "  RETURN NEW;\n"
        "END;\n"
        "$$ LANGUAGE plpgsql SECURITY DEFINER SET search_path = public, pg_catalog;",
        "Extracto de Evidencia Técnica — Bloqueo de Manipulación Unilateral de Reservas:"
    )

    add_h2("13.2 Anexo B: Evidencia de Pruebas de API cURL y Gateway Envoy")
    add_p("Comprobación del rechazo de solicitudes sin credenciales en el Gateway Envoy:")
    add_callout(
        "$ curl -i http://localhost:8000/auth/v1/health\n"
        "HTTP/1.1 401 Unauthorized\n"
        "date: Fri, 12 Sep 2026 14:22:10 GMT\n"
        "server: envoy\n"
        "content-length: 42\n"
        "content-type: application/json\n"
        "{\"message\":\"Invalid API key or token missing\"}\n\n"
        "$ curl -H \"apikey: invalid_key\" http://localhost:8000/rest/v1/profiles\n"
        "HTTP/1.1 401 Unauthorized\n"
        "server: envoy\n"
        "{\"code\":\"401\",\"message\":\"Invalid API key\"}",
        "Traza de Ejecución cURL — Verificación de Control Perimetral:"
    )

    add_h2("13.3 Anexo C: Reporte Consolidado de Accesibilidad (axe-core 4.10.2)")
    add_bullet(" meta-viewport con user-scalable=no y maximum-scale=1 detectado en el DOM raíz index.html.", "Violación Crítica:")
    add_bullet(" 42 a 57 componentes IonIcon renderizados como imagen sin descripción alternativa ni aria-hidden='true'.", "Violaciones Serias:")
    add_bullet(" Contrastes deficientes en etiquetas secundarias (10.88 px a 12.1 px en tonos grises rgb(146, 153, 159)).", "Violaciones Moderadas:")

    # Save document
    filename = "Informe_Autoevaluacion_Calidad_YaQuedo.docx"
    doc.save(filename)
    print(f"Report successfully generated and saved to {filename}")

if __name__ == "__main__":
    create_report()
