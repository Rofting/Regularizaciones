"""
prueba_facturas.py
==================
Script de prueba para procesar todos los PDFs de la carpeta entrada/.
Ejecutar desde la carpeta core/:
    python prueba_facturas.py
"""
import sys, shutil, os
sys.path.insert(0, '.')

from lector_pdf import procesar_archivo
from gestor_bd import (conectar, buscar_comunidad_por_cif, insertar_factura,
                        marcar_archivo_procesado, archivo_ya_procesado)

CARPETA_ENTRADA    = '../entrada'
CARPETA_PROCESADOS = '../procesados'
CARPETA_CUARENTENA = '../cuarentena'   # PDFs que no se pudieron procesar
RUTA_BD            = '../data/gestion.db'
RUTA_PROVEEDORES   = '../config/proveedores.json'

os.makedirs(CARPETA_PROCESADOS, exist_ok=True)
os.makedirs(CARPETA_CUARENTENA, exist_ok=True)

con = conectar(RUTA_BD)
archivos = sorted([f for f in os.listdir(CARPETA_ENTRADA) if f.lower().endswith('.pdf')])

if not archivos:
    print('📂 No hay PDFs en la carpeta entrada/')
    con.close()
    sys.exit(0)

print(f'📂 {len(archivos)} archivo(s) encontrado(s)\n')

ok_count = dup_count = err_count = pago_count = lect_count = 0

for nombre in archivos:
    ruta = os.path.join(CARPETA_ENTRADA, nombre)

    # ¿Ya procesado?
    if archivo_ya_procesado(con, nombre):
        print(f'⏭️  Ya procesado: {nombre}')
        dup_count += 1
        continue

    # Detectar comunidad por nombre del archivo
    # Soporta: '644_baser.pdf', '644-baser.pdf', '644 baser.pdf'
    import re as _re
    m = _re.match(r'^(\d{3,4})[_\-\s]', nombre)
    codigo = m.group(1) if m else '644'

    resultado = procesar_archivo(ruta, codigo, con_bd=con,
                                  ruta_proveedores=RUTA_PROVEEDORES)

    # ── FACTURA NORMAL ──────────────────────────────────────────────────────
    if resultado.get('ok') and resultado.get('tipo') == 'FACTURA':
        datos = resultado['datos']

        # Buscar comunidad por CIF (doble seguridad)
        cif = resultado.get('cif_pdf')
        com = buscar_comunidad_por_cif(con, cif) if cif else None

        if not com:
            print(f'⚠️  {nombre}: CIF {cif} no registrado en BD')
            print(f'   → Registra la comunidad primero con el script de inicialización')
            shutil.move(ruta, os.path.join(CARPETA_CUARENTENA, nombre))
            err_count += 1
            continue

        # Buscar periodo abierto
        per = con.execute(
            "SELECT id_periodo FROM periodos "
            "WHERE id_comunidad=? AND estado='abierto' "
            "ORDER BY fecha_inicio DESC LIMIT 1",
            (com['id_comunidad'],)
        ).fetchone()

        if not per:
            print(f'⚠️  {nombre}: sin periodo abierto para comunidad {codigo}')
            print(f'   → Crea el periodo con: python inicializar_comunidades.py')
            err_count += 1
            continue

        datos['id_comunidad'] = com['id_comunidad']
        datos['id_periodo']   = per['id_periodo']
        id_fac = insertar_factura(con, datos)

        if id_fac:
            marcar_archivo_procesado(con, nombre, resultado['hash_md5'], id_fac)
            shutil.move(ruta, os.path.join(CARPETA_PROCESADOS, nombre))
            tipo = datos.get('tipo_suministro', '?')
            imp  = datos.get('importe_total', 0)
            prov = resultado.get('proveedor_clave', '?')
            print(f'✅ {nombre:<45} {tipo:<15} {imp:>9.2f} €  [{prov}]')
            ok_count += 1
        else:
            print(f'⚠️  {nombre}: duplicado (ya estaba en BD)')
            dup_count += 1

    # ── LECTURA METRIGEST ───────────────────────────────────────────────────
    elif resultado.get('ok') and resultado.get('tipo') == 'LECTURA_METRIGEST':
        nvecs     = len(resultado['datos'].get('vecinos', []))
        tipo_lec  = resultado['datos'].get('tipo', '?')
        cabecera  = resultado['datos'].get('cabecera', {})
        periodo   = f"{cabecera.get('fecha_inicio','?')} → {cabecera.get('fecha_fin','?')}"
        marcar_archivo_procesado(con, nombre, resultado['hash_md5'],
                                  resultado='ok', notas=f'LECTURA {tipo_lec}')
        print(f'📊 {nombre:<45} LECTURA {tipo_lec:<12} {nvecs} vecinos  [{periodo}]')
        lect_count += 1

    # ── JUSTIFICANTE DE PAGO (no es factura) ───────────────────────────────
    elif resultado.get('ok') and resultado.get('tipo') == 'JUSTIFICANTE_PAGO':
        datos = resultado['datos']
        imp   = datos.get('importe_total', 0)
        prov  = resultado.get('proveedor_clave', '?')
        marcar_archivo_procesado(con, nombre, resultado['hash_md5'],
                                  resultado='ok', notas='JUSTIFICANTE_PAGO')
        print(f'💳 {nombre:<45} PAGO          {imp:>9.2f} €  [{prov}] (no es factura)')
        pago_count += 1

    # ── ERROR ───────────────────────────────────────────────────────────────
    else:
        motivo  = resultado.get('motivo', 'ERROR_DESCONOCIDO')
        detalle = resultado.get('detalle', '')
        ocr_msg = '  → Necesita OCR (ver instrucciones abajo)' if resultado.get('requiere_ocr') else ''

        print(f'❌ {nombre:<45} {motivo}')
        if detalle:
            print(f'   {detalle}')
        if ocr_msg:
            print(ocr_msg)

        # Mover a cuarentena si el fallo es definitivo
        if motivo in ('PROVEEDOR_NO_IDENTIFICADO', 'SIN_TEXTO', 'CUPS_NO_COINCIDE',
                      'COMUNIDAD_CONTRADICTORIA', 'COMUNIDAD_NO_IDENTIFICADA'):
            dest = os.path.join(CARPETA_CUARENTENA, nombre)
            if not os.path.exists(dest):
                shutil.copy2(ruta, dest)   # copia, no mueve — para revisión
            print(f'   → Copiado a cuarentena/ para revisión')
        err_count += 1

con.close()

print(f"""
{'─'*60}
  ✅ Facturas procesadas:  {ok_count}
  📊 Lecturas Metrigest:   {lect_count}
  💳 Justificantes pago:   {pago_count}
  ⏭️  Ya procesados antes:  {dup_count}
  ❌ Errores / cuarentena: {err_count}
{'─'*60}
""")

if err_count > 0:
    print("📁 Los archivos con error están en ../cuarentena/ para revisión.")
    print("   Para el agua escaneada: pip install pytesseract")
    print("   y descarga Tesseract de https://github.com/UB-Mannheim/tesseract/wiki")
