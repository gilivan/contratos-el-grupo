"""
app.py — Servidor web para generar documentos El Grupo v2
Tipos: Contrato Proveedor · Contrato Freefan · Acta de Cierre

Uso:
    pip install flask python-docx gunicorn
    python3 app.py
    → Abrir http://localhost:8080 en el navegador
"""

import os
import tempfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify
import generar_contrato as gc

app = Flask(__name__)
BASE_DIR = Path(__file__).parent


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generar", methods=["POST"])
def generar():
    try:
        # Ajustar ASSETS_DIR al directorio de la app
        gc.ASSETS_DIR = BASE_DIR / "assets"

        tipo    = request.form.get("tipo", "").strip()
        empresa = request.form.get("empresa", "").strip()

        if tipo not in ("proveedor", "freefan", "acta_cierre"):
            return jsonify({"error": "Tipo de documento no válido."}), 400
        if empresa not in ("fiera", "perez_villa"):
            return jsonify({"error": "Selecciona la razón social del contratista."}), 400

        # ── PROVEEDOR ─────────────────────────────────────────────────
        if tipo == "proveedor":
            actividades_raw = request.form.getlist("actividades")
            actividades = [a.strip() for a in actividades_raw if a.strip()]
            if not actividades:
                return jsonify({"error": "Agrega al menos una actividad al alcance."}), 400

            data = {
                "tipo":                          tipo,
                "empresa":                       empresa,
                "empresa_contratante":           request.form.get("empresa_contratante", "").strip().upper(),
                "nombre_rl":                     request.form.get("nombre_rl", "").strip().upper(),
                "cedula_rl":                     request.form.get("cedula_rl", "").strip(),
                "direccion":                     request.form.get("direccion", "").strip(),
                "nit":                           request.form.get("nit", "").strip(),
                "contacto_contratante_nombre":   request.form.get("contacto_nombre", "").strip(),
                "contacto_contratante_telefono": request.form.get("contacto_telefono", "").strip(),
                "contacto_contratante_correo":   request.form.get("contacto_correo", "").strip(),
                "actividades":                   actividades,
                "duracion_cantidad":             request.form.get("duracion_cantidad", "").strip(),
                "duracion_unidad":               request.form.get("duracion_unidad", "meses"),
                "fecha_inicio":                  request.form.get("fecha_inicio", "").strip(),
                "valor_pago":                    request.form.get("valor_pago", "").strip(),
                "firma_dia":                     request.form.get("firma_dia", "").strip(),
                "firma_mes":                     request.form.get("firma_mes", "").strip(),
                "firma_anio":                    request.form.get("firma_anio", "").strip(),
            }

            required = [
                ("empresa_contratante", "Nombre de la empresa contratante"),
                ("nombre_rl",           "Representante Legal"),
                ("cedula_rl",           "Cédula del RL"),
                ("nit",                 "NIT"),
                ("duracion_cantidad",   "Duración"),
                ("fecha_inicio",        "Fecha de inicio"),
                ("valor_pago",          "Valor y forma de pago"),
                ("firma_dia",           "Día de firma"),
                ("firma_mes",           "Mes de firma"),
                ("firma_anio",          "Año de firma"),
            ]
            for campo, label in required:
                if not data.get(campo):
                    return jsonify({"error": f"El campo «{label}» es obligatorio."}), 400

            template_key = empresa
            doc = gc.Document(str(gc.ASSETS_DIR / gc.TEMPLATE_FILES[template_key]))
            gc.process_body(doc, data)
            gc.process_table(doc, data, empresa)

            empresa_label  = "FieraS.A.S" if empresa == "fiera" else "PerezVilla"
            nombre_cliente = data["empresa_contratante"].replace(" ", "")[:20]
            filename = f"Contrato_Proveedor_{nombre_cliente}_{empresa_label}.docx"

        # ── FREEFAN ───────────────────────────────────────────────────
        elif tipo == "freefan":
            data = {
                "tipo":               tipo,
                "empresa":            empresa,
                "contratista_nombre": request.form.get("ff_contratista_nombre", "").strip().upper(),
                "contratista_cc":     request.form.get("ff_contratista_cc", "").strip(),
                "numero_contrato":    request.form.get("ff_numero_contrato", "").strip(),
                "fecha_inicio":       request.form.get("ff_fecha_inicio", "").strip(),
                "fecha_fin":          request.form.get("ff_fecha_fin", "").strip(),
                "duracion_dias":      request.form.get("ff_duracion_dias", "").strip(),
                "valor":              request.form.get("ff_valor", "").strip(),
                "firma_dia":          request.form.get("ff_firma_dia", "").strip(),
                "firma_mes":          request.form.get("ff_firma_mes", "").strip(),
                "firma_anio":         request.form.get("ff_firma_anio", "").strip(),
            }

            required = [
                ("contratista_nombre", "Nombre del contratista"),
                ("contratista_cc",     "Cédula del contratista"),
                ("numero_contrato",    "Número de contrato"),
                ("fecha_inicio",       "Fecha de inicio"),
                ("fecha_fin",          "Fecha de fin"),
                ("duracion_dias",      "Duración en días"),
                ("valor",              "Valor del contrato"),
                ("firma_dia",          "Día de firma"),
                ("firma_mes",          "Mes de firma"),
                ("firma_anio",         "Año de firma"),
            ]
            for campo, label in required:
                if not data.get(campo):
                    return jsonify({"error": f"El campo «{label}» es obligatorio."}), 400

            doc = gc.Document(str(gc.ASSETS_DIR / gc.TEMPLATE_FILES["freefan"]))
            gc.process_freefan(doc, data)

            empresa_label = "FieraS.A.S" if empresa == "fiera" else "PerezVilla"
            filename = f"Contrato_Freefan_{data['numero_contrato']}_{empresa_label}.docx"

        # ── ACTA DE CIERRE ────────────────────────────────────────────
        elif tipo == "acta_cierre":
            data = {
                "tipo":               tipo,
                "empresa":            empresa,
                "contratista_nombre": request.form.get("ac_contratista_nombre", "").strip().upper(),
                "contratista_cc":     request.form.get("ac_contratista_cc", "").strip(),
                "numero_acta":        request.form.get("ac_numero_acta", "").strip(),
                "fecha_actual":       request.form.get("ac_fecha_actual", "").strip(),
                "fecha_contrato":     request.form.get("ac_fecha_contrato", "").strip(),
                "fecha_finalizacion": request.form.get("ac_fecha_finalizacion", "").strip(),
                "numero_contrato":    request.form.get("ac_numero_contrato", "").strip(),
                "valor":              request.form.get("ac_valor", "").strip(),
            }

            required = [
                ("contratista_nombre", "Nombre del contratista"),
                ("contratista_cc",     "Cédula del contratista"),
                ("numero_acta",        "Número de acta"),
                ("fecha_actual",       "Fecha del acta"),
                ("fecha_contrato",     "Fecha del contrato original"),
                ("fecha_finalizacion", "Fecha de finalización"),
                ("numero_contrato",    "Número de contrato"),
                ("valor",              "Valor del contrato"),
            ]
            for campo, label in required:
                if not data.get(campo):
                    return jsonify({"error": f"El campo «{label}» es obligatorio."}), 400

            doc = gc.Document(str(gc.ASSETS_DIR / gc.TEMPLATE_FILES["acta_cierre"]))
            gc.process_acta(doc, data)

            empresa_label = "FieraS.A.S" if empresa == "fiera" else "PerezVilla"
            filename = f"Acta_Cierre_{data['numero_acta']}_{empresa_label}.docx"

        # ── Guardar y enviar ──────────────────────────────────────────
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_path = tmp.name
        doc.save(tmp_path)

        return send_file(
            tmp_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        return jsonify({"error": f"Error al generar el documento: {str(e)}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    print(f"\n✅  Servidor iniciado — abre http://localhost:{port} en tu navegador\n")
    app.run(debug=False, host="0.0.0.0", port=port)
