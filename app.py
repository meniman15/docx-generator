# -*- coding: utf-8 -*-
import os
from flask import Flask, request, jsonify, send_file
from templateToDocx import fill_template_from_data, SCRIPT_DIR

app = Flask(__name__)

# Template path — resolved relative to templateToDocx.py's directory
TEMPLATE_PATH = os.environ.get(
    "TEMPLATE_PATH",
    os.path.join(SCRIPT_DIR, "templateV2.docx"),
)


@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint."""
    return jsonify({"status": "ok"})


@app.route("/api/generate-docx", methods=["POST"])
def generate_docx():
    """
    Accept JSON body with meeting data, return a filled Meeting_Summary.docx.

    Expected JSON structure:
    {
        "פרטי הפגישה": { "שם הפגישה": "...", "תאריך": "dd/mm/yy", ... },
        "משתתפים": [ { "משתתף": "...", "תפקיד": "...", "ארגון": "..." }, ... ],
        "הנחיות": [ { "הנחיה": "...", "למי": "...", "תאריך": "..." }, ... ],
        "רשימת תפוצה": [ { "משתתף": "...", "תפקיד": "...", "ארגון": "..." }, ... ]
    }
    """
    # --- Parse JSON body ---
    if not request.is_json:
        # Handle possible BOM in raw body
        raw = request.get_data()
        if raw[:3] == b'\xef\xbb\xbf':
            raw = raw[3:]
        try:
            import json
            data = json.loads(raw.decode("utf-8"))
        except Exception as e:
            return jsonify({"error": f"Invalid JSON: {str(e)}"}), 400
    else:
        data = request.get_json(silent=True)
        if data is None:
            return jsonify({"error": "Request body must be valid JSON"}), 400

    # --- Validate required key ---
    # Origami API format: must have "data" list
    # Legacy format: must have "פרטי הפגישה"
    is_origami = "data" in data and isinstance(data.get("data"), list)
    is_legacy = "פרטי הפגישה" in data
    is_flat = isinstance(data, dict) and len(data) > 0

    if not is_origami and not is_legacy and not is_flat:
        return jsonify({
            "error": "Invalid JSON structure",
            "hint": "Expected a non-empty JSON object.",
        }), 400

    # --- Verify template exists ---
    if not os.path.isfile(TEMPLATE_PATH):
        return jsonify({
            "error": f"Template file not found: {TEMPLATE_PATH}",
        }), 500

    # --- Generate the document ---
    try:
        buf = fill_template_from_data(TEMPLATE_PATH, data)
    except Exception as e:
        return jsonify({"error": f"Document generation failed: {str(e)}"}), 500

    # --- Return the .docx file ---
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="Meeting_Summary.docx",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    print(f"🚀 Starting server on http://0.0.0.0:{port}")
    print(f"📄 Template: {TEMPLATE_PATH}")
    app.run(host="0.0.0.0", port=port, debug=False)
