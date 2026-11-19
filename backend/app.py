from flask import Flask, request, send_file, jsonify
from docx import Document
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
import io
import os
import openpyxl
from datetime import datetime

app = Flask(__name__)
CORS(app)

app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://root:rootpassword@db/paynet_db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

class DocumentRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    type = db.Column(db.String(255))
    number = db.Column(db.String(100))
    client_name = db.Column(db.String(255))
    idnp = db.Column(db.String(50))
    operator = db.Column(db.String(100))
    status = db.Column(db.String(50), default='pending')
    file_path = db.Column(db.String(255), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id, 'type': self.type, 'number': self.number,
            'name': self.client_name, 'idnp': self.idnp,
            'operator': self.operator, 'status': self.status,
            'created': self.created_at.isoformat()
        }

import time
with app.app_context():
    retries = 15
    while retries > 0:
        try:
            db.create_all()
            # Încercăm să adăugăm coloana file_path dacă tabela există deja dar coloana lipsește
            try:
                db.session.execute(text("ALTER TABLE document_record ADD COLUMN file_path VARCHAR(255)"))
                db.session.commit()
                print("✅ Coloana 'file_path' a fost adăugată în tabela existentă.")
            except Exception as e:
                db.session.rollback() # Dacă dă eroare (ex: coloana există deja), ignorăm

            print("✅ Baza de date a fost conectată cu succes!")
            break
        except Exception as e:
            print(f"⏳ Se așteaptă baza de date... ({retries} încercări rămase): {e}")
            time.sleep(2)
            retries -= 1

@app.route('/api/ping', methods=['GET'])
def ping():
    return jsonify({"status": "ok", "message": "Serverul Python rulează!"})

TEMPLATES_DIR = '/app/templates'

def replace_text_in_paragraph(paragraph, replacements):
    inline_text = "".join([run.text for run in paragraph.runs])
    for key, val in replacements.items():
        inline_text = inline_text.replace(key, str(val))

    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = inline_text
    else:
        paragraph.add_run(inline_text)

def replace_in_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                replace_text_in_paragraph(p, replacements)

@app.route('/api/generate-certificate', methods=['POST'])
def generate_cert():
    data = request.json
    new_doc = DocumentRecord(
        type=data.get('type'), number=data.get('number'),
        client_name=data.get('name'), idnp=data.get('idnp'), operator=data.get('operator')
    )
    db.session.add(new_doc)
    db.session.commit()

    cert_type = data.get('type', '')
    if 'card închis' in cert_type:
        template_path = os.path.join(TEMPLATES_DIR, 'Certificat_card_inchis_template.docx')
    elif 'deținerea cardului' in cert_type:
        template_path = os.path.join(TEMPLATES_DIR, 'Certificat_card_template.docx')
    elif 'rechizitele' in cert_type:
        template_path = os.path.join(TEMPLATES_DIR, 'Certificat_rechizite_bancare_template.docx')
    else:
        return jsonify({"error": "Șablon negăsit"}), 400

    if not os.path.exists(template_path):
        return jsonify({"error": f"Fișierul {template_path} nu există."}), 404

    doc = Document(template_path)
    title_full = "Dlui" if data.get('title') == 'Dl' else "Dnei"
    replacements = {
        'Nr. 000/26-O din zz.ll.aaaa': f"Nr. {data.get('number')} din {data.get('date')}",
        'Nume prenume client': data.get('name', ''),
        'Adresa de domiciliu client': data.get('address', ''),
        'Adresa de email client': data.get('email', ''),
        '000000000000': data.get('idnp', ''),
        '123456#####1234': data.get('cardNumber', ''),
        '000000#####0000': data.get('cardNumber', ''),
        'll/zz': data.get('cardExpiry', ''),
        'zz.ll.aaaa - zz.ll.aaaa': f"{data.get('closedStart')} - {data.get('closedEnd')}",
        '11111111111111': data.get('account', ''),
        '1111111111111': data.get('account', ''),
        'Igor MUNTEANU': data.get('signer', 'Igor MUNTEANU'),
        'Dlui ': f"{title_full} ",
        'Dlui / Dnei ': f"{title_full} ",
    }
    for p in doc.paragraphs: replace_text_in_paragraph(p, replacements)
    for table in doc.tables: replace_in_table(table, replacements)

    # Salvăm pe disc
    safe_filename = f"cert_{new_doc.id}_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
    file_path = os.path.join('/app/uploads', safe_filename)
    doc.save(file_path)

    # Actualizăm calea în baza de date
    new_doc.file_path = safe_filename
    db.session.commit()

    # Trimitem către browser
    return send_file(file_path, as_attachment=True, download_name='Certificat_Paynet.docx')

@app.route('/api/get-documents', methods=['GET'])
def get_documents():
    docs = DocumentRecord.query.order_by(DocumentRecord.created_at.desc()).all()
    return jsonify([d.to_dict() for d in docs])

@app.route('/api/process-statement', methods=['POST'])
def process_statement():
    if 'file' not in request.files:
        return jsonify({"error": "Nu s-a primit niciun fișier."}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "Niciun fișier selectat."}), 400

    # Preluăm datele suplimentare trimise din frontend (dacă există)
    client_name = request.form.get('client_name', 'Nespecificat')
    operator = request.form.get('operator', 'Operator')

    try:
        # Procesăm fișierul Excel
        wb = openpyxl.load_workbook(file)

        # (Opțional) Aici poți adăuga logica de formatare a extrasului
        # wb.active['A1'] = "PAYNET"

        # Salvăm fișierul pe disc
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        safe_filename = f"extras_{timestamp}.xlsx"
        file_path = os.path.join('/app/uploads', safe_filename)
        wb.save(file_path)

        # Salvăm în baza de date
        new_doc = DocumentRecord(
            type="Extras de cont",
            number="-",
            client_name=client_name,
            idnp="-",
            operator=operator,
            status="signed", # Il marcam ca semnat/finalizat
            file_path=safe_filename
        )
        db.session.add(new_doc)
        db.session.commit()

        # Trimitem fișierul către browser pentru descărcare
        return send_file(file_path, as_attachment=True, download_name='Extras_Procesat.xlsx')
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/download/<int:doc_id>', methods=['GET'])
def download_file(doc_id):
    doc = DocumentRecord.query.get(doc_id)
    if not doc or not doc.file_path:
        return jsonify({"error": "Fișierul nu a fost găsit."}), 404

    file_path = os.path.join('/app/uploads', doc.file_path)
    if not os.path.exists(file_path):
        return jsonify({"error": "Fișierul fizic lipsește de pe server."}), 404

    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8765)
