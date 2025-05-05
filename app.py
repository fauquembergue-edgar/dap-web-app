# app.py - Application web Flask pour gestion du DAP
from flask import Flask, render_template, request, send_file, redirect, url_for
from docxtpl import DocxTemplate
from datetime import datetime
import os, uuid
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__)
UPLOAD_FOLDER = "temp"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

with open("structure_formulaire_updated.json", "r") as f:
    fields = json.load(f)

@app.route("/")
def home():
    return redirect(url_for("entreprise1"))

@app.route("/entreprise1", methods=["GET", "POST"])
def entreprise1():
    if request.method == "POST":
        data = request.form.to_dict()
        token = str(uuid.uuid4())
        filepath = os.path.join(UPLOAD_FOLDER, f"dap_{token}.docx")
        doc = DocxTemplate("modele_dap_template_final.docx")
        doc.render(data)
        doc.save(filepath)
        envoyer_mail(data.get("email_entreprise2", ""), filepath, token)
        return f"Formulaire envoyé à l'entreprise 2 avec le lien : /entreprise2/{token}"
    return render_template("formulaire.html", fields=fields, titre="Entreprise 1")

@app.route("/entreprise2/<token>", methods=["GET", "POST"])
def entreprise2(token):
    filepath = os.path.join(UPLOAD_FOLDER, f"dap_{token}.docx")
    if not os.path.exists(filepath): return "Fichier introuvable."
    if request.method == "POST":
        data = request.form.to_dict()
        doc = DocxTemplate(filepath)
        data["signature_entreprise_2"] = f"Signé par Entreprise 2 le {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        doc.render(data)
        doc.save(filepath)
        envoyer_mail(data.get("email_entreprise1", ""), filepath, token)
        return f"Formulaire signé et envoyé à l'entreprise 1."
    return render_template("formulaire.html", fields=fields, titre="Entreprise 2")

@app.route("/finalisation/<token>", methods=["GET", "POST"])
def finalisation(token):
    filepath = os.path.join(UPLOAD_FOLDER, f"dap_{token}.docx")
    if not os.path.exists(filepath): return "Fichier introuvable."
    if request.method == "POST":
        data = request.form.to_dict()
        doc = DocxTemplate(filepath)
        data["signature_entreprise_1"] = f"Signé par Entreprise 1 le {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        doc.render(data)
        doc.save(filepath)
        envoyer_mail(data.get("email_entreprise2", ""), filepath, token, final=True)
        return "Formulaire finalisé et renvoyé."
    return render_template("formulaire.html", fields=fields, titre="Signature finale Entreprise 1")

def envoyer_mail(to_addr, filepath, token, final=False):
    if not to_addr: return
    from_addr = os.environ.get("FROM_EMAIL")
    password = os.environ.get("FROM_PASSWORD")
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = "Formulaire DAP finalisé" if final else "Formulaire DAP à remplir"

    with open(filepath, "rb") as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(filepath)}')
        msg.attach(part)

    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login(from_addr, password)
        server.send_message(msg)

@app.route("/telecharger/<token>")
def telecharger(token):
    path = os.path.join(UPLOAD_FOLDER, f"dap_{token}.docx")
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
