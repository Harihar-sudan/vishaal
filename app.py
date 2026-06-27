from flask import Flask, render_template, request, send_file
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import io

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///company.db'
db = SQLAlchemy(app)

# ------------------ MODELS ------------------
class Delivery(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.String(20))
    dno = db.Column(db.String(20))
    pcs = db.Column(db.Integer)
    weight = db.Column(db.Float)

class Received(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.String(20))
    dno = db.Column(db.String(20))
    weight = db.Column(db.Float)

# ------------------ ROUTES ------------------
@app.route("/")
def home():
    return render_template("home.html")

# --- Delivery ---
@app.route("/delivery", methods=["GET", "POST"])
def delivery():
    if request.method == "POST":
        new_entry = Delivery(
            date=request.form["date"],
            dno=request.form["dno"],
            pcs=request.form["pcs"],
            weight=request.form["weight"]
        )
        db.session.add(new_entry)
        db.session.commit()
        return "Delivery entry saved!"
    return render_template("delivery.html")

@app.route("/delivery/old")
def delivery_old():
    deliveries = Delivery.query.all()
    return render_template("old_data.html", deliveries=deliveries, show_pcs=True)

# --- Received ---
@app.route("/received", methods=["GET", "POST"])
def received():
    if request.method == "POST":
        new_entry = Received(
            date=request.form["date"],
            dno=request.form["dno"],
            weight=request.form["weight"]
        )
        db.session.add(new_entry)
        db.session.commit()
        return "Received entry saved!"
    return render_template("received.html")

@app.route("/received/old")
def received_old():
    received_entries = Received.query.all()
    return render_template("old_data.html", deliveries=received_entries, show_pcs=False)

# --- Delete Entry ---
@app.route("/delete/<int:id>")
def delete(id):
    # Try deleting from Delivery first, then Received
    entry = Delivery.query.get(id)
    if entry:
        db.session.delete(entry)
        db.session.commit()
        return "Delivery entry deleted!"
    entry = Received.query.get(id)
    if entry:
        db.session.delete(entry)
        db.session.commit()
        return "Received entry deleted!"
    return "Entry not found!"

# --- Download Excel ---
@app.route("/download/delivery")
def download_delivery():
    deliveries = Delivery.query.all()
    data = [{
        "Date": d.date,
        "D.No": d.dno,
        "PCS": d.pcs,
        "Weight": d.weight
    } for d in deliveries]

    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Deliveries")
    output.seek(0)

    return send_file(output, as_attachment=True, download_name="delivery_data.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/download/received")
def download_received():
    received_entries = Received.query.all()
    data = [{
        "Date": r.date,
        "D.No": r.dno,
        "Weight": r.weight
    } for r in received_entries]

    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Received")
    output.seek(0)

    return send_file(output, as_attachment=True, download_name="received_data.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/delete_all/delivery")
def delete_all_delivery():
    Delivery.query.delete()
    db.session.commit()
    return "All Delivery data deleted!"

@app.route("/delete_all/received")
def delete_all_received():
    Received.query.delete()
    db.session.commit()
    return "All Received data deleted!"

@app.route("/edit/<int:id>", methods=["GET", "POST"])
def edit(id):
    # Try Delivery first
    entry = Delivery.query.get(id)
    if not entry:
        entry = Received.query.get(id)

    if request.method == "POST":
        entry.date = request.form["date"]
        entry.dno = request.form["dno"]
        if isinstance(entry, Delivery):  # Only Delivery has PCS
            entry.pcs = request.form["pcs"]
        entry.weight = request.form["weight"]
        db.session.commit()
        return "Entry updated successfully!"

    # Render edit form
    return render_template("edit.html", entry=entry, show_pcs=isinstance(entry, Delivery))


# ------------------ MAIN ------------------
if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True)
