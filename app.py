from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from flask_cors import CORS
from sheets_client import SheetsClient
import os
import time

app = Flask(__name__)
CORS(app)
app.secret_key = os.environ.get("FLASK_SECRET", "dev-secret-change-me")

# Initialize Sheets client
try:
    sheets = SheetsClient()
except Exception as e:
    sheets = None
    print("Sheets init failed:", e)

# --------------orders------------------
@app.route("/orders")
def orders():
    if not sheets:
        return render_template("error.html", message="Sheets not available")

    pivot = sheets.get_pivot_data()
    return render_template(
        "orders.html",
        products=pivot["products"],
        parties=pivot["parties"],
        pivot=pivot["pivot"]
    )
# ---------------- HEALTH ----------------
@app.route("/_health")
def health():
    return jsonify({
        "status": "ok",
        "sheets_initialized": sheets is not None
    }), 200


# ---------------- HOME ----------------
@app.route("/")
def index():
    if not sheets:
        return render_template("error.html", message="Sheets not initialized")

    lists = sheets.load_lists()
    recent_orders = sheets.get_recent_orders(50)

    return render_template(
        "index.html",
        products=lists["products"],
        companies=lists["companies"],
        brands=lists["brands"],
        recent_orders=recent_orders
    )


# ---------------- SUBMIT ORDER ----------------
@app.route("/submit", methods=["POST"])
def submit():
    if not sheets:
        flash("Sheets not available", "danger")
        return redirect(url_for("index"))

    company = request.form.get("company", "").strip()
    includes_gst = request.form.get("includes_gst") == "on"

    if not company:
        flash("Company required", "warning")
        return redirect(url_for("index"))

    success = 0

    for key in request.form:
        if not key.startswith("orders[") or "[product]" not in key:
            continue

        idx = key[key.find("[")+1:key.find("]")]
        product = request.form.get(f"orders[{idx}][product]", "").strip()
        brand = request.form.get(f"orders[{idx}][brand]", "").strip()
        qty = request.form.get(f"orders[{idx}][quantity]", "").strip()
        price = request.form.get(f"orders[{idx}][price]", "").strip()

        if not product or not qty or not price:
            continue

        try:
            qty = int(qty)
            price = float(price)
            if includes_gst:
                price = round(price / 1.05, 2)

            order_number = f"{int(time.time())}-{idx}"

            sheets.add_order(
                company=company,
                product=product,
                quantity=qty,
                price=price,
                brand=brand,
            )
            success += 1
        except Exception as e:
            print("Order error:", e)

    if success:
        flash(f"{success} orders added successfully", "success")

    return redirect(url_for("index"))


# ---------------- APIs ----------------
@app.route("/api/products")
def api_products():
    return jsonify({"products": sheets.load_lists()["products"] if sheets else []})


@app.route("/api/companies")
def api_companies():
    return jsonify({"companies": sheets.load_lists()["companies"] if sheets else []})


@app.route("/api/orders_by_product")
def api_orders_by_product():
    product = request.args.get("product", "")
    return jsonify({"orders": sheets.get_orders_by_product(product) if sheets else []})


@app.route("/api/orders_by_party")
def api_orders_by_party():
    company = request.args.get("company", "")
    return jsonify({"orders": sheets.get_orders_by_party(company) if sheets else []})


@app.route("/api/pivot_data")
def api_pivot_data():
    if not sheets:
        return jsonify({"pivot": [], "products": [], "parties": []})

    pf = request.args.get("product_filter", "")
    cf = request.args.get("party_filter", "")
    return jsonify(sheets.get_pivot_data(pf, cf))


# ---------------- DISPATCH ----------------
@app.route("/dispatch")
def dispatch():
    return render_template("dispatch.html")


@app.route("/dispatch/save", methods=["POST"])
def save_dispatch():
    if not sheets:
        return jsonify({"ok": False}), 400

    data = request.json
    for d in data.get("dispatches", []):
        sheets.add_dispatch(
            company=d["company"],
            product=d["product"],
            quantity=int(d["quantity"]),
            order_number=d["order_number"]
        )

    return jsonify({"ok": True})
@app.route("/api/parties_with_pending")
def parties_with_pending():
    if not sheets:
        return jsonify({"companies": []})

    pivot = sheets.get_pivot_data()
    return jsonify({"companies": pivot["parties"]})


if __name__ == "__main__":
    app.run(port=8000, debug=True)
