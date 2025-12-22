from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from flask_cors import CORS
from sheets_client import SheetsClient
import os
import time
from collections import defaultdict
from collections import deque
import hashlib
_recent_submissions = deque(maxlen=200)
DEDUP_WINDOW = 5  # seconds

app = Flask(__name__)
CORS(app)
app.secret_key = os.environ.get("FLASK_SECRET", "dev-secret-change-me")
_rate_limit = defaultdict(list)
RATE_LIMIT = 30  # requests
WINDOW = 60      # seconds
def rate_limited(key):
    now = time.time()
    window = _rate_limit[key]

    # remove old
    _rate_limit[key] = [t for t in window if now - t < WINDOW]

    if len(_rate_limit[key]) >= RATE_LIMIT:
        return True

    _rate_limit[key].append(now)
    return False

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

    lists = sheets.load_lists()
    return render_template(
        "orders.html",
        products=lists["products"],
        parties=lists["companies"]
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

    # ---------------- BUILD ORDER LINES ----------------
    order_lines = []

    for key in request.form:
        if not key.startswith("orders[") or "[product]" not in key:
            continue

        idx = key[key.find("[") + 1 : key.find("]")]
        product = request.form.get(f"orders[{idx}][product]", "").strip()
        brand = request.form.get(f"orders[{idx}][brand]", "").strip()
        qty = request.form.get(f"orders[{idx}][quantity]", "").strip()
        price = request.form.get(f"orders[{idx}][price]", "").strip()

        if not product or not qty or not price:
            continue

        try:
            qty = int(qty)
            price = float(price)
            if qty <= 0:
                continue

            if includes_gst:
                price = round(price / 1.05, 2)

            order_lines.append((product, brand, qty, price))
        except ValueError:
            continue

    if not order_lines:
        flash("No valid order items", "warning")
        return redirect(url_for("index"))

    # ---------------- DEDUPLICATION ----------------
    now = time.time()

    # Create stable fingerprint (order-insensitive)
    fingerprint_data = (
        company,
        tuple(sorted((p, b, q, pr) for p, b, q, pr in order_lines))
    )

    fingerprint = hashlib.sha256(
        repr(fingerprint_data).encode()
    ).hexdigest()

    for ts, fp in list(_recent_submissions):
        if fp == fingerprint and now - ts < DEDUP_WINDOW:
            return redirect(url_for("index"))

    _recent_submissions.append((now, fingerprint))

    # ---------------- WRITE TO SHEET ----------------
    success = 0

    for product, brand, qty, price in order_lines:
        try:
            sheets.add_order(
                company=company,
                product=product,
                brand=brand,
                quantity=qty,
                price=price
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
    if rate_limited("pivot"):
        return jsonify({"error": "Rate limit exceeded"}), 429

    product = request.args.get("product", "")
    return jsonify({"orders": sheets.get_orders_by_product(product) if sheets else []})


@app.route("/api/orders_by_party")
def api_orders_by_party():
    if rate_limited("pivot"):
        return jsonify({"error": "Rate limit exceeded"}), 429
    company = request.args.get("company", "")
    return jsonify({"orders": sheets.get_orders_by_party(company) if sheets else []})


@app.route("/api/pivot_data")
def api_pivot_data():
    if rate_limited("pivot"):
        return jsonify({"error": "Rate limit exceeded"}), 429

    if not sheets:
        print ("Sheets not initialized")
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
        return jsonify({"ok": False, "error": "Sheets not initialized"}), 500

    data = request.get_json(force=True, silent=True)
    if not data or "dispatches" not in data:
        return jsonify({"ok": False, "error": "Invalid payload"}), 400

    written = 0
    errors = []

    for d in data["dispatches"]:
        try:
            company= d.get("company", "").strip()
            serial = str(d.get("order_number", "")).strip()
            product = str(d.get("product", "")).strip()
            qty = int(d.get("quantity", 0))

            if not serial or not product or qty <= 0:
                continue

            sheets.add_dispatch(
                company=company,              # company is optional for logging
                product=product,
                quantity=qty,
                order_number=serial
            )
            written += 1

        except Exception as e:
            errors.append(str(e))

    if written == 0:
        return jsonify({
            "ok": False,
            "error": "No dispatch rows written",
            "details": errors
        }), 400

    return jsonify({
        "ok": True,
        "rows_written": written
    })


@app.route("/api/parties_with_pending")
def parties_with_pending():
    if not sheets:
        return jsonify({"companies": []})

    pivot = sheets.get_pivot_data()
    return jsonify({"companies": pivot["parties"]})
@app.route("/api/products_with_pending")
def products_with_pending():
    if not sheets:
        return jsonify({"products": []})

    pivot = sheets.get_pivot_data()
    return jsonify({"products": pivot["products"]})

@app.route("/api/recent_orders")
def api_recent_orders():
    if not sheets:
        return jsonify({"orders": []})

    return jsonify({
        "orders": sheets.get_recent_orders_with_row()
    })


@app.route("/api/update_order", methods=["POST"])
def api_update_order():
    data = request.get_json(force=True)

    sheets.update_order_row(
        row=int(data["row"]),
        product=data.get("product", ""),
        brand=data.get("brand", ""),
        quantity=data.get("quantity", 0),
        price=data.get("price", 0)
    )

    return jsonify({"ok": True})


@app.route("/api/delete_order", methods=["POST"])
def api_delete_order():
    data = request.get_json(force=True)
    sheets.delete_order_row(int(data["row"]))
    return jsonify({"ok": True})

@app.route("/api/undo_delete_order", methods=["POST"])
def api_undo_delete_order():
    data = request.get_json(force=True)

    sheets.restore_order_row(
        row=int(data["row"]),
        data=data["data"]
    )
    return jsonify({"ok": True})


if __name__ == "__main__":
    app.run(port=8000, debug=True)