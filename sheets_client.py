import os
import datetime
import gspread
from google.oauth2.service_account import Credentials
import time
import random
from googleapiclient.errors import HttpError



SERVICE_ACCOUNT_FILE = "service_account.json"
SHEET_ID = "1WcNZPg8upKAn1eVlpM2faaHl1HxldEz4p8upgmGSOas"

COL_ORDER = "Order Number"
COL_DATE = "Date"
COL_COMPANY = "Company"
COL_PRODUCT = "Product"
COL_BRAND = "Brand"
COL_QUANTITY = "Quantity"
COL_PRICE = "Price"
COL_BALANCE = "Balance Order"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]


class SheetsClient:
    def _retry(self, fn, retries=3):
        for i in range(retries):
            try:
                return fn()
            except HttpError as e:
                if i == retries - 1:
                    raise
                sleep = (2 ** i) + random.random()
                time.sleep(sleep)

    def __init__(self):
        creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES
        )
        self.client = gspread.authorize(creds)
        self.sheet = self.client.open_by_key(SHEET_ID).worksheet("orders")
        self.requirements_sheet = self.client.open_by_key(SHEET_ID).worksheet("requirement")
        
        self._cache = {}
        self._cache_ttl = 15  # seconds (safe)

        try:
            self.dispatch_ws = self.client.open_by_key(SHEET_ID).worksheet("dispatch")
        except:
            self.dispatch_ws = self.client.open_by_key(SHEET_ID).add_worksheet(
                title="dispatch", rows=1000, cols=10
            )
            self.dispatch_ws.append_row(
                ["Date", "Company", "Product", "Quantity", "Order Number"]
            )
    def _cached(self, key, fn):
        now = time.time()

        if key in self._cache:
            value, ts = self._cache[key]
            if now - ts < self._cache_ttl:
                return value

        value = fn()
        self._cache[key] = (value, now)
        return value
    def _invalidate_cache(self):
        self._cache.clear()


    # ---------------- LOAD LISTS ----------------
    def load_lists(self):
        ss = self.client.open_by_key(SHEET_ID)
        out = {"products": [], "companies": [], "brands": []}

        for k in out.keys():
            try:
                ws = ss.worksheet(k)
                out[k] = ws.col_values(1)[1:]
            except:
                pass

        return out
    def _norm_company(self, s):
        return (
            (s or "")
            .strip()
            .lower()
            .replace("&", "and")
            .replace(" ", "")
        )
    

    # ---------------- DISPATCH ----------------
    def add_dispatch(self, company, product, quantity, order_number):
        self.dispatch_ws.append_row([
            datetime.date.today().isoformat(),
            company,
            product,
            quantity,
            order_number
        ], value_input_option="USER_ENTERED")
        self._invalidate_cache()

    # ---------------- AGGREGATIONS ----------------
    def _dispatch_map(self):
        rows = self.dispatch_ws.get_all_values()
        dispatch = {}

        for r in rows[1:]:
            if len(r) < 5:
                continue

            serial = (r[4] or "").strip()
            product = self._norm(r[2])

            try:
                qty = int(float(r[3]))
            except Exception:
                continue

            if not serial or not product:
                continue

            key = (serial, product)
            dispatch[key] = dispatch.get(key, 0) + qty

        return dispatch



    def get_orders_by_party(self, company):
        dispatch = self._dispatch_map()
        out = []

        rows = self._cached("orders_rows", lambda: self.sheet.get_all_values())


        for r in rows[1:]:
            if len(r) < 6:
                continue

            serial = (r[0] or "").strip()
            party = (r[2] or "").strip()
            product = r[3] or ""

            if self._norm_company(party) != self._norm_company(company):
                continue

            try:
                ordered = int(float(r[5]))
            except Exception:
                ordered = 0

            dispatched = dispatch.get((serial, self._norm(product)), 0)
            remaining = ordered - dispatched

            # 🔑 THIS LINE FIXES "already dispatched but still visible"
            if remaining <= 0:
                continue

            out.append({
                "company": party,
                "product": product,
                "serial": serial,
                "ordered": ordered,
                "dispatched": dispatched,
                "remaining": remaining,
                "price": r[6] if len(r) > 6 else ""
            })

        return out


    def get_orders_by_product(self, product):
        dispatch = self._dispatch_map()
        out = []

        target = self._norm(product)
        rows = self._cached("orders_rows", lambda: self.sheet.get_all_values())


        for r in rows[1:]:
            if len(r) < 6:
                continue

            serial = (r[0] or "").strip()
            prod = r[3] or ""

            if self._norm(prod) != target:
                continue

            party = r[2] or ""

            try:
                ordered = int(float(r[5]))
            except Exception:
                ordered = 0

            dispatched = dispatch.get((serial, target), 0)
            remaining = ordered - dispatched

            if remaining <= 0:
                continue

            out.append({
                "company": party,          
                "product": prod,
                "serial": serial,
                "ordered": ordered,
                "dispatched": dispatched,
                "remaining": remaining,
                "price": r[6] if len(r) > 6 else ""
            })

        return out






    def get_pivot_data(self, product_filter="", party_filter=""):
        rows = self._cached("orders_rows", lambda: self.sheet.get_all_values())

        dispatch = self._dispatch_map()

        data = {}

        product_filter = product_filter.lower()
        party_filter = party_filter.lower()

        for r in rows[1:]:
            

            if len(r) < 6:
                continue

            # Column mapping
            serial = r[0]
            date = r[1]
            company = r[2].strip()
            product = r[3]

            if not date.strip():
                continue  # ignore empty rows

            product_filters = [p for p in product_filter.split(",") if p]
            party_filters = [p for p in party_filter.split(",") if p]

            if product_filters and not any(p in product.lower() for p in product_filters):
                continue

            if party_filters and not any(p in company.lower() for p in party_filters):
                continue

            try:
                ordered = int(float(r[5]))
            except Exception:

                ordered = 0

            dispatched = dispatch.get((serial, self._norm(product)), 0)
            pending = ordered - dispatched
            if pending <= 0:
                continue

            data.setdefault(company, {})
            data[company][product] = data[company].get(product, 0) + pending

        products = sorted({p for c in data.values() for p in c})
        parties = sorted(data.keys())

        pivot = []
        for party in parties:
            row = []
            for product in products:
                row.append(data[party].get(product, 0))
            pivot.append(row)

        return {
            "products": products,
            "parties": parties,
            "pivot": pivot
        }

    def get_recent_orders(self, limit=50):
        rows = self._cached("orders_rows", lambda: self.sheet.get_all_values())

        if len(rows) <= 1:
            return []

        data = []

        for r in rows[1:]:
            # Column B = Date
            if len(r) < 2 or not r[1].strip():
                continue

            serial = r[0] if len(r) > 0 else ""

            data.append({
                "serial": serial,
                "date": r[1],
                "company": r[2] if len(r) > 2 else "",
                "product": r[3] if len(r) > 3 else "",
                "brand": r[4] if len(r) > 4 else "",
                "quantity": r[5] if len(r) > 5 else "",
                "price": r[6] if len(r) > 6 else "",
                "total": (
                    float(r[5]) * float(r[6])
                    if len(r) > 6 and r[5] and r[6]
                    else ""
                )
            })

        data.reverse()
        return data[:limit]

    # ---------------- ADD ORDER ----------------
    def add_order(self, company, product, quantity, price, brand):
        """
        Insert order in the first row where Date column is empty.
        Assumes:
        Column A = Serial (already filled / formula)
        Column B = Date
        """
        sheet = self.sheet
        all_vals = sheet.get_all_values()

        # Find first empty Date cell (Column B)
        target_row = None
        for i in range(1, len(all_vals)):
            row = all_vals[i]
            if len(row) < 2 or not row[1].strip():
                target_row = i + 1  # sheets are 1-indexed
                break

        if not target_row:
            target_row = len(all_vals) + 1

        today = datetime.date.today().isoformat()

        # Columns:
        # B = Date, C = Company, D = Product, E = Brand, F = Quantity, G = Price
        sheet.update(
            f"B{target_row}:G{target_row}",
            [[today, company, product, brand, int(quantity), float(price)]],
            value_input_option="USER_ENTERED"
        )
        

        # ---------------- RECENT ORDERS ----------------
    def _norm(self, s):
        return (s or "").strip().lower().replace(" ", "")


    def get_recent_orders_with_row(self, limit=15):
        rows = self._cached("orders_rows", lambda: self.sheet.get_all_values())
        out = []

        for i, r in enumerate(rows[1:], start=2):  # sheet rows are 1-indexed
            if len(r) < 2 or not r[1].strip():
                continue

            out.append({
                "row": i,                  # 👈 IMPORTANT
                "serial": r[0],
                "date": r[1],
                "company": r[2] if len(r) > 2 else "",
                "product": r[3] if len(r) > 3 else "",
                "brand": r[4] if len(r) > 4 else "",
                "quantity": r[5] if len(r) > 5 else "",
                "price": r[6] if len(r) > 6 else "",
                "total": (
                    float(r[5]) * float(r[6]) * 1.05
                    if len(r) > 6 and r[5] and r[6]
                    else ""
                )
            })

        out.reverse()
        return out[:limit]


    def update_order_row(self, row, product, brand, quantity, price):
        self.sheet.update(
        f"D{row}:G{row}",
        [[
            product,
            brand,
            int(quantity),
            float(price)
        ]],
        value_input_option="USER_ENTERED"
    )

        self._invalidate_cache()



    def delete_order_row(self, row):
        # Soft delete: clear Date + data, keep serial & formulas
        self.sheet.update(
            f"B{row}:G{row}",
            [["", "", "", "", "", ""]],
            value_input_option="USER_ENTERED"
        )
        self._invalidate_cache()
    def restore_order_row(self, row, data):
        """
        data = dict with keys:
        date, company, product, brand, quantity, price
        """
        self.sheet.update(
            f"B{row}:G{row}",
            [[
                data.get("date", ""),
                data.get("company", ""),
                data.get("product", ""),
                data.get("brand", ""),
                data.get("quantity", ""),
                data.get("price", "")
            ]],
            value_input_option="USER_ENTERED"
        )
        self._invalidate_cache()
    def get_inventory_requirements(self):
        if "requirements_rows" in self._cache:
            rows, ts = self._cache["requirements_rows"]
            if time.time() - ts < self._cache_ttl:
                return rows
        rows = self.requirements_sheet.get_all_records()
        # rows must include: product, width, thickness, weight
        return rows