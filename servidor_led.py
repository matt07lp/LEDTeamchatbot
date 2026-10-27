"""
Servidor MCP - Venta y renta de pantallas LED (v0.3)

Tools:
  1. evaluate_request        -> classify BUY / RENT / unsupported, creates the lead
  2. handle_purchase_request -> schedule call | showroom visit | send quotation
  3. handle_rental_request   -> dates + measurements + location -> quotation

Storage (Excel):
  data/leads.xlsx    -> sheets: Leads, Appointments, Quotations   (written by the bot)
  data/pricing.xlsx  -> sheets: Purchase, Rental, Delivery, Settings (edited by your team)

Install:  pip install mcp openpyxl
Run:      python servidor_led.py            (stdio)
"""

import functools
import json
import os
import smtplib
import threading
import uuid
from datetime import datetime, date, time, timedelta
from email.message import EmailMessage
from pathlib import Path
from typing import Literal, Optional

try:
    from mcp.server.mcpserver import MCPServer          # SDK nuevo (formato del profesor)
except ImportError:
    from mcp.server.fastmcp import FastMCP as MCPServer  # SDK 1.x: misma clase, otro nombre
from openpyxl import Workbook, load_workbook

DATA_DIR = Path(os.getenv(
    "LED_DATA_DIR",
    "/Users/migueltorres/Documents/Documents - Miguel MacBook Pro/"
    "Diplomado AI and LLM for FM/Modulo 4/Proyecto final",
))
LEADS_FILE = DATA_DIR / "leads.xlsx"
PRICING_FILE = DATA_DIR / "pricing.xlsx"

# El nombre aparece cuando un Host MCP lista los servidores conectados.
servidor = MCPServer("Servidor LED Ventas y Renta")
_lock = threading.Lock()

# ---------------------------------------------------------------------------
# EXCEL LAYER
# ---------------------------------------------------------------------------
LEADS_HEADERS = {
    "Leads": ["lead_id", "created_at", "updated_at", "intent", "summary",
              "name", "email", "phone", "company", "status"],
    "Appointments": ["appointment_id", "lead_id", "type", "datetime", "status"],
    "Quotations": ["quote_id", "lead_id", "type", "details_json", "subtotal",
                   "tax", "total", "currency", "valid_until", "delivery_status"],
}

# PLACEHOLDER prices: replace with your real price list in pricing.xlsx
PRICING_DEFAULTS = {
    "Purchase": (["product_id", "name", "screen_type", "pixel_pitch_mm", "price_per_m2"], [
        ["P2.5-IN", "Indoor P2.5", "indoor", 2.5, 1800],
        ["P3-IN", "Indoor P3", "indoor", 3.0, 1500],
        ["P4-IN", "Indoor P4", "indoor", 4.0, 1200],
        ["P5-OUT", "Outdoor P5", "outdoor", 5.0, 1400],
        ["P6-OUT", "Outdoor P6", "outdoor", 6.0, 1100],
        ["P10-OUT", "Outdoor P10", "outdoor", 10.0, 800],
    ]),
    "Rental": (["screen_type", "daily_rate_per_m2", "min_charge"], [
        ["indoor", 45, 300],
        ["outdoor", 40, 300],
    ]),
    "Delivery": (["city", "fee"], [
        ["Mexico City", 150],
        ["DEFAULT", 500],
    ]),
    "Settings": (["key", "value"], [
        ["currency", "MXN"],
        ["tax_rate", 0.16],
        ["quote_valid_days", 15],
        ["showroom_open", "09:00"],
        ["showroom_close", "18:00"],
        ["showroom_days", "0,1,2,3,4"],   # 0=Mon ... 6=Sun
        ["slot_minutes", 60],
    ]),
}


# Support-team contact shown to customers whose request is not supported (repairs, etc.).
# PLACEHOLDERS: replace them with your real data in pricing.xlsx > Settings.
AJUSTES_SOPORTE = [
    ["support_phone", "+52 55 0000 0000"],
    ["support_email", "soporte@example.com"],
    ["support_hours", "Lunes a viernes, 9:00 a 18:00"],
]


def _completar_ajustes() -> None:
    """Add any missing support-contact rows to an existing Settings sheet (never overwrites)."""
    try:
        wb = load_workbook(PRICING_FILE)
        ws = wb["Settings"]
        existentes = {c.value for c in ws["A"] if c.value}
        faltan = [r for r in AJUSTES_SOPORTE if r[0] not in existentes]
        if faltan:
            for r in faltan:
                ws.append(r)
            wb.save(PRICING_FILE)
    except Exception:
        pass  # never block the tools because of this convenience


def ensure_files() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not LEADS_FILE.exists():
        wb = Workbook()
        wb.remove(wb.active)
        for sheet, headers in LEADS_HEADERS.items():
            wb.create_sheet(sheet).append(headers)
        wb.save(LEADS_FILE)
    if not PRICING_FILE.exists():
        wb = Workbook()
        wb.remove(wb.active)
        for sheet, (headers, rows) in PRICING_DEFAULTS.items():
            ws = wb.create_sheet(sheet)
            ws.append(headers)
            for r in rows:
                ws.append(r)
        wb.save(PRICING_FILE)
    _completar_ajustes()


def _rows(wb, sheet: str) -> list[dict]:
    ws = wb[sheet]
    headers = [c.value for c in ws[1]]
    return [dict(zip(headers, [c.value for c in row]))
            for row in ws.iter_rows(min_row=2)
            if any(c.value is not None for c in row)]


def _append(wb, sheet: str, data: dict) -> None:
    wb[sheet].append([data.get(h) for h in LEADS_HEADERS[sheet]])


def _update_lead(wb, lead_id: str, **fields) -> None:
    ws = wb["Leads"]
    headers = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2):
        if row[0].value == lead_id:
            fields["updated_at"] = _now()
            for k, v in fields.items():
                row[headers.index(k)].value = v
            return


def _get_lead(wb, lead_id: str) -> Optional[dict]:
    return next((r for r in _rows(wb, "Leads") if r["lead_id"] == lead_id), None)


def _settings() -> dict:
    wb = load_workbook(PRICING_FILE, data_only=True)
    return {r["key"]: r["value"] for r in _rows(wb, "Settings")}


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _to_time(v) -> time:
    return v if isinstance(v, time) else datetime.strptime(str(v), "%H:%M").time()


def safe_tool(fn):
    """Serialize Excel access and turn failures into messages the LLM can use."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        with _lock:
            try:
                ensure_files()
                return fn(*args, **kwargs)
            except PermissionError:
                return {"error": "A data file is open in another program (Excel). "
                                 "Tell the customer to try again in a moment."}
            except Exception as e:
                return {"error": f"Internal error: {e}"}
    return wrapper


# ---------------------------------------------------------------------------
# BUSINESS LOGIC
# ---------------------------------------------------------------------------
def _missing_contact(name, email, phone) -> list[str]:
    return [k for k, v in {"name": name, "email": email, "phone": phone}.items()
            if not (v or "").strip()]


def _check_showroom_slot(when: datetime, appointments: list[dict], s: dict) -> Optional[str]:
    days = [int(d) for d in str(s["showroom_days"]).split(",")]
    if when.weekday() not in days:
        return "The showroom is closed that day."
    if not (_to_time(s["showroom_open"]) <= when.time() <= _to_time(s["showroom_close"])):
        return (f"The showroom is open {_to_time(s['showroom_open']):%H:%M}-"
                f"{_to_time(s['showroom_close']):%H:%M}.")
    slot = timedelta(minutes=int(s["slot_minutes"]))
    for a in appointments:
        if a["type"] == "showroom_visit" and a["status"] == "scheduled":
            if abs(datetime.fromisoformat(a["datetime"]) - when) < slot:
                return "That time slot is already taken."
    return None


def _finalize_quote(wb, lead: dict, qtype: str, details: dict, subtotal: float) -> dict:
    s = _settings()
    tax = round(subtotal * float(s["tax_rate"]), 2)
    total = round(subtotal + tax, 2)
    valid_until = (date.today() + timedelta(days=int(s["quote_valid_days"]))).isoformat()
    quote_id = f"Q-{uuid.uuid4().hex[:6].upper()}"
    delivery = _deliver_quote(lead, quote_id, qtype, details, subtotal, tax, total,
                              s["currency"], valid_until)
    _append(wb, "Quotations", {
        "quote_id": quote_id, "lead_id": lead["lead_id"], "type": qtype,
        "details_json": json.dumps(details), "subtotal": round(subtotal, 2),
        "tax": tax, "total": total, "currency": s["currency"],
        "valid_until": valid_until, "delivery_status": delivery})
    _update_lead(wb, lead["lead_id"], status=f"{qtype}_quoted")
    return {"quote_id": quote_id, "subtotal": round(subtotal, 2), "tax": tax,
            "total": total, "currency": s["currency"], "valid_until": valid_until,
            "delivery_status": delivery,
            "next_step": ("Tell the customer the quotation was emailed to them."
                          if delivery == "emailed" else
                          "Tell the customer the quotation is ready and our team "
                          "will send it shortly.")}


def _deliver_quote(lead, quote_id, qtype, details, subtotal, tax, total, currency, valid) -> str:
    """Email the quote via SMTP if configured (SMTP_HOST/USER/PASSWORD/FROM env vars)."""
    host = os.getenv("SMTP_HOST")
    if not host:
        return "pending_manual_send"
    msg = EmailMessage()
    msg["Subject"] = f"Your LED screen {qtype} quotation {quote_id}"
    msg["From"] = os.getenv("SMTP_FROM", os.getenv("SMTP_USER", ""))
    msg["To"] = lead["email"]
    lines = "\n".join(f"  {k}: {v}" for k, v in details.items())
    msg.set_content(
        f"Hello {lead['name']},\n\nQuotation {quote_id}\n{lines}\n\n"
        f"Subtotal: {subtotal:,.2f} {currency}\nTax: {tax:,.2f} {currency}\n"
        f"TOTAL: {total:,.2f} {currency}\nValid until: {valid}\n")
    try:
        with smtplib.SMTP(host, int(os.getenv("SMTP_PORT", "587"))) as smtp:
            smtp.starttls()
            smtp.login(os.getenv("SMTP_USER", ""), os.getenv("SMTP_PASSWORD", ""))
            smtp.send_message(msg)
        return "emailed"
    except Exception:
        return "email_failed"


def _contacto_soporte() -> dict:
    """Support-team contact, read from pricing.xlsx > Settings."""
    try:
        ajustes = _settings()
    except Exception:
        ajustes = {}
    return {k: str(ajustes.get(f"support_{k}") or "").strip() for k in ("phone", "email", "hours")}


# ---------------------------------------------------------------------------
# TOOL 1: EVALUATE REQUEST
# ---------------------------------------------------------------------------
@servidor.tool()
@safe_tool
def evaluate_request(intent: Literal["buy", "rent", "other"], confidence: float,
                     summary: str, lead_id: Optional[str] = None) -> dict:
    """Classify the customer's request. Call this FIRST, once the intent is clear.
    Only 'buy' and 'rent' of LED screens are supported; use 'other' for anything else.
    Returns a lead_id that MUST be passed to the other tools.

    Args:
        intent: buy, rent, or other (unsupported).
        confidence: 0.0-1.0 certainty about the intent.
        summary: One-line summary of what the customer asked.
        lead_id: Pass the existing lead_id if the intent is being corrected.
    """
    if intent != "other" and confidence < 0.7:
        return {"status": "needs_clarification",
                "next_step": "Ask ONE short question: does the customer want to buy or rent?"}

    stored = intent if intent != "other" else "unsupported"
    wb = load_workbook(LEADS_FILE)
    if lead_id and _get_lead(wb, lead_id):
        _update_lead(wb, lead_id, intent=stored, summary=summary)
    else:
        lead_id = f"L-{uuid.uuid4().hex[:8].upper()}"
        _append(wb, "Leads", {"lead_id": lead_id, "created_at": _now(),
                              "updated_at": _now(), "intent": stored,
                              "summary": summary, "status": "new"})
    wb.save(LEADS_FILE)

    if stored == "unsupported":
        soporte = _contacto_soporte()
        if soporte["phone"] or soporte["email"]:
            siguiente = ("Politely explain that this chat only handles buying or renting LED screens. "
                         "For anything else (repairs, technical support, other services) give the customer "
                         "the support team contact in 'support_contact', exactly as written, and invite "
                         "them to call or write. Do NOT invent prices, repair times or any other contact data.")
        else:
            siguiente = ("Politely explain that this chat only handles buying or renting LED screens. "
                         "No support contact is configured, so do NOT invent one.")
        return {"status": "unsupported", "lead_id": lead_id,
                "support_contact": soporte, "next_step": siguiente}
    if stored == "buy":
        return {"status": "supported", "intent": "buy", "lead_id": lead_id,
                "next_step": "Ask whether they prefer: (a) a call, (b) a showroom visit, "
                             "or (c) a quotation. Then call handle_purchase_request."}
    return {"status": "supported", "intent": "rent", "lead_id": lead_id,
            "next_step": "Collect rental dates, screen width/height in meters, indoor or "
                         "outdoor, event location, and contact info. Then call "
                         "handle_rental_request."}


# ---------------------------------------------------------------------------
# TOOL 2: PURCHASE REQUEST
# ---------------------------------------------------------------------------
@servidor.tool()
@safe_tool
def handle_purchase_request(
    lead_id: str,
    action: Literal["schedule_call", "showroom_visit", "send_quotation"],
    name: str, email: str, phone: str, company: str = "",
    preferred_datetime: str = "",
    screen_type: Optional[Literal["indoor", "outdoor"]] = None,
    pixel_pitch_mm: Optional[float] = None,
    width_m: float = 0, height_m: float = 0, quantity: int = 1,
) -> dict:
    """For BUY leads only. Schedule a call, schedule a showroom visit, or send a quotation.
    Requires lead_id plus the customer's name, email and phone. Never invent prices.

    Args:
        lead_id: From evaluate_request.
        action: schedule_call, showroom_visit, or send_quotation.
        preferred_datetime: ISO 8601 (e.g. 2026-10-20T15:00). Required for call/visit.
        screen_type: indoor or outdoor (required for send_quotation).
        pixel_pitch_mm: e.g. 2.5, 3, 4, 5, 6, 10 (required for send_quotation).
        width_m: Screen width in meters (required for send_quotation).
        height_m: Screen height in meters (required for send_quotation).
        quantity: Number of screens.
    """
    wb = load_workbook(LEADS_FILE)
    lead = _get_lead(wb, lead_id)
    if not lead:
        return {"error": "Unknown lead_id. Call evaluate_request first."}
    if lead["intent"] != "buy":
        return {"error": f"This lead's intent is '{lead['intent']}', not 'buy'."}
    missing = _missing_contact(name, email, phone)
    if missing:
        return {"error": f"Missing contact fields: {missing}. Ask the customer."}

    _update_lead(wb, lead_id, name=name, email=email, phone=phone, company=company)
    lead.update(name=name, email=email, phone=phone)

    # --- call / showroom ---
    if action in ("schedule_call", "showroom_visit"):
        if not preferred_datetime:
            return {"error": "preferred_datetime is required. Ask for a date and time."}
        try:
            when = datetime.fromisoformat(preferred_datetime).replace(tzinfo=None)
        except ValueError:
            return {"error": "Invalid datetime. Use ISO 8601, e.g. 2026-10-20T15:00"}
        if when <= datetime.now():
            return {"error": "That time is in the past. Ask for another."}
        if action == "showroom_visit":
            problem = _check_showroom_slot(when, _rows(wb, "Appointments"), _settings())
            if problem:
                return {"error": problem + " Ask the customer for another time."}
        appt_id = f"A-{uuid.uuid4().hex[:6].upper()}"
        _append(wb, "Appointments", {"appointment_id": appt_id, "lead_id": lead_id,
                                     "type": action, "datetime": when.isoformat(),
                                     "status": "scheduled"})
        _update_lead(wb, lead_id, status=f"{action}_booked")
        wb.save(LEADS_FILE)
        return {"status": "scheduled", "appointment_id": appt_id, "type": action,
                "when": when.isoformat(),
                "next_step": "Confirm the appointment details to the customer."}

    # --- quotation ---
    if not screen_type or width_m <= 0 or height_m <= 0 or quantity < 1:
        return {"error": "For a quotation I need screen_type, width_m, height_m and "
                         "quantity. Ask the customer."}
    products = [p for p in _rows(load_workbook(PRICING_FILE, data_only=True), "Purchase")
                if p["screen_type"] == screen_type]
    if pixel_pitch_mm is None:
        return {"error": "pixel_pitch_mm is required. Offer these options.",
                "options": [{"pixel_pitch_mm": p["pixel_pitch_mm"], "name": p["name"]}
                            for p in products]}
    product = next((p for p in products
                    if abs(float(p["pixel_pitch_mm"]) - pixel_pitch_mm) < 0.01), None)
    if not product:
        return {"error": "No product with that pixel pitch. Offer these options.",
                "options": [{"pixel_pitch_mm": p["pixel_pitch_mm"], "name": p["name"]}
                            for p in products]}

    area = round(width_m * height_m * quantity, 2)
    subtotal = area * float(product["price_per_m2"])
    details = {"product": product["name"], "size_m": f"{width_m} x {height_m}",
               "quantity": quantity, "total_area_m2": area,
               "price_per_m2": product["price_per_m2"]}
    result = _finalize_quote(wb, lead, "purchase", details, subtotal)
    wb.save(LEADS_FILE)
    return {"status": "quotation_created", **result}


# ---------------------------------------------------------------------------
# TOOL 3: RENTAL REQUEST
# ---------------------------------------------------------------------------
@servidor.tool()
@safe_tool
def handle_rental_request(
    lead_id: str, start_date: str, end_date: str, location: str,
    screen_type: Literal["indoor", "outdoor"],
    width_m: float, height_m: float,
    name: str, email: str, phone: str, company: str = "", quantity: int = 1,
) -> dict:
    """For RENT leads only. Needs rental dates, screen measurements in meters, indoor
    or outdoor, event location, and contact info. Creates and sends a rental quotation.

    Args:
        lead_id: From evaluate_request.
        start_date: First rental day, YYYY-MM-DD.
        end_date: Last rental day, YYYY-MM-DD (same as start for a 1-day rental).
        location: City and venue of the event.
        screen_type: indoor or outdoor.
        width_m: Screen width in meters.
        height_m: Screen height in meters.
        quantity: Number of screens.
    """
    wb = load_workbook(LEADS_FILE)
    lead = _get_lead(wb, lead_id)
    if not lead:
        return {"error": "Unknown lead_id. Call evaluate_request first."}
    if lead["intent"] != "rent":
        return {"error": f"This lead's intent is '{lead['intent']}', not 'rent'."}
    missing = _missing_contact(name, email, phone)
    if missing:
        return {"error": f"Missing contact fields: {missing}. Ask the customer."}
    try:
        start, end = date.fromisoformat(start_date), date.fromisoformat(end_date)
    except ValueError:
        return {"error": "Dates must be YYYY-MM-DD. Ask the customer to clarify."}
    if start < date.today():
        return {"error": "The start date is in the past."}
    if end < start:
        return {"error": "The end date is before the start date."}
    if width_m <= 0 or height_m <= 0 or quantity < 1:
        return {"error": "Dimensions and quantity must be positive."}

    pricing = load_workbook(PRICING_FILE, data_only=True)
    rate = next((r for r in _rows(pricing, "Rental") if r["screen_type"] == screen_type), None)
    if not rate:
        return {"error": f"No rental rate configured for '{screen_type}'."}
    delivery_rows = _rows(pricing, "Delivery")
    fee_row = (next((d for d in delivery_rows
                     if str(d["city"]).lower() in location.lower()
                     and d["city"] != "DEFAULT"), None)
               or next(d for d in delivery_rows if d["city"] == "DEFAULT"))

    days = (end - start).days + 1
    area = round(width_m * height_m * quantity, 2)
    rental_cost = max(area * days * float(rate["daily_rate_per_m2"]),
                      float(rate["min_charge"]))
    subtotal = rental_cost + float(fee_row["fee"])
    details = {"dates": f"{start_date} to {end_date}", "days": days,
               "location": location, "screen_type": screen_type,
               "size_m": f"{width_m} x {height_m}", "quantity": quantity,
               "total_area_m2": area, "rental_cost": round(rental_cost, 2),
               "delivery_fee": float(fee_row["fee"])}

    _update_lead(wb, lead_id, name=name, email=email, phone=phone, company=company)
    lead.update(name=name, email=email, phone=phone)
    result = _finalize_quote(wb, lead, "rental", details, subtotal)
    wb.save(LEADS_FILE)
    return {"status": "quotation_created", **result}


# ---------------------------------------------------------------------------
# PROMPT (the host/chatbot loads this as its system instructions)
# ---------------------------------------------------------------------------
def sales_assistant() -> str:
    """System instructions for the LED screens sales/rental assistant."""
    return f"""You are the virtual assistant of a company that sells and rents LED screens.
Today is {date.today().isoformat()}.

- Be friendly and concise. Ask one or two questions at a time.
- First find out if the customer wants to BUY or RENT, then call evaluate_request.
- Anything other than buying or renting is unsupported: say so politely.
- BUY: offer three options: a call, a showroom visit, or a quotation.
- RENT: collect dates, screen measurements (meters), indoor/outdoor, location, contact info.
- Always pass the lead_id returned by evaluate_request to the other tools.
- Never invent prices, stock or availability. Only report prices returned by the tools.
- Do not call an action tool until you have ALL its required information.
- If a tool returns an error, follow its instructions and ask the customer again.
"""


# Expose the instructions as an MCP prompt (skipped if the SDK version has no prompts)
if hasattr(servidor, "prompt"):
    servidor.prompt()(sales_assistant)


# Punto de entrada: cuando el cliente lo lanza como subproceso, corre en modo stdio.
if __name__ == "__main__":
    ensure_files()
    # Later, for the web page, the transport can change to "streamable-http"
    servidor.run(transport="stdio")
