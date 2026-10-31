import os
import re
import logging
from datetime import datetime, timezone
from io import BytesIO

from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters,
)
from supabase import create_client, Client
from openpyxl import load_workbook

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("auto_format_bot")

BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
SUPABASE_URL = os.getenv("SUPABASE_URL", "").strip()
SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY", "").strip()
ADMIN_IDS = {
    int(x.strip()) for x in os.getenv("ADMIN_IDS", "").split(",")
    if x.strip().isdigit()
}
OWNER_ID = int(os.getenv("BOT_OWNER_ID", "0") or 0)

if not BOT_TOKEN:
    raise RuntimeError("Environment variable TELEGRAM_BOT_TOKEN belum diisi.")
if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
    raise RuntimeError("SUPABASE_URL dan SUPABASE_SERVICE_ROLE_KEY wajib diisi.")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)

MAIN_MENU = [
    [KeyboardButton("👤 Akun"), KeyboardButton("✨ Buat Format")],
    [KeyboardButton("📄 Hasil Format"), KeyboardButton("🧾 Hasil Format + Akun")],
    [KeyboardButton("📂 Data Saya"), KeyboardButton("📊 Hasil Data Saya")],
    [KeyboardButton("🛡️ Panel Admin")],
]
FORMAT_MENU = [
    [KeyboardButton("1. Buat Manual"), KeyboardButton("2. Buat File Excel")],
    [KeyboardButton("3. Setting Template")],
    [KeyboardButton("4. Setting Kode Atas"), KeyboardButton("5. Setting Kode Bawah")],
    [KeyboardButton("6. Kode Aktif Saya")],
    [KeyboardButton("8. Kembali ke Menu Utama")],
]
DATA_MENU = [
    [KeyboardButton("1. Kirim Data"), KeyboardButton("2. Cari Data")],
    [KeyboardButton("3. Hasil Data Saya")],
    [KeyboardButton("4. Kembali ke Menu Utama")],
]
MY_DATA_RESULTS_MENU = [
    [KeyboardButton("1. Data Format"), KeyboardButton("2. Format + Akun")],
    [KeyboardButton("3. Hapus Data Saya")],
    [KeyboardButton("4. Kembali")],
]
FORMAT_RESULTS_MENU = [
    [KeyboardButton("1. Cari Data"), KeyboardButton("2. Tampilkan Semua")],
    [KeyboardButton("3. Riwayat Penjualan")],
    [KeyboardButton("4. Kembali ke Menu Utama")],
]
ADMIN_MENU = [
    [KeyboardButton("1. Tambah Admin"), KeyboardButton("2. Tambah Grup")],
    [KeyboardButton("3. Daftar Admin"), KeyboardButton("4. Daftar Grup")],
    [KeyboardButton("5. Tambah Anggota"), KeyboardButton("6. Hapus Anggota")],
    [KeyboardButton("7. Hapus Grup")],
    [KeyboardButton("8. Kembali ke Menu Utama")],
]

DEFAULT_TEMPLATE = ""

HELP_TEXT = (
    "Kirim data berurutan menggunakan spasi, tanpa koma. Contoh:\n\n"
    "Bandung Antapani Regol 12 juta PT Makmur Jaya\n\n"
    "Bot akan mengisi kolom berdasarkan urutan placeholder {1}, {2}, {3}, dan seterusnya. "
    "Bot mengenali beberapa kelompok umum seperti angka + satuan (12 juta) dan nama perusahaan "
    "yang diawali PT. Untuk nilai multi-kata lain yang ambigu, kirim satu nilai per baris agar "
    "tidak tertukar."
)

def main_keyboard():
    return ReplyKeyboardMarkup(MAIN_MENU, resize_keyboard=True)

def format_keyboard():
    return ReplyKeyboardMarkup(FORMAT_MENU, resize_keyboard=True)

def data_keyboard():
    return ReplyKeyboardMarkup(DATA_MENU, resize_keyboard=True)

def my_data_results_keyboard():
    return ReplyKeyboardMarkup(MY_DATA_RESULTS_MENU, resize_keyboard=True)

def admin_keyboard():
    return ReplyKeyboardMarkup(ADMIN_MENU, resize_keyboard=True)

def format_results_keyboard():
    return ReplyKeyboardMarkup(FORMAT_RESULTS_MENU, resize_keyboard=True)

def is_owner(user_id: int) -> bool:
    # Set BOT_OWNER_ID in Railway Variables to the Telegram ID of the bot creator.
    return bool(OWNER_ID) and user_id == OWNER_ID

def is_admin(user_id: int) -> bool:
    if is_owner(user_id) or user_id in ADMIN_IDS:
        return True
    try:
        rows = supabase.table("bot_users").select("is_admin").eq("telegram_id", user_id).limit(1).execute().data
        return bool(rows and rows[0].get("is_admin"))
    except Exception:
        return False

def is_staff(user_id: int) -> bool:
    return is_owner(user_id) or is_admin(user_id)

def role_name(user_id: int) -> str:
    if is_owner(user_id):
        return "Pemilik Bot"
    if is_admin(user_id):
        return "Admin"
    return "Anggota"

def get_user(tg_user):
    uid = tg_user.id
    row = supabase.table("bot_users").select("*").eq("telegram_id", uid).limit(1).execute().data
    now = datetime.now(timezone.utc).isoformat()
    data = {
        "telegram_id": uid,
        "username": tg_user.username,
        "first_name": tg_user.first_name or "",
        "last_name": tg_user.last_name or "",
        "is_admin": is_admin(uid),
        "last_seen_at": now,
    }
    if row:
        supabase.table("bot_users").update(data).eq("telegram_id", uid).execute()
        return {**row[0], **data}
    data["created_at"] = now
    supabase.table("bot_users").insert(data).execute()
    return data

def get_settings(uid: int):
    rows = supabase.table("user_settings").select("*").eq("telegram_id", uid).limit(1).execute().data
    if rows:
        return rows[0]
    default = {
        "telegram_id": uid,
        "template_text": "",
        "code_prefix": "",
        "next_code_number": 1,
        "code_width": 3,
        "code_bottom": "",
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    supabase.table("user_settings").insert(default).execute()
    return default

def save_settings(uid: int, updates: dict):
    updates["updated_at"] = datetime.now(timezone.utc).isoformat()
    existing = supabase.table("user_settings").select("telegram_id").eq("telegram_id", uid).limit(1).execute().data
    if existing:
        supabase.table("user_settings").update(updates).eq("telegram_id", uid).execute()
    else:
        default = get_settings(uid)
        default.update(updates)
        supabase.table("user_settings").upsert(default, on_conflict="telegram_id").execute()

def parse_input(text: str, template: str = "") -> list[str]:
    """Parse ordered values. Multi-line values are unambiguous; one-line values use
    spaces plus conservative grouping for common patterns (amounts and PT companies).
    Completely arbitrary multi-word fields cannot be inferred perfectly without delimiters.
    """
    text = (text or "").strip()
    if not text:
        return []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    numbered = {}
    for line in lines:
        match = re.match(r"^\s*(\d{1,3})\s*[.)-]\s*(.*)$", line)
        if match:
            idx, value = int(match.group(1)), match.group(2).strip()
            if ":" in value:
                value = value.split(":", 1)[1].strip()
            numbered[idx] = value
    if numbered:
        highest = max(numbered)
        return [numbered.get(i, "") for i in range(1, highest + 1)]
    # Multiple lines are safest: one line = one field.
    if len(lines) > 1:
        if all(":" not in line for line in lines):
            return lines
        keyed = []
        for line in lines:
            if ":" in line:
                keyed.append(line.split(":", 1)[1].strip())
            else:
                keyed.append(line)
        return keyed
    # Support legacy comma-delimited input too, but it is not required.
    if "," in text:
        return [part.strip() for part in text.split(",") if part.strip()]
    tokens = text.split()
    if not tokens:
        return []
    # Group common multi-word patterns, preserving the original sequence.
    out = []
    i = 0
    while i < len(tokens):
        token = tokens[i]
        if token.upper() == "PT" and i + 1 < len(tokens):
            company = [tokens[i], tokens[i + 1]]
            i += 2
            # Common company suffix/name tokens following PT are kept together.
            while i < len(tokens) and not re.match(r"^\d+(?:[.,]\d+)?$", tokens[i]):
                company.append(tokens[i])
                i += 1
            out.append(" ".join(company))
            continue
        if re.match(r"^\d+(?:[.,]\d+)?$", token) and i + 1 < len(tokens) and re.match(r"^[A-Za-zÀ-ÿ]+$", tokens[i + 1]):
            out.append(token + " " + tokens[i + 1])
            i += 2
            continue
        out.append(token)
        i += 1
    # If the template has a known number of fields and we have excess tokens,
    # leave the values ordered; do not silently concatenate arbitrary fields.
    return out


def render_template(template: str, values: list[str]) -> str:
    # Replace longer placeholders first so {1} cannot affect {10}.
    result = template
    for index in range(len(values), 0, -1):
        result = result.replace("{" + str(index) + "}", values[index - 1])
    # Any placeholder not supplied stays visibly empty.
    result = re.sub(r"\{(\d+)\}", "", result)
    # Remove trailing spaces and excessive blank lines.
    result = "\n".join(line.rstrip() for line in result.splitlines())
    result = re.sub(r"\n{3,}", "\n\n", result).strip()
    return result

def code_for(settings: dict, number: int) -> str:
    prefix = (settings.get("code_prefix") or "").strip()
    if not prefix:
        return ""
    width = int(settings.get("code_width") or 3)
    return f"{prefix}_{number:0{width}d}".upper()

def split_regular_and_full(rendered: str, template: str):
    """Regular output stops at the PT field if present; custom templates without PT show all fields.
    For fully custom cutoffs, put a line containing {BATAS} in the template.
    The marker is omitted from the final output.
    """
    marker = "{BATAS}"
    if marker in template:
        template_before = template.split(marker, 1)[0]
        rendered_before = render_template(template_before, [])
        # Placeholder values have already been filled in rendered, so split rendered at marker position
        # by rendering each side separately is safer in caller; this function is a fallback.
        return rendered, rendered
    lines = rendered.splitlines()
    cutoff = None
    for i, line in enumerate(lines):
        if re.match(r"^\s*PT\s*:", line, flags=re.IGNORECASE):
            cutoff = i
            break
    if cutoff is None:
        return rendered, rendered
    return "\n".join(lines[:cutoff + 1]).strip(), rendered

def make_outputs(template: str, values: list[str], settings: dict, number: int):
    # Optional explicit cutoff marker {BATAS} divides Hasil Format and full result.
    if "{BATAS}" in template:
        before, after = template.split("{BATAS}", 1)
        regular = render_template(before, values)
        full = (regular + "\n" + render_template(after, values)).strip()
    else:
        full = render_template(template, values)
        regular, full = split_regular_and_full(full, template)
    code = code_for(settings, number)
    bottom = (settings.get("code_bottom") or "").strip()
    def decorate(body):
        pieces = []
        if code:
            pieces.append(code)
        if body:
            pieces.append(body)
        if bottom:
            pieces.append(bottom)
        return "\n".join(pieces).strip().upper()
    return decorate(regular), decorate(full), code

def next_number(settings: dict) -> int:
    return int(settings.get("next_code_number") or 1)

def store_record(uid: int, input_text: str, regular: str, full: str, code: str):
    return supabase.table("format_records").insert({
        "telegram_id": uid,
        "input_text": input_text,
        "regular_output": regular,
        "full_output": full,
        "generated_code": code,
        "is_saved": False,
        "created_at": datetime.now(timezone.utc).isoformat(),
    }).select("id,is_saved").execute().data

def record_actions(record_id: int, already_saved: bool = False):
    save_label = "✅ Sudah Disimpan" if already_saved else "💾 Simpan"
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✏️ Edit", callback_data=f"edit:{record_id}"),
            InlineKeyboardButton("🗑️ Hapus", callback_data=f"delete:{record_id}"),
            InlineKeyboardButton(save_label, callback_data=f"save:{record_id}"),
        ],
        [
            InlineKeyboardButton("🔄 Buat Lagi", callback_data="again"),
            InlineKeyboardButton("🏠 Menu Utama", callback_data="main_menu"),
        ],
    ])

def saved_result_actions(record_id: int, is_sold: bool = False, full: bool = True, history: bool = False):
    # Hasil Format biasa hanya menyediakan salin; riwayat juga hanya menyediakan salin.
    if history:
        return [InlineKeyboardButton("📋 Salin", callback_data=f"copyrec:{record_id}" if full else f"copyreg:{record_id}")]
    if not full:
        return [InlineKeyboardButton("📋 Salin", callback_data=f"copyreg:{record_id}")]
    sale_button = (
        InlineKeyboardButton("✅ TERJUAL", callback_data=f"soldinfo:{record_id}")
        if is_sold else InlineKeyboardButton("💰 Jual", callback_data=f"sell:{record_id}")
    )
    return [
        InlineKeyboardButton("🗑️ Hapus", callback_data=f"delrec:{record_id}"),
        sale_button,
        InlineKeyboardButton("📋 Salin", callback_data=f"copyrec:{record_id}"),
    ]


def decorate_edited_output(body: str, generated_code: str, settings: dict) -> str:
    pieces = []
    if generated_code:
        pieces.append(generated_code)
    if body:
        pieces.append(body)
    bottom = (settings.get("code_bottom") or "").strip()
    if bottom:
        pieces.append(bottom)
    return "\n".join(pieces).strip().upper()

async def notify_sale_team(context: ContextTypes.DEFAULT_TYPE, record: dict, seller_id: int):
    """Notify members of groups connected to the record owner or the seller."""
    owner_id = int(record.get("telegram_id") or 0)
    member_ids = set()
    for person_id in {owner_id, int(seller_id)}:
        memberships = supabase.table("bot_group_members").select("group_id").eq(
            "telegram_id", person_id
        ).execute().data or []
        for membership in memberships:
            members = supabase.table("bot_group_members").select("telegram_id").eq(
                "group_id", membership["group_id"]
            ).execute().data or []
            member_ids.update(int(m["telegram_id"]) for m in members)
    # Semua admin/pemilik yang terdaftar menerima notifikasi penjualan.
    try:
        admin_rows = supabase.table("bot_users").select("telegram_id,is_admin").limit(5000).execute().data or []
        member_ids.update(int(r["telegram_id"]) for r in admin_rows if r.get("is_admin") or is_owner(int(r["telegram_id"])))
    except Exception:
        log.exception("Gagal mengambil daftar admin untuk notifikasi penjualan")
    member_ids.update(ADMIN_IDS)
    if OWNER_ID:
        member_ids.add(OWNER_ID)
    member_ids.discard(int(seller_id))
    code = record.get("generated_code") or f"ID {record.get('id')}"
    notice = (
        "🔔 NOTIFIKASI DATA TERJUAL\n\n"
        f"Kode data: {code}\n"
        f"Pemilik data ID: {owner_id}\n"
        f"Dijual oleh Telegram ID: {seller_id}\n"
        "Status: TERJUAL"
    )
    for member_id in member_ids:
        try:
            await context.bot.send_message(chat_id=member_id, text=notice)
        except Exception as exc:
            # Telegram users who have never opened the bot may not be reachable.
            log.info("Tidak dapat mengirim notifikasi penjualan ke %s: %s", member_id, exc)


async def handle_record_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    if not query:
        return
    await query.answer()
    uid = update.effective_user.id
    data = query.data or ""

    if data == "again":
        context.user_data["awaiting"] = "manual"
        await query.message.reply_text(
            "🔄 Silakan kirim data acak berikutnya. Template dan pengaturan kode tetap digunakan.",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True),
        )
        return

    if data == "main_menu":
        context.user_data.pop("awaiting", None)
        context.user_data.pop("editing_record_id", None)
        await query.message.reply_text("Menu utama:", reply_markup=main_keyboard())
        return

    match = re.fullmatch(r"(edit|delete|save|sell|sellyes|sellno|delrec|deleteyes|deleteno|copyrec|copyreg|soldinfo):(\d+)", data)
    if not match:
        return
    action, raw_id = match.groups()
    record_id = int(raw_id)

    try:
        rows = supabase.table("format_records").select("*").eq("id", record_id).limit(1).execute().data
        if not rows:
            await query.message.reply_text("Data tidak ditemukan.")
            return
        row = rows[0]
        owner_id = int(row.get("telegram_id") or 0)
        if uid != owner_id and not is_staff(uid):
            await query.message.reply_text("⛔ Anda tidak memiliki izin untuk mengelola data ini.")
            return

        if action == "edit":
            if uid != owner_id:
                await query.message.reply_text("Hanya pemilik data yang boleh mengedit hasil ini.")
                return
            context.user_data["awaiting"] = "edit_record"
            context.user_data["editing_record_id"] = record_id
            await query.message.reply_text(
                "✏️ EDIT HASIL FORMAT\n\nKirim ulang data yang benar. Hasil ini akan diperbarui, bukan dibuat sebagai data baru.\n"
                "Kirim ❌ Batal jika tidak jadi.",
                reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True),
            )
            return

        if action == "delete":
            # Delete is immediate only for the unsaved preview owned by its creator.
            if uid != owner_id:
                await query.message.reply_text("Hanya pemilik data yang boleh menghapus pratinjau ini.")
                return
            supabase.table("format_records").delete().eq("id", record_id).eq("telegram_id", uid).execute()
            await query.message.reply_text("🗑️ Pratinjau hasil format dihapus.")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            return

        if action == "delrec":
            await query.message.reply_text(
                f"🗑️ Apakah Anda yakin ingin menghapus data {row.get('generated_code') or record_id}?",
                reply_markup=InlineKeyboardMarkup([[
                    InlineKeyboardButton("✅ Ya, hapus", callback_data=f"deleteyes:{record_id}"),
                    InlineKeyboardButton("❌ Batal", callback_data=f"deleteno:{record_id}"),
                ]])
            )
            return

        if action == "deleteyes":
            supabase.table("format_records").delete().eq("id", record_id).execute()
            await query.message.reply_text("✅ Data berhasil dihapus.")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            return

        if action == "deleteno":
            await query.message.reply_text("Penghapusan dibatalkan. Data tetap tersimpan.")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            return

        if action == "save":
            if row.get("is_saved"):
                await query.message.reply_text(
                    "✅ Data ini sudah tersimpan.",
                    reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔄 Buat Format Lagi", callback_data="again")]])
                )
                return
            supabase.table("format_records").update({"is_saved": True}).eq("id", record_id).execute()
            await query.message.reply_text(
                "✅ BERHASIL! DATA ANDA DISIMPAN.\n\nData masuk ke menu Hasil Format dan Hasil Format + Akun.",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔄 Buat Format Lagi", callback_data="again")]])
            )
            try:
                await query.edit_message_reply_markup(reply_markup=record_actions(record_id, already_saved=True))
            except Exception:
                pass
            return

        if action in ("copyrec", "copyreg"):
            output = (row.get("regular_output") if action == "copyreg" else row.get("full_output")) or row.get("regular_output") or "(hasil kosong)"
            await query.message.reply_text("📋 SALIN DATA — tekan dan tahan teks di bawah ini untuk menyalin:\n\n" + output[:3900])
            return

        if action == "soldinfo":
            await query.message.reply_text(f"ℹ️ Data {row.get('generated_code') or record_id} sudah ditandai TERJUAL oleh Telegram ID {row.get('sold_by') or 'tidak diketahui'}.")
            return

        if action == "sell":
            if row.get("is_sold"):
                await query.message.reply_text("Data ini sudah berstatus TERJUAL.")
                return
            await query.message.reply_text(
                "💰 KONFIRMASI PENJUALAN\n\nApakah Anda yakin ingin menjual data ini? Setelah dikonfirmasi, anggota tim terkait akan menerima notifikasi.",
                reply_markup=InlineKeyboardMarkup([[
                    InlineKeyboardButton("✅ Iya, yakin", callback_data=f"sellyes:{record_id}"),
                    InlineKeyboardButton("❌ Batal", callback_data=f"sellno:{record_id}"),
                ]])
            )
            return

        if action == "sellno":
            await query.message.reply_text("Penjualan dibatalkan.")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            return

        if action == "sellyes":
            if row.get("is_sold"):
                await query.message.reply_text("Data ini sudah berstatus TERJUAL.")
                return
            now = datetime.now(timezone.utc).isoformat()
            supabase.table("format_records").update({
                "is_sold": True,
                "sold_by": uid,
                "sold_at": now,
            }).eq("id", record_id).execute()
            row.update({"is_sold": True, "sold_by": uid, "sold_at": now})
            await query.message.reply_text(
                f"✅ DATA BERHASIL DITANDAI TERJUAL.\nKode: {row.get('generated_code') or record_id}\nDijual oleh Telegram ID: {uid}"
            )
            await notify_sale_team(context, row, uid)
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            return
    except Exception:
        log.exception("Gagal menjalankan aksi hasil format")
        await query.message.reply_text("Aksi gagal dijalankan. Pastikan SQL terbaru sudah dijalankan di Supabase dan periksa Railway Logs.")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        get_user(update.effective_user)
        get_settings(update.effective_user.id)
    except Exception:
        log.exception("Gagal menyiapkan pengguna di Supabase")
        await update.message.reply_text(
            "Bot belum dapat menyimpan data. Periksa konfigurasi Supabase dan SQL tabel.",
            reply_markup=main_keyboard(),
        )
        return
    await update.message.reply_text(
        "🤖 BOT AUTO FORMAT\n\nPilih menu di bawah ini.",
        reply_markup=main_keyboard(),
    )

async def show_account(update: Update, context: ContextTypes.DEFAULT_TYPE):
    u = update.effective_user
    get_user(u)
    username = f"@{u.username}" if u.username else "(belum ada)"
    role = role_name(u.id)
    await update.message.reply_text(
        f"👤 AKUN SAYA\n\nNama: {(u.full_name or '').strip()}\n"
        f"Username: {username}\nTelegram ID: {u.id}\nPeran: {role}",
        reply_markup=main_keyboard(),
    )

async def ask_manual(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["awaiting"] = "manual"
    await update.message.reply_text(
        "✨ KIRIM DATA UNTUK DIFORMAT\n\n" + HELP_TEXT,
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True),
    )

async def handle_format_input(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str):
    uid = update.effective_user.id
    settings = get_settings(uid)
    template = (settings.get("template_text") or "").strip()
    if not template:
        context.user_data["awaiting"] = "template"
        await update.message.reply_text(
            "🛠️ Template belum diatur. Kirim template buatan Anda terlebih dahulu, "
            "misalnya:\nNAMA : {1}\nSALDO : {2}\nPT : {3}\n\n"
            "Urutan dan nama kolom sepenuhnya Anda tentukan.",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True),
        )
        return
    values = parse_input(text, template)
    if not values:
        await update.message.reply_text("Data kosong. Silakan kirim data lagi.")
        return

    editing_id = context.user_data.get("editing_record_id") if context.user_data.get("awaiting") == "edit_record" else None
    try:
        if editing_id:
            existing = supabase.table("format_records").select("*").eq("id", int(editing_id)).eq(
                "telegram_id", uid
            ).limit(1).execute().data
            if not existing:
                context.user_data.pop("editing_record_id", None)
                context.user_data.pop("awaiting", None)
                await update.message.reply_text("Data yang diedit tidak ditemukan. Silakan buat format baru.", reply_markup=main_keyboard())
                return
            old = existing[0]
            no_decoration = {**settings, "code_prefix": "", "code_bottom": ""}
            regular_raw, full_raw, _ = make_outputs(template, values, no_decoration, 1)
            regular = decorate_edited_output(regular_raw, old.get("generated_code") or "", settings)
            full = decorate_edited_output(full_raw, old.get("generated_code") or "", settings)
            supabase.table("format_records").update({
                "input_text": text,
                "regular_output": regular,
                "full_output": full,
            }).eq("id", int(editing_id)).eq("telegram_id", uid).execute()
            record_id = int(editing_id)
            is_saved = bool(old.get("is_saved"))
            context.user_data.pop("editing_record_id", None)
            context.user_data.pop("awaiting", None)
        else:
            number = next_number(settings)
            regular, full, generated_code = make_outputs(template, values, settings, number)
            inserted = store_record(uid, text, regular, full, generated_code)
            if not inserted:
                raise RuntimeError("Supabase tidak mengembalikan record yang baru disimpan.")
            record_id = int(inserted[0]["id"])
            is_saved = False
            if settings.get("code_prefix"):
                save_settings(uid, {"next_code_number": number + 1})
            context.user_data.pop("awaiting", None)
    except Exception:
        log.exception("Gagal menyimpan hasil format")
        await update.message.reply_text(
            "Gagal menyimpan hasil format ke Supabase. Periksa tabel dan Railway logs.",
            reply_markup=main_keyboard(),
        )
        return

    await update.message.reply_text(
        "✅ FORMAT BERHASIL" + (" — HASIL DIPERBARUI" if editing_id else "") + "\n\n" + regular +
        "\n\nPilih tindakan untuk hasil ini:",
        reply_markup=record_actions(record_id, already_saved=is_saved),
    )

async def ask_template(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["awaiting"] = "template"
    settings = get_settings(update.effective_user.id)
    await update.message.reply_text(
        "🛠️ SETTING TEMPLATE\n\nKirim template baru. Gunakan {1}, {2}, {3}, dan seterusnya.\n"
        "Untuk menentukan batas Hasil Format biasa secara manual, letakkan {BATAS} pada baris batas.\n"
        "Jika tidak memakai {BATAS}, bot akan berhenti sampai baris PT jika ada; jika tidak ada PT, seluruh template ditampilkan.\n\n"
        "Template saat ini:\n" + (settings.get("template_text") or "(belum diatur; tidak ada template default)"),
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True),
    )

async def ask_code_top(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["awaiting"] = "code_prefix"
    await update.message.reply_text(
        "🔢 SETTING KODE ATAS\n\nKirim contoh kode, misalnya:\nABC 001\n\n"
        "Bot akan membuat ABC_001, ABC_002, ABC_003, dan seterusnya.",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True),
    )

async def ask_code_bottom(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["awaiting"] = "code_bottom"
    await update.message.reply_text(
        "🔻 SETTING KODE BAWAH\n\nKirim teks yang ingin ditambahkan di akhir setiap hasil. "
        "Kirim tanda - jika ingin mengosongkan.",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True),
    )

async def show_active_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    s = get_settings(update.effective_user.id)
    prefix = s.get("code_prefix") or "(belum diatur)"
    next_code = code_for(s, int(s.get("next_code_number") or 1)) or "(belum diatur)"
    bottom = s.get("code_bottom") or "(kosong)"
    await update.message.reply_text(
        f"🔢 KODE AKTIF SAYA\n\nAwalan: {prefix}\nKode berikutnya: {next_code}\n"
        f"Kode bawah: {bottom}",
        reply_markup=format_keyboard(),
    )

def fetch_saved_records(uid: int, sold_only: bool = False):
    """Fetch every matching saved record in pages instead of silently stopping at 1,000."""
    all_rows = []
    offset = 0
    page_size = 1000
    while True:
        query = supabase.table("format_records").select("*").eq("is_saved", True)
        if not is_staff(uid):
            query = query.eq("telegram_id", uid)
        if sold_only:
            query = query.eq("is_sold", True)
        page = query.order("id", desc=False).range(offset, offset + page_size - 1).execute().data or []
        all_rows.extend(page)
        if len(page) < page_size:
            break
        offset += page_size
    return all_rows

async def send_format_records(update: Update, context: ContextTypes.DEFAULT_TYPE, rows: list[dict], full: bool, history: bool = False):
    if not rows:
        await update.message.reply_text("Tidak ada data yang cocok.", reply_markup=format_results_keyboard())
        return
    title = "📈 RIWAYAT PENJUALAN" if history else ("🧾 HASIL FORMAT + AKUN" if full else "📄 HASIL FORMAT")
    await update.message.reply_text(f"{title}\nJumlah data: {len(rows)}. Data ditampilkan dalam kelompok pesan agar tombol setiap data tetap berfungsi.")
    owner_rows = supabase.table("bot_users").select("telegram_id,username,first_name,last_name").limit(5000).execute().data or []
    owner_map = {int(item["telegram_id"]): item for item in owner_rows}
    batches: list[tuple[str, list[list[InlineKeyboardButton]]]] = []
    current_text = ""
    current_keyboard: list[list[InlineKeyboardButton]] = []
    for index, row in enumerate(rows, start=1):
        owner_id = int(row.get("telegram_id") or 0)
        owner = owner_map.get(owner_id, {})
        username = f"@{owner.get('username')}" if owner.get("username") else "(belum ada)"
        name = " ".join(x for x in [owner.get("first_name", ""), owner.get("last_name", "")] if x).strip() or "(tidak diketahui)"
        output = row.get("full_output") if full else row.get("regular_output")
        status = "TERJUAL" if row.get("is_sold") else "BELUM TERJUAL"
        header = f"\nAKUN PEMILIK DATA\nNama: {name}\nUsername: {username}\nTelegram ID: {owner_id}" if full else ""
        if history:
            sold_by = row.get("sold_by") or "tidak diketahui"
            sold_at = row.get("sold_at") or "waktu tidak tercatat"
            header += f"\nDijual oleh Telegram ID: {sold_by}\nWaktu: {sold_at}"
        record_text = f"#{index} | ID DATA: {row['id']} | KODE: {row.get('generated_code') or '-'} | STATUS: {status}{header}\n{output or '(hasil kosong)'}"
        separator = "\n\n━━━━━━━━━━━━━━━━━━\n\n"
        if current_text and len(current_text) + len(separator) + len(record_text) > 3400:
            batches.append((current_text, current_keyboard))
            current_text, current_keyboard = "", []
        if current_text:
            current_text += separator
        current_text += record_text
        # The keyboard row belongs to this exact record. Regular results only offer copy.
        current_keyboard.append(saved_result_actions(int(row["id"]), bool(row.get("is_sold")), full=full, history=history))
    if current_text:
        batches.append((current_text, current_keyboard))
    for body, keyboard_rows in batches:
        markup = InlineKeyboardMarkup(keyboard_rows) if keyboard_rows else None
        await update.message.reply_text(body[:3900], reply_markup=markup)

async def show_latest(update: Update, context: ContextTypes.DEFAULT_TYPE, full: bool):
    uid = update.effective_user.id
    context.user_data["results_full_mode"] = bool(full)
    await update.message.reply_text(
        "Pilih tindakan untuk " + ("Hasil Format + Akun" if full else "Hasil Format") + ":",
        reply_markup=format_results_keyboard(),
    )

async def show_all_results(update: Update, context: ContextTypes.DEFAULT_TYPE, full: bool, history: bool = False):
    uid = update.effective_user.id
    rows = fetch_saved_records(uid, sold_only=history)
    await send_format_records(update, context, rows, full=full, history=history)

async def search_format_results(update: Update, context: ContextTypes.DEFAULT_TYPE, keyword: str, full: bool):
    uid = update.effective_user.id
    keyword = keyword.strip().upper()
    rows = fetch_saved_records(uid)
    matches = []
    for row in rows:
        haystack = "\n".join(str(row.get(k) or "") for k in ("generated_code", "regular_output", "full_output", "input_text")).upper()
        # Normalisasi spasi/underscore/tanda baca agar ABC 00056 juga cocok dengan ABC_00056.
        norm_keyword = re.sub(r"[\W_]+", " ", keyword).strip()
        norm_haystack = re.sub(r"[\W_]+", " ", haystack).strip()
        if keyword in haystack or (norm_keyword and norm_keyword in norm_haystack):
            matches.append(row)
    await update.message.reply_text(f"🔎 Pencarian: {keyword}\nDitemukan: {len(matches)} data.")
    await send_format_records(update, context, matches, full=full)


async def show_my_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_user_data_menu(update, context)

async def show_my_data_results(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_user_data_results_menu(update, context)


async def show_user_data_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📂 DATA SAYA\n\nPilih tindakan:", reply_markup=data_keyboard())

async def show_user_data_results_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📊 HASIL DATA SAYA\n\nPilih tampilan:", reply_markup=my_data_results_keyboard())

async def save_raw_user_data(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str):
    uid = update.effective_user.id
    supabase.table("user_data").insert({
        "telegram_id": uid,
        "data_text": text,
        "created_at": datetime.now(timezone.utc).isoformat(),
    }).execute()
    context.user_data.pop("awaiting", None)
    await update.message.reply_text("✅ Data mentah berhasil disimpan. Data ini hanya terlihat oleh akun Anda.", reply_markup=data_keyboard())

async def search_user_data(update: Update, context: ContextTypes.DEFAULT_TYPE, keyword: str):
    uid = update.effective_user.id
    rows = supabase.table("user_data").select("id,data_text,created_at").eq("telegram_id", uid).ilike("data_text", f"%{keyword}%").order("id", desc=True).limit(10).execute().data
    context.user_data.pop("awaiting", None)
    if not rows:
        await update.message.reply_text("Tidak ada data milik Anda yang cocok dengan kata kunci itu.", reply_markup=data_keyboard())
        return
    for row in rows:
        msg = f"📄 ID {row['id']} | {row.get('created_at','')}\n\n{row.get('data_text','')}"
        await update.message.reply_text(msg[:3900], reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🗑️ Hapus", callback_data=f"delud:{row['id']}")]]))

async def list_raw_user_data(update: Update, context: ContextTypes.DEFAULT_TYPE, with_account: bool = False):
    uid = update.effective_user.id
    rows = supabase.table("user_data").select("id,data_text,created_at").eq("telegram_id", uid).order("id", desc=True).limit(10).execute().data
    if not rows:
        await update.message.reply_text("Belum ada data di Data Saya.", reply_markup=my_data_results_keyboard())
        return
    u = update.effective_user
    account = f"👤 AKUN\nNama: {u.full_name}\nUsername: @{u.username if u.username else '(belum ada)'}\nTelegram ID: {uid}\n\n" if with_account else ""
    for row in rows:
        await update.message.reply_text((account + f"📄 ID {row['id']}\n\n" + row.get("data_text", ""))[:3900], reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🗑️ Hapus", callback_data=f"delud:{row['id']}")]]))

async def delete_raw_data_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    rows = supabase.table("user_data").select("id,data_text").eq("telegram_id", uid).order("id", desc=True).limit(10).execute().data
    if not rows:
        await update.message.reply_text("Tidak ada Data Saya yang perlu dihapus.", reply_markup=my_data_results_keyboard())
        return
    for row in rows:
        preview = (row.get("data_text") or "").replace("\n", " ")[:100]
        await update.message.reply_text(f"ID {row['id']}: {preview}", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🗑️ Hapus data ini", callback_data=f"delud:{row['id']}")]]))

async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not is_staff(uid):
        await update.message.reply_text("⛔ Menu ini khusus admin.", reply_markup=main_keyboard())
        return
    try:
        users = supabase.table("bot_users").select("telegram_id", count="exact").execute()
        records = supabase.table("format_records").select("id", count="exact").execute()
        await update.message.reply_text(
            "🛡️ PANEL ADMIN\n\n"
            f"Jumlah pengguna: {users.count if users.count is not None else len(users.data or [])}\n"
            f"Jumlah hasil format: {records.count if records.count is not None else len(records.data or [])}\n\nPilih pengelolaan:",
            reply_markup=admin_keyboard())
    except Exception:
        log.exception("Gagal membuka panel admin")
        await update.message.reply_text("Gagal membaca statistik. Pastikan tabel Supabase tersedia.", reply_markup=main_keyboard())

async def admin_add_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update.effective_user.id):
        await update.message.reply_text("⛔ Hanya pemilik bot yang boleh menambah admin.", reply_markup=admin_keyboard())
        return
    context.user_data["awaiting"] = "add_admin"
    await update.message.reply_text("Kirim Telegram ID orang yang akan dijadikan admin. Hanya angka ID.", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True))

async def admin_add_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_staff(update.effective_user.id):
        return
    context.user_data["awaiting"] = "new_group_name"
    await update.message.reply_text("Kirim nama grup/tim yang ingin dibuat, misalnya TIM BANDUNG.", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True))

async def admin_list_admins(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_staff(update.effective_user.id):
        return
    rows = supabase.table("bot_users").select("telegram_id,username,first_name,last_name,is_admin").eq("is_admin", True).order("telegram_id").limit(200).execute().data
    ids = {int(r["telegram_id"]) for r in rows}
    ids.update(ADMIN_IDS)
    if OWNER_ID:
        ids.add(OWNER_ID)
    lines = ["👑 DAFTAR ADMIN"]
    for aid in sorted(ids):
        row = next((r for r in rows if int(r["telegram_id"]) == int(aid)), {})
        name = " ".join(x for x in [row.get("first_name", ""), row.get("last_name", "")] if x).strip() or row.get("username") or "Belum pernah /start"
        lines.append(f"{aid} — {name}" + (" (Pemilik)" if is_owner(int(aid)) else ""))
    await update.message.reply_text("\n".join(lines)[:3900], reply_markup=admin_keyboard())

async def admin_list_groups(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_staff(update.effective_user.id):
        return
    rows = supabase.table("bot_groups").select("id,group_name,created_by,created_at").order("id", desc=True).limit(100).execute().data
    if not rows:
        await update.message.reply_text("Belum ada grup. Pilih 2. Tambah Grup untuk membuat tim.", reply_markup=admin_keyboard())
        return
    for group in rows:
        members = supabase.table("bot_group_members").select("telegram_id", count="exact").eq("group_id", group["id"]).execute()
        count = members.count if members.count is not None else len(members.data or [])
        await update.message.reply_text(f"👥 {group['group_name']}\nID grup: {group['id']}\nAnggota: {count}", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("➕ Tambah Anggota", callback_data=f"addmem:{group['id']}"), InlineKeyboardButton("🗑️ Hapus Grup", callback_data=f"delgrp:{group['id']}")]]))

async def admin_add_member(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_staff(update.effective_user.id):
        return
    rows = supabase.table("bot_groups").select("id,group_name").order("id", desc=True).limit(100).execute().data
    if not rows:
        await update.message.reply_text("Buat grup terlebih dahulu melalui 2. Tambah Grup.", reply_markup=admin_keyboard())
        return
    buttons = [[InlineKeyboardButton(r["group_name"], callback_data=f"addmem:{r['id']}")] for r in rows]
    await update.message.reply_text("Pilih grup untuk menambahkan anggota:", reply_markup=InlineKeyboardMarkup(buttons))

async def admin_remove_member(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_staff(update.effective_user.id):
        return
    context.user_data["awaiting"] = "remove_member"
    await update.message.reply_text("Kirim ID Telegram anggota yang ingin dikeluarkan dari semua grup.", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True))

async def admin_delete_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_staff(update.effective_user.id):
        return
    rows = supabase.table("bot_groups").select("id,group_name").order("id", desc=True).limit(100).execute().data
    if not rows:
        await update.message.reply_text("Belum ada grup.", reply_markup=admin_keyboard())
        return
    buttons = [[InlineKeyboardButton("🗑️ " + r["group_name"], callback_data=f"delgrp:{r['id']}")] for r in rows]
    await update.message.reply_text("Pilih grup yang akan dihapus:", reply_markup=InlineKeyboardMarkup(buttons))

async def handle_admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = update.effective_user.id
    data = q.data or ""
    if data.startswith("delud:"):
        rid = int(data.split(":", 1)[1])
        supabase.table("user_data").delete().eq("id", rid).eq("telegram_id", uid).execute()
        await q.message.reply_text("🗑️ Data berhasil dihapus jika memang milik akun Anda.")
        return
    if data.startswith("addmem:"):
        if not is_staff(uid):
            await q.message.reply_text("Akses ditolak.")
            return
        gid = int(data.split(":", 1)[1])
        groups = supabase.table("bot_groups").select("id,group_name").eq("id", gid).limit(1).execute().data
        if not groups:
            await q.message.reply_text("Grup tidak ditemukan.")
            return
        context.user_data["awaiting"] = "add_members_to_group"
        context.user_data["selected_group_id"] = gid
        await q.message.reply_text(f"Kirim ID Telegram anggota untuk grup {groups[0]['group_name']}. Boleh satu ID per baris.", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True))
        return
    if data.startswith("delgrp:"):
        if not is_staff(uid):
            await q.message.reply_text("Akses ditolak.")
            return
        gid = int(data.split(":", 1)[1])
        supabase.table("bot_group_members").delete().eq("group_id", gid).execute()
        supabase.table("bot_groups").delete().eq("id", gid).execute()
        await q.message.reply_text("🗑️ Grup dan keanggotaan grup berhasil dihapus.", reply_markup=admin_keyboard())
        return


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if context.user_data.get("awaiting") != "excel":
        await update.message.reply_text("Untuk memproses Excel, pilih Buat Format → 2. Buat File Excel terlebih dahulu.", reply_markup=main_keyboard())
        return
    doc = update.message.document
    settings = get_settings(update.effective_user.id)
    if not (settings.get("template_text") or "").strip():
        context.user_data.pop("awaiting", None)
        await update.message.reply_text(
            "Template belum diatur. Buka Buat Format → Setting Template terlebih dahulu, "
            "lalu kirim file Excel lagi.",
            reply_markup=format_keyboard(),
        )
        return
    if not doc.file_name.lower().endswith(".xlsx"):
        await update.message.reply_text("Kirim file .xlsx. File .xls lama belum didukung.")
        return
    try:
        tg_file = await context.bot.get_file(doc.file_id)
        bio = BytesIO()
        await tg_file.download_to_memory(out=bio)
        bio.seek(0)
        wb = load_workbook(bio, read_only=True, data_only=True)
        ws = wb.active
        made = 0
        for row in ws.iter_rows(values_only=True):
            vals = [str(v).strip() for v in row if v is not None and str(v).strip()]
            if not vals:
                continue
            text = ", ".join(vals)
            settings = get_settings(update.effective_user.id)
            number = next_number(settings)
            regular, full, code = make_outputs(settings.get("template_text") or "", vals, settings, number)
            store_record(update.effective_user.id, text, regular, full, code)
            if settings.get("code_prefix"):
                save_settings(update.effective_user.id, {"next_code_number": number + 1})
            made += 1
        context.user_data.pop("awaiting", None)
        await update.message.reply_text(f"✅ Excel selesai diproses. {made} baris disimpan.", reply_markup=main_keyboard())
    except Exception:
        log.exception("Gagal memproses Excel")
        await update.message.reply_text("Gagal memproses file Excel. Pastikan file .xlsx valid dan lihat Railway logs.", reply_markup=main_keyboard())

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.effective_user:
        return
    text = (update.message.text or "").strip()
    uid = update.effective_user.id
    if text == "❌ Batal":
        context.user_data.pop("awaiting", None)
        context.user_data.pop("editing_record_id", None)
        context.user_data.pop("selected_group_id", None)
        await update.message.reply_text("Dibatalkan.", reply_markup=main_keyboard())
        return

    awaiting = context.user_data.get("awaiting")
    menu_texts = {
        "👤 Akun", "✨ Buat Format", "📄 Hasil Format", "🧾 Hasil Format + Akun",
        "📂 Data Saya", "📊 Hasil Data Saya", "🛡️ Panel Admin",
        "1. Buat Manual", "2. Buat File Excel", "3. Setting Template",
        "4. Setting Kode Atas", "5. Setting Kode Bawah", "6. Kode Aktif Saya",
        "8. Kembali ke Menu Utama", "1. Kirim Data", "2. Cari Data",
        "3. Hasil Data Saya", "4. Kembali", "1. Data Format", "2. Format + Akun",
        "3. Hapus Data Saya", "1. Tambah Admin", "2. Tambah Grup", "3. Daftar Admin",
        "4. Daftar Grup", "5. Tambah Anggota", "6. Hapus Anggota", "7. Hapus Grup",
        "4. Kembali ke Menu Utama", "1. Cari Data", "2. Tampilkan Semua", "3. Riwayat Penjualan"
    }
    if awaiting and not (text in menu_texts and awaiting in ("manual", "search_user_data", "search_format_regular", "search_format_full")):
        try:
            if awaiting in ("manual", "edit_record"):
                await handle_format_input(update, context, text)
                return
            if awaiting == "template":
                if not re.search(r"\{\d+\}", text):
                    await update.message.reply_text("Template harus berisi placeholder seperti {1}, {2}, dan seterusnya. Coba lagi.")
                    return
                save_settings(uid, {"template_text": text})
                context.user_data.pop("awaiting", None)
                await update.message.reply_text("✅ Template berhasil disimpan.", reply_markup=format_keyboard())
                return
            if awaiting == "code_prefix":
                match = re.match(r"^\s*(.*?)\s+(\d+)\s*$", text)
                if not match:
                    match = re.match(r"^\s*(.*?)_+(\d+)\s*$", text)
                if not match or not match.group(1).strip():
                    await update.message.reply_text("Format tidak dikenali. Contoh: ABC 001")
                    return
                prefix = re.sub(r"[\s_-]+", "_", match.group(1).strip()).strip("_").upper()
                digits = match.group(2)
                save_settings(uid, {"code_prefix": prefix, "next_code_number": int(digits), "code_width": len(digits)})
                context.user_data.pop("awaiting", None)
                await update.message.reply_text(f"✅ Kode disimpan. Kode pertama: {prefix}_{int(digits):0{len(digits)}d}", reply_markup=format_keyboard())
                return
            if awaiting == "code_bottom":
                save_settings(uid, {"code_bottom": "" if text == "-" else text.upper()})
                context.user_data.pop("awaiting", None)
                await update.message.reply_text("✅ Kode bawah berhasil disimpan.", reply_markup=format_keyboard())
                return
            if awaiting == "send_user_data":
                await save_raw_user_data(update, context, text)
                return
            if awaiting == "search_user_data":
                await search_user_data(update, context, text)
                return
            if awaiting in ("search_format_regular", "search_format_full"):
                context.user_data.pop("awaiting", None)
                await search_format_results(update, context, text, full=(awaiting == "search_format_full"))
                return
            if awaiting == "add_admin":
                if not is_owner(uid):
                    context.user_data.pop("awaiting", None)
                    await update.message.reply_text("⛔ Hanya pemilik bot yang dapat menambah admin.", reply_markup=main_keyboard())
                    return
                if not text.isdigit():
                    await update.message.reply_text("ID harus berupa angka Telegram. Coba lagi.")
                    return
                target = int(text)
                existing = supabase.table("bot_users").select("telegram_id").eq("telegram_id", target).limit(1).execute().data
                if not existing:
                    supabase.table("bot_users").insert({"telegram_id": target, "is_admin": True, "first_name": "", "last_name": "", "username": ""}).execute()
                else:
                    supabase.table("bot_users").update({"is_admin": True}).eq("telegram_id", target).execute()
                context.user_data.pop("awaiting", None)
                await update.message.reply_text(f"✅ ID {target} sekarang menjadi admin. Admin dapat melihat Hasil Format dan Hasil Format + Akun yang tersimpan, tetapi tidak otomatis dapat melihat Data Saya milik pengguna lain.", reply_markup=admin_keyboard())
                return
            if awaiting == "new_group_name":
                if not is_staff(uid):
                    context.user_data.pop("awaiting", None)
                    await update.message.reply_text("Akses ditolak.", reply_markup=main_keyboard())
                    return
                name = text.strip()[:100]
                inserted = supabase.table("bot_groups").insert({"group_name": name, "created_by": uid}).select("id,group_name").execute().data
                if not inserted:
                    raise RuntimeError("Database tidak mengembalikan grup baru")
                context.user_data["selected_group_id"] = int(inserted[0]["id"])
                context.user_data["awaiting"] = "create_group_ids"
                await update.message.reply_text(f"✅ Grup {name} dibuat. Sekarang kirim ID Telegram anggota, satu ID per baris. Kirim - jika ingin menambahkan anggota nanti.", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True))
                return
            if awaiting in ("create_group_ids", "add_members_to_group"):
                gid = int(context.user_data.get("selected_group_id", 0))
                ids = [int(x.strip()) for x in text.splitlines() if x.strip().isdigit()]
                if text.strip() == "-":
                    ids = []
                if not ids and text.strip() != "-":
                    await update.message.reply_text("Tidak ada ID angka yang terbaca. Kirim satu ID per baris, atau - untuk selesai.")
                    return
                added = 0
                for target in ids:
                    # Ensure FK row exists. The member can later /start to update their profile.
                    exists = supabase.table("bot_users").select("telegram_id").eq("telegram_id", target).limit(1).execute().data
                    if not exists:
                        supabase.table("bot_users").insert({"telegram_id": target, "is_admin": False, "first_name": "", "last_name": "", "username": ""}).execute()
                    old = supabase.table("bot_group_members").select("id").eq("group_id", gid).eq("telegram_id", target).limit(1).execute().data
                    if not old:
                        supabase.table("bot_group_members").insert({"group_id": gid, "telegram_id": target, "added_by": uid}).execute()
                        added += 1
                context.user_data.pop("awaiting", None)
                context.user_data.pop("selected_group_id", None)
                await update.message.reply_text(f"✅ Selesai. Anggota baru ditambahkan: {added}.", reply_markup=admin_keyboard())
                return
            if awaiting == "remove_member":
                if not text.isdigit():
                    await update.message.reply_text("Kirim ID Telegram berupa angka.")
                    return
                target = int(text)
                supabase.table("bot_group_members").delete().eq("telegram_id", target).execute()
                context.user_data.pop("awaiting", None)
                await update.message.reply_text(f"✅ ID {target} dikeluarkan dari grup yang diikuti.", reply_markup=admin_keyboard())
                return
        except Exception:
            log.exception("Gagal memproses input untuk state %s", awaiting)
            await update.message.reply_text("Operasi gagal. Pastikan SQL tabel terbaru sudah dijalankan di Supabase, lalu cek Railway Logs.", reply_markup=main_keyboard())
            return

    if text == "/start":
        await start(update, context)
    elif text == "👤 Akun":
        await show_account(update, context)
    elif text == "✨ Buat Format":
        context.user_data.pop("awaiting", None)
        await update.message.reply_text("✨ MENU BUAT FORMAT\nPilih fitur:", reply_markup=format_keyboard())
    elif text == "1. Buat Manual":
        await ask_manual(update, context)
    elif text == "2. Buat File Excel":
        context.user_data["awaiting"] = "excel"
        await update.message.reply_text("📊 Kirim file Excel .xlsx. Setiap baris akan menjadi satu hasil format.", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True))
    elif text == "3. Setting Template":
        await ask_template(update, context)
    elif text == "4. Setting Kode Atas":
        await ask_code_top(update, context)
    elif text == "5. Setting Kode Bawah":
        await ask_code_bottom(update, context)
    elif text == "6. Kode Aktif Saya":
        await show_active_code(update, context)
    elif text == "8. Kembali ke Menu Utama" or text == "4. Kembali ke Menu Utama":
        context.user_data.pop("awaiting", None)
        await update.message.reply_text("Menu utama:", reply_markup=main_keyboard())
    elif text == "📄 Hasil Format":
        await show_latest(update, context, False)
    elif text == "🧾 Hasil Format + Akun":
        await show_latest(update, context, True)
    elif text == "1. Cari Data":
        full = bool(context.user_data.get("results_full_mode", False))
        context.user_data["awaiting"] = "search_format_full" if full else "search_format_regular"
        await update.message.reply_text("Kirim kode atau kata kunci yang ingin dicari, misalnya ABC_00056, Bandung, atau Antapani.", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True))
    elif text == "2. Tampilkan Semua":
        await show_all_results(update, context, full=bool(context.user_data.get("results_full_mode", False)))
    elif text == "3. Riwayat Penjualan":
        await show_all_results(update, context, full=bool(context.user_data.get("results_full_mode", False)), history=True)
    elif text == "📂 Data Saya":
        context.user_data.pop("awaiting", None)
        await show_user_data_menu(update, context)
    elif text == "📊 Hasil Data Saya" or text == "3. Hasil Data Saya":
        context.user_data.pop("awaiting", None)
        await show_user_data_results_menu(update, context)
    elif text == "1. Kirim Data":
        context.user_data["awaiting"] = "send_user_data"
        await update.message.reply_text("Silakan kirim data yang ingin disimpan. Data disimpan persis seperti yang kamu kirim.", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True))
    elif text == "2. Cari Data":
        context.user_data["awaiting"] = "search_user_data"
        await update.message.reply_text("Kirim kata kunci untuk mencari data milikmu.", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Batal")]], resize_keyboard=True))
    elif text == "1. Data Format":
        await list_raw_user_data(update, context, False)
    elif text == "2. Format + Akun":
        await list_raw_user_data(update, context, True)
    elif text == "3. Hapus Data Saya":
        await delete_raw_data_menu(update, context)
    elif text == "4. Kembali":
        await show_user_data_menu(update, context)
    elif text == "🛡️ Panel Admin":
        await admin_panel(update, context)
    elif text == "1. Tambah Admin":
        await admin_add_admin(update, context)
    elif text == "2. Tambah Grup":
        await admin_add_group(update, context)
    elif text == "3. Daftar Admin":
        await admin_list_admins(update, context)
    elif text == "4. Daftar Grup":
        await admin_list_groups(update, context)
    elif text == "5. Tambah Anggota":
        await admin_add_member(update, context)
    elif text == "6. Hapus Anggota":
        await admin_remove_member(update, context)
    elif text == "7. Hapus Grup":
        await admin_delete_group(update, context)
    else:
        await update.message.reply_text("Pilih menu terlebih dahulu.", reply_markup=main_keyboard())


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    log.error("Unhandled exception: %s", context.error, exc_info=context.error)

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(handle_admin_callback, pattern=r"^(delud|addmem|delgrp):"))
    app.add_handler(CallbackQueryHandler(handle_record_callback))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_error_handler(error_handler)
    log.info("Bot AUTO FORMAT starting")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
