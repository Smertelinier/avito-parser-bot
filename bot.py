import asyncio
import json
import os
import random
import re
from datetime import datetime, timedelta
from io import BytesIO
from urllib.parse import quote

import httpx
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import BufferedInputFile, LabeledPrice, PreCheckoutQuery
from bs4 import BeautifulSoup
from openpyxl import Workbook

TOKEN = os.getenv("BOT_TOKEN", "8874282024:AAFfeftxs2MDuRjEz2N6PXiR21LpYwN6M5c")
FREE_DAYS = 3
PRICE = "300⭐/мес"

USER_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.json")
bot = Bot(token=TOKEN)
dp = Dispatcher()

_LAST_REQUEST = 0.0
_users_cache = None
_users_dirty = False
_users_lock = asyncio.Lock()

_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/121.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:122.0) Gecko/20100101 Firefox/122.0",
]


def _load_users():
    if not os.path.exists(USER_FILE):
        return {}
    with open(USER_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_users(users):
    import tempfile
    tmp = USER_FILE + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
        os.replace(tmp, USER_FILE)
    except PermissionError:
        pass


async def _flush_users():
    global _users_cache, _users_dirty
    async with _users_lock:
        if _users_dirty and _users_cache is not None:
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(None, _save_users, dict(_users_cache))
            _users_dirty = False


async def _periodic_flush():
    while True:
        await asyncio.sleep(30)
        await _flush_users()


def _get_users():
    global _users_cache
    if _users_cache is None:
        _users_cache = _load_users()
    return _users_cache


def _mark_dirty():
    global _users_dirty
    _users_dirty = True


def get_or_create_user(user_id):
    users = _get_users()
    uid = str(user_id)
    if uid not in users:
        users[uid] = {
            "trial_start": datetime.now().isoformat(),
            "subscribed_until": None,
            "searches": 0,
            "last_seen": None,
        }
        _mark_dirty()
    return users[uid]


def is_subscribed(user_id):
    user = get_or_create_user(user_id)
    if user.get("blocked"):
        return False
    if user.get("subscribed_until"):
        until = datetime.fromisoformat(user["subscribed_until"])
        if datetime.now() < until:
            return True
    trial_start = datetime.fromisoformat(user["trial_start"])
    if datetime.now() - trial_start < timedelta(days=FREE_DAYS):
        return True
    return False


async def search_avito(query):
    global _LAST_REQUEST
    now = datetime.now().timestamp()
    since_last = now - _LAST_REQUEST
    if since_last < 8.0:
        await asyncio.sleep(8.0 - since_last)
    _LAST_REQUEST = datetime.now().timestamp()

    headers = {
        "User-Agent": random.choice(_AGENTS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.5",
    }
    url = f"https://www.avito.ru/all?q={quote(query)}"
    try:
        async with httpx.AsyncClient(timeout=25, follow_redirects=True) as client:
            resp = await client.get(url, headers=headers)
            if resp.status_code != 200:
                return None
            html = resp.text
    except Exception:
        return None

    soup = BeautifulSoup(html, "html.parser")
    items = soup.select("[data-marker='item']")
    if not items:
        items = soup.select(".iva-item-root")
    if not items:
        items = soup.select("[itemprop='itemListElement']")

    results = []
    for item in items[:30]:
        try:
            title_el = item.select_one("[itemprop='name']") or item.select_one(".title-root") or item.select_one("a[title]")
            title = title_el.get("title") or title_el.text.strip() if title_el else "Нет названия"

            price_el = item.select_one("[itemprop='price']") or item.select_one(".price-text") or item.select_one("[content]")
            if price_el:
                price = price_el.get("content") or price_el.text.strip()
            else:
                price = "0"

            link_el = item.select_one("a[href]") or title_el
            link = "https://www.avito.ru" + link_el.get("href") if link_el and link_el.get("href", "").startswith("/") else ""

            location_el = item.select_one(".geo-address") or item.select_one("[data-marker='item-address']")
            location = location_el.text.strip() if location_el else ""

            price_num = 0
            price_clean = re.sub(r"[^\d]", "", str(price))
            if price_clean:
                price_num = int(price_clean)

            results.append({
                "title": title[:80],
                "price": price_num,
                "location": location[:50],
                "link": link,
            })
        except Exception:
            continue

    return results if results else None


def make_excel(items):
    wb = Workbook()
    ws = wb.active
    ws.title = "Avito Radar"
    ws.append(["Название", "Цена (₽)", "Локация", "Ссылка"])
    for it in items:
        ws.append([it["title"], it["price"], it["location"], it["link"]])
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    get_or_create_user(message.from_user.id)
    await message.answer(
        f"📊 Avito Radar\n\n"
        f"Отправь поисковый запрос — я соберу объявления в Excel.\n"
        f"Пример: \"iPhone 15\" или \"квартира\"\n\n"
        f"🎁 Первые {FREE_DAYS} дня — бесплатно.\n"
        f"💳 /pay — подписка {PRICE}"
    )


@dp.message(Command("help"))
async def cmd_help(message: types.Message):
    await message.answer(
        f"Отправь текстовый запрос — бот ищет на Avito.\n"
        f"Пришлёт Excel: Название, Цена, Локация, Ссылка\n\n"
        f"🎁 {FREE_DAYS} дня бесплатно\n"
        f"💳 {PRICE}\n"
        f"📩 @Saidikcs"
    )


@dp.message(Command("pay"))
async def cmd_pay(message: types.Message):
    kb = types.InlineKeyboardMarkup(inline_keyboard=[
        [types.InlineKeyboardButton(text="💳 Купить подписку 300⭐/мес", callback_data="buy_sub")],
    ])
    await message.answer(f"💰 Подписка Avito Radar\n⭐ 300 Stars в месяц", reply_markup=kb)


@dp.callback_query(lambda c: c.data == "buy_sub")
async def buy_sub(callback: types.CallbackQuery):
    await callback.message.delete()
    await callback.message.answer_invoice(
        title="Avito Radar — 1 месяц",
        description="Поиск объявлений и Excel. 30 дней доступа.",
        payload=f"sub_{callback.from_user.id}",
        provider_token="",
        currency="XTR",
        prices=[LabeledPrice(label="Подписка 1 месяц", amount=300)],
    )
    await callback.answer()


@dp.pre_checkout_query()
async def pre_checkout(pq: PreCheckoutQuery):
    await pq.answer(ok=True)


@dp.message(F.successful_payment)
async def paid(message: types.Message):
    users = _get_users()
    uid = str(message.from_user.id)
    if uid not in users:
        users[uid] = {"trial_start": datetime.now().isoformat(), "subscribed_until": None}
    users[uid]["subscribed_until"] = (datetime.now() + timedelta(days=30)).isoformat()
    _mark_dirty()
    await _flush_users()
    await message.answer("✅ Оплата прошла! Подписка на 30 дней активирована.")


@dp.message(Command("admin"))
async def cmd_admin(message: types.Message):
    if message.from_user.username != "Saidikcs":
        return
    kb = types.InlineKeyboardMarkup(inline_keyboard=[
        [types.InlineKeyboardButton(text="Пользователи", callback_data="admin_users")],
    ])
    await message.answer("Админ:", reply_markup=kb)


@dp.callback_query(lambda c: c.data == "admin_users")
async def admin_users(cb: types.CallbackQuery):
    if cb.from_user.username != "Saidikcs":
        return
    users = _get_users()
    text = "\n".join([f"{'✅' if is_subscribed(int(uid)) else '❌'} {uid}" for uid in users]) or "Нет пользователей"
    await cb.message.edit_text(text)
    await cb.answer()


@dp.message()
async def handle_search(message: types.Message):
    if not message.text or message.text.startswith("/"):
        return
    if not is_subscribed(message.from_user.id):
        await message.answer(f"😔 Бесплатный период закончился.\n💳 /pay — продлить за {PRICE}")
        return

    await message.answer(f"🔍 Ищу на Avito: \"{message.text[:50]}\"...")
    users = _get_users()
    uid = str(message.from_user.id)
    if uid in users:
        users[uid]["searches"] = users[uid].get("searches", 0) + 1
        users[uid]["last_seen"] = datetime.now().isoformat()
        _mark_dirty()

    try:
        items = await search_avito(message.text)
    except Exception:
        await message.answer("Ошибка при поиске. Попробуй позже.")
        return

    if not items:
        await message.answer("Ничего не найдено. Попробуй другой запрос.")
        return

    excel = make_excel(items)
    top = items[0]
    await message.answer_document(
        BufferedInputFile(excel.read(), filename="avito_search.xlsx"),
        caption=f"✅ Найдено: {len(items)}\n🏆 {top['title']} — {top['price']}₽",
    )


def migrate():
    users = _get_users()
    changed = False
    for u in users.values():
        for k in ("searches", "last_seen", "blocked"):
            if k not in u:
                u[k] = None if k == "last_seen" else (False if k == "blocked" else 0)
                changed = True
    if changed:
        _save_users(users)


async def main():
    migrate()
    asyncio.create_task(_periodic_flush())
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
