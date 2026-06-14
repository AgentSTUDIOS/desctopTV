import asyncio
import os
import requests
import base64
import subprocess
import io
from datetime import datetime, timedelta
from flask import Flask, jsonify, render_template_string, send_file

from winsdk.windows.media.control import GlobalSystemMediaTransportControlsSessionManager as SessionManager
from winsdk.windows.storage.streams import DataReader, Buffer
import win32gui

app = Flask(__name__, static_folder='static')

# ========================== ЯРЛЫКИ ==========================
LINKS_DIR = os.path.join(os.path.dirname(__file__), 'links')

@app.route('/api/links/<int:num>/icon')
def get_link_icon(num):
    """Извлекает иконку из .lnk-ярлыка через PowerShell и возвращает PNG."""
    lnk_path = os.path.join(LINKS_DIR, str(num))

    lnk_file = None
    if os.path.isdir(lnk_path):
        for f in os.listdir(lnk_path):
            if f.lower().endswith('.lnk'):
                lnk_file = os.path.join(lnk_path, f)
                break

    if not lnk_file:
        return '', 404

    ps_script = f"""
Add-Type -AssemblyName System.Drawing
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut('{lnk_file.replace("'", "''")}')
$target = $shortcut.TargetPath
$iconLoc = $shortcut.IconLocation
$iconPath = $iconLoc.Split(',')[0].Trim()
if ($iconPath -eq '' -or -not (Test-Path $iconPath)) {{
    $iconPath = $target
}}
try {{
    $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
    $bmp = $icon.ToBitmap()
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    [Convert]::ToBase64String($ms.ToArray())
}} catch {{
    ''
}}
"""
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-Command', ps_script],
            capture_output=True, text=True, timeout=8
        )
        b64 = result.stdout.strip()
        if b64:
            img_bytes = base64.b64decode(b64)
            return send_file(io.BytesIO(img_bytes), mimetype='image/png')
    except Exception:
        pass
    return '', 404


@app.route('/api/links/<int:num>/launch', methods=['POST'])
def launch_link(num):
    """Запускает ярлык из папки links/<num>/."""
    lnk_path = os.path.join(LINKS_DIR, str(num))

    lnk_file = None
    if os.path.isdir(lnk_path):
        for f in os.listdir(lnk_path):
            if f.lower().endswith('.lnk'):
                lnk_file = os.path.join(lnk_path, f)
                break

    if not lnk_file:
        return jsonify({"success": False, "error": "Ярлык не найден"})

    try:
        os.startfile(lnk_file)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


# ========================== НАСТРОЙКИ ==========================
WEATHER_URL = "https://api.open-meteo.com/v1/forecast?latitude=55.7558&longitude=37.6173&current_weather=true&daily=temperature_2m_max&timezone=Europe%2FMoscow"
MOEX_BASE = "https://iss.moex.com/iss"

CURRENCY_TICKERS = [
    "USD000UTSTOM", "EUR_RUB__TOM", "CNYRUB_TOM",
    "GBPRUB_TOM", "JPYRUB_TOM", "CHFRUB_TOM", "TRYRUB_TOM",
    "KZTRUB_TOM", "UAHRUB_TOM", "BYNRUB_TOM"
]

STOCK_TICKERS = [
    "SBER", "GAZP", "LKOH", "ROSN", "GMKN",
    "VTBR", "TATN", "NLMK", "MGNT", "YNDX"
]

CRYPTO_IDS = [
    "bitcoin", "ethereum", "binancecoin", "solana", "ripple",
    "cardano", "dogecoin", "polkadot", "litecoin", "matic-network"
]
CRYPTO_NAMES = {
    "bitcoin": "Bitcoin", "ethereum": "Ethereum", "binancecoin": "BNB",
    "solana": "Solana", "ripple": "XRP", "cardano": "Cardano",
    "dogecoin": "Dogecoin", "polkadot": "Polkadot", "litecoin": "Litecoin",
    "matic-network": "Polygon"
}


# ========================== ФИНАНСЫ ==========================
def fetch_moex_securities(tickers: list, engine: str, market: str, board: str) -> dict:
    result = {}
    today = datetime.now().strftime("%Y-%m-%d")
    from_date = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")

    for ticker in tickers:
        try:
            url_q = (
                f"{MOEX_BASE}/engines/{engine}/markets/{market}"
                f"/boards/{board}/securities/{ticker}.json"
                f"?iss.meta=off&iss.only=marketdata&marketdata.columns=LAST,LCLOSEPRICE"
            )
            r = requests.get(url_q, timeout=5)
            last_price = None
            if r.status_code == 200:
                data = r.json()
                rows = data.get("marketdata", {}).get("data", [])
                if rows and rows[0]:
                    last_price = rows[0][0] or rows[0][1]

            url_h = (
                f"{MOEX_BASE}/engines/{engine}/markets/{market}"
                f"/boards/{board}/securities/{ticker}/candles.json"
                f"?from={from_date}&till={today}&interval=24&iss.meta=off"
                f"&candles.columns=close"
            )
            rh = requests.get(url_h, timeout=5)
            history = []
            if rh.status_code == 200:
                hdata = rh.json()
                history = [row[0] for row in hdata.get("candles", {}).get("data", []) if row[0]]
                history = history[-15:]
                if not last_price and history:
                    last_price = history[-1]

            result[ticker] = {"price": last_price, "history": history}
        except Exception:
            result[ticker] = {"price": None, "history": []}
    return result


def get_currency_code(ticker: str) -> str:
    if ticker == "USD000UTSTOM":
        return "USD"
    if ticker == "EUR_RUB__TOM":
        return "EUR"
    if ticker == "CNYRUB_TOM":
        return "CNY"
    if ticker.endswith("RUB_TOM"):
        return ticker[:3]
    return ticker[:3]


def build_currency_list(raw: dict) -> list:
    out = []
    for ticker, data in raw.items():
        code = get_currency_code(ticker)
        price = data.get("price")
        history = data.get("history", [])
        value = f"{price:.2f} ₽" if price is not None else "— ₽"
        out.append({"id": code.lower(), "name": code, "value": value, "history": history})
    return out


def build_stock_list(raw: dict) -> list:
    names = {
        "SBER": "Сбербанк", "GAZP": "Газпром", "LKOH": "Лукойл",
        "ROSN": "Роснефть", "GMKN": "Норникель", "VTBR": "ВТБ",
        "TATN": "Татнефть", "NLMK": "НЛМК", "MGNT": "Магнит",
        "YNDX": "Яндекс"
    }
    out = []
    for ticker, data in raw.items():
        name = names.get(ticker, ticker)
        price = data.get("price")
        history = data.get("history", [])
        value = f"{price:.2f} ₽" if price is not None else "— ₽"
        out.append({"id": ticker.lower(), "name": name, "value": value, "history": history})
    return out


def fetch_crypto_prices() -> list:
    ids = ",".join(CRYPTO_IDS)
    url = f"https://api.coingecko.com/api/v3/simple/price?ids={ids}&vs_currencies=usd"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    try:
        r = requests.get(url, timeout=6, headers=headers)
        if r.status_code == 200:
            data = r.json()
            result = []
            for crypto_id in CRYPTO_IDS:
                name = CRYPTO_NAMES.get(crypto_id, crypto_id.capitalize())
                price_usd = data.get(crypto_id, {}).get("usd")
                value = f"${price_usd:,.2f}" if price_usd is not None else "— USD"
                result.append({
                    "id": crypto_id.split('-')[0],
                    "name": name,
                    "value": value,
                    "history": []
                })
            return result
    except Exception:
        pass
    return [{"id": cid[:3], "name": CRYPTO_NAMES.get(cid, cid.capitalize()), "value": "— USD", "history": []}
            for cid in CRYPTO_IDS]


# ========================== МЕДИАПЛЕЕР ==========================
async def get_media_info():
    try:
        manager = await SessionManager.request_async()
        session = manager.get_current_session()
        if session:
            info = await session.try_get_media_properties_async()
            playback = session.get_playback_info()
            is_playing = playback.playback_status == 4
            title  = info.title  if info.title  else "Неизвестный трек"
            artist = info.artist if info.artist else "Неизвестный исполнитель"

            cover_base64 = ""
            if info.thumbnail:
                try:
                    stream_ref = info.thumbnail
                    stream = await stream_ref.open_read_async()
                    size = stream.size
                    if size > 0:
                        buffer = Buffer(size)
                        await stream.read_async(buffer, size, 0)
                        reader = DataReader.from_buffer(buffer)
                        bytes_data = bytearray(size)
                        reader.read_bytes(bytes_data)
                        cover_base64 = "data:image/jpeg;base64," + base64.b64encode(bytes_data).decode('utf-8')
                except Exception:
                    pass

            return {"track": f"{artist} - {title}", "is_playing": is_playing, "cover": cover_base64}
    except Exception:
        pass
    return {"track": "Плеер Windows неактивен", "is_playing": False, "cover": ""}


async def control_media(action):
    try:
        manager = await SessionManager.request_async()
        session = manager.get_current_session()
        if session:
            if action == 'play_pause':
                if session.get_playback_info().playback_status == 4:
                    await session.try_pause_async()
                else:
                    await session.try_play_async()
            elif action == 'next':
                await session.try_skip_next_async()
            elif action == 'prev':
                await session.try_skip_previous_async()
            return True
    except Exception:
        pass
    return False


# ========================== FLASK ROUTES ==========================
@app.route('/')
def index():
    with open('index.html', 'r', encoding='utf-8') as f:
        return render_template_string(f.read())


@app.route('/api/player/<action>', methods=['POST'])
def player_control(action):
    success = asyncio.run(control_media(action))
    return jsonify({"success": success})


@app.route('/api/player/info')
def player_info():
    return jsonify(asyncio.run(get_media_info()))


@app.route('/api/system/data')
def get_system_data():
    weather_data = {"current": "—°C", "forecast": ["—°C"] * 7}
    try:
        r = requests.get(WEATHER_URL, timeout=3)
        if r.status_code == 200:
            res = r.json()
            weather_data["current"]  = f"{res['current_weather']['temperature']}°C"
            weather_data["forecast"] = [f"{t}°C" for t in res['daily']['temperature_2m_max']]
    except Exception:
        pass

    raw_fx     = fetch_moex_securities(CURRENCY_TICKERS, "currency", "selt",  "CETS")
    raw_stocks = fetch_moex_securities(STOCK_TICKERS,    "stock",    "shares", "TQBR")

    currencies_list = build_currency_list(raw_fx)
    stocks_list     = build_stock_list(raw_stocks)
    crypto_list     = fetch_crypto_prices()

    today = datetime.now()
    dates_list = [(today - timedelta(days=i)).strftime("%d.%m") for i in range(14, -1, -1)]

    notifications = []
    def win_enum_handler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and len(notifications) < 6 and title not in [
                "Program Manager", "Settings", "Служба глобальных сеансов мультимедиа"
            ]:
                notifications.append(title)
    try:
        win32gui.EnumWindows(win_enum_handler, None)
    except Exception:
        notifications = ["Нет активных уведомлений"]

    return jsonify({
        "weather":       weather_data,
        "currencies":    currencies_list,
        "stocks":        stocks_list,
        "crypto":        crypto_list,
        "dates":         dates_list,
        "notifications": notifications if notifications else ["Лента уведомлений пуста"],
    })


@app.route('/api/debug/moex')
def debug_moex():
    results = {}
    ticker = "USD000UTSTOM"
    url_q = f"https://iss.moex.com/iss/engines/currency/markets/selt/boards/CETS/securities/{ticker}.json?iss.meta=off&iss.only=marketdata&marketdata.columns=LAST,LCLOSEPRICE"
    r = requests.get(url_q, timeout=5)
    results["currency_status"] = r.status_code
    results["currency_raw"] = r.json() if r.status_code == 200 else r.text

    ticker2 = "SBER"
    url_q2 = f"https://iss.moex.com/iss/engines/stock/markets/shares/boards/TQBR/securities/{ticker2}.json?iss.meta=off&iss.only=marketdata&marketdata.columns=LAST,LCLOSEPRICE"
    r2 = requests.get(url_q2, timeout=5)
    results["stock_status"] = r2.status_code
    results["stock_raw"] = r2.json() if r2.status_code == 200 else r2.text

    return jsonify(results)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
