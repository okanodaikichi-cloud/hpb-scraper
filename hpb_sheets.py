"""
HotPepper Beauty HBL空き率集計 → Google Sheets書き込み
GitHub Actions で週次実行される
"""

import asyncio
import json
import os
import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Dict, List, Optional

import gspread
from google.oauth2.service_account import Credentials
from playwright.async_api import Page, async_playwright

# ──────────────────────────────────────────────
# ⚙️ 設定
# ──────────────────────────────────────────────

START_DATE = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
DAYS = 14

SALON_URLS = [
    "https://beauty.hotpepper.jp/kr/slnH000579091/",
    "https://beauty.hotpepper.jp/kr/slnH000622977/",
    "https://beauty.hotpepper.jp/kr/slnH000703399/",
    "https://beauty.hotpepper.jp/kr/slnH000771305/",
]

# Google SheetsのスプレッドシートID（URLの /d/〇〇〇/ の部分）
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "YOUR_SPREADSHEET_ID_HERE")

HBL_KEYWORDS = ["HBL", "ハリウッドブロウリフト"]
DAYS = max(1, min(14, DAYS))
TARGET_DATES = [START_DATE + timedelta(days=i) for i in range(DAYS)]
TARGET_DATE_STRS = {d.strftime("%Y%m%d") for d in TARGET_DATES}


# ── データクラス ──

def is_weekend(dt: datetime) -> bool:
    return dt.weekday() >= 5


@dataclass
class DayData:
    date: datetime
    maru: int = 0
    batsu: int = 0
    is_holiday: bool = False

    @property
    def total_operating(self) -> int:
        return self.maru + self.batsu

    @property
    def is_closed(self) -> bool:
        return self.is_holiday or self.total_operating == 0


@dataclass
class SalonData:
    url: str
    name: str
    coupon_name: str = ""
    has_hbl: bool = False
    error: Optional[str] = None
    days: List[DayData] = field(default_factory=list)

    def _open_days(self, weekend: Optional[bool]) -> List[DayData]:
        days = [d for d in self.days if not d.is_closed]
        if weekend is None:
            return days
        return [d for d in days if is_weekend(d.date) == weekend]

    def vacancy_rate(self, weekend: Optional[bool] = None) -> Optional[float]:
        days = self._open_days(weekend)
        m = sum(d.maru for d in days)
        t = sum(d.total_operating for d in days)
        return m / t if t > 0 else None


def pct(v: Optional[float]) -> str:
    return f"{v:.1%}" if v is not None else "データなし"


# ── JavaScript: カレンダーデータ抽出 ──

EXTRACT_JS = """
() => {
    try {
        const timeTable = document.querySelector('.moreInnerTable.timeTableLeft');
        const timeLabels = timeTable
            ? Array.from(timeTable.querySelectorAll('th.timeCell'))
                .map(th => th.innerText.trim().replace(/：/g, ':'))
            : [];

        const months = Array.from(document.querySelectorAll('.monthCell'))
            .map(mc => mc.innerText.trim());

        const dateCells = Array.from(
            document.querySelectorAll('.dayCellContainer th')
        ).filter(th => !th.classList.contains('weekPaging'));

        let dateInfos = [];
        for (const th of dateCells) {
            const text = th.innerText.trim();
            const dayMatch = text.match(/(\\d+)/);
            if (!dayMatch) continue;
            dateInfos.push({ day: parseInt(dayMatch[1]), colspan: parseInt(th.colSpan) || 1 });
        }

        let curYear = 2026, curMonth = new Date().getMonth() + 1;
        const firstMonthText = months[0] || '';
        const ymm = firstMonthText.match(/(\\d{4})年(\\d{1,2})月/);
        if (ymm) { curYear = parseInt(ymm[1]); curMonth = parseInt(ymm[2]); }

        let resolvedDates = [];
        for (let i = 0; i < dateInfos.length; i++) {
            if (i > 0 && dateInfos[i].day < dateInfos[i-1].day) {
                curMonth++;
                if (curMonth > 12) { curMonth = 1; curYear++; }
            }
            const dateStr = `${curYear}${String(curMonth).padStart(2,'0')}${String(dateInfos[i].day).padStart(2,'0')}`;
            resolvedDates.push({ dateStr, colspan: dateInfos[i].colspan });
        }

        const dataTables = Array.from(
            document.querySelectorAll('.moreInnerTable:not(.timeTableLeft)')
        );

        const tableToDate = [];
        let tableIdx = 0;
        for (const di of resolvedDates) {
            for (let s = 0; s < di.colspan; s++) {
                tableToDate[tableIdx++] = di.dateStr;
            }
        }

        const raw = {};
        resolvedDates.forEach(di => {
            if (!raw[di.dateStr])
                raw[di.dateStr] = { maruTimes: new Set(), batsuRows: new Set(), isHoliday: false };
        });

        dataTables.forEach((table, idx) => {
            const dateStr = tableToDate[idx];
            if (!dateStr || !raw[dateStr]) return;
            const dr = raw[dateStr];

            if (table.querySelector('.closeColInner')) {
                dr.isHoliday = true;
                return;
            }

            table.querySelectorAll('td.open a[href]').forEach(a => {
                const dm = a.href.match(/rsvRequestDate1=(\\d{8})/);
                const tm = a.href.match(/rsvRequestTime1=(\\d{4})/);
                if (!dm || !tm) return;
                const t = tm[1];
                dr.maruTimes.add(t.slice(0,2) + ':' + t.slice(2));
            });

            const rows = table.querySelectorAll('tbody tr');
            rows.forEach((row, rowIdx) => {
                if (row.querySelector('td.closed')) dr.batsuRows.add(rowIdx);
            });
        });

        const result = {};
        for (const [dateStr, dr] of Object.entries(raw)) {
            if (dr.isHoliday) {
                result[dateStr] = { maru: 0, batsu: 0, isHoliday: true };
                continue;
            }
            const maraTimes = dr.maruTimes;
            const batsuTimes = new Set();
            for (const rowIdx of dr.batsuRows) {
                const tLabel = timeLabels[rowIdx];
                if (tLabel && !maraTimes.has(tLabel)) batsuTimes.add(tLabel);
            }
            result[dateStr] = { maru: maraTimes.size, batsu: batsuTimes.size, isHoliday: false };
        }

        return JSON.stringify({
            ok: true,
            dates: resolvedDates.map(d => d.dateStr),
            tableCount: dataTables.length,
            result
        });
    } catch(e) {
        return JSON.stringify({ ok: false, error: e.toString() });
    }
}
"""


# ── ヘルパー ──

async def get_salon_name(page: Page) -> str:
    for sel in [".slnName", "h1.ttlSalon", ".salonName", "h1[class*='name']", "h1"]:
        try:
            elem = await page.query_selector(sel)
            if elem:
                text = (await elem.inner_text()).strip()
                if text and len(text) < 100:
                    return text
        except Exception:
            pass
    return ""


async def find_hbl_reserve_url(page: Page):
    js = """
    () => {
        const keywords = ['HBL', 'ハリウッドブロウリフト'];
        const nameElems = document.querySelectorAll('.couponMenuName, [class*="couponMenuName"]');
        for (const elem of nameElems) {
            const text = elem.innerText || '';
            if (!keywords.some(kw => text.includes(kw))) continue;
            let p = elem.parentElement;
            for (let i = 0; i < 8; i++) {
                if (!p) break;
                const link = p.querySelector('a[href*="add=0"]');
                if (link) return { name: text.trim().substring(0, 80), href: link.href };
                p = p.parentElement;
            }
        }
        const allElems = document.querySelectorAll('*');
        for (const elem of allElems) {
            const text = elem.innerText || '';
            if (!keywords.some(kw => text.includes(kw))) continue;
            if (text.length > 400) continue;
            const link = elem.querySelector('a[href*="add=0"]');
            if (link) return { name: text.trim().substring(0, 80), href: link.href };
        }
        return null;
    }
    """
    result = await page.evaluate(js)
    if result:
        return result.get("href"), result.get("name", "")
    return None, ""


async def get_next_week_url(page: Page) -> Optional[str]:
    js = """() => {
        const a = document.querySelector('a.arrowPagingWeekR, a.jscCalNavi[href*="week"]');
        return a ? a.href : null;
    }"""
    return await page.evaluate(js)


# ── メインスクレイピング ──

async def scrape_salon(browser, salon_url: str) -> SalonData:
    salon_id = salon_url.rstrip("/").split("/")[-1]
    context = await browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        viewport={"width": 1280, "height": 900},
        locale="ja-JP",
    )
    page = await context.new_page()
    result = SalonData(url=salon_url, name=salon_id)

    try:
        print(f"\n[{salon_id}] サロン名取得中...")
        await page.goto(salon_url, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(2)
        name = await get_salon_name(page)
        if name:
            result.name = name
        print(f"  → {result.name}")

        await page.goto(salon_url.rstrip("/") + "/coupon/", wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(2)

        body = await page.content()
        result.has_hbl = any(kw in body for kw in HBL_KEYWORDS)

        if not result.has_hbl:
            result.error = "HBLクーポンなし"
            return result

        reserve_url, coupon_name = await find_hbl_reserve_url(page)
        result.coupon_name = coupon_name

        if not reserve_url:
            result.error = "HBL予約リンクが見つかりません"
            return result

        await page.goto(reserve_url, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(2)

        all_day_data: Dict[str, DayData] = {}
        collected = set()

        for week in range(3):
            raw = await page.evaluate(EXTRACT_JS)
            data = json.loads(raw)

            if not data.get("ok"):
                break

            result_week = data.get("result", {})
            for date_str, day_info in result_week.items():
                if date_str not in TARGET_DATE_STRS or date_str in collected:
                    continue
                dt = datetime.strptime(date_str, "%Y%m%d")
                dd = DayData(
                    date=dt,
                    maru=day_info.get("maru", 0),
                    batsu=day_info.get("batsu", 0),
                    is_holiday=day_info.get("isHoliday", False),
                )
                all_day_data[date_str] = dd
                collected.add(date_str)

            if TARGET_DATE_STRS.issubset(collected):
                break

            next_url = await get_next_week_url(page)
            if not next_url:
                break
            await page.goto(next_url, wait_until="domcontentloaded", timeout=30000)
            await asyncio.sleep(2)

        for d in TARGET_DATES:
            ds = d.strftime("%Y%m%d")
            result.days.append(all_day_data.get(ds, DayData(date=d, is_holiday=True)))
        result.days.sort(key=lambda d: d.date)

    except Exception as e:
        import traceback
        result.error = str(e)
        print(f"  エラー: {traceback.format_exc()}")
    finally:
        await context.close()

    return result


# ── Google Sheets 書き込み ──

def write_to_sheets(salon_results: List[SalonData]):
    # 環境変数からGoogle認証情報を取得
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise ValueError("環境変数 GOOGLE_CREDENTIALS が設定されていません")

    creds_dict = json.loads(creds_json)
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    wday = ["月", "火", "水", "木", "金", "土", "日"]
    run_date = datetime.today().strftime("%Y/%m/%d %H:%M")

    # ── シート1: 集計結果 ──
    try:
        ws1 = spreadsheet.worksheet("集計結果")
        ws1.clear()
    except gspread.WorksheetNotFound:
        ws1 = spreadsheet.add_worksheet(title="集計結果", rows=100, cols=10)

    headers1 = ["集計日時", "サロン名", "総稼働率", "土日稼働率", "平日稼働率", "備考", "サロンURL"]
    ws1.append_row(headers1)
    for s in salon_results:
        ws1.append_row([
            run_date,
            s.name,
            pct(s.vacancy_rate()),
            pct(s.vacancy_rate(True)),
            pct(s.vacancy_rate(False)),
            s.error or "",
            s.url,
        ])
    print("✔ 集計結果シート書き込み完了")

    # ── シート2: 日別詳細 ──
    try:
        ws2 = spreadsheet.worksheet("日別詳細")
        ws2.clear()
    except gspread.WorksheetNotFound:
        ws2 = spreadsheet.add_worksheet(title="日別詳細", rows=1000, cols=10)

    headers2 = ["集計日時", "サロン名", "日付", "曜日", "週末/平日",
                "◎(空き)コマ", "×(埋まり)コマ", "合計コマ", "空き率"]
    ws2.append_row(headers2)

    rows2 = []
    for s in salon_results:
        for d in s.days:
            rate = d.maru / d.total_operating if d.total_operating > 0 else None
            rows2.append([
                run_date,
                s.name,
                d.date.strftime("%Y/%m/%d"),
                wday[d.date.weekday()],
                "週末" if is_weekend(d.date) else "平日",
                "" if d.is_holiday else d.maru,
                "" if d.is_holiday else d.batsu,
                "" if d.is_holiday else d.total_operating,
                "休業日" if d.is_holiday else pct(rate),
            ])
    if rows2:
        ws2.append_rows(rows2)
    print("✔ 日別詳細シート書き込み完了")


# ── メイン実行 ──

async def main():
    print("=" * 55)
    print(f"HBL空き率集計")
    print(f"期間: {START_DATE.strftime('%Y/%m/%d')} 〜 "
          f"{(START_DATE + timedelta(days=DAYS-1)).strftime('%Y/%m/%d')} ({DAYS}日間)")
    print(f"対象サロン: {len(SALON_URLS)}件")
    print("=" * 55)

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"],
        )
        results: List[SalonData] = []
        for url in SALON_URLS:
            data = await scrape_salon(browser, url)
            results.append(data)
        await browser.close()

    write_to_sheets(results)
    print("\n✔ 全処理完了")


if __name__ == "__main__":
    asyncio.run(main())
