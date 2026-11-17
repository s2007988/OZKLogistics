import asyncio, pathlib
from playwright.async_api import async_playwright
OUT = pathlib.Path(r"c:\Antigravity\lab_5_ozk_logistics\screens")
URLS = {"sat": "https://www.sat.ua/", "meest": "https://ua.meest.com/", "raben": "https://ukraine.raben-group.com/", "lardi": "https://lardi-trans.com/", "ukrposhta": "https://www.ukrposhta.ua/ua", "dhl": "https://www.dhl.com/ua-uk/home.html"}
async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch()
        ctx = await b.new_context(viewport={"width": 1440, "height": 860}, locale="uk-UA",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36")
        for n, u in URLS.items():
            pg = await ctx.new_page()
            try:
                await pg.goto(u, wait_until="domcontentloaded", timeout=45000); await pg.wait_for_timeout(6000)
                for txt in ["Відхилити все", "Прийняти все", "Прийняти", "Accept all", "Погоджуюсь", "Зрозуміло"]:
                    try:
                        loc = pg.get_by_role("button", name=txt, exact=True)
                        if await loc.count(): await loc.first.click(timeout=2000); await pg.wait_for_timeout(1000); break
                    except Exception: pass
                await pg.screenshot(path=OUT / f"c_{n}.png"); print("OK", n, pg.url, await pg.title())
            except Exception as e: print("FAIL", n, str(e)[:120])
            await pg.close()
        await b.close()
asyncio.run(main())
