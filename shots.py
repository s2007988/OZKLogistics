import asyncio, pathlib
from playwright.async_api import async_playwright

ROOT = pathlib.Path(r"c:\Antigravity\lab_5_ozk_logistics")
OUT = ROOT / "screens"; OUT.mkdir(exist_ok=True)
SITE = (ROOT / "site" / "index.html").as_uri()

COMPETITORS = {
    "novaposhta": "https://novaposhta.ua/",
    "delivery": "https://www.delivery-auto.com/uk-UA",
    "sat": "https://www.sat.ua/",
    "meest": "https://meest.com/",
    "raben": "https://polska.raben-group.com/uk/",
    "gunsel": "https://www.dhl.com/ua-uk/home.html",
}

async def own(p):
    b = await p.chromium.launch()
    pg = await b.new_page(viewport={"width": 1440, "height": 860})
    errs = []
    pg.on("console", lambda m: errs.append(m.text) if m.type == "error" else None)
    pg.on("pageerror", lambda e: errs.append(str(e)))
    await pg.goto(SITE, wait_until="networkidle")
    await pg.wait_for_timeout(1200)
    await pg.screenshot(path=OUT / "own_hero.png")
    for sel, name in [("#services", "own_services"), ("#calculator", "own_calc"), ("#advantages", "own_adv")]:
        await pg.locator(sel).scroll_into_view_if_needed(); await pg.evaluate(f"document.querySelector('{sel}').scrollIntoView()")
        await pg.wait_for_timeout(900)
        await pg.screenshot(path=OUT / f"{name}.png")
    await pg.click("[data-demo='OZK-2026-0417']")
    await pg.evaluate("document.querySelector('#tracking').scrollIntoView()")
    await pg.wait_for_timeout(1500)
    await pg.screenshot(path=OUT / "own_tracking.png")
    await pg.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    await pg.wait_for_timeout(900)
    await pg.screenshot(path=OUT / "own_footer.png")
    # form test
    await pg.evaluate("document.querySelector('#order').scrollIntoView()")
    await pg.fill("#orderForm [name=name]", "Олег, ТОВ Будмакс")
    await pg.fill("#orderForm [name=phone]", "+380 67 123 45 67")
    await pg.check("#orderForm [name=agree]")
    await pg.click("#orderForm button[type=submit]")
    await pg.wait_for_timeout(500)
    print("FORM:", await pg.inner_text("#formMsg"))
    print("PRICE:", await pg.inner_text("#price"))
    await pg.screenshot(path=OUT / "own_order.png")
    m = await b.new_page(viewport={"width": 390, "height": 844}, device_scale_factor=2)
    await m.goto(SITE, wait_until="networkidle"); await m.wait_for_timeout(800)
    await m.screenshot(path=OUT / "own_mobile.png")
    print("ERRORS:", errs)
    await b.close()

async def comp(p):
    b = await p.chromium.launch()
    ctx = await b.new_context(viewport={"width": 1440, "height": 860}, locale="uk-UA",
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36")
    for name, url in COMPETITORS.items():
        pg = await ctx.new_page()
        try:
            await pg.goto(url, wait_until="domcontentloaded", timeout=45000)
            await pg.wait_for_timeout(6000)
            await pg.screenshot(path=OUT / f"c_{name}.png")
            print("OK", name, await pg.title())
        except Exception as e:
            print("FAIL", name, e)
        await pg.close()
    await b.close()

async def main():
    async with async_playwright() as p:
        await own(p)
        await comp(p)

asyncio.run(main())
