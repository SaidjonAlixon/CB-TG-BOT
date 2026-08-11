# FILIAL ATTENDANCE

Telegram bot orqali filial xodimlarining kelish-ketishini qayd etadi, smena qoidalari bo‘yicha kechikish va erta ketishni hisoblaydi hamda professional Excel hisobotlarini yaratadi.

## Run & Operate

- `python main.py` — Telegram botni lokal ishga tushirish (`.env` kerak)
- `uv run python main.py` — xuddi shu, uv orqali
- Required env: `DATABASE_URL`, `TELEGRAM_BOT_TOKEN`
- Optional env: `ADMIN_TELEGRAM_IDS` — vergul bilan ajratilgan Telegram ID raqamlari

## Railway deploy

1. Railway projectda **PostgreSQL** bo‘lsin.
2. Bot service yarating (GitHub repo yoki `railway up`).
3. Bot service **Variables**:
   - `TELEGRAM_BOT_TOKEN` — BotFather token
   - `DATABASE_URL` = `${{Postgres.DATABASE_URL}}` (Variable Reference)
   - `ADMIN_TELEGRAM_IDS` — admin Telegram ID(lar)
4. Build: `Dockerfile` / `railway.toml` orqali avtomatik.
5. Start command: `python main.py` (worker; HTTP port kerak emas).
6. Deploy keyin Logs’da `Application started` ko‘rinsin.

### Ma’lumotlar saqlanishi

- Barcha filial / xodim / davomat / admin yozuvlari **PostgreSQL**da.
- Bot restart, redeploy yoki kodni qayta o‘rnatish — **mavjud yozuvlarni o‘chirmaydi** (`CREATE TABLE IF NOT EXISTS`).
- O‘chadi faqat: Railway’da **Postgres service yoki volume** ni o‘chirib tashlasangiz, yoki bazani qo‘lda DROP qilsangiz.
- Shu sabab Postgres’ni o‘chirmang; faqat bot service’ni redeploy qiling.

Ichki `postgres.railway.internal` faqat Railway ichida ishlaydi. Lokal uchun Public URL yoki alohida Postgres kerak.

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- API: Express 5
- DB: PostgreSQL + Drizzle ORM
- Validation: Zod (`zod/v4`), `drizzle-zod`
- API codegen: Orval (from OpenAPI spec)
- Build: esbuild (CJS bundle)

## Where things live

_Populate as you build — short repo map plus pointers to the source-of-truth file for DB schema, API contracts, theme files, etc._

## Architecture decisions

_Populate as you build — non-obvious choices a reader couldn't infer from the code (3-5 bullets)._

## Product

- Xodimlar telefon, F.I.Sh., filial, lavozim va smena orqali ro‘yxatdan o‘tadi.
- “KELDIM” va “KETDIM” tugmalari real vaqtni saqlaydi.
- 1-smena (08:15 / 17:00) va 2-smena (17:15 / 23:45, yarim tundan o‘tish bilan) hisoblanadi.
- Admin filiallarni Excel’dan import qiladi va kunlik/haftalik/oylik Excel hisobotlarini oladi.

## User preferences

_Populate as you build — explicit user instructions worth remembering across sessions._

## Gotchas

_Populate as you build — sharp edges, "always run X before Y" rules._

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
