# 🎓 Akademik Işler Boty

Telegram boty — Referat, Doklad we Prezentasiýa taýarlaýar.

## Faýllar

```
├── bot.py           # Esasy bot kody
├── template.docx    # Kapylyk sahypasy şablony
├── requirements.txt # Python garaşlylyklary
├── render.yaml      # Render.com sazlama
└── .gitignore
```

## Hyzmatlary

| Hyzmat | Baha |
|--------|------|
| 📄 Referat | 149 ★ Telegram Stars |
| 🎤 Doklad | 149 ★ Telegram Stars |
| 📊 Prezentasiýa | 149 ★ Telegram Stars |

🎁 **1-nji sargyt MUGT!**

## Diller

Bot 3 dilde işleýär: 🇹🇲 Türkmen / 🇷🇺 Rus / 🇬🇧 Iňlis

## Deploy (Render.com)

### 1. GitHub-a ýükläň
```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/SIZIŇ_USERNAME/REPO_ADY.git
git push -u origin main
```

### 2. Render.com sazlamasy
- render.com → New → Background Worker
- GitHub repo saýlaň
- **Disks** → Add: Name=`bot-data`, Mount=/data, Size=1GB

### 3. Environment Variables

| Key | Value |
|-----|-------|
| `BOT_TOKEN` | @BotFather-dan |
| `DEEPSEEK_API_KEY` | platform.deepseek.com |
| `STABILITY_API_KEY` | platform.stability.ai |
| `DB_PATH` | `/data/bot.db` |

### 4. Deploy
Manual Deploy → Deploy latest commit

## Admin komandalar

```
/admin   — kömek
/stats   — statistika
/users   — ulanyjylar
/orders  — garaşyp duranlar
/balance — girdeji
```
