import asyncio
import json
import os
from datetime import datetime
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import FSInputFile, BotCommand
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from google import genai
from google.genai import types as genai_types

# ================= SOZLAMALAR =================
BOT_TOKEN = "9183078963:AAHG4gMv_RVsj8LyEf2K0pVm_1LGFwyLHDA"
GEMINI_API_KEY = "AQ.Ab8RN6I-1vjinSidjP3RPmykXMbezIClJWd414ZkdXtnghjwHABOT"
ADMIN_ID = 510826259
PAROL = "OMBORXONA"

FILIALLAR = {
    "andijon_shahar": {
        "nomi": "Andijon shahar",
        "masullar": ["elyor aka", "alimov7722", "jasur aka", "998990548674"]
    },
    "andijon_tuman": {
        "nomi": "Andijon tuman",
        "masullar": ["ramazonovanvar", "anvar aka", "sarvar aka"]
    },
    "asaka_tuman": {
        "nomi": "Asaka tumani",
        "masullar": ["sharof_imamov", "sharof aka", "farxod aka", "998977859310"]
    },
    "oltinkol_tuman": {
        "nomi": "Oltinko'l tumani",
        "masullar": ["javlon3355", "javlon aka", "nurmuxammad aka", "770066454", "esonov03", "samandar aka"]
    }
}
# ==============================================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
ai_client = genai.Client(api_key=GEMINI_API_KEY)

parol_kutish = {}

def excelga_yozish(filial_kodi, data, user_full_name, asl_xabar):
    filial_nomi = FILIALLAR[filial_kodi]["nomi"]
    fayl_nomi = f"Omborxona_{filial_nomi}.xlsx"
    
    operatsiya_turi = data.get("operatsiya", "Kirim").upper()
    sheet_name = "KIRIM" if "KIRIM" in operatsiya_turi else "CHIQIM"
    
    if not os.path.exists(fayl_nomi):
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "KIRIM"
        ws2 = wb.create_sheet(title="CHIQIM")
        
        headers = [
            "No", "Sana", "MAHSULOT NOMI", "SONI", 
            "BIRLIGI", "NARXI", "SUMMASI", 
            "JOY", "ETKAZUVCHI", "MAS'UL", "Asl xabar"
        ]
        
        for ws in [ws1, ws2]:
            ws.append([f"OMBORXONA {ws.title} - {filial_nomi}"])
            ws.append(headers)
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            for col_num, cell in enumerate(ws[2], 1):
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
        wb.save(fayl_nomi)
        
    wb = openpyxl.load_workbook(fayl_nomi)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(title=sheet_name)
    else:
        ws = wb[sheet_name]
        
    row_num = ws.max_row + 1
    
    kelgan_sana = data.get("sana", "")
    sana = kelgan_sana if kelgan_sana else datetime.now().strftime("%d.%m.%Y %H:%M")
    
    mahsulot = data.get("mahsulot", "")
    miqdor = data.get("miqdor", 1)
    birlik = data.get("birlik", "dona")
    narx = data.get("narx", 0)
    summa = f"=F{row_num}*D{row_num}" if narx else 0
    joy = data.get("joy", "-")
    etkazuvchi = data.get("etkazuvchi", "-")
    
    qator = [
        row_num - 2,
        sana,
        mahsulot,
        miqdor,
        birlik,
        narx,
        summa,
        joy,
        etkazuvchi,
        user_full_name,
        asl_xabar
    ]
    ws.append(qator)
    wb.save(fayl_nomi)

async def xabarni_tahlil_qilish(matn):
    prompt = f"""
    Quyidagi Telegram guruh xabarini tahlil qil. 
    MUHIM QOIDALAR:
    1. Agar xabar naqd pulda olingan xaridlarga, korporativ kartadan qilingan xaridlarga yoki omborxona kirim-chiqimiga aloqador bo'lmagan boshqa maishiy/umumiy xaridlarga tegishli bo'lsa, unda "mahsulot" maydoniga "RAD" deb yoz.
    2. Agar bu haqiqiy omborxona omborga kelgan yoki ombordan chiqqan ehtiyot qism / material bo'lsa, uni quyidagi formatda JSON qilib qaytar:
    
    Xabar: "{matn}"
    
    Parametrlar:
    - operatsiya: "Kirim" yoki "Chiqim"
    - mahsulot: Mahsulot nomi (agar naqd/korporativ karta bo'lsa aynan "RAD" deb yozgin)
    - miqdor: Son sifatida
    - birlik: O'lchov birligi
    - narx: Narxi (agar yo'q bo'lsa 0)
    - joy: Qayerdan olingani yoki qayerga berilgani
    - etkazuvchi: Yetkazib beruvchi shaxs
    - sana: Xabarda ko'rsatilgan sana va vaqt (masalan: "01.11.2026"). Agar xabarda sana aniq yozilgan bo'lsa, o'shani yoz, bo'sh bo'lsa "".
    """
    
    schema = {
        "type": "OBJECT",
        "properties": {
            "operatsiya": {"type": "STRING", "enum": ["Kirim", "Chiqim"]},
            "mahsulot": {"type": "STRING"},
            "miqdor": {"type": "NUMBER"},
            "birlik": {"type": "STRING"},
            "narx": {"type": "NUMBER"},
            "joy": {"type": "STRING"},
            "etkazuvchi": {"type": "STRING"},
            "sana": {"type": "STRING"}
        },
        "required": ["operatsiya", "mahsulot", "miqdor"]
    }

    response = ai_client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=genai_types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=schema,
            temperature=0.1
        )
    )
    return json.loads(response.text)

def filialni_aniqla(uid, uname, fname):
    for kod, info in FILIALLAR.items():
        for m in info["masullar"]:
            if m in uname or m in fname or m in uid:
                return kod
    return None

# Guruhdagi har qanday xabarni tutib qolish
@dp.message(F.text)
async def guruh_xabari(message: types.Message):
    try:
        # Faqat guruh yoki superguruh ekanligini tekshiramiz
        if message.chat.type not in {"group", "supergroup"}:
            return

        user = message.from_user
        if not user:
            return

        user_id_str = str(user.id)
        username = (user.username or "").lower()
        full_name = (user.full_name or "").lower()
        
        filial_kodi = filialni_aniqla(user_id_str, username, full_name)
        
        if filial_kodi:
            print(f"[{FILIALLAR[filial_kodi]['nomi']}] Xabar keldi: {message.text}")
            tahlil = await xabarni_tahlil_qilish(message.text)
            mahsulot = tahlil.get("mahsulot", "")
            
            if tahlil and mahsulot and mahsulot.upper() != "RAD":
                excelga_yozish(filial_kodi, tahlil, user.full_name, message.text)
                print("-> Excelga yozildi!")
    except Exception as e:
        print(f"Xatolik guruhda: {e}")

@dp.message(F.chat.type == "private", Command("hisobot"))
async def hisobot_sorash(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        return
    
    parol_kutish[message.from_user.id] = True
    await message.answer("?? Iltimos, hisobotlarni olish uchun parolni kiriting:")

@dp.message(F.chat.type == "private", F.text)
async def parolni_tekshirish(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        return
    
    if parol_kutish.get(message.from_user.id):
        if message.text.strip() == PAROL:
            parol_kutish[message.from_user.id] = False
            
            yuborildi = False
            for kod, info in FILIALLAR.items():
                fayl_nomi = f"Omborxona_{info['nomi']}.xlsx"
                if os.path.exists(fayl_nomi):
                    await message.answer_document(FSInputFile(fayl_nomi))
                    yuborildi = True
                    
            if not yuborildi:
                await message.answer("?? Hozircha hech qanday ma'lumot yozilmagan.")
        else:
            await message.answer("? Noto'g'ri parol! Qaytadan urinib ko'ring yoki /hisobot buyrug'ini bosing.")

async def main():
    print("Omborxona boti tezkor rejimda ishga tushdi...")
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())