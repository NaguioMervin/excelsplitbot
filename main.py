import os
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import math
import asyncio

TOKEN = os.environ.get("BOT_TOKEN")  # use Replit secret for safety

# --- Create folder to store uploads ---
if not os.path.exists("uploads"):
    os.makedirs("uploads")


# --- /start command ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Hi! Send me an Excel file (.xlsx) and I'll split it into 2,500-row chunks."
    )


# --- Handle Excel uploads ---
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = await update.message.document.get_file()
    file_path = f"uploads/{update.message.document.file_name}"
    await file.download_to_drive(file_path)

    await update.message.reply_text("📥 File received. Processing...")

    # --- Split Excel file ---
    rows_per_file = 2500
    df = pd.read_excel(file_path, dtype=str, engine="openpyxl")
    total_rows = len(df)
    num_files = math.ceil(total_rows / rows_per_file)
    base_name = os.path.splitext(os.path.basename(file_path))[0]

    for i in range(num_files):
        start = i * rows_per_file
        end = min(start + rows_per_file, total_rows)
        part_df = df.iloc[start:end]
        part_df["Batch Label"] = f"Batch V{i+1}"
        out_name = f"{base_name}_V{i+1}.xlsx"
        part_df.to_excel(out_name, index=False)
        await update.message.reply_document(document=open(out_name, "rb"))
        os.remove(out_name)
        await asyncio.sleep(1)

    await update.message.reply_text("✅ All split files sent!")


# --- Clear temp files ---
async def clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    for f in os.listdir("uploads"):
        os.remove(os.path.join("uploads", f))
    await update.message.reply_text("🧹 Temporary files cleared.")


# --- Start bot ---
app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("clear", clear))
app.add_handler(
    MessageHandler(filters.Document.FileExtension("xlsx"), handle_file))

print("✅ Bot is running...")
app.run_polling()
