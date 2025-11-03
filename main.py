from keep_alive import keep_alive
import os
import pandas as pd
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes,
)
import math
import asyncio

TOKEN = os.environ.get("BOT_TOKEN")  # Secure token storage

# --- Create uploads folder ---
if not os.path.exists("uploads"):
    os.makedirs("uploads")


# ================================
# 📍 COMMANDS
# ================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Hi! I'm your Excel Assistant Bot.\n\n"
        "You can choose what I should do:\n"
        "1️⃣ /split - Split Excel into smaller parts\n"
        "2️⃣ /merge - Merge multiple Excel files\n"
        "3️⃣ /compare - Compare two Excel files for data verification\n"
        "4️⃣ /clear - Clear uploaded temporary files")


# --- Split mode ---
async def split_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["mode"] = "split"
    await update.message.reply_text(
        "📤 Split mode activated!\n\n"
        "Send me an Excel file (.xlsx), and I'll split it into 2,500-row parts by default.\n"
        "You can change this setting later using /setrows <number>.")


# --- Merge mode ---
async def merge_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["mode"] = "merge"
    context.user_data["merge_files"] = []
    await update.message.reply_text(
        "📥 Merge mode activated!\n\n"
        "Send me multiple Excel files (.xlsx), and I’ll combine them into one.\n"
        "Once done, I’ll automatically send you the merged file.")


# --- Compare mode ---
async def compare_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["mode"] = "compare"
    context.user_data["compare_files"] = []
    await update.message.reply_text(
        "📊 Compare mode activated!\n\n"
        "Send me *2 Excel files* you want to compare.\n"
        "I'll check if they are identical (even if rows are not in the same order)."
    )


# --- Set rows per file for split ---
async def setrows(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args or len(context.args) == 0:
        current = context.user_data.get("rows_per_file", 2500)
        await update.message.reply_text(
            f"📊 Current setting: {current} rows per file\n\n"
            f"To change it, use: /setrows <number>\n"
            f"Example: /setrows 5000")
        return

    try:
        rows = int(context.args[0])
        if rows < 1:
            await update.message.reply_text(
                "❌ Please enter a number greater than 0")
            return

        context.user_data["rows_per_file"] = rows
        await update.message.reply_text(f"✅ Set to {rows} rows per file!")
    except ValueError:
        await update.message.reply_text(
            "❌ Please enter a valid number\nExample: /setrows 5000")


# ================================
# 📁 FILE HANDLING
# ================================
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mode = context.user_data.get("mode")

    if not mode:
        await update.message.reply_text(
            "⚠️ Please choose a mode first using /split, /merge, or /compare.")
        return

    file = await update.message.document.get_file()
    file_path = f"uploads/{update.message.document.file_name}"
    await file.download_to_drive(file_path)

    if mode == "split":
        await update.message.reply_text("📥 File received. Processing split...")
        await process_split(update, context, file_path)

    elif mode == "merge":
        context.user_data["merge_files"].append(file_path)
        await update.message.reply_text(
            f"📎 Added: {update.message.document.file_name}")
        # If user sends "done", merge automatically
        if len(context.user_data["merge_files"]) > 1:
            await process_merge(update, context)

    elif mode == "compare":
        context.user_data["compare_files"].append(file_path)
        await update.message.reply_text(
            f"📎 Added for comparison: {update.message.document.file_name}")

        if len(context.user_data["compare_files"]) == 2:
            await compare_excels(update, context)


# ================================
# 🔹 SPLIT EXCEL
# ================================
async def process_split(update: Update, context: ContextTypes.DEFAULT_TYPE,
                        file_path):
    rows_per_file = context.user_data.get("rows_per_file", 2500)
    df = pd.read_excel(file_path, dtype=str, engine="openpyxl")
    total_rows = len(df)
    num_files = math.ceil(total_rows / rows_per_file)
    base_name = os.path.splitext(os.path.basename(file_path))[0]

    for i in range(num_files):
        start = i * rows_per_file
        end = min(start + rows_per_file, total_rows)
        part_df = df.iloc[start:end]
        out_name = f"{base_name}_V{i+1}.xlsx"
        part_df.to_excel(out_name, index=False)
        await update.message.reply_document(document=open(out_name, "rb"))
        os.remove(out_name)
        await asyncio.sleep(1)

    await update.message.reply_text("✅ All split files sent!")


# ================================
# 🔹 MERGE EXCEL
# ================================
async def process_merge(update: Update, context: ContextTypes.DEFAULT_TYPE):
    files = context.user_data.get("merge_files", [])
    if len(files) < 2:
        await update.message.reply_text(
            "⚠️ Please send at least 2 Excel files to merge.")
        return

    try:
        dfs = []
        for i, file_path in enumerate(files):
            df = pd.read_excel(file_path, dtype=str, engine="openpyxl")
            # Remove header row after the first file
            if i > 0:
                df = df.iloc[1:]
            dfs.append(df)

        merged_df = pd.concat(dfs, ignore_index=True)
        base_name = os.path.splitext(os.path.basename(files[0]))[0]
        out_name = f"{base_name}_merged.xlsx"

        merged_df.to_excel(out_name, index=False)
        await update.message.reply_document(document=open(out_name, "rb"))
        os.remove(out_name)
        await update.message.reply_text("✅ Merge complete!")

    except Exception as e:
        await update.message.reply_text(f"❌ Error during merge: {e}")
    finally:
        context.user_data["merge_files"] = []


# ================================
# 🔹 COMPARE EXCEL
# ================================
async def compare_excels(update: Update, context: ContextTypes.DEFAULT_TYPE):
    files = context.user_data["compare_files"]

    if len(files) != 2:
        await update.message.reply_text(
            "⚠️ Please upload exactly 2 Excel files for comparison.")
        return

    try:
        df1 = pd.read_excel(files[0], dtype=str, engine="openpyxl").fillna("")
        df2 = pd.read_excel(files[1], dtype=str, engine="openpyxl").fillna("")

        if df1.shape != df2.shape:
            await update.message.reply_text(
                f"❌ Files differ in structure:\n"
                f"File 1: {df1.shape[0]} rows × {df1.shape[1]} cols\n"
                f"File 2: {df2.shape[0]} rows × {df2.shape[1]} cols")
            return

        df1.columns = df1.columns.str.strip().str.lower()
        df2.columns = df2.columns.str.strip().str.lower()

        if sorted(df1.columns) != sorted(df2.columns):
            await update.message.reply_text(
                "❌ Columns are not identical across both files.")
            return
        df2 = df2[df1.columns]

        df1_sorted = df1.sort_values(by=list(df1.columns)).reset_index(
            drop=True)
        df2_sorted = df2.sort_values(by=list(df2.columns)).reset_index(
            drop=True)

        comparison = (df1_sorted == df2_sorted)

        if comparison.all().all():
            await update.message.reply_text(
                "✅ The two Excel files are *identical* in all data values!")
        else:
            diff_rows = (~comparison).any(axis=1).sum()
            await update.message.reply_text(
                f"⚠️ Files have inconsistencies in {diff_rows} rows.\n"
                "Generating difference file...")
            diff_df = pd.concat([df1_sorted, df2_sorted],
                                keys=["File1", "File2"]).reset_index()
            diff_df.to_excel("comparison_result.xlsx", index=False)
            await update.message.reply_document(
                document=open("comparison_result.xlsx", "rb"))
            os.remove("comparison_result.xlsx")

    except Exception as e:
        await update.message.reply_text(f"❌ Error during comparison: {e}")

    finally:
        context.user_data["compare_files"] = []


# ================================
# 🧹 CLEAR TEMP FILES
# ================================
async def clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    for f in os.listdir("uploads"):
        os.remove(os.path.join("uploads", f))
    await update.message.reply_text("🧹 Temporary files cleared.")


# ================================
# 🚀 START BOT
# ================================
app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("split", split_mode))
app.add_handler(CommandHandler("merge", merge_mode))
app.add_handler(CommandHandler("compare", compare_mode))
app.add_handler(CommandHandler("setrows", setrows))
app.add_handler(CommandHandler("clear", clear))
app.add_handler(
    MessageHandler(filters.Document.FileExtension("xlsx"), handle_file))

print("✅ Bot is running...")
keep_alive()
app.run_polling()
