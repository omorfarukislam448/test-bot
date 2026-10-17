import datetime
import hashlib
import hmac
import json
import os
import re
import time
from urllib.parse import urlencode
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import pytz
import requests
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

TOKEN = "8740758744:AAF551GI6y42aXKDWdvHrM04yanX4Wo0Uj0"
DEFAULT_ADMIN_ID = 6945991387

SETTINGS_FILE = "bot_settings.json"
USER_PAYMENTS_FILE = "user_payments.json"
GLOBAL_IDS_FILE = "global_submitted_ids.json"
SUBMISSION_LOGS_FILE = "submission_logs.json"
BLOCKED_USERS_FILE = "blocked_users.json"
ALL_USERS_FILE = "all_users.json"
ALL_SUBMITTED_ROWS_FILE = "all_submitted_rows_data.json"
REFERRALS_FILE = "referrals_data.json"
ADMINS_FILE = "admins.json"
USED_TXIDS_FILE = "used_crypto_txids.json"
CRYPTO_WITHDRAWALS_FILE = "crypto_withdrawals.json"

TZ = pytz.timezone("Asia/Dhaka")
MAINTENANCE_MODE = False
CATEGORY_SHIFTS = {
    "ig_2fa": {"start": "00:00", "end": "23:59"},
    "ig_cookies": {"start": "00:00", "end": "23:59"},
    "fb_2fa": {"start": "00:00", "end": "23:59"},
    "fb_cookies": {"start": "00:00", "end": "23:59"},
    "pc_clone": {"start": "00:00", "end": "23:59"},
}


def load_admins():
  if os.path.exists(ADMINS_FILE):
    with open(ADMINS_FILE, "r", encoding="utf-8") as f:
      try:
        return json.load(f)
      except:
        return [DEFAULT_ADMIN_ID]
  return [DEFAULT_ADMIN_ID]


def save_admins(admins_list):
  with open(ADMINS_FILE, "w", encoding="utf-8") as f:
    json.dump(admins_list, f, ensure_ascii=False, indent=4)


def is_admin(user_id):
  admins = load_admins()
  return user_id in admins or user_id == DEFAULT_ADMIN_ID


def is_shift_active(category_key):
  if category_key not in CATEGORY_SHIFTS:
    return True
  now = datetime.datetime.now(TZ)
  current_time_str = now.strftime("%H:%M")
  shift = CATEGORY_SHIFTS[category_key]
  return shift["start"] <= current_time_str <= shift["end"]


def load_used_txids():
  if os.path.exists(USED_TXIDS_FILE):
    with open(USED_TXIDS_FILE, "r", encoding="utf-8") as f:
      try:
        return json.load(f)
      except:
        return []
  return []


def save_used_txids(tx_list):
  with open(USED_TXIDS_FILE, "w", encoding="utf-8") as f:
    json.dump(tx_list, f, ensure_ascii=False, indent=4)


def load_crypto_withdrawals():
  if os.path.exists(CRYPTO_WITHDRAWALS_FILE):
    with open(CRYPTO_WITHDRAWALS_FILE, "r", encoding="utf-8") as f:
      try:
        return json.load(f)
      except:
        return {}
  return {}


def save_crypto_withdrawals(data):
  with open(CRYPTO_WITHDRAWALS_FILE, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)


def extract_crypto_amount(amount_str):
  nums = re.findall(r"[\d\.]+", str(amount_str))
  if nums:
    try:
      return float(nums[0])
    except:
      return 0.0
  return 0.0


def load_settings():
  default_data = {
      "channels": {
          "join_channel": "@OMORTECH",
          "submission_channel": "@OMORTECH",
          "withdrawal_channel": "@OMORTECH",
          "support_link": "https://t.me/OMOR_BUYER1",
      },
      "binance_api": {"api_key": "", "api_secret": ""},
      "custom_messages": {
          "welcome_msg": "👋 **স্বাগতম {name}!** আমাদের প্রিমিয়াম ফাইল সাবমিট সিস্টেমে আপনাকে স্বাগতম। 🌟\n\n📌 ফাইল সাবমিট করতে বা অন্যান্য সুবিধা পেতে নিচের বাটনগুলো ব্যবহার করুন:",
          "crypto_verify_msg": "🔍 **Binance API-এর মাধ্যমে ট্রানজেকশন যাচাই করা হচ্ছে...** প্রিয় {name}, দয়া করে অপেক্ষা করুন ⏳",
          "crypto_success_msg": "🎉 **অভিনন্দন {name}!** আপনার **{amount}** এর ট্রানজেকশনটি **Binance API**-এর মাধ্যমে সফলভাবে ভেরিফাই হয়েছে এবং উইথড্র রিকোয়েস্টটি অ্যাডমিন চ্যানেলে পাঠানো হয়েছে! 🚀",
          "crypto_duplicate_warning": "⚠️ **সতর্কতা {name}!** এই ট্রানজেকশন আইডি (প্রুফ) টি ইতিপূর্বেই ব্যবহার করা হয়েছে!\n\nএকই ট্রানজেকশন প্রুফ দিয়ে বারবার উইথড্র রিকোয়েস্ট পাঠানো কঠোরভাবে নিষিদ্ধ। সঠিক নতুন প্রুফ দিয়ে আবার চেষ্টা করুন।"
      },
      "categories": {
          "ig_2fa": {
              "name": "Instagram 2FA",
              "status": True,
              "price": 5.0,
              "show_price": True,
              "popup_closed": "দুঃখিত {name}, বর্তমানে {category} এর ফাইল জমা নেওয়া বন্ধ রয়েছে! ❌",
              "popup_shift": "দুঃখিত {name}, বর্তমানে {category} শিফটের বাইরে রয়েছে! ⏰",
          },
          "ig_cookies": {
              "name": "Instagram Cookies",
              "status": True,
              "price": 5.0,
              "show_price": True,
              "popup_closed": "দুঃখিত {name}, বর্তমানে {category} এর ফাইল জমা নেওয়া বন্ধ রয়েছে! ❌",
              "popup_shift": "দুঃখিত {name}, বর্তমানে {category} শিফটের বাইরে রয়েছে! ⏰",
          },
          "fb_2fa": {
              "name": "Facebook 2FA",
              "status": True,
              "price": 6.0,
              "show_price": True,
              "popup_closed": "দুঃখিত {name}, বর্তমানে {category} এর ফাইল জমা নেওয়া বন্ধ রয়েছে! ❌",
              "popup_shift": "দুঃখিত {name}, বর্তমানে {category} শিফটের বাইরে রয়েছে! ⏰",
          },
          "fb_cookies": {
              "name": "Facebook Cookies",
              "status": True,
              "price": 6.0,
              "show_price": True,
              "popup_closed": "দুঃখিত {name}, বর্তমানে {category} এর ফাইল জমা নেওয়া বন্ধ রয়েছে! ❌",
              "popup_shift": "দুঃখিত {name}, বর্তমানে {category} শিফটের বাইরে রয়েছে! ⏰",
          },
          "pc_clone": {
              "name": "PC Clone",
              "status": True,
              "price": 10.0,
              "show_price": True,
              "popup_closed": "দুঃখিত {name}, বর্তমানে {category} এর ফাইল জমা নেওয়া বন্ধ রয়েছে! ❌",
              "popup_shift": "দুঃখিত {name}, বর্তমানে {category} শিফটের বাইরে রয়েছে! ⏰",
          },
      },
      "payments": {
          "bkash": {"name": "Bkash", "status": True},
          "nagad": {"name": "Nagad", "status": True},
          "binance": {"name": "Binance", "status": True},
          "rocket": {"name": "Rocket", "status": True},
      },
      "wallets": {
          "Binance Pay": {
              "desc": "বিনান্স পে আইডি বা ইমেইল",
              "address": "omor.tech.binance@gmail.com",
          },
          "USDT (TRC20)": {
              "desc": "ট্রন নেটওয়ার্ক ইউএসডিটি",
              "address": "TQn9Y2khEsLJW1ChVWFMSMeRDow5KcbLSE",
          },
      },
      "referral": {
          "status": True,
          "bonus_amount": 10.0,
          "commission_percent": 5.0,
          "commission_status": True,
          "min_refs_to_withdraw": 10,
          "min_amount_to_withdraw": 50.0,
      },
  }
  if os.path.exists(SETTINGS_FILE):
    with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
      try:
        data = json.load(f)
        for cat_key, cat_val in default_data["categories"].items():
          if cat_key not in data.get("categories", {}):
            data.setdefault("categories", {})[cat_key] = cat_val
          else:
            for k, v in cat_val.items():
              if k not in data["categories"][cat_key]:
                data["categories"][cat_key][k] = v
        if "channels" not in data:
          data["channels"] = default_data["channels"]
        if "binance_api" not in data:
          data["binance_api"] = default_data["binance_api"]
        if "custom_messages" not in data:
          data["custom_messages"] = default_data["custom_messages"]
        if "referral" not in data:
          data["referral"] = default_data["referral"]
        if "wallets" not in data:
          data["wallets"] = default_data["wallets"]
        return data
      except:
        return default_data
  return default_data


def save_settings(data):
  with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)


def load_user_payments():
  if os.path.exists(USER_PAYMENTS_FILE):
    with open(USER_PAYMENTS_FILE, "r", encoding="utf-8") as f:
      try:
        return json.load(f)
      except:
        return {}
  return {}


def save_user_payments(data):
  with open(USER_PAYMENTS_FILE, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)


def load_global_ids():
  if os.path.exists(GLOBAL_IDS_FILE):
    with open(GLOBAL_IDS_FILE, "r", encoding="utf-8") as f:
      try:
        return json.load(f)
      except:
        return []
  return []


def save_global_ids(ids_list):
  with open(GLOBAL_IDS_FILE, "w", encoding="utf-8") as f:
    json.dump(ids_list, f, ensure_ascii=False, indent=4)


def load_logs():
  if os.path.exists(SUBMISSION_LOGS_FILE):
    with open(SUBMISSION_LOGS_FILE, "r", encoding="utf-8") as f:
      try:
        return json.load(f)
      except:
        return []
  return []


def save_logs(logs_list):
  with open(SUBMISSION_LOGS_FILE, "w", encoding="utf-8") as f:
    json.dump(logs_list, f, ensure_ascii=False, indent=4)


def load_blocked_users():
  if os.path.exists(BLOCKED_USERS_FILE):
    with open(BLOCKED_USERS_FILE, "r", encoding="utf-8") as f:
      try:
        return json.load(f)
      except:
        return []
  return []


def save_blocked_users(blocked_list):
  with open(BLOCKED_USERS_FILE, "w", encoding="utf-8") as f:
    json.dump(blocked_list, f, ensure_ascii=False, indent=4)


def load_all_users():
  if os.path.exists(ALL_USERS_FILE):
    with open(ALL_USERS_FILE, "r", encoding="utf-8") as f:
      try:
        return json.load(f)
      except:
        return {}
  return {}


def save_all_users(data):
  with open(ALL_USERS_FILE, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)


def load_all_rows_db():
  if os.path.exists(ALL_SUBMITTED_ROWS_FILE):
    with open(ALL_SUBMITTED_ROWS_FILE, "r", encoding="utf-8") as f:
      try:
        return json.load(f)
      except:
        return {}
  return {}


def save_all_rows_db(data):
  with open(ALL_SUBMITTED_ROWS_FILE, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)


def load_referrals():
  if os.path.exists(REFERRALS_FILE):
    with open(REFERRALS_FILE, "r", encoding="utf-8") as f:
      try:
        return json.load(f)
      except:
        return {}
  return {}


def save_referrals(data):
  with open(REFERRALS_FILE, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)


def verify_binance_deposit(tx_id, api_key, api_secret):
  if not api_key or not api_secret:
    return False, "API Keys not set"
  url = "https://api.binance.com/sapi/v1/capital/deposit/hisrec"
  timestamp = int(time.time() * 1000) - 2000
  params = {"timestamp": timestamp, "recvWindow": 60000}
  query_string = urlencode(params)
  signature = hmac.new(
      api_secret.encode("utf-8"),
      query_string.encode("utf-8"),
      hashlib.sha256,
  ).hexdigest()
  headers = {"X-MBX-APIKEY": api_key}
  try:
    response = requests.get(
        f"{url}?{query_string}&signature={signature}",
        headers=headers,
        timeout=10,
    )
    data = response.json()
    if isinstance(data, list):
      for item in data:
        if str(item.get("txId", "")).strip().lower() == str(tx_id).strip().lower():
          if item.get("status") == 1:
            amt = item.get("amount", "0")
            coin = item.get("coin", "USDT")
            return True, {
                "amount": f"{amt} {coin}",
                "raw": item,
                "type": "Deposit",
            }
      return False, "TxHash not found or not completed in Binance deposit history"
    else:
      return False, f"API Error: {data.get('msg', 'Unknown error')}"
  except Exception as e:
    return False, str(e)


def verify_binance_pay(pay_id_or_order_id, api_key, api_secret):
  if not api_key or not api_secret:
    return False, "API Keys not set"
  url = "https://api.binance.com/sapi/v1/pay/transactions"
  timestamp = int(time.time() * 1000) - 2000
  params = {"timestamp": timestamp, "recvWindow": 60000}
  query_string = urlencode(params)
  signature = hmac.new(
      api_secret.encode("utf-8"),
      query_string.encode("utf-8"),
      hashlib.sha256,
  ).hexdigest()
  headers = {"X-MBX-APIKEY": api_key}
  try:
    response = requests.get(
        f"{url}?{query_string}&signature={signature}",
        headers=headers,
        timeout=10,
    )
    data = response.json()
    if data.get("code") == "000000" or data.get("success") == True:
      transactions = data.get("data", [])
      for tx in transactions:
        if (
            str(tx.get("orderId", "")).strip()
            == str(pay_id_or_order_id).strip()
            or str(tx.get("transactionId", "")).strip()
            == str(pay_id_or_order_id).strip()
        ):
          amt = tx.get("amount", "0")
          currency = tx.get("currency", "USDT")
          return True, {
              "amount": f"{amt} {currency}",
              "raw": tx,
              "type": "Binance Pay",
          }
      return (
          False,
          "Pay ID / Order ID not found in Binance Pay transaction history",
      )
    else:
      return False, f"API Error: {data.get('message', 'Unknown error')}"
  except Exception as e:
    return False, str(e)


def format_channel_link(ch_val):
  ch_val = str(ch_val).strip()
  if ch_val.startswith("@"):
    username = ch_val[1:]
    return f"https://t.me/{username}", f"@{username}"
  elif "t.me/" in ch_val:
    if ch_val.startswith("http://") or ch_val.startswith("https://"):
      return ch_val, ch_val
    return f"https://t.me/{ch_val}", f"https://t.me/{ch_val}"
  else:
    return f"https://t.me/{ch_val}", ch_val


def parse_chat_id(ch_val):
  ch_str = str(ch_val).strip()
  if ch_str.isdigit() or (ch_str.startswith("-") and ch_str[1:].isdigit()):
    return int(ch_str)
  return ch_str


async def check_channel_membership(bot, user_id):
  if is_admin(user_id):
    return True
  settings = load_settings()
  ch = settings.get("channels", {}).get("join_channel", "@OMORTECH")
  target_chat = parse_chat_id(ch)
  try:
    member = await bot.get_chat_member(chat_id=target_chat, user_id=user_id)
    if member.status in ["creator", "administrator", "member"]:
      return True
  except Exception:
    pass
  return False


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
  user = update.effective_user
  bot = context.bot

  blocked_users = load_blocked_users()
  if user.id in blocked_users:
    msg = f"⛔ **দুঃখিত {user.first_name}!** আপনাকে এই বট ব্যবহার করা থেকে ব্লক করা হয়েছে।"
    if update.callback_query:
      await update.callback_query.message.reply_text(msg, parse_mode="Markdown")
    else:
      await update.message.reply_text(msg, parse_mode="Markdown")
    return

  if MAINTENANCE_MODE and not is_admin(user.id):
    msg = (
        f"দুঃখিত {user.first_name}, বটটি বর্তমানে রক্ষণাবেক্ষণ (Maintenance) মোডে রয়েছে। অনুগ্রহ"
        " করে পরে চেষ্টা করুন।"
    )
    if update.callback_query:
      await update.callback_query.message.reply_text(msg)
    else:
      await update.message.reply_text(msg)
    return

  settings = load_settings()
  join_ch = settings.get("channels", {}).get("join_channel", "@OMORTECH")
  link_url, link_display = format_channel_link(join_ch)

  is_member = await check_channel_membership(bot, user.id)
  if not is_member:
    join_msg = (
        f"⛔ **সম্মানিত {user.first_name}, চ্যানেলে জয়েন করা বাধ্যতামূলক!**\n\n"
        "বটটি ব্যবহার করতে হলে প্রথমে আমাদের অফিসিয়াল চ্যানেলে জয়েন করতে"
        " হবে। নিচের বাটনে ক্লিক করে জয়েন করুন:\n\n"
        f"✨ চ্যানেলে জয়েন করার পর নিচের **🔄 স্টার্ট / রিফ্রেশ করুন** বাটনে"
        " ক্লিক করুন."
    )
    join_kbd = [
        [InlineKeyboardButton("📢 চ্যানেলে জয়েন করুন", url=link_url)],
        [
            InlineKeyboardButton(
                "🔄 স্টার্ট / রিফ্রেশ করুন", callback_data="back_home"
            )
        ],
    ]
    if update.callback_query:
      try:
        await update.callback_query.message.edit_text(
            join_msg,
            reply_markup=InlineKeyboardMarkup(join_kbd),
            parse_mode="Markdown",
            disable_web_page_preview=True,
        )
      except:
        await update.callback_query.message.reply_text(
            join_msg,
            reply_markup=InlineKeyboardMarkup(join_kbd),
            parse_mode="Markdown",
            disable_web_page_preview=True,
        )
    else:
      await update.message.reply_text(
          join_msg,
          reply_markup=InlineKeyboardMarkup(join_kbd),
          parse_mode="Markdown",
          disable_web_page_preview=True,
      )
    return

  all_users = load_all_users()
  user_id_str = str(user.id)
  if user_id_str not in all_users:
    all_users[user_id_str] = user.first_name
    save_all_users(all_users)

    if context.args:
      arg = context.args[0]
      if arg.startswith("ref_"):
        try:
          ref_by_id = int(arg.replace("ref_", ""))
          if ref_by_id != user.id:
            referrals = load_referrals()
            if user_id_str not in referrals.get("referred_users", []):
              referrals.setdefault("referred_users", []).append(user_id_str)

              bonus_now = settings.get("referral", {}).get("bonus_amount", 10.0)

              ref_by_str = str(ref_by_id)
              if ref_by_str not in referrals.get("referrers", {}):
                referrals.setdefault("referrers", {})[ref_by_str] = {
                    "refs": [],
                    "withdrawn": 0.0,
                    "commission_earned": 0.0,
                    "manual_adjustment": 0.0,
                }

              referrals["referrers"][ref_by_str]["refs"].append({
                  "uid": user_id_str,
                  "bonus": bonus_now,
                  "date": datetime.datetime.now(TZ).strftime("%Y-%m-%d"),
              })
              save_referrals(referrals)
        except:
          pass

  welcome_tpl = settings.get("custom_messages", {}).get(
      "welcome_msg", "👋 **স্বাগতম {name}!**"
  )
  welcome_msg = welcome_tpl.replace("{name}", user.first_name)

  support_link_val = (
      settings.get("channels", {})
      .get("support_link", "https://t.me/OMOR_BUYER1")
      .strip()
  )

  keyboard = []
  for key, cat in settings["categories"].items():
    if cat["status"] and is_shift_active(key):
      btn_text = f"🟢 {cat['name']}"
    else:
      btn_text = f"🔴 {cat['name']}"
    keyboard.append([InlineKeyboardButton(btn_text, callback_data=f"cat_{key}")])

  if support_link_val.lower() == "bot":
    keyboard.append([
        InlineKeyboardButton(
            "📞 সাপোর্ট আইডি", callback_data="user_bot_support"
        )
    ])
  else:
    sup_url, sup_disp = format_channel_link(support_link_val)
    keyboard.append([InlineKeyboardButton("📞 সাপোর্ট আইডি", url=sup_url)])

  keyboard.append([
      InlineKeyboardButton("⚙ পেমেন্ট সেটিংস", callback_data="payment_menu"),
      InlineKeyboardButton(
          "💳 আমার পেমেন্ট ও হিস্ট্রি", callback_data="my_payment_tracking"
      ),
  ])
  keyboard.append([
      InlineKeyboardButton(
          "📁 আমার হিস্ট্রি এক্সেল ফাইল", callback_data="export_my_history"
      ),
      InlineKeyboardButton(
          "🎁 রেফারেল লিংক ও ব্যালেন্স", callback_data="referral_menu_user"
      ),
  ])
  keyboard.append([
      InlineKeyboardButton(
          "🌐 ক্রিপ্টো উইথড্র অ্যাড্রেস", callback_data="user_crypto_wallets"
      )
  ])

  if is_admin(user.id):
    keyboard.append([
        InlineKeyboardButton(
            "🛠 অ্যাডমিন প্যানেল সেটিংস", callback_data="admin_panel_menu"
        )
    ])

  reply_markup = InlineKeyboardMarkup(keyboard)

  if update.callback_query:
    try:
      await update.callback_query.message.edit_text(
          welcome_msg,
          reply_markup=reply_markup,
          parse_mode="Markdown",
          disable_web_page_preview=True,
      )
    except:
      await update.callback_query.message.reply_text(
          welcome_msg,
          reply_markup=reply_markup,
          parse_mode="Markdown",
          disable_web_page_preview=True,
      )
  else:
    await update.message.reply_text(
        welcome_msg,
        reply_markup=reply_markup,
        parse_mode="Markdown",
        disable_web_page_preview=True,
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
  user = update.effective_user
  await update.message.reply_text(
      f"💡 **প্রিয় {user.first_name}, সাহায্য ও নির্দেশিকা:**\n\n1. প্রথমে বট মেনু থেকে **⚙ পেমেন্ট সেটিংস** এ"
      " গিয়ে আপনার পেমেন্ট নম্বর সেট করুন।\n2. পছন্দের ক্যাটাগরিতে ক্লিক করে"
      " সঠিক ফরম্যাটের এক্সেল ফাইল আপলোড করুন।\n3. কোনো সমস্যা হলে সাপোর্ট আইডিতে"
      " যোগাযোগ করুন।",
      parse_mode="Markdown",
  )


async def status_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
  if not is_admin(update.effective_user.id):
    return
  logs = load_logs()
  users = load_all_users()
  await update.message.reply_text(
      f"📊 **বট স্ট্যাটাস রিপোর্ট:**\n\n• মোট রেজিস্টার্ড ইউজার: `{len(users)} জন`\n•"
      f" মোট সাবমিশন লগ: `{len(logs)} টি`\n• মেইনটেনেন্স মোড:"
      f" `{'ON' if MAINTENANCE_MODE else 'OFF'}`",
      parse_mode="Markdown",
  )


async def users_list_command(
    update: Update, context: ContextTypes.DEFAULT_TYPE
):
  if not is_admin(update.effective_user.id):
    return
  users = load_all_users()
  text = f"👥 **বট ব্যবহারকারী সকল মেম্বারদের তালিকা (মোট: {len(users)} জন):**\n\n"
  for uid, name in list(users.items()):
    text += (
        f"👤 [{name}](tg://user?id={uid})\n"
        f"🆔 User ID: `{uid}`\n"
        "-------------------------\n"
    )
  await update.message.reply_text(text, parse_mode="Markdown")


async def reset_all_logs_command(
    update: Update, context: ContextTypes.DEFAULT_TYPE
):
  if not is_admin(update.effective_user.id):
    return
  save_logs([])
  save_all_rows_db({})
  await update.message.reply_text(
      "✅ সকল সাবমিশন হিস্ট্রি এবং মূল এক্সেল ডাটাবেস সম্পূর্ণ রিসেট করা হয়েছে!"
  )


async def clear_global_ids_command(
    update: Update, context: ContextTypes.DEFAULT_TYPE
):
  if not is_admin(update.effective_user.id):
    return
  save_global_ids([])
  save_all_rows_db({})
  await update.message.reply_text(
      "✅ গ্লোবাল সাবমিটেড আইডি এবং মূল এক্সেল ডাটাবেস সম্পূর্ণ ক্লিয়ার করা হয়েছে!"
  )


async def top_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
  logs = load_logs()
  user_counts = {}
  for l in logs:
    uid = l["user_id"]
    name = l["user_name"]
    cnt = l.get("unique", 0)
    user_counts[uid] = user_counts.get(uid, {"name": name, "total": 0})
    user_counts[uid]["total"] += cnt

  sorted_users = sorted(
      user_counts.items(), key=lambda x: x[1]["total"], reverse=True
  )[:10]
  text = "🏆 **টপ সাবমিটর লিডারবোর্ড:**\n\n"
  for rank, (uid, info) in enumerate(sorted_users, 1):
    text += (
        f"{rank}. [{info['name']}](tg://user?id={uid}) —"
        f" `{info['total']} টি ভ্যালিড আইডি`\n"
    )
  await update.message.reply_text(text, parse_mode="Markdown")


async def block_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
  if not is_admin(update.effective_user.id):
    return
  if not context.args:
    await update.message.reply_text("ব্যবহার: `/block <user_id>`")
    return
  try:
    uid = int(context.args[0])
    blocked = load_blocked_users()
    if uid not in blocked:
      blocked.append(uid)
      save_blocked_users(blocked)
    await update.message.reply_text(f"⛔ ইউজার `{uid}` কে ব্লক করা হয়েছে।")
  except:
    await update.message.reply_text("সঠিক ইউজার আইডি দিন।")


async def unblock_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
  if not is_admin(update.effective_user.id):
    return
  if not context.args:
    await update.message.reply_text("ব্যবহার: `/unblock <user_id>`")
    return
  try:
    uid = int(context.args[0])
    blocked = load_blocked_users()
    if uid in blocked:
      blocked.remove(uid)
      save_blocked_users(blocked)
    await update.message.reply_text(f"✅ ইউজার `{uid}` কে আনব্লক করা হয়েছে।")
  except:
    await update.message.reply_text("সঠিক ইউজার আইডি দিন।")


async def broadcast_command(
    update: Update, context: ContextTypes.DEFAULT_TYPE
):
  if not is_admin(update.effective_user.id):
    return
  if not context.args:
    await update.message.reply_text(
        "ব্যবহার: `/broadcast <আপনার বার্তা>`", parse_mode="Markdown"
    )
    return
  msg = " ".join(context.args)
  users = load_all_users()
  success = 0
  for uid in users:
    try:
      await context.bot.send_message(
          chat_id=int(uid),
          text=f"📢 **ব্রডকাস্ট নোটিফিকেশন:**\n\n{msg}",
          parse_mode="Markdown",
      )
      success += 1
    except:
      pass
  await update.message.reply_text(
      f"✅ সফলভাবে {success} জন ইউজারের কাছে বার্তা পাঠানো হয়েছে!"
  )


async def send_specific_message(
    update: Update, context: ContextTypes.DEFAULT_TYPE
):
  if not is_admin(update.effective_user.id):
    return
  if len(context.args) < 2:
    await update.message.reply_text("ব্যবহার: `/send <user_id> <বার্তা>`")
    return
  try:
    uid = int(context.args[0])
    msg = " ".join(context.args[1:])
    await context.bot.send_message(chat_id=uid, text=msg, parse_mode="Markdown")
    await update.message.reply_text("✅ বার্তা পাঠানো হয়েছে!")
  except Exception as e:
    await update.message.reply_text(f"ত্রুটি: {e}")


async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
  global MAINTENANCE_MODE
  query = update.callback_query
  data = query.data
  user = update.effective_user
  user_id = user.id

  waiting_keys = [
      "waiting_for_popup_closed",
      "waiting_for_popup_shift",
      "waiting_for_ref_bonus",
      "waiting_for_min_refs",
      "waiting_for_min_amt",
      "waiting_for_search_user",
      "waiting_for_shift_cat",
      "waiting_for_report_info",
      "waiting_for_custom_date_cat",
      "waiting_for_new_price_cat",
      "waiting_for_payment_method",
      "waiting_for_cat_popup_edit",
      "waiting_for_add_wallet_name_step",
      "waiting_for_add_wallet_desc_step",
      "waiting_for_add_wallet_addr_step",
      "waiting_for_user_edit_data",
      "waiting_for_user_reset_id",
      "waiting_for_new_cat_key",
      "waiting_for_new_cat_name",
      "waiting_for_new_cat_price",
      "waiting_for_comm_percent",
      "waiting_for_join_channel",
      "waiting_for_sub_channel",
      "waiting_for_with_channel",
      "waiting_for_support_link",
      "waiting_for_new_admin_id",
      "waiting_for_support_msg",
      "waiting_for_binance_api_key",
      "waiting_for_binance_api_secret",
      "waiting_for_edit_msg_key",
  ]
  for k in waiting_keys:
    context.user_data.pop(k, None)

  if user_id in load_blocked_users():
    await query.answer(f"⛔ {user.first_name}, আপনাকে ব্লক করা হয়েছে!", show_alert=True)
    return

  if MAINTENANCE_MODE and not is_admin(user_id):
    await query.answer(
        f"দুঃখিত {user.first_name}, বটটি বর্তমানে রক্ষণাবেক্ষণ মোডে রয়েছে।",
        show_alert=True,
    )
    return

  is_member = await check_channel_membership(context.bot, user_id)
  settings = load_settings()
  join_ch = settings.get("channels", {}).get("join_channel", "@OMORTECH")
  link_url, link_disp = format_channel_link(join_ch)

  if not is_member and data != "back_home":
    await query.answer(
        f"⛔ সম্মানিত {user.first_name}, প্রথমে আমাদের চ্যানেলে জয়েন করুন: {link_disp}", show_alert=True
    )
    return

  user_id_str = str(user_id)
  user_payments = load_user_payments()

  if data == "user_bot_support":
    await query.answer()
    context.user_data["waiting_for_support_msg"] = True
    kbd = [[InlineKeyboardButton("🔙 মূল মেনু", callback_data="back_home")]]
    await query.message.edit_text(
        f"📞 **সাপোর্ট সেন্টার - সম্মানিত {user.first_name}**\n\n"
        "আপনার কি সমস্যা হচ্ছে তা অনুগ্রহ করে বিস্তারিত লিখে পাঠান। "
        "আপনি চাইলে **মেসেজ, স্ক্রিনশট ছবি অথবা উভয়টি** একসাথে পাঠাতে পারেন। "
        "আপনার মেসেজটি সরাসরি আমাদের অ্যাডমিন টিমের কাছে পৌঁছে যাবে। 📩",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("cat_"):
    cat_key = data.replace("cat_", "")
    if cat_key not in settings["categories"]:
      await query.answer("⛔ ক্যাটাগরি পাওয়া যায়নি!", show_alert=True)
      return

    cat_info = settings["categories"][cat_key]
    status_open = cat_info["status"]
    shift_open = is_shift_active(cat_key)

    if (not status_open or not shift_open) and not is_admin(user_id):
      if not status_open:
        tpl = cat_info.get(
            "popup_closed",
            "দুঃখিত {name}, বর্তমানে {category} এর ফাইল জমা নেওয়া বন্ধ রয়েছে! ❌",
        )
      else:
        tpl = cat_info.get(
            "popup_shift",
            "দুঃখিত {name}, বর্তমানে {category} শিফটের বাইরে রয়েছে! ⏰",
        )
      custom_msg = tpl.replace("{category}", cat_info["name"]).replace("{name}", user.first_name)
      await query.answer(custom_msg, show_alert=True)
      return

    await query.answer()
    context.user_data["selected_category_key"] = cat_key
    context.user_data["selected_category"] = cat_info["name"]
    context.user_data["selected_category_price"] = cat_info["price"]

    if any(
        x in cat_key for x in ["mail", "gmail", "yahoo", "outlook", "hotmail"]
    ):
      col_a_desc = "কলাম A: Email Address"
    elif "ig" in cat_key:
      col_a_desc = "কলাম A: Username"
    else:
      col_a_desc = "কলাম A: UID"

    col_c_desc = (
        "কলাম C: Cookies"
        if ("cookies" in cat_key or cat_key == "pc_clone")
        else "কলাম C: 2FA Key"
    )

    if cat_info.get("show_price", True):
      rate_text = f"\n💰 ক্যাটাগরি রেট: **{cat_info['price']}৳ প্রতি আইডি**"
    else:
      rate_text = ""

    keyboard = [
        [InlineKeyboardButton("🔙 মূল মেনুতে ফিরে যান", callback_data="back_home")]
    ]
    await query.message.edit_text(
        f"📂 ক্যাটাগরি: **{cat_info['name']}**{rate_text}\n\n📌 **সম্মানিত {user.first_name}, সঠিক নিয়মাবলী:**\n• {col_a_desc}\n• কলাম B: Password\n• {col_c_desc}\n\n📁"
        " এখন আপনার এক্সেল ফাইলটি (.xlsx) সরাসরি এই বটে সেন্ড করুন। 📥",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown",
    )
    return

  elif data == "user_crypto_wallets":
    await query.answer()
    wallets = settings.get("wallets", {})
    text = (
        f"🌐 **ক্রিপ্টো ও ওয়ালেট উইথড্র অ্যাড্রেসসমূহ ({user.first_name}):**\n\nনিচের মাধ্যম বা"
        " ওয়ালেটগুলো থেকে আপনার পছন্দমতো মাধ্যমে উইথড্র করতে পারেন:\n\n"
    )
    kbd = []
    for w_name, w_info in wallets.items():
      if isinstance(w_info, str):
        w_desc = ""
      else:
        w_desc = w_info.get("desc", "")
      kbd.append([
          InlineKeyboardButton(
              f"🔹 {w_name} ({w_desc})", callback_data=f"view_w_{w_name}"
          )
      ])
    kbd.append([InlineKeyboardButton("🔙 মূল মেনু", callback_data="back_home")])
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
    return

  elif data.startswith("view_w_"):
    w_name = data.replace("view_w_", "")
    wallets = settings.get("wallets", {})
    w_info = wallets.get(w_name, {})

    if isinstance(w_info, str):
      w_addr = w_info
      w_desc = ""
    else:
      w_addr = w_info.get("address", "N/A")
      w_desc = w_info.get("desc", "")

    await query.answer()
    text = (
        f"🌐 **ওয়ালেট মাধ্যম:** {w_name}\n"
        f"📝 বিবরণ: {w_desc}\n\n"
        f"📌 **অ্যাড্রেস / অ্যাকাউন্ট:**\n`{w_addr}`\n\n"
        f"💡 *প্রিয় {user.first_name}, এই অ্যাড্রেসে পেমেন্ট করে ট্রানজেকশন আইডি (TxHash বা Pay ID / Order ID) সরাসরি এই বটে সেন্ড করুন। আমাদের সিস্টেম স্বয়ংক্রিয়ভাবে Binance API দিয়ে যাচাই করে উইথড্র এপ্রুভ করবে।*"
    )
    kbd = [
        [
            InlineKeyboardButton(
                "📋 অ্যাড্রেস কপি করুন", callback_data=f"copy_w_addr_{w_name}"
            )
        ],
        [
            InlineKeyboardButton(
                "📤 ট্রানজেকশন আইডি বা প্রুফ পাঠান",
                callback_data=f"send_crypto_proof_{w_name}",
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 ওয়ালেট লিস্ট", callback_data="user_crypto_wallets"
            )
        ],
    ]
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
    return

  elif data.startswith("copy_w_addr_"):
    w_name = data.replace("copy_w_addr_", "")
    wallets = settings.get("wallets", {})
    w_info = wallets.get(w_name, {})
    if isinstance(w_info, str):
      addr = w_info
    else:
      addr = w_info.get("address", "N/A")
    await query.answer(f"অ্যাড্রেস: {addr}", show_alert=True)
    return

  elif data.startswith("send_crypto_proof_"):
    user_p = user_payments.get(user_id_str, {})
    has_valid_payment = any(
        p_v != "সেট করা হয়নি" for p_k, p_v in user_p.items()
    )
    if not has_valid_payment and not is_admin(user_id):
      await query.answer(
          f"⛔ {user.first_name}, উইথড্র করার পূর্বে অন্তত একটি পেমেন্ট নম্বর বা ওয়ালেট তথ্য সেট করা বাধ্যতামূলক!",
          show_alert=True,
      )
      return

    w_name = data.replace("send_crypto_proof_", "")
    await query.answer()
    context.user_data["waiting_for_crypto_proof"] = w_name
    kbd = [[
        InlineKeyboardButton(
            "🔙 ওয়ালেট লিস্ট", callback_data="user_crypto_wallets"
        )
    ]]
    await query.message.edit_text(
        f"✍ প্রিয় **{user.first_name}**, আপনি **{w_name}** মাধ্যমে উইথড্র রিকোয়েস্ট পাঠাতে চাচ্ছেন。\n\n"
        "দয়া করে এখন আপনার পেমেন্ট সম্পন্ন করার পর **TxHash / Txid** অথবা **Binance Pay ID / Order ID** টেক্সট আকারে এই চ্যাটে সেন্ড করুন (বট স্বয়ংক্রিয়ভাবে Binance API দিয়ে চেক করবে):",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "admin_panel_menu":
    if not is_admin(user_id):
      await query.answer("⛔ আপনার এই অপশনে প্রবেশাধিকার নেই!", show_alert=True)
      return
    await query.answer()

    m_status_text = "🟢 ON" if MAINTENANCE_MODE else "🔴 OFF"
    admin_keyboard = [
        [
            InlineKeyboardButton(
                f"⚠️ মেইনটেনেন্স মোড: {m_status_text}",
                callback_data="adm_toggle_maintenance",
            )
        ],
        [
            InlineKeyboardButton(
                "💬 বটের সমস্ত মেসেজ এডিট প্যানেল", callback_data="adm_custom_messages_menu"
            )
        ],
        [
            InlineKeyboardButton(
                "🔑 Binance API সেটিংস", callback_data="adm_binance_api_menu"
            )
        ],
        [
            InlineKeyboardButton(
                "📂 ক্যাটাগরি ম্যানেজমেন্ট ও পপআপ",
                callback_data="adm_cat_mgmt_hub",
            )
        ],
        [
            InlineKeyboardButton(
                "📥 ফাইল ডাউনলোড হাব (Download Hub)",
                callback_data="adm_download_hub",
            )
        ],
        [
            InlineKeyboardButton(
                "💰 পেমেন্ট ও রিপোর্ট ম্যাচিং প্যানেল",
                callback_data="adm_pay_report_hub",
            )
        ],
        [
            InlineKeyboardButton(
                "👤 ইউজার ব্যালেন্স/রেফার ও ডাটা কন্ট্রোল",
                callback_data="adm_user_manage_menu",
            )
        ],
        [
            InlineKeyboardButton(
                "🌐 ক্রিপ্টো উইথড্র হিস্ট্রি ও পরিসংখ্যান",
                callback_data="adm_crypto_stats",
            )
        ],
        [
            InlineKeyboardButton(
                "🌐 ক্রিপ্টো ওয়ালেট অ্যাড্রেস ম্যানেজ",
                callback_data="adm_wallet_mgmt",
            )
        ],
        [
            InlineKeyboardButton(
                "📢 চ্যানেল ও সাপোর্ট রাউটিং সেটআপ",
                callback_data="adm_channel_mgmt",
            )
        ],
        [
            InlineKeyboardButton(
                "👑 অ্যাডমিন কন্ট্রোল ও রিমুভ প্যানেল",
                callback_data="adm_admin_control_panel",
            )
        ],
        [
            InlineKeyboardButton(
                "⏰ শিফট টাইম সেটআপ", callback_data="adm_shift_menu"
            )
        ],
        [
            InlineKeyboardButton(
                "🎁 রেফারেল সিস্টেম ও কমিশন",
                callback_data="adm_referral_settings",
            )
        ],
        [
            InlineKeyboardButton(
                "🎛 পেমেন্ট মাধ্যম অন/অফ", callback_data="adm_toggle_payments"
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 মূল মেনুতে ফিরে যান", callback_data="back_home"
            )
        ],
    ]
    await query.message.edit_text(
        "🛠 **অ্যাডমিন কন্ট্রোল প্যানেল (মাস্টার হাব)**\n\nনিচের অপশনগুলো থেকে আপনার"
        " বট সম্পূর্ণ নিয়ন্ত্রণ করুন:",
        reply_markup=InlineKeyboardMarkup(admin_keyboard),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_custom_messages_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    msgs = settings.get("custom_messages", {})
    text = (
        "💬 **বটের সমস্ত মেসেজ কাস্টমাইজেশন হাব**\n\n"
        "বটের যেকোনো মেসেজ পছন্দমত এডিট করতে নিচের বাটনগুলোতে ক্লিক করুন:\n"
        "*(মেসেজে `{name}` বা `{amount}` ডায়নামিক ভ্যারিয়েবল ব্যবহার করতে পারেন)*"
    )
    kbd = [
        [InlineKeyboardButton("✏️ ওয়েলকাম মেসেজ", callback_data="msg_edit_welcome_msg")],
        [InlineKeyboardButton("✏️ ক্রিপ্টো ভেরিফাই প্রোগ্রেস মেসেজ", callback_data="msg_edit_crypto_verify_msg")],
        [InlineKeyboardButton("✏️ ক্রিপ্টো ভেরিফাই সাকসেস মেসেজ", callback_data="msg_edit_crypto_success_msg")],
        [InlineKeyboardButton("✏️ ডুপ্লিকেট প্রুফ ওয়ার্নিং মেসেজ", callback_data="msg_edit_crypto_duplicate_warning")],
        [InlineKeyboardButton("🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu")],
    ]
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
    return

  elif data.startswith("msg_edit_"):
    if not is_admin(user_id):
      return
    await query.answer()
    msg_key = data.replace("msg_edit_", "")
    context.user_data["waiting_for_edit_msg_key"] = msg_key
    curr_text = settings.get("custom_messages", {}).get(msg_key, "N/A")
    kbd = [[InlineKeyboardButton("🔙 ফিরে যান", callback_data="adm_custom_messages_menu")]]
    await query.message.edit_text(
        f"✍ **বর্তমানে সেট করা মেসেজ:**\n`{curr_text}`\n\n"
        "নতুন মেসেজটি নিচে লিখে চ্যাটে সেন্ড করুন (ইউজারের নাম দেখাতে `{name}` ও টাকার পরিমাণের জন্য `{amount}` ব্যবহার করুন):",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_crypto_stats":
    if not is_admin(user_id):
      return
    await query.answer()
    crypto_w_data = load_crypto_withdrawals()
    if not crypto_w_data:
      await query.answer("⛔ কোনো ক্রিপ্টো উইথড্র হিস্ট্রি পাওয়া যায়নি!", show_alert=True)
      return

    all_users = load_all_users()
    text = "🌐 **ক্রিপ্টো ওয়ালেট উইথড্র পরিসংখ্যান (আজ পর্যন্ত):**\n\n"
    grand_total_usdt = 0.0

    for uid_str, w_list in crypto_w_data.items():
      u_name = all_users.get(uid_str, "Unknown User")
      user_total = 0.0
      text += f"👤 [{u_name}](tg://user?id={uid_str}) (`{uid_str}`)\n"
      for w_item in w_list:
        amt_str = w_item.get("amount_str", "0")
        val = extract_crypto_amount(amt_str)
        user_total += val
        grand_total_usdt += val
        text += f"   • {w_item.get('wallet', 'Crypto')}: `{amt_str}` ({w_item.get('date', '')})\n"
      text += f"   📌 **মোট উত্তোলন:** `{user_total:.2f} USDT`\n-------------------------\n"

    text += f"\n💎 **সর্বমোট ক্রিপ্টো উইথড্র:** `{grand_total_usdt:.2f} USDT`"
    kbd = [[InlineKeyboardButton("🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu")]]
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown", disable_web_page_preview=True
    )
    return

  elif data == "adm_binance_api_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    b_api = settings.get("binance_api", {})
    has_key = "✅ সেট করা আছে" if b_api.get("api_key") else "❌ সেট করা নেই"
    has_secret = (
        "✅ সেট করা আছে" if b_api.get("api_secret") else "❌ সেট করা নেই"
    )

    text = (
        "🔑 **Binance API Configuration**\n\n"
        f"• API Key: `{has_key}`\n"
        f"• API Secret: `{has_secret}`\n\n"
        "স্বয়ংক্রিয় পেমেন্ট ভেরিফিকেশনের জন্য বাইন্যান্স অ্যাকাউন্ট থেকে API Key এবং Secret Key সেট করুন:"
    )
    kbd = [
        [
            InlineKeyboardButton(
                "✏️ API Key সেট করুন", callback_data="adm_set_b_key"
            )
        ],
        [
            InlineKeyboardButton(
                "✏️ API Secret সেট করুন", callback_data="adm_set_b_secret"
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu"
            )
        ],
    ]
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
    return

  elif data == "adm_set_b_key":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_binance_api_key"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_binance_api_menu"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ আপনার Binance **API Key** টি চ্যাটে লিখে পাঠান:",
        reply_markup=InlineKeyboardMarkup(kbd),
    )
    return

  elif data == "adm_set_b_secret":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_binance_api_secret"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_binance_api_menu"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ আপনার Binance **API Secret Key** টি চ্যাটে লিখে পাঠান:",
        reply_markup=InlineKeyboardMarkup(kbd),
    )
    return

  elif data == "adm_admin_control_panel":
    if not is_admin(user_id):
      return
    await query.answer()
    admins = load_admins()
    text = "👑 **অ্যাডমিন তালিকা ও এক্সেস কন্ট্রোল:**\n\n"
    kbd = []
    for aid in admins:
      text += f"👤 **ID:** `{aid}`\n"
      if aid != DEFAULT_ADMIN_ID and aid != user_id:
        kbd.append([
            InlineKeyboardButton(
                f"❌ রিমুভ {aid}", callback_data=f"adm_rm_admin_{aid}"
            )
        ])
    kbd.append([
        InlineKeyboardButton(
            "➕ নতুন অ্যাডমিন যোগ করুন", callback_data="adm_add_admin_prompt"
        )
    ])
    kbd.append(
        [InlineKeyboardButton("🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu")]
    )
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
    return

  elif data.startswith("adm_rm_admin_"):
    if not is_admin(user_id):
      return
    try:
      target_aid = int(data.replace("adm_rm_admin_", ""))
      admins = load_admins()
      if target_aid in admins:
        admins.remove(target_aid)
        save_admins(admins)
        await query.answer(
            f"✅ অ্যাডমিন {target_aid} সফলভাবে রিমুভ করা হয়েছে!", show_alert=True
        )
    except Exception as e:
      await query.answer(f"ত্রুটি: {e}", show_alert=True)
    admins = load_admins()
    text = "👑 **অ্যাডমিন তালিকা ও এক্সেস কন্ট্রোল:**\n\n"
    kbd = []
    for aid in admins:
      text += f"👤 **ID:** `{aid}`\n"
      if aid != DEFAULT_ADMIN_ID and aid != user_id:
        kbd.append([
            InlineKeyboardButton(
                f"❌ রিমুভ {aid}", callback_data=f"adm_rm_admin_{aid}"
            )
        ])
    kbd.append([
        InlineKeyboardButton(
            "➕ নতুন অ্যাডমিন যোগ করুন", callback_data="adm_add_admin_prompt"
        )
    ])
    kbd.append(
        [InlineKeyboardButton("🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu")]
    )
    try:
      await query.message.edit_text(
          text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
      )
    except:
      pass
    return

  elif data == "adm_cat_mgmt_hub":
    if not is_admin(user_id):
      return
    await query.answer()
    kbd = [
        [
            InlineKeyboardButton(
                "🟢 ক্যাটাগরি চালু/বন্ধ প্যানেল", callback_data="adm_toggle_panel"
            )
        ],
        [
            InlineKeyboardButton(
                "➕ নতুন ক্যাটাগরি যোগ করুন", callback_data="adm_add_cat_prompt"
            ),
            InlineKeyboardButton(
                "🗑️ ক্যাটাগরি ডিলিট করুন", callback_data="adm_del_cat_menu"
            ),
        ],
        [
            InlineKeyboardButton(
                "💵 ক্যাটাগরি প্রাইস ও ডিসপ্লে ম্যানেজ",
                callback_data="adm_set_prices",
            )
        ],
        [
            InlineKeyboardButton(
                "💬 ক্যাটাগরি পপআপ মেসেজ (Closed & Shift)",
                callback_data="adm_custom_popup_menu",
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu"
            )
        ],
    ]
    await query.message.edit_text(
        "📂 **ক্যাটাগরি ম্যানেজমেন্ট হাব:**",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_del_cat_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    kbd = []
    for key, cat in settings["categories"].items():
      kbd.append([
          InlineKeyboardButton(
              f"❌ ডিলিট: {cat['name']}", callback_data=f"adm_delcat_{key}"
          )
      ])
    kbd.append([
        InlineKeyboardButton("🔙 ক্যাটাগরি মেনু", callback_data="adm_cat_mgmt_hub")
    ])
    await query.message.edit_text(
        "🗑️ **ক্যাটাগরি ডিলিট প্যানেল**\n\nযে ক্যাটাগরি স্থায়ীভাবে ডিলিট করতে চান তাতে ক্লিক করুন:",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("adm_delcat_"):
    if not is_admin(user_id):
      return
    cat_key = data.replace("adm_delcat_", "")
    if cat_key in settings["categories"]:
      cat_name = settings["categories"][cat_key]["name"]
      del settings["categories"][cat_key]
      save_settings(settings)
      await query.answer(
          f"✅ {cat_name} সফলভাবে ডিলিট করা হয়েছে!", show_alert=True
      )
    else:
      await query.answer("⛔ ক্যাটাগরি পাওয়া যায়নি!", show_alert=True)

    kbd = []
    for key, cat in settings["categories"].items():
      kbd.append([
          InlineKeyboardButton(
              f"❌ ডিলিট: {cat['name']}", callback_data=f"adm_delcat_{key}"
          )
      ])
    kbd.append([
        InlineKeyboardButton("🔙 ক্যাটাগরি মেনু", callback_data="adm_cat_mgmt_hub")
    ])
    try:
      await query.message.edit_text(
          "🗑️ **ক্যাটাগরি ডিলিট প্যানেল**\n\nযে ক্যাটাগরি স্থায়ীভাবে ডিলিট করতে চান তাতে ক্লিক করুন:",
          reply_markup=InlineKeyboardMarkup(kbd),
          parse_mode="Markdown",
      )
    except:
      pass
    return

  elif data == "adm_download_hub":
    if not is_admin(user_id):
      return
    await query.answer()
    kbd = [
        [
            InlineKeyboardButton(
                "📥 সব সময়ের জমা দেওয়া আইডির মার্জড ফাইল",
                callback_data="adm_download_menu",
            )
        ],
        [
            InlineKeyboardButton(
                "📅 আজকের ক্যাটাগরি ফাইল ডাউনলোড",
                callback_data="adm_today_download_menu",
            )
        ],
        [
            InlineKeyboardButton(
                "🗓 কাস্টম তারিখ অনুযায়ী ফাইল ডাউনলোড",
                callback_data="adm_custom_date_download_menu",
            )
        ],
        [
            InlineKeyboardButton(
                "📂 সকল ইউজারের সাবমিশন হিস্ট্রি (এক্সেল)",
                callback_data="adm_export_all_history",
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu"
            )
        ],
    ]
    await query.message.edit_text(
        "📥 **ফাইল ডাউনলোড হাব:**",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_pay_report_hub":
    if not is_admin(user_id):
      return
    await query.answer()
    kbd = [
        [
            InlineKeyboardButton(
                "📂 অটোমেটেড রেজাল্ট রিপোর্ট ও পেমেন্ট ম্যাচিং",
                callback_data="adm_report_menu",
            )
        ],
        [
            InlineKeyboardButton(
                "💰 ক্যাটাগরি পেমেন্ট ক্লিয়ার করুন (Paid)",
                callback_data="adm_pay_menu",
            )
        ],
        [
            InlineKeyboardButton(
                "👥 সকল ইউজারের পেমেন্ট স্ট্যাটাস (Global)",
                callback_data="adm_global_pay_list",
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu"
            )
        ],
    ]
    await query.message.edit_text(
        "💰 **পেমেন্ট ও রিপোর্ট ম্যাচিং প্যানেল:**",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_channel_mgmt":
    if not is_admin(user_id):
      return
    await query.answer()
    ch_data = settings.get("channels", {})
    join_ch = ch_data.get("join_channel", "@OMORTECH")
    sub_ch = ch_data.get("submission_channel", "@OMORTECH")
    with_ch = ch_data.get("withdrawal_channel", "@OMORTECH")
    supp_link = ch_data.get("support_link", "https://t.me/OMOR_BUYER1")

    text = (
        "📢 **চ্যানেল ও সাপোর্ট রাউটিং ম্যানেজমেন্ট**\n\n"
        f"• জয়েন চ্যানেল (মেম্বার চেক): `{join_ch}`\n"
        f"• সাবমিশন চ্যানেল: `{sub_ch}`\n"
        f"• উইথড্র ও নোটিফিকেশন চ্যানেল: `{with_ch}`\n"
        f"• সাপোর্ট আইডি/লিংক: `{supp_link}`\n\n"
        "*(টিপস: সাপোর্ট লিংকে `bot` লিখলে ইউজার সরাসরি বটের ভেতরই সাপোর্ট মেসেজ পাঠাতে পারবে)*\n\n"
        "পরিবর্তন করতে নিচের বাটনগুলোতে ক্লিক করুন:"
    )
    kbd = [
        [
            InlineKeyboardButton(
                "✏️ জয়েন চ্যানেল পরিবর্তন", callback_data="adm_set_join_ch"
            )
        ],
        [
            InlineKeyboardButton(
                "✏️ সাবমিশন চ্যানেল পরিবর্তন", callback_data="adm_set_sub_ch"
            )
        ],
        [
            InlineKeyboardButton(
                "✏️ উইথড্র চ্যানেল পরিবর্তন", callback_data="adm_set_with_ch"
            )
        ],
        [
            InlineKeyboardButton(
                "✏️ সাপোর্ট আইডি / লিংক পরিবর্তন",
                callback_data="adm_set_support_link",
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu"
            )
        ],
    ]
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
    return

  elif data == "adm_set_join_ch":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_join_channel"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_channel_mgmt"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ নতুন জয়েন চ্যানেলের ইউজারনেম, আইডি বা ইনভাইট লিংক লিখে পাঠান (যেমন:"
        " `@MyChannel` বা ব্যক্তিগত আইডি `-100...`):",
        reply_markup=InlineKeyboardMarkup(kbd),
    )
    return

  elif data == "adm_set_sub_ch":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_sub_channel"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_channel_mgmt"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ নতুন সাবমিশন চ্যানেলের ইউজারনেম বা আইডি লিখে পাঠান (যেমন:"
        " `@MyChannel` বা `-100...`):",
        reply_markup=InlineKeyboardMarkup(kbd),
    )
    return

  elif data == "adm_set_with_ch":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_with_channel"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_channel_mgmt"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ নতুন উইথড্র চ্যানেল বা গ্রুপের ইউজারনেম/আইডি লিখে পাঠান (যেমন:"
        " `@MyWithdrawChannel` বা `-100...`):",
        reply_markup=InlineKeyboardMarkup(kbd),
    )
    return

  elif data == "adm_set_support_link":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_support_link"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_channel_mgmt"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ নতুন সাপোর্ট আইডি, ইউজারনেম বা লিংক লিখে পাঠান (যেমন: `https://t.me/YourUsername` বা `@YourUsername`)。\n\n"
        "💡 *যদি চান ইউজার সরাসরি বটের ভেতরেই মেসেজ ও ছবি পাঠাবে, তবে শুধু `bot` লিখে পাঠান।*",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_add_admin_prompt":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_new_admin_id"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu"
            )
        ]
    ]
    await query.message.edit_text(
        "👑 নতুন অ্যাডমিনের **Telegram User ID** (যেমন: `6945991387`) চ্যাটে লিখে"
        " পাঠান:",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_add_cat_prompt":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_new_cat_key"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ক্যাটাগরি মেনু", callback_data="adm_cat_mgmt_hub"
            )
        ]
    ]
    instructions = (
        "➕ **নতুন ক্যাটাগরি যোগ করার নিয়ম:**\n\n"
        "১. প্রথমে ক্যাটাগরির একটি ইউনিক ইংরেজি কোড দিন (যেমন: `gmail_2fa` বা"
        " `telegram`)\n"
        "২. এরপর সুন্দর ডিসপ্লে নাম দিন (যেমন: `Gmail 2FA`)\n"
        "৩. সবশেষে প্রতি আইডির মূল্য বা রেট (যেমন `5.0`) দিন。\n\n"
        "দয়া করে প্রথম ধাপের জন্য ক্যাটাগরির ইংরেজি কোড (Key) লিখে পাঠান:"
    )
    await query.message.edit_text(
        instructions, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
    return

  elif data == "adm_wallet_mgmt":
    if not is_admin(user_id):
      return
    await query.answer()
    await adm_wallet_mgmt_render(query, settings)
    return

  elif data == "adm_add_wallet_prompt":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_add_wallet_name_step"] = True
    kbd = [[InlineKeyboardButton("🔙 ফিরে যান", callback_data="adm_wallet_mgmt")]]
    await query.message.edit_text(
        "✍ নতুন ওয়ালেটের নাম (যেমন: `Binance Pay` বা `USDT (TRC20)`) লিখে চ্যাটে পাঠান:",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("adm_del_wallet_"):
    if not is_admin(user_id):
      return
    w_name = data.replace("adm_del_wallet_", "")
    if "wallets" in settings and w_name in settings["wallets"]:
      del settings["wallets"][w_name]
      save_settings(settings)
    await query.answer(f"✅ {w_name} সফলভাবে ডিলিট করা হয়েছে!", show_alert=True)
    await adm_wallet_mgmt_render(query, settings)
    return

  elif data == "adm_user_manage_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    kbd = [
        [
            InlineKeyboardButton(
                "👤 নির্দিষ্ট ইউজারের ব্যালেন্স/রেফার এডিট (+/-)",
                callback_data="adm_edit_user_bal",
            )
        ],
        [
            InlineKeyboardButton(
                "🔄 নির্দিষ্ট ইউজারের উইথড্র হিস্ট্রি রিসেট (Withdrawn=0)",
                callback_data="adm_reset_user_withdrawn",
            )
        ],
        [
            InlineKeyboardButton(
                "🎁 নির্দিষ্ট ইউজারের সমস্ত রেফারেল হিস্ট্রি রিসেট",
                callback_data="adm_reset_user_referrals_all",
            )
        ],
        [
            InlineKeyboardButton(
                "🗑️ নির্দিষ্ট ইউজারের সাবমিশন হিস্ট্রি রিসেট",
                callback_data="adm_reset_user_sub",
            )
        ],
        [
            InlineKeyboardButton(
                "🔍 ইউজার সার্চ প্যানেল", callback_data="adm_search_panel"
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu"
            )
        ],
    ]
    await query.message.edit_text(
        "👤 **ইউজার ডাটা কন্ট্রোল প্যানেল:**",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_edit_user_bal":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_user_edit_data"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_user_manage_menu"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ নির্দিষ্ট ইউজারের ব্যালেন্স বা রেফার বাড়াতে/কমাতে এই ফরম্যাটে লিখে পাঠান:\n\n"
        "`USER_ID +50` (টাকা যোগ করতে বা টাকা কাটতে `-50`)\n"
        "`USER_ID ref+2` (রেফার বাড়াতে)\n"
        "`USER_ID ref-1` (ফেক রেফার কাটতে)\n\n"
        "উদাহরণ: `8276619021 +20` বা `8276619021 -50`",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_reset_user_withdrawn":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_user_reset_withdrawn"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_user_manage_menu"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ নির্দিষ্ট ইউজারের উইথড্র করা অ্যামাউন্ট রিসেট করে ০ (শূন্য) করতে তার **User ID** লিখে পাঠান:\n\nউদাহরণ: `8276619021`",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_reset_user_referrals_all":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_user_reset_referrals_all"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_user_manage_menu"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ নির্দিষ্ট ইউজারের সমস্ত রেফারেল হিস্ট্রি (রেফার লিস্ট, বোনাস, কমিশন, উইথড্র) সম্পূর্ণ রিসেট বা মুছে ফেলতে তার **User ID** লিখে পাঠান:\n\nউদাহরণ: `8276619021`",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_reset_user_sub":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_user_reset_id"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_user_manage_menu"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ নির্দিষ্ট ইউজারের সমস্ত সাবমিশন হিস্ট্রি ক্লিয়ার বা রিসেট করতে তার **User"
        " ID** লিখে পাঠান:\n\nউদাহরণ: `6945991387`",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_custom_popup_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    popup_kbd = []
    for key, cat in settings["categories"].items():
      popup_kbd.append([
          InlineKeyboardButton(
              f"💬 {cat['name']} পপআপ এডিট",
              callback_data=f"adm_edit_cat_popup_{key}",
          )
      ])
    popup_kbd.append([
        InlineKeyboardButton("🔙 ক্যাটাগরি মেনু", callback_data="adm_cat_mgmt_hub")
    ])
    await query.message.edit_text(
        "💬 **ক্যাটাগরিভিত্তিক পপআপ মেসেজ ম্যানেজমেন্ট**\n\nযে ক্যাটাগরির বন্ধ থাকার"
        " বা শিফটের পপআপ মেসেজ এডিট করতে চান তাতে ক্লিক করুন:",
        reply_markup=InlineKeyboardMarkup(popup_kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("adm_edit_cat_popup_"):
    if not is_admin(user_id):
      return
    await query.answer()
    cat_key = data.replace("adm_edit_cat_popup_", "")
    cat_name = settings["categories"][cat_key]["name"]
    p_closed = settings["categories"][cat_key].get(
        "popup_closed", "দুঃখিত {name}, বর্তমানে {category} বন্ধ রয়েছে!"
    )
    p_shift = settings["categories"][cat_key].get(
        "popup_shift", "দুঃখিত {name}, বর্তমানে {category} শিফটের বাইরে রয়েছে!"
    )

    context.user_data["waiting_for_popup_choice"] = cat_key
    kbd = [
        [
            InlineKeyboardButton(
                "🔒 ক্যাটাগরি অফ পপআপ এডিট",
                callback_data=f"popclosed_{cat_key}",
            )
        ],
        [
            InlineKeyboardButton(
                "⏰ শিফট টাইম অফ পপআপ এডিট",
                callback_data=f"popshift_{cat_key}",
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_custom_popup_menu"
            )
        ],
    ]
    await query.message.edit_text(
        f"✍ আপনি **{cat_name}** ক্যাটাগরির পপআপ এডিট করতে চান。\n\n"
        f"• **বর্তমান অফ পপআপ:** `{p_closed}`\n"
        f"• **বর্তমান শিফট পপআপ:** `{p_shift}`\n\nকোন পপআপটি এডিট করতে চান তা সিলেক্ট"
        " করুন:",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("popclosed_"):
    if not is_admin(user_id):
      return
    await query.answer()
    cat_key = data.replace("popclosed_", "")
    context.user_data["waiting_for_cat_popup_edit"] = ("closed", cat_key)
    cat_name = settings["categories"][cat_key]["name"]
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data=f"adm_edit_cat_popup_{cat_key}"
            )
        ]
    ]
    await query.message.edit_text(
        f"✍ **{cat_name}** ক্যাটাগরি ম্যানুয়ালি অফ থাকলে যে পপআপ দেখাবে তা লিখে"
        " পাঠান ( `{category}` এবং `{name}` ব্যবহার করতে পারেন):",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("popshift_"):
    if not is_admin(user_id):
      return
    await query.answer()
    cat_key = data.replace("popshift_", "")
    context.user_data["waiting_for_cat_popup_edit"] = ("shift", cat_key)
    cat_name = settings["categories"][cat_key]["name"]
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data=f"adm_edit_cat_popup_{cat_key}"
            )
        ]
    ]
    await query.message.edit_text(
        f"✍ **{cat_name}** শিফটের বাইরে থাকলে যে পপআপ দেখাবে তা লিখে পাঠান"
        " (`{category}` এবং `{name}` ব্যবহার করতে পারেন):",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_referral_settings":
    if not is_admin(user_id):
      return
    await query.answer()
    ref_data = settings.get("referral", {})
    status_str = "🟢 চালু" if ref_data.get("status", True) else "⛔ বন্ধ"
    comm_status_str = (
        "🟢 চালু" if ref_data.get("commission_status", True) else "⛔ বন্ধ"
    )
    text = (
        "🎁 **রেফারেল সিস্টেম ও কমিশন ম্যানেজমেন্ট**\n\n"
        f"• রেফারেল স্ট্যাটাস: `{status_str}`\n"
        f"• প্রতি রেফার বোনাস: `৳{ref_data.get('bonus_amount', 10.0)}`\n"
        f"• রেফারেল কমিশন (%): `{ref_data.get('commission_percent', 5.0)}%`\n"
        f"• কমিশন স্ট্যাটাস: `{comm_status_str}`\n"
        f"• ন্যূনতম রেফার শর্ত: `{ref_data.get('min_refs_to_withdraw', 10)} টি`\n"
        f"• ন্যূনতম উইথড্র টাকা: `৳{ref_data.get('min_amount_to_withdraw', 50.0)}`\n\n"
        "*(নোট: রেফারেল বোনাস এবং কাজের কমিশন সম্পূর্ণ আলাদাভাবে হিসাব হয়ে স্বয়ংক্রিয়ভাবে ব্যালেন্স যোগ হয়)*"
    )
    kbd = [
        [
            InlineKeyboardButton(
                f"টগল রেফারেল: {status_str}",
                callback_data="adm_ref_toggle_status",
            )
        ],
        [
            InlineKeyboardButton(
                f"টগল কমিশন (%): {comm_status_str}",
                callback_data="adm_ref_toggle_comm",
            )
        ],
        [
            InlineKeyboardButton(
                "✏️ কমিশন পার্সেন্ট (%)", callback_data="adm_ref_set_comm_pct"
            ),
            InlineKeyboardButton(
                "✏️ বোনাস অ্যামাউন্ট", callback_data="adm_ref_set_bonus"
            ),
        ],
        [
            InlineKeyboardButton(
                "✏️ ন্যূনতম রেফার শর্ত", callback_data="adm_ref_set_min_refs"
            ),
            InlineKeyboardButton(
                "✏️ ন্যূনতম উইথড্র টাকা", callback_data="adm_ref_set_min_amt"
            ),
        ],
        [
            InlineKeyboardButton(
                "🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu"
            )
        ],
    ]
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
    return

  elif data == "adm_ref_toggle_status":
    if not is_admin(user_id):
      return
    settings["referral"]["status"] = not settings["referral"].get(
        "status", True
    )
    save_settings(settings)
    await query.answer("স্ট্যাটাস পরিবর্তন হয়েছে!", show_alert=True)
    await adm_referral_menu_render(query, settings)
    return

  elif data == "adm_ref_toggle_comm":
    if not is_admin(user_id):
      return
    settings["referral"]["commission_status"] = not settings["referral"].get(
        "commission_status", True
    )
    save_settings(settings)
    await query.answer("কমিশন স্ট্যাটাস পরিবর্তন হয়েছে!", show_alert=True)
    await adm_referral_menu_render(query, settings)
    return

  elif data == "adm_ref_set_comm_pct":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_comm_percent"] = True
    kbd = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_referral_settings"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ উইথড্র থেকে কত % কমিশন রেফারার পাবে তা সংখ্যায় লিখে পাঠান (যেমন: `5`"
        " বা `10`):",
        reply_markup=InlineKeyboardMarkup(kbd),
    )
    return

  elif data == "adm_ref_set_bonus":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_ref_bonus"] = True
    keyboard = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_referral_settings"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ প্রতি রেফারে কত টাকা বোনাস দিতে চান তা সংখ্যায় লিখে পাঠান:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return

  elif data == "adm_ref_set_min_refs":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_min_refs"] = True
    keyboard = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_referral_settings"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ উইথড্র করার জন্য ন্যূনতম কয়টি রেফার লাগবে তা সংখ্যায় লিখে পাঠান"
        " (যেমন: `10`):",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return

  elif data == "adm_ref_set_min_amt":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_min_amt"] = True
    keyboard = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_referral_settings"
            )
        ]
    ]
    await query.message.edit_text(
        "✍ উইথড্র করার জন্য ন্যূনতম কত টাকা থাকতে হবে তা সংখ্যায় লিখে পাঠান"
        " (যেমন: `50`):",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return

  elif data == "referral_menu_user":
    await query.answer()
    ref_data = settings.get("referral", {"status": True})
    if not ref_data.get("status", True):
      await query.answer(
          f"⛔ দুঃখিত {user.first_name}, রেফারেল সিস্টেম বর্তমানে বন্ধ রয়েছে!", show_alert=True
      )
      return

    referrals = load_referrals()
    ref_info = referrals.get("referrers", {}).get(
        user_id_str,
        {
            "refs": [],
            "withdrawn": 0.0,
            "commission_earned": 0.0,
            "manual_adjustment": 0.0,
        },
    )

    user_refs = ref_info.get("refs", [])
    total_refs_count = len(user_refs)

    total_earned_ref_bonus = sum(
        float(r.get("bonus", 10.0)) for r in user_refs
    )
    commission_earned = float(ref_info.get("commission_earned", 0.0))
    manual_adj = float(ref_info.get("manual_adjustment", 0.0))

    total_combined_earned = (
        total_earned_ref_bonus + commission_earned + manual_adj
    )
    already_withdrawn = float(ref_info.get("withdrawn", 0.0))
    current_available_bal = max(0.0, total_combined_earned - already_withdrawn)

    bot_username = (await context.bot.get_me()).username
    ref_link = f"https://t.me/{bot_username}?start=ref_{user_id}"

    profile_text = (
        "👤 **রেফারেল ও কমিশন প্রোফাইল**\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        f"🔹 **নাম:** {user.first_name}\n"
        f"🆔 **User ID:** `{user_id}`\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        f"👥 **মোট রেফার করেছেন:** `{total_refs_count} জন`\n"
        f"🎁 **রেফার বোনাস আয়:** `৳{total_earned_ref_bonus}`\n"
        f"💼 **কমিশন আয়:** `৳{commission_earned}`\n"
        f"💰 **মোট সম্মিলিত আয় : ৳{max(0.0, total_combined_earned)}**\n"
        f"💸 **উইথড্র করা হয়েছে:** `৳{already_withdrawn}`\n"
        f"💵 **উত্তোলনযোগ্য ব্যালেন্স:** `৳{current_available_bal}`\n\n"
        f"🔗 **আপনার রেফারেল লিংক:**\n`{ref_link}`\n\n"
        "*(লিংকটিতে ক্লিক করলেই কপি হয়ে যাবে)*"
    )
    keyboard = [
        [
            InlineKeyboardButton(
                "💸 উইথড্র করুন (Withdraw)",
                callback_data="withdraw_ref_bonus",
            )
        ],
        [InlineKeyboardButton("🔙 মূল মেনু", callback_data="back_home")],
    ]
    await query.message.edit_text(
        profile_text,
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown",
    )
    return

  elif data == "withdraw_ref_bonus":
    ref_data = settings.get("referral", {})
    referrals = load_referrals()
    ref_info = referrals.get("referrers", {}).get(
        user_id_str,
        {
            "refs": [],
            "withdrawn": 0.0,
            "commission_earned": 0.0,
            "manual_adjustment": 0.0,
        },
    )

    user_refs = ref_info.get("refs", [])
    total_refs_count = len(user_refs)

    total_earned_ref_bonus = sum(
        float(r.get("bonus", 10.0)) for r in user_refs
    )
    commission_earned = float(ref_info.get("commission_earned", 0.0))
    manual_adj = float(ref_info.get("manual_adjustment", 0.0))
    total_combined_earned = (
        total_earned_ref_bonus + commission_earned + manual_adj
    )

    already_withdrawn = float(ref_info.get("withdrawn", 0.0))
    current_available_bal = max(0.0, total_combined_earned - already_withdrawn)

    min_refs = ref_data.get("min_refs_to_withdraw", 10)
    min_amt = ref_data.get("min_amount_to_withdraw", 50.0)

    if total_refs_count < min_refs and current_available_bal < min_amt:
      await query.answer(
          f"⛔ সম্মানিত {user.first_name}, উইথড্র শর্ত পূরণ হয়নি!\nকমপক্ষে {min_refs}টি রেফার এবং ন্যূনতম ৳{min_amt} ব্যালেন্স থাকতে হবে।",
          show_alert=True,
      )
      return

    if current_available_bal < min_amt:
      await query.answer(
          f"⛔ সম্মানিত {user.first_name}, উইথড্র শর্ত পূরণ হয়নি!\nন্যূনতম ৳{min_amt} উত্তোলনযোগ্য ব্যালেন্স থাকতে হবে।",
          show_alert=True,
      )
      return

    user_p = user_payments.get(user_id_str, {})
    has_method = any(p_v != "সেট করা হয়নি" for p_k, p_v in user_p.items())
    if not has_method and not is_admin(user_id):
      await query.answer(
          f"⛔ {user.first_name}, পেমেন্ট নম্বর সেট করা নেই! প্রথমে পেমেন্ট মেনু থেকে নম্বর সেট করুন।",
          show_alert=True,
      )
      return

    withdraw_amount = current_available_bal

    ref_info["withdrawn"] = already_withdrawn + current_available_bal
    referrals.setdefault("referrers", {})[user_id_str] = ref_info
    save_referrals(referrals)

    await query.answer(
        f"🎉 প্রিয় {user.first_name}, আপনার রেফারেল উইথড্র রিকোয়েস্ট সফলভাবে জমা হয়েছে!",
        show_alert=True,
    )

    active_payments = []
    for p_k, p_v in user_p.items():
      if p_v != "সেট করা হয়নি":
        active_payments.append(f"• **{p_k.capitalize()}**: `{p_v}`")
    p_info = (
        "\n".join(active_payments)
        if active_payments
        else "কোনো নম্বর সেট করা নেই"
    )

    with_ch_raw = settings.get("channels", {}).get(
        "withdrawal_channel", DEFAULT_ADMIN_ID
    )
    with_ch = parse_chat_id(with_ch_raw)

    wd_kbd = [
        [
            InlineKeyboardButton(
                "✅ Pay Done (Accept)",
                callback_data=f"adm_wd_done_{user_id}_{withdraw_amount}",
            )
        ]
    ]

    try:
      await context.bot.send_message(
          chat_id=with_ch,
          text=(
              "📥 **নতুন রেফারেল উইথড্র রিকোয়েস্ট!**\n\n"
              f"👤 ইউজার: [{user.first_name}](tg://user?id={user_id})"
              f" (`{user_id}`)\n"
              f"💰 উইথড্র পরিমাণ: `৳{withdraw_amount}`\n"
              f"👥 মোট রেফার: `{total_refs_count} জন`\n\n"
              f"📌 **পেমেন্ট তথ্য:**\n{p_info}"
          ),
          reply_markup=InlineKeyboardMarkup(wd_kbd),
          parse_mode="Markdown",
      )
    except Exception as e:
      print(f"Error sending withdrawal request to channel: {e}")
    return

  elif data.startswith("adm_wd_done_"):
    if not is_admin(user_id):
      await query.answer("⛔ আপনার অনুমতি নেই!", show_alert=True)
      return

    parts = data.replace("adm_wd_done_", "").split("_")
    target_uid = int(parts[0])
    wd_amount = parts[1]

    await query.answer(
        "✅ পেমেন্ট সফলভাবে এপ্রুভ ও পে ডান করা হয়েছে!", show_alert=True
    )

    try:
      current_text = query.message.text
      new_text = current_text + "\n\n📌 **Status: ✅ PAID & ACCEPTED BY ADMIN**"
      await query.message.edit_text(
          new_text, reply_markup=None, parse_mode="Markdown"
      )
    except:
      pass

    try:
      all_u = load_all_users()
      target_name = all_u.get(str(target_uid), "Member")
      success_msg = (
          f"🎉 **অভিনন্দন {target_name}! আপনার উইথড্র রিকোয়েস্ট সফলভাবে পেমেন্ট করা হয়েছে।** 💸\n\n"
          f"💰 পরিমাণ: `৳{wd_amount}`\n"
          "📌 স্ট্যাটাস: ✅ **SUCCESSFULLY PAID**"
      )
      await context.bot.send_message(
          chat_id=target_uid, text=success_msg, parse_mode="Markdown"
      )
    except Exception as e:
      print(f"Error notifying user about withdrawal: {e}")
    return

  elif data.startswith("adm_crypto_done_"):
    if not is_admin(user_id):
      await query.answer("⛔ আপনার অনুমতি নেই!", show_alert=True)
      return

    parts = data.replace("adm_crypto_done_", "").split("_", 1)
    target_uid = int(parts[0])
    w_name = parts[1]

    await query.answer(
        "✅ ক্রিপ্টো পেমেন্ট সফলভাবে এপ্রুভ ও পে ডান করা হয়েছে!", show_alert=True
    )

    try:
      if query.message.caption:
        current_caption = query.message.caption or ""
        new_caption = (
            current_caption
            + f"\n\n📌 **Status: ✅ PAID & ACCEPTED BY ADMIN ({w_name})**"
        )
        await query.message.edit_caption(
            caption=new_caption, reply_markup=None, parse_mode="Markdown"
        )
      else:
        current_text = query.message.text or ""
        new_text = (
            current_text
            + f"\n\n📌 **Status: ✅ PAID & ACCEPTED BY ADMIN ({w_name})**"
        )
        await query.message.edit_text(
            text=new_text, reply_markup=None, parse_mode="Markdown"
        )
    except Exception as e:
      print(f"Error updating crypto caption/text: {e}")

    try:
      all_u = load_all_users()
      target_name = all_u.get(str(target_uid), "Member")
      success_msg = (
          f"🎉 **অভিনন্দন {target_name}! আপনার {w_name} ক্রিপ্টো উইথড্র রিকোয়েস্ট সফলভাবে পেমেন্ট করা হয়েছে।** 💸\n\n"
          "📌 স্ট্যাটাস: ✅ **SUCCESSFULLY PAID**"
      )
      await context.bot.send_message(
          chat_id=target_uid, text=success_msg, parse_mode="Markdown"
      )
    except Exception as e:
      print(f"Error notifying user about crypto withdrawal: {e}")
    return

  elif data == "adm_toggle_panel":
    if not is_admin(user_id):
      return
    await query.answer()
    await adm_toggle_panel_render(query, settings)
    return

  elif data.startswith("adm_tglcat_"):
    if not is_admin(user_id):
      return
    cat_key = data.replace("adm_tglcat_", "")
    if cat_key in settings["categories"]:
      current_status = settings["categories"][cat_key]["status"]
      settings["categories"][cat_key]["status"] = not current_status
      save_settings(settings)

    await query.answer("স্ট্যাটাস পরিবর্তন হয়েছে!", show_alert=True)
    await adm_toggle_panel_render(query, settings)
    return

  elif data == "adm_search_panel":
    if not is_admin(user_id):
      return
    await query.answer()
    context.user_data["waiting_for_search_user"] = True
    keyboard = [
        [
            InlineKeyboardButton(
                "🔙 ইউজার মেনু", callback_data="adm_user_manage_menu"
            )
        ]
    ]
    await query.message.edit_text(
        "🔍 **ইউজার সার্চ প্যানেল**\n\nদয়া করে যে ইউজারের হিস্ট্রি দেখতে চান তার"
        " **User ID** (যেমন: `6945991387`) চ্যাটে লিখে পাঠান:",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_toggle_maintenance":
    if not is_admin(user_id):
      return
    MAINTENANCE_MODE = not MAINTENANCE_MODE
    status_str = "ENABLED (চালু)" if MAINTENANCE_MODE else "DISABLED (বন্ধ)"
    await query.answer(f"Maintenance Mode is now {status_str}", show_alert=True)
    m_status_text = "🟢 ON" if MAINTENANCE_MODE else "🔴 OFF"
    admin_keyboard = [
        [
            InlineKeyboardButton(
                f"⚠️ মেইনটেনেন্স মোড: {m_status_text}",
                callback_data="adm_toggle_maintenance",
            )
        ],
        [
            InlineKeyboardButton(
                "💬 বটের সমস্ত মেসেজ এডিট প্যানেল", callback_data="adm_custom_messages_menu"
            )
        ],
        [
            InlineKeyboardButton(
                "🔑 Binance API সেটিংস", callback_data="adm_binance_api_menu"
            )
        ],
        [
            InlineKeyboardButton(
                "📂 ক্যাটাগরি ম্যানেজমেন্ট ও পপআপ",
                callback_data="adm_cat_mgmt_hub",
            )
        ],
        [
            InlineKeyboardButton(
                "📥 ফাইল ডাউনলোড হাব (Download Hub)",
                callback_data="adm_download_hub",
            )
        ],
        [
            InlineKeyboardButton(
                "💰 পেমেন্ট ও রিপোর্ট ম্যাচিং প্যানেল",
                callback_data="adm_pay_report_hub",
            )
        ],
        [
            InlineKeyboardButton(
                "👤 ইউজার ব্যালেন্স/রেফার ও ডাটা কন্ট্রোল",
                callback_data="adm_user_manage_menu",
            )
        ],
        [
            InlineKeyboardButton(
                "🌐 ক্রিপ্টো উইথড্র হিস্ট্রি ও পরিসংখ্যান",
                callback_data="adm_crypto_stats",
            )
        ],
        [
            InlineKeyboardButton(
                "🌐 ক্রিপ্টো ওয়ালেট অ্যাড্রেস ম্যানেজ",
                callback_data="adm_wallet_mgmt",
            )
        ],
        [
            InlineKeyboardButton(
                "📢 চ্যানেল ও সাপোর্ট রাউটিং সেটআপ",
                callback_data="adm_channel_mgmt",
            )
        ],
        [
            InlineKeyboardButton(
                "👑 অ্যাডমিন কন্ট্রোল ও রিমুভ প্যানেল",
                callback_data="adm_admin_control_panel",
            )
        ],
        [
            InlineKeyboardButton(
                "⏰ শিফট টাইম সেটআপ", callback_data="adm_shift_menu"
            )
        ],
        [
            InlineKeyboardButton(
                "🎁 রেফারেল সিস্টেম ও কমিশন",
                callback_data="adm_referral_settings",
            )
        ],
        [
            InlineKeyboardButton(
                "🎛 পেমেন্ট মাধ্যম অন/অফ", callback_data="adm_toggle_payments"
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 মূল মেনুতে ফিরে যান", callback_data="back_home"
            )
        ],
    ]
    try:
      await query.message.edit_text(
          "🛠 **অ্যাডমিন কন্ট্রোল প্যানেল (মাস্টার হাব)**\n\nনিচের অপশনগুলো থেকে আপনার"
          " বট সম্পূর্ণ নিয়ন্ত্রণ করুন:",
          reply_markup=InlineKeyboardMarkup(admin_keyboard),
          parse_mode="Markdown",
      )
    except:
      pass
    return

  elif data == "adm_shift_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    shift_kbd = []
    for key, cat in settings["categories"].items():
      s_info = CATEGORY_SHIFTS.get(key, {"start": "00:00", "end": "23:59"})
      shift_kbd.append([
          InlineKeyboardButton(
              f"⏰ {cat['name']} ({s_info['start']} - {s_info['end']})",
              callback_data=f"adm_set_shift_{key}",
          )
      ])
    shift_kbd.append([
        InlineKeyboardButton("🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu")
    ])
    await query.message.edit_text(
        "⏰ **ক্যাটাগরি শিফট শিডিউলার (মিনিটসহ)**\n\nযে ক্যাটাগরির সময় পরিবর্তন"
        " করতে চান তাতে ক্লিক করুন:",
        reply_markup=InlineKeyboardMarkup(shift_kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("adm_set_shift_"):
    if not is_admin(user_id):
      return
    await query.answer()
    cat_key = data.replace("adm_set_shift_", "")
    context.user_data["waiting_for_shift_cat"] = cat_key
    cat_name = settings["categories"][cat_key]["name"]
    keyboard = [
        [InlineKeyboardButton("🔙 ফিরে যান", callback_data="adm_shift_menu")]
    ]
    await query.message.edit_text(
        f"✍ আপনি **{cat_name}** ক্যাটাগরির শিফট টাইম সেট করতে চান。\n\nদয়া করে"
        " শুরু এবং শেষের সময় **ঘণ্টা:মিনিট** ফরম্যাটে চ্যাটে লিখে পাঠান:\n`HH:MM"
        " HH:MM` (উদাহরণস্বরূপ: `10:30 22:45`)",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_report_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    rep_kbd = []
    for key, cat in settings["categories"].items():
      rep_kbd.append([
          InlineKeyboardButton(
              f"📂 {cat['name']} রেজাল্ট রিপোর্ট আপলোড",
              callback_data=f"adm_upload_rep_{key}",
          )
      ])
    rep_kbd.append([
        InlineKeyboardButton("🔙 পেমেন্ট হাব", callback_data="adm_pay_report_hub")
    ])
    await query.message.edit_text(
        "📂 **অটোমেটেড রেজাল্ট রিপোর্ট ও পেমেন্ট ম্যাচিং**\n\nযে ক্যাটাগরির"
        " ভ্যালিড আইডি ফাইল আপলোড করে মেলাতে চান তা সিলেক্ট করুন:",
        reply_markup=InlineKeyboardMarkup(rep_kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("adm_upload_rep_"):
    if not is_admin(user_id):
      return
    await query.answer()
    cat_key = data.replace("adm_upload_rep_", "")
    cat_name = settings["categories"][cat_key]["name"]

    date_kbd = []
    today = datetime.datetime.now(TZ)
    for i in range(7):
      dt = today - datetime.timedelta(days=i)
      dt_str = dt.strftime("%Y-%m-%d")
      display_str = dt.strftime("%d-%m-%Y")

      logs = load_logs()
      reported = any(
          l["category"] == cat_name
          and l["date"].startswith(dt_str)
          and l.get("has_report", False)
          for l in logs
      )
      rep_status_icon = "🟢 Reported" if reported else "🔴 Pending"

      date_kbd.append([
          InlineKeyboardButton(
              f"📅 {display_str} [{rep_status_icon}]",
              callback_data=f"repdate_{cat_key}_{dt_str}",
          )
      ])
    date_kbd.append(
        [InlineKeyboardButton("🔙 ফিরে যান", callback_data="adm_report_menu")]
    )

    await query.message.edit_text(
        f"📅 **{cat_name}** ক্যাটাগরির কোন তারিখের কাজের রিপোর্ট আপলোড করতে চান"
        " তা নিচের ৭ দিনের তালিকা থেকে সিলেক্ট করুন:",
        reply_markup=InlineKeyboardMarkup(date_kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("repdate_"):
    if not is_admin(user_id):
      return
    await query.answer()

    remainder = data.replace("repdate_", "", 1)
    target_date_str = remainder[-10:]
    cat_key = remainder[:-11]

    if cat_key not in settings["categories"]:
      await query.answer("⛔ ক্যাটাগরি পাওয়া যায়নি!", show_alert=True)
      return

    cat_name = settings["categories"][cat_key]["name"]
    context.user_data["waiting_for_report_info"] = {
        "cat_key": cat_key,
        "cat_name": cat_name,
        "date": target_date_str,
    }

    display_date = datetime.datetime.strptime(
        target_date_str, "%Y-%m-%d"
    ).strftime("%d-%m-%Y")

    kbd = [[InlineKeyboardButton("🔙 ফিরে যান", callback_data="adm_report_menu")]]
    await query.message.edit_text(
        f"📂 **{cat_name}** ({display_date}) তারিখের ভ্যালিড আইডির তালিকা ফাইল"
        " (.txt বা .xlsx) এই চ্যাটে সরাসরি আপলোড করুন。\n\n*(উল্লেখ্য: শুধুমাত্র"
        " এই নির্দিষ্ট তারিখের সাবমিশনগুলোরই হিসাব হবে)*",
        reply_markup=InlineKeyboardMarkup(kbd),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_export_all_history":
    if not is_admin(user_id):
      return
    logs = load_logs()
    if not logs:
      await query.answer("⛔ কোনো সাবমিশন লগ পাওয়া যায়নি!", show_alert=True)
      return
    await query.answer()

    export_wb = openpyxl.Workbook()
    export_ws = export_wb.active
    export_ws.title = "All Users History"

    export_ws.append([
        "তারিখ ও সময়",
        "ইউজার আইডি",
        "ইউজার নাম",
        "ক্যাটাগরি",
        "মোট ফাইল",
        "ভ্যালিড",
        "রিজেক্ট",
        "ডুপ্লিকেট",
        "মোট টাকা (৳)",
        "রিপোর্ট স্ট্যাটাস",
        "পেমেন্ট স্ট্যাটাস",
    ])
    for l in logs:
      p_status = "Paid" if l.get("is_paid", False) else "Unpaid"
      rep_status = (
          "Reported" if l.get("has_report", False) else "Pending Report"
      )
      export_ws.append([
          l.get("date", ""),
          l.get("user_id", ""),
          l.get("user_name", ""),
          l.get("category", ""),
          l.get("total", 0),
          l.get("unique", 0),
          l.get("rejected", 0),
          l.get("duplicates", 0),
          l.get("earned_amount", 0),
          rep_status,
          p_status,
      ])

    file_name = "All_Users_Submission_History.xlsx"
    export_wb.save(file_name)

    with open(file_name, "rb") as f:
      await context.bot.send_document(
          chat_id=user_id,
          document=f,
          caption=(
              "📁 **সকল মেম্বারদের সমস্ত কাজের সাবমিশন ও Paid/Unpaid হিস্ট্রি"
              " ফাইল হিসেবে সংযুক্ত করা হলো।** ✨"
          ),
          parse_mode="Markdown",
      )
    os.remove(file_name)
    return

  elif data == "adm_pay_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    pay_kbd = []
    for key, cat in settings["categories"].items():
      pay_kbd.append([
          InlineKeyboardButton(
              f"💳 {cat['name']} সাবমিশন লিস্ট",
              callback_data=f"adm_paysublist_{key}",
          )
      ])
    pay_kbd.append([
        InlineKeyboardButton("🔙 পেমেন্ট হাব", callback_data="adm_pay_report_hub")
    ])
    await query.message.edit_text(
        "💰 **ক্যাটাগরিভিত্তিক পৃথক সাবমিশন পেমেন্ট প্যানেল**\n\nযে ক্যাটাগরির প্রতিটি"
        " কাজের আলাদা পেমেন্ট লিস্ট দেখতে চান তাতে ক্লিক করুন:",
        reply_markup=InlineKeyboardMarkup(pay_kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("adm_paysublist_"):
    if not is_admin(user_id):
      return
    cat_key = data.replace("adm_paysublist_", "")
    if cat_key not in settings["categories"]:
      return
    cat_name = settings["categories"][cat_key]["name"]

    logs = load_logs()
    cat_logs_with_index = []
    for idx, l in enumerate(logs):
      if l["category"] == cat_name and not l.get("is_paid", False):
        cat_logs_with_index.append((idx, l))

    if not cat_logs_with_index:
      await query.answer(
          f"⛔ {cat_name} ক্যাটাগরিতে কোনো বকেয়া (Unpaid) সাবমিশন নেই!",
          show_alert=True,
      )
      return
    await query.answer()

    text = f"💰 **{cat_name} ক্যাটাগরির পৃথক সাবমিশন ও বকেয়া তালিকা:**\n\n"
    kbd = []

    for idx, l in cat_logs_with_index[:10]:
      user_link = f"[{l['user_name']}](tg://user?id={l['user_id']})"
      rep_badge = "🟢 Reported" if l.get("has_report", False) else "🔴 Pending"
      text += (
          f"👤 **{user_link}** (`{l['user_id']}`)\n   - তারিখ: `{l['date']}` |"
          f" ভ্যালিড: `{l['unique']}` টি | রিজেক্ট: `{l.get('rejected', 0)}` টি |"
          f" {rep_badge} | 🔴 বকেয়া: `৳{l.get('earned_amount', 0)}`\n\n"
      )
      kbd.append([
          InlineKeyboardButton(
              f"💳 {l['user_name']} ({l['date']} -"
              f" {l.get('earned_amount', 0)}৳) Paid",
              callback_data=f"payprev_idx_{idx}",
          )
      ])

    kbd.append([
        InlineKeyboardButton(
            "⚡ এক ক্লিকে সব পেমেন্ট সম্পন্ন করুন (Pay All)",
            callback_data=f"pay_bulk_{cat_key}",
        )
    ])
    kbd.append(
        [InlineKeyboardButton("🔙 পেমেন্ট মেনু", callback_data="adm_pay_menu")]
    )
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
    return

  elif data.startswith("pay_bulk_"):
    if not is_admin(user_id):
      return
    cat_key = data.replace("pay_bulk_", "")
    if cat_key not in settings["categories"]:
      return
    cat_name = settings["categories"][cat_key]["name"]

    logs = load_logs()
    count = 0
    for idx, l in enumerate(logs):
      if l["category"] == cat_name and not l.get("is_paid", False):
        l["is_paid"] = True
        count += 1
        try:
          await context.bot.send_message(
              chat_id=l["user_id"],
              text=(
                  f"🎉 **প্রিয় {l['user_name']}, আপনার {cat_name} কাজের ৳{l.get('earned_amount', 0)}"
                  " পেমেন্ট সফলভাবে সম্পন্ন হয়েছে!**"
              ),
              parse_mode="Markdown",
          )
        except:
          pass

    save_logs(logs)
    await query.answer(
        f"🎉 {count} টি সাবমিশনের পেমেন্ট এক ক্লিকে সম্পূর্ণ করা হয়েছে!",
        show_alert=True,
    )

    query.data = f"adm_paysublist_{cat_key}"
    await button_callback(update, context)
    return

  elif data.startswith("payprev_idx_"):
    if not is_admin(user_id):
      return
    await query.answer()
    log_idx = int(data.replace("payprev_idx_", ""))
    logs = load_logs()
    if log_idx >= len(logs):
      await query.answer("⛔ লগ পাওয়া যায়নি!", show_alert=True)
      return

    l = logs[log_idx]
    target_uid = l["user_id"]
    target_name = l["user_name"]
    cat_name = l["category"]
    amount = l.get("earned_amount", 0.0)
    sub_date = l.get("date", "")

    target_uid_str = str(target_uid)
    user_p = user_payments.get(target_uid_str, {})

    active_payments_list = []
    for p_k, p_v in user_p.items():
      if p_v != "সেট করা হয়নি":
        active_payments_list.append(f"• {p_k.capitalize()}: `{p_v}`")

    user_payment_details = (
        "\n".join(active_payments_list)
        if active_payments_list
        else "কোনো নম্বর সেট করা নেই"
    )

    preview_text = (
        f"💳 **পেমেন্ট কনফার্মেশন**\n"
        f"👤 শ্রমিক: [{target_name}](tg://user?id={target_uid}) (`{target_uid}`)\n"
        f"📂 ক্যাটাগরি: {cat_name} | 📅 তারিখ: `{sub_date}`\n"
        f"💰 বকেয়া: **৳{amount}**\n\n"
        f"📌 **নম্বর:**\n{user_payment_details}"
    )

    cat_key_matched = "ig_2fa"
    for k, v in settings["categories"].items():
      if v["name"] == cat_name:
        cat_key_matched = k
        break

    preview_kbd = [
        [
            InlineKeyboardButton(
                f"✅ Confirm Payment ({amount}৳)",
                callback_data=f"paynow_idx_{log_idx}",
            )
        ],
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান",
                callback_data=f"adm_paysublist_{cat_key_matched}",
            )
        ],
    ]

    await query.message.edit_text(
        preview_text,
        reply_markup=InlineKeyboardMarkup(preview_kbd),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("paynow_idx_"):
    if not is_admin(user_id):
      return
    log_idx = int(data.replace("paynow_idx_", ""))
    logs = load_logs()
    if log_idx >= len(logs):
      return

    l = logs[log_idx]
    l["is_paid"] = True
    save_logs(logs)

    target_uid = l["user_id"]
    target_name = l["user_name"]
    cat_name = l["category"]
    amount = l.get("earned_amount", 0.0)
    sub_date = l.get("date", "")
    valid_count = l.get("unique", 0)
    total_count = l.get("total", 0)
    rejected_count = l.get("rejected", 0)
    duplicates_count = l.get("duplicates", 0)

    ref_data = settings.get("referral", {})
    if ref_data.get("status", True) and ref_data.get(
        "commission_status", True
    ):
      comm_pct = float(ref_data.get("commission_percent", 5.0))
      referrals = load_referrals()
      target_uid_str = str(target_uid)

      for ref_by_str, r_info in referrals.get("referrers", {}).items():
        referred_list = r_info.get("refs", [])
        is_referred = any(
            r.get("uid") == target_uid_str for r in referred_list
        )
        if is_referred:
          commission_amt = (amount * comm_pct) / 100.0
          current_comm = float(r_info.get("commission_earned", 0.0))
          r_info["commission_earned"] = current_comm + commission_amt
          save_referrals(referrals)
          try:
            all_u = load_all_users()
            ref_name = all_u.get(ref_by_str, "Referrer")
            await context.bot.send_message(
                chat_id=int(ref_by_str),
                text=(
                    f"🎉 **প্রিয় {ref_name}, রেফারেল কমিশন প্রাপ্তি!**\n\nআপনার রেফার করা মেম্বার"
                    f" [{target_name}](tg://user?id={target_uid}) এর পেমেন্ট"
                    f" সম্পন্ন থাকায় আপনি পেয়েছেন কাজের ওপর **৳{commission_amt}**"
                    " কমিশন!"
                ),
                parse_mode="Markdown",
            )
          except:
            pass
          break

    cat_rate = 5.0
    for k, v in settings["categories"].items():
      if v["name"] == cat_name:
        cat_rate = v["price"]
        break

    target_uid_str = str(target_uid)
    user_p = user_payments.get(target_uid_str, {})
    p_info_str = "bKash / Nagad"
    for p_k, p_v in user_p.items():
      if p_v != "সেট করা হয়নি":
        p_info_str = f"{p_k.capitalize()} ({p_v})"
        break

    try:
      invoice_msg = (
          "🧾 OFFICIAL PAYMENT INVOICE 🧾\n"
          "━━━━━━━━━━━━━━━━━━━━━\n"
          f"👤 Worker Name: {target_name}\n"
          f"🆔 Worker ID: `{target_uid}`\n"
          "━━━━━━━━━━━━━━━━━━━━━\n"
          f"📂 Category: {cat_name}\n"
          f"📅 Submission Date: {sub_date}\n"
          f"🔢 Total Submitted: {total_count} Pcs\n"
          f"✅ Valid IDs: {valid_count} Pcs\n"
          f"❌ Rejected IDs: {rejected_count} Pcs\n"
          f"♻️ Duplicates: {duplicates_count} Pcs\n"
          f"💰 Rate per ID: {cat_rate}৳\n"
          f"💸 Total Paid Amount: ৳{amount}\n"
          "━━━━━━━━━━━━━━━━━━━━━\n"
          f"💳 Payment Method: {p_info_str}\n"
          "📌 Status: ✅ SUCCESSFULLY PAID"
      )
      await context.bot.send_message(
          chat_id=target_uid, text=invoice_msg, parse_mode="Markdown"
      )
    except Exception as e:
      print(f"Error sending invoice to user: {e}")

    await query.answer(
        "✅ সফলভাবে এই সাবমিশনের পেমেন্ট Paid করা হয়েছে!", show_alert=True
    )

    cat_key_matched = "ig_2fa"
    for k, v in settings["categories"].items():
      if v["name"] == cat_name:
        cat_key_matched = k
        break

    query.data = f"adm_paysublist_{cat_key_matched}"
    await button_callback(update, context)
    return

  elif data == "adm_global_pay_list":
    if not is_admin(user_id):
      return
    logs = load_logs()
    if not logs:
      await query.answer("⛔ কোনো সাবমিশন লগ পাওয়া যায়নি!", show_alert=True)
      return
    await query.answer()

    user_summary = {}
    for l in logs:
      uid = l["user_id"]
      name = l["user_name"]
      earned = l.get("earned_amount", 0.0)
      is_paid = l.get("is_paid", False)

      if uid not in user_summary:
        user_summary[uid] = {
            "name": name,
            "total": 0.0,
            "unpaid": 0.0,
            "has_unpaid": False,
        }

      user_summary[uid]["total"] += earned
      if not is_paid:
        user_summary[uid]["unpaid"] += earned
        user_summary[uid]["has_unpaid"] = True

    text = "👥 **সকল মেম্বারদের সার্বিক পেমেন্ট ও বকেয়া তালিকা:**\n\n"
    for uid, udata in user_summary.items():
      status_mark = (
          "🟢 সম্পূর্ণ পরিশোধিত (Paid)"
          if not udata["has_unpaid"]
          else f"🔴 বকেয়া: ৳{udata['unpaid']}"
      )
      user_link = f"[{udata['name']}](tg://user?id={uid})"
      text += (
          f"👤 **{user_link}** (`{uid}`)\n   - মোট আয়: ৳{udata['total']} |"
          f" {status_mark}\n\n"
      )

    keyboard = [
        [
            InlineKeyboardButton(
                "🔙 পেমেন্ট হাব", callback_data="adm_pay_report_hub"
            )
        ]
    ]
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown"
    )
    return

  elif data == "adm_download_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    dl_keyboard = []
    for key, cat in settings["categories"].items():
      dl_keyboard.append([
          InlineKeyboardButton(
              f"📁 {cat['name']} ফাইল ডাউনলোড", callback_data=f"adm_dl_{key}"
          )
      ])
    dl_keyboard.append([
        InlineKeyboardButton(
            "🔙 ডাউনলোড হাব", callback_data="adm_download_hub"
        )
    ])
    await query.message.edit_text(
        "📥 **সব সময়ের জমা দেওয়া আইডির মার্জড ফাইল ডাউনলোড**\n\nযে ক্যাটাগরির সব"
        " আইডি একসাথে ফাইল আকারে নিতে চান তাতে ক্লিক করুন:",
        reply_markup=InlineKeyboardMarkup(dl_keyboard),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_today_download_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    td_keyboard = []
    for key, cat in settings["categories"].items():
      td_keyboard.append([
          InlineKeyboardButton(
              f"📅 আজকের {cat['name']} ফাইল", callback_data=f"adm_toddl_{key}"
          )
      ])
    td_keyboard.append([
        InlineKeyboardButton(
            "🔙 ডাউনলোড হাব", callback_data="adm_download_hub"
        )
    ])
    await query.message.edit_text(
        "📅 **তারিখ অনুযায়ী ক্যাটাগরি ফাইল ডাউনলোড**\n\nআজকের জমা হওয়া নির্দিষ্ট"
        " ক্যাটাগরির ফাইল ডাউনলোড করতে ক্লিক করুন:",
        reply_markup=InlineKeyboardMarkup(td_keyboard),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_custom_date_download_menu":
    if not is_admin(user_id):
      return
    await query.answer()
    c_keyboard = []
    for key, cat in settings["categories"].items():
      c_keyboard.append([
          InlineKeyboardButton(
              f"🗓 {cat['name']} (কাস্টম তারিখ)",
              callback_data=f"adm_custdl_{key}",
          )
      ])
    c_keyboard.append([
        InlineKeyboardButton(
            "🔙 ডাউনলোড হাব", callback_data="adm_download_hub"
        )
    ])
    await query.message.edit_text(
        "🗓 **নির্দিষ্ট তারিখ অনুযায়ী ফাইল ডাউনলোড**\n\nযে ক্যাটাগরির ফাইল নিতে"
        " চান তা সিলেক্ট করুন, এরপর নির্দিষ্ট তারিখ চাওয়া হবে:",
        reply_markup=InlineKeyboardMarkup(c_keyboard),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("adm_custdl_"):
    if not is_admin(user_id):
      return
    await query.answer()
    cat_key = data.replace("adm_custdl_", "")
    if cat_key not in settings["categories"]:
      return

    context.user_data["waiting_for_custom_date_cat"] = cat_key
    cat_name = settings["categories"][cat_key]["name"]

    keyboard = [
        [
            InlineKeyboardButton(
                "🔙 ফিরে যান", callback_data="adm_custom_date_download_menu"
            )
        ]
    ]
    await query.message.edit_text(
        f"🗓 আপনি **{cat_name}** ক্যাটাগরির ফাইল ডাউনলোড করতে চান。\n\nদয়া করে"
        " নির্দিষ্ট তারিখটি এই ফরম্যাটে চ্যাটে লিখে পাঠান: `DD-MM-YYYY`\n(উদাহরণস্বরূপ:"
        " `10-10-2026`)",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("adm_toddl_"):
    if not is_admin(user_id):
      return
    cat_key = data.replace("adm_toddl_", "")
    if cat_key not in settings["categories"]:
      return
    cat_name = settings["categories"][cat_key]["name"]
    today_str = datetime.datetime.now(TZ).strftime("%Y-%m-%d")

    logs = load_logs()
    today_logs = [
        l
        for l in logs
        if l["category"] == cat_name and l["date"].startswith(today_str)
    ]

    if not today_logs:
      await query.answer(
          f"⛔ আজ {cat_name} ক্যাটাগরিতে কোনো এন্ট্রি জমা পড়েনি!",
          show_alert=True,
      )
      return
    await query.answer()

    today_valid_ids = set()
    for l in today_logs:
      for uid_val in l.get("submitted_ids_list", []):
        today_valid_ids.add(uid_val)

    all_rows_db = load_all_rows_db()
    cat_rows = all_rows_db.get(cat_name, [])

    today_rows = []
    for r in cat_rows:
      if r and len(r) > 0:
        val_str = str(r[0]).strip()
        if val_str.endswith(".0"):
          val_str = val_str[:-2]
        if val_str in today_valid_ids:
          today_rows.append(r)

    if not today_rows:
      await query.answer(
          f"⛔ আজ {cat_name} ক্যাটাগরিতে কোনো এন্ট্রি পাওয়া যায়নি!",
          show_alert=True,
      )
      return

    merged_wb = openpyxl.Workbook()
    merged_ws = merged_wb.active
    merged_ws.title = "Today_Export"

    for r_data in today_rows:
      merged_ws.append(r_data)

    current_date_str = datetime.datetime.now(TZ).strftime("%d-%m-%Y")
    safe_name = cat_name.replace(" ", "_")
    file_name = f"Today_{safe_name}_{current_date_str}.xlsx"
    merged_wb.save(file_name)

    with open(file_name, "rb") as f:
      await context.bot.send_document(
          chat_id=user_id,
          document=f,
          caption=(
              f"📅 **{cat_name}** ক্যাটাগরির আজকের ({today_str}) জমা হওয়া সমস্ত আইডি"
              " সংযুক্ত করা হলো! 🚀"
          ),
          parse_mode="Markdown",
      )
    os.remove(file_name)
    return

  elif data.startswith("adm_dl_"):
    if not is_admin(user_id):
      return
    cat_key = data.replace("adm_dl_", "")
    if cat_key not in settings["categories"]:
      return
    cat_name = settings["categories"][cat_key]["name"]
    all_rows_db = load_all_rows_db()
    cat_rows = all_rows_db.get(cat_name, [])

    if not cat_rows:
      await query.answer(
          f"⛔ {cat_name} ক্যাটাগরিতে এখনো কোনো আইডি জমা পড়েনি!",
          show_alert=True,
      )
      return
    await query.answer()

    merged_wb = openpyxl.Workbook()
    merged_ws = merged_wb.active
    merged_ws.title = cat_name[:30]

    for r_data in cat_rows:
      merged_ws.append(r_data)

    current_date_str = datetime.datetime.now(TZ).strftime("%d-%m-%Y")
    safe_name = cat_name.replace(" ", "_")
    merged_file_name = f"Merged_All_{safe_name}_{current_date_str}.xlsx"
    merged_wb.save(merged_file_name)

    with open(merged_file_name, "rb") as f:
      await context.bot.send_document(
          chat_id=user_id,
          document=f,
          caption=(
              f"📁 **{cat_name}** ক্যাটাগরির আজ পর্যন্ত জমা হওয়া সমস্ত ইউনিক আইডি"
              " একসাথে সংযুক্ত করা হলো! 🚀"
          ),
          parse_mode="Markdown",
      )
    os.remove(merged_file_name)
    return

  elif data == "adm_set_prices":
    if not is_admin(user_id):
      return
    await query.answer()
    await adm_set_prices_render(query, settings)
    return

  elif data.startswith("adm_toggle_show_"):
    if not is_admin(user_id):
      return
    cat_key = data.replace("adm_toggle_show_", "")
    if cat_key in settings["categories"]:
      current = settings["categories"][cat_key].get("show_price", True)
      settings["categories"][cat_key]["show_price"] = not current
      save_settings(settings)

    await query.answer("স্ট্যাটাস পরিবর্তন হয়েছে!", show_alert=True)
    await adm_set_prices_render(query, settings)
    return

  elif data.startswith("adm_price_"):
    if not is_admin(user_id):
      return
    await query.answer()
    cat_key = data.replace("adm_price_", "")
    context.user_data["waiting_for_new_price_cat"] = cat_key
    cat_name = settings["categories"][cat_key]["name"]

    keyboard = [
        [InlineKeyboardButton("🔙 ফিরে যান", callback_data="adm_set_prices")]
    ]
    await query.message.edit_text(
        f"✍ আপনি **{cat_name}** ক্যাটাগরির বর্তমান রেট পরিবর্তন করতে চান。\n\nদয়া"
        " করে নতুন মূল্য (সংখ্যায়, যেমন: `6.5` বা `10`) চ্যাটে লিখে পাঠান:",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown",
    )
    return

  elif data == "adm_toggle_payments":
    if not is_admin(user_id):
      return
    await query.answer()
    await adm_toggle_payments_render(query, settings)
    return

  elif data.startswith("adm_toggle_p_"):
    if not is_admin(user_id):
      return
    p_key = data.replace("adm_toggle_p_", "")
    settings["payments"][p_key]["status"] = not settings["payments"][p_key][
        "status"
    ]
    save_settings(settings)

    await query.answer("পেমেন্ট গেটওয়ে স্ট্যাটাস পরিবর্তন হয়েছে!", show_alert=True)
    await adm_toggle_payments_render(query, settings)
    return

  elif data == "payment_menu":
    await query.answer()
    user_p = user_payments.get(user_id_str, {})
    p_keyboard = []
    for p_key, p_val in settings["payments"].items():
      if p_val["status"]:
        status_icon = "🟢"
        saved_num = user_p.get(p_key, "সেট করা হয়নি")
        btn_text = f"{status_icon} {p_val['name']} ({saved_num})"
      else:
        btn_text = f"🔴 {p_val['name']} (Closed)"

      p_keyboard.append(
          [InlineKeyboardButton(btn_text, callback_data=f"set_p_{p_key}")]
      )
    p_keyboard.append(
        [InlineKeyboardButton("🔙 মূল মেনুতে ফিরে যান", callback_data="back_home")]
    )

    pay_msg = (
        f"💳 **পেমেন্ট নম্বর ম্যানেজমেন্ট ({user.first_name})**\n\nআপনার পেমেন্ট নম্বর সেট বা"
        " পরিবর্তন করতে নিচের মাধ্যমটিতে ক্লিক করুন:\n(নম্বরগুলোতে ক্লিক করলেই"
        " কপি হয়ে যাবে)"
    )
    await query.message.edit_text(
        pay_msg,
        reply_markup=InlineKeyboardMarkup(p_keyboard),
        parse_mode="Markdown",
    )
    return

  elif data.startswith("set_p_"):
    p_key = data.replace("set_p_", "")
    p_info = settings["payments"].get(p_key, {})

    if not p_info.get("status", True) and not is_admin(user_id):
      await query.answer(
          f"দুঃখিত {user.first_name}, এই পেমেন্ট মাধ্যমটি বর্তমানে বন্ধ রয়েছে! ❌", show_alert=True
      )
      return

    await query.answer()
    context.user_data["waiting_for_payment_method"] = p_key

    method_names = {
        "bkash": "Bkash",
        "nagad": "Nagad",
        "binance": "Binance",
        "rocket": "Rocket",
    }
    m_name = method_names.get(p_key, p_key)

    keyboard = [
        [InlineKeyboardButton("🔙 ফিরে যান", callback_data="payment_menu")]
    ]

    if p_key == "binance":
      prompt_text = (
          f"✍ প্রিয় **{user.first_name}**, আপনি আপনার **{m_name}** নম্বর বা পেমেন্ট আইডি সেট করতে চান。\n\nদয়া"
          " করে আপনার বিনান্স পে আইডি বা বাইন্যান্স ইমেইল/নম্বরটি চ্যাটে লিখে"
          " সেন্ড করুন:"
      )
    else:
      prompt_text = (
          f"✍ প্রিয় **{user.first_name}**, আপনি আপনার **{m_name}** নম্বরটি সেট করতে চান。\n\nদয়া করে সঠিক"
          " **১১ সংখ্যার বাংলাদেশি মোবাইল নম্বর** দিন যা `01` দিয়ে শুরু:"
      )

    await query.message.edit_text(
        prompt_text,
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown",
    )
    return

  elif data == "my_payment_tracking":
    await query.answer()
    logs = load_logs()
    user_logs = [l for l in logs if l["user_id"] == user_id]
    total_earned = sum(l.get("earned_amount", 0.0) for l in user_logs)
    paid_earned = sum(
        l.get("earned_amount", 0.0) for l in user_logs if l.get("is_paid", False)
    )
    due_earned = total_earned - paid_earned

    user_p = user_payments.get(user_id_str, {})
    active_payments = []
    for p_k, p_v in user_p.items():
      if p_v != "সেট করা হয়নি":
        active_payments.append(f"🔹 {p_k.capitalize()}: `{p_v}`")
    payment_info_str = (
        "\n".join(active_payments)
        if active_payments
        else "কোনো নম্বর সেট করা নেই"
    )

    text = (
        f"💳 **প্রিয় {user.first_name}, আপনার পেমেন্ট ও আয়ের তথ্য**:\n\n"
        f"💰 সর্বমোট আয়: **৳{total_earned}**\n"
        f"🟢 পরিশোধিত (Paid): **৳{paid_earned}**\n"
        f"🔴 বকেয়া (Unpaid): **৳{due_earned}**\n\n"
        f"📌 **আপনার পেমেন্ট নম্বরসমূহ:**\n{payment_info_str}\n\n"
        "নিচের এক্সেল হিস্ট্রি থেকে প্রতিটি কাজের Paid/Unpaid স্ট্যাটাস দেখতে"
        " পারবেন।"
    )
    keyboard = [[InlineKeyboardButton("🔙 মূল মেনু", callback_data="back_home")]]
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown"
    )
    return

  elif data == "export_my_history":
    logs = load_logs()
    user_logs = [l for l in logs if l["user_id"] == user_id]
    if not user_logs:
      await query.answer(f"⛔ {user.first_name}, আপনার কোনো কাজের হিস্ট্রি পাওয়া যায়নি!", show_alert=True)
      return
    await query.answer()

    export_wb = openpyxl.Workbook()
    export_ws = export_wb.active
    export_ws.title = "My Submission History"

    export_ws.append([
        "তারিখ ও সময়",
        "ক্যাটাগরি",
        "মোট ফাইল",
        "ভ্যালিড",
        "রিজেক্ট",
        "ডুপ্লিকেট",
        "মোট টাকা (৳)",
        "রিপোর্ট স্ট্যাটাস",
        "পেমেন্ট স্ট্যাটাস",
    ])
    for l in user_logs:
      p_status = "Paid" if l.get("is_paid", False) else "Unpaid"
      rep_status = (
          "Reported" if l.get("has_report", False) else "Pending Report"
      )
      export_ws.append([
          l["date"],
          l["category"],
          l["total"],
          l["unique"],
          l.get("rejected", 0),
          l["duplicates"],
          l.get("earned_amount", 0),
          rep_status,
          p_status,
      ])

    file_name = f"My_History_{user_id}.xlsx"
    export_wb.save(file_name)

    with open(file_name, "rb") as f:
      await context.bot.send_document(
          chat_id=user_id,
          document=f,
          caption=(
              f"📁 **প্রিয় {user.first_name}, আপনার সমস্ত কাজের সাবমিশন ও Paid/Unpaid হিস্ট্রি ফাইল হিসেবে"
              " সংযুক্ত করা হলো।** ✨"
          ),
          parse_mode="Markdown",
      )
    os.remove(file_name)
    return

  elif data == "back_home":
    await query.answer()
    await start(update, context)
    return


async def adm_toggle_panel_render(query, settings):
  toggle_keyboard = []
  for key, cat in settings["categories"].items():
    status_icon = "🟢 चालू" if cat["status"] else "⛔ বন্ধ"
    toggle_keyboard.append([
        InlineKeyboardButton(
            f"{cat['name']}: {status_icon}",
            callback_data=f"adm_tglcat_{key}",
        )
    ])
  toggle_keyboard.append([
      InlineKeyboardButton("🔙 ক্যাটাগরি মেনু", callback_data="adm_cat_mgmt_hub")
  ])
  try:
    await query.message.edit_text(
        "🔄 **🟢 ক্যাটাগরি চালু/বন্ধ প্যানেল**\n\nযে ক্যাটাগরি অন বা অফ করতে চান"
        " তাতে ক্লিক করুন:",
        reply_markup=InlineKeyboardMarkup(toggle_keyboard),
        parse_mode="Markdown",
    )
  except:
    pass


async def adm_set_prices_render(query, settings):
  p_keyboard = []
  for key, cat in settings["categories"].items():
    disp_status = (
        "👀 প্রাইস দৃশ্যমান" if cat.get("show_price", True) else "🙈 প্রাইস লুকানো"
    )
    p_keyboard.append([
        InlineKeyboardButton(
            f"✏️ {cat['name']} ({cat['price']}৳)",
            callback_data=f"adm_price_{key}",
        ),
        InlineKeyboardButton(
            disp_status, callback_data=f"adm_toggle_show_{key}"
        ),
    ])
  p_keyboard.append([
      InlineKeyboardButton("🔙 ক্যাটাগরি মেনু", callback_data="adm_cat_mgmt_hub")
  ])
  try:
    await query.message.edit_text(
        "💰 **ক্যাটাগরি প্রাইস ও ডিসপ্লে ম্যানেজমেন্ট**\n\nরেট পরিবর্তন করতে বা"
        " প্রাইস শো অন/অফ করতে নিচের বাটনগুলোতে ক্লিক করুন:",
        reply_markup=InlineKeyboardMarkup(p_keyboard),
        parse_mode="Markdown",
    )
  except:
    pass


async def adm_toggle_payments_render(query, settings):
  pay_keyboard = []
  for p_key, p_val in settings["payments"].items():
    status_icon = "🟢 চালু" if p_val["status"] else "⛔ বন্ধ"
    pay_keyboard.append([
        InlineKeyboardButton(
            f"{p_val['name']}: {status_icon}",
            callback_data=f"adm_toggle_p_{p_key}",
        )
    ])
  pay_keyboard.append([
      InlineKeyboardButton("🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu")
  ])
  try:
    await query.message.edit_text(
        "🎛 **পেমেন্ট গেটওয়ে কন্ট্রোল প্যানেল**\n\nযে মাধ্যমটি অন বা অফ করতে চান তাতে"
        " ক্লিক করুন:",
        reply_markup=InlineKeyboardMarkup(pay_keyboard),
        parse_mode="Markdown",
    )
  except:
    pass


async def adm_referral_menu_render(query, settings):
  ref_data = settings.get("referral", {})
  status_str = "🟢 চালু" if ref_data.get("status", True) else "⛔ বন্ধ"
  comm_status_str = (
      "🟢 চালু" if ref_data.get("commission_status", True) else "⛔ বন্ধ"
  )
  text = (
      "🎁 **রেফারেল সিস্টেম ও কমিশন ম্যানেজমেন্ট**\n\n"
      f"• রেফারেল স্ট্যাটাস: `{status_str}`\n"
      f"• প্রতি রেফার বোনাস: `৳{ref_data.get('bonus_amount', 10.0)}`\n"
      f"• রেফারেল কমিশন (%): `{ref_data.get('commission_percent', 5.0)}%`\n"
      f"• কমিশন স্ট্যাটাস: `{comm_status_str}`\n"
      f"• ন্যূনতম রেফার শর্ত: `{ref_data.get('min_refs_to_withdraw', 10)} টি`\n"
      f"• ন্যূনতম উইথড্র টাকা: `৳{ref_data.get('min_amount_to_withdraw', 50.0)}`\n\n"
      "*(নোট: রেফারেল বোনাস এবং কাজের কমিশন সম্পূর্ণ আলাদাভাবে হিসাব হয়ে স্বয়ংক্রিয়ভাবে ব্যালেন্স যোগ হয়)*"
  )
  kbd = [
      [
          InlineKeyboardButton(
              f"টগল রেফারেল: {status_str}", callback_data="adm_ref_toggle_status"
          )
      ],
      [
          InlineKeyboardButton(
              f"টগল কমিশন (%): {comm_status_str}",
              callback_data="adm_ref_toggle_comm",
          )
      ],
      [
          InlineKeyboardButton(
              "✏️ কমিশন পার্সেন্ট (%)", callback_data="adm_ref_set_comm_pct"
          ),
          InlineKeyboardButton(
              "✏️ বোনাস অ্যামাউন্ট", callback_data="adm_ref_set_bonus"
          ),
      ],
      [
          InlineKeyboardButton(
              "✏️ ন্যূনতম রেফার শর্ত", callback_data="adm_ref_set_min_refs"
          ),
          InlineKeyboardButton(
              "✏️ ন্যূনতম উইথড্র টাকা", callback_data="adm_ref_set_min_amt"
          ),
      ],
      [InlineKeyboardButton("🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu")],
  ]
  try:
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
  except:
    pass


async def adm_wallet_mgmt_render(query, settings):
  wallets = settings.get("wallets", {})
  text = "🌐 **ক্রিপ্টো ও ওয়ালেট অ্যাড্রেস ম্যানেজমেন্ট:**\n\n"
  kbd = []
  for w_name, w_info in wallets.items():
    if isinstance(w_info, str):
      w_desc = ""
      w_addr = w_info
    else:
      w_desc = w_info.get("desc", "")
      w_addr = w_info.get("address", "")
    text += f"• **{w_name}** ({w_desc}): `{w_addr}`\n"
    kbd.append([
        InlineKeyboardButton(
            f"❌ ডিলিট: {w_name}", callback_data=f"adm_del_wallet_{w_name}"
        )
    ])
  kbd.append([
      InlineKeyboardButton(
          "➕ নতুন ওয়ালেট মাধ্যম যোগ করুন", callback_data="adm_add_wallet_prompt"
      )
  ])
  kbd.append([
      InlineKeyboardButton("🔙 অ্যাডমিন মেনু", callback_data="admin_panel_menu")
  ])
  try:
    await query.message.edit_text(
        text, reply_markup=InlineKeyboardMarkup(kbd), parse_mode="Markdown"
    )
  except:
    pass


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
  user = update.effective_user
  bot = context.bot
  if user.id in load_blocked_users():
    return

  settings = load_settings()

  if context.user_data.get("waiting_for_support_msg"):
    context.user_data["waiting_for_support_msg"] = None
    photo_file = update.message.photo[-1]
    caption = update.message.caption or "কোনো বিবরণ দেওয়া হয়নি"

    with_ch_raw = settings.get("channels", {}).get(
        "withdrawal_channel", DEFAULT_ADMIN_ID
    )
    with_ch = parse_chat_id(with_ch_raw)

    user_link = f"[{user.first_name}](tg://user?id={user.id})"
    forward_caption = (
        "📩 **নতুন সাপোর্ট মেসেজ (ছবি সহ)!**\n\n"
        f"👤 ইউজার: {user_link}\n"
        f"🆔 আইডি: `{user.id}`\n\n"
        f"💬 সমস্যা / বিবরণ:\n{caption}"
    )

    try:
      await bot.send_photo(
          chat_id=with_ch,
          photo=photo_file.file_id,
          caption=forward_caption,
          parse_mode="Markdown",
      )
      await update.message.reply_text(
          f"✅ প্রিয় {user.first_name}, আপনার সাপোর্ট মেসেজ ও ছবি সফলভাবে অ্যাডমিনের কাছে পাঠানো হয়েছে! শীঘ্রই আপনার সাথে যোগাযোগ করা হবে। 🙏",
          parse_mode="Markdown"
      )
    except Exception as e:
      await update.message.reply_text("❌ মেসেজ পাঠাতে সমস্যা হয়েছে: " + str(e))
    return

  user_payments = load_user_payments()
  user_p = user_payments.get(str(user.id), {})
  has_valid_payment = any(
      p_v != "সেট করা হয়নি" for p_k, p_v in user_p.items()
  )
  if not has_valid_payment and not is_admin(user.id):
    await update.message.reply_text(
        f"⛔ {user.first_name}, প্রুফ পাঠানোর পূর্বে অন্তত একটি পেমেন্ট নম্বর সেট করা বাধ্যতামূলক!"
        " দয়া করে বট মেনু থেকে **⚙ পেমেন্ট সেটিংস** এ গিয়ে আপনার পেমেন্ট নম্বরটি"
        " সেট করুন।",
        parse_mode="Markdown"
    )
    return

  crypto_wallet = context.user_data.get("waiting_for_crypto_proof")
  if not crypto_wallet:
    return

  caption = update.message.caption or ""
  if not caption.strip():
    await update.message.reply_text(
        f"⛔ {user.first_name}, ছবির সাথে ট্রানজেকশন আইডি (TxHash বা Pay ID / Order ID) ক্যাপশনে লিখে দিন অথবা সরাসরি টেক্সট মেসেজ পাঠান।"
    )
    return

  update.message.text = caption
  await handle_text_message(update, context)


async def handle_text_message(
    update: Update, context: ContextTypes.DEFAULT_TYPE
):
  user = update.effective_user
  user_id = user.id
  if user_id in load_blocked_users():
    return

  settings = load_settings()

  if is_admin(user_id) and context.user_data.get("waiting_for_edit_msg_key"):
    msg_key = context.user_data.get("waiting_for_edit_msg_key")
    new_txt = update.message.text.strip()
    context.user_data["waiting_for_edit_msg_key"] = None
    settings.setdefault("custom_messages", {})[msg_key] = new_txt
    save_settings(settings)
    await update.message.reply_text(
        f"✅ মেসেজটি সফলভাবে আপডেট করা হয়েছে!\n\n`{new_txt}`",
        parse_mode="Markdown"
    )
    return

  if is_admin(user_id) and context.user_data.get("waiting_for_binance_api_key"):
    key_val = update.message.text.strip()
    context.user_data["waiting_for_binance_api_key"] = None
    settings.setdefault("binance_api", {})["api_key"] = key_val
    save_settings(settings)
    await update.message.reply_text(
        "✅ Binance API Key সফলভাবে সেভ করা হয়েছে!"
    )
    return

  if is_admin(user_id) and context.user_data.get(
      "waiting_for_binance_api_secret"
  ):
    secret_val = update.message.text.strip()
    context.user_data["waiting_for_binance_api_secret"] = None
    settings.setdefault("binance_api", {})["api_secret"] = secret_val
    save_settings(settings)
    await update.message.reply_text(
        "✅ Binance API Secret Key সফলভাবে সেভ করা হয়েছে!"
    )
    return

  if context.user_data.get("waiting_for_support_msg"):
    context.user_data["waiting_for_support_msg"] = None
    text_content = update.message.text.strip()

    with_ch_raw = settings.get("channels", {}).get(
        "withdrawal_channel", DEFAULT_ADMIN_ID
    )
    with_ch = parse_chat_id(with_ch_raw)

    user_link = f"[{user.first_name}](tg://user?id={user_id})"
    forward_text = (
        "📩 **নতুন সাপোর্ট মেসেজ!**\n\n"
        f"👤 ইউজার: {user_link}\n"
        f"🆔 আইডি: `{user_id}`\n\n"
        f"💬 সমস্যা / বিবরণ:\n`{text_content}`"
    )

    try:
      await context.bot.send_message(
          chat_id=with_ch, text=forward_text, parse_mode="Markdown"
      )
      await update.message.reply_text(
          f"✅ প্রিয় {user.first_name}, আপনার সাপোর্ট মেসেজ সফলভাবে অ্যাডমিনের কাছে পাঠানো হয়েছে! শীঘ্রই আপনার সাথে যোগাযোগ করা হবে। 🙏",
          parse_mode="Markdown"
      )
    except Exception as e:
      await update.message.reply_text("❌ মেসেজ পাঠাতে সমস্যা হয়েছে: " + str(e))
    return

  crypto_wallet = context.user_data.get("waiting_for_crypto_proof")
  if crypto_wallet:
    user_payments = load_user_payments()
    user_p = user_payments.get(str(user_id), {})
    has_valid_payment = any(
        p_v != "সেট করা হয়নি" for p_k, p_v in user_p.items()
    )
    if not has_valid_payment and not is_admin(user_id):
      await update.message.reply_text(
          f"⛔ {user.first_name}, প্রুফ পাঠানোর পূর্বে অন্তত একটি পেমেন্ট নম্বর সেট করা বাধ্যতামূলক!"
      )
      return

    tx_id_input = update.message.text.strip()

    used_txids = load_used_txids()
    clean_tx_check = tx_id_input.lower()
    if clean_tx_check in used_txids:
      dup_tpl = settings.get("custom_messages", {}).get(
          "crypto_duplicate_warning", "⚠️ **সতর্কতা {name}!** এই প্রুফটি ইতিপূর্বেই ব্যবহার করা হয়েছে!"
      )
      await update.message.reply_text(
          dup_tpl.replace("{name}", user.first_name),
          parse_mode="Markdown"
      )

      with_ch_raw = settings.get("channels", {}).get(
          "withdrawal_channel", DEFAULT_ADMIN_ID
      )
      with_ch = parse_chat_id(with_ch_raw)
      user_link = f"[{user.first_name}](tg://user?id={user_id})"
      warning_alert_text = (
          "🚨 **সতর্কতা: ডুপ্লিকেট ক্রিপ্টো প্রুফ ব্যবহারের চেষ্টা!**\n\n"
          f"👤 ইউজার: {user_link}\n"
          f"🆔 আইডি: `{user_id}`\n"
          f"🔹 ওয়ালেট মাধ্যম: `{crypto_wallet}`\n"
          f"⚠️ রিইউজ করা TxID / Order ID: `{tx_id_input}`\n\n"
          "*(এই ইউজার ইতিমধ্যে ব্যবহৃত ট্রানজেকশন আইডি দিয়ে আবার উইথড্র নেওয়ার চেষ্টা করেছে)*"
      )
      try:
        await context.bot.send_message(
            chat_id=with_ch, text=warning_alert_text, parse_mode="Markdown"
        )
      except:
        pass
      return

    v_progress_tpl = settings.get("custom_messages", {}).get(
        "crypto_verify_msg", "🔍 **Binance API-এর মাধ্যমে ট্রানজেকশন যাচাই করা হচ্ছে...** দয়া করে অপেক্ষা করুন ⏳"
    )
    await update.message.reply_text(
        v_progress_tpl.replace("{name}", user.first_name),
        parse_mode="Markdown"
    )

    binance_api_cfg = settings.get("binance_api", {})
    api_key = binance_api_cfg.get("api_key", "")
    api_secret = binance_api_cfg.get("api_secret", "")

    verified = False
    tx_details = None
    error_reason = ""

    if not api_key or not api_secret:
      verified = True
      error_reason = "API keys not configured by admin, forwarded for manual check."
      received_amount = "Manual Check"
    else:
      is_pay = tx_id_input.isdigit() or "pay" in crypto_wallet.lower()

      if is_pay:
        verified, tx_details = verify_binance_pay(
            tx_id_input, api_key, api_secret
        )
      else:
        verified, tx_details = verify_binance_deposit(
            tx_id_input, api_key, api_secret
        )

      if not verified and not is_pay:
        v_pay, d_pay = verify_binance_pay(tx_id_input, api_key, api_secret)
        if v_pay:
          verified, tx_details = v_pay, d_pay

      if not verified:
        error_reason = (
            str(tx_details) if tx_details else "Transaction ID not found"
        )
      else:
        received_amount = tx_details.get("amount", "Unknown Amount")

    if not verified:
      await update.message.reply_text(
          f"❌ **Binance API ভেরিফিকেশন ব্যর্থ হয়েছে!**\n\nকারণ: `{error_reason}`\n\nদয়া করে সঠিক TxHash বা Pay ID / Order ID দিয়ে আবার চেষ্টা করুন।",
          parse_mode="Markdown"
      )
      return

    used_txids.append(clean_tx_check)
    save_used_txids(used_txids)

    crypto_w_data = load_crypto_withdrawals()
    user_w_list = crypto_w_data.setdefault(str(user_id), [])
    user_w_list.append({
        "wallet": crypto_wallet,
        "amount_str": received_amount,
        "tx_id": tx_id_input,
        "date": datetime.datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
    })
    save_crypto_withdrawals(crypto_w_data)

    context.user_data["waiting_for_crypto_proof"] = None
    with_ch_raw = settings.get("channels", {}).get(
        "withdrawal_channel", DEFAULT_ADMIN_ID
    )
    with_ch = parse_chat_id(with_ch_raw)

    active_payments = []
    for p_k, p_v in user_p.items():
      if p_v != "সেট করা হয়নি":
        active_payments.append(f"• {p_k.capitalize()}: `{p_v}`")
    p_info_str = (
        "\n".join(active_payments) if active_payments else "কোনো নম্বর সেট করা নেই"
    )

    user_link = f"[{user.first_name}](tg://user?id={user_id})"
    forward_text = (
        "🌐 **অটো-ভেরিফাইড ক্রিপ্টো উইথড্র রিকোয়েস্ট!** ✅\n\n"
        f"👤 ইউজার: {user_link}\n"
        f"🆔 আইডি: `{user_id}`\n"
        f"🔹 ওয়ালেট মাধ্যম: `{crypto_wallet}`\n"
        f"💵 **প্রাপ্ত ক্রিপ্টো পরিমাণ:** `{received_amount}`\n"
        f"🔍 TxID / Pay ID / Order ID: `{tx_id_input}`\n"
        f"🟢 Binance API Status: `SUCCESSFULLY VERIFIED`\n\n"
        f"📌 **ইউজারের সেভ করা পেমেন্ট তথ্য:**\n{p_info_str}"
    )

    wd_kbd = [
        [
            InlineKeyboardButton(
                "✅ Pay Done (Accept)",
                callback_data=f"adm_crypto_done_{user_id}_{crypto_wallet}",
            )
        ]
    ]

    try:
      await context.bot.send_message(
          chat_id=with_ch,
          text=forward_text,
          reply_markup=InlineKeyboardMarkup(wd_kbd),
          parse_mode="Markdown",
      )
      succ_tpl = settings.get("custom_messages", {}).get(
          "crypto_success_msg", "🎉 **অভিনন্দন {name}!** আপনার **{amount}** এর ট্রানজেকশনটি সফলভাবে ভেরিফাই হয়েছে!"
      )
      formatted_succ = succ_tpl.replace("{name}", user.first_name).replace("{amount}", received_amount)
      await update.message.reply_text(formatted_succ, parse_mode="Markdown")
    except Exception as e:
      await update.message.reply_text("❌ রিকোয়েস্ট পাঠাতে সমস্যা হয়েছে: " + str(e))
    return

  if is_admin(user_id) and context.user_data.get("waiting_for_join_channel"):
    ch_val = update.message.text.strip()
    context.user_data["waiting_for_join_channel"] = None
    settings.setdefault("channels", {})["join_channel"] = ch_val
    save_settings(settings)
    await update.message.reply_text(
        f"✅ নতুন জয়েন চ্যানেল সফলভাবে সেট করা হয়েছে: `{ch_val}`",
        parse_mode="Markdown",
    )
    return

  if is_admin(user_id) and context.user_data.get("waiting_for_sub_channel"):
    ch_val = update.message.text.strip()
    context.user_data["waiting_for_sub_channel"] = None
    settings.setdefault("channels", {})["submission_channel"] = ch_val
    save_settings(settings)
    await update.message.reply_text(
        f"✅ নতুন সাবমিশন চ্যানেল সফলভাবে সেট করা হয়েছে: `{ch_val}`",
        parse_mode="Markdown",
    )
    return

  if is_admin(user_id) and context.user_data.get("waiting_for_with_channel"):
    ch_val = update.message.text.strip()
    context.user_data["waiting_for_with_channel"] = None
    settings.setdefault("channels", {})["withdrawal_channel"] = ch_val
    save_settings(settings)
    await update.message.reply_text(
        f"✅ নতুন উইথড্র চ্যানেল সফলভাবে সেট করা হয়েছে: `{ch_val}`",
        parse_mode="Markdown",
    )
    return

  if is_admin(user_id) and context.user_data.get("waiting_for_support_link"):
    sup_val = update.message.text.strip()
    context.user_data["waiting_for_support_link"] = None
    settings.setdefault("channels", {})["support_link"] = sup_val
    save_settings(settings)
    await update.message.reply_text(
        f"✅ সফলভাবে নতুন সাপোর্ট আইডি / লিংক সেট করা হয়েছে: `{sup_val}`",
        parse_mode="Markdown",
    )
    return

  if is_admin(user_id) and context.user_data.get("waiting_for_new_admin_id"):
    try:
      new_adm_id = int(update.message.text.strip())
      admins = load_admins()
      if new_adm_id not in admins:
        admins.append(new_adm_id)
        save_admins(admins)
      context.user_data["waiting_for_new_admin_id"] = None
      await update.message.reply_text(
          f"👑 সফলভাবে নতুন অ্যাডমিন আইডি `{new_adm_id}` যুক্ত করা হয়েছে!",
          parse_mode="Markdown",
      )
      return
    except:
      await update.message.reply_text(
          "⛔ সঠিক সংখ্যায় টেলিগ্রাম ইউজার আইডি দিন।"
      )
      return

  if is_admin(user_id) and context.user_data.get("waiting_for_comm_percent"):
    try:
      pct = float(update.message.text.strip())
      settings["referral"]["commission_percent"] = pct
      save_settings(settings)
      context.user_data["waiting_for_comm_percent"] = None
      await update.message.reply_text(
          f"✅ সফলভাবে রেফারেল কমিশন পার্সেন্ট `{pct}%` সেট করা হয়েছে!"
      )
      return
    except:
      await update.message.reply_text("⛔ সঠিক সংখ্যা লিখুন (যেমন: `5`)।")
      return

  if is_admin(user_id) and context.user_data.get("waiting_for_new_cat_key"):
    c_key = update.message.text.strip().lower().replace(" ", "_")
    context.user_data["waiting_for_new_cat_key"] = None
    context.user_data["waiting_for_new_cat_name"] = c_key
    await update.message.reply_text(
        f"✍ ধাপ ২: এখন ক্যাটাগরির সুন্দর ডিসপ্লে নাম (যেমন: `Gmail 2FA`)"
        " লিখে পাঠান:"
    )
    return

  if is_admin(user_id) and context.user_data.get("waiting_for_new_cat_name"):
    c_name = update.message.text.strip()
    c_key = context.user_data.get("waiting_for_new_cat_name")
    context.user_data["waiting_for_new_cat_name"] = None
    context.user_data["waiting_for_new_cat_price"] = {
        "key": c_key,
        "name": c_name,
    }
    await update.message.reply_text(
        f"✍ ধাপ ৩: এখন **{c_name}** ক্যাটাগরির প্রতি আইডির রেট (যেমন: `5.0` বা"
        " `10`) সংখ্যায় লিখে পাঠান:"
    )
    return

  if is_admin(user_id) and context.user_data.get("waiting_for_new_cat_price"):
    info = context.user_data.get("waiting_for_new_cat_price")
    context.user_data["waiting_for_new_cat_price"] = None
    try:
      c_price = float(update.message.text.strip())
      c_key = info["key"]
      c_name = info["name"]

      settings.setdefault("categories", {})[c_key] = {
          "name": c_name,
          "status": True,
          "price": c_price,
          "show_price": True,
          "popup_closed": (
              f"দুঃখিত {{name}}, বর্তমানে {c_name} এর ফাইল জমা নেওয়া বন্ধ রয়েছে! ❌"
          ),
          "popup_shift": f"দুঃখিত {{name}}, বর্তমানে {c_name} শিফটের বাইরে রয়েছে! ⏰",
      }
      save_settings(settings)
      await update.message.reply_text(
          f"✅ সফলভাবে নতুন ক্যাটাগরি **{c_name}** ({c_price}৳ রেটসহ) যুক্ত করা"
          " হয়েছে! 🎉\n\nএখন মেম্বাররা `/start` দিলেই এই ক্যাটাগরি দেখতে পাবে।",
          parse_mode="Markdown",
      )
    except Exception as e:
      await update.message.reply_text(f"❌ ভুল মূল্য দিয়েছেন! আবার চেষ্টা করুন: {e}")
    return

  if is_admin(user_id) and context.user_data.get(
      "waiting_for_add_wallet_name_step"
  ):
    w_name = update.message.text.strip()
    context.user_data["waiting_for_add_wallet_name_step"] = None
    context.user_data["temp_wallet_name"] = w_name
    context.user_data["waiting_for_add_wallet_desc_step"] = True
    await update.message.reply_text(
        f"✍ এখন **{w_name}** এর সংক্ষিপ্ত বিবরণ (যেমন: `ইউএসডিটি নেটওয়ার্ক বা পে"
        " আইডি`) লিখে পাঠান:"
    )
    return

  if is_admin(user_id) and context.user_data.get(
      "waiting_for_add_wallet_desc_step"
  ):
    w_desc = update.message.text.strip()
    context.user_data["waiting_for_add_wallet_desc_step"] = None
    context.user_data["temp_wallet_desc"] = w_desc
    context.user_data["waiting_for_add_wallet_addr_step"] = True
    await update.message.reply_text(
        f"✍ এখন মূল **অ্যাড্রেস বা অ্যাকাউন্ট নম্বর** লিখে পাঠান:"
    )
    return

  if is_admin(user_id) and context.user_data.get(
      "waiting_for_add_wallet_addr_step"
  ):
    w_addr = update.message.text.strip()
    context.user_data["waiting_for_add_wallet_addr_step"] = None
    w_name = context.user_data.pop("temp_wallet_name", "Wallet")
    w_desc = context.user_data.pop("temp_wallet_desc", "Wallet Account")

    settings.setdefault("wallets", {})[w_name] = {
        "desc": w_desc,
        "address": w_addr,
    }
    save_settings(settings)
    await update.message.reply_text(
        f"✅ সফলভাবে **{w_name}** ওয়ালেট অ্যাড্রেস সেভ করা হয়েছে!"
    )
    return

  if is_admin(user_id) and context.user_data.get("waiting_for_user_edit_data"):
    text = update.message.text.strip()
    context.user_data["waiting_for_user_edit_data"] = None
    try:
      parts = text.split()
      if len(parts) < 2:
        raise ValueError("পর্যাপ্ত আর্গুমেন্ট নেই")

      target_id = parts[0].strip()
      action_token = parts[1].strip()

      referrals = load_referrals()
      ref_info = referrals.setdefault("referrers", {}).setdefault(
          target_id,
          {
              "refs": [],
              "withdrawn": 0.0,
              "commission_earned": 0.0,
              "manual_adjustment": 0.0,
          },
      )

      if "ref" in action_token.lower():
        num_val = int(action_token.lower().replace("ref", ""))
        if num_val < 0:
          remove_count = abs(num_val)
          ref_info["refs"] = ref_info["refs"][:-remove_count]
        else:
          for _ in range(abs(num_val)):
            ref_info["refs"].append({
                "uid": "admin_added",
                "bonus": float(
                    settings.get("referral", {}).get("bonus_amount", 10.0)
                ),
                "date": datetime.datetime.now(TZ).strftime("%Y-%m-%d"),
            })
        save_referrals(referrals)
        await update.message.reply_text(
            f"✅ ইউজার `{target_id}` এর রেফারেল সফলভাবে আপডেট করা হয়েছে! (বর্তমান মোট রেফার: {len(ref_info['refs'])})",
            parse_mode="Markdown",
        )
      else:
        amount = float(
            action_token.replace("+", "").replace("-", "").strip()
        )
        current_adj = float(ref_info.get("manual_adjustment", 0.0))
        if "-" in action_token:
          ref_info["manual_adjustment"] = current_adj - amount
        else:
          ref_info["manual_adjustment"] = current_adj + amount

        save_referrals(referrals)
        await update.message.reply_text(
            f"✅ ইউজার `{target_id}` এর ব্যালেন্স সফলভাবে আপডেট করা হয়েছে! (বর্তমান ব্যালেন্স এডজাস্টমেন্ট: ৳{ref_info['manual_adjustment']})",
            parse_mode="Markdown",
        )
    except Exception as e:
      await update.message.reply_text(
          f"❌ ফরম্যাট সঠিক নয়! সঠিক নিয়মে লিখুন (যেমন: `8276619021 +20` বা `8276619021 -50`): {e}",
          parse_mode="Markdown",
      )
    return

  if is_admin(user_id) and context.user_data.get(
      "waiting_for_user_reset_withdrawn"
  ):
    target_id = update.message.text.strip()
    context.user_data["waiting_for_user_reset_withdrawn"] = None
    referrals = load_referrals()
    if target_id in referrals.get("referrers", {}):
      referrals["referrers"][target_id]["withdrawn"] = 0.0
      save_referrals(referrals)
      await update.message.reply_text(
          f"✅ ইউজার `{target_id}` এর উইথড্র হিস্ট্রি সফলভাবে রিসেট করে ৳0 করা হয়েছে!",
          parse_mode="Markdown",
      )
    else:
      await update.message.reply_text(
          "⛔ এই ইউজারের কোনো রেফারেল বা উইথড্র ডাটা পাওয়া যায়নি!"
      )
    return

  if is_admin(user_id) and context.user_data.get(
      "waiting_for_user_reset_referrals_all"
  ):
    target_id = update.message.text.strip()
    context.user_data["waiting_for_user_reset_referrals_all"] = None
    referrals = load_referrals()
    if target_id in referrals.get("referrers", {}):
      del referrals["referrers"][target_id]
      save_referrals(referrals)
      await update.message.reply_text(
          f"✅ ইউজার `{target_id}` এর সমস্ত রেফারেল হিস্ট্রি সফলভাবে রিসেট ও মুছে ফেলা হয়েছে!",
          parse_mode="Markdown",
      )
    else:
      await update.message.reply_text(
          "⛔ এই ইউজারের কোনো রেফারেল ডাটা পাওয়া যায়নি!"
      )
    return

  if is_admin(user_id) and context.user_data.get("waiting_for_user_reset_id"):
    target_id = update.message.text.strip()
    context.user_data["waiting_for_user_reset_id"] = None
    logs = load_logs()
    new_logs = [l for l in logs if str(l.get("user_id")) != target_id]
    save_logs(new_logs)
    await update.message.reply_text(
        f"✅ ইউজার `{target_id}` এর সমস্ত সাবমিশন হিস্ট্রি রিসেট করা হয়েছে!",
        parse_mode="Markdown",
    )
    return

  waiting_popup_info = context.user_data.get("waiting_for_cat_popup_edit")
  if waiting_popup_info and is_admin(user_id):
    popup_type, cat_key = waiting_popup_info
    new_popup = update.message.text.strip()
    if cat_key in settings["categories"]:
      if popup_type == "closed":
        settings["categories"][cat_key]["popup_closed"] = new_popup
      else:
        settings["categories"][cat_key]["popup_shift"] = new_popup
      save_settings(settings)

    context.user_data["waiting_for_cat_popup_edit"] = None
    cat_display_name = settings["categories"][cat_key]["name"]
    await update.message.reply_text(
        f"✅ **{cat_display_name}** ক্যাটাগরির জন্য নতুন পপআপ সফলভাবে সেভ করা"
        f" হয়েছে!\n\n`{new_popup}`",
        parse_mode="Markdown",
    )
    return

  waiting_ref_bonus = context.user_data.get("waiting_for_ref_bonus")
  if waiting_ref_bonus and is_admin(user_id):
    try:
      val = float(update.message.text.strip())
      settings["referral"]["bonus_amount"] = val
      save_settings(settings)
      context.user_data["waiting_for_ref_bonus"] = None
      await update.message.reply_text(
          f"✅ সফলভাবে নতুনদের জন্য রেফারেল বোনাস ৳{val} সেট করা হয়েছে!"
      )
      return
    except:
      await update.message.reply_text("⛔ সঠিক সংখ্যা লিখুন (যেমন: `10`)।")
      return

  waiting_min_refs = context.user_data.get("waiting_for_min_refs")
  if waiting_min_refs and is_admin(user_id):
    try:
      val = int(update.message.text.strip())
      settings["referral"]["min_refs_to_withdraw"] = val
      save_settings(settings)
      context.user_data["waiting_for_min_refs"] = None
      await update.message.reply_text(
          f"✅ সফলভাবে ন্যূনতম রেফার শর্ত {val} টি সেট করা হয়েছে!"
      )
      return
    except:
      await update.message.reply_text("⛔ সঠিক সংখ্যা লিখুন (যেমন: `10`)।")
      return

  waiting_min_amt = context.user_data.get("waiting_for_min_amt")
  if waiting_min_amt and is_admin(user_id):
    try:
      val = float(update.message.text.strip())
      settings["referral"]["min_amount_to_withdraw"] = val
      save_settings(settings)
      context.user_data["waiting_for_min_amt"] = None
      await update.message.reply_text(
          f"✅ সফলভাবে ন্যূনতম উইথড্র পরিমাণ ৳{val} সেট করা হয়েছে!"
      )
      return
    except:
      await update.message.reply_text("⛔ সঠিক সংখ্যা লিখুন (যেমন: `50`)।")
      return

  waiting_search = context.user_data.get("waiting_for_search_user")
  if waiting_search and is_admin(user_id):
    context.user_data["waiting_for_search_user"] = None
    try:
      target_uid = int(update.message.text.strip())
      logs = load_logs()
      user_logs = [l for l in logs if l["user_id"] == target_uid]

      if not user_logs:
        await update.message.reply_text(
            f"⛔ ইউজার আইডি `{target_uid}` এর কোনো সাবমিশন হিস্ট্রি পাওয়া যায়নি!",
            parse_mode="Markdown",
        )
        return

      user_name = user_logs[0]["user_name"]
      total_earned = sum(l.get("earned_amount", 0.0) for l in user_logs)
      paid_earned = sum(
          l.get("earned_amount", 0.0)
          for l in user_logs
          if l.get("is_paid", False)
      )
      due_earned = total_earned - paid_earned

      text = (
          f"🔍 **ইউজার সার্চ ফলাফল:**\n\n"
          f"👤 নাম: [{user_name}](tg://user?id={target_uid})\n"
          f"🆔 আইডি: `{target_uid}`\n"
          f"💰 মোট আয়: ৳{total_earned} | 🟢 পরিশোধিত: ৳{paid_earned} | 🔴 বকেয়া:"
          f" ৳{due_earned}\n\n"
          f"📁 **সাম্প্রতিক সাবমিশনসমূহ:**\n\n"
      )

      for l in user_logs[-10:]:
        p_status = "Paid" if l.get("is_paid", False) else "Unpaid"
        rep_status = "Reported" if l.get("has_report", False) else "Pending"
        text += (
            f"• ক্যাটাগরি: `{l['category']}`\n"
            f"  তারিখ: `{l['date']}`\n"
            f"  ভ্যালিড: `{l['unique']}` | টাকা: `৳{l.get('earned_amount', 0)}`"
            f" ({p_status} | {rep_status})\n\n"
        )

      await update.message.reply_text(text, parse_mode="Markdown")
      return
    except Exception as e:
      await update.message.reply_text(
          f"⛔ ভুল ইউজার আইডি দিয়েছেন অথবা সার্চ করতে সমস্যা হয়েছে: {str(e)}"
      )
      return

  waiting_shift_cat = context.user_data.get("waiting_for_shift_cat")
  if waiting_shift_cat and is_admin(user_id):
    try:
      parts = update.message.text.strip().split()
      start_t, end_t = parts[0], parts[1]
      datetime.datetime.strptime(start_t, "%H:%M")
      datetime.datetime.strptime(end_t, "%H:%M")

      CATEGORY_SHIFTS[waiting_shift_cat] = {"start": start_t, "end": end_t}
      context.user_data["waiting_for_shift_cat"] = None
      cat_name = settings["categories"][waiting_shift_cat]["name"]
      await update.message.reply_text(
          f"✅ সফলভাবে **{cat_name}** ক্যাটাগরির শিফট সময় সেট করা হয়েছে: প্রতিদিন"
          f" `{start_t}` থেকে `{end_t}` পর্যন্ত। ⏰",
          parse_mode="Markdown",
      )
      return
    except:
      await update.message.reply_text(
          "⛔ ভুল ফরম্যাট! দয়া করে সঠিক ফরম্যাটে সময় দিন যেমন: `10:30 22:45`",
          parse_mode="Markdown",
      )
      return

  is_member = await check_channel_membership(context.bot, user_id)
  if not is_member:
    join_ch = settings.get("channels", {}).get("join_channel", "@OMORTECH")
    link_url, link_disp = format_channel_link(join_ch)
    join_msg = (
        f"⛔ **সম্মানিত {user.first_name}, চ্যানেলে জয়েন করা বাধ্যতামূলক!**\n\n"
        "বটটি ব্যবহার করতে হলে প্রথমে আমাদের অফিসিয়াল চ্যানেলে জয়েন করতে"
        " হবে। নিচে বাটনে ক্লিক করে জয়েন করুন:\n\n"
        f"✨ চ্যানেলে জয়েন করার পর নিচের **🔄 স্টার্ট / রিফ্রেশ করুন** বাটনে"
        " ক্লিক করুন."
    )
    join_kbd = [
        [InlineKeyboardButton("📢 চ্যানেলে জয়েন করুন", url=link_url)],
        [
            InlineKeyboardButton(
                "🔄 স্টার্ট / রিফ্রেশ করুন", callback_data="back_home"
            )
        ],
    ]
    await update.message.reply_text(
        join_msg,
        reply_markup=InlineKeyboardMarkup(join_kbd),
        parse_mode="Markdown",
        disable_web_page_preview=True,
    )
    return

  waiting_custom_date_cat = context.user_data.get("waiting_for_custom_date_cat")
  if waiting_custom_date_cat and is_admin(user_id):
    date_text = update.message.text.strip()
    context.user_data["waiting_for_custom_date_cat"] = None

    cat_name = settings["categories"][waiting_custom_date_cat]["name"]

    target_date = date_text.replace("/", "-").strip()
    try:
      dt_obj = datetime.datetime.strptime(target_date, "%d-%m-%Y")
      search_str = dt_obj.strftime("%Y-%m-%d")
    except:
      search_str = target_date

    logs = load_logs()
    matching_logs = [
        l
        for l in logs
        if l["category"] == cat_name and search_str in l["date"]
    ]

    if not matching_logs:
      await update.message.reply_text(
          f"⛔ তারিখ `{date_text}` এ **{cat_name}** ক্যাটাগরিতে কোনো সাবমিশন"
          " পাওয়া যায়নি!",
          parse_mode="Markdown",
      )
      return

    matching_valid_ids = set()
    for l in matching_logs:
      for u_id in l.get("submitted_ids_list", []):
        matching_valid_ids.add(u_id)

    all_rows_db = load_all_rows_db()
    cat_rows = all_rows_db.get(cat_name, [])

    matching_rows = []
    for r in cat_rows:
      if r and len(r) > 0:
        val_str = str(r[0]).strip()
        if val_str.endswith(".0"):
          val_str = val_str[:-2]
        if val_str in matching_valid_ids:
          matching_rows.append(r)

    if not matching_rows:
      await update.message.reply_text(
          f"⛔ তারিখ `{date_text}` এ **{cat_name}** ক্যাটাগরিতে কোনো এন্ট্রি নেই!",
          parse_mode="Markdown",
      )
      return

    export_wb = openpyxl.Workbook()
    export_ws = export_wb.active
    export_ws.title = "Custom Date Export"

    for r_data in matching_rows:
      export_ws.append(r_data)

    safe_name = cat_name.replace(" ", "_")
    file_name = f"Custom_Date_{safe_name}_{date_text}.xlsx"
    export_wb.save(file_name)

    with open(file_name, "rb") as f:
      await context.bot.send_document(
          chat_id=user_id,
          document=f,
          caption=(
              f"🗓 **{cat_name}** ক্যাটাগরির কাস্টম তারিখ `{date_text}` এর সমস্ত"
              " জমা হওয়া আইডি ফাইলের মাধ্যমে সংযুক্ত করা হলো! 🚀"
          ),
          parse_mode="Markdown",
      )
    os.remove(file_name)
    return

  waiting_price_cat = context.user_data.get("waiting_for_new_price_cat")
  if waiting_price_cat and is_admin(user_id):
    text_val = update.message.text.strip()
    try:
      new_price = float(text_val)
      settings["categories"][waiting_price_cat]["price"] = new_price
      save_settings(settings)

      context.user_data["waiting_for_new_price_cat"] = None
      cat_name = settings["categories"][waiting_price_cat]["name"]

      await update.message.reply_text(
          f"✅ সফলভাবে **{cat_name}** ক্যাটাগরির নতুন রেট নির্ধারণ করা হয়েছে:"
          f" **{new_price}৳** 🌟",
          parse_mode="Markdown",
      )
      return
    except ValueError:
      await update.message.reply_text(
          "⛔ ভুল মূল্য দিয়েছেন! দয়া করে সঠিক সংখ্যা লিখুন (যেমন: `5` বা"
          " `6.5`)。",
          parse_mode="Markdown",
      )
      return

  waiting_method = context.user_data.get("waiting_for_payment_method")
  if waiting_method:
    number = update.message.text.strip()

    if waiting_method in ["bkash", "nagad", "rocket"]:
      if not (
          number.isdigit() and len(number) == 11 and number.startswith("01")
      ):
        await update.message.reply_text(
            f"⛔ সম্মানিত {user.first_name}, ভুল নম্বর দেওয়া হয়েছে!\n\nদয়া করে সঠিক **১১ সংখ্যার** বাংলাদেশি"
            " মোবাইল নম্বর দিন যা `01` দিয়ে শুরু.",
            parse_mode="Markdown"
        )
        return

    user_id_str = str(user_id)
    user_payments = load_user_payments()
    if user_id_str not in user_payments:
      user_payments[user_id_str] = {}

    user_payments[user_id_str][waiting_method] = number
    save_user_payments(user_payments)

    context.user_data["waiting_for_payment_method"] = None

    await update.message.reply_text(
        f"✅ প্রিয় {user.first_name}, আপনার পেমেন্ট নম্বর সফলভাবে সেভ করা হয়েছে! 🌟\n\nফাইল সাবমিট করার"
        " জন্য আবার `/start` দিন অথবা মেনু থেকে ক্যাটাগরি সিলেক্ট করুন.",
        parse_mode="Markdown",
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
  user = update.effective_user
  bot = context.bot

  if user.id in load_blocked_users():
    await update.message.reply_text(f"⛔ {user.first_name}, আপনাকে ব্লক করা হয়েছে।")
    return

  settings = load_settings()
  join_ch = settings.get("channels", {}).get("join_channel", "@OMORTECH")
  link_url, link_disp = format_channel_link(join_ch)

  is_member = await check_channel_membership(bot, user.id)
  if not is_member:
    join_msg = (
        f"⛔ **সম্মানিত {user.first_name}, চ্যানেলে জয়েন করা বাধ্যতামূলক!**\n\n"
        "বটটি ব্যবহার করতে হলে প্রথমে আমাদের অফিসিয়াল চ্যানেলে জয়েন করতে"
        " হবে। নিচে বাটনে ক্লিক করে জয়েন করুন:\n\n"
        f"✨ চ্যানেলে জয়েন করার পর নিচের **🔄 স্টার্ট / রিফ্রেশ করুন** বাটনে"
        " ক্লিক করুন."
    )
    join_kbd = [
        [InlineKeyboardButton("📢 চ্যানেলে জয়েন করুন", url=link_url)],
        [
            InlineKeyboardButton(
                "🔄 স্টার্ট / রিফ্রেশ করুন", callback_data="back_home"
            )
        ],
    ]
    await update.message.reply_text(
        join_msg,
        reply_markup=InlineKeyboardMarkup(join_kbd),
        parse_mode="Markdown",
        disable_web_page_preview=True,
    )
    return

  doc = update.message.document
  if doc:
    file_name_lower = doc.file_name.lower()
    if not (
        file_name_lower.endswith(".xlsx") or file_name_lower.endswith(".txt")
    ):
      original_name = doc.file_name
      wrong_file_msg = (
          f"⛔ **{user.first_name}, ফাইল গ্রহণ করা হয়নি!**\n\n"
          f"আপনার পাঠানো ফাইলটি (`{original_name}`) সঠিক এক্সেল ফাইল নয়। "
          "অনুগ্রহ করে শুধুমাত্র সঠিক **.xlsx** ফরম্যাতের এক্সেল ফাইলটি তৈরি করে"
          " আবার আপলোড করুন। 📂✨"
      )
      await update.message.reply_text(wrong_file_msg, parse_mode="Markdown")
      return

  waiting_report_info = context.user_data.get("waiting_for_report_info")
  if waiting_report_info and is_admin(user.id):
    try:
      cat_name = waiting_report_info["cat_name"]
      target_date = waiting_report_info["date"]

      file = await update.message.document.get_file()
      file_path = "downloaded_report_file"
      await file.download_to_drive(file_path)

      valid_ids = set()
      if update.message.document.file_name.endswith(".xlsx"):
        df_rep = pd.read_excel(file_path, header=None)
        for r_idx in range(len(df_rep)):
          val = df_rep.iloc[r_idx, 0]
          if not pd.isna(val):
            v_str = str(val).strip()
            if v_str.endswith(".0"):
              v_str = v_str[:-2]
            valid_ids.add(v_str)
      else:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
          for line in f:
            clean_l = line.strip()
            if clean_l:
              if clean_l.endswith(".0"):
                clean_l = clean_l[:-2]
              valid_ids.add(clean_l)

      context.user_data["waiting_for_report_info"] = None

      rate = 5.0
      for k, v in settings["categories"].items():
        if v["name"] == cat_name:
          rate = v["price"]
          break

      logs = load_logs()
      display_date = datetime.datetime.strptime(
          target_date, "%Y-%m-%d"
      ).strftime("%d-%m-%Y")
      summary = (
          f"📊 **অটোমেটেড রিপোর্ট ম্যাচিং সামারি ({cat_name} -"
          f" {display_date}):**\n\n"
      )

      match_found = False
      for l in logs:
        if l["category"] == cat_name and l["date"].startswith(target_date):
          match_found = True
          target_uid = l["user_id"]
          user_submitted_list = l.get("submitted_ids_list", [])
          original_unique = l.get("unique", 0)

          net_valid = 0
          for u_id in user_submitted_list:
            if u_id in valid_ids:
              net_valid += 1

          net_rejected = max(0, original_unique - net_valid)

          l["unique"] = net_valid
          l["rejected"] = net_rejected
          payout_amt = net_valid * rate
          l["earned_amount"] = float(payout_amt)
          l["has_report"] = True

          summary += (
              f"👤 [{l['user_name']}](tg://user?id={target_uid}) (`{target_uid}`)\n"
              f"  ✅ ভ্যালিড আইডি: {net_valid} টি\n"
              f"  ❌ রিজেক্ট আইডি: {net_rejected} টি\n"
              f"  💰 পেমেন্ট হিসাব: `{payout_amt}৳`\n\n"
          )

      save_logs(logs)
      if not match_found:
        summary += "⛔ এই তারিখের কোনো সাবমিশন পাওয়া যায়নি।"

      await update.message.reply_text(summary, parse_mode="Markdown")
      if os.path.exists(file_path):
        os.remove(file_path)
      return
    except Exception as e:
      await update.message.reply_text(
          f"❌ রিপোর্ট প্রসেস করতে সমস্যা হয়েছে: {str(e)}"
      )
      return

  user_id_str = str(user.id)
  user_payments = load_user_payments()
  user_p = user_payments.get(user_id_str, {})
  has_valid_payment = any(
      p_v != "সেট করা হয়নি" for p_k, p_v in user_p.items()
  )

  if not has_valid_payment and not is_admin(user.id):
    no_pay_msg = (
        f"⛔ **{user.first_name}, ফাইল সাবমিট ব্যর্থ হয়েছে!**\n\n"
        "ফাইল জমা দেওয়ার আগে অন্তত একটি পেমেন্ট মাধ্যম বা নম্বর (যেমন: bKash বা"
        " Nagad) যুক্ত করা বাধ্যতামূলক। "
        "দয়া করে বট মেনু থেকে **⚙ পেমেন্ট সেটিংস** এ গিয়ে আপনার পেমেন্ট নম্বরটি"
        " সেট করে নিন।"
    )
    await update.message.reply_text(no_pay_msg, parse_mode="Markdown")
    return

  selected_cat_key = context.user_data.get("selected_category_key")
  selected_cat = context.user_data.get("selected_category")
  selected_cat_price = context.user_data.get("selected_category_price", 5.0)

  if not selected_cat_key:
    await update.message.reply_text(
        f"⛔ প্রিয় {user.first_name}, দয়া করে প্রথমে `/start` দিয়ে ক্যাটাগরি সিলেক্ট করুন।",
        parse_mode="Markdown",
    )
    return

  cat_info = settings["categories"].get(selected_cat_key, {})
  status_open = cat_info.get("status", True)
  shift_open = is_shift_active(selected_cat_key)

  if (not status_open or not shift_open) and not is_admin(user.id):
    if not status_open:
      tpl = cat_info.get(
          "popup_closed",
          "দুঃখিত {name}, বর্তমানে {category} এর ফাইল জমা নেওয়া বন্ধ রয়েছে! ❌",
      )
    else:
      tpl = cat_info.get(
          "popup_shift", "দুঃখিত {name}, বর্তমানে {category} শিফটের বাইরে রয়েছে! ⏰"
      )
    custom_msg = tpl.replace("{category}", cat_info.get("name", selected_cat)).replace("{name}", user.first_name)
    await update.message.reply_text(custom_msg, parse_mode="Markdown")
    return

  progress_msg = await update.message.reply_text(
      f"⌛ **প্রিয় {user.first_name}, ফাইল প্রসেসিং শুরু হয়েছে...**\n`[▒▒▒▒▒▒▒▒▒▒] 0%`", parse_mode="Markdown"
  )

  file = await update.message.document.get_file()
  input_path = "downloaded_file.xlsx"
  await file.download_to_drive(input_path)

  try:
    await progress_msg.edit_text(
        "🔍 **ফাইল রিডিং ও কলাম চেক করা হচ্ছে...**\n`[████▒▒▒▒▒▒] 40%`",
        parse_mode="Markdown",
    )

    df = pd.read_excel(input_path, engine="openpyxl", header=None)

    if df.shape[1] < 3:
      await progress_msg.edit_text(
          f"⛔ {user.first_name}, ফাইল রিজেক্ট করা হয়েছে!\nকারণ: ফাইলে পর্যাপ্ত কলাম (A, B, C) নেই।",
          parse_mode="Markdown"
      )
      return

    total_rows = len(df)
    if total_rows == 0:
      await progress_msg.edit_text(f"⛔ {user.first_name}, ফাইলটি খালি রয়েছে!", parse_mode="Markdown")
      return

    first_data_idx = -1
    last_data_idx = -1

    for row_idx in range(total_rows):
      val_a = df.iloc[row_idx, 0]
      val_b = df.iloc[row_idx, 1]
      val_c = df.iloc[row_idx, 2]
      is_empty = pd.isna(val_a) and pd.isna(val_b) and pd.isna(val_c)

      has_extra = False
      if df.shape[1] > 3:
        for c_idx in range(3, df.shape[1]):
          if (
              not pd.isna(df.iloc[row_idx, c_idx])
              and str(df.iloc[row_idx, c_idx]).strip() != ""
          ):
            has_extra = True
            break

      if not (is_empty and not has_extra):
        if first_data_idx == -1:
          first_data_idx = row_idx
        last_data_idx = row_idx

    blank_row_nums = []
    if first_data_idx != -1 and last_data_idx != -1:
      for row_idx in range(first_data_idx, last_data_idx + 1):
        val_a = df.iloc[row_idx, 0]
        val_b = df.iloc[row_idx, 1]
        val_c = df.iloc[row_idx, 2]
        is_empty = pd.isna(val_a) and pd.isna(val_b) and pd.isna(val_c)

        has_extra = False
        if df.shape[1] > 3:
          for c_idx in range(3, df.shape[1]):
            if (
                not pd.isna(df.iloc[row_idx, c_idx])
                and str(df.iloc[row_idx, c_idx]).strip() != ""
            ):
              has_extra = True
              break

        if is_empty and not has_extra:
          blank_row_nums.append(row_idx + 1)

    raw_errors = []

    if blank_row_nums:
      sorted_blanks = sorted(list(set(blank_row_nums)))
      ranges = []
      start_r = sorted_blanks[0]
      end_r = sorted_blanks[0]
      for n in sorted_blanks[1:]:
        if n == end_r + 1:
          end_r = n
        else:
          ranges.append((start_r, end_r))
          start_r = n
          end_r = n
      ranges.append((start_r, end_r))

      for s, e in ranges:
        if s == e:
          raw_errors.append((
              s,
              f"{s} নং সারি: মাঝখানে কোনো ফাঁকা সারি (Blank Row) রাখা যাবে না।"
              " সিরিয়াল অনুযায়ী ডেটা দিন.",
          ))
        else:
          raw_errors.append((
              s,
              f"{s}-{e} নং সারি: মাঝখানে কোনো ফাঁকা সারি (Blank Row) রাখা যাবে"
              " না। সিরিয়াল অনুযায়ী ডেটা দিন।",
          ))

    for row_idx in range(total_rows):
      excel_row_num = row_idx + 1

      val_a = df.iloc[row_idx, 0]
      val_b = df.iloc[row_idx, 1]
      val_c = df.iloc[row_idx, 2]

      is_row_empty = pd.isna(val_a) and pd.isna(val_b) and pd.isna(val_c)

      has_extra_data = False
      if df.shape[1] > 3:
        for col_idx in range(3, df.shape[1]):
          ex_v = df.iloc[row_idx, col_idx]
          if not pd.isna(ex_v) and str(ex_v).strip() != "":
            has_extra_data = True
            break

      if is_row_empty and not has_extra_data:
        continue

      if any(
          x in selected_cat_key
          for x in ["mail", "gmail", "yahoo", "outlook", "hotmail"]
      ):
        col_a_label = "Email Address"
      elif "ig" in selected_cat_key:
        col_a_label = "Username"
      else:
        col_a_label = "UID"

      col_c_label = "2FA Key" if "2fa" in selected_cat_key else "Cookies"

      row_issue_parts = []

      is_a_empty = pd.isna(val_a) or str(val_a).strip() == ""
      is_b_empty = pd.isna(val_b) or str(val_b).strip() == ""
      is_c_empty = pd.isna(val_c) or str(val_c).strip() == ""

      if is_a_empty:
        row_issue_parts.append(f"কলাম A ({col_a_label}) ফাঁকা রয়েছে")
      if is_b_empty:
        row_issue_parts.append("কলাম B (Password) ফাঁকা রয়েছে")
      if is_c_empty:
        row_issue_parts.append(f"কলাম C ({col_c_label}) ফাঁকা রয়েছে")

      if row_issue_parts:
        combined_msg = " এবং ".join(row_issue_parts) + "।"
        raw_errors.append((excel_row_num, combined_msg))
        continue

      row_val_a = str(val_a).strip()
      if row_val_a.endswith(".0"):
        row_val_a = row_val_a[:-2]

      row_val_c = str(val_c).strip()

      if any(
          x in selected_cat_key
          for x in ["mail", "gmail", "yahoo", "outlook", "hotmail"]
      ):
        if not ("@" in row_val_a and "." in row_val_a and len(row_val_a) > 5):
          raw_errors.append((
              excel_row_num,
              "কলাম A-তে সঠিক Email Address দেওয়া হয়নি।",
          ))
      elif selected_cat_key in ["fb_2fa", "fb_cookies", "pc_clone"]:
        if not (
            row_val_a.isdigit()
            and len(row_val_a) == 14
            and row_val_a.startswith("61")
        ):
          raw_errors.append((
              excel_row_num,
              "কলাম A-তে সঠিক ১৪ সংখ্যার UID (যা ৬টি ১ বা ৬১ দিয়ে শুরু) দেওয়া হয়নি।",
          ))

      elif selected_cat_key in ["ig_2fa", "ig_cookies"]:
        if row_val_a.isdigit() and len(row_val_a) > 10:
          raw_errors.append(
              (excel_row_num, "কলাম A-তে সঠিক Username দেওয়া হয়নি।")
          )

      if selected_cat_key in ["ig_2fa", "fb_2fa"]:
        if (
            len(row_val_c) > 50
            or "c_user" in row_val_c.lower()
            or "datr" in row_val_c.lower()
            or "sb=" in row_val_c.lower()
        ):
          raw_errors.append(
              (excel_row_num, "কলাম C-তে সঠিক 2FA Key দেওয়া হয়নি।")
          )

      elif selected_cat_key in ["ig_cookies", "fb_cookies", "pc_clone"]:
        if len(row_val_c) < 15 or not (
            "=" in row_val_c
            or ";" in row_val_c
            or "c_user" in row_val_c.lower()
        ):
          raw_errors.append(
              (excel_row_num, "কলাম C-তে সঠিক Cookies দেওয়া হয়নি।")
          )

      if df.shape[1] > 3:
        for col_idx in range(3, df.shape[1]):
          extra_val = df.iloc[row_idx, col_idx]
          if not pd.isna(extra_val) and str(extra_val).strip() != "":
            col_letter = openpyxl.utils.get_column_letter(col_idx + 1)
            raw_errors.append((
                excel_row_num,
                f"অতিরিক্ত কলাম {col_letter}-তে ভুল বা অতিরিক্ত ডেটা রয়েছে।",
            ))
            break

    raw_errors.sort(key=lambda x: x[0])

    if raw_errors:
      formatted_error_lines = []
      error_limit = 20
      for r_num, err_msg in raw_errors[:error_limit]:
        formatted_error_lines.append(f"• **{r_num} নং সারি:** {err_msg}")

      if len(raw_errors) > error_limit:
        formatted_error_lines.append("\n... এবং আরও ত্রুটি রয়েছে।")

      final_error_text = "\n".join(formatted_error_lines)

      rejection_msg = (
          f"⛔ **{user.first_name}, আপনার ফাইল প্রত্যাখ্যান করা হয়েছে!**\n\n⚠ ফাইলের সুনির্দিষ্ট ত্রুটিসমূহ:\n"
          f"━━━━━━━━━━━━━━━━━━━━\n{final_error_text}\n━━━━━━━━━━━━━━━━━━━━\n💡"
          " দয়া করে উপরোক্ত সারির ভুলগুলো সংশোধন করে ফাইলটি আবার সেন্ড করুন."
      )
      await progress_msg.edit_text(rejection_msg, parse_mode="Markdown")
      return

    await progress_msg.edit_text(
        "🔄 **ডুপ্লিকেট ও গ্লোবাল আইডি ম্যাচিং চলছে...**\n`[████████▒▒] 80%`",
        parse_mode="Markdown",
    )

    rows_data = []
    for row_idx in range(total_rows):
      r_vals = [df.iloc[row_idx, col] for col in range(df.shape[1])]
      cleaned_vals = []
      for v in r_vals:
        if pd.isna(v):
          cleaned_vals.append("")
        else:
          try:
            if hasattr(v, "item"):
              v = v.item()
            if isinstance(v, float) and v.is_integer():
              v = int(v)
            elif isinstance(v, (int, np.integer)):
              v = int(v)
            else:
              v = str(v).strip()
              if v.endswith(".0"):
                v = v[:-2]
          except:
            v = str(v).strip()
          cleaned_vals.append(v)
      rows_data.append(cleaned_vals)

    all_entries = [
        str(r[0]).strip()
        for r in rows_data
        if not (pd.isna(r[0]) or str(r[0]).strip() == "")
    ]
    total_count = len(all_entries)

    global_ids = load_global_ids()

    seen_in_file = set()
    unique_entries = []
    file_duplicates = 0
    already_submitted = 0
    admin_rows_data = []
    unique_rows_data = []

    for idx, item in enumerate(all_entries):
      clean_item = item.strip()
      if clean_item.endswith(".0"):
        clean_item = clean_item[:-2]

      if clean_item in seen_in_file:
        file_duplicates += 1
      elif clean_item in global_ids:
        already_submitted += 1
        seen_in_file.add(clean_item)
      else:
        seen_in_file.add(clean_item)
        unique_entries.append(clean_item)
        global_ids.append(clean_item)
        for r_data in rows_data:
          if not pd.isna(r_data[0]) and str(r_data[0]).strip() == clean_item:
            admin_rows_data.append(r_data)
            unique_rows_data.append(r_data)
            break

    save_global_ids(global_ids)

    all_rows_db = load_all_rows_db()
    if selected_cat not in all_rows_db:
      all_rows_db[selected_cat] = []
    all_rows_db[selected_cat].extend(unique_rows_data)
    save_all_rows_db(all_rows_db)

    net_unique = len(set(unique_entries))
    total_duplicates = file_duplicates + already_submitted

    earned_amount = net_unique * selected_cat_price

    logs = load_logs()
    logs.append({
        "date": datetime.datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "user_id": int(user.id),
        "user_name": str(user.first_name),
        "category": str(selected_cat),
        "total": int(total_count),
        "unique": int(net_unique),
        "rejected": 0,
        "duplicates": int(total_duplicates),
        "earned_amount": float(earned_amount),
        "is_paid": False,
        "has_report": False,
        "submitted_ids_list": unique_entries,
    })
    save_logs(logs)

    wb = openpyxl.load_workbook(input_path)
    ws = wb.active
    red_fill = PatternFill(
        start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
    )

    checked_set = set()
    existing_global_set = set(
        global_ids[: -len(unique_entries)] if unique_entries else global_ids
    )

    for row in range(1, ws.max_row + 1):
      val = ws.cell(row=row, column=1).value
      if val is not None:
        val_str = str(val).strip()
        if val_str.endswith(".0"):
          val_str = val_str[:-2]
        if (
            val_str in checked_set
            or all_entries.count(val_str) > 1
            or val_str in existing_global_set
        ):
          for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = red_fill
        else:
          checked_set.add(val_str)

    current_date_str = datetime.datetime.now(TZ).strftime("%d-%m-%Y")
    safe_cat_name = selected_cat.replace(" ", "_")
    highlighted_path = f"{current_date_str}_{safe_cat_name}.xlsx"
    wb.save(highlighted_path)

    admin_wb = openpyxl.Workbook()
    admin_ws = admin_wb.active
    for r_data in admin_rows_data:
      admin_ws.append(r_data)

    admin_file_path = f"admin_{current_date_str}_{safe_cat_name}.xlsx"
    admin_wb.save(admin_file_path)

    await progress_msg.delete()

    report_text = (
        f"📊 **প্রিয় {user.first_name}, আপনার ফাইল রিপোর্ট সফলভাবে তৈরি হয়েছে!**\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        f"📂 ক্যাটাগরি: {selected_cat}\n"
        f"🔢 মোট সাবমিট: {total_count}\n"
        f"✅ নতুন ইউনিক: {net_unique}\n"
        f"💵 প্রতি রেট: {selected_cat_price}৳\n"
        f"💰 মোট পাওনা: **{earned_amount}৳**\n"
        f"♻️ নিজের ডুপ্লিকেট: {file_duplicates}\n"
        f"⚠️ আগে জমা দিয়েছেন: {already_submitted}\n"
        f"📁 স্ট্যাটাস: সফলভাবে জমা হয়েছে! 🎉"
    )
    await update.message.reply_text(report_text, parse_mode="Markdown")

    if file_duplicates > 0:
      caption_text = (
          f"📥 {user.first_name}, আপনার ফাইলের {file_duplicates} টি ডুপ্লিকেট আইডি লাল রঙে মার্ক"
          " করা হয়েছে"
      )
      with open(highlighted_path, "rb") as f:
        await bot.send_document(
            chat_id=user.id, document=f, caption=caption_text
        )
    else:
      with open(highlighted_path, "rb") as f:
        await bot.send_document(chat_id=user.id, document=f)

    user_payments = load_user_payments()
    user_p = user_payments.get(user_id_str, {})

    active_payments = []
    for p_key, p_val in user_p.items():
      if p_val != "সেট করা হয়নি":
        active_payments.append(f"🔹 {p_key.capitalize()}: `{p_val}`")

    if active_payments:
      payment_info_str = "\n" + "\n".join(active_payments)
    else:
      payment_info_str = " সেট করা হয়নি"

    user_link = f"[{user.first_name}](tg://user?id={user.id})"

    caption = (
        "📥 **নতুন ফাইল সাবমিট হয়েছে!**\n\n"
        f"👤 ইউজার: {user_link}\n"
        f"🆔 আইডি: `{user.id}`\n"
        f"📂 ক্যাটাগরি: {selected_cat}\n"
        f"💳 পেমেন্ট নম্বর:{payment_info_str}\n"
        f"📊 মোট: {total_count} | ইউনিক: {net_unique} | টাকা: **{earned_amount}৳**"
    )

    sub_ch_raw = settings.get("channels", {}).get(
        "submission_channel", DEFAULT_ADMIN_ID
    )
    sub_ch = parse_chat_id(sub_ch_raw)
    with open(admin_file_path, "rb") as admin_f:
      await bot.send_document(
          chat_id=sub_ch,
          document=admin_f,
          caption=caption,
          parse_mode="Markdown",
      )

  except Exception as e:
    try:
      await progress_msg.edit_text(
          f"⛔ ফাইল প্রসেস করতে গিয়ে সমস্যা হয়েছে!\nত্রুটি: {str(e)}"
      )
    except:
      await update.message.reply_text(
          f"⛔ ফাইল প্রসেস করতে গিয়ে সমস্যা হয়েছে!\nত্রুটি: {str(e)}"
      )


def main():
  application = (
      ApplicationBuilder()
      .token(TOKEN)
      .read_timeout(30)
      .write_timeout(30)
      .connect_timeout(30)
      .pool_timeout(30)
      .build()
  )

  application.add_handler(CommandHandler("start", start))
  application.add_handler(CommandHandler("help", help_command))
  application.add_handler(CommandHandler("status", status_command))
  application.add_handler(CommandHandler("users", users_list_command))
  application.add_handler(CommandHandler("resetall", reset_all_logs_command))
  application.add_handler(CommandHandler("clearids", clear_global_ids_command))
  application.add_handler(CommandHandler("top", top_command))
  application.add_handler(CommandHandler("block", block_command))
  application.add_handler(CommandHandler("unblock", unblock_command))
  application.add_handler(CommandHandler("broadcast", broadcast_command))
  application.add_handler(CommandHandler("send", send_specific_message))

  application.add_handler(CallbackQueryHandler(button_callback))
  application.add_handler(MessageHandler(filters.PHOTO, handle_photo))
  application.add_handler(
      MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text_message)
  )
  application.add_handler(MessageHandler(filters.Document.ALL, handle_document))

  print("🤖 Bot is running smoothly...")
  application.run_polling()


if __name__ == "__main__":
  main()